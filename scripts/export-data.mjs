import fs from 'node:fs'
import path from 'node:path'
import process from 'node:process'
import ExcelJS from 'exceljs'

const root = path.resolve(import.meta.dirname, '..')
const sourceWorkbook =
  process.env.CYBERSMART_WORKBOOK ??
  path.join(root, 'source', 'WeST-CyberSmart-Final-Validated-Assessment.xlsx')
const sourceServerInventory = path.join(root, 'source', 'server-inventory.json')
const outputDirectory = path.join(root, 'public', 'data')
const downloadDirectory = path.join(root, 'public', 'downloads')
const downloadFile = path.join(downloadDirectory, 'WeST-CyberSmart-Final-Validated-Assessment.xlsx')

function keyOf(header) {
  const words = String(header)
    .replaceAll('%', ' percent ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .split(/[^a-zA-Z0-9]+/)
    .filter(Boolean)

  return words
    .map((word, index) => {
      const lower = word.toLowerCase()
      return index === 0 ? lower : `${lower[0].toUpperCase()}${lower.slice(1)}`
    })
    .join('')
}

function valueOf(cell) {
  const value = cell.value
  if (value === null || value === undefined) return ''
  if (typeof value === 'number' || typeof value === 'boolean') return value
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object') {
    if ('result' in value && value.result !== undefined) return value.result
    if ('text' in value && typeof value.text === 'string') return value.text
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((item) => item.text).join('')
    }
  }
  return cell.text
}

function rowsFrom(workbook, sheetName) {
  const sheet = workbook.getWorksheet(sheetName)
  if (!sheet) throw new Error(`Missing worksheet: ${sheetName}`)

  const headers = []
  sheet.getRow(5).eachCell({ includeEmpty: true }, (cell, column) => {
    headers[column] = cell.text.trim()
  })

  const rows = []
  for (let rowNumber = 6; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber)
    if (!row.getCell(1).text.trim()) continue
    const record = {}
    headers.forEach((header, column) => {
      if (!header || !column) return
      record[keyOf(header)] = valueOf(row.getCell(column))
    })
    rows.push(record)
  }
  return rows
}

function numberOf(value) {
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : 0
}

function buildSites(navigatorRows) {
  return navigatorRows.map((row) => {
    const passChecks = {}
    const failChecks = {}
    const remainingByType = {}

    Object.entries(row).forEach(([key, value]) => {
      if (key.startsWith('pass') && !['passing', 'passPercent'].includes(key)) {
        passChecks[key.slice(4)] = numberOf(value)
      } else if (key.startsWith('fail') && key !== 'failing') {
        failChecks[key.slice(4)] = numberOf(value)
      } else if (key.startsWith('remaining') && key !== 'remainingToInstall') {
        remainingByType[key.slice(9)] = numberOf(value)
      }
    })

    return {
      site: row.openSite,
      compatibleDevices: numberOf(row.totalCompatibleDevices),
      cyberSmartDevices: numberOf(row.cyberSmartDevices),
      serverInstalls: numberOf(row.serverCyberSmartInstalls),
      serialMatched: numberOf(row.serialMatchTotal),
      notSerialMatched: numberOf(row.notSerialMatched),
      remainingToInstall: numberOf(row.remainingToInstall),
      passing: numberOf(row.passing),
      failing: numberOf(row.failing),
      passPercent: numberOf(row.passPercent),
      possibleUnmanaged: numberOf(row.possibleUnmanaged),
      priority: row.priority,
      passChecks,
      failChecks,
      remainingByType,
    }
  })
}

function summaryFromPending(rows) {
  const values = Object.fromEntries(
    rows.map((row) => [keyOf(row.statusMeasure), numberOf(row.deviceCount)]),
  )

  return {
    portalTotal: values.totalCyberSmartDevices,
    detailedDevices: values.installedReporting,
    pending: values.pending,
    inventoryInstallationGap: values.inventoryInstallationGap,
    portalSecure: values.portalSecure,
    csvPassing: values.csvDerivedPassing,
    csvFailing: values.csvDerivedFailing,
    recentlyActive: values.recentlyActive,
  }
}

function identity(value) {
  return String(value ?? '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '')
}

function validIdentity(value) {
  const normalized = identity(value)
  return normalized && !['NA', 'NONE', 'UNKNOWN', 'NOTRECORDED'].includes(normalized)
}

function buildServerMatches(cyberSmartDevices, serverInventoryPayload) {
  const inventory = serverInventoryPayload.inventory.map((server) => ({
    ...server,
    schoolKey: identity(server.school),
    serialKey: validIdentity(server.serialNumber) ? identity(server.serialNumber) : '',
    hostKey: identity(server.serverName),
  }))
  const cyberServers = cyberSmartDevices.filter((device) => device.deviceType === 'Server')

  const matches = cyberServers.map((device) => {
    const schoolKey = identity(device.reportSite)
    const serialKey = validIdentity(device.serial) ? identity(device.serial) : ''
    const hostKey = identity(device.host)
    const sameSite = inventory.filter((server) => server.schoolKey === schoolKey)

    const sameSiteSerial = serialKey
      ? sameSite.filter((server) => server.serialKey === serialKey)
      : []
    const globalSerial = serialKey
      ? inventory.filter((server) => server.serialKey === serialKey)
      : []
    const sameSiteHost = hostKey
      ? sameSite.filter((server) => server.hostKey === hostKey)
      : []
    const similarHost = hostKey.length >= 5
      ? sameSite.filter(
          (server) =>
            server.hostKey.length >= 5 &&
            (server.hostKey.includes(hostKey) || hostKey.includes(server.hostKey)),
        )
      : []

    let matched = null
    let method = 'Unmatched'
    if (sameSiteSerial.length === 1) {
      matched = sameSiteSerial[0]
      method = 'Exact serial - same site'
    } else if (globalSerial.length === 1) {
      matched = globalSerial[0]
      method = 'Exact serial - different site'
    } else if (sameSiteHost.length === 1) {
      matched = sameSiteHost[0]
      method = 'Exact hostname - same site'
    } else if (similarHost.length === 1) {
      matched = similarHost[0]
      method = 'Hostname similarity - same site'
    }

    return {
      site: device.reportSite,
      host: device.host,
      serial: device.serial,
      status: device.overallStatus,
      os: `${device.osName} ${device.osVersion}`.trim(),
      failedChecks: device.failedChecks,
      failureReasons: device.failureReasons,
      lastCheckIn: device.lastCheckIn,
      matchMethod: method,
      matched: Boolean(matched),
      inventoryItemId: matched?.inventoryItemId ?? '',
      inventorySchool: matched?.school ?? '',
      inventoryServerName: matched?.serverName ?? '',
      inventorySerial: matched?.serialNumber ?? '',
      physicalVirtual: matched?.physicalVirtual ?? '',
      roleWorkload: matched?.roleWorkload ?? '',
      location: matched?.location ?? '',
      manufacturer: matched?.manufacturer ?? '',
      model: matched?.model ?? '',
      osHypervisor: matched?.osHypervisor ?? '',
    }
  })

  const inventoryCoverage = serverInventoryPayload.inventory.map((server) => {
    const installed = matches.filter((match) => match.inventoryItemId === server.inventoryItemId)
    return {
      ...server,
      cyberSmartInstalled: installed.length > 0,
      cyberSmartInstallCount: installed.length,
      cyberSmartHosts: installed.map((match) => match.host).join('; '),
      cyberSmartStatuses: [...new Set(installed.map((match) => match.status))].join('; '),
      coverageStatus: installed.length ? 'CyberSmart matched' : 'No matched CyberSmart install',
    }
  })
  const siteNames = [...new Set([
    ...serverInventoryPayload.inventory.map((server) => server.school),
    ...matches.map((match) => match.site),
  ])].filter(Boolean).sort()
  const siteSummary = siteNames.map((site) => {
    const siteInventory = inventoryCoverage.filter((server) => server.school === site)
    const siteInstalls = matches.filter((match) => match.site === site)
    return {
      site,
      compatibleServersAndVms: siteInventory.length,
      physicalServers: siteInventory.filter((server) => server.physicalVirtual === 'PHYSICAL').length,
      virtualMachines: siteInventory.filter((server) => server.physicalVirtual === 'VIRTUAL').length,
      cyberSmartServerInstalls: siteInstalls.length,
      matchedInstalls: siteInstalls.filter((match) => match.matched).length,
      unmatchedInstalls: siteInstalls.filter((match) => !match.matched).length,
      passingInstalls: siteInstalls.filter((match) => match.status === 'Pass').length,
      failingInstalls: siteInstalls.filter((match) => match.status === 'Fail').length,
      inventoryWithCyberSmart: siteInventory.filter((server) => server.cyberSmartInstalled).length,
      inventoryWithoutCyberSmart: siteInventory.filter((server) => !server.cyberSmartInstalled).length,
    }
  })

  return {
    summary: {
      cyberSmartServerInstalls: cyberServers.length,
      inventoryRecords: serverInventoryPayload.records,
      physicalServers: serverInventoryPayload.physicalServers,
      virtualMachines: serverInventoryPayload.virtualMachines,
      matchedServers: matches.filter((match) => match.matched).length,
      unmatchedServers: matches.filter((match) => !match.matched).length,
      failingServers: matches.filter((match) => match.status === 'Fail').length,
      passingServers: matches.filter((match) => match.status === 'Pass').length,
      inventoryWithCyberSmart: inventoryCoverage.filter((server) => server.cyberSmartInstalled).length,
      inventoryWithoutCyberSmart: inventoryCoverage.filter((server) => !server.cyberSmartInstalled).length,
    },
    inventory: inventoryCoverage,
    matches,
    siteSummary,
  }
}

function writeJson(filename, value) {
  const file = path.join(outputDirectory, filename)
  fs.writeFileSync(file, JSON.stringify(value))
  return { file, bytes: fs.statSync(file).size }
}

async function main() {
  if (!fs.existsSync(sourceWorkbook)) {
    throw new Error(`Workbook not found: ${sourceWorkbook}`)
  }
  if (!fs.existsSync(sourceServerInventory)) {
    throw new Error(`Server inventory not found: ${sourceServerInventory}`)
  }

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(sourceWorkbook)

  const pendingStatus = rowsFrom(workbook, 'Pending Status')
  const sites = buildSites(rowsFrom(workbook, 'Site Navigator'))
  const installedDevices = rowsFrom(workbook, 'Installed Device Finder')
  const failingDevices = rowsFrom(workbook, 'Failing Device Finder')
  const remainingDevices = rowsFrom(workbook, 'Remaining Device Finder')
  const checks = rowsFrom(workbook, 'Site Check Breakdown')
  const failureCauses = rowsFrom(workbook, 'Site Failure Causes')
  const allocations = rowsFrom(workbook, 'Allocation by Site')
  const inventory = rowsFrom(workbook, 'Inventory Coverage')
  const cyberSmartDevices = rowsFrom(workbook, 'CyberSmart Devices')
  const reconciliation = rowsFrom(workbook, 'Reconciliation')
  const methodology = rowsFrom(workbook, 'Methodology')
  const serverInventoryPayload = JSON.parse(fs.readFileSync(sourceServerInventory, 'utf8'))
  const servers = buildServerMatches(cyberSmartDevices, serverInventoryPayload)
  const clientCompute = {
    records: inventory.length,
    quantity: inventory.reduce((total, row) => total + numberOf(row.quantity), 0),
    bySite: Object.values(
      inventory.reduce((result, row) => {
        const site = row.school || row.site || 'Unknown'
        const current = result[site] ?? { site, records: 0, quantity: 0, byCategory: {} }
        current.records += 1
        current.quantity += numberOf(row.quantity)
        const category = row.deviceCategory || 'Unknown'
        current.byCategory[category] = (current.byCategory[category] ?? 0) + numberOf(row.quantity)
        result[site] = current
        return result
      }, {}),
    ),
  }
  const reportSummary = {
    ...summaryFromPending(pendingStatus),
    schoolCyberSmartDevices: sites.reduce((sum, site) => sum + site.cyberSmartDevices, 0),
    centralCyberSmartDevices:
      cyberSmartDevices.length -
      sites.reduce((sum, site) => sum + site.cyberSmartDevices, 0),
    compatibleDevices: sites.reduce((sum, site) => sum + site.compatibleDevices, 0),
    serverInstalls: sites.reduce((sum, site) => sum + site.serverInstalls, 0),
    serialMatched: sites.reduce((sum, site) => sum + site.serialMatched, 0),
    notSerialMatched: sites.reduce((sum, site) => sum + site.notSerialMatched, 0),
    remainingToInstall: sites.reduce((sum, site) => sum + site.remainingToInstall, 0),
  }
  const schoolDueDates = [
    '2027-01-31',
    '2027-02-28',
    '2027-03-31',
    '2027-04-30',
    '2027-05-31',
    '2027-06-30',
    '2027-07-31',
  ]
  const sortedSites = [...sites].sort((a, b) => a.site.localeCompare(b.site))
  const schoolSchedule = sortedSites.map((site, index) => ({
    site: site.site,
    dueDate: schoolDueDates[Math.min(6, Math.floor(index * 7 / sortedSites.length))],
  }))
  const dueBySite = Object.fromEntries(schoolSchedule.map((item) => [item.site, item.dueDate]))
  const rolloutTypeLabels = {
    Windows: 'Windows',
    Chromebook: 'Chromebook',
    AppleMac: 'Apple / Mac',
    IPadIOs: 'iPad / iOS',
    AndroidTablet: 'Android / tablet',
    PossibleTypeUnknown: 'possible unknown / unmanaged',
  }
  const rolloutOwners = {
    'South Dartmoor Community College for WeST': 'Richard Penhale',
    'Callington Community College for WeST': 'David Mahoney',
    "Sir James Smith's School for WeST": 'David Mahoney',
    'Plymstock School for WeST': 'David Wainwright',
    'Coombe Dean School for WeST': 'David Wainwright',
    'Ivybridge Community College for WeST': 'Christian Rowlands',
    'Eggbuckland Community College': 'Adam Hedley',
    "Hele's School for WeST": 'David Lyns',
  }
  const programmeTasks = [
    {
      scope: 'WeST Central',
      site: 'WeST Central',
      category: 'Central devices',
      task: 'Reconcile and remediate all WeST Central CyberSmart devices',
      measure: `${reportSummary.centralCyberSmartDevices} detailed central devices`,
      owner: 'Mark Kent',
      dueDate: '2026-12-01',
      status: 'Not started',
    },
    {
      scope: 'Trust-wide',
      site: 'WeST Central',
      category: 'Portal reconciliation',
      task: 'Resolve CyberSmart devices pending without detailed export evidence',
      measure: `${reportSummary.pending} pending devices`,
      owner: 'Mark Kent',
      dueDate: '2026-12-01',
      status: 'Not started',
    },
    {
      scope: 'Trust-wide',
      site: 'WeST Central',
      category: 'Portal reconciliation',
      task: 'Reconcile portal Secure with all-exported-checks-pass reporting',
      measure: `${reportSummary.portalSecure - reportSummary.csvPassing} device difference`,
      owner: 'Mark Kent',
      dueDate: '2026-12-01',
      status: 'Not started',
    },
    {
      scope: 'Trust-wide',
      site: 'WeST Central',
      category: 'BYOD',
      task: 'Approve and publish a trust-wide BYOD policy',
      measure: 'Policy approved and communicated',
      owner: 'Mark Kent',
      dueDate: '2026-12-01',
      status: 'Not started',
    },
    {
      scope: 'Trust-wide',
      site: 'WeST Central',
      category: 'BYOD',
      task: 'Define and implement trust-wide BYOD technical controls',
      measure: 'Controls implemented and evidenced',
      owner: 'Mark Kent',
      dueDate: '2026-12-01',
      status: 'Not started',
    },
  ]
  for (const site of sortedSites) {
    const dueDate = dueBySite[site.site]
    const serverSite = servers.siteSummary.find((item) => item.site === site.site)
    for (const [deviceType, quantity] of Object.entries(site.remainingByType)) {
      if (quantity <= 0) continue
      const label = rolloutTypeLabels[deviceType] ?? deviceType
      programmeTasks.push({
        scope: 'School',
        site: site.site,
        category: `CyberSmart rollout - ${label}`,
        task: deviceType === 'PossibleTypeUnknown'
          ? 'Identify compatible devices and complete CyberSmart installation for possible unknown / unmanaged inventory'
          : `Install CyberSmart on remaining ${label} compatible devices`,
        measure: `${quantity} remaining ${label} devices`,
        owner: rolloutOwners[site.site] ?? 'Mark Kent',
        dueDate,
        status: 'Not started',
      })
    }
    if (site.failing > 0) {
      programmeTasks.push({
        scope: 'School',
        site: site.site,
        category: 'Compliance remediation',
        task: 'Remediate all CyberSmart failing client devices',
        measure: `${site.failing} failing devices`,
        owner: 'Mark Kent',
        dueDate,
        status: 'Not started',
      })
    }
    if (site.notSerialMatched > 0) {
      programmeTasks.push({
        scope: 'School',
        site: site.site,
        category: 'Inventory reconciliation',
        task: 'Reconcile CyberSmart installations not matched by serial',
        measure: `${site.notSerialMatched} unmatched installations`,
        owner: 'Mark Kent',
        dueDate,
        status: 'Not started',
      })
    }
    if ((serverSite?.inventoryWithoutCyberSmart ?? 0) > 0 || (serverSite?.unmatchedInstalls ?? 0) > 0) {
      programmeTasks.push({
        scope: 'School',
        site: site.site,
        category: 'Server CyberSmart',
        task: 'Reconcile server and VM CyberSmart coverage',
        measure: `${serverSite?.inventoryWithoutCyberSmart ?? 0} infrastructure records without match; ${serverSite?.unmatchedInstalls ?? 0} unmatched agents`,
        owner: 'Mark Kent',
        dueDate,
        status: 'Not started',
      })
    }
    programmeTasks.push({
      scope: 'School',
      site: site.site,
      category: 'MFA',
      task: 'Implement and evidence MFA for the site',
      measure: 'MFA implementation and evidence complete',
      owner: 'Mark Kent',
      dueDate,
      status: 'Not started',
    })
  }

  const core = {
    generatedAt: new Date().toISOString(),
    source: {
      workbook: path.basename(sourceWorkbook),
      notice:
        'Public full-detail report requested by the data owner. Contains device identifiers and security findings.',
    },
    summary: reportSummary,
    sites,
    checks,
    failureCauses,
    allocations,
    reconciliation,
    methodology,
    serverSummary: servers.summary,
    serverSites: servers.siteSummary,
    clientCompute,
    taskProgramme: {
      owner: 'Mark Kent',
      rolloutOwners,
      centralDueDate: '2026-12-01',
      schoolCompletionDeadline: '2027-07-31',
      schoolSchedule,
      tasks: programmeTasks,
    },
  }

  fs.mkdirSync(outputDirectory, { recursive: true })
  fs.mkdirSync(downloadDirectory, { recursive: true })
  const legacyReport = path.join(outputDirectory, 'report.json')
  if (fs.existsSync(legacyReport)) fs.unlinkSync(legacyReport)
  const outputs = [
    writeJson('core.json', core),
    writeJson('inventory.json', inventory),
    writeJson('cyber-devices.json', cyberSmartDevices),
    writeJson('installed-devices.json', installedDevices),
    writeJson('failing-devices.json', failingDevices),
    writeJson('remaining-devices.json', remainingDevices),
    writeJson('servers.json', servers),
  ]
  fs.copyFileSync(sourceWorkbook, downloadFile)

  console.log(
    JSON.stringify(
      {
        outputs,
        sites: sites.length,
        inventory: inventory.length,
        cyberSmartDevices: cyberSmartDevices.length,
        installedDevices: installedDevices.length,
        failingDevices: failingDevices.length,
        remainingRows: remainingDevices.length,
      },
      null,
      2,
    ),
  )
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
