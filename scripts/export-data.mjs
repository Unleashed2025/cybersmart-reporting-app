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
const downloadFile = path.join(downloadDirectory, 'WeST-CyberSmart-Assessment-and-Programme-Tracker.xlsx')
const colours = {
  navy: 'FF0B1736',
  blue: 'FF2F6FED',
  pink: 'FFFF2D7D',
  green: 'FF19A974',
  amber: 'FFF4A62A',
  red: 'FFE5484D',
  paleBlue: 'FFEAF1FF',
  paleAmber: 'FFFFF3DC',
  paleGreen: 'FFE6F7F1',
  white: 'FFFFFFFF',
  line: 'FFD9E1EC',
  text: 'FF14213D',
  muted: 'FF64748B',
}

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

function buildCentralReport(centralDevices, inventory) {
  const morleyInventory = inventory.filter((row) => row.school === 'Morley Meadow')
  const serialIndex = new Map()
  const assetIndex = new Map()
  for (const row of morleyInventory) {
    if (validIdentity(row.serialNumber)) serialIndex.set(identity(row.serialNumber), row)
    if (validIdentity(row.assetTag)) assetIndex.set(identity(row.assetTag), row)
  }

  const devices = centralDevices.map((device) => {
    const serialKey = validIdentity(device.serial) ? identity(device.serial) : ''
    const hostKey = identity(device.host)
    let match = serialKey ? serialIndex.get(serialKey) : undefined
    let method = match ? 'Exact serial to Morley inventory' : ''
    if (!match && serialKey) {
      match = assetIndex.get(serialKey)
      if (match) method = 'CyberSmart serial to Morley asset tag'
    }
    if (!match) {
      match = morleyInventory.find((row) =>
        [row.serialNumber, row.assetTag].some((value) => {
          const key = identity(value)
          return key.length >= 5 && hostKey.includes(key)
        }))
      if (match) method = 'CyberSmart hostname contains Morley identifier'
    }
    return {
      ...device,
      morleyInventoryMatch: match ? 'Matched' : 'No match',
      morleyMatchMethod: method || 'No Morley inventory match',
      morleyInventoryCategory: match?.deviceCategory ?? '',
      morleyInventorySerial: match?.serialNumber ?? '',
      morleyAssetTag: match?.assetTag ?? '',
      morleyInventoryModel: match?.model ?? '',
    }
  })

  const group = (key) => Object.fromEntries(
    Object.entries(devices.reduce((result, row) => {
      const value = row[key] || 'Not recorded'
      result[value] = (result[value] ?? 0) + 1
      return result
    }, {})).sort((a, b) => b[1] - a[1]),
  )
  const failedControls = {}
  for (const row of devices) {
    for (const control of String(row.failedChecks ?? '').split(';').map((value) => value.trim()).filter(Boolean)) {
      failedControls[control] = (failedControls[control] ?? 0) + 1
    }
  }
  const matchedToMorley = devices.filter((row) => row.morleyInventoryMatch === 'Matched').length
  return {
    summary: {
      reportSite: 'West Country Schools Trust / Central',
      associatedSite: 'Morley Meadow',
      inventoryAvailable: false,
      totalAgents: devices.length,
      passingAgents: devices.filter((row) => row.overallStatus === 'Pass').length,
      failingAgents: devices.filter((row) => row.overallStatus === 'Fail').length,
      serialRecorded: devices.filter((row) => Boolean(row.serial)).length,
      morleyInventoryRecords: morleyInventory.length,
      morleyInventoryQuantity: morleyInventory.reduce((total, row) => total + numberOf(row.quantity), 0),
      matchedToMorley,
      unmatchedToMorley: devices.length - matchedToMorley,
      deviceTypes: group('deviceType'),
      failedControls: Object.fromEntries(Object.entries(failedControls).sort((a, b) => b[1] - a[1])),
    },
    devices,
  }
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
    const globalHost = hostKey
      ? inventory.filter((server) => server.hostKey === hostKey)
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
    } else if (globalHost.length === 1) {
      matched = globalHost[0]
      method = 'Exact hostname - different site'
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

function excelDate(value) {
  return value ? new Date(`${value}T00:00:00Z`) : ''
}

function prepareTrackingSheet(workbook, name, title, description, headers, widths, tabColour) {
  const existing = workbook.getWorksheet(name)
  if (existing) workbook.removeWorksheet(existing.id)
  const sheet = workbook.addWorksheet(name, {
    properties: { tabColor: { argb: tabColour } },
    views: [{ state: 'frozen', ySplit: 4 }],
  })
  sheet.mergeCells(1, 1, 1, headers.length)
  sheet.getCell(1, 1).value = title
  sheet.getCell(1, 1).font = { name: 'Aptos Display', size: 20, bold: true, color: { argb: colours.white } }
  sheet.getCell(1, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colours.navy } }
  sheet.getCell(1, 1).alignment = { vertical: 'middle' }
  sheet.getRow(1).height = 34

  sheet.mergeCells(2, 1, 2, headers.length)
  sheet.getCell(2, 1).value = description
  sheet.getCell(2, 1).font = { name: 'Aptos', size: 10, color: { argb: colours.muted } }
  sheet.getCell(2, 1).alignment = { vertical: 'middle', wrapText: true }
  sheet.getRow(2).height = 30

  sheet.mergeCells(3, 1, 3, headers.length)
  sheet.getCell(3, 1).value = 'Yellow cells are editable progress fields. Filters are enabled on every column.'
  sheet.getCell(3, 1).font = { name: 'Aptos', size: 9, italic: true, color: { argb: colours.text } }
  sheet.getCell(3, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colours.paleAmber } }
  sheet.getRow(3).height = 22

  const headerRow = sheet.getRow(4)
  headers.forEach((header, index) => {
    const cell = headerRow.getCell(index + 1)
    cell.value = header
    cell.font = { name: 'Aptos', size: 9, bold: true, color: { argb: colours.white } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colours.blue } }
    cell.alignment = { vertical: 'middle', wrapText: true }
    cell.border = { bottom: { style: 'thin', color: { argb: colours.line } } }
    sheet.getColumn(index + 1).width = widths[index] ?? 16
  })
  headerRow.height = 30
  return sheet
}

function styleTrackingRows(sheet, startRow, endRow, editableColumns = []) {
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber)
    row.font = { name: 'Aptos', size: 9, color: { argb: colours.text } }
    row.alignment = { vertical: 'top', wrapText: true }
    row.height = 27
    row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
      cell.border = { bottom: { style: 'hair', color: { argb: colours.line } } }
      if (editableColumns.includes(columnNumber)) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colours.paleAmber } }
      } else if (rowNumber % 2 === 0) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFD' } }
      }
    })
  }
}

function addStatusValidation(sheet, column, startRow, endRow) {
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    sheet.getCell(rowNumber, column).dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: ['"Not started,Assigned,In progress,Blocked,Complete"'],
    }
  }
}

function addProgressValidation(sheet, column, startRow, endRow) {
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const cell = sheet.getCell(rowNumber, column)
    cell.dataValidation = {
      type: 'whole',
      operator: 'between',
      allowBlank: false,
      formulae: [0, 100],
    }
    cell.numFmt = '0"%"'
  }
}

function addProgrammeTrackingSheets(workbook, core, remainingDevices) {
  workbook.calcProperties.fullCalcOnLoad = true
  workbook.calcProperties.forceFullCalc = true

  const tasks = core.taskProgramme.tasks
  const taskHeaders = [
    'Task ID', 'Scope', 'Site', 'Category', 'Task', 'Completion measure', 'Owner', 'Due date',
    'Status', 'Progress %', 'Actual completion date', 'RAG', 'Blocker / dependency',
    'Latest update / evidence', 'Last updated',
  ]
  const taskSheet = prepareTrackingSheet(
    workbook,
    'Programme Tracker',
    'CyberSmart Programme Tracker',
    'The complete app task allocation with editable status, progress, completion, blocker and evidence fields.',
    taskHeaders,
    [12, 14, 31, 25, 50, 34, 20, 14, 16, 12, 18, 12, 28, 42, 16],
    colours.pink,
  )
  tasks.forEach((task, index) => {
    const rowNumber = index + 5
    taskSheet.addRow([
      `CS-${String(index + 1).padStart(3, '0')}`,
      task.scope,
      task.site,
      task.category,
      task.task,
      task.measure,
      task.owner,
      excelDate(task.dueDate),
      task.status,
      0,
      '',
      { formula: `IF(I${rowNumber}="Complete","Green",IF(I${rowNumber}="Blocked","Red",IF(H${rowNumber}<TODAY(),"Red",IF(H${rowNumber}-TODAY()<=30,"Amber","Green"))))` },
      '',
      '',
      '',
    ])
    taskSheet.getCell(rowNumber, 8).numFmt = 'dd mmm yyyy'
    taskSheet.getCell(rowNumber, 11).numFmt = 'dd mmm yyyy'
    taskSheet.getCell(rowNumber, 15).numFmt = 'dd mmm yyyy'
  })
  const taskEndRow = tasks.length + 4
  taskSheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: taskEndRow, column: taskHeaders.length } }
  styleTrackingRows(taskSheet, 5, taskEndRow, [9, 10, 11, 13, 14, 15])
  addStatusValidation(taskSheet, 9, 5, taskEndRow)
  addProgressValidation(taskSheet, 10, 5, taskEndRow)

  const schoolHeaders = [
    'Site', 'Remaining-install owner', 'Deadline', 'Compatible clients', 'CyberSmart client installs',
    'Serial matched', 'Without serial match', 'Remaining clients', 'Known remaining', 'Explicit unmanaged remaining',
    'Unallocated / uncertain gap', 'Install tasks', 'Compliance tasks', 'Inventory tasks', 'Server tasks',
    'MFA tasks', 'Total tasks', 'Confirmed server installs', 'Remaining server installs',
    'Average progress %', 'Overall status', 'RAG', 'Latest update',
  ]
  const schoolSheet = prepareTrackingSheet(
    workbook,
    'School Progress',
    'School CyberSmart Progress',
    'One row per school, combining coverage, assigned workload, deadlines and formula-driven progress from the Programme Tracker.',
    schoolHeaders,
    [31, 23, 14, 15, 20, 14, 18, 16, 16, 20, 20, 13, 15, 14, 12, 11, 12, 18, 18, 17, 16, 11, 36],
    colours.blue,
  )
  core.sites.forEach((site, index) => {
    const rowNumber = index + 5
    const schoolTasks = tasks.filter((task) => task.site === site.site)
    const server = core.serverSites.find((item) => item.site === site.site)
    const schedule = core.taskProgramme.schoolSchedule.find((item) => item.site === site.site)
    const clientInstalls = Math.max(0, site.cyberSmartDevices - site.serverInstalls)
    const unmatchedClients = Math.max(0, site.notSerialMatched - site.serverInstalls)
    schoolSheet.addRow([
      site.site,
      core.taskProgramme.rolloutOwners[site.site] ?? core.taskProgramme.owner,
      excelDate(schedule?.dueDate),
      site.compatibleDevices,
      clientInstalls,
      site.serialMatched,
      unmatchedClients,
      site.remainingToInstall,
      site.knownRemaining,
      site.explicitUnmanagedRemaining,
      site.unallocatedInventoryGap,
      schoolTasks.filter((task) => task.category.startsWith('CyberSmart rollout')).length,
      schoolTasks.filter((task) => task.category === 'Compliance remediation').length,
      schoolTasks.filter((task) => task.category === 'Inventory reconciliation').length,
      schoolTasks.filter((task) => task.category === 'Server CyberSmart').length,
      schoolTasks.filter((task) => task.category === 'MFA').length,
      schoolTasks.length,
      server?.inventoryWithCyberSmart ?? 0,
      server?.inventoryWithoutCyberSmart ?? 0,
      { formula: `IFERROR(AVERAGEIF('Programme Tracker'!$C:$C,A${rowNumber},'Programme Tracker'!$J:$J),0)` },
      { formula: `IF(T${rowNumber}>=100,"Complete",IF(T${rowNumber}>0,"In progress","Not started"))` },
      { formula: `IF(U${rowNumber}="Complete","Green",IF(C${rowNumber}<TODAY(),"Red",IF(C${rowNumber}-TODAY()<=30,"Amber","Green")))` },
      '',
    ])
    schoolSheet.getCell(rowNumber, 3).numFmt = 'dd mmm yyyy'
    schoolSheet.getCell(rowNumber, 20).numFmt = '0"%"'
  })
  const schoolEndRow = core.sites.length + 4
  schoolSheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: schoolEndRow, column: schoolHeaders.length } }
  styleTrackingRows(schoolSheet, 5, schoolEndRow, [23])

  const rolloutTasks = tasks.filter((task) => task.category.startsWith('CyberSmart rollout'))
  const rolloutHeaders = [
    'Task ID', 'Site', 'Device type', 'Remaining quantity', 'Owner', 'Due date', 'Status',
    'Progress %', 'Actual completion date', 'Blocker / dependency', 'Latest update / evidence', 'Last updated',
  ]
  const rolloutSheet = prepareTrackingSheet(
    workbook,
    'Device Rollout Tracker',
    'CyberSmart Device Rollout Tracker',
    'A concise site and device-type view of the remaining compatible-device installation work.',
    rolloutHeaders,
    [12, 31, 22, 18, 20, 14, 16, 12, 18, 28, 42, 16],
    colours.amber,
  )
  rolloutTasks.forEach((task, index) => {
    const sourceIndex = tasks.indexOf(task)
    const quantity = Number(task.measure.match(/^\d+/)?.[0] ?? 0)
    rolloutSheet.addRow([
      `CS-${String(sourceIndex + 1).padStart(3, '0')}`,
      task.site,
      task.category.replace('CyberSmart rollout - ', ''),
      quantity,
      task.owner,
      excelDate(task.dueDate),
      task.status,
      0,
      '',
      '',
      '',
      '',
    ])
    const rowNumber = index + 5
    rolloutSheet.getCell(rowNumber, 6).numFmt = 'dd mmm yyyy'
    rolloutSheet.getCell(rowNumber, 9).numFmt = 'dd mmm yyyy'
    rolloutSheet.getCell(rowNumber, 12).numFmt = 'dd mmm yyyy'
  })
  const rolloutEndRow = rolloutTasks.length + 4
  rolloutSheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: rolloutEndRow, column: rolloutHeaders.length } }
  styleTrackingRows(rolloutSheet, 5, rolloutEndRow, [7, 8, 9, 10, 11, 12])
  addStatusValidation(rolloutSheet, 7, 5, rolloutEndRow)
  addProgressValidation(rolloutSheet, 8, 5, rolloutEndRow)

  return {
    programmeTasks: tasks.length,
    schoolRows: core.sites.length,
    rolloutTasks: rolloutTasks.length,
    remainingEvidenceRows: remainingDevices.length,
  }
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
  const remainingEvidenceBySite = remainingDevices.reduce((result, row) => {
    const site = row.schoolSite
    const current = result[site] ?? { explicitUnmanagedRemaining: 0, unallocatedInventoryGap: 0 }
    const quantity = numberOf(row.allocateQuantity)
    const unallocated = row.deviceCategory === 'Unallocated inventory gap' ||
      String(row.evidenceConfidence).includes('Gap exceeds identifiable unmatched inventory rows')
    const unmanaged = !unallocated && (
      String(row.unmanaged).toLowerCase() === 'yes' ||
      String(row.deviceCategory).toLowerCase().includes('unmanaged') ||
      String(row.deviceType).toLowerCase().includes('unknown')
    )
    if (unallocated) current.unallocatedInventoryGap += quantity
    else if (unmanaged) current.explicitUnmanagedRemaining += quantity
    result[site] = current
    return result
  }, {})
  sites.forEach((site) => {
    const evidence = remainingEvidenceBySite[site.site] ?? { explicitUnmanagedRemaining: 0, unallocatedInventoryGap: 0 }
    site.explicitUnmanagedRemaining = evidence.explicitUnmanagedRemaining
    site.unallocatedInventoryGap = evidence.unallocatedInventoryGap
    site.knownRemaining = Math.max(0, site.remainingToInstall - evidence.explicitUnmanagedRemaining - evidence.unallocatedInventoryGap)
  })
  const checks = rowsFrom(workbook, 'Site Check Breakdown')
  const failureCauses = rowsFrom(workbook, 'Site Failure Causes')
  const allocations = rowsFrom(workbook, 'Allocation by Site')
  const inventory = rowsFrom(workbook, 'Inventory Coverage')
  const cyberSmartDevices = rowsFrom(workbook, 'CyberSmart Devices')
  const centralCyberSmartDevices = cyberSmartDevices.filter(
    (device) => device.reportSite === 'West Country Schools Trust / Central',
  )
  const centralReport = buildCentralReport(centralCyberSmartDevices, inventory)
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
    centralCyberSmartDevices: centralCyberSmartDevices.length,
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
    centralReport: centralReport.summary,
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
    writeJson('tracker.json', {
      taskProgramme: core.taskProgramme,
      sites: core.sites.map((site) => ({
        site: site.site,
        compatibleDevices: site.compatibleDevices,
        cyberSmartDevices: site.cyberSmartDevices,
        serverInstalls: site.serverInstalls,
        serialMatched: site.serialMatched,
        notSerialMatched: site.notSerialMatched,
        remainingToInstall: site.remainingToInstall,
        knownRemaining: site.knownRemaining,
        explicitUnmanagedRemaining: site.explicitUnmanagedRemaining,
        unallocatedInventoryGap: site.unallocatedInventoryGap,
      })),
      serverSites: core.serverSites,
      centralReport: core.centralReport,
    }),
    writeJson('inventory.json', inventory),
    writeJson('cyber-devices.json', cyberSmartDevices),
    writeJson('central-devices.json', centralReport.devices),
    writeJson('installed-devices.json', installedDevices),
    writeJson('failing-devices.json', failingDevices),
    writeJson('remaining-devices.json', remainingDevices),
    writeJson('servers.json', servers),
  ]
  let workbookTracking = {
    file: downloadFile,
    generated: false,
    note: 'Generate the enhanced workbook with npm run spreadsheet:build on Windows to preserve the original Excel workbook exactly.',
  }
  if (process.env.CYBERSMART_EXCELJS_TRACKER === '1') {
    workbookTracking = {
      file: downloadFile,
      generated: true,
      ...addProgrammeTrackingSheets(workbook, core, remainingDevices),
    }
    await workbook.xlsx.writeFile(downloadFile)
  }

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
        workbook: {
          bytes: fs.existsSync(downloadFile) ? fs.statSync(downloadFile).size : 0,
          ...workbookTracking,
        },
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
