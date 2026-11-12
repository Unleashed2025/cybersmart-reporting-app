export type Page =
  | 'overview'
  | 'programme'
  | 'sites'
  | 'inventory'
  | 'installed'
  | 'failures'
  | 'remaining'
  | 'servers'
  | 'checks'
  | 'guide'

export interface Summary {
  portalTotal: number
  detailedDevices: number
  pending: number
  inventoryInstallationGap: number
  portalSecure: number
  csvPassing: number
  csvFailing: number
  recentlyActive: number
  schoolCyberSmartDevices: number
  centralCyberSmartDevices: number
  compatibleDevices: number
  serverInstalls: number
  serialMatched: number
  notSerialMatched: number
  remainingToInstall: number
}

export interface Site {
  site: string
  compatibleDevices: number
  cyberSmartDevices: number
  serverInstalls: number
  serialMatched: number
  notSerialMatched: number
  remainingToInstall: number
  passing: number
  failing: number
  passPercent: number
  possibleUnmanaged: number
  priority: string
  passChecks: Record<string, number>
  failChecks: Record<string, number>
  remainingByType: Record<string, number>
}

export interface CheckRow extends Record<string, unknown> {
  schoolSite: string
  cyberSmartCheck: string
  devicesAssessed: number
  passingDevices: number
  failingDevices: number
  unknownDevices: number
  checkPassPercent: number
  interpretation: string
}

export interface CoreData {
  generatedAt: string
  sourceWorkbook: string
  summary: Summary
  sites: Site[]
  checks: CheckRow[]
  failureCauses: Array<Record<string, unknown>>
  allocations: Array<Record<string, unknown>>
  reconciliation: Array<Record<string, unknown>>
  methodology: Array<Record<string, unknown>>
  serverSummary: ServerSummary
  serverSites: ServerSiteSummary[]
  clientCompute: {
    records: number
    quantity: number
    bySite: Array<{
      site: string
      records: number
      quantity: number
      byCategory: Record<string, number>
    }>
  }
  taskProgramme: {
    owner: string
    rolloutOwners: Record<string, string>
    centralDueDate: string
    schoolCompletionDeadline: string
    schoolSchedule: Array<{ site: string; dueDate: string }>
    tasks: ProgrammeTask[]
  }
}

export interface ProgrammeTask extends Record<string, unknown> {
  scope: string
  site: string
  category: string
  task: string
  measure: string
  owner: string
  dueDate: string
  status: string
}

export interface InventoryRow extends Record<string, unknown> {
  school: string
  site: string
  deviceCategory: string
  compatibleType: string
  quantity: number
  possibleInstall: string
  exactCyberSmartMatch: string
  matchedCyberSmartStatus: string
  serialNumber: string
  assetTag: string
  manufacturer: string
  model: string
  operatingSystem: string
  location: string
  inventoryId: string
}

export interface InstalledRow extends Record<string, unknown> {
  allocationCategory: string
  schoolSite: string
  host: string
  serial: string
  deviceType: string
  overallStatus: string
  os: string
  failedChecks: string
  matchMethod: string
  cyberSmartDeviceId: string
}

export interface FailingRow extends Record<string, unknown> {
  schoolSite: string
  host: string
  serial: string
  deviceType: string
  os: string
  failedChecks: string
  failureReasons: string
  lastCheckIn: string
  matchMethod: string
  cyberSmartDeviceId: string
}

export interface RemainingRow extends Record<string, unknown> {
  allocationCategory: string
  schoolSite: string
  deviceCategory: string
  deviceType: string
  allocateQuantity: number
  inventoryQuantity: number
  serial: string
  assetTag: string
  operatingSystem: string
  manufacturer: string
  model: string
  location: string
  evidenceConfidence: string
  allocationStatus: string
}

export interface ServerMatch extends Record<string, unknown> {
  site: string
  host: string
  serial: string
  status: string
  os: string
  failedChecks: string
  matchMethod: string
  matched: boolean
  inventorySchool: string
  inventoryServerName: string
  inventorySerial: string
  physicalVirtual: string
  roleWorkload: string
  location: string
}

export interface ServerInventory extends Record<string, unknown> {
  school: string
  serverName: string
  serialNumber: string
  physicalVirtual: string
  roleWorkload: string
  location: string
  manufacturer: string
  model: string
  osHypervisor: string
  inventoryItemId: string
  cyberSmartInstalled: boolean
  cyberSmartInstallCount: number
  cyberSmartHosts: string
  cyberSmartStatuses: string
  coverageStatus: string
}

export interface ServerSummary {
  cyberSmartServerInstalls: number
  inventoryRecords: number
  physicalServers: number
  virtualMachines: number
  matchedServers: number
  unmatchedServers: number
  failingServers: number
  passingServers: number
  inventoryWithCyberSmart: number
  inventoryWithoutCyberSmart: number
}

export interface ServerData {
  summary: ServerSummary
  inventory: ServerInventory[]
  matches: ServerMatch[]
  siteSummary: ServerSiteSummary[]
}

export interface ServerSiteSummary extends Record<string, unknown> {
  site: string
  compatibleServersAndVms: number
  physicalServers: number
  virtualMachines: number
  cyberSmartServerInstalls: number
  matchedInstalls: number
  unmatchedInstalls: number
  passingInstalls: number
  failingInstalls: number
  inventoryWithCyberSmart: number
  inventoryWithoutCyberSmart: number
}

export interface Remediation {
  owner: string
  priority: string
  status: string
  dueDate: string
  notes: string
}
