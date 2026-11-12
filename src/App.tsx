import { useEffect, useMemo, useState } from 'react'
import type { ReactNode } from 'react'
import {
  AlertTriangle, BarChart3, BookOpen, CheckCircle2, ChevronRight, CircleGauge,
  Download, Laptop, ListChecks, Menu, MonitorCheck, Search, Server,
  ShieldAlert, ShieldCheck, X,
} from 'lucide-react'
import {
  Bar, BarChart, CartesianGrid, Legend, ResponsiveContainer, Tooltip, XAxis, YAxis,
} from 'recharts'
import './App.css'
import type {
  CheckRow, CoreData, FailingRow, InstalledRow, InventoryRow, Page, Remediation,
  RemainingRow, ServerData, ServerInventory, ServerMatch, ServerSiteSummary,
} from './types'

const nav: Array<{ id: Page; label: string; icon: typeof CircleGauge }> = [
  { id: 'overview', label: 'Executive overview', icon: CircleGauge },
  { id: 'sites', label: 'Site explorer', icon: BarChart3 },
  { id: 'inventory', label: 'Compute inventory', icon: Laptop },
  { id: 'installed', label: 'CyberSmart installs', icon: MonitorCheck },
  { id: 'failures', label: 'Failures & allocation', icon: ShieldAlert },
  { id: 'remaining', label: 'Remaining installs', icon: AlertTriangle },
  { id: 'servers', label: 'Server CyberSmart report', icon: Server },
  { id: 'checks', label: 'Control performance', icon: ListChecks },
  { id: 'guide', label: 'Guide & methodology', icon: BookOpen },
]
const formatNumber = new Intl.NumberFormat('en-GB')
const formatPercent = new Intl.NumberFormat('en-GB', { maximumFractionDigits: 1 })
const dataUrl = (file: string) => `${import.meta.env.BASE_URL}data/${file}`
const display = (value: unknown) => value === null || value === undefined || value === '' ? '—' : typeof value === 'boolean' ? value ? 'Yes' : 'No' : String(value)

function useDataset<T>(file: string, enabled: boolean) {
  const [data, setData] = useState<T | null>(null)
  const [error, setError] = useState('')
  useEffect(() => {
    if (!enabled || data) return
    fetch(dataUrl(file))
      .then((response) => {
        if (!response.ok) throw new Error(`${response.status} ${response.statusText}`)
        return response.json() as Promise<T>
      })
      .then(setData)
      .catch((reason: unknown) => setError(reason instanceof Error ? reason.message : String(reason)))
  }, [data, enabled, file])
  return { data, error, loading: enabled && !data && !error }
}

function Loader({ error = '' }: { error?: string }) {
  return error
    ? <div className="error">Unable to load report data: {error}</div>
    : <div className="loader"><i />Loading detailed reporting data…</div>
}

function Header({ title, description }: { title: string; description: string }) {
  return (
    <div className="page-heading">
      <div>
        <p className="eyebrow">CyberSmart coverage assessment</p>
        <h1>{title}</h1>
        <p>{description}</p>
      </div>
      <a className="primary" href={`${import.meta.env.BASE_URL}downloads/WeST-CyberSmart-Final-Validated-Assessment.xlsx`}>
        <Download size={17} /> Download full workbook
      </a>
    </div>
  )
}

function Metric({ label, value, note, tone = '' }: { label: string; value: number | string; note?: string; tone?: string }) {
  return (
    <article className={`metric ${tone}`}>
      <span>{label}</span>
      <strong>{typeof value === 'number' ? formatNumber.format(value) : value}</strong>
      {note && <small>{note}</small>}
    </article>
  )
}

function Badge({ value }: { value: string }) {
  const lower = value.toLowerCase()
  const tone = lower.includes('pass') || lower.includes('exact') || lower.includes('matched') && !lower.includes('unmatched')
    ? 'good'
    : lower.includes('fail') || lower.includes('unmatched') || lower.includes('no exact')
      ? 'bad' : 'neutral'
  return <span className={`badge ${tone}`}>{value || 'Not recorded'}</span>
}

interface Column<T> { key: keyof T & string; label: string; render?: (row: T) => ReactNode }

function DataTable<T extends Record<string, unknown>>({ rows, columns }: { rows: T[]; columns: Array<Column<T>> }) {
  const [page, setPage] = useState(0)
  const size = 100
  const maxPage = Math.max(0, Math.ceil(rows.length / size) - 1)
  const current = Math.min(page, maxPage)
  if (!rows.length) return <div className="empty">No records match the selected filters.</div>
  return (
    <>
      <div className="table-wrap">
        <table>
          <thead><tr>{columns.map((column) => <th key={column.key}>{column.label}</th>)}</tr></thead>
          <tbody>
            {rows.slice(current * size, current * size + size).map((row, index) => (
              <tr key={`${current}-${index}`}>
                {columns.map((column) => <td key={column.key}>{column.render ? column.render(row) : display(row[column.key])}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="pagination">
        <span>{formatNumber.format(current * size + 1)}–{formatNumber.format(Math.min(rows.length, current * size + size))} of {formatNumber.format(rows.length)}</span>
        <div>
          <button type="button" disabled={current === 0} onClick={() => setPage(current - 1)}>Previous</button>
          <button type="button" disabled={current >= maxPage} onClick={() => setPage(current + 1)}>Next</button>
        </div>
      </div>
    </>
  )
}

function exportCsv(filename: string, rows: Array<Record<string, unknown>>) {
  if (!rows.length) return
  const headers = Object.keys(rows[0])
  const quote = (value: unknown) => `"${String(value ?? '').replaceAll('"', '""')}"`
  const csv = [headers.map(quote).join(','), ...rows.map((row) => headers.map((header) => quote(row[header])).join(','))].join('\r\n')
  const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }))
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
}

function matchesSearch(row: Record<string, unknown>, search: string) {
  const needle = search.toLowerCase()
  return !needle || Object.values(row).some((value) => String(value ?? '').toLowerCase().includes(needle))
}

function FilteredTable<T extends Record<string, unknown>>({
  rows, columns, siteKey, typeKey, filename,
}: {
  rows: T[]; columns: Array<Column<T>>; siteKey: keyof T & string; typeKey: keyof T & string; filename: string
}) {
  const [search, setSearch] = useState('')
  const [site, setSite] = useState('')
  const [type, setType] = useState('')
  const sites = useMemo(() => [...new Set(rows.map((row) => display(row[siteKey])).filter((value) => value !== '—'))].sort(), [rows, siteKey])
  const types = useMemo(() => [...new Set(rows.map((row) => display(row[typeKey])).filter((value) => value !== '—'))].sort(), [rows, typeKey])
  const filtered = useMemo(() => rows.filter((row) =>
    (!site || display(row[siteKey]) === site) &&
    (!type || display(row[typeKey]) === type) &&
    matchesSearch(row, search)), [rows, search, site, siteKey, type, typeKey])
  return (
    <section className="panel">
      <div className="filters">
        <label className="search"><Search size={17} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search hostname, serial, model, asset tag…" /></label>
        <select value={site} onChange={(event) => setSite(event.target.value)}><option value="">All sites</option>{sites.map((value) => <option key={value}>{value}</option>)}</select>
        <select value={type} onChange={(event) => setType(event.target.value)}><option value="">All device types</option>{types.map((value) => <option key={value}>{value}</option>)}</select>
        <button className="secondary" type="button" onClick={() => exportCsv(filename, filtered)}><Download size={16} /> Export filtered CSV</button>
      </div>
      <DataTable rows={filtered} columns={columns} />
    </section>
  )
}

function Overview({ core, navigate }: { core: CoreData; navigate: (page: Page) => void }) {
  const s = core.summary
  const chart = [...core.sites].sort((a, b) => b.compatibleDevices - a.compatibleDevices).slice(0, 10).map((site) => ({
    name: site.site.replace(/ School| Academy| College/g, ''),
    'Installed clients': Math.max(0, site.cyberSmartDevices - site.serverInstalls),
    'Remaining clients': site.remainingToInstall,
    Servers: site.serverInstalls,
  }))
  return (
    <>
      <Header title="Trust-wide inventory and coverage" description="All reporting starts with the authoritative compatible client-compute estate, then reconciles CyberSmart installation, serial matching, compliance and remaining rollout." />
      <h2 className="section-title">Inventory and installation headlines</h2>
      <div className="metrics">
        <Metric label="All client-compute items" value={core.clientCompute.quantity} note={`${formatNumber.format(core.clientCompute.records)} source records, including non-compatible categories`} tone="slate" />
        <Metric label="Compatible client devices" value={s.compatibleDevices} note="Authoritative client denominator" />
        <Metric label="School CyberSmart devices" value={s.schoolCyberSmartDevices} note="Includes separately reported servers" />
        <Metric label="Serial-matched installs" value={s.serialMatched} note="Passing and failing devices" tone="green" />
        <Metric label="Not serial matched" value={s.notSerialMatched} note="Includes server installs" tone="amber" />
        <Metric label="Remaining client installs" value={s.remainingToInstall} note="Provisional inventory allocation" tone="red" />
        <Metric label="CyberSmart server installs" value={s.serverInstalls} note="Separate from client compute" tone="purple" />
        <Metric label="Compatible servers & VMs" value={core.serverSummary.inventoryRecords} note={`${formatNumber.format(core.serverSummary.physicalServers)} physical · ${formatNumber.format(core.serverSummary.virtualMachines)} virtual`} tone="purple" />
      </div>
      <h2 className="section-title">Compliance and portal headlines</h2>
      <div className="metrics">
        <Metric label="Portal total devices" value={s.portalTotal} />
        <Metric label="Detailed CSV devices" value={s.detailedDevices} />
        <Metric label="Pending devices" value={s.pending} tone="amber" />
        <Metric label="Portal secure" value={s.portalSecure} tone="green" />
        <Metric label="CSV passing" value={s.csvPassing} tone="green" />
        <Metric label="CSV failing" value={s.csvFailing} tone="red" />
        <Metric label="Recently active" value={s.recentlyActive} />
        <Metric label="Central / Trust devices" value={s.centralCyberSmartDevices} tone="slate" />
        <Metric label="Inventory gap" value={s.inventoryInstallationGap} tone="red" />
      </div>
      <section className="panel">
        <h2>Three-category reconciliation</h2>
        <p>Server installations are visible within unmatched totals but never reduce the client-compute installation gap.</p>
        <div className="flow">
          <button onClick={() => navigate('inventory')}><strong>{formatNumber.format(s.compatibleDevices)}</strong><span>Compatible client devices</span></button><ChevronRight />
          <button onClick={() => navigate('installed')}><strong>{formatNumber.format(s.serialMatched)}</strong><span>Installed + serial matched</span></button><ChevronRight />
          <button onClick={() => navigate('installed')}><strong>{formatNumber.format(s.notSerialMatched)}</strong><span>Installed, no serial match</span></button><ChevronRight />
          <button onClick={() => navigate('remaining')}><strong>{formatNumber.format(s.remainingToInstall)}</strong><span>Remaining allocation</span></button>
        </div>
      </section>
      <div className="two-column">
        <section className="panel">
          <h2>Largest compatible client estates</h2><p>Installed client agents, remaining allocation and separately reported server agents.</p>
          <ResponsiveContainer width="100%" height={420}>
            <BarChart data={chart} layout="vertical" margin={{ left: 18, right: 12, top: 20 }}>
              <CartesianGrid strokeDasharray="3 3" horizontal={false} /><XAxis type="number" />
              <YAxis dataKey="name" type="category" width={150} tick={{ fontSize: 11 }} />
              <Tooltip /><Legend /><Bar dataKey="Installed clients" stackId="a" fill="#2878ff" />
              <Bar dataKey="Remaining clients" stackId="a" fill="#f3a83b" /><Bar dataKey="Servers" fill="#7d66d9" />
            </BarChart>
          </ResponsiveContainer>
        </section>
        <section className="panel">
          <h2>Portal reconciliation</h2><p>Portal headlines and detailed export evidence are not treated as identical measures.</p>
          <div className="breakdown">
            <div><span>Portal total</span><strong>{formatNumber.format(s.portalTotal)}</strong></div>
            <div><span>Detailed device rows</span><strong>{formatNumber.format(s.detailedDevices)}</strong></div>
            <div><span>Pending</span><strong>{formatNumber.format(s.pending)}</strong></div>
            <div><span>Portal secure</span><strong>{formatNumber.format(s.portalSecure)}</strong></div>
            <div><span>All exported checks pass</span><strong>{formatNumber.format(s.csvPassing)}</strong></div>
          </div>
          <section className="panel">
            <h2>Server CyberSmart coverage by site</h2>
            <p>Compatible infrastructure inventory and CyberSmart server-agent evidence remain separate from client-compute coverage.</p>
            <DataTable<ServerSiteSummary> rows={core.serverSites} columns={[
              { key: 'site', label: 'Site' },
              { key: 'compatibleServersAndVms', label: 'Compatible servers & VMs' },
              { key: 'physicalServers', label: 'Physical' },
              { key: 'virtualMachines', label: 'Virtual' },
              { key: 'cyberSmartServerInstalls', label: 'CyberSmart installs' },
              { key: 'matchedInstalls', label: 'Matched installs' },
              { key: 'unmatchedInstalls', label: 'Unmatched installs' },
              { key: 'passingInstalls', label: 'Passing' },
              { key: 'failingInstalls', label: 'Failing' },
              { key: 'inventoryWithoutCyberSmart', label: 'Inventory without matched install' },
            ]} />
          </section>
          <div className="notice"><AlertTriangle size={18} /><p>The 55-device difference between portal Secure and CSV-derived Pass needs portal-rule or export-timing reconciliation.</p></div>
        </section>
      </div>
    </>
  )
}

function SiteExplorer({
  core, installed, failing, remaining, servers,
}: {
  core: CoreData; installed: InstalledRow[] | null; failing: FailingRow[] | null; remaining: RemainingRow[] | null; servers: ServerData | null
}) {
  const [name, setName] = useState(core.sites[0]?.site ?? '')
  const site = core.sites.find((item) => item.site === name) ?? core.sites[0]
  if (!site) return null
  const fullInventory = core.clientCompute.bySite.find((item) => item.site === site.site)
  const siteInstalled = installed?.filter((row) => row.schoolSite === site.site) ?? []
  const siteFailing = failing?.filter((row) => row.schoolSite === site.site) ?? []
  const siteRemaining = remaining?.filter((row) => row.schoolSite === site.site) ?? []
  const siteServerInventory = servers?.inventory.filter((row) => row.school === site.site) ?? []
  const siteServerInstalls = servers?.matches.filter((row) => row.site === site.site) ?? []
  const siteServerSummary = core.serverSites.find((row) => row.site === site.site)
  const controls = Object.entries(site.failChecks).sort((a, b) => b[1] - a[1])
  return (
    <>
      <Header title="Complete site report" description="Every headline and evidence list for one site: total compute, matched and unmatched installations, compliance, remaining rollout, servers and failed controls." />
      <section className="panel site-picker">
        <label>Reporting site<select value={name} onChange={(event) => setName(event.target.value)}>{core.sites.map((item) => <option key={item.site}>{item.site}</option>)}</select></label>
        <Badge value={`${site.priority} priority`} />
      </section>
      <h2 className="section-title">Site inventory and installation headlines</h2>
      <div className="metrics">
        <Metric label="All client-compute items" value={fullInventory?.quantity ?? site.compatibleDevices} note={`${formatNumber.format(fullInventory?.records ?? 0)} source records`} tone="slate" />
        <Metric label="Compatible client devices" value={site.compatibleDevices} />
        <Metric label="CyberSmart devices" value={site.cyberSmartDevices} />
        <Metric label="Serial matched" value={site.serialMatched} tone="green" />
        <Metric label="Not serial matched" value={site.notSerialMatched} tone="amber" />
        <Metric label="Remaining to install" value={site.remainingToInstall} tone="red" />
        <Metric label="Server installs" value={site.serverInstalls} note="Separate from client compute" tone="purple" />
        <Metric label="Compatible servers & VMs" value={siteServerSummary?.compatibleServersAndVms ?? siteServerInventory.length} note={`${siteServerSummary?.physicalServers ?? 0} physical · ${siteServerSummary?.virtualMachines ?? 0} virtual`} tone="purple" />
      </div>
      <h2 className="section-title">Site compliance headlines</h2>
      <div className="metrics">
        <Metric label="Passing devices" value={site.passing} tone="green" />
        <Metric label="Failing devices" value={site.failing} tone="red" />
        <Metric label="Pass percentage" value={`${formatPercent.format(site.passPercent)}%`} />
        <Metric label="Possible unmanaged" value={site.possibleUnmanaged} tone="amber" />
        <Metric label="Passing server agents" value={siteServerSummary?.passingInstalls ?? 0} tone="green" />
        <Metric label="Failing server agents" value={siteServerSummary?.failingInstalls ?? 0} tone="red" />
        <Metric label="Server inventory without match" value={siteServerSummary?.inventoryWithoutCyberSmart ?? 0} tone="amber" />
      </div>
      <div className="two-column">
        <section className="panel">
          <h2>Full client-compute inventory by category</h2><p>All source inventory quantities for this site, including categories outside the compatible denominator.</p>
          <div className="breakdown">{Object.entries(fullInventory?.byCategory ?? {}).sort((a, b) => b[1] - a[1]).map(([label, value]) => <div key={label}><span>{label}</span><strong>{formatNumber.format(value)}</strong></div>)}</div>
        </section>
        <section className="panel">
          <h2>Remaining installations by device type</h2><p>Provisional allocation from the compatible client inventory.</p>
          <div className="breakdown">{Object.entries(site.remainingByType).map(([label, value]) => <div key={label}><span>{label.replace(/([a-z])([A-Z])/g, '$1 $2')}</span><strong>{formatNumber.format(value)}</strong></div>)}</div>
        </section>
      </div>
      <div className="two-column">
        <section className="panel">
          <h2>Failed controls</h2><p>Pass and fail totals for every control with evidence at this site.</p>
          <div className="breakdown">{controls.map(([label, fails]) => <div key={label}><span>{label.replace(/([a-z])([A-Z])/g, '$1 $2')}</span><strong>{formatNumber.format(site.passChecks[label] ?? 0)} pass / {formatNumber.format(fails)} fail</strong></div>)}</div>
        </section>
      </div>
      {!installed || !failing || !remaining || !servers ? <Loader /> : (
        <>
          <details open className="panel"><summary>All CyberSmart installations ({formatNumber.format(siteInstalled.length)})</summary>
            <DataTable rows={siteInstalled} columns={[
              { key: 'allocationCategory', label: 'Category' }, { key: 'host', label: 'Host' }, { key: 'serial', label: 'Serial' },
              { key: 'deviceType', label: 'Type' }, { key: 'overallStatus', label: 'Compliance', render: (row) => <Badge value={row.overallStatus} /> },
              { key: 'matchMethod', label: 'Inventory match' }, { key: 'failedChecks', label: 'Failed controls' },
            ]} />
          </details>
          <details className="panel"><summary>Failing devices ({formatNumber.format(siteFailing.length)})</summary>
            <DataTable rows={siteFailing} columns={[
              { key: 'host', label: 'Host' }, { key: 'serial', label: 'Serial' }, { key: 'deviceType', label: 'Type' },
              { key: 'os', label: 'Operating system' }, { key: 'failedChecks', label: 'Failed controls' },
              { key: 'failureReasons', label: 'Reason' }, { key: 'lastCheckIn', label: 'Last check-in' },
            ]} />
          </details>
          <details className="panel"><summary>Remaining inventory allocation ({formatNumber.format(siteRemaining.reduce((total, row) => total + Number(row.allocateQuantity || 0), 0))} devices)</summary>
            <DataTable rows={siteRemaining} columns={[
              { key: 'deviceCategory', label: 'Inventory category' }, { key: 'deviceType', label: 'Type' }, { key: 'allocateQuantity', label: 'Quantity' },
              { key: 'serial', label: 'Serial' }, { key: 'assetTag', label: 'Asset tag' }, { key: 'manufacturer', label: 'Manufacturer' },
              { key: 'model', label: 'Model' }, { key: 'location', label: 'Location' },
            ]} />
          </details>
          <details className="panel"><summary>Compatible server and VM inventory ({formatNumber.format(siteServerInventory.length)})</summary>
            <DataTable rows={siteServerInventory} columns={[
              { key: 'serverName', label: 'Server / VM' }, { key: 'serialNumber', label: 'Serial' }, { key: 'physicalVirtual', label: 'Physical / virtual' },
              { key: 'roleWorkload', label: 'Role / workload' }, { key: 'osHypervisor', label: 'Operating system / hypervisor' },
              { key: 'location', label: 'Location' }, { key: 'coverageStatus', label: 'CyberSmart coverage', render: (row) => <Badge value={row.coverageStatus} /> },
              { key: 'cyberSmartHosts', label: 'Matched CyberSmart host' },
            ]} />
          </details>
          <details className="panel"><summary>CyberSmart server installations ({formatNumber.format(siteServerInstalls.length)})</summary>
            <DataTable rows={siteServerInstalls} columns={[
              { key: 'host', label: 'CyberSmart host' }, { key: 'serial', label: 'Serial' }, { key: 'status', label: 'Compliance', render: (row) => <Badge value={row.status} /> },
              { key: 'matchMethod', label: 'Infrastructure match' }, { key: 'inventoryServerName', label: 'Inventory server / VM' },
              { key: 'physicalVirtual', label: 'Physical / virtual' }, { key: 'failedChecks', label: 'Failed controls' },
            ]} />
          </details>
        </>
      )}
    </>
  )
}

function BrowserPage<T extends Record<string, unknown>>({
  title, description, rows, columns, siteKey, typeKey, filename, metrics,
}: {
  title: string; description: string; rows: T[]; columns: Array<Column<T>>;
  siteKey: keyof T & string; typeKey: keyof T & string; filename: string; metrics?: ReactNode
}) {
  return <><Header title={title} description={description} />{metrics}<FilteredTable rows={rows} columns={columns} siteKey={siteKey} typeKey={typeKey} filename={filename} /></>
}

type AssignedFailure = FailingRow & {
  owner: string; priority: string; remediationStatus: string; dueDate: string; notes: string
}

function Failures({ rows }: { rows: FailingRow[] }) {
  const storageKey = 'cybersmart-remediation-v1'
  const [remediation, setRemediation] = useState<Record<string, Remediation>>(() => {
    try { return JSON.parse(localStorage.getItem(storageKey) ?? '{}') as Record<string, Remediation> } catch { return {} }
  })
  const update = (id: string, key: keyof Remediation, value: string) => {
    const current = remediation[id] ?? { owner: '', priority: 'Medium', status: 'Not started', dueDate: '', notes: '' }
    const next = { ...remediation, [id]: { ...current, [key]: value } }
    setRemediation(next)
    localStorage.setItem(storageKey, JSON.stringify(next))
  }
  const assigned: AssignedFailure[] = rows.map((row) => {
    const item = remediation[row.cyberSmartDeviceId]
    return { ...row, owner: item?.owner ?? '', priority: item?.priority ?? 'Medium', remediationStatus: item?.status ?? 'Not started', dueDate: item?.dueDate ?? '', notes: item?.notes ?? '' }
  })
  return (
    <BrowserPage
      title="Failures & remediation allocation"
      description="Locate the exact failing device, review its controls and assign remediation locally in this browser without modifying source evidence."
      rows={assigned} siteKey="schoolSite" typeKey="deviceType" filename="cybersmart-failures-and-remediation.csv"
      metrics={<div className="metrics compact">
        <Metric label="Failing school devices" value={rows.length} tone="red" />
        <Metric label="Assigned to an owner" value={Object.values(remediation).filter((item) => item.owner).length} />
        <Metric label="Marked complete" value={Object.values(remediation).filter((item) => item.status === 'Complete').length} tone="green" />
      </div>}
      columns={[
        { key: 'schoolSite', label: 'Site' }, { key: 'host', label: 'Host' }, { key: 'serial', label: 'Serial' }, { key: 'deviceType', label: 'Type' },
        { key: 'failedChecks', label: 'Failed controls' },
        { key: 'owner', label: 'Owner', render: (row) => <input className="cell-input" value={row.owner} onChange={(event) => update(row.cyberSmartDeviceId, 'owner', event.target.value)} placeholder="Assign owner" /> },
        { key: 'priority', label: 'Priority', render: (row) => <select className="cell-select" value={row.priority} onChange={(event) => update(row.cyberSmartDeviceId, 'priority', event.target.value)}><option>Critical</option><option>High</option><option>Medium</option><option>Low</option></select> },
        { key: 'remediationStatus', label: 'Remediation', render: (row) => <select className="cell-select" value={row.remediationStatus} onChange={(event) => update(row.cyberSmartDeviceId, 'status', event.target.value)}><option>Not started</option><option>Assigned</option><option>In progress</option><option>Blocked</option><option>Complete</option></select> },
        { key: 'dueDate', label: 'Due', render: (row) => <input className="cell-input" type="date" value={row.dueDate} onChange={(event) => update(row.cyberSmartDeviceId, 'dueDate', event.target.value)} /> },
        { key: 'notes', label: 'Notes', render: (row) => <input className="cell-input notes" value={row.notes} onChange={(event) => update(row.cyberSmartDeviceId, 'notes', event.target.value)} placeholder="Local notes" /> },
      ]}
    />
  )
}

function Checks({ rows }: { rows: CheckRow[] }) {
  const [site, setSite] = useState('')
  const sites = [...new Set(rows.map((row) => row.schoolSite))].sort()
  const filtered = rows.filter((row) => !site || row.schoolSite === site)
  const aggregate = Object.values(filtered.reduce<Record<string, CheckRow>>((result, row) => {
    const current = result[row.cyberSmartCheck] ?? { ...row, schoolSite: 'Selected sites', devicesAssessed: 0, passingDevices: 0, failingDevices: 0, unknownDevices: 0, checkPassPercent: 0 }
    current.devicesAssessed += row.devicesAssessed
    current.passingDevices += row.passingDevices
    current.failingDevices += row.failingDevices
    current.unknownDevices += row.unknownDevices
    current.checkPassPercent = current.devicesAssessed ? current.passingDevices / current.devicesAssessed * 100 : 0
    result[row.cyberSmartCheck] = current
    return result
  }, {})).sort((a, b) => b.failingDevices - a.failingDevices)
  return (
    <>
      <Header title="Control performance" description="Trust-wide or site-level Pass, Fail and Unknown totals for every CyberSmart control." />
      <section className="panel">
        <div className="filters"><select value={site} onChange={(event) => setSite(event.target.value)}><option value="">All sites</option>{sites.map((value) => <option key={value}>{value}</option>)}</select></div>
        <DataTable rows={aggregate} columns={[
          { key: 'cyberSmartCheck', label: 'CyberSmart control' }, { key: 'devicesAssessed', label: 'Assessed' },
          { key: 'passingDevices', label: 'Pass' }, { key: 'failingDevices', label: 'Fail' }, { key: 'unknownDevices', label: 'Unknown' },
          { key: 'checkPassPercent', label: 'Pass rate', render: (row) => `${formatPercent.format(row.checkPassPercent)}%` },
        ]} />
      </section>
    </>
  )
}

function ServerPage({ data }: { data: ServerData }) {
  return (
    <>
      <Header title="Server CyberSmart report" description="A standalone report for the complete compatible server and VM estate, CyberSmart server installations, infrastructure matching, compliance and remaining coverage." />
      <div className="metrics">
        <Metric label="Infrastructure records" value={data.summary.inventoryRecords} note={`${data.summary.physicalServers} physical · ${data.summary.virtualMachines} virtual`} />
        <Metric label="CyberSmart server installs" value={data.summary.cyberSmartServerInstalls} tone="purple" />
        <Metric label="Matched to infrastructure" value={data.summary.matchedServers} tone="green" />
        <Metric label="Unmatched server agents" value={data.summary.unmatchedServers} tone="amber" />
        <Metric label="Inventory with CyberSmart match" value={data.summary.inventoryWithCyberSmart} tone="green" />
        <Metric label="Inventory without CyberSmart match" value={data.summary.inventoryWithoutCyberSmart} tone="red" />
        <Metric label="Passing server agents" value={data.summary.passingServers} tone="green" />
        <Metric label="Failing server agents" value={data.summary.failingServers} tone="red" />
      </div>
      <h2 className="section-title">Complete compatible infrastructure inventory</h2>
      <FilteredTable<ServerInventory>
        rows={data.inventory} siteKey="school" typeKey="physicalVirtual" filename="compatible-server-vm-inventory.csv"
        columns={[
          { key: 'school', label: 'Site' }, { key: 'serverName', label: 'Server / VM' }, { key: 'serialNumber', label: 'Serial' },
          { key: 'physicalVirtual', label: 'Physical / virtual' }, { key: 'roleWorkload', label: 'Role / workload' },
          { key: 'manufacturer', label: 'Manufacturer' }, { key: 'model', label: 'Model' }, { key: 'osHypervisor', label: 'Operating system / hypervisor' },
          { key: 'location', label: 'Location' }, { key: 'coverageStatus', label: 'CyberSmart coverage', render: (row) => <Badge value={row.coverageStatus} /> },
          { key: 'cyberSmartHosts', label: 'Matched CyberSmart hosts' },
        ]}
      />
      <h2 className="section-title">CyberSmart server installation evidence</h2>
      <FilteredTable<ServerMatch>
      rows={data.matches} siteKey="site" typeKey="physicalVirtual" filename="cybersmart-server-reconciliation.csv"
      columns={[
        { key: 'site', label: 'Site' }, { key: 'host', label: 'CyberSmart host' }, { key: 'serial', label: 'CyberSmart serial' },
        { key: 'status', label: 'Compliance', render: (row) => <Badge value={row.status} /> }, { key: 'matchMethod', label: 'Infrastructure match' },
        { key: 'inventoryServerName', label: 'Inventory server' }, { key: 'physicalVirtual', label: 'Physical / virtual' },
        { key: 'roleWorkload', label: 'Role / workload' }, { key: 'location', label: 'Location' }, { key: 'failedChecks', label: 'Failed controls' },
      ]}
      />
    </>
  )
}

function Guide({ core }: { core: CoreData }) {
  const items = [
    ['Start with total compute', 'Compatible client compute is the denominator for each site. Screens and servers are excluded; possible unmanaged devices remain visible.'],
    ['Match CyberSmart', 'Category 1 is installed with an exact same-site serial match. Category 2 is installed without an exact serial match, regardless of compliance status.'],
    ['Allocate remaining devices', 'Category 3 prioritises unmatched serialised inventory, then quantity-only records. Confirm local identity before deployment.'],
    ['Remediate failures', 'Use hostname, serial, failed controls and check-in evidence. Assignment fields are browser-local and available in CSV export.'],
    ['Keep servers separate', 'Server and VM agents are reconciled to infrastructure where possible but never reduce the client installation gap.'],
    ['Read portal totals carefully', `${formatNumber.format(core.summary.pending)} portal devices have no detailed export rows. Portal Secure and all-exported-checks-pass are different measures.`],
  ]
  return (
    <>
      <Header title="Guide & methodology" description="How to interpret trust-wide and site reporting, matching, gaps, controls and remediation evidence." />
      <div className="guide-grid">{items.map(([title, body], index) => <section className="panel" key={title}><span className="step">{index + 1}</span><h2>{title}</h2><p>{body}</p></section>)}</div>
      <section className="panel public-warning"><AlertTriangle size={22} /><div><h2>Public reporting notice</h2><p>This static report contains hostnames, serial numbers and device-level security findings. Treat the shared URL as sensitive operational information.</p></div></section>
    </>
  )
}

function App() {
  const [page, setPage] = useState<Page>('overview')
  const [menu, setMenu] = useState(false)
  const core = useDataset<CoreData>('core.json', true)
  const inventory = useDataset<InventoryRow[]>('inventory.json', page === 'inventory')
  const installed = useDataset<InstalledRow[]>('installed-devices.json', page === 'installed' || page === 'sites')
  const failing = useDataset<FailingRow[]>('failing-devices.json', page === 'failures' || page === 'sites')
  const remaining = useDataset<RemainingRow[]>('remaining-devices.json', page === 'remaining' || page === 'sites')
  const servers = useDataset<ServerData>('servers.json', page === 'servers' || page === 'sites')
  const navigate = (next: Page) => { setPage(next); setMenu(false); window.scrollTo({ top: 0, behavior: 'smooth' }) }

  if (!core.data) return <main className="startup"><ShieldCheck size={42} /><h1>Loading CyberSmart reporting</h1><Loader error={core.error} /></main>

  let content: ReactNode
  if (page === 'overview') content = <Overview core={core.data} navigate={navigate} />
  else if (page === 'sites') content = <SiteExplorer core={core.data} installed={installed.data} failing={failing.data} remaining={remaining.data} servers={servers.data} />
  else if (page === 'inventory') content = inventory.data ? <BrowserPage title="Authoritative compute inventory" description="Every client-compute record that forms the site denominator, including serialised and quantity-only evidence." rows={inventory.data} siteKey="school" typeKey="compatibleType" filename="compatible-compute-inventory.csv" columns={[
    { key: 'school', label: 'Site' }, { key: 'deviceCategory', label: 'Inventory category' }, { key: 'compatibleType', label: 'Compatible type' },
    { key: 'quantity', label: 'Quantity' }, { key: 'serialNumber', label: 'Serial' }, { key: 'assetTag', label: 'Asset tag' },
    { key: 'manufacturer', label: 'Manufacturer' }, { key: 'model', label: 'Model' }, { key: 'operatingSystem', label: 'Operating system' },
    { key: 'exactCyberSmartMatch', label: 'CyberSmart match', render: (row) => <Badge value={row.exactCyberSmartMatch} /> }, { key: 'location', label: 'Location' },
  ]} /> : <Loader error={inventory.error} />
  else if (page === 'installed') content = installed.data ? <BrowserPage title="CyberSmart installations" description="All school installations split between exact serial matches and devices not matched by serial. Pass or fail does not alter match status." rows={installed.data} siteKey="schoolSite" typeKey="deviceType" filename="cybersmart-installations.csv" columns={[
    { key: 'schoolSite', label: 'Site' }, { key: 'allocationCategory', label: 'Category' }, { key: 'host', label: 'Host' },
    { key: 'serial', label: 'Serial' }, { key: 'deviceType', label: 'Type' }, { key: 'overallStatus', label: 'Compliance', render: (row) => <Badge value={row.overallStatus} /> },
    { key: 'os', label: 'Operating system' }, { key: 'matchMethod', label: 'Inventory match' }, { key: 'failedChecks', label: 'Failed controls' },
  ]} /> : <Loader error={installed.error} />
  else if (page === 'failures') content = failing.data ? <Failures rows={failing.data} /> : <Loader error={failing.error} />
  else if (page === 'remaining') content = remaining.data ? <BrowserPage title="Remaining installation allocation" description="Provisional inventory-first list of client devices still requiring CyberSmart. Quantity-only records must be identified locally." rows={remaining.data} siteKey="schoolSite" typeKey="deviceType" filename="remaining-cybersmart-installations.csv" columns={[
    { key: 'schoolSite', label: 'Site' }, { key: 'deviceCategory', label: 'Inventory category' }, { key: 'deviceType', label: 'Type' },
    { key: 'allocateQuantity', label: 'Allocate' }, { key: 'serial', label: 'Serial' }, { key: 'assetTag', label: 'Asset tag' },
    { key: 'operatingSystem', label: 'Operating system' }, { key: 'manufacturer', label: 'Manufacturer' }, { key: 'model', label: 'Model' },
    { key: 'location', label: 'Location' }, { key: 'evidenceConfidence', label: 'Evidence confidence' },
  ]} /> : <Loader error={remaining.error} />
  else if (page === 'servers') content = servers.data ? <ServerPage data={servers.data} /> : <Loader error={servers.error} />
  else if (page === 'checks') content = <Checks rows={core.data.checks} />
  else content = <Guide core={core.data} />

  return (
    <div className="app">
      <aside className={`sidebar ${menu ? 'open' : ''}`}>
        <div className="brand"><span className="brand-icon"><ShieldCheck size={24} /></span><div><strong>WeST CyberSmart</strong><small>Coverage & compliance</small></div></div>
        <button className="close-menu" type="button" onClick={() => setMenu(false)}><X /></button>
        <nav>{nav.map((item) => { const Icon = item.icon; return <button key={item.id} className={page === item.id ? 'active' : ''} onClick={() => navigate(item.id)}><Icon size={18} />{item.label}</button> })}</nav>
        <div className="sidebar-footer"><span><CheckCircle2 size={12} /> Validated evidence snapshot</span><small>Generated {new Date(core.data.generatedAt).toLocaleDateString('en-GB')}</small></div>
      </aside>
      {menu && <button className="scrim" onClick={() => setMenu(false)} />}
      <div className="content-shell">
        <header className="mobile-header"><button type="button" onClick={() => setMenu(true)}><Menu /></button><strong>WeST CyberSmart</strong></header>
        <main className="content">{content}</main>
      </div>
    </div>
  )
}

export default App
