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
  CheckRow, CoreData, CyberDeviceRow, FailingRow, InstalledRow, InventoryRow, Page, Remediation,
  ProgrammeTask, RemainingRow, ServerData, ServerInventory, ServerMatch,
} from './types'

const nav: Array<{ id: Page; label: string; icon: typeof CircleGauge }> = [
  { id: 'overview', label: 'Executive overview', icon: CircleGauge },
  { id: 'programme', label: 'Remediation programme', icon: CheckCircle2 },
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
      <div className="download-actions">
        <a className="primary" href={`${import.meta.env.BASE_URL}downloads/WeST-CyberSmart-Project-Delivery-Plan.pptx`}>
          <Download size={17} /> Project PowerPoint
        </a>
        <a className="secondary" href={`${import.meta.env.BASE_URL}downloads/WeST-CyberSmart-Assessment-and-Programme-Tracker.xlsx`}>
          <Download size={17} /> Excel assessment & tracker
        </a>
        <a className="secondary" href={`${import.meta.env.BASE_URL}downloads/WeST-CyberSmart-Project-Delivery-Plan.xlsx`}>
          <Download size={17} /> Project delivery plan
        </a>
      </div>
    </div>
  )
}

function ReadingGuide({ items, title = 'How to read this page' }: { items: string[]; title?: string }) {
  return (
    <aside className="reading-guide">
      <div><BookOpen size={18} /><strong>{title}</strong></div>
      <ul>{items.map((item) => <li key={item}>{item}</li>)}</ul>
    </aside>
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

function HighlightStat({
  icon, label, value, tone = '',
}: {
  icon: ReactNode; label: string; value: number; tone?: string
}) {
  return (
    <article className={`highlight-stat ${tone}`}>
      <div className="highlight-icon">{icon}</div>
      <div><span>{label}</span><strong>{formatNumber.format(value)}</strong></div>
    </article>
  )
}

function readableLabel(value: string) {
  const known: Record<string, string> = {
    AppleMac: 'Apple / Mac',
    IPadIOs: 'iPad / iOS',
    AndroidTablet: 'Android / tablet',
    PossibleTypeUnknown: 'Possible / unknown',
  }
  return known[value] ?? value.replace(/([a-z0-9])([A-Z])/g, '$1 $2')
}

interface Column<T> { key: keyof T & string; label: string; description?: string; render?: (row: T) => ReactNode }

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
          <thead><tr>{columns.map((column) => <th key={column.key} title={column.description}>{column.label}</th>)}</tr></thead>
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
  rows, columns, siteKey, typeKey, filename, ownerKey, failureKey, typeLabel = 'device types',
}: {
  rows: T[]; columns: Array<Column<T>>; siteKey: keyof T & string; typeKey: keyof T & string; filename: string
  ownerKey?: keyof T & string; failureKey?: keyof T & string; typeLabel?: string
}) {
  const [search, setSearch] = useState('')
  const [site, setSite] = useState('')
  const [type, setType] = useState('')
  const [owner, setOwner] = useState('')
  const [failure, setFailure] = useState('')
  const sites = useMemo(() => [...new Set(rows.map((row) => display(row[siteKey])).filter((value) => value !== '—'))].sort(), [rows, siteKey])
  const types = useMemo(() => [...new Set(rows.map((row) => display(row[typeKey])).filter((value) => value !== '—'))].sort(), [rows, typeKey])
  const owners = useMemo(() => ownerKey ? [...new Set(rows.map((row) => display(row[ownerKey])).filter((value) => value !== '—'))].sort() : [], [ownerKey, rows])
  const failures = useMemo(() => failureKey ? [...new Set(rows.flatMap((row) => String(row[failureKey] ?? '').split(';').map((value) => value.trim()).filter(Boolean)))].sort() : [], [failureKey, rows])
  const filtered = useMemo(() => rows.filter((row) =>
    (!site || display(row[siteKey]) === site) &&
    (!type || display(row[typeKey]) === type) &&
    (!owner || !ownerKey || display(row[ownerKey]) === owner) &&
    (!failure || !failureKey || String(row[failureKey] ?? '').split(';').map((value) => value.trim()).includes(failure)) &&
    matchesSearch(row, search)), [failure, failureKey, owner, ownerKey, rows, search, site, siteKey, type, typeKey])
  return (
    <section className="panel">
      <div className="filters">
        <label className="search"><Search size={17} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search hostname, serial, model, asset tag…" /></label>
        <select value={site} onChange={(event) => setSite(event.target.value)}><option value="">All sites</option>{sites.map((value) => <option key={value}>{value}</option>)}</select>
        {ownerKey && <select value={owner} onChange={(event) => setOwner(event.target.value)}><option value="">All task owners</option>{owners.map((value) => <option key={value}>{value}</option>)}</select>}
        <select value={type} onChange={(event) => setType(event.target.value)}><option value="">All {typeLabel}</option>{types.map((value) => <option key={value}>{value}</option>)}</select>
        {failureKey && <select value={failure} onChange={(event) => setFailure(event.target.value)}><option value="">All failed controls</option>{failures.map((value) => <option key={value}>{value}</option>)}</select>}
        <button className="secondary" type="button" onClick={() => exportCsv(filename, filtered)}><Download size={16} /> Export filtered CSV</button>
      </div>
      <DataTable rows={filtered} columns={columns} />
    </section>
  )
}

function Overview({ core, navigate }: { core: CoreData; navigate: (page: Page) => void }) {
  const [selectedSiteName, setSelectedSiteName] = useState(core.sites[0]?.site ?? '')
  const [selectedServerSiteName, setSelectedServerSiteName] = useState(core.serverSites[0]?.site ?? '')
  const s = core.summary
  const selectedSite = core.sites.find((site) => site.site === selectedSiteName) ?? core.sites[0]
  const selectedServerSite = core.serverSites.find((site) => site.site === selectedServerSiteName) ?? core.serverSites[0]
  const clientInstalls = selectedSite ? Math.max(0, selectedSite.cyberSmartDevices - selectedSite.serverInstalls) : 0
  const clientUnmatched = selectedSite ? Math.max(0, selectedSite.notSerialMatched - selectedSite.serverInstalls) : 0
  const remainingByType = selectedSite
    ? Object.entries(selectedSite.remainingByType)
      .filter(([, value]) => value > 0)
      .map(([type, value]) => ({ name: readableLabel(type), Remaining: value }))
      .sort((a, b) => b.Remaining - a.Remaining)
    : []
  const controlPerformance = selectedSite
    ? [...new Set([...Object.keys(selectedSite.passChecks), ...Object.keys(selectedSite.failChecks)])]
      .map((control) => ({
        name: readableLabel(control),
        Passing: selectedSite.passChecks[control] ?? 0,
        Failing: selectedSite.failChecks[control] ?? 0,
      }))
      .filter((control) => control.Passing + control.Failing > 0)
      .sort((a, b) => b.Failing - a.Failing || b.Passing - a.Passing)
      .slice(0, 8)
    : []
  const serverCoveragePercent = selectedServerSite?.compatibleServersAndVms
    ? Math.round((selectedServerSite.inventoryWithCyberSmart / selectedServerSite.compatibleServersAndVms) * 100)
    : 0
  const chart = [...core.sites].sort((a, b) => b.compatibleDevices - a.compatibleDevices).slice(0, 10).map((site) => ({
    name: site.site.replace(/ School| Academy| College/g, ''),
    'Installed clients': Math.max(0, site.cyberSmartDevices - site.serverInstalls),
    'Remaining clients': site.remainingToInstall,
  }))
  const taskBreakdown = core.sites.map((site) => {
    const tasks = core.taskProgramme.tasks.filter((task) => task.site === site.site)
    return {
      site: site.site,
      installOwner: core.taskProgramme.rolloutOwners[site.site] ?? core.taskProgramme.owner,
      dueDate: core.taskProgramme.schoolSchedule.find((item) => item.site === site.site)?.dueDate ?? '',
      remainingDevices: site.remainingToInstall,
      knownRemaining: site.knownRemaining,
      explicitUnmanagedRemaining: site.explicitUnmanagedRemaining,
      unallocatedInventoryGap: site.unallocatedInventoryGap,
      installationTasks: tasks.filter((task) => task.category.startsWith('CyberSmart rollout')).length,
      complianceTasks: tasks.filter((task) => task.category === 'Compliance remediation').length,
      inventoryTasks: tasks.filter((task) => task.category === 'Inventory reconciliation').length,
      serverTasks: tasks.filter((task) => task.category === 'Server CyberSmart').length,
      mfaTasks: tasks.filter((task) => task.category === 'MFA').length,
      totalTasks: tasks.length,
    }
  })
  const taskColumnGuide = [
    { label: 'School', description: 'The school or delivery site covered by that row.' },
    { label: 'Remaining-install owner', description: 'The person accountable for deploying CyberSmart to the remaining compatible client devices. Other task types remain assigned to their named task owner.' },
    { label: 'Deadline', description: 'The target date by which the school’s full CyberSmart work package should be completed.' },
    { label: 'Remaining devices', description: 'Compatible client devices that do not yet have a confirmed CyberSmart installation allocation.' },
    { label: 'Known remaining', description: 'Remaining devices tied to an identifiable managed inventory category or device row.' },
    { label: 'Explicit unmanaged', description: 'Remaining inventory explicitly described as unmanaged or unknown in the source evidence.' },
    { label: 'Unallocated / uncertain gap', description: 'A calculated quantity where the coverage gap is larger than the identifiable remaining inventory rows. These are not blank rows and must be reconciled locally.' },
    { label: 'Device-type install tasks', description: 'Separate rollout tasks for device categories that still need installations, such as Windows, Chromebook, Apple or iPad.' },
    { label: 'Compliance tasks', description: 'Tasks to correct CyberSmart control failures found on installed devices.' },
    { label: 'Inventory tasks', description: 'Tasks to reconcile unmatched agents, serial numbers and inventory evidence.' },
    { label: 'Server tasks', description: 'Tasks to confirm and complete CyberSmart coverage across the school’s servers and virtual machines.' },
    { label: 'MFA tasks', description: 'The site task to implement or evidence multi-factor authentication.' },
    { label: 'Total tasks', description: 'The total number of programme tasks assigned to the school across all task types.' },
  ]
  return (
    <>
      <Header title="Trust-wide inventory and coverage" description="All reporting starts with the authoritative compatible client-compute estate, then reconciles CyberSmart installation, serial matching, compliance and remaining rollout." />
      <ReadingGuide items={[
        'Compatible client devices are the main coverage denominator; servers and VMs are reported separately.',
        'Serial matched is the strongest installation evidence. Without serial match means an agent exists but still needs inventory reconciliation.',
        'Remaining installs are compatible devices without a confirmed CyberSmart allocation, split by device type for delivery.',
        'Passing and failing figures come from exported CyberSmart controls; a device fails overall if one or more assessed controls fail.',
      ]} />
      <h2 className="section-title">Inventory and installation headlines</h2>
      <div className="metrics">
        <Metric label="All client-compute items" value={core.clientCompute.quantity} note={`${formatNumber.format(core.clientCompute.records)} source records, including non-compatible categories`} tone="slate" />
        <Metric label="Compatible client devices" value={s.compatibleDevices} note="Authoritative client denominator" />
        <Metric label="School CyberSmart devices" value={s.schoolCyberSmartDevices} note="Includes separately reported servers" />
        <Metric label="Serial-matched installs" value={s.serialMatched} note="Passing and failing devices" tone="green" />
        <Metric label="Not serial matched" value={s.notSerialMatched} note="Includes server installs" tone="amber" />
        <Metric label="Remaining client installs" value={s.remainingToInstall} note="Provisional inventory allocation" tone="red" />
        <Metric label="Total compatible servers & VMs" value={core.serverSummary.inventoryRecords} note={`${formatNumber.format(core.serverSummary.physicalServers)} physical · ${formatNumber.format(core.serverSummary.virtualMachines)} virtual`} tone="purple" />
        <Metric label="Confirmed CyberSmart server installs" value={core.serverSummary.inventoryWithCyberSmart} note="Unique infrastructure records with a confirmed agent match" tone="green" />
        <Metric label="Remaining server CyberSmart installs" value={core.serverSummary.inventoryWithoutCyberSmart} note="Compatible infrastructure without a confirmed match" tone="red" />
        <Metric label="Server-labelled agents (raw)" value={s.serverInstalls} note="Reconciliation evidence only — not installed-server coverage" tone="amber" />
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
      {selectedSite && (
        <section className="panel coverage-highlight">
          <div className="highlight-heading">
            <div>
              <h2><Laptop size={20} /> CyberSmart client coverage by school</h2>
              <p>Select a school for a wide dashboard view of installation coverage, remaining device types and control performance.</p>
            </div>
            <label className="dashboard-select">
              School
              <select value={selectedSite.site} onChange={(event) => setSelectedSiteName(event.target.value)}>
                {core.sites.map((site) => <option key={site.site}>{site.site}</option>)}
              </select>
            </label>
          </div>
          <div className="highlight-grid">
            <HighlightStat icon={<Laptop size={20} />} label="Compatible devices" value={selectedSite.compatibleDevices} />
            <HighlightStat icon={<MonitorCheck size={20} />} label="CyberSmart installs" value={clientInstalls} tone="blue" />
            <HighlightStat icon={<CheckCircle2 size={20} />} label="Serial matched" value={selectedSite.serialMatched} tone="green" />
            <HighlightStat icon={<Search size={20} />} label="Without serial match" value={clientUnmatched} tone="amber" />
            <HighlightStat icon={<AlertTriangle size={20} />} label="Remaining installs" value={selectedSite.remainingToInstall} tone="red" />
          </div>
          <div className="coverage-summary">
            <div className="coverage-status">
              <span><ShieldCheck size={17} /> Overall passing <strong>{formatNumber.format(selectedSite.passing)}</strong></span>
              <span><ShieldAlert size={17} /> Overall failing <strong>{formatNumber.format(selectedSite.failing)}</strong></span>
              <span>Pass rate <strong>{formatPercent.format(selectedSite.passPercent)}%</strong></span>
            </div>
            <button className="secondary" type="button" onClick={() => navigate('sites')}>Open complete site report <ChevronRight size={16} /></button>
          </div>
          <div className="highlight-charts">
            <div className="chart-card">
              <h3>Remaining installs by device type</h3>
              <p>Compatible devices still requiring CyberSmart, grouped for delivery planning.</p>
              {remainingByType.length ? (
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart data={remainingByType} layout="vertical" margin={{ left: 18, right: 24, top: 16 }}>
                    <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                    <XAxis type="number" />
                    <YAxis dataKey="name" type="category" width={120} tick={{ fontSize: 11 }} />
                    <Tooltip /><Bar dataKey="Remaining" fill="#f3a83b" radius={[0, 5, 5, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              ) : <div className="empty compact-empty">No remaining device allocation for this school.</div>}
            </div>
            <div className="chart-card">
              <h3>Highest-volume control results</h3>
              <p>Passing and failing device counts for the most material assessed controls.</p>
              {controlPerformance.length ? (
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart data={controlPerformance} layout="vertical" margin={{ left: 18, right: 24, top: 16 }}>
                    <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                    <XAxis type="number" />
                    <YAxis dataKey="name" type="category" width={170} tick={{ fontSize: 10 }} />
                    <Tooltip /><Legend />
                    <Bar dataKey="Passing" stackId="controls" fill="#16a474" />
                    <Bar dataKey="Failing" stackId="controls" fill="#dd4b57" radius={[0, 5, 5, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              ) : <div className="empty compact-empty">No assessed control results for this school.</div>}
            </div>
          </div>
        </section>
      )}
      <section className="panel">
        <h2>Remediation programme</h2>
        <p>{core.taskProgramme.owner} is the programme owner. Named local owners are assigned to remaining compatible-device installations at designated secondary schools. WeST Central is due 1 December 2026; school waves run through 31 July 2027.</p>
        <div className="flow">
          <button onClick={() => navigate('programme')}><strong>{formatNumber.format(core.taskProgramme.tasks.length)}</strong><span>Total programme tasks</span></button><ChevronRight />
          <button onClick={() => navigate('programme')}><strong>{formatNumber.format(core.taskProgramme.tasks.filter((task) => task.site === 'WeST Central').length)}</strong><span>Central and trust-wide tasks</span></button><ChevronRight />
          <button onClick={() => navigate('programme')}><strong>{formatNumber.format(core.sites.length)}</strong><span>School completion plans</span></button><ChevronRight />
          <button onClick={() => navigate('programme')}><strong>31 Jul 2027</strong><span>Final completion deadline</span></button>
        </div>
      </section>
      <section className="panel">
        <h2>School task types, owners and deadlines</h2>
        <p>Dashboard breakdown of the delivery workload for every school. Installation tasks are split by remaining device type.</p>
        <ReadingGuide title="How to read this table" items={[
          'Each row is one school. Remaining devices is a device quantity; the task columns are counts of delivery tasks.',
          'Known remaining is identifiable inventory; Explicit unmanaged is labelled that way in source data; Unallocated / uncertain gap is a calculated shortfall requiring reconciliation.',
          'Remaining-install owner is accountable only for device installation tasks. Other work stays with the owner named in the full programme.',
          'Device-type install tasks are separated by categories such as Windows, Chromebook, Apple or iPad.',
          'Deadline is the target for the school’s complete work package; Total tasks adds all installation, compliance, inventory, server and MFA tasks.',
        ]} />
        <div className="filters">
          <button className="secondary" type="button" onClick={() => exportCsv('school-task-type-owner-breakdown.csv', taskBreakdown)}>
            <Download size={16} /> Export school task matrix
          </button>
        </div>
        <DataTable rows={taskBreakdown} columns={[
          { key: 'site', label: 'School', description: taskColumnGuide[0].description },
          { key: 'installOwner', label: 'Remaining-install owner', description: taskColumnGuide[1].description },
          { key: 'dueDate', label: 'Deadline', description: taskColumnGuide[2].description },
          { key: 'remainingDevices', label: 'Remaining devices', description: taskColumnGuide[3].description },
          { key: 'knownRemaining', label: 'Known remaining', description: taskColumnGuide[4].description },
          { key: 'explicitUnmanagedRemaining', label: 'Explicit unmanaged', description: taskColumnGuide[5].description },
          { key: 'unallocatedInventoryGap', label: 'Unallocated / uncertain gap', description: taskColumnGuide[6].description },
          { key: 'installationTasks', label: 'Device-type install tasks', description: taskColumnGuide[7].description },
          { key: 'complianceTasks', label: 'Compliance tasks', description: taskColumnGuide[8].description },
          { key: 'inventoryTasks', label: 'Inventory tasks', description: taskColumnGuide[9].description },
          { key: 'serverTasks', label: 'Server tasks', description: taskColumnGuide[10].description },
          { key: 'mfaTasks', label: 'MFA tasks', description: taskColumnGuide[11].description },
          { key: 'totalTasks', label: 'Total tasks', description: taskColumnGuide[12].description },
        ]} />
      </section>
      <div className="two-column">
        <section className="panel">
          <h2>Largest compatible client estates</h2><p>Installed client agents and remaining client installation allocation. Server evidence is reported separately.</p>
          <ResponsiveContainer width="100%" height={420}>
            <BarChart data={chart} layout="vertical" margin={{ left: 18, right: 12, top: 20 }}>
              <CartesianGrid strokeDasharray="3 3" horizontal={false} /><XAxis type="number" />
              <YAxis dataKey="name" type="category" width={150} tick={{ fontSize: 11 }} />
              <Tooltip /><Legend /><Bar dataKey="Installed clients" stackId="a" fill="#2878ff" />
              <Bar dataKey="Remaining clients" stackId="a" fill="#f3a83b" />
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
          <div className="notice"><AlertTriangle size={18} /><p>The 55-device difference between portal Secure and CSV-derived Pass needs portal-rule or export-timing reconciliation.</p></div>
        </section>
      </div>
      {selectedServerSite && (
        <section className="panel coverage-highlight server-highlight">
          <div className="highlight-heading">
            <div>
              <h2><Server size={20} /> Server CyberSmart coverage by site</h2>
              <p>A full-width infrastructure view. Confirmed installs are unique authoritative server or VM records with a matched CyberSmart agent.</p>
            </div>
            <label className="dashboard-select">
              Site
              <select value={selectedServerSite.site} onChange={(event) => setSelectedServerSiteName(event.target.value)}>
                {core.serverSites.map((site) => <option key={site.site}>{site.site}</option>)}
              </select>
            </label>
          </div>
          <div className="highlight-grid four">
            <HighlightStat icon={<Server size={20} />} label="Total servers & VMs" value={selectedServerSite.compatibleServersAndVms} />
            <HighlightStat icon={<ShieldCheck size={20} />} label="Confirmed installs" value={selectedServerSite.inventoryWithCyberSmart} tone="green" />
            <HighlightStat icon={<AlertTriangle size={20} />} label="Remaining installs" value={selectedServerSite.inventoryWithoutCyberSmart} tone="red" />
            <HighlightStat icon={<Search size={20} />} label="Raw server-labelled agents" value={selectedServerSite.cyberSmartServerInstalls} tone="amber" />
          </div>
          <div className="server-coverage-visual">
            <div className="coverage-bar-heading">
              <span>Confirmed infrastructure coverage</span>
              <strong>{serverCoveragePercent}%</strong>
            </div>
            <div className="coverage-bar">
              <span style={{ width: `${serverCoveragePercent}%` }} />
            </div>
            <div className="coverage-legend">
              <span><i className="confirmed" /> {formatNumber.format(selectedServerSite.inventoryWithCyberSmart)} confirmed</span>
              <span><i className="remaining" /> {formatNumber.format(selectedServerSite.inventoryWithoutCyberSmart)} remaining</span>
              <span>{formatNumber.format(selectedServerSite.physicalServers)} physical · {formatNumber.format(selectedServerSite.virtualMachines)} virtual</span>
              <button className="secondary" type="button" onClick={() => navigate('servers')}>Open server report <ChevronRight size={16} /></button>
            </div>
          </div>
        </section>
      )}
    </>
  )
}

function SiteExplorer({
  core, installed, failing, remaining, servers, centralDevices,
}: {
  core: CoreData; installed: InstalledRow[] | null; failing: FailingRow[] | null; remaining: RemainingRow[] | null; servers: ServerData | null
  centralDevices: CyberDeviceRow[] | null
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
  const showCentral = site.site === 'Morley Meadow'
  const centralPassing = centralDevices?.filter((row) => row.overallStatus === 'Pass').length ?? 0
  const centralFailing = centralDevices?.filter((row) => row.overallStatus === 'Fail').length ?? 0
  const centralSerials = centralDevices?.filter((row) => Boolean(row.serial)).length ?? 0
  const centralMorleyMatches = centralDevices?.filter((row) => row.morleyInventoryMatch === 'Matched').length ?? 0
  const centralDeviceTypes = Object.entries((centralDevices ?? []).reduce<Record<string, number>>((result, row) => {
    result[row.deviceType || 'Not recorded'] = (result[row.deviceType || 'Not recorded'] ?? 0) + 1
    return result
  }, {})).sort((a, b) => b[1] - a[1])
  const centralFailedControls = Object.entries((centralDevices ?? []).reduce<Record<string, number>>((result, row) => {
    for (const control of String(row.failedChecks ?? '').split(';').map((value) => value.trim()).filter(Boolean)) {
      result[control] = (result[control] ?? 0) + 1
    }
    return result
  }, {})).sort((a, b) => b[1] - a[1])
  return (
    <>
      <Header title="Complete site report" description="Every headline and evidence list for one site: total compute, matched and unmatched installations, compliance, remaining rollout, servers and failed controls." />
      <ReadingGuide items={[
        'Choose a school first; every headline, breakdown and evidence table below updates to that site.',
        'Compatible clients, servers and VMs are separate denominators and should not be added together when judging client coverage.',
        'Matched installations have stronger inventory evidence; unmatched installations require hostname or serial reconciliation.',
        'Remaining allocation is the planned rollout quantity, while failed controls identify remediation needed on devices already installed.',
      ]} />
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
        <Metric label="Total compatible servers & VMs" value={siteServerSummary?.compatibleServersAndVms ?? siteServerInventory.length} note={`${siteServerSummary?.physicalServers ?? 0} physical · ${siteServerSummary?.virtualMachines ?? 0} virtual`} tone="purple" />
        <Metric label="Confirmed CyberSmart server installs" value={siteServerSummary?.inventoryWithCyberSmart ?? 0} note="Unique infrastructure records with a confirmed match" tone="green" />
        <Metric label="Remaining server CyberSmart installs" value={siteServerSummary?.inventoryWithoutCyberSmart ?? 0} note="Compatible infrastructure without a confirmed match" tone="red" />
        <Metric label="Server-labelled agents (raw)" value={site.serverInstalls} note="Reconciliation evidence only" tone="amber" />
      </div>
      <h2 className="section-title">Site compliance headlines</h2>
      <div className="metrics">
        <Metric label="Passing devices" value={site.passing} tone="green" />
        <Metric label="Failing devices" value={site.failing} tone="red" />
        <Metric label="Pass percentage" value={`${formatPercent.format(site.passPercent)}%`} />
        <Metric label="Possible unmanaged" value={site.possibleUnmanaged} tone="amber" />
        <Metric label="Passing server agents" value={siteServerSummary?.passingInstalls ?? 0} tone="green" />
        <Metric label="Failing server agents" value={siteServerSummary?.failingInstalls ?? 0} tone="red" />
      </div>
      {showCentral && (
        <section className="panel associated-site-report">
          <div className="associated-site-heading">
            <div>
              <p className="eyebrow">Associated CyberSmart report</p>
              <h2>West Country Schools Trust / Central</h2>
              <p>Central is operationally connected to Morley Meadow, so its CyberSmart evidence is shown on this page. Its agents have also been checked against Morley’s inventory. No Central audit inventory is available, therefore Central coverage percentage and remaining-install counts cannot yet be calculated.</p>
            </div>
            <Badge value="Evidence only — no inventory denominator" />
          </div>
          {!centralDevices ? <Loader /> : (
            <>
              <div className="highlight-grid">
                <HighlightStat icon={<MonitorCheck size={20} />} label="CyberSmart agents" value={centralDevices.length} />
                <HighlightStat icon={<ShieldCheck size={20} />} label="Passing agents" value={centralPassing} tone="green" />
                <HighlightStat icon={<ShieldAlert size={20} />} label="Failing agents" value={centralFailing} tone="red" />
                <HighlightStat icon={<Search size={20} />} label="Serial recorded" value={centralSerials} tone="amber" />
                <HighlightStat icon={<CheckCircle2 size={20} />} label="Matched to Morley inventory" value={centralMorleyMatches} tone={centralMorleyMatches ? 'green' : 'slate'} />
              </div>
              <div className="notice central-match-notice">
                <AlertTriangle size={18} />
                <p><strong>Morley reconciliation:</strong> {formatNumber.format(centralMorleyMatches)} of {formatNumber.format(centralDevices.length)} Central agents match Morley’s {formatNumber.format(core.centralReport.morleyInventoryRecords)} inventory records by serial, asset tag or hostname evidence. Unmatched Central agents remain separate and must not reduce Morley’s remaining-install gap.</p>
              </div>
              <div className="two-column associated-breakdowns">
                <section className="chart-card">
                  <h3>Central devices by type</h3>
                  <div className="breakdown">{centralDeviceTypes.map(([label, value]) => <div key={label}><span>{label}</span><strong>{formatNumber.format(value)}</strong></div>)}</div>
                </section>
                <section className="chart-card">
                  <h3>Central failed controls</h3>
                  <div className="breakdown">
                    {centralFailedControls.length
                      ? centralFailedControls.map(([label, value]) => <div key={label}><span>{label}</span><strong>{formatNumber.format(value)}</strong></div>)
                      : <div><span>No failed controls recorded</span><strong>0</strong></div>}
                  </div>
                </section>
              </div>
              <details open>
                <summary>All WeST Central CyberSmart devices ({formatNumber.format(centralDevices.length)})</summary>
                <DataTable rows={centralDevices} columns={[
                  { key: 'host', label: 'Host' }, { key: 'serial', label: 'Serial' }, { key: 'deviceType', label: 'Type' },
                  { key: 'overallStatus', label: 'Compliance', render: (row) => <Badge value={row.overallStatus} /> },
                  { key: 'osName', label: 'Operating system' }, { key: 'osVersion', label: 'OS version' },
                  { key: 'failedChecks', label: 'Failed controls' }, { key: 'lastCheckIn', label: 'Last check-in' },
                  { key: 'morleyInventoryMatch', label: 'Morley match', render: (row) => <Badge value={row.morleyInventoryMatch} /> },
                  { key: 'morleyMatchMethod', label: 'Morley match method' }, { key: 'morleyInventoryCategory', label: 'Morley inventory category' },
                  { key: 'morleyInventorySerial', label: 'Morley serial' }, { key: 'morleyAssetTag', label: 'Morley asset tag' },
                ]} />
              </details>
            </>
          )}
        </section>
      )}
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
  title, description, guide, rows, columns, siteKey, typeKey, filename, metrics, ownerKey, failureKey, typeLabel,
}: {
  title: string; description: string; guide: string[]; rows: T[]; columns: Array<Column<T>>;
  siteKey: keyof T & string; typeKey: keyof T & string; filename: string; metrics?: ReactNode
  ownerKey?: keyof T & string; failureKey?: keyof T & string; typeLabel?: string
}) {
  return <><Header title={title} description={description} /><ReadingGuide items={guide} />{metrics}<FilteredTable rows={rows} columns={columns} siteKey={siteKey} typeKey={typeKey} filename={filename} ownerKey={ownerKey} failureKey={failureKey} typeLabel={typeLabel} /></>
}

type AssignedFailure = FailingRow & {
  owner: string; priority: string; remediationStatus: string; dueDate: string; notes: string
}

function Failures({ rows, core }: { rows: FailingRow[]; core: CoreData }) {
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
    const scheduledDueDate = core.taskProgramme.schoolSchedule.find((schedule) => schedule.site === row.schoolSite)?.dueDate ?? core.taskProgramme.centralDueDate
    return { ...row, owner: item?.owner || core.taskProgramme.owner, priority: item?.priority ?? 'Medium', remediationStatus: item?.status ?? 'Not started', dueDate: item?.dueDate || scheduledDueDate, notes: item?.notes ?? '' }
  })
  return (
    <BrowserPage
      title="Failures & remediation allocation"
      description="Locate the exact failing device, review its controls and assign remediation locally in this browser without modifying source evidence."
      guide={[
        'Each row is one installed device with at least one failed CyberSmart control.',
        'Failed controls states what needs fixing; hostname, serial, operating system and last check-in help identify the device.',
        'Owner, priority, status, due date and notes are editable planning fields stored only in this browser.',
        'Use the filters to isolate a school, owner, device type or specific failed control before exporting.',
      ]}
      rows={assigned} siteKey="schoolSite" typeKey="deviceType" filename="cybersmart-failures-and-remediation.csv"
      ownerKey="owner" failureKey="failedChecks"
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

function Programme({ core }: { core: CoreData }) {
  const tasks = core.taskProgramme.tasks
  const central = tasks.filter((task) => task.site === 'WeST Central')
  const mfa = tasks.filter((task) => task.category === 'MFA')
  const byod = tasks.filter((task) => task.category === 'BYOD')
  const rollout = tasks
    .filter((task) => task.category.startsWith('CyberSmart rollout'))
    .sort((a, b) => {
      const aDelegated = a.owner === core.taskProgramme.owner ? 1 : 0
      const bDelegated = b.owner === core.taskProgramme.owner ? 1 : 0
      return aDelegated - bDelegated || a.owner.localeCompare(b.owner) || a.site.localeCompare(b.site)
    })
  const delegated = Object.entries(core.taskProgramme.rolloutOwners).map(([site, owner]) => {
    const ownerTasks = rollout.filter((task) => task.site === site)
    return {
      site,
      owner,
      deviceTypes: ownerTasks.map((task) => task.category.replace('CyberSmart rollout - ', '')).join(', '),
      remainingDevices: ownerTasks.reduce((total, task) => total + Number(task.measure.match(/^\d+/)?.[0] ?? 0), 0),
      dueDate: ownerTasks[0]?.dueDate ?? '',
    }
  })
  return (
    <>
      <Header title="Remediation programme" description="A single owner and deadline-led delivery plan for Central, all schools, CyberSmart remediation, server coverage, MFA and BYOD." />
      <ReadingGuide items={[
        'Each programme row is a delivery task, not a device. The completion measure explains the quantity or evidence required.',
        'Installation tasks are split by school and device type; named local owners apply only where explicitly shown.',
        'Mark Kent owns all other schools and all non-installation work unless a task row names someone else.',
        'School deadlines are completion dates for the full local work package; Central and trust-wide work has its own earlier deadline.',
      ]} />
      <div className="metrics">
        <Metric label="Programme owner" value={core.taskProgramme.owner} />
        <Metric label="Installation task owners" value={new Set(tasks.filter((task) => task.category.startsWith('CyberSmart rollout')).map((task) => task.owner)).size} note="Mark Kent plus named local owners" />
        <Metric label="Total tasks" value={tasks.length} />
        <Metric label="Central / trust-wide tasks" value={central.length} note="Due 1 December 2026" tone="purple" />
        <Metric label="School tasks" value={tasks.length - central.length} note="Monthly waves January–July 2027" />
        <Metric label="MFA site tasks" value={mfa.length} note="One for every school" tone="amber" />
        <Metric label="Trust-wide BYOD tasks" value={byod.length} note="Policy and technical controls" tone="amber" />
      </div>
      <h2 className="section-title">Delegated remaining-install owners</h2>
      <section className="panel allocation-highlight">
        <p>These owners are responsible specifically for installing CyberSmart on the remaining compatible devices at their schools. Mark Kent owns all other schools and all non-installation work.</p>
        <DataTable rows={delegated} columns={[
          { key: 'owner', label: 'Installation task owner' }, { key: 'site', label: 'School' },
          { key: 'deviceTypes', label: 'Assigned device types' }, { key: 'remainingDevices', label: 'Remaining devices' },
          { key: 'dueDate', label: 'All installation tasks due' },
        ]} />
      </section>
      <h2 className="section-title">School completion schedule</h2>
      <section className="panel">
        <DataTable rows={core.taskProgramme.schoolSchedule.map((item) => ({
          ...item,
          installOwner: core.taskProgramme.rolloutOwners[item.site] ?? core.taskProgramme.owner,
          remainingDevices: core.sites.find((site) => site.site === item.site)?.remainingToInstall ?? 0,
          installTaskCount: rollout.filter((task) => task.site === item.site).length,
          taskCount: tasks.filter((task) => task.site === item.site).length,
        }))} columns={[
          { key: 'site', label: 'School' }, { key: 'installOwner', label: 'Remaining-install owner' },
          { key: 'remainingDevices', label: 'Remaining devices' }, { key: 'installTaskCount', label: 'Device-type install tasks' },
          { key: 'dueDate', label: 'All tasks due' }, { key: 'taskCount', label: 'Total tasks' },
        ]} />
      </section>
      <h2 className="section-title">Remaining CyberSmart installation tasks</h2>
      <FilteredTable<ProgrammeTask> rows={rollout} siteKey="site" typeKey="category" ownerKey="owner" typeLabel="device types" filename="remaining-cybersmart-installation-tasks.csv" columns={[
        { key: 'owner', label: 'Task owner' }, { key: 'site', label: 'School' }, { key: 'category', label: 'Device type' },
        { key: 'task', label: 'Installation task' }, { key: 'measure', label: 'Remaining quantity' },
        { key: 'dueDate', label: 'Due date' }, { key: 'status', label: 'Status', render: (row) => <Badge value={row.status} /> },
      ]} />
      <h2 className="section-title">All programme tasks</h2>
      <FilteredTable<ProgrammeTask> rows={tasks} siteKey="site" typeKey="category" ownerKey="owner" typeLabel="task types" filename="cybersmart-remediation-programme.csv" columns={[
        { key: 'scope', label: 'Scope' }, { key: 'site', label: 'Site' }, { key: 'category', label: 'Category' },
        { key: 'task', label: 'Task' }, { key: 'measure', label: 'Completion measure' }, { key: 'owner', label: 'Owner' },
        { key: 'dueDate', label: 'Due date' }, { key: 'status', label: 'Status', render: (row) => <Badge value={row.status} /> },
      ]} />
    </>
  )
}

type AssignedRemaining = RemainingRow & { taskOwner: string; dueDate: string }

function RemainingPage({ rows, core }: { rows: RemainingRow[]; core: CoreData }) {
  const assigned: AssignedRemaining[] = rows.map((row) => ({
    ...row,
    taskOwner: core.taskProgramme.rolloutOwners[row.schoolSite] ?? core.taskProgramme.owner,
    dueDate: core.taskProgramme.schoolSchedule.find((item) => item.site === row.schoolSite)?.dueDate ?? core.taskProgramme.schoolCompletionDeadline,
  }))
  return (
    <BrowserPage
      title="Remaining installation allocation"
      description="Every provisional remaining compatible-device allocation now includes its responsible task owner and school completion deadline."
      guide={[
        'Each row is an inventory item or quantity contributing to the remaining compatible-device rollout.',
        'Allocate is the number of devices represented by the row; it may be greater than one for quantity-only inventory.',
        'Task owner and school deadline come from the delivery programme and apply to installing that remaining allocation.',
        'Evidence confidence shows how precise the source inventory is; serialised records are stronger than quantity-only evidence.',
      ]}
      rows={assigned} siteKey="schoolSite" typeKey="deviceType" filename="remaining-cybersmart-installations-with-owners.csv"
      ownerKey="taskOwner"
      metrics={<div className="metrics compact">
        <Metric label="Remaining client devices" value={core.summary.remainingToInstall} tone="red" />
        <Metric label="Device-type installation tasks" value={core.taskProgramme.tasks.filter((task) => task.category.startsWith('CyberSmart rollout')).length} tone="amber" />
        <Metric label="Installation task owners" value={new Set(core.taskProgramme.tasks.filter((task) => task.category.startsWith('CyberSmart rollout')).map((task) => task.owner)).size} />
      </div>}
      columns={[
        { key: 'taskOwner', label: 'Task owner' }, { key: 'dueDate', label: 'School deadline' },
        { key: 'schoolSite', label: 'Site' }, { key: 'deviceCategory', label: 'Inventory category' }, { key: 'deviceType', label: 'Type' },
        { key: 'allocateQuantity', label: 'Allocate' }, { key: 'serial', label: 'Serial' }, { key: 'assetTag', label: 'Asset tag' },
        { key: 'operatingSystem', label: 'Operating system' }, { key: 'manufacturer', label: 'Manufacturer' }, { key: 'model', label: 'Model' },
        { key: 'location', label: 'Location' }, { key: 'evidenceConfidence', label: 'Evidence confidence' },
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
      <ReadingGuide items={[
        'Select All sites for trust-wide results or choose one school for a local control view.',
        'Assessed is the number of devices with evidence for that control; it is not always the full installed estate.',
        'Pass and Fail are explicit CyberSmart outcomes. Unknown means the export did not provide a usable result.',
        'Pass rate is Pass divided by Assessed, and the table is ordered with the highest failing counts first.',
      ]} />
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
      <ReadingGuide items={[
        'Total compatible servers and VMs is the authoritative infrastructure denominator.',
        'Confirmed installs counts unique inventory records with a reliable CyberSmart match; Remaining is the unmatched infrastructure balance.',
        'Raw server-labelled agents are reconciliation evidence only and may include stale, duplicate or differently reported records.',
        'Use the inventory table for coverage decisions and the raw-agent table to investigate matching and failed controls.',
      ]} />
      <div className="metrics">
        <Metric label="Total compatible servers & VMs" value={data.summary.inventoryRecords} note={`${data.summary.physicalServers} physical · ${data.summary.virtualMachines} virtual`} />
        <Metric label="Confirmed CyberSmart server installs" value={data.summary.inventoryWithCyberSmart} note="Unique matched infrastructure records" tone="green" />
        <Metric label="Remaining server CyberSmart installs" value={data.summary.inventoryWithoutCyberSmart} note="Infrastructure without a confirmed match" tone="red" />
        <Metric label="Server-labelled agents (raw)" value={data.summary.cyberSmartServerInstalls} note="Agent evidence, not confirmed coverage" tone="amber" />
        <Metric label="Matched agent records" value={data.summary.matchedServers} tone="green" />
        <Metric label="Unmatched agent records" value={data.summary.unmatchedServers} tone="amber" />
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
      <h2 className="section-title">Raw CyberSmart server-agent evidence</h2>
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
      <ReadingGuide items={[
        'Read the numbered cards in order: inventory denominator, installation match, remaining allocation, remediation and server separation.',
        'Counts describe a fixed evidence snapshot and may differ from live CyberSmart portal figures after the generation date.',
        'Use the dedicated pages and CSV exports when you need device-level evidence behind a headline.',
      ]} />
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
  const centralDevices = useDataset<CyberDeviceRow[]>('central-devices.json', page === 'sites')
  const navigate = (next: Page) => { setPage(next); setMenu(false); window.scrollTo({ top: 0, behavior: 'smooth' }) }

  if (!core.data) return <main className="startup"><ShieldCheck size={42} /><h1>Loading CyberSmart reporting</h1><Loader error={core.error} /></main>

  let content: ReactNode
  if (page === 'overview') content = <Overview core={core.data} navigate={navigate} />
  else if (page === 'programme') content = <Programme core={core.data} />
  else if (page === 'sites') content = <SiteExplorer core={core.data} installed={installed.data} failing={failing.data} remaining={remaining.data} servers={servers.data} centralDevices={centralDevices.data} />
  else if (page === 'inventory') content = inventory.data ? <BrowserPage title="Authoritative compute inventory" description="Every client-compute record that forms the site denominator, including serialised and quantity-only evidence." guide={[
    'Each row is a source inventory record; Quantity can represent one device or a grouped quantity.',
    'Compatible type is the normalised CyberSmart delivery category used for coverage and task planning.',
    'CyberSmart match shows whether the inventory record has reliable installation evidence; blank identifiers reduce matching confidence.',
    'Filter by site or type, then use serial, asset tag, model and location to validate individual devices.',
  ]} rows={inventory.data} siteKey="school" typeKey="compatibleType" filename="compatible-compute-inventory.csv" columns={[
    { key: 'school', label: 'Site' }, { key: 'deviceCategory', label: 'Inventory category' }, { key: 'compatibleType', label: 'Compatible type' },
    { key: 'quantity', label: 'Quantity' }, { key: 'serialNumber', label: 'Serial' }, { key: 'assetTag', label: 'Asset tag' },
    { key: 'manufacturer', label: 'Manufacturer' }, { key: 'model', label: 'Model' }, { key: 'operatingSystem', label: 'Operating system' },
    { key: 'exactCyberSmartMatch', label: 'CyberSmart match', render: (row) => <Badge value={row.exactCyberSmartMatch} /> }, { key: 'location', label: 'Location' },
  ]} /> : <Loader error={inventory.error} />
  else if (page === 'installed') content = installed.data ? <BrowserPage title="CyberSmart installations" description="All school installations split between exact serial matches and devices not matched by serial. Pass or fail does not alter match status." guide={[
    'Each row is one CyberSmart client agent reported for a school.',
    'Category and Inventory match describe evidence confidence; serial matched is stronger than hostname or unmatched evidence.',
    'Compliance is independent of inventory matching: a correctly matched device can still fail one or more controls.',
    'Failed controls lists the checks requiring remediation; use site and device-type filters to narrow the estate.',
  ]} rows={installed.data} siteKey="schoolSite" typeKey="deviceType" filename="cybersmart-installations.csv" columns={[
    { key: 'schoolSite', label: 'Site' }, { key: 'allocationCategory', label: 'Category' }, { key: 'host', label: 'Host' },
    { key: 'serial', label: 'Serial' }, { key: 'deviceType', label: 'Type' }, { key: 'overallStatus', label: 'Compliance', render: (row) => <Badge value={row.overallStatus} /> },
    { key: 'os', label: 'Operating system' }, { key: 'matchMethod', label: 'Inventory match' }, { key: 'failedChecks', label: 'Failed controls' },
  ]} /> : <Loader error={installed.error} />
  else if (page === 'failures') content = failing.data ? <Failures rows={failing.data} core={core.data} /> : <Loader error={failing.error} />
  else if (page === 'remaining') content = remaining.data ? <RemainingPage rows={remaining.data} core={core.data} /> : <Loader error={remaining.error} />
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
