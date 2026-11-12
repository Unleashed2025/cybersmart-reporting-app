import fs from 'node:fs'
import path from 'node:path'
import PptxGenJS from 'pptxgenjs'

const root = path.resolve(import.meta.dirname, '..')
const core = JSON.parse(fs.readFileSync(path.join(root, 'public', 'data', 'core.json'), 'utf8'))
const remainingDevices = JSON.parse(fs.readFileSync(path.join(root, 'public', 'data', 'remaining-devices.json'), 'utf8'))
const output = path.join(root, 'public', 'downloads', 'WeST-CyberSmart-Multisite-Programme.pptx')
const projectOutput = path.join(root, 'public', 'downloads', 'WeST-CyberSmart-Project-Delivery-Plan.pptx')

const pptx = new PptxGenJS()
pptx.layout = 'LAYOUT_WIDE'
pptx.author = 'Unleashed'
pptx.company = 'Unleashed'
pptx.subject = 'CyberSmart multisite coverage and remediation programme'
pptx.title = 'WeST CyberSmart Multisite Programme'
pptx.lang = 'en-GB'
pptx.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos', lang: 'en-GB' }

const W = 13.333
const H = 7.5
const C = {
  navy: '0B1736',
  ink: '14213D',
  muted: '64748B',
  bg: 'F4F7FB',
  white: 'FFFFFF',
  pink: 'FF2D7D',
  blue: '2F6FED',
  cyan: '00BFEA',
  purple: '7E3AED',
  green: '19A974',
  amber: 'F4A62A',
  red: 'E5484D',
  line: 'D9E1EC',
  paleBlue: 'EAF1FF',
  paleGreen: 'E6F7F1',
  paleAmber: 'FFF3DC',
  paleRed: 'FDE8E9',
}
const nf = new Intl.NumberFormat('en-GB')
let page = 0

function nextSlide() {
  page += 1
  return pptx.addSlide()
}

function addFooter(slide, scope = 'WeST CyberSmart') {
  slide.addText('U N L E A S H E D', {
    x: 0.48, y: 7.08, w: 2, h: 0.16, fontSize: 6.5, bold: true,
    color: C.navy, charSpacing: 2, margin: 0,
  })
  slide.addText(scope.toUpperCase(), {
    x: 9.1, y: 7.08, w: 3.45, h: 0.16, fontSize: 5.5,
    color: C.muted, align: 'right', margin: 0, fit: 'shrink',
  })
  slide.addText(String(page), {
    x: 12.75, y: 7.08, w: 0.25, h: 0.16, fontSize: 6,
    color: C.muted, align: 'right', margin: 0,
  })
}

function addHeader(slide, eyebrow, title, subtitle, scope = 'WeST CyberSmart') {
  slide.background = { color: C.bg }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: W, h: 0.13, fill: { color: C.pink }, line: { color: C.pink },
  })
  slide.addText(eyebrow.toUpperCase(), {
    x: 0.55, y: 0.42, w: 7, h: 0.18, fontSize: 6.5, bold: true,
    color: C.pink, charSpacing: 2, margin: 0,
  })
  slide.addText(title, {
    x: 0.55, y: 0.72, w: 12.1, h: 0.46, fontFace: 'Aptos Display',
    fontSize: 22, bold: true, color: C.navy, margin: 0, fit: 'shrink',
  })
  slide.addText(subtitle, {
    x: 0.55, y: 1.22, w: 12.1, h: 0.3, fontSize: 8.5,
    color: C.muted, margin: 0, fit: 'shrink',
  })
  addFooter(slide, scope)
}

function addCard(slide, x, y, w, h, accent = C.blue, fill = C.white) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.04, y: y + 0.05, w, h, rectRadius: 0.06,
    fill: { color: '000000', transparency: 92 }, line: { color: '000000', transparency: 100 },
  })
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06, fill: { color: fill }, line: { color: C.line, width: 0.6 },
  })
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w: 0.09, h, fill: { color: accent }, line: { color: accent },
  })
}

function addMetric(slide, x, y, w, value, label, detail, accent = C.blue) {
  addCard(slide, x, y, w, 1.18, accent)
  slide.addText(String(value), {
    x: x + 0.24, y: y + 0.18, w: w - 0.38, h: 0.34,
    fontFace: 'Aptos Display', fontSize: 20, bold: true, color: C.navy, margin: 0, fit: 'shrink',
  })
  slide.addText(label, {
    x: x + 0.24, y: y + 0.58, w: w - 0.38, h: 0.18,
    fontSize: 7.3, bold: true, color: C.ink, margin: 0, fit: 'shrink',
  })
  slide.addText(detail, {
    x: x + 0.24, y: y + 0.83, w: w - 0.38, h: 0.16,
    fontSize: 6.5, color: C.muted, margin: 0, fit: 'shrink',
  })
}

function addSectionLabel(slide, text, x, y, w = 5) {
  slide.addText(text.toUpperCase(), {
    x, y, w, h: 0.2, fontSize: 7, bold: true, color: C.pink, charSpacing: 1.3, margin: 0,
  })
}

function addBars(slide, items, { x, y, w, h, color = C.blue, valueColor = C.navy, maxItems = 12 }) {
  const rows = items.slice(0, maxItems)
  const max = Math.max(1, ...rows.map((item) => item.value))
  const rowH = h / Math.max(1, rows.length)
  rows.forEach((item, index) => {
    const rowY = y + index * rowH
    slide.addText(item.label, {
      x, y: rowY + 0.02, w: w * 0.32, h: rowH - 0.03, fontSize: 6.6,
      color: C.ink, margin: 0, valign: 'mid', fit: 'shrink',
    })
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + w * 0.34, y: rowY + rowH * 0.28, w: w * 0.53, h: Math.max(0.08, rowH * 0.35),
      rectRadius: 0.03, fill: { color: C.line }, line: { color: C.line },
    })
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + w * 0.34, y: rowY + rowH * 0.28,
      w: Math.max(0.03, w * 0.53 * item.value / max), h: Math.max(0.08, rowH * 0.35),
      rectRadius: 0.03, fill: { color: item.color ?? color }, line: { color: item.color ?? color },
    })
    slide.addText(nf.format(item.value), {
      x: x + w * 0.89, y: rowY + 0.02, w: w * 0.1, h: rowH - 0.03,
      fontSize: 7, bold: true, color: valueColor, align: 'right', margin: 0, valign: 'mid',
    })
  })
}

function addTable(slide, headers, rows, { x = 0.55, y = 1.68, w = 12.2, h = 4.95, fontSize = 6.5 } = {}) {
  const weights = headers.length === 11
    ? [2.35, 1.45, 1, 0.85, 2.15, 0.8, 0.8, 0.8, 0.8, 0.7, 0.75]
    : headers.length === 10
    ? [2.5, 1, 1, 1, 1, 1, 1, 0.65, 1.75, 1.2]
    : headers.length === 4
      ? [2.5, 1.45, 1.75, 0.65]
      : headers.length === 3
        ? [2.2, 2.45, 1.25]
        : headers.map(() => 1)
  const weightTotal = weights.reduce((sum, value) => sum + value, 0)
  const widths = weights.map((value) => w * value / weightTotal)
  const allRows = [headers, ...rows]
  const rowH = h / allRows.length
  allRows.forEach((row, rowIndex) => {
    let cellX = x
    row.forEach((value, columnIndex) => {
      const cellW = widths[columnIndex] ?? w / headers.length
      const header = rowIndex === 0
      slide.addShape(pptx.ShapeType.rect, {
        x: cellX,
        y: y + rowIndex * rowH,
        w: cellW,
        h: rowH,
        fill: { color: header ? C.navy : rowIndex % 2 === 1 ? C.white : 'F8FAFD' },
        line: { color: header ? C.navy : C.line, width: 0.45 },
      })
      slide.addText(String(value ?? ''), {
        x: cellX + 0.05,
        y: y + rowIndex * rowH + 0.02,
        w: Math.max(0.05, cellW - 0.1),
        h: Math.max(0.05, rowH - 0.04),
        fontFace: 'Aptos',
        fontSize: header ? fontSize : Math.max(4.5, fontSize - 0.3),
        bold: header,
        color: header ? C.white : C.ink,
        margin: 0,
        valign: 'mid',
        breakLine: false,
        fit: 'shrink',
      })
      cellX += cellW
    })
  })
}

function dateLabel(value) {
  return new Date(`${value}T00:00:00Z`).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' })
}

function shortSite(value) {
  return value
    .replace(' Community College for WeST', '')
    .replace(' School for WeST', '')
    .replace(' Primary School', '')
    .replace(' Community Primary School', '')
    .replace(' Community College', '')
}

// Cover
{
  const slide = nextSlide()
  slide.background = { color: C.navy }
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.2, h: H, fill: { color: C.pink }, line: { color: C.pink } })
  slide.addShape(pptx.ShapeType.ellipse, { x: 10.6, y: 0.35, w: 2.8, h: 2.8, fill: { color: C.cyan, transparency: 76 }, line: { color: C.cyan, transparency: 100 } })
  slide.addShape(pptx.ShapeType.ellipse, { x: 11.45, y: 1.1, w: 1.7, h: 1.7, fill: { color: C.purple, transparency: 45 }, line: { color: C.purple, transparency: 100 } })
  slide.addShape(pptx.ShapeType.ellipse, { x: 10.25, y: 2.15, w: 0.9, h: 0.9, fill: { color: C.pink, transparency: 15 }, line: { color: C.pink, transparency: 100 } })
  slide.addText('U N L E A S H E D', { x: 0.85, y: 0.8, w: 2.3, h: 0.2, fontSize: 8, bold: true, charSpacing: 2.3, color: C.white, margin: 0 })
  slide.addText('WEST CYBERSMART', { x: 0.85, y: 1.72, w: 7, h: 0.2, fontSize: 8, bold: true, charSpacing: 2, color: C.pink, margin: 0 })
  slide.addText('Multisite coverage &\nremediation programme', {
    x: 0.85, y: 2.18, w: 8.9, h: 1.45, fontFace: 'Aptos Display',
    fontSize: 33, bold: true, color: C.white, margin: 0, breakLine: false, fit: 'shrink',
  })
  slide.addText('Client compute · CyberSmart installs · Servers & VMs · Compliance · Tasks · Owners · Deadlines', {
    x: 0.85, y: 4.12, w: 9.5, h: 0.42, fontSize: 12, color: 'CAD5E5', margin: 0,
  })
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.85, y: 5.05, w: 4.1, h: 0.8, rectRadius: 0.05, fill: { color: C.pink }, line: { color: C.pink } })
  slide.addText(`${core.sites.length} schools · ${nf.format(core.taskProgramme.tasks.length)} programme tasks`, {
    x: 1.12, y: 5.28, w: 3.55, h: 0.25, fontSize: 11, bold: true, color: C.white, align: 'center', margin: 0,
  })
  slide.addText(`Evidence snapshot ${new Date(core.generatedAt).toLocaleDateString('en-GB')}`, {
    x: 0.85, y: 6.55, w: 4.5, h: 0.2, fontSize: 7.5, color: '94A3B8', margin: 0,
  })
}

// Executive summary
{
  const slide = nextSlide()
  addHeader(slide, 'Executive overview', 'The estate, the gap and the delivery plan', 'A single trust-wide view across client devices, server infrastructure, CyberSmart evidence and assigned remediation.')
  const s = core.summary
  addMetric(slide, 0.55, 1.75, 2.35, nf.format(core.clientCompute.quantity), 'ALL CLIENT-COMPUTE ITEMS', `${nf.format(core.clientCompute.records)} source records`, C.navy)
  addMetric(slide, 3.02, 1.75, 2.35, nf.format(s.compatibleDevices), 'COMPATIBLE CLIENT DEVICES', 'Coverage denominator', C.blue)
  addMetric(slide, 5.49, 1.75, 2.35, nf.format(s.schoolCyberSmartDevices), 'SCHOOL CYBERSMART DEVICES', 'Includes server agents', C.cyan)
  addMetric(slide, 7.96, 1.75, 2.35, nf.format(s.remainingToInstall), 'REMAINING CLIENT INSTALLS', 'Allocated by site and type', C.amber)
  addMetric(slide, 10.43, 1.75, 2.35, nf.format(core.serverSummary.inventoryRecords), 'TOTAL SERVERS & VMS', `${core.serverSummary.inventoryWithCyberSmart} confirmed installs · ${core.serverSummary.inventoryWithoutCyberSmart} remaining`, C.purple)
  addMetric(slide, 0.55, 3.15, 2.35, nf.format(s.csvPassing), 'PASSING DEVICES', 'All detailed checks pass', C.green)
  addMetric(slide, 3.02, 3.15, 2.35, nf.format(s.csvFailing), 'FAILING DEVICES', 'Detailed remediation evidence', C.red)
  addMetric(slide, 5.49, 3.15, 2.35, nf.format(s.serialMatched), 'SERIAL-MATCHED INSTALLS', 'Status-neutral exact matches', C.green)
  addMetric(slide, 7.96, 3.15, 2.35, nf.format(s.notSerialMatched), 'NOT SERIAL MATCHED', 'Includes server agents', C.amber)
  addMetric(slide, 10.43, 3.15, 2.35, nf.format(core.taskProgramme.tasks.length), 'PROGRAMME TASKS', `${new Set(core.taskProgramme.tasks.map((task) => task.owner)).size} assigned owners`, C.pink)
  addCard(slide, 0.55, 4.75, 12.23, 1.55, C.pink, C.white)
  slide.addText('Delivery commitment', { x: 0.85, y: 5.02, w: 2.3, h: 0.25, fontSize: 13, bold: true, color: C.navy, margin: 0 })
  slide.addText('WeST Central and trust-wide work', { x: 3.3, y: 4.98, w: 2.1, h: 0.18, fontSize: 7, bold: true, color: C.muted, margin: 0 })
  slide.addText('1 DEC 2026', { x: 3.3, y: 5.28, w: 2.1, h: 0.35, fontSize: 20, bold: true, color: C.pink, margin: 0 })
  slide.addText('First school wave', { x: 6.05, y: 4.98, w: 1.8, h: 0.18, fontSize: 7, bold: true, color: C.muted, margin: 0 })
  slide.addText('31 JAN 2027', { x: 6.05, y: 5.28, w: 2.1, h: 0.35, fontSize: 20, bold: true, color: C.blue, margin: 0 })
  slide.addText('All schools complete', { x: 9.15, y: 4.98, w: 1.8, h: 0.18, fontSize: 7, bold: true, color: C.muted, margin: 0 })
  slide.addText('31 JUL 2027', { x: 9.15, y: 5.28, w: 2.1, h: 0.35, fontSize: 20, bold: true, color: C.green, margin: 0 })
}

// Client coverage across sites
{
  const slide = nextSlide()
  addHeader(slide, 'How to read this report', 'Five rules for interpreting the numbers', 'Use these definitions consistently across trust-wide, school, server and programme slides.')
  const guide = [
    ['1', 'Start with the denominator', 'Compatible client devices are the client coverage denominator. Servers and VMs always remain separate.'],
    ['2', 'Read installation evidence', 'Serial matched is the strongest client evidence. Without serial match means an agent exists but needs reconciliation.'],
    ['3', 'Understand remaining quantities', 'Remaining installs are compatible devices without a confirmed allocation and are split by school and device type.'],
    ['4', 'Separate coverage from compliance', 'Installed does not mean passing. Pass and fail come from the exported CyberSmart control results.'],
    ['5', 'Read tasks as delivery work', 'Task columns count actions, not devices. The completion measure gives quantity; owner, due date and progress drive delivery.'],
  ]
  guide.forEach(([number, title, body], index) => {
    const y = 1.72 + index * 0.96
    addCard(slide, 0.68, y, 11.98, 0.78, index === 4 ? C.pink : C.blue, C.white)
    slide.addShape(pptx.ShapeType.ellipse, { x: 0.94, y: y + 0.16, w: 0.42, h: 0.42, fill: { color: index === 4 ? C.pink : C.blue }, line: { color: index === 4 ? C.pink : C.blue } })
    slide.addText(number, { x: 1.04, y: y + 0.25, w: 0.22, h: 0.16, fontSize: 8, bold: true, color: C.white, align: 'center', margin: 0 })
    slide.addText(title, { x: 1.58, y: y + 0.13, w: 2.45, h: 0.22, fontSize: 10, bold: true, color: C.navy, margin: 0 })
    slide.addText(body, { x: 4.08, y: y + 0.13, w: 8.18, h: 0.42, fontSize: 8.5, color: C.muted, margin: 0, fit: 'shrink', valign: 'mid' })
  })
}

// Client coverage across sites
{
  const slide = nextSlide()
  const central = core.centralReport
  addHeader(slide, 'Associated site evidence', 'WeST Central CyberSmart report', 'Shown with Morley Meadow because the operations are connected. Central has no independent audit inventory, so this is agent evidence rather than measured coverage.')
  addMetric(slide, 0.55, 1.72, 2.35, nf.format(central.totalAgents), 'CENTRAL CYBERSMART AGENTS', 'Detailed exported agent records', C.blue)
  addMetric(slide, 3.02, 1.72, 2.35, nf.format(central.passingAgents), 'PASSING AGENTS', 'All assessed controls pass', C.green)
  addMetric(slide, 5.49, 1.72, 2.35, nf.format(central.failingAgents), 'FAILING AGENTS', 'Control remediation required', C.red)
  addMetric(slide, 7.96, 1.72, 2.35, nf.format(central.serialRecorded), 'SERIAL RECORDED', 'Identifier present in CyberSmart', C.amber)
  addMetric(slide, 10.43, 1.72, 2.35, nf.format(central.matchedToMorley), 'MATCHED TO MORLEY', `${central.morleyInventoryRecords} Morley records checked`, central.matchedToMorley ? C.green : C.amber)
  addSectionLabel(slide, 'Central device types', 0.65, 3.2)
  addBars(slide, Object.entries(central.deviceTypes).map(([label, value]) => ({ label, value, color: C.blue })), { x: 0.65, y: 3.45, w: 5.65, h: 2.25, maxItems: 8 })
  addSectionLabel(slide, 'Central failed controls', 6.7, 3.2)
  addBars(slide, Object.entries(central.failedControls).map(([label, value]) => ({ label, value, color: C.red })), { x: 6.7, y: 3.45, w: 5.95, h: 2.25, maxItems: 8 })
  addCard(slide, 0.65, 5.93, 12, 0.66, C.amber, C.white)
  slide.addText(`Morley reconciliation result: ${central.matchedToMorley} of ${central.totalAgents} Central agents matched Morley Meadow inventory by serial, asset tag or hostname evidence. Central agents remain separate and do not reduce Morley’s installation gap.`, {
    x: 0.92, y: 6.12, w: 11.45, h: 0.24, fontSize: 8.3, bold: true, color: C.navy, margin: 0, fit: 'shrink',
  })
}

// Client coverage across sites
for (const [part, sites] of [[1, core.sites.slice(0, 16)], [2, core.sites.slice(16)]]) {
  const slide = nextSlide()
  addHeader(slide, 'Multisite client coverage', `Compatible devices and remaining installations · ${part} of 2`, 'The blue bar shows the compatible client estate; the amber bar shows the remaining CyberSmart installation allocation.')
  addBars(slide, sites.map((site) => ({ label: shortSite(site.site), value: site.compatibleDevices, color: C.blue })), { x: 0.65, y: 1.75, w: 5.9, h: 4.9, color: C.blue, maxItems: 16 })
  addBars(slide, sites.map((site) => ({ label: shortSite(site.site), value: site.remainingToInstall, color: C.amber })), { x: 6.8, y: 1.75, w: 5.9, h: 4.9, color: C.amber, maxItems: 16 })
  addSectionLabel(slide, 'Compatible client devices', 0.65, 1.55)
  addSectionLabel(slide, 'Remaining CyberSmart installs', 6.8, 1.55)
}

// Device type gap
{
  const slide = nextSlide()
  addHeader(slide, 'Installation gap', 'Remaining compatible devices by technology', 'Trust-wide totals built from every site allocation. Unknown and unmanaged items require local identification before deployment.')
  const totals = {}
  for (const site of core.sites) {
    for (const [type, value] of Object.entries(site.remainingByType)) totals[type] = (totals[type] ?? 0) + value
  }
  const labels = {
    Windows: 'Windows', Chromebook: 'Chromebook', AppleMac: 'Apple / Mac',
    IPadIOs: 'iPad / iOS', AndroidTablet: 'Android / tablet', PossibleTypeUnknown: 'Unknown / unmanaged',
  }
  addBars(slide, Object.entries(totals).map(([type, value]) => ({ label: labels[type] ?? type, value, color: type === 'PossibleTypeUnknown' ? C.amber : C.blue })).sort((a, b) => b.value - a.value), { x: 0.8, y: 1.95, w: 7.3, h: 3.7, maxItems: 8 })
  addCard(slide, 8.55, 1.85, 3.95, 4.3, C.pink, C.white)
  slide.addText('What the task programme creates', { x: 8.88, y: 2.15, w: 3.25, h: 0.45, fontSize: 17, bold: true, color: C.navy, margin: 0, fit: 'shrink' })
  slide.addText([
    { text: `${core.taskProgramme.tasks.filter((task) => task.category.startsWith('CyberSmart rollout')).length} device-type installation tasks\n`, options: { bold: true, color: C.blue } },
    { text: 'Separate task for each site and remaining technology cohort\n\n', options: { color: C.muted } },
    { text: 'Exact quantity, owner and school deadline\n', options: { bold: true, color: C.green } },
    { text: 'Windows · Chromebook · Apple/Mac · iPad/iOS · Android/tablet · unknown/unmanaged', options: { color: C.muted } },
  ], { x: 8.88, y: 2.9, w: 3.25, h: 2.45, fontSize: 11, breakLine: false, margin: 0, valign: 'top', fit: 'shrink' })
}

// Server coverage
{
  const slide = nextSlide()
  addHeader(slide, 'Inventory reconciliation', 'Unallocated and uncertain remaining inventory', 'The Excel tracker contains every relevant row. This slide shows the scale and representative examples without inventing asset identities.')
  const explicitUnmanaged = core.sites.reduce((total, site) => total + site.explicitUnmanagedRemaining, 0)
  const unallocatedGap = core.sites.reduce((total, site) => total + site.unallocatedInventoryGap, 0)
  const totalsBased = remainingDevices
    .filter((row) => String(row.evidenceConfidence).startsWith('Totals-based'))
    .reduce((total, row) => total + Number(row.allocateQuantity || 0), 0)
  addMetric(slide, 0.65, 1.72, 3.65, nf.format(explicitUnmanaged), 'EXPLICIT UNMANAGED / UNKNOWN', 'Source inventory labels these rows unmanaged or unknown', C.amber)
  addMetric(slide, 4.52, 1.72, 3.65, nf.format(totalsBased), 'QUANTITY-ONLY EVIDENCE', 'No individual serial match is available', C.purple)
  addMetric(slide, 8.39, 1.72, 3.65, nf.format(unallocatedGap), 'UNALLOCATED / UNCERTAIN GAP', 'Calculated shortfall with no item-level identity', C.red)

  const gapExamples = remainingDevices.filter((row) => row.deviceCategory === 'Unallocated inventory gap').slice(0, 2)
  const unmanagedExamples = remainingDevices.filter((row) => String(row.deviceCategory).toLowerCase().includes('unmanaged')).slice(0, 2)
  const totalsExamples = remainingDevices.filter((row) => String(row.evidenceConfidence).startsWith('Totals-based')).slice(0, 2)
  const exampleRows = [...gapExamples, ...unmanagedExamples, ...totalsExamples].map((row) => {
    const classification = row.deviceCategory === 'Unallocated inventory gap'
      ? 'Unallocated gap'
      : String(row.deviceCategory).toLowerCase().includes('unmanaged')
        ? 'Explicit unmanaged'
        : 'Quantity-only'
    const identifier = row.serial || row.assetTag || 'No item-level identifier'
    const action = row.deviceCategory === 'Unallocated inventory gap'
      ? 'Reconcile local inventory before allocation'
      : row.allocationStatus
    return [shortSite(row.schoolSite), classification, row.deviceCategory, row.allocateQuantity, identifier, action]
  })
  addSectionLabel(slide, 'Representative source / allocation rows', 0.65, 3.18, 8)
  addTable(slide,
    ['Site', 'Classification', 'Inventory category', 'Qty', 'Serial / asset evidence', 'Required action'],
    exampleRows,
    { x: 0.65, y: 3.45, w: 12.05, h: 2.75, fontSize: 6.2 },
  )
  slide.addText('Important: Ivybridge’s 438 is one aggregate gap row with blank serial, asset, model and location fields. It is not 438 blank spreadsheet rows and must not be presented as 438 known unmanaged devices.', {
    x: 0.75, y: 6.35, w: 11.85, h: 0.3, fontSize: 7.5, bold: true, color: C.red, margin: 0, fit: 'shrink',
  })
}

// Server coverage
{
  const slide = nextSlide()
  addHeader(slide, 'Server CyberSmart', 'Compatible server and VM coverage', 'A separate infrastructure report: server agents never reduce the client-compute installation gap.')
  const ss = core.serverSummary
  addMetric(slide, 0.55, 1.72, 2.35, nf.format(ss.inventoryRecords), 'TOTAL COMPATIBLE SERVERS & VMS', `${ss.physicalServers} physical · ${ss.virtualMachines} virtual`, C.purple)
  addMetric(slide, 3.02, 1.72, 2.35, nf.format(ss.inventoryWithCyberSmart), 'CONFIRMED CYBERSMART INSTALLS', 'Unique matched infrastructure records', C.green)
  addMetric(slide, 5.49, 1.72, 2.35, nf.format(ss.inventoryWithoutCyberSmart), 'REMAINING SERVER INSTALLS', 'No confirmed CyberSmart match', C.red)
  addMetric(slide, 7.96, 1.72, 2.35, nf.format(ss.cyberSmartServerInstalls), 'RAW SERVER-LABELLED AGENTS', 'Reconciliation evidence only', C.amber)
  addMetric(slide, 10.43, 1.72, 2.35, nf.format(ss.unmatchedServers), 'UNMATCHED AGENT RECORDS', 'Duplicates, stale or unmapped evidence', C.amber)
  const top = [...core.serverSites].sort((a, b) => b.compatibleServersAndVms - a.compatibleServersAndVms).slice(0, 12)
  addSectionLabel(slide, 'Largest server and VM estates', 0.65, 3.2)
  addBars(slide, top.map((site) => ({ label: shortSite(site.site), value: site.compatibleServersAndVms, color: C.purple })), { x: 0.65, y: 3.45, w: 6.0, h: 3.0, color: C.purple })
  addSectionLabel(slide, 'Confirmed server installs by site', 6.85, 3.2)
  addBars(slide, [...core.serverSites].sort((a, b) => b.inventoryWithCyberSmart - a.inventoryWithCyberSmart).slice(0, 12).map((site) => ({ label: shortSite(site.site), value: site.inventoryWithCyberSmart, color: C.green })), { x: 6.85, y: 3.45, w: 5.85, h: 3.0, color: C.green })
}

// Programme ownership
{
  const slide = nextSlide()
  addHeader(slide, 'Remediation programme', 'Task ownership and workload', 'Mark Kent owns the programme and all non-delegated work. Named local owners receive only remaining compatible-device installation tasks.')
  const ownerCounts = Object.entries(core.taskProgramme.tasks.reduce((result, task) => {
    result[task.owner] = (result[task.owner] ?? 0) + 1
    return result
  }, {})).map(([label, value]) => ({ label, value, color: label === 'Mark Kent' ? C.pink : C.blue })).sort((a, b) => b.value - a.value)
  addBars(slide, ownerCounts, { x: 0.75, y: 1.9, w: 6.1, h: 3.8, color: C.blue, maxItems: 12 })
  addCard(slide, 7.25, 1.8, 5.15, 4.35, C.pink, C.white)
  slide.addText('Ownership rules', { x: 7.58, y: 2.12, w: 4.4, h: 0.35, fontSize: 18, bold: true, color: C.navy, margin: 0 })
  const delegated = Object.entries(core.taskProgramme.rolloutOwners).map(([site, owner]) => `${owner} — ${shortSite(site)}`)
  slide.addText(`Programme owner\nMARK KENT\n\nDelegated remaining-install owners\n${delegated.join('\n')}`, {
    x: 7.58, y: 2.68, w: 4.35, h: 2.95, fontSize: 10, color: C.muted,
    bold: false, breakLine: false, margin: 0, fit: 'shrink',
  })
}

// Completion waves
{
  const slide = nextSlide()
  addHeader(slide, 'Delivery schedule', 'Monthly school completion waves', 'Every task for a school shares its completion date. WeST Central and trust-wide tasks complete first.')
  const waves = Object.entries(core.taskProgramme.schoolSchedule.reduce((result, item) => {
    result[item.dueDate] = [...(result[item.dueDate] ?? []), item.site]
    return result
  }, {}))
  const maxCount = Math.max(...waves.map(([, schools]) => schools.length))
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.65, y: 1.75, w: 1.35, h: 4.55, rectRadius: 0.06, fill: { color: C.navy }, line: { color: C.navy } })
  slide.addText('WEST\nCENTRAL', { x: 0.82, y: 2.1, w: 1.0, h: 0.65, fontSize: 16, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink' })
  slide.addText('1 DEC\n2026', { x: 0.82, y: 3.03, w: 1.0, h: 0.65, fontSize: 18, bold: true, color: C.pink, align: 'center', margin: 0 })
  slide.addText(`${core.taskProgramme.tasks.filter((task) => task.site === 'WeST Central').length} tasks`, { x: 0.82, y: 4.2, w: 1.0, h: 0.25, fontSize: 9, color: 'CBD5E1', align: 'center', margin: 0 })
  waves.forEach(([dueDate, schools], index) => {
    const x = 2.25 + index * 1.47
    const height = 1.5 + 2.35 * schools.length / maxCount
    const y = 6.15 - height
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w: 1.18, h: height, rectRadius: 0.05, fill: { color: index % 2 === 0 ? C.blue : C.purple }, line: { color: index % 2 === 0 ? C.blue : C.purple } })
    slide.addText(String(schools.length), { x: x + 0.15, y: y + 0.25, w: 0.88, h: 0.4, fontSize: 24, bold: true, color: C.white, align: 'center', margin: 0 })
    slide.addText('schools', { x: x + 0.15, y: y + 0.75, w: 0.88, h: 0.2, fontSize: 8, color: C.white, align: 'center', margin: 0 })
    slide.addText(dateLabel(dueDate).replace(' 2027', ''), { x, y: 6.28, w: 1.18, h: 0.24, fontSize: 7.5, bold: true, color: C.navy, align: 'center', margin: 0 })
  })
}

// All schools summary tables
const summaryRows = core.sites.map((site) => {
  const server = core.serverSites.find((item) => item.site === site.site)
  const schedule = core.taskProgramme.schoolSchedule.find((item) => item.site === site.site)
  const tasks = core.taskProgramme.tasks.filter((task) => task.site === site.site)
  const installOwner = core.taskProgramme.rolloutOwners[site.site] ?? core.taskProgramme.owner
  const clientInstalls = Math.max(0, site.cyberSmartDevices - site.serverInstalls)
  return [
    shortSite(site.site), site.compatibleDevices, clientInstalls, site.serialMatched,
    site.remainingToInstall, server?.compatibleServersAndVms ?? 0, `${server?.inventoryWithCyberSmart ?? 0} / ${server?.inventoryWithoutCyberSmart ?? 0}`,
    tasks.length, installOwner, dateLabel(schedule?.dueDate),
  ]
})
for (let index = 0; index < summaryRows.length; index += 11) {
  const slide = nextSlide()
  addHeader(slide, 'Multisite management view', `All-school coverage, tasks and ownership · ${Math.floor(index / 11) + 1} of ${Math.ceil(summaryRows.length / 11)}`, 'Client and server coverage, remaining tasks, local installation ownership and completion dates.')
  addTable(slide,
    ['School', 'Compatible\nclients', 'CyberSmart client\ninstalls', 'Serial\nmatched', 'Remaining\nclients', 'Servers\n& VMs', 'Server installs\nconfirmed / remaining', 'Tasks', 'Install owner', 'All tasks due'],
    summaryRows.slice(index, index + 11),
    { y: 1.68, h: 4.95, fontSize: 6.2 },
  )
}

// School task matrix
const taskMatrixRows = core.sites.map((site) => {
  const tasks = core.taskProgramme.tasks.filter((task) => task.site === site.site)
  const schedule = core.taskProgramme.schoolSchedule.find((item) => item.site === site.site)
  return [
    shortSite(site.site),
    core.taskProgramme.rolloutOwners[site.site] ?? core.taskProgramme.owner,
    dateLabel(schedule?.dueDate),
    site.remainingToInstall,
    `${site.knownRemaining} known / ${site.explicitUnmanagedRemaining} unmanaged / ${site.unallocatedInventoryGap} uncertain`,
    tasks.filter((task) => task.category.startsWith('CyberSmart rollout')).length,
    tasks.filter((task) => task.category === 'Compliance remediation').length,
    tasks.filter((task) => task.category === 'Inventory reconciliation').length,
    tasks.filter((task) => task.category === 'Server CyberSmart').length,
    tasks.filter((task) => task.category === 'MFA').length,
    tasks.length,
  ]
})
for (let index = 0; index < taskMatrixRows.length; index += 11) {
  const slide = nextSlide()
  addHeader(slide, 'Delivery management', `School task types, owners and deadlines · ${Math.floor(index / 11) + 1} of ${Math.ceil(taskMatrixRows.length / 11)}`, 'Remaining devices is a device quantity; the task columns count actions. Install owner applies to rollout tasks, while the deadline covers the full school work package.')
  addTable(slide,
    ['School', 'Install owner', 'Deadline', 'Remaining\ndevices', 'Remaining breakdown\nknown / unmanaged / uncertain', 'Install\ntasks', 'Compliance', 'Inventory', 'Server', 'MFA', 'Total\ntasks'],
    taskMatrixRows.slice(index, index + 11),
    { y: 1.68, h: 4.95, fontSize: 6.2 },
  )
}

// Site detail slides
for (const site of core.sites) {
  const slide = nextSlide()
  const server = core.serverSites.find((item) => item.site === site.site)
  const schedule = core.taskProgramme.schoolSchedule.find((item) => item.site === site.site)
  const tasks = core.taskProgramme.tasks.filter((task) => task.site === site.site)
  const installOwner = core.taskProgramme.rolloutOwners[site.site] ?? core.taskProgramme.owner
  const clientInstalls = Math.max(0, site.cyberSmartDevices - site.serverInstalls)
  const unmatchedClients = Math.max(0, site.notSerialMatched - site.serverInstalls)
  addHeader(slide, 'School delivery plan', site.site, `${tasks.length} tasks · programme owner ${core.taskProgramme.owner} · remaining-install owner ${installOwner} · all tasks due ${dateLabel(schedule?.dueDate)}`, site.site)
  addMetric(slide, 0.55, 1.65, 1.9, nf.format(site.compatibleDevices), 'COMPATIBLE CLIENTS', 'Authoritative denominator', C.blue)
  addMetric(slide, 2.58, 1.65, 1.9, nf.format(clientInstalls), 'CYBERSMART CLIENT INSTALLS', `${site.serialMatched} serial matched`, C.cyan)
  addMetric(slide, 4.61, 1.65, 1.9, nf.format(site.remainingToInstall), 'REMAINING INSTALLS', `${unmatchedClients} client installs unmatched`, C.amber)
  addMetric(slide, 6.64, 1.65, 1.9, nf.format(site.failing), 'FAILING DEVICES', `${site.passing} passing`, C.red)
  addMetric(slide, 8.67, 1.65, 1.9, nf.format(server?.compatibleServersAndVms ?? 0), 'TOTAL SERVERS & VMS', `${server?.physicalServers ?? 0} physical · ${server?.virtualMachines ?? 0} virtual`, C.purple)
  addMetric(slide, 10.7, 1.65, 1.9, nf.format(server?.inventoryWithCyberSmart ?? 0), 'CONFIRMED SERVER INSTALLS', `${server?.inventoryWithoutCyberSmart ?? 0} remaining · ${server?.cyberSmartServerInstalls ?? 0} raw agents`, C.green)
  addSectionLabel(slide, 'Remaining client installs by device type', 0.6, 3.08)
  const labels = { Windows: 'Windows', Chromebook: 'Chromebook', AppleMac: 'Apple / Mac', IPadIOs: 'iPad / iOS', AndroidTablet: 'Android / tablet', PossibleTypeUnknown: 'Unknown / unmanaged' }
  addBars(slide, Object.entries(site.remainingByType).filter(([, value]) => value > 0).map(([type, value]) => ({ label: labels[type] ?? type, value, color: type === 'PossibleTypeUnknown' ? C.amber : C.blue })), { x: 0.65, y: 3.35, w: 4.6, h: 2.7, maxItems: 7 })
  addSectionLabel(slide, 'Assigned tasks', 5.52, 3.08)
  addTable(slide, ['Task type', 'Completion measure', 'Owner'], tasks.map((task) => [task.category, task.measure, task.owner]), { x: 5.52, y: 3.35, w: 7.15, h: 2.75, fontSize: 5.8 })
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.65, y: 6.25, w: 12.02, h: 0.42, rectRadius: 0.04, fill: { color: C.navy }, line: { color: C.navy } })
  slide.addText(`DELIVERY DEADLINE  ${dateLabel(schedule?.dueDate).toUpperCase()}     ·     PROGRAMME OWNER  ${core.taskProgramme.owner.toUpperCase()}     ·     INSTALL OWNER  ${installOwner.toUpperCase()}`, {
    x: 0.92, y: 6.37, w: 11.48, h: 0.15, fontSize: 7.2, bold: true, color: C.white, align: 'center', margin: 0, fit: 'shrink',
  })
}

// Closing slide
{
  const slide = nextSlide()
  slide.background = { color: C.navy }
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.15, fill: { color: C.pink }, line: { color: C.pink } })
  slide.addText('THE DELIVERY OUTCOME', { x: 0.85, y: 1.05, w: 5, h: 0.2, fontSize: 8, bold: true, color: C.pink, charSpacing: 2, margin: 0 })
  slide.addText('Every compatible device covered.\nEvery failure owned.\nEvery school complete.', {
    x: 0.85, y: 1.62, w: 8.4, h: 2.1, fontFace: 'Aptos Display', fontSize: 32, bold: true, color: C.white, margin: 0, breakLine: false,
  })
  slide.addText('WeST Central by 1 December 2026 · first school wave by 31 January 2027 · all schools by 31 July 2027', {
    x: 0.85, y: 4.2, w: 10.5, h: 0.45, fontSize: 13, color: 'CBD5E1', margin: 0,
  })
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.85, y: 5.18, w: 4.35, h: 0.85, rectRadius: 0.06, fill: { color: C.pink }, line: { color: C.pink } })
  slide.addText(`${nf.format(core.taskProgramme.tasks.length)} tasks · ${new Set(core.taskProgramme.tasks.map((task) => task.owner)).size} owners · ${core.sites.length} schools`, {
    x: 1.1, y: 5.45, w: 3.85, h: 0.26, fontSize: 12, bold: true, color: C.white, align: 'center', margin: 0,
  })
  slide.addText('U N L E A S H E D', { x: 0.85, y: 6.75, w: 2.3, h: 0.2, fontSize: 8, bold: true, charSpacing: 2.3, color: C.white, margin: 0 })
}

fs.mkdirSync(path.dirname(output), { recursive: true })
await pptx.writeFile({ fileName: output, compression: true })
fs.copyFileSync(output, projectOutput)
console.log(JSON.stringify({
  output,
  projectOutput,
  slides: page,
  bytes: fs.statSync(output).size,
}, null, 2))
