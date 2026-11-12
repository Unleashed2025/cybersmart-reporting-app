$ErrorActionPreference = 'Stop'

$root = Split-Path $PSScriptRoot -Parent
$source = Join-Path $root 'source\WeST-CyberSmart-Final-Validated-Assessment.xlsx'
$output = Join-Path $root 'public\downloads\WeST-CyberSmart-Assessment-and-Programme-Tracker.xlsx'
$projectOutput = Join-Path $root 'public\downloads\WeST-CyberSmart-Project-Delivery-Plan.xlsx'
$corePath = Join-Path $root 'public\data\tracker.json'
$remainingPath = Join-Path $root 'public\data\remaining-devices.json'

if (-not (Test-Path -LiteralPath $source)) { throw "Source workbook not found: $source" }
if (-not (Test-Path -LiteralPath $corePath)) { throw "Core data not found: $corePath" }
if (-not (Test-Path -LiteralPath $remainingPath)) { throw "Remaining-device data not found: $remainingPath" }

$core = Get-Content -LiteralPath $corePath -Raw | ConvertFrom-Json
$remaining = Get-Content -LiteralPath $remainingPath -Raw | ConvertFrom-Json

Add-Type -AssemblyName System.Drawing
function Get-OleColour([string]$hex) {
  $colour = [System.Drawing.ColorTranslator]::FromHtml("#$hex")
  return [System.Drawing.ColorTranslator]::ToOle($colour)
}

$navy = Get-OleColour '0B1736'
$blue = Get-OleColour '2F6FED'
$pink = Get-OleColour 'FF2D7D'
$amber = Get-OleColour 'F4A62A'
$paleAmber = Get-OleColour 'FFF3DC'
$paleBlue = Get-OleColour 'EAF1FF'
$white = Get-OleColour 'FFFFFF'
$muted = Get-OleColour '64748B'
$line = Get-OleColour 'D9E1EC'

function New-TrackingSheet {
  param(
    $Workbook,
    [string]$Name,
    [string]$Title,
    [string]$Description,
    [string[]]$Headers,
    [object[]]$Rows,
    [double[]]$Widths,
    [int[]]$EditableColumns,
    [int]$TabColour
  )

  foreach ($existing in @($Workbook.Worksheets)) {
    if ($existing.Name -eq $Name) {
      $existing.Delete()
      break
    }
  }

  $sheet = $Workbook.Worksheets.Add()
  $sheet.Name = $Name
  $sheet.Tab.Color = $TabColour
  $columnCount = $Headers.Count
  $rowCount = $Rows.Count

  $sheet.Range($sheet.Cells(1, 1), $sheet.Cells(1, $columnCount)).Merge()
  $sheet.Cells(1, 1).Value2 = $Title
  $sheet.Cells(1, 1).Font.Name = 'Aptos Display'
  $sheet.Cells(1, 1).Font.Size = 20
  $sheet.Cells(1, 1).Font.Bold = $true
  $sheet.Cells(1, 1).Font.Color = $white
  $sheet.Cells(1, 1).Interior.Color = $navy
  $sheet.Rows.Item(1).RowHeight = 32

  $sheet.Range($sheet.Cells(2, 1), $sheet.Cells(2, $columnCount)).Merge()
  $sheet.Cells(2, 1).Value2 = $Description
  $sheet.Cells(2, 1).Font.Name = 'Aptos'
  $sheet.Cells(2, 1).Font.Size = 10
  $sheet.Cells(2, 1).Font.Color = $muted
  $sheet.Cells(2, 1).WrapText = $true
  $sheet.Rows.Item(2).RowHeight = 30

  $sheet.Range($sheet.Cells(3, 1), $sheet.Cells(3, $columnCount)).Merge()
  $sheet.Cells(3, 1).Value2 = 'Yellow cells are editable progress fields. Filters are enabled on every column.'
  $sheet.Cells(3, 1).Font.Name = 'Aptos'
  $sheet.Cells(3, 1).Font.Size = 9
  $sheet.Cells(3, 1).Font.Italic = $true
  $sheet.Cells(3, 1).Interior.Color = $paleAmber

  for ($column = 1; $column -le $columnCount; $column++) {
    $cell = $sheet.Cells(4, $column)
    $cell.Value2 = $Headers[$column - 1]
    $cell.Font.Name = 'Aptos'
    $cell.Font.Size = 9
    $cell.Font.Bold = $true
    $cell.Font.Color = $white
    $cell.Interior.Color = $blue
    $cell.WrapText = $true
    $cell.VerticalAlignment = -4108
    $sheet.Columns.Item($column).ColumnWidth = $Widths[$column - 1]
  }
  $sheet.Rows.Item(4).RowHeight = 30

  if ($rowCount -gt 0) {
    for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
      $row = $Rows[$rowIndex]
      for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) {
        $cell = $sheet.Cells.Item(5 + $rowIndex, 1 + $columnIndex)
        $value = $row[$columnIndex]
        if ($null -eq $value) {
          $cell.Value2 = ''
        } elseif ($value -is [byte] -or $value -is [int16] -or $value -is [int32] -or $value -is [int64] -or $value -is [single] -or $value -is [double] -or $value -is [decimal]) {
          $cell.Value2 = [double]$value
        } else {
          $cell.Value2 = [string]$value
        }
      }
    }
    $body = $sheet.Range($sheet.Cells(5, 1), $sheet.Cells(4 + $rowCount, $columnCount))
    $body.Font.Name = 'Aptos'
    $body.Font.Size = 9
    $body.WrapText = $true
    $body.VerticalAlignment = -4160
    $body.Borders.Item(9).Color = $line
    $body.RowHeight = 27

    foreach ($column in $EditableColumns) {
      $sheet.Range($sheet.Cells(5, $column), $sheet.Cells(4 + $rowCount, $column)).Interior.Color = $paleAmber
    }
  }

  $sheet.Range($sheet.Cells(4, 1), $sheet.Cells([Math]::Max(4, 4 + $rowCount), $columnCount)).AutoFilter() | Out-Null
  $sheet.Activate()
  $sheet.Application.ActiveWindow.SplitRow = 4
  $sheet.Application.ActiveWindow.FreezePanes = $true
  return $sheet
}

function Get-SafeSheetName {
  param([string]$Name, [hashtable]$UsedNames)
  $safe = $Name -replace '[:\\/\?\*\[\]]', ' '
  $safe = ($safe -replace '\s+', ' ').Trim()
  if ($safe.Length -gt 31) { $safe = $safe.Substring(0, 31).Trim() }
  $candidate = $safe
  $suffix = 2
  while ($UsedNames.ContainsKey($candidate)) {
    $suffixText = " $suffix"
    $baseLength = [Math]::Min($safe.Length, 31 - $suffixText.Length)
    $candidate = $safe.Substring(0, $baseLength).Trim() + $suffixText
    $suffix++
  }
  $UsedNames[$candidate] = $true
  return $candidate
}

$programmeRows = New-Object System.Collections.ArrayList
$taskIndex = 0
foreach ($task in $core.taskProgramme.tasks) {
  $taskIndex++
  [void]$programmeRows.Add(@(
    ('CS-{0:D3}' -f $taskIndex),
    $task.scope,
    $task.site,
    $task.category,
    $task.task,
    $task.measure,
    $task.owner,
    $task.dueDate,
    $task.status,
    0,
    '',
    'Green',
    '',
    '',
    ''
  ))
}

$schoolRows = New-Object System.Collections.ArrayList
foreach ($site in $core.sites) {
  $siteTasks = @($core.taskProgramme.tasks | Where-Object { $_.site -eq $site.site })
  $schedule = $core.taskProgramme.schoolSchedule | Where-Object { $_.site -eq $site.site } | Select-Object -First 1
  $server = $core.serverSites | Where-Object { $_.site -eq $site.site } | Select-Object -First 1
  $installOwner = $core.taskProgramme.owner
  if ($core.taskProgramme.rolloutOwners.PSObject.Properties.Name -contains $site.site) {
    $installOwner = $core.taskProgramme.rolloutOwners.($site.site)
  }
  $clientInstalls = [Math]::Max(0, [int]$site.cyberSmartDevices - [int]$site.serverInstalls)
  $unmatchedClients = [Math]::Max(0, [int]$site.notSerialMatched - [int]$site.serverInstalls)
  [void]$schoolRows.Add(@(
    $site.site,
    $installOwner,
    $schedule.dueDate,
    [int]$site.compatibleDevices,
    $clientInstalls,
    [int]$site.serialMatched,
    $unmatchedClients,
    [int]$site.remainingToInstall,
    [int]$site.knownRemaining,
    [int]$site.explicitUnmanagedRemaining,
    [int]$site.unallocatedInventoryGap,
    @($siteTasks | Where-Object { $_.category -like 'CyberSmart rollout*' }).Count,
    @($siteTasks | Where-Object { $_.category -eq 'Compliance remediation' }).Count,
    @($siteTasks | Where-Object { $_.category -eq 'Inventory reconciliation' }).Count,
    @($siteTasks | Where-Object { $_.category -eq 'Server CyberSmart' }).Count,
    @($siteTasks | Where-Object { $_.category -eq 'MFA' }).Count,
    $siteTasks.Count,
    [int]$server.inventoryWithCyberSmart,
    [int]$server.inventoryWithoutCyberSmart,
    0,
    'Not started',
    'Green',
    ''
  ))
}

$rolloutRows = New-Object System.Collections.ArrayList
$taskIndex = 0
foreach ($task in $core.taskProgramme.tasks) {
  $taskIndex++
  if ($task.category -notlike 'CyberSmart rollout*') { continue }
  $quantity = 0
  if ([string]$task.measure -match '^(\d+)') { $quantity = [int]$Matches[1] }
  [void]$rolloutRows.Add(@(
    ('CS-{0:D3}' -f $taskIndex),
    $task.site,
    ([string]$task.category).Replace('CyberSmart rollout - ', ''),
    $quantity,
    $task.owner,
    $task.dueDate,
    $task.status,
    0,
    '',
    '',
    '',
    ''
  ))
}

$uncertainRows = New-Object System.Collections.ArrayList
foreach ($row in $remaining) {
  $isGap = $row.deviceCategory -eq 'Unallocated inventory gap' -or [string]$row.evidenceConfidence -like 'Gap exceeds*'
  $isTotalsBased = [string]$row.evidenceConfidence -like 'Totals-based*'
  $isUnmanaged = [string]$row.deviceCategory -like '*unmanaged*' -or [string]$row.deviceType -like '*unknown*'
  if (-not ($isGap -or $isTotalsBased -or $isUnmanaged)) { continue }

  $classification = if ($isGap) { 'Unallocated / uncertain gap' } elseif ($isUnmanaged) { 'Explicit unmanaged / unknown inventory' } else { 'Quantity-only / no serial match' }
  $owner = $core.taskProgramme.owner
  if ($core.taskProgramme.rolloutOwners.PSObject.Properties.Name -contains $row.schoolSite) {
    $owner = $core.taskProgramme.rolloutOwners.($row.schoolSite)
  }
  $schedule = $core.taskProgramme.schoolSchedule | Where-Object { $_.site -eq $row.schoolSite } | Select-Object -First 1
  $explanation = if ($isGap) {
    'Calculated coverage shortfall. No individual source asset can be named; identifiers remain blank pending local reconciliation.'
  } elseif ($isUnmanaged) {
    'Source inventory explicitly identifies this row as unmanaged or unknown.'
  } else {
    'Quantity-based inventory evidence without an individual serial match.'
  }
  [void]$uncertainRows.Add(@(
    $classification,
    $row.schoolSite,
    $row.deviceCategory,
    $row.deviceType,
    [int]$row.allocateQuantity,
    $row.serial,
    $row.assetTag,
    $row.operatingSystem,
    $row.manufacturer,
    $row.model,
    $row.location,
    $row.evidenceConfidence,
    $row.allocationStatus,
    $explanation,
    $owner,
    $schedule.dueDate,
    'Not started',
    0,
    '',
    ''
  ))
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Force }
  $workbook = $excel.Workbooks.Open($source)
  $workbook.SaveAs($output, 51)

  $programmeSheet = New-TrackingSheet -Workbook $workbook -Name 'Programme Tracker' -Title 'CyberSmart Programme Tracker' -Description 'The complete app task allocation with editable status, progress, completion, blocker and evidence fields.' -Headers @(
    'Task ID','Scope','Site','Category','Task','Completion measure','Owner','Due date','Status','Progress %','Actual completion date','RAG','Blocker / dependency','Latest update / evidence','Last updated'
  ) -Rows $programmeRows.ToArray() -Widths @(12,14,31,25,50,34,20,14,16,12,18,12,28,42,16) -EditableColumns @(9,10,11,13,14,15) -TabColour $pink
  $programmeSheet.Columns.Item(10).NumberFormat = '0"%"'

  $schoolSheet = New-TrackingSheet -Workbook $workbook -Name 'School Progress' -Title 'School CyberSmart Progress' -Description 'One row per school, including known remaining inventory, explicit unmanaged items and unallocated or uncertain gaps.' -Headers @(
    'Site','Remaining-install owner','Deadline','Compatible clients','CyberSmart client installs','Serial matched','Without serial match','Remaining clients','Known remaining','Explicit unmanaged remaining','Unallocated / uncertain gap','Install tasks','Compliance tasks','Inventory tasks','Server tasks','MFA tasks','Total tasks','Confirmed server installs','Remaining server installs','Average progress %','Overall status','RAG','Latest update'
  ) -Rows $schoolRows.ToArray() -Widths @(31,23,14,15,20,14,18,16,16,20,20,13,15,14,12,11,12,18,18,17,16,11,36) -EditableColumns @(23) -TabColour $blue
  $schoolSheet.Columns.Item(20).NumberFormat = '0"%"'

  $rolloutSheet = New-TrackingSheet -Workbook $workbook -Name 'Device Rollout Tracker' -Title 'CyberSmart Device Rollout Tracker' -Description 'A site and device-type view of the remaining compatible-device installation work.' -Headers @(
    'Task ID','Site','Device type','Remaining quantity','Owner','Due date','Status','Progress %','Actual completion date','Blocker / dependency','Latest update / evidence','Last updated'
  ) -Rows $rolloutRows.ToArray() -Widths @(12,31,22,18,20,14,16,12,18,28,42,16) -EditableColumns @(7,8,9,10,11,12) -TabColour $amber
  $rolloutSheet.Columns.Item(8).NumberFormat = '0"%"'

  $uncertainSheet = New-TrackingSheet -Workbook $workbook -Name 'Unallocated & Uncertain' -Title 'Unallocated and Uncertain Inventory' -Description 'Every source/allocation row classed as explicitly unmanaged, quantity-only without serial evidence, or an unallocated calculated gap. Blank identifiers are intentionally preserved.' -Headers @(
    'Classification','Site','Inventory category','Device type','Quantity','Serial','Asset tag','Operating system','Manufacturer','Model','Location','Evidence confidence','Allocation status','Explanation','Owner','Deadline','Status','Progress %','Latest update / evidence','Last updated'
  ) -Rows $uncertainRows.ToArray() -Widths @(30,31,25,20,11,18,16,24,18,22,20,34,28,55,20,14,16,12,42,16) -EditableColumns @(17,18,19,20) -TabColour $amber
  $uncertainSheet.Columns.Item(18).NumberFormat = '0"%"'

  $workbook.Save()
  $sheetCount = $workbook.Worksheets.Count
  $workbook.Close($true)

  if (Test-Path -LiteralPath $projectOutput) { Remove-Item -LiteralPath $projectOutput -Force }
  $projectWorkbook = $excel.Workbooks.Add()
  $defaultSheetNames = @($projectWorkbook.Worksheets | ForEach-Object { $_.Name })

  [void](New-TrackingSheet -Workbook $projectWorkbook -Name 'Portfolio Overview' -Title 'WeST CyberSmart Project Delivery Plan' -Description 'One row per school showing coverage, remaining inventory classifications, task allocation, owners, deadlines and progress.' -Headers @(
    'Site','Remaining-install owner','Deadline','Compatible clients','CyberSmart client installs','Serial matched','Without serial match','Remaining clients','Known remaining','Explicit unmanaged remaining','Unallocated / uncertain gap','Install tasks','Compliance tasks','Inventory tasks','Server tasks','MFA tasks','Total tasks','Confirmed server installs','Remaining server installs','Average progress %','Overall status','RAG','Latest update'
  ) -Rows $schoolRows.ToArray() -Widths @(31,23,14,15,20,14,18,16,16,20,20,13,15,14,12,11,12,18,18,17,16,11,36) -EditableColumns @(20,21,22,23) -TabColour $pink)

  [void](New-TrackingSheet -Workbook $projectWorkbook -Name 'Task Register' -Title 'Trust-wide CyberSmart Task Register' -Description 'Every central, trust-wide and school task with completion measure, owner, timescale and editable delivery tracking.' -Headers @(
    'Task ID','Scope','Site','Category','Task','Completion measure','Owner','Due date','Status','Progress %','Actual completion date','RAG','Blocker / dependency','Latest update / evidence','Last updated'
  ) -Rows $programmeRows.ToArray() -Widths @(12,14,31,25,50,34,20,14,16,12,18,12,28,42,16) -EditableColumns @(9,10,11,12,13,14,15) -TabColour $blue)

  [void](New-TrackingSheet -Workbook $projectWorkbook -Name 'Device Rollout' -Title 'Remaining Device Rollout' -Description 'The installation work grouped by school and device type, including owner, deadline and progress fields.' -Headers @(
    'Task ID','Site','Device type','Remaining quantity','Owner','Due date','Status','Progress %','Actual completion date','Blocker / dependency','Latest update / evidence','Last updated'
  ) -Rows $rolloutRows.ToArray() -Widths @(12,31,22,18,20,14,16,12,18,28,42,16) -EditableColumns @(7,8,9,10,11,12) -TabColour $amber)

  $waveRows = New-Object System.Collections.ArrayList
  foreach ($schedule in $core.taskProgramme.schoolSchedule) {
    $site = $core.sites | Where-Object { $_.site -eq $schedule.site } | Select-Object -First 1
    $siteTasks = @($core.taskProgramme.tasks | Where-Object { $_.site -eq $schedule.site })
    $owner = $core.taskProgramme.owner
    if ($core.taskProgramme.rolloutOwners.PSObject.Properties.Name -contains $schedule.site) {
      $owner = $core.taskProgramme.rolloutOwners.($schedule.site)
    }
    [void]$waveRows.Add(@(
      $schedule.dueDate,
      $schedule.site,
      $owner,
      [int]$site.remainingToInstall,
      [int]$site.knownRemaining,
      [int]$site.explicitUnmanagedRemaining,
      [int]$site.unallocatedInventoryGap,
      $siteTasks.Count,
      'Not started',
      0,
      '',
      ''
    ))
  }
  [void](New-TrackingSheet -Workbook $projectWorkbook -Name 'Completion Waves' -Title 'School Completion Waves' -Description 'The planned school delivery sequence from January to July 2027, with owners, remaining quantities and progress tracking.' -Headers @(
    'Completion deadline','Site','Remaining-install owner','Remaining devices','Known remaining','Explicit unmanaged','Unallocated / uncertain','Total tasks','Status','Progress %','Blocker / dependency','Latest update'
  ) -Rows $waveRows.ToArray() -Widths @(18,31,23,17,16,18,20,12,16,12,28,42) -EditableColumns @(9,10,11,12) -TabColour $blue)

  [void](New-TrackingSheet -Workbook $projectWorkbook -Name 'Unallocated & Uncertain' -Title 'Unallocated and Uncertain Inventory' -Description 'Every relevant line from the source allocation: unmanaged rows, quantity-only rows and calculated gaps with blank identifiers preserved.' -Headers @(
    'Classification','Site','Inventory category','Device type','Quantity','Serial','Asset tag','Operating system','Manufacturer','Model','Location','Evidence confidence','Allocation status','Explanation','Owner','Deadline','Status','Progress %','Latest update / evidence','Last updated'
  ) -Rows $uncertainRows.ToArray() -Widths @(30,31,25,20,11,18,16,24,18,22,20,34,28,55,20,14,16,12,42,16) -EditableColumns @(17,18,19,20) -TabColour $amber)

  $usedSheetNames = @{}
  foreach ($sheetName in @('Portfolio Overview','Task Register','Device Rollout','Completion Waves','Unallocated & Uncertain')) {
    $usedSheetNames[$sheetName] = $true
  }
  foreach ($site in $core.sites) {
    $siteTasks = @($core.taskProgramme.tasks | Where-Object { $_.site -eq $site.site })
    $siteTaskRows = New-Object System.Collections.ArrayList
    foreach ($task in $siteTasks) {
      [void]$siteTaskRows.Add(@(
        $task.category,
        $task.task,
        $task.measure,
        $task.owner,
        $task.dueDate,
        $task.status,
        0,
        '',
        '',
        '',
        ''
      ))
    }
    $siteSheetName = Get-SafeSheetName -Name $site.site -UsedNames $usedSheetNames
    $siteDescription = "Compatible clients: $($site.compatibleDevices). Remaining: $($site.remainingToInstall), comprising $($site.knownRemaining) known, $($site.explicitUnmanagedRemaining) explicit unmanaged and $($site.unallocatedInventoryGap) unallocated or uncertain. All tasks use the site completion timescale."
    [void](New-TrackingSheet -Workbook $projectWorkbook -Name $siteSheetName -Title "$($site.site) Delivery Plan" -Description $siteDescription -Headers @(
      'Task category','Task','Completion measure','Owner','Due date','Status','Progress %','Actual completion date','Blocker / dependency','Latest update / evidence','Last updated'
    ) -Rows $siteTaskRows.ToArray() -Widths @(27,52,36,22,14,16,12,18,28,42,16) -EditableColumns @(6,7,8,9,10,11) -TabColour $paleBlue)
  }

  foreach ($sheetName in $defaultSheetNames) {
    foreach ($sheet in @($projectWorkbook.Worksheets)) {
      if ($sheet.Name -eq $sheetName) {
        $sheet.Delete()
        break
      }
    }
  }
  $projectWorkbook.SaveAs($projectOutput, 51)
  $projectSheetCount = $projectWorkbook.Worksheets.Count
  $projectWorkbook.Close($true)

  Write-Output (@{
    output = $output
    sheets = $sheetCount
    projectOutput = $projectOutput
    projectSheets = $projectSheetCount
    programmeTasks = $programmeRows.Count
    schools = $schoolRows.Count
    rolloutTasks = $rolloutRows.Count
    uncertainRows = $uncertainRows.Count
  } | ConvertTo-Json)
}
finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
