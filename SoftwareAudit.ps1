<#
Portable Software Audit Utility - Installed Apps Only, Original Colors
- Uses Windows Installed Apps inventory sources only.
- Does not scan Program Files, Start Menu shortcuts, AppsFolder, or arbitrary executable files.
- Requires Microsoft Excel on the audited PC.
#>

#requires -version 5.1
$ErrorActionPreference = "Stop"

# Original workbook colors. These are Excel/VBA RGB() integer values.
# Blue and green are the same light-tinted Office theme colors used in the reference workbook.
$OriginalBlueColor   = 15653023  # RGB(159,216,238) - software/version columns
$OriginalGreenColor  = 9364099   # RGB(131,226,142) - installed/same-version asset cells
$OriginalYellowColor = 65535     # RGB(255,255,0)   - version mismatch cells
$OriginalRedColor    = 255       # RGB(255,0,0)     - missing software cells

$FilterWindowsDefaultsAndFrameworks = $true

function Normalize-Name {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    $n = $Name.ToLowerInvariant()
    $n = $n -replace '&', 'and'
    $n = $n -replace '\+', 'plus'
    $n = $n -replace '[^a-z0-9]+', ''
    return $n
}

function Normalize-VersionText {
    param($Version)
    if ($null -eq $Version) { return "" }
    return ($Version.ToString().Trim() -replace '\s+', ' ')
}

function Test-NonEmptyVersion {
    param($Version)
    $v = Normalize-VersionText $Version
    if ([string]::IsNullOrWhiteSpace($v)) { return $false }
    if ($v -match '^(installed|unknown|n/a|null|0|0\.0)$') { return $false }
    return $true
}

function Remove-TrailingZeroVersionSegments {
    param([string]$Version)
    $v = Normalize-VersionText $Version
    if ($v -notmatch '^\d+(\.\d+)+$') { return $v.ToLowerInvariant() }
    while ($v -match '^(.+)\.0$') { $v = $matches[1] }
    return $v.ToLowerInvariant()
}

function Test-VersionMatch {
    param([string]$Expected, [string]$Installed)
    $e = Normalize-VersionText $Expected
    $i = Normalize-VersionText $Installed
    if ([string]::IsNullOrWhiteSpace($e) -or [string]::IsNullOrWhiteSpace($i)) { return $false }
    if ($i -eq 'Installed') { return $false }
    if ($e.Equals($i, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    if ((Remove-TrailingZeroVersionSegments $e) -eq (Remove-TrailingZeroVersionSegments $i)) { return $true }
    return $false
}

function Convert-PackedMsiVersion {
    param($PackedVersion)
    if ($null -eq $PackedVersion) { return "" }
    try {
        [UInt32]$num = [UInt32]$PackedVersion
        if ($num -eq 0) { return "" }
        $major = ($num -shr 24) -band 0xFF
        $minor = ($num -shr 16) -band 0xFF
        $build = $num -band 0xFFFF
        if ($major -eq 0 -and $minor -eq 0 -and $build -eq 0) { return "" }
        return "$major.$minor.$build"
    } catch { return "" }
}

function Get-RegistryDisplayVersion {
    param($Item)
    foreach ($p in @('DisplayVersion','BundleVersion','ProductVersion')) {
        try { if (Test-NonEmptyVersion $Item.$p) { return (Normalize-VersionText $Item.$p) } } catch {}
    }
    try {
        if ($null -ne $Item.VersionMajor -and $null -ne $Item.VersionMinor) { return ("{0}.{1}" -f $Item.VersionMajor, $Item.VersionMinor) }
    } catch {}
    try {
        $packed = Convert-PackedMsiVersion $Item.Version
        if (Test-NonEmptyVersion $packed) { return $packed }
    } catch {}
    try {
        if ($Item.QuietDisplayName -match '(\d+(\.\d+){1,5}([\+\-][A-Za-z0-9\.]+)?)') { return $matches[1] }
    } catch {}
    return ""
}

function New-AppEntry {
    param(
        [string]$Name,
        [string]$Version = "",
        [string]$Publisher = "",
        [string]$Source = "Installed Apps",
        [string]$Id = ""
    )
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $cleanName = $Name.Trim()
    [PSCustomObject]@{
        Name       = $cleanName
        Version    = Normalize-VersionText $Version
        Publisher  = if ($Publisher) { $Publisher.ToString().Trim() } else { "" }
        Source     = $Source
        Id         = $Id
        Normalized = Normalize-Name $cleanName
    }
}

$ExplicitKeepNamePatterns = @(
    '^microsoft\s+edge$',
    '^microsoft\s+onedrive$',
    '^microsoft\s+teams$',
    '^microsoft\s+visual\s+studio\s+code$',
    '^visual\s+studio\s+code$',
    '^microsoft\s+visual\s+studio\s+tools\s+for\s+applications',
    '^power\s+automate$',
    '^r\s+for\s+windows',
    '^rstudio$',
    '^nvidia\s+physx',
    '^uxp\s+webview\s+support$'
)

$DefaultWindowsAppNamePatterns = @(
    '^alarms?\s*&?\s*clock$', '^calculator$', '^camera$', '^clipchamp$', '^clock$', '^copilot$', '^cortana$',
    '^dev\s+home$', '^family$', '^feedback\s+hub$', '^get\s+help$', '^get\s+started$', '^mail\s+and\s+calendar$',
    '^maps$', '^media\s+player$', '^microsoft\s+365\s*\(office\)$', '^microsoft\s+bing', '^microsoft\s+clipchamp$',
    '^microsoft\s+copilot$', '^microsoft\s+family$', '^microsoft\s+news$', '^microsoft\s+photos$', '^microsoft\s+solitaire',
    '^microsoft\s+store$', '^microsoft\s+to\s+do$', '^microsoft\s+whiteboard$', '^mixed\s+reality\s+portal$',
    '^movies\s*&\s*tv$', '^msn\s+', '^notepad$', '^office$', '^onenote$', '^outlook\s*\(new\)$', '^paint$',
    '^people$', '^phone\s+link$', '^quick\s+assist$', '^remote\s+desktop$', '^screen\s+sketch$', '^snipping\s+tool$',
    '^sound\s+recorder$', '^sticky\s+notes$', '^terminal$', '^tips$', '^voice\s+recorder$', '^weather$',
    '^windows\s+(security|backup|camera|maps|media|notepad|photos|terminal|web\s+experience)', '^xbox', '^your\s+phone$'
)


$ExcludedMicrosoftPackagePrefixes = @(
    'MicrosoftWindows',
    'Microsoft.Sec',
    'Microsoft.Start',
    'Microsoft.Windows',
    'Microsoft.Win32'
)

$DefaultWindowsPackagePatterns = @(
    '^Clipchamp\.', '^Microsoft\.WindowsAlarms$', '^Microsoft\.WindowsCalculator$', '^Microsoft\.WindowsCamera$',
    '^Microsoft\.WindowsMaps$', '^Microsoft\.WindowsSoundRecorder$', '^Microsoft\.WindowsFeedbackHub$', '^Microsoft\.GetHelp$',
    '^Microsoft\.Getstarted$', '^Microsoft\.People$', '^Microsoft\.Todos$', '^Microsoft\.MicrosoftSolitaireCollection$',
    '^Microsoft\.MicrosoftStickyNotes$', '^Microsoft\.WindowsNotepad$', '^Microsoft\.Paint$', '^Microsoft\.MSPaint$',
    '^Microsoft\.ScreenSketch$', '^Microsoft\.SnippingTool$', '^Microsoft\.Windows\.Photos$', '^Microsoft\.Photos$',
    '^Microsoft\.Bing(News|Weather|Finance|Sports)', '^Microsoft\.Zune(Music|Video)$', '^Microsoft\.Xbox', '^Microsoft\.YourPhone$',
    '^Microsoft\.CorporationII\.QuickAssist$', '^Microsoft\.549981C3F5F10$', '^Microsoft\.Windows\.DevHome$',
    '^Microsoft\.WindowsTerminal$', '^Microsoft\.WindowsStore$', '^Microsoft\.StorePurchaseApp$', '^Microsoft\.DesktopAppInstaller$',
    '^Microsoft\.WebMediaExtensions$', '^Microsoft\.WebpImageExtension$', '^Microsoft\.HEIFImageExtension$',
    '^Microsoft\.HEVCVideoExtension$', '^Microsoft\.VP9VideoExtensions$', '^Microsoft\.VCLibs', '^Microsoft\.UI\.Xaml',
    '^Microsoft\.NET\.Native', '^Microsoft\.Services\.Store'
)

$FrameworkPatterns = @(
    '^microsoft\s+\.net', '^\.net\s+', 'asp\.net', 'windows\s+desktop\s+runtime',
    'microsoft\s+edge\s+webview2\s+runtime', 'webview2\s+runtime',
    '^microsoft\s+visual\s+c\+\+.*redistributable', '^microsoft\s+visual\s+c\+\+.*runtime',
    'visual\s+c\+\+\s+redistributable', '\bvc\+\+\b', '\bvc_redist\b', '^microsoft\s+vclibs',
    '^microsoft\s+ui\s+xaml', '^microsoft\s+windows\s+app\s+runtime', '^windows\s+app\s+runtime',
    '^microsoft\s+update\s+health\s+tools$', '^update\s+for\s+', '^security\s+update\s+for\s+', '^hotfix\s+for\s+',
    '^servicing\s+stack', '^kb\d+', 'language\s+pack$', 'click-to-run\s+extensibility\s+component',
    'office\s+shared\s+.*\s+component', '^microsoft\s+edge\s+update$', 'mozilla\s+maintenance\s+service',
    'google\s+update\s+helper', '^intel\(r\).*driver', '^intel\s+.*driver', '^amd\s+.*driver',
    '^nvidia\s+(graphics|hd\s+audio|frameview).*driver', '^realtek\s+.*driver'
)


function Test-ExcludedMicrosoftPackagePrefix {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $value = $Text.Trim()
    foreach ($prefix in $ExcludedMicrosoftPackagePrefixes) {
        if ($value.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Test-ExplicitKeep {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    $n = $Name.Trim().ToLowerInvariant()
    foreach ($p in $ExplicitKeepNamePatterns) { if ($n -match $p) { return $true } }
    return $false
}

function Test-DefaultWindowsOrFramework {
    param($App)
    if (-not $FilterWindowsDefaultsAndFrameworks) { return $false }
    if ($null -eq $App -or [string]::IsNullOrWhiteSpace($App.Name)) { return $true }
    if (Test-ExplicitKeep $App.Name) { return $false }

    $name = $App.Name.Trim()
    $lower = $name.ToLowerInvariant()
    $id = if ($App.Id) { $App.Id.ToString() } else { "" }
    if ((Test-ExcludedMicrosoftPackagePrefix $name) -or (Test-ExcludedMicrosoftPackagePrefix $id)) { return $true }
    foreach ($p in $DefaultWindowsAppNamePatterns) { if ($lower -match $p) { return $true } }
    foreach ($p in $DefaultWindowsPackagePatterns) { if ($name -match $p -or $id -match $p) { return $true } }
    foreach ($p in $FrameworkPatterns) { if ($lower -match $p) { return $true } }
    return $false
}

function Test-VisibleRegistryInstalledApp {
    param($Item)
    if ($null -eq $Item.DisplayName -or [string]::IsNullOrWhiteSpace($Item.DisplayName.ToString())) { return $false }
    try { if ($Item.SystemComponent -eq 1) { return $false } } catch {}
    try { if ($Item.NoDisplay -eq 1) { return $false } } catch {}
    try { if ($Item.ReleaseType -match 'Update|Hotfix|Security|Service Pack') { return $false } } catch {}
    try { if ($Item.ParentKeyName) { return $false } } catch {}
    return $true
}

function Get-InstalledAppsFromRegistry {
    $apps = @()
    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($path in $paths) {
        try {
            foreach ($item in @(Get-ItemProperty -Path $path -ErrorAction SilentlyContinue)) {
                if (-not (Test-VisibleRegistryInstalledApp $item)) { continue }
                $app = New-AppEntry -Name $item.DisplayName -Version (Get-RegistryDisplayVersion $item) -Publisher $item.Publisher -Source 'Installed Apps - Registry' -Id $item.PSChildName
                if ($app -and -not (Test-DefaultWindowsOrFramework $app)) { $apps += $app }
            }
        } catch {}
    }
    return $apps
}

function Get-AppxFriendlyName {
    param($Package)
    $names = @()
    try {
        $manifest = Get-AppxPackageManifest -Package $Package.PackageFullName -ErrorAction SilentlyContinue
        if ($manifest.Package.Properties.DisplayName) { $names += $manifest.Package.Properties.DisplayName.ToString() }
    } catch {}
    try { $names += $Package.Name } catch {}
    foreach ($n in $names) {
        if ([string]::IsNullOrWhiteSpace($n)) { continue }
        $clean = $n.Trim()
        if ($clean -match '^ms-resource:' -or $clean -match '^@\{') { continue }
        return $clean
    }
    try { return $Package.Name } catch { return "" }
}

function Get-InstalledAppsFromAppx {
    $apps = @()
    try { $packages = @(Get-AppxPackage -PackageTypeFilter Main -ErrorAction SilentlyContinue) }
    catch { try { $packages = @(Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object { -not $_.IsFramework }) } catch { $packages = @() } }

    foreach ($pkg in $packages) {
        try { if ($pkg.IsFramework) { continue } } catch {}
        $name = Get-AppxFriendlyName $pkg
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $publisher = ""; try { $publisher = $pkg.Publisher } catch {}
        $id = ""; try { $id = $pkg.Name } catch {}
        $app = New-AppEntry -Name $name -Version $pkg.Version -Publisher $publisher -Source 'Installed Apps - Store/AppX' -Id $id
        if ($app -and -not (Test-DefaultWindowsOrFramework $app)) { $apps += $app }
    }
    return $apps
}

function Get-AppVersionScore { param($App) if (Test-NonEmptyVersion $App.Version) { return 20 } return 0 }
function Get-AppSourceScore { param($App) if ($App.Source -like '*Registry*') { return 10 } if ($App.Source -like '*Store*' -or $App.Source -like '*AppX*') { return 5 } return 1 }

function Merge-InstalledApps {
    param([array]$Apps)
    return @($Apps |
        Where-Object { $_ -and $_.Normalized -ne "" } |
        Group-Object Normalized |
        ForEach-Object {
            $_.Group | Sort-Object `
                @{Expression={Get-AppVersionScore $_};Descending=$true}, `
                @{Expression={Get-AppSourceScore $_};Descending=$true}, `
                @{Expression={if ($_.Name) { $_.Name.Length } else { 9999 }};Ascending=$true}, `
                Name | Select-Object -First 1
        } |
        Sort-Object Name)
}

function Get-InstalledApps {
    return @(Merge-InstalledApps -Apps @((Get-InstalledAppsFromRegistry) + (Get-InstalledAppsFromAppx)))
}

function Get-RequiredAliases {
    param([string]$Name)
    $req = Normalize-Name $Name
    $aliases = @($req)
    $map = @{
        'adobepremierpro' = @('adobepremierepro')
        'adobepremierpro2025' = @('adobepremierepro2025')
        'adobeacrobat64bit' = @('adobeacrobat')
        'microsoftvisualstudiocode' = @('visualstudiocode')
        'vscode' = @('visualstudiocode')
        'adobedigitaledition' = @('adobedigitaleditions')
        'adobedigitaledition45' = @('adobedigitaleditions45')
        'ibmspssstatics' = @('ibmspssstatistics')
    }
    if ($map.ContainsKey($req)) { $aliases += $map[$req] }
    return @($aliases | Where-Object { $_ } | Select-Object -Unique)
}

function Get-MatchScore {
    param([string]$Key, $App)
    if ([string]::IsNullOrWhiteSpace($Key) -or $null -eq $App -or [string]::IsNullOrWhiteSpace($App.Normalized)) { return 0 }
    $i = $App.Normalized
    if ($i -eq $Key) { return 100 }
    if ($i.StartsWith($Key)) { return 85 }
    if ($Key.StartsWith($i) -and $i.Length -ge 6) { return 80 }
    if ($i.Contains($Key) -and $Key.Length -ge 6) { return 70 }
    if ($Key.Contains($i) -and $i.Length -ge 8) { return 65 }
    return 0
}

function Find-AppMatch {
    param([string]$RequiredName, [array]$InstalledApps)
    $hits = @()
    foreach ($app in $InstalledApps) {
        foreach ($key in (Get-RequiredAliases $RequiredName)) {
            $score = Get-MatchScore -Key $key -App $app
            if ($score -gt 0) {
                $hits += [PSCustomObject]@{ App=$app; Score=$score; VersionScore=(Get-AppVersionScore $app); SourceScore=(Get-AppSourceScore $app); Len=$app.Name.Length }
            }
        }
    }
    if ($hits.Count -eq 0) { return $null }
    return ($hits | Sort-Object @{Expression={$_.Score};Descending=$true}, @{Expression={$_.VersionScore};Descending=$true}, @{Expression={$_.SourceScore};Descending=$true}, @{Expression={$_.Len};Ascending=$true} | Select-Object -First 1).App
}

function Set-FillColor {
    param($RangeOrCell, [int]$Color)
    try {
        $RangeOrCell.Interior.Pattern = 1
        $RangeOrCell.Interior.Color = $Color
    } catch {}
}

function Clear-CellComment {
    param($Cell)
    try { if ($Cell.Comment()) { $Cell.Comment().Delete() | Out-Null } } catch {}
}

function Set-CellComment {
    param($Cell, [string]$Text)
    Clear-CellComment -Cell $Cell
    if (-not [string]::IsNullOrWhiteSpace($Text)) { try { $Cell.AddComment($Text) | Out-Null } catch {} }
}

# Audited asset cells are always explicitly marked as one of three outcomes: green, red, or yellow.
function Set-AuditSame {
    param($Cell)
    try { $Cell.ClearContents() | Out-Null } catch {}
    Clear-CellComment -Cell $Cell
    Set-FillColor -RangeOrCell $Cell -Color $OriginalGreenColor
}

function Set-AuditMissing {
    param($Cell, [string]$Comment)
    try { $Cell.ClearContents() | Out-Null } catch {}
    Set-FillColor -RangeOrCell $Cell -Color $OriginalRedColor
    Set-CellComment -Cell $Cell -Text $Comment
}

function Set-AuditDifferent {
    param($Cell, [string]$Version, [string]$Comment)
    $value = if (Test-NonEmptyVersion $Version) { $Version } else { 'Installed - version unavailable' }
    try { $Cell.Value2 = [string]$value } catch {}
    Set-FillColor -RangeOrCell $Cell -Color $OriginalYellowColor
    Set-CellComment -Cell $Cell -Text $Comment
}

function Apply-OriginalDefaultFormatting {
    param($Sheet, [int]$FirstAppRow, [int]$LastAppRow, [int]$LastCol)
    if ($LastAppRow -lt $FirstAppRow) { return }
    try {
        $blueEnd = [Math]::Min(2, $LastCol)
        if ($blueEnd -ge 1) { Set-FillColor -RangeOrCell ($Sheet.Range($Sheet.Cells.Item($FirstAppRow,1), $Sheet.Cells.Item($LastAppRow,$blueEnd))) -Color $OriginalBlueColor }
        if ($LastCol -ge 3) { Set-FillColor -RangeOrCell ($Sheet.Range($Sheet.Cells.Item($FirstAppRow,3), $Sheet.Cells.Item($LastAppRow,$LastCol))) -Color $OriginalGreenColor }
    } catch {}
}

function Get-AutoAssetTag {
    $bad = @('', 'default string', 'none', 'no asset tag', 'no asset information', 'to be filled by o.e.m.', 'system serial number', 'unknown')
    try {
        $tag = (Get-CimInstance Win32_SystemEnclosure -ErrorAction Stop | Select-Object -First 1 -ExpandProperty SMBIOSAssetTag)
        if ($tag) { $tag = $tag.ToString().Trim(); if ($bad -notcontains $tag.ToLowerInvariant()) { return $tag } }
    } catch {}
    try {
        $serial = (Get-CimInstance Win32_BIOS -ErrorAction Stop | Select-Object -First 1 -ExpandProperty SerialNumber)
        if ($serial) { $serial = $serial.ToString().Trim(); if ($bad -notcontains $serial.ToLowerInvariant()) { return $serial } }
    } catch {}
    return $null
}

function Get-AssetTagChoice {
    $auto = Get-AutoAssetTag
    if ($auto) {
        Write-Host "Detected asset tag / serial: $auto" -ForegroundColor Green
        $accept = Read-Host "Use this value? [Y/n]"
        if ([string]::IsNullOrWhiteSpace($accept) -or $accept.Trim().ToLowerInvariant().StartsWith('y')) { return $auto }
    } else {
        Write-Host "No usable BIOS asset tag was found. Switching to manual entry." -ForegroundColor Yellow
    }

    do {
        $manual = (Read-Host "Enter asset tag for this device").Trim()
        if ([string]::IsNullOrWhiteSpace($manual)) { Write-Host "Asset tag cannot be blank." -ForegroundColor Yellow }
    } until (-not [string]::IsNullOrWhiteSpace($manual))
    return $manual
}

function Get-MainMenuChoice {
    Write-Host "Software Audit Utility - Installed Apps Only" -ForegroundColor Cyan
    Write-Host "  1. Audit existing spreadsheet"
    Write-Host "  2. Export filtered Installed Apps inventory to CSV"
    do { $choice = Read-Host "Choose 1 or 2" } until ($choice -in @('1','2'))
    return $choice
}

function Get-WorkbookPath {
    $dir = Split-Path -Parent $MyInvocation.ScriptName
    $files = @(Get-ChildItem -Path $dir -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '*_backup_*' -and $_.FullName -notmatch '\\Backups\\' -and $_.FullName -notmatch '\\Exports\\' })
    if ($files.Count -eq 1) { return $files[0].FullName }
    if ($files.Count -gt 1) {
        Write-Host "Multiple spreadsheets found:" -ForegroundColor Cyan
        for ($i=0; $i -lt $files.Count; $i++) { Write-Host ("  {0}. {1}" -f ($i+1), $files[$i].Name) }
        do { $raw = Read-Host "Select spreadsheet number"; $n=0; $ok=[int]::TryParse($raw,[ref]$n) -and $n -ge 1 -and $n -le $files.Count } until ($ok)
        return $files[$n-1].FullName
    }
    $manual = Read-Host "Workbook path"
    if (-not (Test-Path $manual)) { throw "Workbook not found: $manual" }
    return $manual
}

function Get-WorksheetSelection {
    param($Workbook)
    $sheets = @(); for ($i=1; $i -le $Workbook.Worksheets.Count; $i++) { $sheets += $Workbook.Worksheets.Item($i) }
    if ($sheets.Count -eq 1) { return @($sheets[0]) }
    Write-Host "Worksheets found:" -ForegroundColor Cyan
    for ($i=0; $i -lt $sheets.Count; $i++) { Write-Host ("  {0}. {1}" -f ($i+1), $sheets[$i].Name) }
    Write-Host "  A. All worksheets"
    do {
        $raw = Read-Host "Select worksheet number or A"
        if ($raw -match '^[Aa]$') { return @($sheets) }
        $n=0; $ok=[int]::TryParse($raw,[ref]$n) -and $n -ge 1 -and $n -le $sheets.Count
    } until ($ok)
    return @($sheets[$n-1])
}

function Test-HeaderRow {
    param($Sheet)
    $a = Normalize-Name ($Sheet.Cells.Item(1,1).Text)
    $b = Normalize-Name ($Sheet.Cells.Item(1,2).Text)
    return (($a -in @('software','requiredsoftware','softwarename')) -and ($b -in @('version','expectedversion')))
}

function Get-UsedLastRow { param($Sheet) try { $u=$Sheet.UsedRange; return ($u.Row + $u.Rows.Count - 1) } catch { return 1 } }
function Get-UsedLastCol { param($Sheet) try { $u=$Sheet.UsedRange; return ($u.Column + $u.Columns.Count - 1) } catch { return 2 } }

function Get-AppLastRow {
    param($Sheet, [int]$FirstAppRow)
    $last = Get-UsedLastRow $Sheet
    for ($r=$FirstAppRow; $r -le $last; $r++) {
        $v = $Sheet.Cells.Item($r,1).Text.Trim()
        if ([string]::IsNullOrWhiteSpace($v) -or $v -match '^notes?:\s*$') { return ($r-1) }
    }
    return $last
}

function Get-SheetLayout {
    param($Sheet)
    $hasHeader = Test-HeaderRow $Sheet
    $first = if ($hasHeader) { 2 } else { 1 }
    $last = Get-AppLastRow -Sheet $Sheet -FirstAppRow $first
    [PSCustomObject]@{ HasHeader=$hasHeader; FirstAppRow=$first; LastAppRow=$last; LastColumn=[Math]::Max(2,(Get-UsedLastCol $Sheet)) }
}

function Ensure-AuditHeader {
    param($Sheet)
    $layout = Get-SheetLayout $Sheet
    if ($layout.HasHeader) { return (Get-SheetLayout $Sheet) }
    Write-Host "Worksheet '$($Sheet.Name)' has no header. Inserting one so the asset column can be labeled." -ForegroundColor Yellow
    $Sheet.Rows.Item(1).Insert() | Out-Null
    $Sheet.Cells.Item(1,1).Value2 = 'Software '
    $Sheet.Cells.Item(1,2).Value2 = 'Version'
    try { $Sheet.Range($Sheet.Cells.Item(1,1),$Sheet.Cells.Item(1,2)).Font.Bold = $true; $Sheet.Range($Sheet.Cells.Item(1,1),$Sheet.Cells.Item(1,2)).Borders.LineStyle = 1 } catch {}
    return (Get-SheetLayout $Sheet)
}

function Copy-FormatRangeSafe {
    param($SourceRange, $DestinationRange, $Excel)
    try { $SourceRange.Copy() | Out-Null; $DestinationRange.PasteSpecial(-4122) | Out-Null; $Excel.CutCopyMode=$false; return $true }
    catch { try { $Excel.CutCopyMode=$false } catch {}; return $false }
}

function Invoke-AuditWorkbook {
    param([array]$InstalledApps)
    $asset = Get-AssetTagChoice
    if ([string]::IsNullOrWhiteSpace($asset)) { throw "Asset tag cannot be blank." }
    $path = Get-WorkbookPath
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $null
    try {
        $workbook = $excel.Workbooks.Open($path)
        $sheets = @(Get-WorksheetSelection -Workbook $workbook)
        $totalRows=0; $totalSame=0; $totalMissing=0; $totalDiff=0
        foreach ($sheet in $sheets) {
            $layout = Ensure-AuditHeader $sheet
            if ($layout.LastAppRow -lt $layout.FirstAppRow) { Write-Host "Skipping '$($sheet.Name)': no app rows." -ForegroundColor Yellow; continue }
            $lastCol = [Math]::Max($layout.LastColumn, 2)
            $assetCol = $null
            for ($c=3; $c -le $lastCol; $c++) { if ($sheet.Cells.Item(1,$c).Text.Trim() -eq $asset) { $assetCol=$c; break } }
            if (-not $assetCol) {
                $assetCol = $lastCol + 1
                if ($lastCol -ge 3) {
                    Copy-FormatRangeSafe -SourceRange ($sheet.Range($sheet.Cells.Item(1,$lastCol),$sheet.Cells.Item($layout.LastAppRow,$lastCol))) -DestinationRange ($sheet.Range($sheet.Cells.Item(1,$assetCol),$sheet.Cells.Item($layout.LastAppRow,$assetCol))) -Excel $excel | Out-Null
                    try { $sheet.Columns.Item($assetCol).ColumnWidth = $sheet.Columns.Item($lastCol).ColumnWidth } catch {}
                } else {
                    try { $sheet.Columns.Item($assetCol).ColumnWidth = [Math]::Min([Math]::Max($sheet.Columns.Item(2).ColumnWidth,12),18) } catch {}
                }
                $sheet.Cells.Item(1,$assetCol).Value2 = [string]$asset
                $sheet.Cells.Item(1,$assetCol).Font.Bold = $true
            }

            $same=0; $missing=0; $diff=0; $rows=0
            for ($r=$layout.FirstAppRow; $r -le $layout.LastAppRow; $r++) {
                $name = $sheet.Cells.Item($r,1).Text.Trim()
                $expected = $sheet.Cells.Item($r,2).Text.Trim()
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $rows++
                $cell = $sheet.Cells.Item($r,$assetCol)
                $match = Find-AppMatch -RequiredName $name -InstalledApps $InstalledApps
                if ($match) {
                    $version = Normalize-VersionText $match.Version
                    if ([string]::IsNullOrWhiteSpace($version)) { $version = 'Installed' }
                    if (Test-VersionMatch -Expected $expected -Installed $version) {
                        $same++; Set-AuditSame -Cell $cell
                    } else {
                        $diff++
                        $expectedLabel = if ([string]::IsNullOrWhiteSpace($expected)) { '(blank)' } else { $expected }
                        Set-AuditDifferent -Cell $cell -Version $version -Comment "Detected as: $($match.Name)`nInstalled version: $version`nExpected version: $expectedLabel`nSource: $($match.Source)`nAudit time: $timestamp"
                    }
                } else {
                    $missing++
                    Set-AuditMissing -Cell $cell -Comment "No matching Installed Apps entry found for '$name'. Audit time: $timestamp"
                }
            }
            Write-Host "Worksheet '$($sheet.Name)': rows=$rows, same=$same, missing=$missing, different=$diff"
            $totalRows += $rows; $totalSame += $same; $totalMissing += $missing; $totalDiff += $diff
        }
        $backupDir = Join-Path ([System.IO.Path]::GetDirectoryName($path)) 'Backups'
        if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
        $backupPath = Join-Path $backupDir ([System.IO.Path]::GetFileNameWithoutExtension($path) + '_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx')
        $workbook.SaveCopyAs($backupPath)
        $workbook.Save()
        Write-Host "Audit complete." -ForegroundColor Green
        Write-Host "Rows checked: $totalRows"
        Write-Host "Same version: $totalSame"
        Write-Host "Missing: $totalMissing"
        Write-Host "Different / version unavailable: $totalDiff"
        Write-Host "Workbook updated: $path"
        Write-Host "Backup saved: $backupPath"
    } finally {
        if ($workbook) { $workbook.Close($true) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    }
}

function Export-InstalledAppsCsv {
    param([array]$InstalledApps)
    $dir = Split-Path -Parent $MyInvocation.ScriptName
    $exportDir = Join-Path $dir 'Exports'
    if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }
    $machine = $env:COMPUTERNAME -replace '[^a-zA-Z0-9_-]','_'
    if ([string]::IsNullOrWhiteSpace($machine)) { $machine = 'Computer' }
    $path = Join-Path $exportDir ('InstalledApps_Filtered_' + $machine + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv')
    $InstalledApps | Select-Object Name, Version, Publisher, Source, Id | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    Write-Host "Filtered Installed Apps CSV exported:" -ForegroundColor Green
    Write-Host $path
}

$choice = Get-MainMenuChoice
Write-Host "Scanning Windows Installed Apps inventory..."
$installedApps = @(Get-InstalledApps)
Write-Host "Detected $($installedApps.Count) filtered Installed Apps entries."
if ($choice -eq '1') { Invoke-AuditWorkbook -InstalledApps $installedApps }
else { Export-InstalledAppsCsv -InstalledApps $installedApps }
