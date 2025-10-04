<# 
    Fix-WindowPositions.ps1  (PowerShell 5.1 compatible)
    ----------------------------------------------------
    1) (Optional) Create a system restore point (if admin & enabled)
    2) Export (backup) and clear window-position related registry caches (HKCU)
    3) Clear jump-list caches (can hold window metadata)
    4) Force Windows to recalc display topology (/internal then /extend)
    5) Restart Explorer cleanly
    6) Enumerate all top-level windows; restore if minimized; move/resize anything
       outside/partially outside the current monitors.
#>

[CmdletBinding()]
param(
    [switch]$NoRestorePoint,
    [switch]$DryRun
)

function Write-Info($msg){ Write-Host "[*] $msg" -ForegroundColor Cyan }
function Write-OK($msg){ Write-Host "[+] $msg" -ForegroundColor Green }
function Write-Warn($msg){ Write-Host "[!] $msg" -ForegroundColor Yellow }
function Write-Err($msg){ Write-Host "[-] $msg" -ForegroundColor Red }

# --- 0) Environment checks
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$role = "Standard user"
if ($IsAdmin) { $role = "Administrator" }
Write-Info ("Running as {0}" -f $role)
if ($DryRun) { Write-Warn "DRY RUN: No changes will be applied." }

# --- 1) Optional restore point
if ($IsAdmin -and -not $NoRestorePoint) {
    try {
        Write-Info "Attempting to create a System Restore Point..."
        Checkpoint-Computer -Description "Fix-WindowPositions" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        Write-OK "Restore point created."
    } catch {
        Write-Warn "Could not create a restore point (disabled or time-limited). Continuing..."
    }
}

# --- 2) Backup & clear registry keys
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir = Join-Path $env:TEMP ("WindowPosBackup_$timestamp")
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null

$keysToClear = @(
  'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Streams',
  'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3',
  'HKCU\Control Panel\Desktop\WindowMetrics'
)

Write-Info "Backing up and clearing window position caches..."
foreach ($key in $keysToClear) {
    $safeName = ($key -replace '[\\/:*?"<>| ]','_')
    $outFile = Join-Path $backupDir "$safeName.reg"
    try {
        cmd.exe /c "reg export `"$key`" `"$outFile`" /y" | Out-Null
        if (Test-Path $outFile) { Write-OK "Backed up: $key -> $outFile" }

        if (-not $DryRun) {
            cmd.exe /c "reg delete `"$key`" /f" | Out-Null
            Write-OK "Cleared: $key"
        } else {
            Write-Info "Would clear: $key"
        }
    } catch {
        Write-Warn "Could not process: $key ($($_.Exception.Message))"
    }
}

# --- 3) Clear jump-list cache (sometimes stores placement metadata)
$autoDest = Join-Path $env:APPDATA "Microsoft\Windows\Recent\AutomaticDestinations"
try {
    if (Test-Path $autoDest) {
        if (-not $DryRun) {
            Get-ChildItem $autoDest -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            Write-OK "Cleared AutomaticDestinations jump-list cache."
        } else {
            Write-Info "Would clear jump-list cache at: $autoDest"
        }
    }
} catch {
    Write-Warn "Could not clear jump-list cache: $($_.Exception.Message)"
}

# --- 4) Force Windows to rebuild display map (flip projection)
function Invoke-DisplayRemap {
    param([switch]$SkipFlip)
    if ($SkipFlip) { 
        Write-Info "Skipping display flip." 
        return 
    }

    $displaySwitch = "$env:SystemRoot\System32\DisplaySwitch.exe"
    if (Test-Path $displaySwitch) {
        Write-Info "Forcing display topology rebuild (DisplaySwitch /internal then /extend)..."
        if (-not $DryRun) {
            Start-Process -FilePath $displaySwitch -ArgumentList "/internal" -WindowStyle Hidden -Wait
            Start-Sleep -Milliseconds 800
            Start-Process -FilePath $displaySwitch -ArgumentList "/extend" -WindowStyle Hidden -Wait
            Start-Sleep -Milliseconds 800
            Write-OK "Display topology refreshed."
        } else {
            Write-Info "Would run: DisplaySwitch /internal -> /extend"
        }
    } else {
        Write-Warn "DisplaySwitch.exe not found; skipping projection flip."
    }
}
Invoke-DisplayRemap

# --- 5) Restart Explorer (clean shell refresh)
function Restart-Explorer {
    Write-Info "Restarting Windows Explorer..."
    if (-not $DryRun) {
        Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 800
        Start-Process explorer.exe
        Start-Sleep -Seconds 2
        Write-OK "Explorer restarted."
    } else {
        Write-Info "Would restart explorer.exe"
    }
}
Restart-Explorer

# --- 6) Enumerate and re-center out-of-bounds windows
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$signature = @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class Win32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern IntPtr GetShellWindow();
    [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    public const int SW_RESTORE = 9;
    public const uint GW_OWNER = 4;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left; public int Top; public int Right; public int Bottom;
    }
}
"@

Add-Type -TypeDefinition $signature -Language CSharp -IgnoreWarnings

function Get-WindowTitle([IntPtr]$hWnd){
    $len = [Win32]::GetWindowTextLength($hWnd)
    if ($len -le 0) { return "" }
    $sb = New-Object System.Text.StringBuilder ($len + 2)
    [void][Win32]::GetWindowText($hWnd, $sb, $sb.Capacity)
    return $sb.ToString()
}

function Get-Rect([IntPtr]$hWnd){
    $r = New-Object Win32+RECT
    if ([Win32]::GetWindowRect($hWnd, [ref]$r)) { return $r } else { return $null }
}

function Rect-ToString($r){ return "L=$($r.Left), T=$($r.Top), R=$($r.Right), B=$($r.Bottom), W=$($r.Right-$r.Left), H=$($r.Bottom-$r.Top)" }

# Virtual screen bounds (union of all monitors)
$Screens = [System.Windows.Forms.Screen]::AllScreens

# Guard in case enumeration fails
if (-not $Screens -or $Screens.Count -eq 0) {
    Write-Warn "No screens detected via .NET. Continuing with window recenter logic on primary screen coordinates (0,0)."
}

# Compute union safely without nested Measure-Object props
$lefts   = @()
$tops    = @()
$rights  = @()
$bottoms = @()
foreach ($s in $Screens) {
    $lefts   += $s.Bounds.Left
    $tops    += $s.Bounds.Top
    $rights  += $s.Bounds.Right
    $bottoms += $s.Bounds.Bottom
}
if ($lefts.Count -eq 0) { $lefts = @(0) }
if ($tops.Count -eq 0) { $tops = @(0) }
if ($rights.Count -eq 0) { $rights = @(1920) }   # sensible default
if ($bottoms.Count -eq 0) { $bottoms = @(1080) } # sensible default

$VirtualLeft   = ($lefts   | Measure-Object -Minimum).Minimum
$VirtualTop    = ($tops    | Measure-Object -Minimum).Minimum
$VirtualRight  = ($rights  | Measure-Object -Maximum).Maximum
$VirtualBottom = ($bottoms | Measure-Object -Maximum).Maximum
$VirtualWidth  = $VirtualRight - $VirtualLeft
$VirtualHeight = $VirtualBottom - $VirtualTop

$Primary = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

Write-Info "Virtual desktop: L=$VirtualLeft T=$VirtualTop W=$VirtualWidth H=$VirtualHeight"

# Helper: does rect intersect any screen?
function Is-InAnyScreen($rect){
    foreach ($s in $Screens){
        $sb = $s.Bounds
        $intersectX = [Math]::Max($rect.Left, $sb.Left) -lt [Math]::Min($rect.Right, $sb.Right)
        $intersectY = [Math]::Max($rect.Top,  $sb.Top)  -lt [Math]::Min($rect.Bottom, $sb.Bottom)
        if ($intersectX -and $intersectY) { return $true }
    }
    return $false
}

# Clamp a rect fully inside a given bounds (keeping size if possible)
function Clamp-IntoBounds([Win32+RECT]$r, [System.Drawing.Rectangle]$b){
    $w = [Math]::Max(100, $r.Right - $r.Left)
    $h = [Math]::Max(80,  $r.Bottom - $r.Top)

    if ($w -gt $b.Width)  { $w = $b.Width  - 40 }
    if ($h -gt $b.Height) { $h = $b.Height - 60 }

    $x = $r.Left; $y = $r.Top
    if ($x -lt $b.Left)                 { $x = $b.Left + 10 }
    if ( ($x + $w) -gt $b.Right )       { $x = $b.Right - $w - 10 }
    if ($y -lt $b.Top)                  { $y = $b.Top + 10 }
    if ( ($y + $h) -gt $b.Bottom )      { $y = $b.Bottom - $h - 10 }

    return @{ X=$x; Y=$y; W=$w; H=$h }
}

# Choose the "best" target screen (primary by default, or nearest intersecting)
function Get-TargetScreen([Win32+RECT]$r){
    $best = $null
    $bestArea = -1
    foreach ($s in $Screens){
        $sb = $s.Bounds
        $ix = [Math]::Max(0, [Math]::Min($r.Right,$sb.Right) - [Math]::Max($r.Left,$sb.Left))
        $iy = [Math]::Max(0, [Math]::Min($r.Bottom,$sb.Bottom) - [Math]::Max($r.Top,$sb.Top))
        $area = $ix * $iy
        if ($area -gt $bestArea) { $best = $s; $bestArea = $area }
    }
    if ($best -eq $null) { return [System.Windows.Forms.Screen]::PrimaryScreen }
    return $best
}

$Shell = [Win32]::GetShellWindow()

$MovedCount = 0
$TotalCount = 0

$delegate = [Win32+EnumWindowsProc]{
    param([IntPtr]$hWnd, [IntPtr]$lParam)

    if ($hWnd -eq $Shell) { return $true }
    if (-not [Win32]::IsWindowVisible($hWnd)) { return $true }
    $owner = [Win32]::GetWindow($hWnd, [Win32]::GW_OWNER)
    if ($owner -ne [IntPtr]::Zero) { return $true }

    $title = Get-WindowTitle $hWnd
    if ([string]::IsNullOrWhiteSpace($title)) { return $true }

    $r = Get-Rect $hWnd
    if ($null -eq $r) { return $true }

    $TotalCount++

    $rectStr = Rect-ToString $r

    if (-not $DryRun) { [Win32]::ShowWindow($hWnd, [Win32]::SW_RESTORE) | Out-Null }

    $targetScreen = Get-TargetScreen $r
    $tb = $targetScreen.Bounds

    $needsMove = $false

    if (-not (Is-InAnyScreen $r)) {
        $needsMove = $true
        Write-Info "Out-of-bounds: '$title' ($rectStr)"
    } else {
        $margin = 5
        if ( ($r.Left  -lt ($tb.Left  - $margin)) -or
             ($r.Top   -lt ($tb.Top   - $margin)) -or
             ($r.Right -gt ($tb.Right + $margin)) -or
             ($r.Bottom-gt ($tb.Bottom+ $margin)) ) {
            $needsMove = $true
            Write-Info "Partially outside: '$title' ($rectStr)"
        }
    }

    if ($needsMove) {
        $pos = Clamp-IntoBounds $r $tb
        if (-not $DryRun) {
            [Win32]::MoveWindow($hWnd, [int]$pos.X, [int]$pos.Y, [int]$pos.W, [int]$pos.H, $true) | Out-Null
            $MovedCount++
            Write-OK "Moved -> X=$($pos.X) Y=$($pos.Y) W=$($pos.W) H=$($pos.H)"
        } else {
            Write-Info "Would move -> X=$($pos.X) Y=$($pos.Y) W=$($pos.W) H=$($pos.H)"
        }
    }

    return $true
}

Write-Info "Scanning & fixing top-level windows..."
[Win32]::EnumWindows($delegate, [IntPtr]::Zero) | Out-Null
Write-OK "Processed $TotalCount windows; moved $MovedCount."
Write-OK "Done. If this happens again after an RDP multi-monitor session, just run this script once more."
