# ==========================================================
#  TypesetToolkit -- verify the toolbar really shows up
#
#  The installer proves the workspace XML is written and valid.
#  What it cannot prove is that CorelDRAW, started the way a user
#  starts it, actually materialises the "排版工具" toolbar from that
#  markup.  This script settles that:
#
#    1. start CorelDRW.exe by hand
#    2. enumerate every top-level window owned by that process
#    3. look for a window whose caption is the toolbar name
#    4. print every caption (so a failure is diagnosable)
#    5. close CorelDRAW again
#
#  ASCII ONLY: the toolbar name is rebuilt from char codes so this
#  file never depends on how PowerShell guesses its encoding.
#
#  usage:
#    powershell -ExecutionPolicy Bypass -File tools\verify_toolbar.ps1
#    powershell -ExecutionPolicy Bypass -File tools\verify_toolbar.ps1 -KeepRunning
# ==========================================================

param(
  [int]$WaitSec = 25,
  [string]$Exe = "C:\Program Files (x86)\CorelDRAW X4\Programs\CorelDRW.exe",
  [switch]$KeepRunning
)

$ErrorActionPreference = 'Continue'

# 排版工具  (avoid any encoding dependency)
$TB = [string][char]0x6392 + [char]0x7248 + [char]0x5DE5 + [char]0x5177

Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class WinEnum {
  delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr p);
  [DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);

  public static List<string> Titles(int wantPid) {
    var res = new List<string>();
    EnumWindows(delegate(IntPtr h, IntPtr p) {
      int pid; GetWindowThreadProcessId(h, out pid);
      if (pid != wantPid) return true;
      var t = new StringBuilder(512);
      GetWindowTextW(h, t, t.Capacity);
      var c = new StringBuilder(256);
      GetClassNameW(h, c, c.Capacity);
      res.Add((IsWindowVisible(h) ? "vis" : "hid") + " | class=" + c.ToString() + " | title=" + t.ToString());
      return true;
    }, IntPtr.Zero);
    return res;
  }
}
"@

$here = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$logPath = Join-Path $here '_toolbar_check.log'
$script:lines = New-Object System.Collections.ArrayList
function L([string]$s) { [void]$script:lines.Add($s); Write-Host $s }
function Flush { try { $script:lines -join "`r`n" | Set-Content -Path $logPath -Encoding Unicode } catch {} }

L "==== verify_toolbar $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') ===="
L "want caption = [$TB]"

if (-not (Test-Path $Exe)) { L "!! exe not found"; Flush; exit 2 }

L "[1] Start-Process $Exe"
try { Start-Process -FilePath $Exe } catch { L "!! start failed: $_"; Flush; exit 3 }

$pid2 = 0
for ($i = 0; $i -lt 40; $i++) {
  Start-Sleep -Milliseconds 500
  $p = Get-Process -Name CorelDRW -ErrorAction SilentlyContinue |
       Where-Object { $_.Id -ne 0 } |
       Sort-Object StartTime -Descending | Select-Object -First 1
  if ($null -ne $p) { $pid2 = $p.Id; break }
}
if ($pid2 -eq 0) { L "!! CorelDRAW never appeared"; Flush; exit 4 }
L "    pid = $pid2"

# let the UI finish building the command bars
$deadline = (Get-Date).AddSeconds($WaitSec)
$titles = @()
$hit = $false
while ((Get-Date) -lt $deadline) {
  Start-Sleep -Seconds 3
  try { $titles = [WinEnum]::Titles($pid2) } catch { $titles = @() }
  $hit = @($titles | Where-Object { $_ -like "*title=$TB*" }).Count -gt 0
  L ("  +{0}s : windows={1} toolbarFound={2}" -f ([int]((Get-Date) - $deadline).TotalSeconds + $WaitSec), $titles.Count, $hit)
  if ($hit) { break }
}

L "[2] window dump (pid $pid2) :"
foreach ($t in $titles) { L "    $t" }

if (-not $KeepRunning) {
  L "[3] closing CorelDRAW ..."
  & taskkill /PID $pid2 2>&1 | ForEach-Object { L "    $_" }
  for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 500
    if (-not (Get-Process -Id $pid2 -ErrorAction SilentlyContinue)) { break }
  }
}

L ""
if ($hit) { L "==== TOOLBAR PASS ====  [$TB] is on screen" } else { L "==== TOOLBAR FAIL ====  no window titled [$TB]" }
Flush
if ($hit) { exit 0 } else { exit 1 }