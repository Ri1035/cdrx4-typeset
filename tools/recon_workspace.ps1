<#
  recon_workspace.ps1 -- recon a reference plugin, step 1 (no CorelDRAW needed)

  READ-ONLY.  Parses CorelDRAW's workspace file DRAWUIConfig.xml and lists
  every toolbar / flyout the plugins installed, every button caption, and
  the macro path each button is bound to.

  WHY THIS IS THE MOST VALUABLE STEP
  ----------------------------------
  A reference plugin's GMS (e.g. the local "排版" one) is usually protected,
  so its source cannot be read.  But CorelDRAW persists every button's macro
  path in PLAIN TEXT inside the workspace:

      <itemData guid="..." dynamicCommand="GlobalMacros.ButtShapesUpModule.中心对齐" .../>

  That single line leaks the reference plugin's project name, module names
  and procedure names -- i.e. its feature list and its code skeleton.  What
  is left is re-implementing each feature from its observed behaviour.

  ASCII-ONLY ON PURPOSE: PowerShell 5 reads a .ps1 as ANSI unless the file
  starts with a UTF-8 BOM, so Chinese literals in here would be mojibake.
  Chinese content flows through from the XML, which IS read as UTF-8.

  USAGE
    powershell -ExecutionPolicy Bypass -File tools\recon_workspace.ps1
    powershell -ExecutionPolicy Bypass -File tools\recon_workspace.ps1 -Config "D:\...\DRAWUIConfig.xml"
    powershell -ExecutionPolicy Bypass -File tools\recon_workspace.ps1 -Out recon_ws.txt
#>
param(
    [string]$Config = "",
    [string]$Out = ""
)

$ErrorActionPreference = "Stop"
$u8 = New-Object System.Text.UTF8Encoding($false)

function Get-Configs {
    param([string]$Explicit)
    if ($Explicit -ne "") {
        if (-not (Test-Path $Explicit)) { throw "no such file: $Explicit" }
        return @(Get-Item $Explicit)
    }
    $root = Join-Path $env:APPDATA "Corel"
    if (-not (Test-Path $root)) { throw "no Corel folder under %APPDATA% - run CorelDRAW once first" }
    $found = @(Get-ChildItem $root -Recurse -Filter "DRAWUIConfig.xml" -File -ErrorAction SilentlyContinue)
    if ($found.Count -eq 0) { throw "no DRAWUIConfig.xml found under $root" }
    return $found
}

# one XML tag -> hashtable of its attributes
function Get-Attrs {
    param([string]$Tag)
    $h = @{}
    foreach ($m in [regex]::Matches($Tag, '([A-Za-z][A-Za-z0-9]*)\s*=\s*"([^"]*)"')) {
        $h[$m.Groups[1].Value] = $m.Groups[2].Value
    }
    return $h
}

function Report-Config {
    param([System.IO.FileInfo]$File, [System.Text.StringBuilder]$Sb)

    $txt = $u8.GetString([System.IO.File]::ReadAllBytes($File.FullName))
    [void]$Sb.AppendLine("===================================================================")
    [void]$Sb.AppendLine("config : " + $File.FullName)
    [void]$Sb.AppendLine("size   : " + $txt.Length)
    [void]$Sb.AppendLine("")

    # ---- 1. every button / menu item, keyed by guid -------------------
    $items = @{}
    foreach ($m in [regex]::Matches($txt, '<itemData\s[^>]*?/?>')) {
        $a = Get-Attrs $m.Value
        if ($a.ContainsKey("guid")) { $items[$a["guid"]] = $a }
    }

    # ---- 2. every command bar (toolbar / flyout / menu) ---------------
    $bars = @{}
    foreach ($m in [regex]::Matches($txt, '<commandBarData\s([^>]*?)>(.*?)</commandBarData>', 'Singleline')) {
        $a = Get-Attrs $m.Groups[1].Value
        if (-not $a.ContainsKey("guid")) { continue }
        $refs = @()
        foreach ($r in [regex]::Matches($m.Groups[2].Value, '<item\s+guidRef="([^"]+)"')) {
            $refs += $r.Groups[1].Value
        }
        $bars[$a["guid"]] = [pscustomobject]@{ Attrs = $a; Refs = $refs }
    }

    # ---- 3. the user-created bars: these ARE the plugin toolbars ------
    [void]$Sb.AppendLine("--- user-created toolbars / flyouts (what the plugins installed) ---")
    $shown = 0
    foreach ($key in $items.Keys) {
        $it = $items[$key]
        if (-not $it.ContainsKey("flyoutBarRef")) { continue }
        $barGuid = $it["flyoutBarRef"]
        $bar = $bars[$barGuid]
        if ($null -eq $bar) { continue }

        # the visible name sits on the parent item; the flyout bar itself
        # is usually still called "新工具栏 N"
        $label = $it["userCaption"]
        if (-not $label) { $label = $bar.Attrs["userCaption"] }
        if (-not $label) { $label = "(unnamed)" }

        [void]$Sb.AppendLine("")
        $head = "[" + $label + "]   flyout guid=" + $barGuid + "   buttons=" + $bar.Refs.Count
        [void]$Sb.AppendLine($head)
        $n = 0
        foreach ($rf in $bar.Refs) {
            $n++
            $child = $items[$rf]
            if ($null -eq $child) {
                [void]$Sb.AppendLine(("  {0,2}. <ref points at a missing itemData: {1}>" -f $n, $rf))
                continue
            }
            $cap = $child["userCaption"]
            if (-not $cap) { $cap = "(no caption)" }
            $cmd = $child["dynamicCommand"]
            if (-not $cmd) { $cmd = "<not a macro: built-in command or submenu>" }
            [void]$Sb.AppendLine(("  {0,2}. {1,-16} -> {2}" -f $n, $cap, $cmd))
        }
        $shown++
    }
    if ($shown -eq 0) { [void]$Sb.AppendLine("(no userCreated toolbar - the plugin may have shipped macros without buttons)") }

    # ---- 4. all macro commands, grouped by project -------------------
    # dynamicCommand = "<Project>.<Module>.<Proc>" for plugin commands.
    # Built-in CorelDRAW commands carry no dot and are filtered out.
    [void]$Sb.AppendLine("")
    [void]$Sb.AppendLine("--- every macro command in the workspace, grouped by project ---")
    [void]$Sb.AppendLine("--- (= the reference plugin's code skeleton) -------------------")
    $byProj = @{}
    foreach ($it in $items.Values) {
        $cmd = $it["dynamicCommand"]
        if (-not $cmd) { continue }
        if ($cmd -notmatch '\.') { continue }
        $proj = $cmd.Split('.')[0]
        if (-not $byProj.ContainsKey($proj)) { $byProj[$proj] = @{} }
        $cap = $it["userCaption"]
        if (-not $cap) { $cap = "" }
        $byProj[$proj][$cmd] = $cap
    }
    if ($byProj.Count -eq 0) {
        [void]$Sb.AppendLine("(no dotted macro command found)")
    }
    foreach ($proj in ($byProj.Keys | Sort-Object)) {
        [void]$Sb.AppendLine("")
        [void]$Sb.AppendLine($proj + "   (" + $byProj[$proj].Count + " commands)")
        foreach ($cmd in ($byProj[$proj].Keys | Sort-Object)) {
            $cap = $byProj[$proj][$cmd]
            $tail = ""
            if ($cap) { $tail = "   caption=" + $cap }
            [void]$Sb.AppendLine("    " + $cmd + $tail)
        }
    }
    [void]$Sb.AppendLine("")
}

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("CorelDRAW workspace recon  " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))

foreach ($c in (Get-Configs $Config)) { Report-Config -File $c -Sb $sb }

$text = $sb.ToString()
Write-Host $text

if ($Out -ne "") {
    $full = [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Out))
    [System.IO.File]::WriteAllText($full, $text, $u8)
    Write-Host "written -> $full"
}