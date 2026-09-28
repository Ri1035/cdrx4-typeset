Option Explicit

'==========================================================
'  recon_vbe.vbs -- recon a reference plugin, step 2
'
'  Answers the one question that decides the whole porting
'  strategy: CAN THE REFERENCE PLUGIN'S SOURCE BE READ?
'
'  It enumerates every VBA project loaded in CorelDRAW and
'  reports, per project:
'     - Name / FileName
'     - Protection   (0 = readable, 1 = password protected)
'     - every component: name, type, line count
'  and, for any UNPROTECTED project, dumps each component's
'  source to _recon_src\<project>\<component>.txt (UTF-16).
'
'  READ-ONLY WITH RESPECT TO THE USER'S WORK:
'    - attaches to a running CorelDRAW if there is one, and
'      then leaves it alone (no Quit);
'    - only launches its own instance when none is running,
'      and quits that one at the end;
'    - never saves, closes or edits a document.
'
'  ASCII ONLY: Windows Script Host parses .vbs as ANSI, so a
'  UTF-8 file containing Chinese would be mis-parsed.  Chinese
'  project / module names come back from COM as real strings
'  and are written to a UTF-16 log, which is fine.
'
'  Result -> _recon_vbe.log (UTF-16) and stdout.
'==========================================================

Dim fso, sh, here, logPath, gLog
Dim app, vbe, ownApp, i, k, waited, ts
Dim pr, comp, prot, lines, nm, fn, dumpRoot, outPath, src
Dim nProj, nComp, nRead, nLocked
Dim cc, cnm, cty, cln

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")
here = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
logPath = here & "\_recon_vbe.log"
dumpRoot = here & "\_recon_src"
gLog = ""
ownApp = False

Sub W(s)
  gLog = gLog & s & vbCrLf
  WScript.Echo s
End Sub

Sub Flush()
  Dim ts
  On Error Resume Next
  Set ts = fso.CreateTextFile(logPath, True, True)
  If Not ts Is Nothing Then
    ts.Write gLog
    ts.Close
  End If
  On Error GoTo 0
End Sub

Sub Die(msg)
  W "!! " & msg
  Flush
  WScript.Quit 1
End Sub

W "==== CorelDRAW VBA recon " & Now & " ===="

' --- 0. attach to a running instance, else start one -----------
Set app = Nothing
On Error Resume Next
Set app = GetObject(, "CorelDRAW.Application.14")
On Error GoTo 0

If Not app Is Nothing Then
  W "[0] attached to the running CorelDRAW (will not quit it)"
Else
  On Error Resume Next
  Err.Clear
  Set app = CreateObject("CorelDRAW.Application.14")
  If Err.Number <> 0 Or app Is Nothing Then
    Die "could not start CorelDRAW - is X4 installed with VBA?"
  End If
  Err.Clear
  app.Visible = True
  app.InitializeVBA
  Err.Clear
  On Error GoTo 0
  ownApp = True
  W "[0] started a private CorelDRAW (will quit it at the end)"
End If

' --- 1. the VBE -----------------------------------------------
Set vbe = Nothing
For waited = 1 To 40
  Set vbe = Nothing
  On Error Resume Next
  Err.Clear
  Set vbe = app.VBE
  Err.Clear
  On Error GoTo 0
  If Not vbe Is Nothing Then Exit For
  WScript.Sleep 500
Next
If vbe Is Nothing Then
  If ownApp Then QuitApp
  Die "no VBE - the VBA component is missing from this CorelDRAW"
End If

' --- 2. every project -----------------------------------------
nProj = 0
nComp = 0
nRead = 0
nLocked = 0

On Error Resume Next
nProj = vbe.VBProjects.Count
Err.Clear
On Error GoTo 0

W ""
W "[1] VBProjects.Count = " & nProj
W ""

For i = 1 To nProj
  Set pr = Nothing
  On Error Resume Next
  Err.Clear
  Set pr = vbe.VBProjects.Item(i)
  Err.Clear
  On Error GoTo 0
  If pr Is Nothing Then
    W "proj " & i & "  <unavailable>"
  Else
    nm = ""
    fn = ""
    prot = -1
    On Error Resume Next
    Err.Clear
    nm = pr.Name
    fn = pr.FileName
    prot = pr.Protection
    Err.Clear
    On Error GoTo 0

    W "-----------------------------------------------------------------"
    W "proj " & i & "  name=[" & nm & "]"
    W "        file=[" & fn & "]"
    W "        protection=" & prot & ProtectionNote(prot)

    If prot = 0 Then nRead = nRead + 1
    If prot = 1 Then nLocked = nLocked + 1

    ' --- components ---
    cc = 0
    On Error Resume Next
    Err.Clear
    cc = pr.VBComponents.Count
    Err.Clear
    On Error GoTo 0
    W "        components=" & cc

    For k = 1 To cc
      Set comp = Nothing
      On Error Resume Next
      Err.Clear
      Set comp = pr.VBComponents.Item(k)
      Err.Clear
      On Error GoTo 0
      If Not comp Is Nothing Then
        cnm = ""
        cty = -1
        cln = -1
        On Error Resume Next
        Err.Clear
        cnm = comp.Name
        cty = comp.Type
        cln = comp.CodeModule.CountOfLines
        Err.Clear
        On Error GoTo 0
        W "          " & k & ". [" & cnm & "] type=" & cty & TypeNote(cty) & " lines=" & cln
        nComp = nComp + 1

        ' dump the source when the project is not protected
        If prot = 0 And cln > 0 Then
          src = ""
          On Error Resume Next
          Err.Clear
          src = comp.CodeModule.Lines(1, cln)
          Err.Clear
          On Error GoTo 0
          If Len(src) > 0 Then
            outPath = DumpPath(dumpRoot, nm, cnm)
            If Len(outPath) > 0 Then
              On Error Resume Next
              Set ts = fso.CreateTextFile(outPath, True, True)
              If Not ts Is Nothing Then
                ts.Write src
                ts.Close
              End If
              Err.Clear
              On Error GoTo 0
            End If
          End If
        End If
      End If
      Err.Clear
    Next
    W ""
  End If
Next

' --- 3. verdict -----------------------------------------------
W "================================================================="
W "SUMMARY"
W "  projects        = " & nProj
W "  components      = " & nComp
W "  readable (prot0)= " & nRead
W "  locked   (prot1)= " & nLocked
If nRead > 0 Then
  W "  source dumped to " & dumpRoot
End If
W ""
If nLocked > 0 And nRead = 0 Then
  W "  VERDICT: every project is password protected.  The source cannot"
  W "           be read, so the port has to be a RE-IMPLEMENTATION from"
  W "           observed behaviour.  Go to the workspace recon + the"
  W "           screenshot / probe loop in the playbook."
ElseIf nRead > 0 Then
  W "  VERDICT: at least one project is readable - read the dumped"
  W "           source first, it beats any amount of guessing."
End If
W "================================================================="

If ownApp Then QuitApp
Flush
WScript.Quit 0


' --- helpers ---------------------------------------------------

Function ProtectionNote(p)
  ProtectionNote = ""
  If p = 0 Then
    ProtectionNote = "  <-- READABLE, source will be dumped"
  ElseIf p = 1 Then
    ProtectionNote = "  <-- password protected"
  Else
    ProtectionNote = "  <-- could not read the flag"
  End If
End Function

Function TypeNote(t)
  TypeNote = ""
  Select Case t
    Case 1
      TypeNote = " (standard module)"
    Case 2
      TypeNote = " (class module)"
    Case 3
      TypeNote = " (userform)"
    Case 100
      TypeNote = " (document / ThisDocument)"
  End Select
End Function

' <root>\<project>\<component>.txt, with the characters a file name
' cannot hold replaced by "_"
Function DumpPath(root, proj, comp)
  Dim d
  DumpPath = ""
  d = root & "\" & Safe(proj)
  On Error Resume Next
  If Not fso.FolderExists(d) Then
    If Not fso.FolderExists(root) Then fso.CreateFolder root
    fso.CreateFolder d
  End If
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0
  DumpPath = d & "\" & Safe(comp) & ".txt"
End Function

Function Safe(s)
  Dim bad, i, ch, r
  bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
  r = s
  For i = 0 To UBound(bad)
    r = Replace(r, bad(i), "_")
  Next
  Safe = r
End Function

Sub QuitApp()
  Dim w
  On Error Resume Next
  app.Quit
  On Error GoTo 0
  For w = 1 To 60
    WScript.Sleep 500
    Set app = Nothing
    On Error Resume Next
    Set app = GetObject(, "CorelDRAW.Application.14")
    On Error GoTo 0
    If app Is Nothing Then Exit For
  Next
End Sub