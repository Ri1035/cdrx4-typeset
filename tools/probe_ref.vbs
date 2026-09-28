Option Explicit

'==========================================================
'  probe_ref.vbs -- behaviour recon for the [PaiBan] plugin
'
'  Runs ONE reference macro (from the encrypted
'  Draw\GlobalMacros.gms) against a scratch document and logs
'  the scene before / after, so the behaviour can be
'  reverse-engineered.
'
'  usage: cscript //nologo tools\probe_ref.vbs <1..6> [scene]
'
'    1 = MainModule.拼版              (imposition)
'    2 = CropMark.裁切线              (crop marks)
'    3 = ThisMacroStorage.书脊计算     (spine thickness)
'    4 = ButtShapesUpModule.中心对齐   (centre align)
'    5 = ButtShapesUpModule.按间距分布  (distribute by spacing)
'    6 = DelSegment.删除线段           (delete segments)
'
'  scene:
'    r = three rectangles, all selected          (default)
'    e = rectangles + ellipse, all selected
'    c = one curve, its first two nodes selected
'    n = three rectangles, nothing selected
'
'  SAFETY:
'    - builds its own scratch document and makes it active;
'    - closes ONLY that scratch document, never saving;
'    - NEVER calls app.Quit (the user's other windows are
'      left alone);
'    - Chinese macro names are built with ChrW() and every
'      comment here is ASCII, so this file parses the same
'      whatever code page WSH happens to read it in.
'
'  Result -> _probe_ref<N>.log  (and echoed to stdout)
'==========================================================

Dim fso, here, root, logPath, logFile
Dim idx, scene, proj, proc, macroFull
Dim app, doc, pg, lay, i, s, shp
Dim nr

Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)

If WScript.Arguments.Count < 1 Then
  WScript.Echo "usage: cscript //nologo tools\probe_ref.vbs <1..6> [r|e|c|n]"
  WScript.Quit 2
End If

idx = CInt(WScript.Arguments(0))
scene = "r"
If WScript.Arguments.Count >= 2 Then scene = LCase(WScript.Arguments(1))

proj = "GlobalMacros"
proc = ProcFor(idx)
If proc = "" Then
  WScript.Echo "bad index " & idx
  WScript.Quit 2
End If
macroFull = proj & "." & proc

logPath = root & "\_probe_ref" & idx & ".log"
Set logFile = fso.CreateTextFile(logPath, True, True)

Sub W(s)
  logFile.WriteLine s
  WScript.Echo s
End Sub

' --- Chinese procedure names, built from code points ----------
Function ProcFor(n)
  Dim cn
  ProcFor = ""
  Select Case n
    Case 1
      cn = ChrW(&H62FC) & ChrW(&H7248)                                 ' pai ban
      ProcFor = "MainModule." & cn
    Case 2
      cn = ChrW(&H88C1) & ChrW(&H5207) & ChrW(&H7EBF)                   ' cai qie xian
      ProcFor = "CropMark." & cn
    Case 3
      cn = ChrW(&H4E66) & ChrW(&H810A) & ChrW(&H8BA1) & ChrW(&H7B97)     ' shu ji ji suan
      ProcFor = "ThisMacroStorage." & cn
    Case 4
      cn = ChrW(&H4E2D) & ChrW(&H5FC3) & ChrW(&H5BF9) & ChrW(&H9F50)     ' zhong xin dui qi
      ProcFor = "ButtShapesUpModule." & cn
    Case 5
      cn = ChrW(&H6309) & ChrW(&H95F4) & ChrW(&H8DDD) & ChrW(&H5206) & ChrW(&H5E03)
      ProcFor = "ButtShapesUpModule." & cn
    Case 6
      cn = ChrW(&H5220) & ChrW(&H9664) & ChrW(&H7EBF) & ChrW(&H6BB5)     ' shan chu xian duan
      ProcFor = "DelSegment." & cn
  End Select
End Function

Sub DumpShapes(prefix, coll)
  Dim n, k, o
  On Error Resume Next
  Err.Clear
  n = 0
  n = coll.Count
  W prefix & " count = " & n & "   (err=" & Err.Number & ")"
  Err.Clear
  On Error GoTo 0
  For k = 1 To n
    On Error Resume Next
    Err.Clear
    Set o = Nothing
    Set o = coll.Item(k)
    If Err.Number = 0 And Not o Is Nothing Then
      W prefix & " [" & k & "] type=" & o.Type & _
        "  pos=" & Round(o.PositionX, 2) & "," & Round(o.PositionY, 2) & _
        "  size=" & Round(o.SizeWidth, 2) & "x" & Round(o.SizeHeight, 2) & _
        "  L/T/R/B=" & Round(o.LeftX, 2) & "/" & Round(o.TopY, 2) & "/" & _
        Round(o.RightX, 2) & "/" & Round(o.BottomY, 2)
      If o.Type = 3 Then
        W prefix & "      nodes=" & o.Curve.Nodes.Count & _
          "  subpaths=" & o.Curve.SubPaths.Count
      End If
    Else
      W prefix & " [" & k & "] read failed [" & Err.Number & "] " & Err.Description
    End If
    Err.Clear
    On Error GoTo 0
  Next
End Sub

Sub DumpScene(tag)
  Dim layr, k, npages

  W "  --- scene " & tag & " ---"

  On Error Resume Next
  Err.Clear
  npages = doc.Pages.Count
  W "      pages = " & npages & "   (err=" & Err.Number & ")"
  Err.Clear
  On Error GoTo 0

  DumpShapes "      page", pg.Shapes

  ' guides live on the master page in X4
  On Error Resume Next
  Err.Clear
  DumpShapes "      master", doc.MasterPage.Shapes
  Err.Clear
  On Error GoTo 0

  ' every layer of the active page, in case the macro wrote elsewhere
  On Error Resume Next
  Err.Clear
  For k = 1 To pg.Layers.Count
    Set layr = Nothing
    Set layr = pg.Layers.Item(k)
    If Err.Number = 0 And Not layr Is Nothing Then
      W "      layer [" & layr.Name & "] shapes=" & layr.Shapes.Count
    End If
    Err.Clear
  Next
  Err.Clear
  On Error GoTo 0

  W "  --- end scene " & tag & " ---"
End Sub

W "==== probe_ref " & idx & "  macro=" & macroFull & "  scene=" & scene & "  " & Now & " ===="

' --- app ------------------------------------------------------
Set app = Nothing
On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
If Err.Number <> 0 Or app Is Nothing Then
  W "!! cannot reach CorelDRAW [" & Err.Number & "] " & Err.Description
  logFile.Close
  WScript.Quit 1
End If
Err.Clear
On Error GoTo 0

On Error Resume Next
Err.Clear
app.InitializeVBA
W "[1] app version = " & app.Version & "   InitializeVBA err=" & Err.Number
W "[1] Documents.Count (before) = " & app.Documents.Count
Err.Clear
On Error GoTo 0

' --- scratch document ----------------------------------------
On Error Resume Next
Err.Clear
Set doc = Nothing
Set doc = app.CreateDocument
W "[2] CreateDocument err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0
If doc Is Nothing Then
  W "!! no scratch document, abort"
  logFile.Close
  WScript.Quit 1
End If

On Error Resume Next
Err.Clear
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer
W "[3] page size = " & Round(pg.SizeWidth, 2) & " x " & Round(pg.SizeHeight, 2) & _
  "   unit = " & doc.Unit & "   active layer = " & lay.Name
Err.Clear
On Error GoTo 0

' --- scene ----------------------------------------------------
Select Case scene
  Case "c"
    ' one rectangle turned into a curve, first two nodes selected
    On Error Resume Next
    Err.Clear
    Set s = Nothing
    Set s = lay.CreateRectangle2(20, 20, 60, 60)
    s.ConvertToCurves
    Set nr = Nothing
    Set nr = s.Curve.Nodes.All
    W "[4] curve nodes = " & nr.Count
    s.Curve.Nodes(1).Selected = True
    s.Curve.Nodes(2).Selected = True
    Err.Clear
    On Error GoTo 0
    W "[4] built scene, shapes = " & pg.Shapes.Count
  Case "p"
    ' just two rectangles
    On Error Resume Next
    Err.Clear
    Set s = Nothing
    Set s = lay.CreateRectangle2(20, 20, 60, 60)
    Err.Clear
    Set s = Nothing
    Set s = lay.CreateRectangle2(180, 50, 220, 90)
    Err.Clear
    On Error GoTo 0
    W "[4] built scene, shapes = " & pg.Shapes.Count
  Case "g"
    ' three rectangles, grouped into one object
    On Error Resume Next
    Err.Clear
    Set s = Nothing
    Set s = lay.CreateRectangle2(20, 20, 60, 60)
    Err.Clear
    Set s = Nothing
    Set s = lay.CreateRectangle2(100, 30, 140, 70)
    Err.Clear
    Set s = Nothing
    Set s = lay.CreateRectangle2(180, 50, 220, 90)
    Err.Clear
    doc.ClearSelection
    For i = 1 To pg.Shapes.Count
      pg.Shapes.Item(i).AddToSelection
    Next
    app.ActiveSelectionRange.Group
    Err.Clear
    On Error GoTo 0
    W "[4] built scene, shapes = " & pg.Shapes.Count
  Case Else
    On Error Resume Next
    Err.Clear
    Set s = Nothing
    Set s = lay.CreateRectangle2(20, 20, 60, 60)
    Err.Clear
    Set s = Nothing
    Set s = lay.CreateRectangle2(100, 30, 140, 70)
    Err.Clear
    Set s = Nothing
    Set s = lay.CreateRectangle2(180, 50, 220, 90)
    Err.Clear
    If scene = "e" Then
      Set s = Nothing
      Set s = lay.CreateEllipse2(260, 60, 40, 40)
    End If
    Err.Clear
    On Error GoTo 0
    W "[4] built scene, shapes = " & pg.Shapes.Count
End Select

' --- selection ------------------------------------------------
On Error Resume Next
Err.Clear
doc.ClearSelection
If scene <> "n" Then
  For i = 1 To pg.Shapes.Count
    Err.Clear
    Set shp = Nothing
    Set shp = pg.Shapes.Item(i)
    If Err.Number = 0 And Not shp Is Nothing Then shp.AddToSelection
    Err.Clear
  Next
End If
W "[5] selection count = " & app.ActiveSelectionRange.Count
Err.Clear
On Error GoTo 0

DumpScene "BEFORE"

' --- run the reference macro ---------------------------------
W ""
W "[6] RunMacro " & macroFull
On Error Resume Next
Err.Clear
app.GMSManager.RunMacro proj, proc
W "      err=" & Err.Number & "  " & Err.Description
Err.Clear
On Error GoTo 0

On Error Resume Next
Err.Clear
app.Refresh
Err.Clear
On Error GoTo 0

W ""
DumpScene "AFTER"
W "[7] shapes after = " & pg.Shapes.Count

' --- cleanup: close ONLY the scratch document ----------------
On Error Resume Next
Err.Clear
doc.Close
W "[8] scratch doc.Close err=" & Err.Number & " " & Err.Description
W "[8] Documents.Count (after) = " & app.Documents.Count
Err.Clear
On Error GoTo 0

W "==== done (CorelDRAW left running) ===="
logFile.Close
WScript.Quit 0