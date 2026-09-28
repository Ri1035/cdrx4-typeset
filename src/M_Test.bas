Attribute VB_Name = "M_Test"
Option Explicit

'==========================================================
' 自检
'   新建一个临时文档，把每个功能的「纯逻辑入口」跑一遍，
'   结果写进 %TEMP%\typeset_selftest.log
'
'   这个模块不挂工具栏按钮，只给构建 / 冒烟测试用。
'   跑完会把临时文档丢掉，不动用户正在编辑的文档。
'==========================================================

Public Sub SelfTest()
    Dim fso As Object
    Dim doc As Document
    Dim pg As Page
    Dim lay As Object
    Dim log As String
    Dim logPath As String

    logPath = Environ$("TEMP") & "\typeset_selftest.log"
    log = "TypesetToolkit selftest " & Now & vbCrLf

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub

    On Error Resume Next
    Err.Clear
    Set doc = Nothing
    Set doc = CorelDRAW.CreateDocument
    log = log & "0 create doc    err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    If doc Is Nothing Then
        WriteLog fso, logPath, log
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear
    Set pg = doc.ActivePage
    Set lay = pg.ActiveLayer
    log = log & "0 page/layer    err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    If lay Is Nothing Then
        WriteLog fso, logPath, log
        Exit Sub
    End If

    TestSpine doc, log
    TestAlign doc, lay, log
    TestDistribute doc, lay, log
    TestDelSegment doc, lay, log
    TestCropMark doc, lay, log
    TestImpose doc, lay, log

    log = log & "done" & vbCrLf
    WriteLog fso, logPath, log

    On Error Resume Next
    doc.Dirty = False
    doc.Close
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 1. 书脊计算：纯函数，直接验算
'    0.135 * (80/100) * (200/2) = 10.8 mm
'----------------------------------------------------------
Private Sub TestSpine(ByVal doc As Document, ByRef log As String)
    Dim t As Double
    Dim w As Double

    On Error Resume Next
    Err.Clear
    t = M_Spine.SpineThicknessMm(80, 200)
    log = log & "1 spine thick   t=" & Format$(t, "0.###") & " want=10.8 err=" & Err.Number & vbCrLf
    Err.Clear
    w = M_Spine.WeightFromSpine(10.8, 200)
    log = log & "1 spine weight  w=" & Format$(w, "0.#") & " want=80 err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 2. 中心对齐：三个矩形，最底层（最先创建）不动，其余 2 个收敛过去
'----------------------------------------------------------
Private Sub TestAlign(ByVal doc As Document, ByVal lay As Object, ByRef log As String)
    Dim a As Shape
    Dim b As Shape
    Dim c As Shape
    Dim ssr As ShapeRange
    Dim n As Long
    Dim cx As Double
    Dim cy As Double

    On Error Resume Next
    Err.Clear
    Set a = lay.CreateRectangle2(20, 20, 60, 60)
    Set b = lay.CreateRectangle2(100, 30, 140, 70)
    Set c = lay.CreateRectangle2(180, 50, 220, 90)
    doc.ClearSelection
    If Not a Is Nothing Then a.AddToSelection
    If Not b Is Nothing Then b.AddToSelection
    If Not c Is Nothing Then c.AddToSelection
    Set ssr = Nothing
    Set ssr = CorelDRAW.ActiveSelectionRange
    log = log & "2 align select  n=" & ssr.Count & " want=3 err=" & Err.Number & vbCrLf
    Err.Clear

    n = M_Align.AlignCentersCore(ssr)
    log = log & "2 align moved   n=" & n & " want=2 err=" & Err.Number & vbCrLf

    ' 三个对象的中心应当都落在 (50,50)
    cx = 0#: cy = 0#
    If Not c Is Nothing Then
        cx = c.LeftX + c.SizeWidth / 2#
        cy = c.BottomY + c.SizeHeight / 2#
    End If
    log = log & "2 align check   c=(" & Format$(cx, "0.#") & "," & Format$(cy, "0.#") & ") want=(50,50)" & vbCrLf

    Err.Clear
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 3. 按间距分布：固定间距模式，最左不动，其余 2 个重排
'----------------------------------------------------------
Private Sub TestDistribute(ByVal doc As Document, ByVal lay As Object, ByRef log As String)
    Dim a As Shape
    Dim b As Shape
    Dim c As Shape
    Dim ssr As ShapeRange
    Dim n As Long
    Dim gap As Double
    Dim wantGap As Double

    On Error Resume Next
    Err.Clear
    Set a = lay.CreateRectangle2(0, 300, 20, 20)
    Set b = lay.CreateRectangle2(60, 300, 20, 20)
    Set c = lay.CreateRectangle2(160, 300, 20, 20)
    doc.ClearSelection
    If Not a Is Nothing Then a.AddToSelection
    If Not b Is Nothing Then b.AddToSelection
    If Not c Is Nothing Then c.AddToSelection
    Set ssr = Nothing
    Set ssr = CorelDRAW.ActiveSelectionRange
    Err.Clear

    n = M_Distribute.DistributeBySpacingCore(ssr, 5)
    log = log & "3 distrib moved n=" & n & " want=2 err=" & Err.Number & vbCrLf

    ' 相邻对象之间的空隙应当 = 5 mm
    gap = -1#
    wantGap = M_Util.MmToDoc(5)
    If Not a Is Nothing And Not b Is Nothing Then gap = b.LeftX - a.RightX
    log = log & "3 distrib gap   gap=" & Format$(gap, "0.####") & _
          " want=" & Format$(wantGap, "0.####") & vbCrLf

    Err.Clear
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 4. 删除线段：单线段对象整体删除；曲线选中节点则删节点+线段
'----------------------------------------------------------
Private Sub TestDelSegment(ByVal doc As Document, ByVal lay As Object, ByRef log As String)
    Dim ln As Shape
    Dim cv As Shape
    Dim ssr As ShapeRange
    Dim n As Long
    Dim before As Long
    Dim after As Long
    Dim nodesBefore As Long
    Dim nodesAfter As Long

    ' 4a 单线段对象
    On Error Resume Next
    Err.Clear
    before = doc.ActivePage.Shapes.Count
    Set ln = lay.CreateLineSegment(0, 500, 50, 500)
    doc.ClearSelection
    If Not ln Is Nothing Then ln.AddToSelection
    Set ssr = Nothing
    Set ssr = CorelDRAW.ActiveSelectionRange
    log = log & "4a delseg isline=" & CStr(M_Util.IsLineSegment(ln)) & " want=True err=" & Err.Number & vbCrLf
    Err.Clear

    n = M_DelSegment.DelSegmentCore(ssr)
    after = doc.ActivePage.Shapes.Count
    log = log & "4a delseg obj   n=" & n & " want=1  shapes " & before & "->" & after & _
          " want=+1-1 err=" & Err.Number & vbCrLf
    Err.Clear

    ' 4b 曲线选中节点
    Set cv = Nothing
    Err.Clear
    Set cv = lay.CreateRectangle2(0, 700, 40, 40)
    If Not cv Is Nothing Then cv.ConvertToCurves
    nodesBefore = 0
    If Not cv Is Nothing Then nodesBefore = cv.Curve.Nodes.Count
    If Not cv Is Nothing Then
        cv.Curve.Nodes(1).Selected = True
        cv.Curve.Nodes(2).Selected = True
    End If
    doc.ClearSelection
    If Not cv Is Nothing Then cv.AddToSelection
    Set ssr = Nothing
    Set ssr = CorelDRAW.ActiveSelectionRange
    log = log & "4b curve nodes  before=" & nodesBefore & " want=4 err=" & Err.Number & vbCrLf
    Err.Clear

    n = M_DelSegment.DelSegmentCore(ssr)
    nodesAfter = 0
    If Not cv Is Nothing Then nodesAfter = cv.Curve.Nodes.Count
    log = log & "4b delseg node  n=" & n & " want=1  nodes " & nodesBefore & "->" & nodesAfter & _
          " want=4->2 err=" & Err.Number & vbCrLf

    Err.Clear
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 5. 裁切线：一个矩形 -> 4 个角共 8 条线
'----------------------------------------------------------
Private Sub TestCropMark(ByVal doc As Document, ByVal lay As Object, ByRef log As String)
    Dim r As Shape
    Dim ssr As ShapeRange
    Dim n As Long
    Dim before As Long
    Dim after As Long

    On Error Resume Next
    Err.Clear
    Set r = lay.CreateRectangle2(200, 400, 100, 60)
    doc.ClearSelection
    If Not r Is Nothing Then r.AddToSelection
    Set ssr = Nothing
    Set ssr = CorelDRAW.ActiveSelectionRange
    before = doc.ActivePage.Shapes.Count
    Err.Clear

    n = M_CropMark.CropMarkCore(ssr, 3, 1)
    after = doc.ActivePage.Shapes.Count
    ' 8 条线最后会打成一个群组：8 个对象 -> 1 个，所以对象数净 +1
    log = log & "5 cropmark line n=" & n & " want=8  shapes " & before & "->" & after & _
          " want=+1(grouped) err=" & Err.Number & vbCrLf

    Err.Clear
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' 6. 拼版：1 个对象 -> 3 列 x 2 行 = 6 个
'----------------------------------------------------------
Private Sub TestImpose(ByVal doc As Document, ByVal lay As Object, ByRef log As String)
    Dim r As Shape
    Dim ssr As ShapeRange
    Dim n As Long
    Dim before As Long
    Dim after As Long

    On Error Resume Next
    Err.Clear
    Set r = lay.CreateRectangle2(0, 1000, 20, 20)
    doc.ClearSelection
    If Not r Is Nothing Then r.AddToSelection
    Set ssr = Nothing
    Set ssr = CorelDRAW.ActiveSelectionRange
    before = doc.ActivePage.Shapes.Count
    Err.Clear

    n = M_Impose.ImposeCore(ssr, 3, 2, 5, 5)
    after = doc.ActivePage.Shapes.Count
    log = log & "6 impose total  n=" & n & " want=6  shapes " & before & "->" & after & _
          " want=+5 err=" & Err.Number & vbCrLf

    Err.Clear
    On Error GoTo 0
End Sub

Private Sub WriteLog(ByVal fso As Object, ByVal p As String, ByVal S As String)
    Dim ts As Object

    On Error Resume Next
    Set ts = fso.CreateTextFile(p, True, True)
    If Not ts Is Nothing Then
        ts.Write S
        ts.Close
    End If
    On Error GoTo 0
End Sub

'==========================================================
' 详细探针：把环境与几何单位、页面尺寸等事实落盘
' 结果写 %TEMP%\typeset_probe.log
'==========================================================
Public Sub Probe()
    Dim fso As Object
    Dim doc As Document
    Dim pg As Page
    Dim log As String
    Dim logPath As String

    logPath = Environ$("TEMP") & "\typeset_probe.log"
    log = "TypesetToolkit probe " & Now & vbCrLf

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub

    log = log & "version     = " & CorelDRAW.Version & vbCrLf

    On Error Resume Next
    Err.Clear
    Set doc = Nothing
    Set doc = CorelDRAW.CreateDocument
    log = log & "create doc  err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    If doc Is Nothing Then
        WriteLog fso, logPath, log
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear
    Set pg = doc.ActivePage
    log = log & "page size   = " & Round(pg.SizeWidth, 3) & " x " & Round(pg.SizeHeight, 3) & vbCrLf
    log = log & "doc unit    = " & doc.Unit & "   mmPerUnit = " & Round(M_Util.MmPerUnit(), 6) & vbCrLf
    log = log & "mm->doc 1mm = " & Round(M_Util.MmToDoc(1), 6) & vbCrLf
    Err.Clear
    On Error GoTo 0

    On Error Resume Next
    doc.Dirty = False
    doc.Close
    On Error GoTo 0

    log = log & "done" & vbCrLf
    WriteLog fso, logPath, log
End Sub

'==========================================================
' 诊断：把「构建会话」的事实落盘（工程名 / 组件清单）
' 结果写 %TEMP%\typeset_diag.log
'==========================================================
Public Sub Diag()
    Dim fso As Object
    Dim log As String
    Dim logPath As String
    Dim p As Object
    Dim i As Long
    Dim c As Object

    logPath = Environ$("TEMP") & "\typeset_diag.log"
    log = "TypesetToolkit diag " & Now & vbCrLf

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub

    log = log & "version  = " & CorelDRAW.Version & vbCrLf
    log = log & "prj name = " & M_Install.PRJ_NAME & vbCrLf
    log = log & "toolbar  = " & M_Install.TOOLBAR_NAME & vbCrLf

    On Error Resume Next
    Err.Clear
    Set p = Nothing
    For i = 1 To CorelDRAW.VBE.VBProjects.Count
        Set c = Nothing
        Set c = CorelDRAW.VBE.VBProjects.Item(i)
        If Err.Number = 0 And Not c Is Nothing Then
            log = log & "proj " & i & " name=" & c.Name & " comps=" & c.VBComponents.Count & vbCrLf
            If StrComp(c.Name, M_Install.PRJ_NAME, vbTextCompare) = 0 Then Set p = c
        End If
        Err.Clear
    Next i
    Err.Clear

    If Not p Is Nothing Then
        For i = 1 To p.VBComponents.Count
            Set c = Nothing
            Err.Clear
            Set c = p.VBComponents.Item(i)
            If Err.Number = 0 And Not c Is Nothing Then
                log = log & "  comp " & i & " name=" & c.Name & " type=" & c.Type & _
                      " lines=" & c.CodeModule.CountOfLines & vbCrLf
            End If
            Err.Clear
        Next i
    End If
    Err.Clear
    On Error GoTo 0

    log = log & "done" & vbCrLf
    WriteLog fso, logPath, log
End Sub