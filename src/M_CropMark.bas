Attribute VB_Name = "M_CropMark"
Option Explicit

'==========================================================
' 裁切线
'
' 参考宏 GlobalMacros.CropMark.裁切线 已加密。侦察
' （tools\probe_ref.vbs 2，scene=n，页面上 3 个矩形、无选中）时
' 它 err=0 但页面上没有多出任何对象 —— 说明它是「按选中对象」
' 工作的（没有选择就静默返回），而不是围绕页面。
'
' 实现：围绕「选中对象的包围盒」（= 裁切框 / trim box）的四个角
' 各画一个 L 形角线（裁切标记），角线整体外移 offset 毫米。
'
' 这是印刷「裁切线 / 裁切标记」的标准含义，也和模块名 CropMark 一致。
'
'   ┌ ─ ─ ─ ─ ─ ─ ┐        每角两条线：
'                （角点在裁切框外 offset 处）
'    │  裁切框     │        颜色 = 套版色（registration）
'                │        线宽 = 0.25 mm
'   └ ─ ─ ─ ─ ─ ─ ┘        画完打成一个群组，方便整体移动/删除
'==========================================================

Public Sub CropMark()
    Dim ssr As ShapeRange
    Dim s1 As String
    Dim s2 As String
    Dim markLen As Double
    Dim off As Double
    Dim n As Long

    If Not M_Util.RequireSelection(1, "裁切线") Then Exit Sub

    Set ssr = Nothing
    On Error Resume Next
    Set ssr = CorelDRAW.ActiveSelectionRange
    On Error GoTo 0
    If ssr Is Nothing Then Exit Sub

    s1 = InputBox("裁切线：角线长度（mm）", "裁切线", "3")
    If Len(Trim$(s1)) = 0 Then Exit Sub
    s2 = InputBox("裁切线：角线与裁切框的间距 / 出血（mm）", "裁切线", "1")
    If Len(Trim$(s2)) = 0 Then Exit Sub

    markLen = Val(s1)
    off = Val(s2)
    If markLen <= 0 Then markLen = 3#
    If off < 0 Then off = 0#

    n = CropMarkCore(ssr, markLen, off)
    M_Util.DoRefresh
    If n = 0 Then
        MsgBox "没有画出裁切线，请确认已选中对象。", vbExclamation, "排版工具"
    End If
End Sub

' 返回画出的线段条数（正常为 8 = 4 个角 x 2 条线）
Public Function CropMarkCore(ByVal ssr As ShapeRange, _
        ByVal markLenMm As Double, ByVal offsetMm As Double) As Long
    Dim doc As Document
    Dim pg As Page
    Dim lay As Object
    Dim sr As New ShapeRange
    Dim sh As Shape
    Dim minL As Double
    Dim maxR As Double
    Dim minB As Double
    Dim maxT As Double
    Dim markLen As Double
    Dim off As Double
    Dim lines As Long
    Dim col As Color

    CropMarkCore = 0
    If ssr Is Nothing Then Exit Function

    Set doc = Nothing
    On Error Resume Next
    Set doc = CorelDRAW.ActiveDocument
    On Error GoTo 0
    If doc Is Nothing Then Exit Function

    If Not M_Util.SelectionBounds(ssr, minL, maxR, minB, maxT) Then Exit Function

    Set pg = M_Util.ActivePageSafe()
    If pg Is Nothing Then Exit Function
    Set lay = Nothing
    On Error Resume Next
    Set lay = pg.ActiveLayer
    On Error GoTo 0
    If lay Is Nothing Then Exit Function

    markLen = M_Util.MmToDoc(markLenMm)
    off = M_Util.MmToDoc(offsetMm)

    lines = 0
    On Error Resume Next
    doc.BeginCommandGroup "裁切线"

    ' 左下角
    If AddCorner(sr, lay, minL - off, minB - off, -1#, -1#, markLen) Then lines = lines + 2
    ' 右下角
    If AddCorner(sr, lay, maxR + off, minB - off, 1#, -1#, markLen) Then lines = lines + 2
    ' 左上角
    If AddCorner(sr, lay, minL - off, maxT + off, -1#, 1#, markLen) Then lines = lines + 2
    ' 右上角
    If AddCorner(sr, lay, maxR + off, maxT + off, 1#, 1#, markLen) Then lines = lines + 2
    Err.Clear
    On Error GoTo 0

    If sr.Count > 0 Then
        Set col = M_Util.RegistrationColor()
        On Error Resume Next
        ' 用属性赋值而不是 SetOutlineProperties(..., Color:=...)：后者带
        ' 命名参数，对 TLB 引用敏感；属性写法在 M_Seal 等模块里已验证可用。
        sr.Outline.Width = M_Util.MmToDoc(0.25)
        If Not col Is Nothing Then sr.Outline.Color.CopyAssign col
        sr.CreateSelection
        sr.Group
        Err.Clear
        On Error GoTo 0
    End If

    On Error Resume Next
    doc.EndCommandGroup
    On Error GoTo 0

    CropMarkCore = lines
End Function

' 在一个角上画 L 形角线：sx / sy 是角线伸出的方向（-1 或 +1）
Private Function AddCorner(ByVal sr As ShapeRange, ByVal lay As Object, _
        ByVal x As Double, ByVal y As Double, _
        ByVal sx As Double, ByVal sy As Double, _
        ByVal markLen As Double) As Boolean
    Dim h As Shape
    Dim v As Shape

    AddCorner = False
    If sr Is Nothing Or lay Is Nothing Then Exit Function

    Set h = Nothing
    Set v = Nothing
    On Error Resume Next
    Err.Clear
    Set h = lay.CreateLineSegment(x, y, x + sx * markLen, y)
    Err.Clear
    Set v = lay.CreateLineSegment(x, y, x, y + sy * markLen)
    Err.Clear
    On Error GoTo 0

    If Not h Is Nothing Then sr.Add h
    If Not v Is Nothing Then sr.Add v
    If Not h Is Nothing And Not v Is Nothing Then AddCorner = True
End Function