Attribute VB_Name = "M_DelSegment"
Option Explicit

'==========================================================
' 删除线段
'
' 参考宏 GlobalMacros.DelSegment.删除线段 在加密的 GlobalMacros.gms
' 里，源码读不出来。侦察（tools\probe_ref.vbs 6，scene=c）时它在
' 「一条转曲后的矩形 + 选中前两个节点」的文档上弹出了模态框，
' 而且没有产生任何几何变化 —— 说明它要么只认「线段对象」，
' 要么对选择有额外要求。这里按「线段」的两种常见含义实现，
' 两种可以同时命中：
'
'   1. 直线段对象：由单条线段组成的曲线（1 条子路径 + 2 个节点）
'      → 整个删掉。
'   2. 曲线上的选中节点：用形状工具在曲线上选中 >= 2 个节点后执行
'      → 删除这些节点以及它们之间的线段（等价于 CorelDRAW 的
'        「删除节点」，NodeRange.Delete）。
'
' X4 的 API 里 Segment / SegmentRange 都没有 Delete
' （见 tools\tlb_dump.ps1 对 IDrawSegment 的转储），删几何只能走
' NodeRange.Delete，所以第 2 种用法会连带删掉节点。
'==========================================================

Public Sub DelSegment()
    Dim ssr As ShapeRange
    Dim n As Long

    If Not M_Util.RequireSelection(1, "删除线段") Then Exit Sub

    Set ssr = Nothing
    On Error Resume Next
    Set ssr = CorelDRAW.ActiveSelectionRange
    On Error GoTo 0
    If ssr Is Nothing Then Exit Sub

    n = DelSegmentCore(ssr)
    M_Util.DoRefresh
    If n = 0 Then
        MsgBox "没有找到可删除的线段。" & vbCrLf & vbCrLf & _
               "用法一：直接选中若干「直线段」对象（一根直线）。" & vbCrLf & _
               "用法二：用形状工具(F10)在曲线上选中节点，再执行。", _
               vbExclamation, "排版工具"
    End If
End Sub

' 返回删除动作数：删掉的对象数 + 处理的节点组数
Public Function DelSegmentCore(ByVal ssr As ShapeRange) As Long
    Dim doc As Document
    Dim i As Long
    Dim n As Long
    Dim sh As Shape
    Dim cv As Object
    Dim nr As Object
    Dim cnt As Long
    Dim done As Long

    DelSegmentCore = 0
    If ssr Is Nothing Then Exit Function

    Set doc = Nothing
    On Error Resume Next
    Set doc = CorelDRAW.ActiveDocument
    On Error GoTo 0
    If doc Is Nothing Then Exit Function

    n = 0
    On Error Resume Next
    n = ssr.Count
    On Error GoTo 0
    If n < 1 Then Exit Function

    done = 0
    On Error Resume Next
    doc.BeginCommandGroup "删除线段"
    On Error GoTo 0

    For i = 1 To n
        Set sh = Nothing
        On Error Resume Next
        Set sh = ssr.Item(i)
        On Error GoTo 0

        If Not sh Is Nothing Then
            ' 只处理曲线对象；矩形/椭圆等先不隐式转曲，避免误伤
            If sh.Type = cdrCurveShape Then
                cnt = SelectedNodeCount(sh)
                If cnt >= 2 Then
                    ' 用法二：删掉选中节点及其之间的线段
                    Set cv = M_Util.ShapeCurve(sh)
                    If Not cv Is Nothing Then
                        Set nr = Nothing
                        On Error Resume Next
                        Err.Clear
                        Set nr = cv.Selection
                        Err.Clear
                        On Error GoTo 0
                        If Not nr Is Nothing Then
                            On Error Resume Next
                            nr.Delete
                            On Error GoTo 0
                            done = done + 1
                        End If
                    End If
                ElseIf M_Util.IsLineSegment(sh) Then
                    ' 用法一：单线段对象，整个删掉
                    On Error Resume Next
                    sh.Delete
                    On Error GoTo 0
                    done = done + 1
                End If
            End If
        End If
    Next i

    On Error Resume Next
    doc.EndCommandGroup
    On Error GoTo 0

    DelSegmentCore = done
End Function

' 形状当前选中的节点数（取不到返回 0）
Private Function SelectedNodeCount(ByVal sh As Shape) As Long
    Dim cv As Object
    Dim nr As Object
    Dim cnt As Long

    SelectedNodeCount = 0
    If sh Is Nothing Then Exit Function

    Set cv = M_Util.ShapeCurve(sh)
    If cv Is Nothing Then Exit Function

    cnt = 0
    On Error Resume Next
    Err.Clear
    Set nr = Nothing
    Set nr = cv.Selection
    If Err.Number = 0 And Not nr Is Nothing Then cnt = nr.Count
    Err.Clear
    On Error GoTo 0
    If cnt < 0 Then cnt = 0
    SelectedNodeCount = cnt
End Function