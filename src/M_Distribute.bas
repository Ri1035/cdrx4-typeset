Attribute VB_Name = "M_Distribute"
Option Explicit

'==========================================================
' 按间距分布
'
' 复用 hongwenjun/corelvba（公有领域）里的两个排列算法：
'   Simple_Train_Arrangement  —— 固定间距的「火车排列」
'   Average_Distance          —— 首尾不动、等间距分布
'
' 两者都会：先按 LeftX 从左到右排序，再把垂直方向统一到
' 最左对象的中心线。区别只在水平间距怎么算。
'
' X4 的 ShapeRange 没有 Sort 方法（参考实现靠 lyvba32.dll
' 算法库排序），这里不引外部依赖，用冒泡排序。
'==========================================================

Public Sub DistributeBySpacing()
    Dim ssr As ShapeRange
    Dim mode As String
    Dim s As String
    Dim sp As Double
    Dim n As Long

    If Not M_Util.RequireSelection(2, "按间距分布") Then Exit Sub

    mode = InputBox("按间距分布" & vbCrLf & vbCrLf & _
                    "1 = 固定间距：最左的对象不动，其余依次排到「前一个右边缘 + 间距」" & vbCrLf & _
                    "2 = 平均分布：最左、最右两个不动，中间的等间距铺开" & vbCrLf & vbCrLf & _
                    "请输入 1 或 2：", "按间距分布", "1")
    mode = Trim$(mode)
    If Len(mode) = 0 Then Exit Sub

    Set ssr = Nothing
    On Error Resume Next
    Set ssr = CorelDRAW.ActiveSelectionRange
    On Error GoTo 0
    If ssr Is Nothing Then Exit Sub

    If mode = "2" Then
        n = DistributeEvenlyCore(ssr)
    Else
        s = InputBox("请输入对象之间的间距（mm）：", "按间距分布", "3")
        If Len(Trim$(s)) = 0 Then Exit Sub
        sp = Val(s)
        If sp < 0 Then
            MsgBox "间距不能为负数。", vbExclamation, "排版工具"
            Exit Sub
        End If
        n = DistributeBySpacingCore(ssr, sp)
    End If

    M_Util.DoRefresh
    If n = 0 Then MsgBox "没有可分布的对象。", vbExclamation, "排版工具"
End Sub

'----------------------------------------------------------
' 模式 1：固定间距（火车排列）
' 返回被重新摆放的对象数
'----------------------------------------------------------
Public Function DistributeBySpacingCore(ByVal ssr As ShapeRange, ByVal spacingMm As Double) As Long
    Dim doc As Document
    Dim arr() As Shape
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As Shape
    Dim sh As Shape
    Dim sp As Double
    Dim baseY As Double
    Dim cursorX As Double
    Dim dx As Double
    Dim dy As Double

    DistributeBySpacingCore = 0
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
    If n < 2 Then Exit Function

    ReDim arr(1 To n)
    For i = 1 To n
        Set arr(i) = Nothing
        On Error Resume Next
        Set arr(i) = ssr.Item(i)
        On Error GoTo 0
        If arr(i) Is Nothing Then Exit Function
    Next i

    SortByLeftX arr, n

    sp = M_Util.MmToDoc(spacingMm)
    baseY = arr(1).CenterY
    cursorX = arr(1).RightX

    On Error Resume Next
    doc.BeginCommandGroup "按间距分布"
    On Error GoTo 0

    For i = 2 To n
        Set sh = arr(i)
        dx = (cursorX + sp) - sh.LeftX
        dy = baseY - sh.CenterY
        On Error Resume Next
        sh.Move dx, dy
        On Error GoTo 0
        cursorX = cursorX + sp + sh.SizeWidth
        DistributeBySpacingCore = DistributeBySpacingCore + 1
    Next i

    On Error Resume Next
    doc.EndCommandGroup
    On Error GoTo 0
End Function

'----------------------------------------------------------
' 模式 2：平均分布（首尾不动，中心等距）
' 返回被重新摆放的对象数
'----------------------------------------------------------
Public Function DistributeEvenlyCore(ByVal ssr As ShapeRange) As Long
    Dim doc As Document
    Dim arr() As Shape
    Dim n As Long
    Dim i As Long
    Dim tmp As Shape
    Dim first As Double
    Dim last As Double
    Dim interval As Double
    Dim cur As Double
    Dim baseY As Double

    DistributeEvenlyCore = 0
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
    If n < 3 Then Exit Function

    ReDim arr(1 To n)
    For i = 1 To n
        Set arr(i) = Nothing
        On Error Resume Next
        Set arr(i) = ssr.Item(i)
        On Error GoTo 0
        If arr(i) Is Nothing Then Exit Function
    Next i

    SortByLeftX arr, n

    first = arr(1).CenterX
    last = arr(n).CenterX
    baseY = arr(1).CenterY
    interval = (last - first) / (n - 1)

    On Error Resume Next
    doc.BeginCommandGroup "按间距分布（平均）"
    On Error GoTo 0

    cur = first
    For i = 1 To n
        On Error Resume Next
        arr(i).CenterY = baseY
        arr(i).CenterX = cur
        On Error GoTo 0
        cur = cur + interval
        If i > 1 And i < n Then DistributeEvenlyCore = DistributeEvenlyCore + 1
    Next i

    On Error Resume Next
    doc.EndCommandGroup
    On Error GoTo 0
End Function

'----------------------------------------------------------
' 按 LeftX 升序冒泡（X4 无 ShapeRange.Sort）
'----------------------------------------------------------
Private Sub SortByLeftX(ByRef arr() As Shape, ByVal n As Long)
    Dim i As Long
    Dim j As Long
    Dim tmp As Shape

    For i = 1 To n - 1
        For j = 1 To n - i
            If arr(j).LeftX > arr(j + 1).LeftX Then
                Set tmp = arr(j)
                Set arr(j) = arr(j + 1)
                Set arr(j + 1) = tmp
            End If
        Next j
    Next i
End Sub