Attribute VB_Name = "M_Align"
Option Explicit

'==========================================================
' 中心对齐
'
' 语义（probe_ref 4 实测）：
'   选中若干对象后执行，所有对象的包围盒中心收敛到
'   「最底层那个对象」的中心。
'
'   实测数据（三个矩形，中心分别是 (290,95) (170,65) (50,50)）：
'     执行后 (290,95) -> (50,50)
'            (170,65) -> (50,50)
'            (50,50)  -> 不动
'   最底层 = 最先创建 = 居中不动的那一个。
'
' 最底层的判定见 M_Util.BottomMostOf（按 Page.Shapes 的
' 「前 -> 后」顺序 + StaticID 反查）。
'==========================================================

Public Sub CenterAlign()
    Dim ssr As ShapeRange
    Dim n As Long

    If Not M_Util.RequireSelection(1, "中心对齐") Then Exit Sub

    Set ssr = Nothing
    On Error Resume Next
    Set ssr = CorelDRAW.ActiveSelectionRange
    On Error GoTo 0
    If ssr Is Nothing Then Exit Sub

    n = AlignCentersCore(ssr)
    M_Util.DoRefresh
    If n = 0 Then MsgBox "没有可对齐的对象。", vbExclamation, "排版工具"
End Sub

' 返回实际移动的对象数
Public Function AlignCentersCore(ByVal ssr As ShapeRange) As Long
    Dim doc As Document
    Dim ref As Shape
    Dim i As Long
    Dim n As Long
    Dim sh As Shape
    Dim cx As Double
    Dim cy As Double
    Dim moved As Long

    AlignCentersCore = 0
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

    ' 基准 = 最底层对象；它的中心就是所有对象的落点
    Set ref = M_Util.BottomMostOf(ssr)
    If ref Is Nothing Then Exit Function

    cx = ref.LeftX + ref.SizeWidth / 2#
    cy = ref.BottomY + ref.SizeHeight / 2#

    moved = 0
    On Error Resume Next
    doc.BeginCommandGroup "中心对齐"
    On Error GoTo 0

    For i = 1 To n
        Set sh = Nothing
        On Error Resume Next
        Set sh = ssr.Item(i)
        On Error GoTo 0
        If Not sh Is Nothing Then
            ' 已经在落点上的（也就是基准自己）不用动，省一次写操作
            If Not (M_Util.NearEq(sh.LeftX + sh.SizeWidth / 2#, cx) And _
                    M_Util.NearEq(sh.BottomY + sh.SizeHeight / 2#, cy)) Then
                M_Util.CenterAt sh, cx, cy
                moved = moved + 1
            End If
        End If
    Next i

    On Error Resume Next
    doc.EndCommandGroup
    On Error GoTo 0

    AlignCentersCore = moved
End Function