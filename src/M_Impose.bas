Attribute VB_Name = "M_Impose"
Option Explicit

'==========================================================
' 拼版
'
' 复用可读参考 ZeroBase\ArrangeForm（CommandButton1_Click +
' arrange_Clone / arrange_Clone_one）的矩阵复制算法：
'
'     先用 ShapeRange.StepAndRepeat 把选区横向复制成一行，
'     再把整行纵向复制成矩阵。
'
'   matrix = Array(列数, 行数, 列间距, 行间距)
'
' 原版是 UserForm（4 个 TextBox）；X4 的 .bas 导不进窗体，所以
' 降级成一次 InputBox 收 4 个数（流程等价）。
'
' 间距 = 对象之间的空隙；复制步长 = 对象尺寸 + 间距。
' 结果：原选区留在左下第一格，其余是副本，副本被选中。
'==========================================================

Public Sub Impose()
    Dim ssr As ShapeRange
    Dim s As String
    Dim a As Variant
    Dim cols As Long
    Dim rows As Long
    Dim gx As Double
    Dim gy As Double
    Dim n As Long

    If Not M_Util.RequireSelection(1, "拼版") Then Exit Sub

    s = InputBox("拼版：把选中的对象按矩阵复制铺开" & vbCrLf & vbCrLf & _
                 "格式：列数 行数 列间距(mm) 行间距(mm)" & vbCrLf & _
                 "例如：3 2 5 5" & vbCrLf & vbCrLf & _
                 "列 = 横向个数，行 = 纵向个数，间距 = 对象之间的空隙。", _
                 "拼版", "3 2 5 5")
    If Len(Trim$(s)) = 0 Then Exit Sub

    ' 允许 3x2 / 3*2 / 3,2 / 3，2 等写法
    s = Replace(s, "x", " ")
    s = Replace(s, "X", " ")
    s = Replace(s, "*", " ")
    s = Replace(s, ",", " ")
    s = Replace(s, ChrW(&HFF0C), " ")      ' 中文逗号
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)

    a = Split(s)
    cols = 3: rows = 2: gx = 5#: gy = 5#
    If UBound(a) >= 0 Then If Val(a(0)) > 0 Then cols = CLng(Val(a(0)))
    If UBound(a) >= 1 Then If Val(a(1)) > 0 Then rows = CLng(Val(a(1)))
    If UBound(a) >= 2 Then gx = Val(a(2))
    If UBound(a) >= 3 Then gy = Val(a(3))
    If cols < 1 Then cols = 1
    If rows < 1 Then rows = 1
    If gx < 0 Then gx = 0#
    If gy < 0 Then gy = 0#
    If cols = 1 And rows = 1 Then Exit Sub

    Set ssr = Nothing
    On Error Resume Next
    Set ssr = CorelDRAW.ActiveSelectionRange
    On Error GoTo 0
    If ssr Is Nothing Then Exit Sub

    n = ImposeCore(ssr, cols, rows, gx, gy)
    M_Util.DoRefresh
    If n = 0 Then MsgBox "拼版没有产生结果。", vbExclamation, "排版工具"
End Sub

' 返回矩阵里最终的对象总数（= 原对象数 x 列 x 行）
Public Function ImposeCore(ByVal ssr As ShapeRange, _
        ByVal cols As Long, ByVal rows As Long, _
        ByVal gapXMm As Double, ByVal gapYMm As Double) As Long
    Dim doc As Document
    Dim base As New ShapeRange
    Dim rowRange As ShapeRange
    Dim all As ShapeRange
    Dim dup As ShapeRange
    Dim i As Long
    Dim n As Long
    Dim sh As Shape
    Dim w As Double
    Dim h As Double
    Dim gx As Double
    Dim gy As Double

    ImposeCore = 0
    If ssr Is Nothing Then Exit Function
    If cols < 1 Or rows < 1 Then Exit Function
    If cols = 1 And rows = 1 Then Exit Function

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

    ' 先把选区快照进一个独立 ShapeRange：StepAndRepeat 会改动选择集，
    ' 直接拿 ActiveSelectionRange 当基准会被它带着跑。
    For i = 1 To n
        Set sh = Nothing
        On Error Resume Next
        Set sh = ssr.Item(i)
        On Error GoTo 0
        If Not sh Is Nothing Then base.Add sh
    Next i
    If base.Count < 1 Then Exit Function

    w = 0#: h = 0#
    On Error Resume Next
    w = base.SizeWidth
    h = base.SizeHeight
    On Error GoTo 0
    If w <= 0# Or h <= 0# Then Exit Function

    gx = M_Util.MmToDoc(gapXMm)
    gy = M_Util.MmToDoc(gapYMm)

    Set rowRange = base

    On Error Resume Next
    Err.Clear
    doc.BeginCommandGroup "拼版"

    ' 横向复制成一行
    If cols > 1 Then
        Set dup = Nothing
        Set dup = base.StepAndRepeat(cols - 1, w + gx, 0#)
        If Not dup Is Nothing Then
            Set rowRange = Nothing
            Set rowRange = doc.CreateShapeRangeFromArray(base, dup)
        End If
    End If

    ' rows = 1 时 all 必须指向「整行」，否则最后只会选中原对象
    Set all = rowRange

    ' 纵向复制成矩阵
    If rows > 1 Then
        Set dup = Nothing
        Set dup = rowRange.StepAndRepeat(rows - 1, 0#, -(h + gy))
        If Not dup Is Nothing Then
            Set all = Nothing
            Set all = doc.CreateShapeRangeFromArray(rowRange, dup)
        End If
    End If

    If Not all Is Nothing Then
        all.CreateSelection
        ImposeCore = all.Count
    End If

    Err.Clear
    doc.EndCommandGroup
    On Error GoTo 0
End Function