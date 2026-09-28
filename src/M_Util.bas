Attribute VB_Name = "M_Util"
Option Explicit

'==========================================================
' 排版工具 TypesetToolkit —— 公共函数
' 供其余模块调用，本身不对应工具栏按钮
'
' 约定：涉及「可能只读的属性」时用晚期绑定（As Object），
'       避免早期绑定在编译期报「不能给只读属性赋值」——
'       那会让整个 VBA 工程编译不过、所有宏一起失效。
'
' 【编译红线】不要写 Dim app As Object : Set app = CorelDRAW
'       一律直接用 CorelDRAW.成员
'==========================================================

' 当前是否有打开的文档
Public Function HasDocument() As Boolean
    On Error Resume Next
    HasDocument = (CorelDRAW.Documents.Count > 0)
    On Error GoTo 0
End Function

'==========================================================
' 单位换算
' X4 实测 doc.Unit：1=英寸 2=英尺 3=毫米 4=厘米 5=像素
'                   6=英里 7=米 8=千米 9=迪多点 11=码 12=派卡
' 页面/对象尺寸跟着文档单位走，所以几何类功能必须先换算。
'==========================================================

Public Function MmPerUnit() As Double
    Dim u As Long
    u = 3
    On Error Resume Next
    u = CorelDRAW.ActiveDocument.Unit
    On Error GoTo 0

    Select Case u
        Case 1:   MmPerUnit = 25.4              ' 英寸
        Case 2:   MmPerUnit = 304.8             ' 英尺
        Case 3:   MmPerUnit = 1                 ' 毫米
        Case 4:   MmPerUnit = 10                ' 厘米
        Case 5:   MmPerUnit = 25.4 / 300        ' 像素（按 300dpi 折算）
        Case 6:   MmPerUnit = 1609344           ' 英里
        Case 7:   MmPerUnit = 1000              ' 米
        Case 8:   MmPerUnit = 1000000           ' 千米
        Case 9:   MmPerUnit = 0.3759259         ' 迪多点
        Case 11:  MmPerUnit = 914.4             ' 码
        Case 12:  MmPerUnit = 4.2333333         ' 派卡
        Case Else: MmPerUnit = 1                ' 兜底当毫米
    End Select
End Function

' 毫米 -> 当前文档单位
Public Function MmToDoc(ByVal v As Double) As Double
    MmToDoc = v / MmPerUnit()
End Function

' 当前文档单位 -> 毫米
Public Function DocToMm(ByVal v As Double) As Double
    DocToMm = v * MmPerUnit()
End Function

' 当前页，取不到就退回第一页
Public Function ActivePageSafe() As Page
    Dim doc As Document
    Dim pg As Page

    On Error Resume Next
    Set doc = CorelDRAW.ActiveDocument
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Set pg = doc.ActivePage
    If Err.Number <> 0 Or pg Is Nothing Then
        Err.Clear
        For Each pg In doc.Pages
            Exit For
        Next pg
    End If
    Err.Clear
    On Error GoTo 0
    Set ActivePageSafe = pg
End Function

'==========================================================
' 摆放 / 样式
'==========================================================

' 把形状的包围盒中心移到 (x, y)，坐标是「当前文档单位」
Public Sub CenterAt(ByVal sh As Shape, ByVal x As Double, ByVal y As Double)
    Dim o As Object
    Dim dx As Double
    Dim dy As Double
    Dim ok As Boolean

    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Err.Clear
    Set o = sh
    o.CenterX = x
    o.CenterY = y
    ok = (Err.Number = 0)
    If Not ok Then
        Err.Clear
        dx = x - (o.LeftX + o.SizeWidth / 2)
        dy = y - (o.BottomY + o.SizeHeight / 2)
        o.Move dx, dy
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' 刷新界面
Public Sub DoRefresh()
    On Error Resume Next
    CorelDRAW.Refresh
    On Error GoTo 0
End Sub

'==========================================================
' 选择集
'==========================================================

Public Function SelectionCount() As Long
    Dim n As Long
    n = 0
    On Error Resume Next
    n = CorelDRAW.ActiveSelectionRange.Count
    On Error GoTo 0
    SelectionCount = n
End Function

' 选中数量不足 minN 就弹提示并返回 False
Public Function RequireSelection(ByVal minN As Long, ByVal what As String) As Boolean
    Dim n As Long

    RequireSelection = False
    If Not HasDocument() Then
        MsgBox "请先打开或新建一个文档。", vbExclamation, "排版工具"
        Exit Function
    End If

    n = SelectionCount()
    If n < minN Then
        MsgBox "请先选中至少 " & minN & " 个对象，再执行「" & what & "」。", _
               vbExclamation, "排版工具"
        Exit Function
    End If
    RequireSelection = True
End Function

'==========================================================
' 颜色小工具
'==========================================================

' 造一个 CMYK 颜色（X4 里 CreateCMYKColor 只能从 Application 上取）
Public Function MakeCMYK(ByVal cc As Long, ByVal mm As Long, _
                         ByVal yy As Long, ByVal kk As Long) As Color
    Dim col As Color
    On Error Resume Next
    Set col = CorelDRAW.CreateCMYKColor(cc, mm, yy, kk)
    On Error GoTo 0
    Set MakeCMYK = col
End Function

' 造一个 RGB 颜色
Public Function MakeRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Color
    Dim col As Color
    On Error Resume Next
    Set col = CorelDRAW.CreateRGBColor(r, g, b)
    On Error GoTo 0
    Set MakeRGB = col
End Function

' 套版色（registration）。X4 上取不到就退回 K100 的四色叠加
Public Function RegistrationColor() As Color
    Dim col As Color
    On Error Resume Next
    Err.Clear
    Set col = CorelDRAW.CreateRegistrationColor
    If Err.Number <> 0 Or col Is Nothing Then
        Err.Clear
        Set col = CorelDRAW.CreateCMYKColor(100, 100, 100, 100)
    End If
    Err.Clear
    On Error GoTo 0
    Set RegistrationColor = col
End Function

' 两个颜色是否一样
Public Function SameColor(ByVal a As Color, ByVal b As Color) As Boolean
    Dim ok As Boolean
    If a Is Nothing Or b Is Nothing Then
        SameColor = False
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    ok = a.IsSame(b)
    If Err.Number <> 0 Then
        Err.Clear
        ok = (a.RGBRed = b.RGBRed) And (a.RGBGreen = b.RGBGreen) And (a.RGBBlue = b.RGBBlue)
        Err.Clear
    End If
    On Error GoTo 0
    SameColor = ok
End Function

'==========================================================
' 杂项
'==========================================================

Public Function DesktopPath() As String
    On Error Resume Next
    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0
End Function

Public Function BaseName(ByVal p As String) As String
    Dim n As String
    Dim i As Long

    If Len(p) = 0 Then
        BaseName = "未命名"
        Exit Function
    End If

    n = p
    i = InStrRev(n, "\")
    If i > 0 Then n = Mid$(n, i + 1)
    i = InStrRev(n, ".")
    If i > 0 Then n = Left$(n, i - 1)
    BaseName = n
End Function

' 取一个形状的曲线对象（不是曲线返回 Nothing）
Public Function ShapeCurve(ByVal sh As Shape) As Object
    Dim o As Object
    Dim cv As Object

    If sh Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    Set o = sh
    Set cv = o.Curve
    If Err.Number <> 0 Then Set cv = Nothing
    Err.Clear
    On Error GoTo 0
    Set ShapeCurve = cv
End Function

' 一个形状是不是「单根直线段」：曲线只有 1 条子路径、2 个节点
Public Function IsLineSegment(ByVal sh As Shape) As Boolean
    Dim cv As Object
    Dim ok As Boolean

    IsLineSegment = False
    If sh Is Nothing Then Exit Function

    Set cv = ShapeCurve(sh)
    If cv Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    ok = (cv.SubPaths.Count = 1) And (cv.Nodes.Count = 2)
    Err.Clear
    On Error GoTo 0
    IsLineSegment = ok
End Function

' 在文档里插入一段居中文字（供计算类功能把结果落到版面）
Public Function InsertNote(ByVal txt As String, ByVal xmm As Double, ByVal ymm As Double) As Shape
    Dim lay As Object
    Dim pg As Page
    Dim sh As Shape

    Set InsertNote = Nothing
    If Not HasDocument() Then Exit Function

    Set pg = ActivePageSafe()
    If pg Is Nothing Then Exit Function
    On Error Resume Next
    Set lay = pg.ActiveLayer
    On Error GoTo 0
    If lay Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    Set sh = lay.CreateArtisticTextWide(MmToDoc(xmm), MmToDoc(ymm), txt)
    Err.Clear
    On Error GoTo 0
    Set InsertNote = sh
End Function

'==========================================================
' 几何 / 身份小工具（供各功能模块共用）
'==========================================================

' 浮点近似相等。位置比较一律用它，别写 = 。
Public Function NearEq(ByVal a As Double, ByVal b As Double) As Boolean
    Dim d As Double
    d = a - b
    If d < 0 Then d = -d
    NearEq = (d < 0.0001)
End Function

' 形状的 StaticID（文档内唯一）。取不到返回 0。
' 用途：判断「某个形状在不在选择集里」——比坐标比较可靠。
Public Function StaticIdOf(ByVal sh As Shape) As Long
    Dim o As Object
    Dim v As Long
    StaticIdOf = 0
    If sh Is Nothing Then Exit Function
    v = 0
    On Error Resume Next
    Err.Clear
    Set o = sh
    v = o.StaticID
    If Err.Number <> 0 Then v = 0
    Err.Clear
    On Error GoTo 0
    StaticIdOf = v
End Function

' 选择集的包围盒。返回 False = 没有可用对象。
Public Function SelectionBounds(ByVal ssr As ShapeRange, _
        ByRef minL As Double, ByRef maxR As Double, _
        ByRef minB As Double, ByRef maxT As Double) As Boolean
    Dim i As Long
    Dim n As Long
    Dim sh As Shape
    Dim got As Boolean

    SelectionBounds = False
    If ssr Is Nothing Then Exit Function

    n = 0
    On Error Resume Next
    n = ssr.Count
    On Error GoTo 0
    If n < 1 Then Exit Function

    got = False
    For i = 1 To n
        Set sh = Nothing
        On Error Resume Next
        Set sh = ssr.Item(i)
        On Error GoTo 0
        If Not sh Is Nothing Then
            If Not got Then
                minL = sh.LeftX
                maxR = sh.RightX
                minB = sh.BottomY
                maxT = sh.TopY
                got = True
            Else
                If sh.LeftX < minL Then minL = sh.LeftX
                If sh.RightX > maxR Then maxR = sh.RightX
                If sh.BottomY < minB Then minB = sh.BottomY
                If sh.TopY > maxT Then maxT = sh.TopY
            End If
        End If
    Next i

    SelectionBounds = got
End Function

' 把形状变成曲线（已经是曲线就不动）。返回是否可用作曲线。
' 注意：矩形/椭圆等 ConvertToCurves 后 Type 变成 cdrCurveShape(3)。
Public Function EnsureCurve(ByVal sh As Shape) As Boolean
    Dim o As Object
    Dim t As Long
    Dim ok As Boolean

    EnsureCurve = False
    If sh Is Nothing Then Exit Function

    ok = False
    On Error Resume Next
    Err.Clear
    Set o = sh
    t = o.Type
    If Err.Number = 0 Then
        If t = cdrCurveShape Then
            ok = True
        Else
            Err.Clear
            o.ConvertToCurves
            If Err.Number = 0 Then ok = True
        End If
    End If
    Err.Clear
    On Error GoTo 0
    EnsureCurve = ok
End Function

' 选择集里「最底层」的那个形状。
'
' 依据（probe_ref 实测）：X4 的 Page.Shapes 是「前 -> 后」排序，
' Item(1) 是最前面（最后创建）的那个，所以从 Count 往回扫，
' 第一个出现在选择集里的就是最底层。
' 用 StaticID 判断归属，不用坐标。
Public Function BottomMostOf(ByVal ssr As ShapeRange) As Shape
    Dim ids() As Long
    Dim n As Long
    Dim i As Long
    Dim pg As Page
    Dim coll As Object
    Dim m As Long
    Dim sh As Shape
    Dim v As Long
    Dim k As Long
    Dim hit As Boolean

    Set BottomMostOf = Nothing
    If ssr Is Nothing Then Exit Function

    n = 0
    On Error Resume Next
    n = ssr.Count
    On Error GoTo 0
    If n < 1 Then Exit Function

    ReDim ids(1 To n)
    For i = 1 To n
        Set sh = Nothing
        On Error Resume Next
        Set sh = ssr.Item(i)
        On Error GoTo 0
        ids(i) = StaticIdOf(sh)
    Next i

    ' 先按 StaticID 从后往前扫
    Set pg = ActivePageSafe()
    If Not pg Is Nothing Then
        Set coll = Nothing
        On Error Resume Next
        Set coll = pg.Shapes
        On Error GoTo 0
        If Not coll Is Nothing Then
            m = 0
            On Error Resume Next
            m = coll.Count
            On Error GoTo 0
            For i = m To 1 Step -1
                Set sh = Nothing
                On Error Resume Next
                Set sh = coll.Item(i)
                On Error GoTo 0
                If Not sh Is Nothing Then
                    v = StaticIdOf(sh)
                    If v <> 0 Then
                        hit = False
                        For k = 1 To n
                            If ids(k) = v Then
                                hit = True
                                Exit For
                            End If
                        Next k
                        If hit Then
                            Set BottomMostOf = sh
                            Exit Function
                        End If
                    End If
                End If
            Next i
        End If
    End If

    ' StaticID 不可用（都为 0）时退回：取选择集最后一项
    On Error Resume Next
    Set BottomMostOf = ssr.Item(n)
    On Error GoTo 0
End Function