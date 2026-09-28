Attribute VB_Name = "M_Install"
Option Explicit

'==========================================================
' 排版工具 TypesetToolkit —— 工具栏安装 / 卸载
'
' 重要：VBA 项目名必须为 TypesetToolkit，
'       否则下面拼出来的命令路径会失效。
'
' 【编译红线】绝对不要写
'       Dim app As Object
'       Set app = CorelDRAW
' X4 的 VBA 把 CorelDRAW 这个全局对象整体赋给变量，会直接报
' 「编译错误：类型不匹配」（VBE 会高亮 CorelDRAW 这个词）。
' 而 VBA 是「一处编译不过 → 整个工程所有宏全废」，于是每次启动
' CorelDRAW 都弹编译错误框、工具栏永远装不上。
'
' 正确写法是直接用 `CorelDRAW.成员`；要晚期绑定就把「取回来的成员」
' 放进 Object 变量，例如：
'       Dim cb As Object
'       Set cb = CorelDRAW.CommandBars(TOOLBAR_NAME)
'==========================================================

Public Const PRJ_NAME As String = "TypesetToolkit"
Public Const TOOLBAR_NAME As String = "排版工具"

' 命令清单：模块.过程 | 按钮名 | 提示
' 顺序与参考插件工作区里的 [排版] 弹出菜单一致：
'     拼版 → 裁切线 → 书脊计算 → 中心对齐 → 按间距分布 → 删除线段
Private Function CmdList() As Variant
    CmdList = Array( _
        Array("M_Impose.Impose", "拼版", "把选中的对象按 行x列 矩阵复制铺开（可选行列间距）"), _
        Array("M_CropMark.CropMark", "裁切线", "围绕选中对象的包围盒画四个角的裁切线（角线）"), _
        Array("M_Spine.SpineCalc", "书脊计算", "按纸张克重与总页数计算书脊厚度，或反算克重"), _
        Array("M_Align.CenterAlign", "中心对齐", "把选中对象的中心对齐到最底层对象的中心"), _
        Array("M_Distribute.DistributeBySpacing", "按间距分布", "把选中对象按固定间距或平均间距排成一行"), _
        Array("M_DelSegment.DelSegment", "删除线段", "删除选中的直线段对象，或曲线上选中节点之间的线段"))
End Function

' 安装工具栏（带提示，手动执行用）
Public Sub InstallToolbar()
    InstallCore True
End Sub

' 静默安装（安装器一次性调用）
Public Sub InstallToolbarSilent()
    InstallCore False
End Sub

'==========================================================
' 注册插件命令（只写 CorelDRAW 的命令表，不碰任何 UI）。
'
' 为什么每次装工具栏都要注册一遍：按钮上的中文名并不属于按钮自己。
' CorelDRAW 把工具栏骨架持久化在
'   User Workspace\CorelDRAW\_default\DRAWUIConfig.xml
' 里，条目只记「这个按钮绑哪个宏」：
'   <itemData dynamicCommand="TypesetToolkit.M_Impose.Impose" .../>
' 标题要在运行时从命令表里查，而命令表本身不持久化。
'
' 【不要把它挂回启动钩子】启动回调触发时 Corel 的命令栏框架还没
' 初始化完，在里面建/删 CommandBars 会让 X4 直接崩溃。这里只在
' 「安装时」跑一次；按钮标题由安装器写进工作区
' （userCaption / userToolTip，见下面的 DumpItems）。
'==========================================================
Public Sub RegisterCommands()
    Dim items As Variant
    Dim i As Long

    items = CmdList()
    On Error Resume Next
    For i = LBound(items) To UBound(items)
        Err.Clear
        CorelDRAW.AddPluginCommand PRJ_NAME & "." & items(i)(0), items(i)(1), items(i)(2)
        Err.Clear
    Next i
    On Error GoTo 0
End Sub

'==========================================================
' 把按钮清单导出到 %TEMP%\typeset_toolbar_items.txt（UTF-8），
' 一行一个，格式：
'     proc|caption|tooltip
'
' 为什么需要：CorelDRAW 把工具栏骨架写进工作区时只记 dynamicCommand
' （绑哪个宏），不记标题；运行时注册的命令表又不会持久化。所以安装器
' 要在 CorelDRAW 退出之后，拿着这份清单往工作区里补
' userCaption / userToolTip，重启后按钮才有中文名和中文提示。
'
' 由 VBA 导出、而不是让安装器再抄一份，是为了让按钮清单只有一个
' 来源（CmdList），两边不会漂移。
'==========================================================
Public Sub DumpItems()
    Dim items As Variant
    Dim st As Object
    Dim i As Long
    Dim S As String

    items = CmdList()
    For i = LBound(items) To UBound(items)
        S = S & items(i)(0) & "|" & items(i)(1) & "|" & items(i)(2) & vbLf
    Next i

    On Error Resume Next
    Err.Clear
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText S
    st.SaveToFile Environ$("TEMP") & "\typeset_toolbar_items.txt", 2
    st.Close
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub InstallCore(ByVal showMsg As Boolean)
    Dim items As Variant
    Dim cb As Object
    Dim btn As Object
    Dim i As Long

    items = CmdList()

    ' 命令注册（重复注册会报错，忽略）
    RegisterCommands

    ' 先删掉旧的再重建，保证按钮顺序和当前 CmdList 一致
    On Error Resume Next
    Err.Clear
    CorelDRAW.CommandBars(TOOLBAR_NAME).Delete
    Err.Clear
    On Error GoTo 0

    Set cb = Nothing
    On Error Resume Next
    Err.Clear
    Set cb = CorelDRAW.CommandBars.Add(TOOLBAR_NAME)
    If Err.Number <> 0 Then Set cb = Nothing
    Err.Clear
    On Error GoTo 0

    If cb Is Nothing Then
        If showMsg Then MsgBox "创建工具栏失败，请把报错反馈给我。", vbExclamation, "排版工具"
        Exit Sub
    End If

    On Error Resume Next
    cb.Visible = True
    On Error GoTo 0

    On Error Resume Next
    For i = LBound(items) To UBound(items)
        Set btn = Nothing
        Err.Clear
        Set btn = cb.Controls.AddCustomButton(cdrCmdCategoryMacros, PRJ_NAME & "." & items(i)(0))
        If Err.Number = 0 And Not btn Is Nothing Then
            ' 直接写中文标题，免去重启后才刷新的等待
            btn.Caption = items(i)(1)
            btn.TooltipText = items(i)(2)
        End If
        Err.Clear
    Next i
    On Error GoTo 0

    ' 清单落盘，交给安装器在 CorelDRAW 退出后补进工作区
    DumpItems

    If showMsg Then
        MsgBox "工具栏「" & TOOLBAR_NAME & "」安装完成，共 " & (UBound(items) + 1) & " 个功能按钮。", _
               vbInformation, "排版工具"
    End If
End Sub

' 删除工具栏
Public Sub DeleteToolbar()
    On Error Resume Next
    Err.Clear
    CorelDRAW.CommandBars(TOOLBAR_NAME).Delete
    Err.Clear
    On Error GoTo 0
End Sub

' 卸载工具栏
Public Sub UninstallToolbar()
    DeleteToolbar
    MsgBox "工具栏「" & TOOLBAR_NAME & "」已卸载。", vbInformation, "排版工具"
End Sub

'==========================================================
' 工具栏自检：装一遍，再把「真实建出来的按钮」逐项写进日志
' 结果写 %TEMP%\typeset_toolbar.log
'
' 为什么不让外面的冒烟脚本查：X4 不把 CommandBars 交给外部自动化
' 客户端，app.CommandBars(名字) 只会返回 Nothing + err=13 类型不匹配，
' 所以工具栏的验收只能由 VBA 自己做完再落盘。
'==========================================================
Public Sub DiagToolbar()
    Dim fso As Object
    Dim items As Variant
    Dim cb As Object
    Dim btn As Object
    Dim i As Long
    Dim n As Long
    Dim cap As String
    Dim ttl As String
    Dim log As String
    Dim logPath As String
    Dim errNo As Long

    logPath = Environ$("TEMP") & "\typeset_toolbar.log"
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub

    items = CmdList()
    log = "TypesetToolkit toolbar " & Now & vbCrLf
    log = log & "expect       n=" & (UBound(items) + 1) & vbCrLf

    ' 期望值也写进同一份日志：外面只需要读一个文件、一种编码就能比对，
    ' 不用再去解析 src\M_Install.bas 的 UTF-8 源码。
    For i = LBound(items) To UBound(items)
        log = log & "exp " & (i + 1) & " caption=[" & items(i)(1) & "] proc=[" & items(i)(0) & "]" & vbCrLf
    Next i

    ' 装一遍（先删旧的重建，幂等）
    On Error Resume Next
    Err.Clear
    InstallCore False
    errNo = Err.Number
    Err.Clear
    On Error GoTo 0
    log = log & "install      err=" & errNo & vbCrLf

    ' 再从工具栏里取回来逐项核对
    Set cb = Nothing
    On Error Resume Next
    Err.Clear
    Set cb = CorelDRAW.CommandBars(TOOLBAR_NAME)
    errNo = Err.Number
    Err.Clear
    On Error GoTo 0

    log = log & "toolbar      found=" & CStr(Not (cb Is Nothing)) & "  err=" & errNo & vbCrLf

    If Not cb Is Nothing Then
        n = 0
        On Error Resume Next
        n = cb.Controls.Count
        Err.Clear
        On Error GoTo 0
        log = log & "controls     n=" & n & vbCrLf

        For i = 1 To n
            cap = ""
            ttl = ""
            On Error Resume Next
            Err.Clear
            Set btn = Nothing
            Set btn = cb.Controls.Item(i)
            If Err.Number = 0 And Not btn Is Nothing Then
                cap = btn.Caption
                ttl = btn.TooltipText
            End If
            errNo = Err.Number
            Err.Clear
            On Error GoTo 0
            log = log & "btn " & i & " caption=[" & cap & "] tooltip=[" & ttl & "] err=" & errNo & vbCrLf
        Next i
    End If

    log = log & "done" & vbCrLf
    WriteText fso, logPath, log
End Sub

'==========================================================
' 只读探测：报告「工具栏当前是否存在、有几个按钮」，写进
' %TEMP%\typeset_toolbar_state.log。**不创建、不删除任何东西。**
'==========================================================
Public Sub ReportToolbar()
    Dim fso As Object
    Dim cb As Object
    Dim btn As Object
    Dim i As Long
    Dim n As Long
    Dim log As String
    Dim logPath As String
    Dim errNo As Long
    Dim found As Boolean
    Dim cap As String

    logPath = Environ$("TEMP") & "\typeset_toolbar_state.log"
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub

    log = "TypesetToolkit toolbar state " & Now & vbCrLf
    log = log & "name=[" & TOOLBAR_NAME & "]" & vbCrLf

    Set cb = Nothing
    On Error Resume Next
    Err.Clear
    Set cb = CorelDRAW.CommandBars(TOOLBAR_NAME)
    errNo = Err.Number
    Err.Clear
    On Error GoTo 0

    found = Not (cb Is Nothing)
    log = log & "exists=" & CStr(found) & " err=" & errNo & vbCrLf

    n = -1
    If found Then
        On Error Resume Next
        Err.Clear
        n = cb.Controls.Count
        Err.Clear
        On Error GoTo 0
        log = log & "controls=" & n & vbCrLf

        On Error Resume Next
        For i = 1 To n
            cap = ""
            Set btn = Nothing
            Err.Clear
            Set btn = cb.Controls.Item(i)
            If Err.Number = 0 And Not btn Is Nothing Then cap = btn.Caption
            Err.Clear
            log = log & "btn " & i & "=[" & cap & "]" & vbCrLf
        Next
        On Error GoTo 0
    End If

    log = log & "done" & vbCrLf
    WriteText fso, logPath, log
End Sub

Private Sub WriteText(ByVal fso As Object, ByVal p As String, ByVal S As String)
    Dim ts As Object

    On Error Resume Next
    Set ts = fso.CreateTextFile(p, True, True)
    If Not ts Is Nothing Then
        ts.Write S
        ts.Close
    End If
    On Error GoTo 0
End Sub