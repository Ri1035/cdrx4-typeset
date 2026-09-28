Attribute VB_Name = "M_Spine"
Option Explicit

'==========================================================
' 书脊计算
'
' 参考实现来自本机「可读」的 Draw\GMS\书脊计算.gms（UserForm1）：
'     书脊厚度 = 0.135 * (纸张克重 / 100) * (总页数 / 2)
'     反算克重 = 书脊厚度 / (总页数 / 2) * 100 / 0.135
'
' 0.135 是胶版纸的经验系数（mm / (g/m²) / 张）。原版是 UserForm，
' X4 的 .bas 导不进窗体，所以降级成 InputBox 顺序问答（流程等价）。
'==========================================================

Public Const SPINE_FACTOR As Double = 0.135

' 已知克重与总页数 -> 书脊厚度（mm）
Public Function SpineThicknessMm(ByVal weightGsm As Double, ByVal pages As Double) As Double
    SpineThicknessMm = SPINE_FACTOR * (weightGsm / 100#) * (pages / 2#)
End Function

' 已知书脊厚度与总页数 -> 纸张克重（g/m²）
Public Function WeightFromSpine(ByVal spineMm As Double, ByVal pages As Double) As Double
    If pages <= 0 Then
        WeightFromSpine = 0#
        Exit Function
    End If
    WeightFromSpine = spineMm / (pages / 2#) * 100# / SPINE_FACTOR
End Function

' 按钮入口
Public Sub SpineCalc()
    Dim mode As String
    Dim s1 As String
    Dim s2 As String
    Dim w As Double
    Dim p As Double
    Dim r As Double

    mode = InputBox("书脊计算" & vbCrLf & vbCrLf & _
                    "1 = 已知 纸张克重 + 总页数  →  算书脊厚度" & vbCrLf & _
                    "2 = 已知 书脊厚度 + 总页数  →  反算纸张克重" & vbCrLf & vbCrLf & _
                    "请输入 1 或 2：", "书脊计算", "1")
    mode = Trim$(mode)
    If Len(mode) = 0 Then Exit Sub

    If mode = "2" Then
        s1 = InputBox("请输入书脊厚度（mm）：", "书脊计算", "10")
        If Len(Trim$(s1)) = 0 Then Exit Sub
        s2 = InputBox("请输入总页数：", "书脊计算", "200")
        If Len(Trim$(s2)) = 0 Then Exit Sub

        r = WeightFromSpine(Val(s1), Val(s2))
        MsgBox "书脊厚度 " & Format$(Val(s1), "0.###") & " mm" & vbCrLf & _
               "总页数 " & Format$(Val(s2), "0") & vbCrLf & vbCrLf & _
               "推算纸张克重 ≈ " & Format$(r, "0.#") & " g/m²", _
               vbInformation, "书脊计算"
    Else
        s1 = InputBox("请输入纸张克重（g/m²，例如 80）：", "书脊计算", "80")
        If Len(Trim$(s1)) = 0 Then Exit Sub
        s2 = InputBox("请输入总页数：", "书脊计算", "200")
        If Len(Trim$(s2)) = 0 Then Exit Sub

        r = SpineThicknessMm(Val(s1), Val(s2))
        MsgBox "纸张克重 " & Format$(Val(s1), "0.#") & " g/m²" & vbCrLf & _
               "总页数 " & Format$(Val(s2), "0") & vbCrLf & vbCrLf & _
               "书脊厚度 ≈ " & Format$(r, "0.###") & " mm", _
               vbInformation, "书脊计算"
    End If
End Sub