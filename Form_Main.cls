VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub CmdEditPrm_Click()
Me.UOMFn.Enabled = True
Me.UomPth.Enabled = True
Me.ZHT1Fn.Enabled = True
Me.ZHT1Pth.Enabled = True
Me.MB52Pth.Enabled = True
End Sub

Private Sub CmdOpnUomPth_Click()
PthBrw Me.UomPth.Value
End Sub
Private Sub CmdOpnMB52Pth_Click()
PthBrw Me.MB52Pth.Value
End Sub

Sub MsgSet(A$)
MsgBoxSet Msg, A
End Sub

Sub MsgClr()
Me.Msg.Value = ""
DoEvents
End Sub

Private Sub CmdOpnZHT1Pth_Click()
PthBrw Me.ZHT1Pth.Value
End Sub

Private Sub CmdQuit_Click()
Quit
End Sub

Private Sub CmdRunReport_Click()
If Not Vdt Then Exit Sub
Rpt
End Sub

Private Function Vdt() As Boolean
MsgClr
DoCmd.RunCommand acCmdSaveRecord
If PthEr(UomPth) Then Exit Function
If FnEr(UOMFn, UomPth) Then Exit Function
If PthEr(MB52Pth) Then Exit Function
If FnEr(MB52Fn, MB52Pth) Then Exit Function
If MB52FnFmtEr(MB52Fn) Then Exit Function
DoCmd.RunCommand acCmdSaveRecord
Vdt = True
End Function

Private Function MB52FnFmtEr(A As TextBox) As Boolean
If Left(A.Value, 5) <> "MB52 " Then
    Me.Msg.Value = "MB52 file name must begin with [MB52 ]"
    A.SetFocus
    MB52FnFmtEr = True
    Exit Function
End If
Dim B$
B = Mid(A.Value, 6, 10)
If Not YYYYMMDD_IsVdt(B) Then
    Me.Msg.Value = "MB52 file name must begin with [MB52 YYYY-MM-DD]"
    A.SetFocus
    MB52FnFmtEr = True
End If
End Function
Private Function FnEr(A As TextBox, P As TextBox) As Boolean
If Dir(PthEnsSfx(P.Value) & A.Value) = "" Then
    A.Enabled = True
    A.SetFocus
    Me.Msg.Value = "File not found"
    FnEr = True
End If
End Function
Private Function PthEr(A As TextBox) As Boolean
If Dir(A.Value, vbDirectory) = "" Then
    A.Enabled = True
    A.SetFocus
    Me.Msg.Value = "Path not found"
    PthEr = True
End If
End Function

Private Sub CmOpnMB52Pth_Click()
PthBrw Me.MB52Pth.Value
End Sub

Private Sub Form_Load()
Me.Msg.Value = ""
If IsDev Then
    Me.MB52Fn.Value = "MB52 2018-07-30.xls"
Else
    Me.MB52Fn.Value = "MB52 " & Format(Now, "YYYY-MM-DD") & ".xls"
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
TblPrm_Cpy_Fm_C_or_N IsDev
Me.RecordSource = "Prm"
End Sub
