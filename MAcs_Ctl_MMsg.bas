Attribute VB_Name = "MAcs_Ctl_MMsg"
Option Explicit
Option Compare Database
Const CMod$ = "MAcs_Ctl_MMsg."

Sub MsgBoxSet(A As Access.TextBox, Msg$)
Dim CrLf$, B$
If A.Value <> "" Then CrLf = vbCrLf
B = LinesKeepLasN(A.Value & CrLf & Now & " " & A, 5)
A.Value = B
DoEvents
End Sub

Private Function ZFrm() As Access.Form
Set ZFrm = Access.Forms("Main")
End Function
Sub Z_ZIsOk()
Debug.Print ZIsOk
End Sub
Private Function ZIsOk() As Boolean
Const CSub$ = CMod & "ZIsOk"
If Not ItrHasNm(Access.CurrentProject.AllForms, "Main") Then
    Msg CSub, "no Main form"
    Exit Function
End If
If Not ItrHasNm(Access.Forms, "Main") Then
    Msg CSub, "Main form is not open"
    Exit Function
End If
If Not ItrHasNm(Access.Forms("Main").Controls, "Msg") Then
    Msg CSub, "no `Msg` textbox in main form"
    Exit Function
End If
ZIsOk = True
End Function
Private Function ZBox() As Access.TextBox
Set ZBox = ZFrm.Controls("Msg")
End Function

Sub MMsgSet(A$)
If ZIsOk Then
    MsgBoxSet ZBox, A
End If
End Sub
Sub MMsgClr()
If ZIsOk Then ZBox.Value = ""
End Sub
Sub MMsgRun(QryNm)
MMsgSet "Running query: " & QryNm
End Sub
