Attribute VB_Name = "MApp_Apn"
Option Compare Database
Option Explicit
Function AppPth$()
AppPth = PthEns(TmpHom & Apn)
End Function
Private Sub Z_Apn()
OpnCurDb SampleFb_ShpRate
Debug.Print Apn
End Sub
Function Apn$()
Static Y$
If Y = "" Then Y = SqlV("Select Apn from [Apn]")
Apn = Y
End Function
Function ApnAcs(A) As Access.Application
'AcsOpn Acs, ApnWFb(A)
'Set ApnAcs = Acs
End Function

Sub ApnBrwWDb(A)
Dim Fb$
Fb = ApnWFb(A)
'AcsOpn Acs, Fb
'AcsVis Acs
End Sub

Function ApnWAcs(A)
Dim O As Access.Application
'AcsOpn O, ApnWFb(A)
Set ApnWAcs = O
End Function

Function ApnWDb(A) As Database
Static X As Boolean, Y As Database
If Not X Then
    X = True
    FbEns ApnWFb(A)
    Set Y = FbDb(ApnWFb(A))
End If
If Not DbIsOk(Y) Then Set Y = FbDb(ApnWFb(A))
Set ApnWDb = Y
End Function

Function ApnWFb$(A)
ApnWFb = ApnWPth(A) & "Wrk.accdb"
End Function

Function ApnWPth$(A)
Dim P$
P = TmpHom & A & "\"
PthEns P
ApnWPth = P
End Function

Function PgmObjPth$()
PgmObjPth = PthEns(CurDbPth & "PgmObj\")
End Function

Private Function PgmPth$()
PgmPth = FfnPth(Excel.Application.Vbe.ActiveVBProject.FileName)
End Function


Private Sub Z()
Z_Apn
End Sub
