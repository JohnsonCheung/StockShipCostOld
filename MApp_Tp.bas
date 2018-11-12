Attribute VB_Name = "MApp_Tp"
Option Compare Database
Const CMod$ = "MApp_Tp."

Option Explicit
Sub TpAddWc()
Dim Wb As Workbook
Set Wb = TpWb
WbAddWc Wb, WFb, "@Main"
WbAddWc Wb, WFb, "@Rate"
WbAddWc Wb, WFb, "@Sku"
WbAddWc Wb, WFb, "@Repack1"
WbAddWc Wb, WFb, "@Repack2"
WbAddWc Wb, WFb, "@Repack3"
WbAddWc Wb, WFb, "@Repack4"
WbAddWc Wb, WFb, "@Repack5"
WbAddWc Wb, WFb, "@Repack6"
Wb.Close True
End Sub

Sub TpExp()
AttExp "Tp", TpFx
End Sub

Function TpFnn$()
TpFnn = Apn & "(Template)"
End Function

Sub TpGenFx(TpFx$, OupFx$, Fb$, ParamArray WbFmtrAp())
Dim Av()
Av = WbFmtrAp
FfnCpy TpFx, OupFx
WbFmt FxRfh(OupFx, Fb), Av
End Sub

Function TpIdxWs() As Worksheet
Set TpIdxWs = WbWsCd(TpWb, "WsIdx")
End Function

Sub TpImp()
Const CSub$ = CMod & "TpImp"
Dim A$
A = TpFx
If Not FfnIsExist(A) Then
    If True Then
        FunMsgNyApDmp CSub, "Tp not exist, no Import", "TpFx", A
    End If
End If
If AttIsOld("Tp", A) Then AttImp "Tp", A '<== Import
End Sub

Function TpMainFbtStr$()
Dim Wb As Workbook, Qt As QueryTable
Set Wb = TpWb
Set Qt = WbMainQt(Wb)
TpMainFbtStr = QtFbtStr(Qt)
WbQuit Wb
End Function

Function TpMainLo() As ListObject
Set TpMainLo = WbMainLo(TpWb)
End Function

Function TpMainQt() As QueryTable
Set TpMainQt = WbMainQt(TpWb)
End Function
'===============================================
Function TpFx$()
TpFx = TpPth & Apn & "(Template).xlsx"
End Function

Function TpFxm$()
TpFxm = TpPth & Apn & "(Template).xlsm"
End Function
Sub TpOpn()
FxOpn TpFx
End Sub

Function TpWsCdNy() As String()
TpWsCdNy = FxWsCdNy(TpFx)
End Function


'==============================================
Sub TpMinLo()
Dim O As Workbook
Set O = TpWb
WbMinLo O
O.Save
WbVis O
End Sub

Function TpPth$()
TpPth = PthEns(CurDbPth & "Template\")
End Function

Sub TpRfh()
WbVis WbRfh(TpWb)
End Sub

Sub TpRfhWc()
FxRfh TpFx, WFb
End Sub

Function TpWb() As Workbook
Set TpWb = FxWb(TpFx)
End Function

Function TpWcSy() As String()
Dim W As Workbook, X As Excel.Application
Set X = New Excel.Application
Set W = X.Workbooks.Open(TpFx)
TpWcSy = WbWcSy_zOle(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Function

Sub TpWrtFfn(Ffn$)
AttExp "Tp", Ffn
End Sub

