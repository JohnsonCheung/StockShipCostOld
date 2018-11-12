Attribute VB_Name = "MApp___Fun"
Option Compare Database
Option Explicit

Function AppDtaHom$()
AppDtaHom = PthUp(TmpHom)
End Function

Function AppDtaPth$()
AppDtaPth = PthEns(AppDtaHom & Apn & "\")
End Function


Function AppFbAy() As String()
Push AppFbAy, AppJJFb
Push AppFbAy, AppStkShpCstFb
Push AppFbAy, AppStkShpRateFb
Push AppFbAy, AppTaxExpCmpFb
Push AppFbAy, AppTaxRateAlertFb
End Function

Function AppMdNy() As String()
AppMdNy = ItrNy(CodeProject.AllModules)
End Function

Function AppPushAppFcmd$()
AppPushAppFcmd = WPth & "PushApp.Cmd"
End Function

Function AppRoot$()
Stop '
End Function

Function AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
SpecEnsTbl

DbLnkCcm CurDb, IsDev
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Function

Sub BrwDtaFb()
'Acs.OpenCurrentDatabase DtaFb
End Sub

Sub CrtDtaFb()
If IsDev Then Exit Sub
If FfnIsExist(DtaFb) Then Exit Sub
FbCrt DtaFb
Dim Src, Tar$, TarFb$
TarFb = DtaFb
Stop
'For Each Src In CcmTny
    Tar = Mid(Src, 2)
    Application.DoCmd.CopyObject TarFb, Tar, acTable, Src
    Debug.Print MsgLin("CrtDtaFb: Cpy [Src] to [Tar]", Src, Tar)
'Next
End Sub

Sub Doc()
'#BNmMIS is Method-B-Nm-of-Missing.
'           Missing means the method is found in FmPj, but not ToPj
'#FmDicB is MthDic-of-MthBNm-zz-MthLines.   It comes from FmPj
'#ToDicA is MthDic-of-MthANm-zz-MthLinesAy. It comes from ToPj
'#ToDicAB is ToDicA and FmDicB
'#ANm is method-a-name, NNN or NNN:YYY
'        If the method is Sub|Fun, just MthNm
'        If the method is Prp    ,      MthNm:MthTy
'        It is from ToPj (#ToA)
'        One MthANm will have one or more MthLines
'#BNm is method-b-name, MMM.NNN or MMM.NNN:YYY
'        MdNm.MthNm[:MthTy]
'        It is from FmPj (#BFm)
'        One MthBNm will have only one MthLines
'#Missing is for each MthBNm found in FmDicB, but its MthNm is not found in any-method-name-in-ToDicA
'#Dif is for each MthBNm found in FmDicB and also its MthANm is found in ToDicA
'        and the MthB's MthLines is dif and any of the MthA's MthLines
'       (Note.MthANm will have one or more MthLines (due to in differmodule))
End Sub

Function DtaDb() As Database
Set DtaDb = DAO.DBEngine.OpenDatabase(DtaFb)
End Function

Function DtaFb$()
DtaFb = AppHom & DtaFn
End Function

Function DtaFn$()
DtaFn = Apn & "_Data.accdb"
End Function

Function IsDev() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then
    X = True
    Y = Not PthIsExist(ProdPth)
End If
IsDev = Y
End Function

Function IsProd() As Boolean
IsProd = Not IsDev
End Function

Private Sub Z_AppFbAy()
Dim F
For Each F In AppFbAy
If Not IsFfnExist(F) Then Stop
Next
End Sub
