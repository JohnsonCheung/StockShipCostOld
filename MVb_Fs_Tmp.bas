Attribute VB_Name = "MVb_Fs_Tmp"
Option Compare Database
Option Explicit

Function TmpCmd$(Optional Fdr$, Optional Fnn$)
TmpCmd = TmpFfn(".cmd", Fdr, Fnn)
End Function

Function TmpFb$(Optional Fdr$, Optional Fnn$)
TmpFb = TmpFfn(".accdb", Fdr, Fnn)
End Function

Function TmpFcsv$(Optional Fdr$, Optional Fnn$)
TmpFcsv = TmpFfn(".csv", Fdr, Fnn)
End Function

Function TmpFfn$(Ext$, Optional Fdr$, Optional Fnn0$)
Dim Fnn$
TmpPth
Fnn = IIf(Fnn0 = "", TmpNm, Fnn0)
TmpFfn = TmpFdrPth(Fdr) & Fnn & Ext
End Function

Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

Function TmpFx$(Optional Fdr$, Optional Fnn$)
TmpFx = TmpFfn(".xlsx", Fdr, Fnn)
End Function

Function TmpFxm$(Optional Fdr$, Optional Fnn0$)
TmpFxm = TmpFfn(".xlsm", Fdr, Fnn0)
End Function
Function TmpRoot$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpRoot = X
End Function

Function TmpHom$()
Static X$
If X = "" Then X = PthEns(TmpRoot & "App")
TmpHom = X
End Function

Sub TmpHomBrw()
PthBrw TmpHom
End Sub

Function TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function

Function TmpFdrPth$(Fdr$)
TmpFdrPth = PthEns(TmpHom & Fdr & "\")
End Function

Function TmpPth$()
TmpPth = PthEns(TmpHom & TmpNm & "\")
End Function

Sub TmpPthBrw()
PthBrw TmpPth
End Sub

Function TmpPthFix$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthFix = X
End Function

Function TmpPthHom$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthHom = X
End Function
