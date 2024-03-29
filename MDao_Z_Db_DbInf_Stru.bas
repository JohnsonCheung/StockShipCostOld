Attribute VB_Name = "MDao_Z_Db_DbInf_Stru"
Option Compare Database
Option Explicit

Function DbInfDtStru(A As Database) As Dt
Dim T$, TT, Dry(), Des$, NRec&, Stru$
For Each TT In DbTny(A)
    T = TT
    Des = DbtDes(A, T)
    Stru = RmvT1(DbtStru(A, T))
    NRec = DbtNRec(A, T)
    PushI Dry, Array(T, NRec, Des, Stru)
Next
Set DbInfDtStru = Dt("Tbl", "Tbl NRec Des", Dry)
End Function

Function DbStru$(A As Database)
DbStru = DbttStru(A, DbTny(A))
End Function

Sub DbStruDmp(A As Database)
D DbStru(A)
End Sub

Function Stru$()
Stru = CurDbStru
End Function

Function StruFld(ParamArray Ap()) As Drs
Dim Dry(), Av(), Ele$, LikFF, LikFld, X
Av = Ap
For Each X In Av
    LinTRstAsg X, Ele, LikFF
    For Each LikFld In SslSy(LikFF)
        PushI Dry, Array(Ele, LikFld)
    Next
Next
Set StruFld = Drs("Ele FldLik", Dry)
End Function

Function TdDefFny(A) As String()
Dim L$, T$
T = LinT1(A)
L = Replace(A, "*", T)
L = Replace(L, "|", " ")
L = RmvT1(L)
TdDefFny = SslSy(L)
End Function

Function TdDefSk(A) As String()
TdDefSk = SslSy(Replace(TakBet(A, "|", "|"), "*", LinT1(A)))
End Function

Function TTStru$(TT)
TTStru = DbttStru(CurDb, TT)
End Function

Sub TTStruDmp(TT)
D TTStru(TT)
End Sub
