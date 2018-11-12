Attribute VB_Name = "MDao_Z_Prp_Fld"
Option Compare Database
Option Explicit


Function FdPrpNy(A As DAO.Field) As String()
FdPrpNy = ItrNy(A.Properties)
End Function

Property Get TfPrp(T$, F$, P$)
TfPrp = DbtfPrp(CurDb, T, F, P)
End Property

Property Let TfPrp(T$, F$, P$, V)
DbtfPrp(CurDb, T, F, P) = V
End Property

Private Sub Z_TfPrp()
RfhTmpTbl
Dim P$
P = "Ele"
Ept = 123
GoSub Tst
Exit Sub
Tst:
    TfPrp("Tmp", "F1", P) = Ept
    Act = TfPrp("Tmp", "F1", P)
    C
    Return
End Sub

Function FdDes$(A As DAO.Field)
If PrpHas(A.Properties, C_Des) Then FdDes = A.Properties(C_Des)
End Function

Private Sub Z()
Z_TfPrp
End Sub

Property Get DbtfDes$(A As Database, T, F)
DbtfDes = DbtfPrp(A, T, F, C_Des)
End Property

Property Let DbtfDes(A As Database, T, F, Des$)
DbtfPrp(A, T, F, C_Des) = Des
End Property

