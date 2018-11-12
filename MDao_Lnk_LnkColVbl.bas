Attribute VB_Name = "MDao_Lnk_LnkColVbl"
Option Explicit
Option Compare Database


Function ColLnk_ImpSql$(A$(), Fm)
'data ColLnk = F T E
Dim Into$, Ny$(), Ey$()
If FstChr(Fm) <> ">" Then Stop
Into = "#I" & Mid(Fm, 2)
Ny = AyTakT1(A)
Ey = AyMapSy(A, "RmvTT")
'ColLnk_ImpSql = SelNyEyIntoFmSql$(Fm, Into, Ny, Ey)
Stop '
End Function

Function LnkColVbl_LnkColAy(A) As LnkCol()
Dim L
For Each L In AyNz(SplitVBar(A))
    PushObj LnkColVbl_LnkColAy, LinLnkCol(L)
Next
End Function

Function LnkColVbl_Ly(A$) As String()
Dim A1$(), A2$(), Ay() As LnkCol
Ay = LnkColVbl_LnkColAy(A)
A1 = LnkColAy_Ny(Ay)
A2 = AyAlignL(AyQuoteSqBkt(LnkColAy_ExtNy(Ay)))
Dim J%, O$()
For J = 0 To UB(A1)
    Push O, A2(J) & "  " & A1(J)
Next
LnkColVbl_Ly = O
End Function


Function LnkColAy_ExtNy(A() As LnkCol) As String()
LnkColAy_ExtNy = OyPrpSy(A, "Extnm")
End Function

Function LnkColAy_Ny(A() As LnkCol) As String()
LnkColAy_Ny = OyPrpSy(A, "Nm")
End Function

Function LinLnkCol(A) As LnkCol
Dim Nm$, TyStr$, ExtNm$, Ty As DAO.DataTypeEnum
Lin2TRstAsg A, Nm, TyStr, ExtNm
ExtNm = RmvSqBkt(Trim(ExtNm))
Ty = DaoTy(TyStr)
Set LinLnkCol = LnkCol(Nm, Ty, IIf(ExtNm = "", Nm, ExtNm))
End Function


Private Sub Z_LinLnkCol()
Dim A$, Act As LnkCol, Exp As LnkCol
A = "AA Txt XX"
Exp = LnkCol("AA", dbText, "AA")
GoSub Tst
Exit Sub
Tst:
Act = LinLnkCol(A)
Debug.Assert LnkColIsEq(Act, Exp)
Return
End Sub


