Attribute VB_Name = "MDao_Lnk_LnkColVbl_Import"
Option Explicit
Option Compare Database

Sub DbtImpMap(A As Database, T, LnkColVbl$, Optional WhBExpr$)
If FstChr(T) <> ">" Then
    Debug.Print "FstChr of T must be >"
    Stop
End If
'Assume [>?] T exist
'Create [#I?] T
DbtDrp A, "#I" & Mid(T, 2)
Q = ZImpSql(LnkColVbl, T, WhBExpr): A.Execute Q
End Sub

Private Function ZImpSql$(A$, T, Optional WhBExpr$)
Dim Ay() As LnkCol
Ay = LnkColVbl_LnkColAy(A)
ZImpSql = ZImpSql1(Ay, T, WhBExpr)
End Function

Private Function ZImpSql1$(A() As LnkCol, T, Optional WhBExpr$)
If FstChr(T) <> ">" Then
    Debug.Print "T must have first char = '>'"
    Stop
End If
Dim Ny$(), ExtNy$(), J%, O$(), S$, N$(), E$()
Ny = LnkColAy_Ny(A)
ExtNy = LnkColAy_ExtNy(A)
N = AyAlignL(Ny)
E = AyAlignL(AyQuoteSqBktIfNeed(ExtNy))
Erase O
For J = 0 To UB(Ny)
    If ExtNy(J) = Ny(J) Then
        Push O, FmtQQ("     ?    ?", Space(Len(E(J))), N(J))
    Else
        Push O, FmtQQ("     ? As ?", E(J), N(J))
    End If
Next
S = Join(O, "," & vbCrLf)
ZImpSql1 = FmtQQ("Select |?| Into [#I?]| From [?] |?", S, RmvFstChr(T), T, X.Wh(WhBExpr))
End Function

