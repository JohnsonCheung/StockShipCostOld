Attribute VB_Name = "MIde_Z_Md_Lines_Add"
Option Compare Database
Option Explicit
Const CMod$ = ""
Sub MdAppDclLin(A As CodeModule, DclLines$)
A.InsertLines A.CountOfDeclarationLines + 1, DclLines
Debug.Print FmtQQ("MdAppDclLin: Module(?) a DclLin is inserted", MdNm(A))
End Sub

Sub MdLinesApp(A As CodeModule, Lines$)
Const CSub$ = CMod & "MdLinesApp"
If Lines = "" Then Exit Sub
Dim Bef&, Aft&, Exp&, Cnt&
Bef = A.CountOfLines
A.InsertLines A.CountOfLines + 1, Lines '<=====
Aft = A.CountOfLines
Cnt = LinesLinCnt(Lines)
Exp = Bef + Cnt
If Exp <> Aft Then
'    Er CSub, "After copy line count are inconsistents, where [Md], [LinCnt-Bef-Cpy], [LinCnt-of-lines], [Exp-LinCnt-Aft-Cpy], [Act-LinCnt-Aft-Cpy], [Lines]", _
        MdNm(A), Bef, Cnt, Exp, Aft, Lines
End If
End Sub

Sub MdAppLy(A As CodeModule, Ly$())
MdLinesApp A, JnCrLf(Ly)
End Sub

