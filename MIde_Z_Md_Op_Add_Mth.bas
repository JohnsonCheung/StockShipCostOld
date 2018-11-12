Attribute VB_Name = "MIde_Z_Md_Op_Add_Mth"
Option Compare Database
Option Explicit
Sub MdAddFun(A As CodeModule, Nm$, Lines)
MdAdd1 A, Nm, Lines, IsFun:=True
End Sub

Private Sub MdAdd1(A As CodeModule, Nm$, Lines, IsFun As Boolean)
Dim L$
    Dim B$
    B = IIf(IsFun, "Function", "Sub")
    L = FmtQQ("? ?()|?|End ?", B, Nm, Lines, B)
MdLinesApp A, L
MthGo Mth(A, Nm)
End Sub

Sub MdAddSub(A As CodeModule, Nm$, Lines)
MdAdd1 A, Nm, Lines, IsFun:=False
End Sub
