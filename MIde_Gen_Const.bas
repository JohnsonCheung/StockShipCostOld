Attribute VB_Name = "MIde_Gen_Const"
Option Compare Database
Option Explicit
Private A_Nm$
Private B_Pj As VBProject
Private B_Md As CodeModule
Private B_Ft$
Private B_ValFmSrc$

Sub ConstEdt(ConstFunNm$)
ZSetAB ConstFunNm
StrWrt B_ValFmSrc, B_Ft, True
FtBrw B_Ft
End Sub

Sub ConstUpdSrc(ConstFunNm$, Optional IsPub As Boolean)
ZSetAB ConstFunNm
MdMthRmv B_Md, A_Nm
B_Md.InsertLines B_Md.CountOfLines + 1, ConstValMthLines(FtLines(B_Ft), A_Nm, IsPub)
End Sub

Private Sub ZSetAB(ConstFunNm$)
A_Nm = ConstFunNm
Dim Pth$
Pth = TmpHom & "GenConst\": PthEns Pth
Set B_Pj = CurPj
Set B_Md = CurMd
B_Ft = Pth & A_Nm & ".txt"
B_ValFmSrc = MthLinesConstVal(MdMthLines(B_Md, A_Nm))
End Sub

Private Sub Z()
Z_ConstEdt
Z_ConstUpdSrc
End Sub

Private Sub Z_ConstEdt()
ConstEdt "ZZ_A"
End Sub

Private Sub Z_ConstUpdSrc()
ConstUpdSrc "ZZ_A"
End Sub
