Attribute VB_Name = "MIde_Mth_Op_Rmv"
Option Compare Database
Option Explicit

Sub MdMthRmv(A As CodeModule, M)
Dim X() As FmCnt: X = MdMthFmCntAyWithTopRmk(A, M)
MsgWh CSub, "Remove method", "Md Mth FmCnt-WithTopRmk", MdNm(A), M, FmCntAyLy(X)
MdRmvFC A, X
End Sub

Private Sub Z_MthRmv()
Const N$ = "ZZModule"
Dim M As CodeModule
Dim M1 As Mth, M2 As Mth
GoSub Crt
Set M = Md(N)
Set M1 = Mth(M, "ZZRmv1")
Set M2 = Mth(M, "ZZRmv2")
MthRmv M1
MthRmv M2
MdEndTrim M
If M.CountOfLines <> 0 Then MsgBox M.CountOfLines
MdDlt M
Exit Sub
Crt:
    CurPjDltMd N
    Set M = CurPjEnsMod(N)
    MdLinesApp M, RplVBar("Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
    Return
End Sub

Sub MthRmv(A As Mth)
MdMthRmv A.Md, A.Nm
End Sub

Private Sub Z()
Z_MthRmv
End Sub
