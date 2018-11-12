Attribute VB_Name = "MIde__Cur_CdPne_Md_Mth"
Option Compare Database
Option Explicit
Private Sub Z()
Z_CurMd
End Sub
Function CurMdNm$()
CurMdNm = CurCmp.Name
End Function

Function MdCurMthNm$(A As CodeModule)
Dim R1&, R2&, C1&, C2&
A.CodePane.GetSelection R1, C1, R2, C2
Dim K As vbext_ProcKind
MdCurMthNm = A.ProcOfLine(R1, K)
End Function

Function CurMd() As CodeModule
Set CurMd = CurCdPne.CodeModule
End Function

Private Sub Z_CurMd()
Ass CurMd.Parent.Name = "Cur_d"
End Sub

Function CurMdDNm$()
CurMdDNm = MdDNm(CurMd)
End Function

Function CurMdWin() As VBIDE.Window
Dim A As CodePane
Set A = CurCdPne
If IsNothing(A) Then Exit Function
Set CurMdWin = A.Window
End Function

Function CurMth() As Mth
Dim M As CodeModule
    Set M = CurMd
Set CurMth = Mth(M, MdCurMthNm(M))
End Function

Function CurMthNm$()
CurMthNm = CurMth.Nm
End Function
Function CurCdPne() As VBIDE.CodePane
Set CurCdPne = CurVbe.ActiveCodePane
End Function
