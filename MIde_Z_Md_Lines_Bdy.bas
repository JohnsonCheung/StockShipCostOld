Attribute VB_Name = "MIde_Z_Md_Lines_Bdy"
Option Compare Database
Option Explicit
Function MdBdyLines$(A As CodeModule)
If MdHasNoMth(A) Then Exit Function
MdBdyLines = A.Lines(MdBdyFmLno(A), A.CountOfLines)
End Function

Function MdBdyFmLno%(A As CodeModule)
MdBdyFmLno = MdDclLinCnt(A) + 1
End Function

Function MdBdyFmCnt(A As CodeModule) As FmCnt
Dim Lno&
Dim Cnt&
Lno = MdBdyFmLno(A)
MdBdyFmCnt = FmCnt(Lno, Cnt)
End Function

Function MdBdyLy(A As CodeModule) As String()
MdBdyLy = SplitCrLf(MdBdyLines(A))
End Function
