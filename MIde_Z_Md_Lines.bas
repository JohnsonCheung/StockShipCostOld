Attribute VB_Name = "MIde_Z_Md_Lines"
Option Compare Database
Option Explicit
Function MdLines$(A As CodeModule)
If A.CountOfLines = 0 Then Exit Function
MdLines = A.Lines(1, A.CountOfLines)
End Function

Function MdLinesByFmCnt$(A As CodeModule, FmCnt As FmCnt)
With FmCnt
    If .Cnt <= 0 Then Exit Function
    MdLinesByFmCnt = A.Lines(.FmLno, .Cnt)
End With
End Function
Function MdLy(A As CodeModule) As String()
MdLy = SplitCrLf(MdLines(A))
End Function

Function MdFTLines$(A As CodeModule, X As FTNo)
Dim Cnt%: Cnt = FTNoLinCnt(X)
If Cnt = 0 Then Exit Function
MdFTLines = A.Lines(X.FmNo, Cnt)
End Function

Function MdFTLy(A As CodeModule, X As FTNo) As String()
MdFTLy = SplitCrLf(MdFTLines(A, X))
End Function


Function MdPatnLy(A As CodeModule, Patn$) As String()
Dim Ix&(): Ix = AyWhPatnIx(MdLy(A), Patn)
Dim O$(), I, Md As CodeModule
Dim N$: N = MdNm(A)
If Sz(Ix) = 0 Then Exit Function
For Each I In Ix
   Push O, FmtQQ("MdGoLno ""?"",??' ?", N, I + 1, vbTab, A.Lines(I + 1, 1))
Next
MdPatnLy = O
End Function
