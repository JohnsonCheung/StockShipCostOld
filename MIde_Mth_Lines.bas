Attribute VB_Name = "MIde_Mth_Lines"
Option Compare Database
Option Explicit

Function CurMthBdyLines$()
CurMthBdyLines = MdMthBdyLines(CurMd, CurMthNm$)
End Function

Function MdMthBdyLines$(A As CodeModule, MthNm)
MdMthBdyLines = SrcMthBdyLines(MdBdyLy(A), MthNm)
End Function

Function MdMthBdyLy(A As CodeModule, MthNm) As String()
MdMthBdyLy = SrcMthBdyLy(MdSrc(A), MthNm)
End Function

Function MdMthLines$(A As CodeModule, MthNm, Optional WithTopRmk As Boolean)
MdMthLines = SrcMthLines(MdSrc(A), MthNm, WithTopRmk)
End Function

Function MdMthLnoLines$(A As CodeModule, MthLno&)
MdMthLnoLines = A.Lines(MthLno, MdMthLinCnt(A, MthLno))
End Function

Function MthBdyLy(A As CodeModule, MthNm) As String()
MthBdyLy = SrcMthBdyLy(MdBdyLy(A), MthNm)
End Function

Function MthDDNmLines$(MthDDNm$)
MthDDNmLines = MthLines(DDNmMth(MthDDNm))
End Function

Function MthEndLin$(MthLin$)
Dim A$
A = LinMthKd(MthLin): If A = "" Then Stop
MthEndLin = "End " & A
End Function

Function MthLinCnt%(A As Mth)
MthLinCnt = FmCntAyLinCnt(MthFC(A))
End Function

Function MthLines$(A As Mth, Optional WithTopRmk As Boolean)
MthLines = SrcMthLines(MdBdyLy(A.Md), A.Nm, WithTopRmk)
End Function

Function MthLinesWithTopRmk$(A As Mth)
MthLinesWithTopRmk = MthLines(A, WithTopRmk:=True)
End Function

'aaa
Private Property Get XX1()

End Property

'BB
Private Property Let XX1(V)

End Property

Private Sub Z()
Z_MthDDNmLines
Z_MthLines
End Sub

Private Sub Z_MthDDNmLines()
GoTo ZZ
ZZ:
Debug.Print MthDDNmLines("QIde.MIde_Mth_Lines.ZZ_MthDDNmLines")
End Sub

'aa
Private Sub Z_MthLines()
Debug.Print MthLines(Mth(CurMd, "XX1"), WithTopRmk:=True)
End Sub
