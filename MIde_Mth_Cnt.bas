Attribute VB_Name = "MIde_Mth_Cnt"
Option Compare Database
Option Explicit

Function MdMthCnt%(A As CodeModule, Optional B As WhMth)
MdMthCnt = SrcMthCnt(MdSrc(A))
End Function

Function MdMthLinCnt%(A As CodeModule, MthLno&)
Dim Kd$, Lin$, EndLin$, J%
Lin = A.Lines(MthLno, 1)
Kd = LinMthKd(Lin)
If Kd = "" Then Stop
EndLin = "End " & Kd
If HasSfx(Lin, EndLin) Then
    MdMthLinCnt = 1
    Exit Function
End If
For J = MthLno + 1 To A.CountOfLines
    If HasSfx(A.Lines(J, 1), EndLin) Then
        MdMthLinCnt = J - MthLno + 1
        Exit Function
    End If
Next
Stop
End Function

Function MdMthPfxCnt%(A As CodeModule)
MdMthPfxCnt = Sz(MdMthPfxAy(A))
End Function

Function MdNMth%(A As CodeModule)
MdNMth = SrcNMth(MdSrc(A))
End Function

Function MdNPubMth%(A As CodeModule)
MdNPubMth = SrcNMth(MdSrc(A), WhMth("Pub"))
End Function

Function PjNPubMth%(A As VBProject)
Dim O%, C As VBComponent
For Each C In A.VBComponents
    O = O + MdNPubMth(C.CodeModule)
Next
PjNPubMth = O
End Function

Function SrcMthCnt%(A$(), Optional B As WhMth)
SrcMthCnt = Sz(SrcMthIxAy(A, B))
End Function

Private Function SrcNMth%(A$(), Optional B As WhMth)
SrcNMth = SrcMthCnt(A, B)
End Function
