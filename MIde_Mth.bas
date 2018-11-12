Attribute VB_Name = "MIde_Mth"
Option Compare Database
Option Explicit

Function MdEnsMth(A As CodeModule, MthNm$, NewMthLines$)
Dim OldMthLines$: OldMthLines = MdMthBdyLines(A, MthNm)
If OldMthLines = NewMthLines Then
    Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is same", MthNm, MdNm(A))
End If
MdMthRmv A, MthNm
MdLinesApp A, NewMthLines
Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is replaced <=========", MthNm, MdNm(A))
End Function


Function MdMthPfxAy(A As CodeModule) As String()
Dim N
For Each N In AyNz(MdMthNy(A))
    PushNoDup MdMthPfxAy, MthPfx(N)
Next
End Function

Function MdHasMth(A As CodeModule, MthNm$, Optional WhMdy$, Optional WhKd$) As Boolean
MdHasMth = SrcHasMth(MdBdyLy(A), MthNm, WhMdy, WhKd)
End Function

Function MdHasTstSub(A As CodeModule) As Boolean
Dim I
For Each I In MdLy(A)
    If I = "Friend Sub Z()" Then MdHasTstSub = True: Exit Function
    If I = "Sub Z()" Then MdHasTstSub = True: Exit Function
Next
End Function



Function MdMthFmCntAy(A As CodeModule, MthNm$) As FmCnt()
MdMthFmCntAy = SrcMthFmCntAy(MdSrc(A), MthNm)
End Function

Private Sub Z_MdMthFmCntAy()
Dim A() As FmCnt: A = MdMthFmCntAy(Md("Md_"), "XX")
Dim J%
For J = 0 To UB(A)
    FmCntDmp A(J)
Next
End Sub

Function MdMthAy(A As CodeModule) As Mth()
Dim N
For Each N In AyNz(MdMthNy(A))
    PushObj MdMthAy, Mth(A, N)
Next
End Function



Function MdMthKeyLinesDic1(A As CodeModule) As Dictionary
'To be delete
'Dim Pfx$: Pfx = MdPjNm(A) & "." & MdNm(A) & "."
'Set MdMthKeyLinesDic = DicAddKeyPfx(SrcMthKeyLinesDic(MdSrc(A)), Pfx)
End Function

Function MdMthKy(A As CodeModule) As String()
MdMthKy = AyAddPfx(SrcMthKy(MdSrc(A)), MdDNm(A) & ".")
End Function




Private Sub Z()
Z_MdMthFmCntAy
End Sub
