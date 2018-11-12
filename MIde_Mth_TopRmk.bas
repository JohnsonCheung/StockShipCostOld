Attribute VB_Name = "MIde_Mth_TopRmk"
Option Compare Database
Option Explicit
Private Sub Z_MthFmCntAyWithTopRmk()
Dim A As CodeModule, M, Ept() As FmCnt, Act() As FmCnt

Set A = Md("IdeMthFmCnt")
M = "Z_MthFmCntAyWithTopRmk "
PushObj Ept, FmCnt(2, 11)
GoSub Tst

Exit Sub
Tst:
    Act = MdMthFmCntAyWithTopRmk(A, M)
    If Not FmCntAyIsEq(Act, Ept) Then Stop
    Return
End Sub

Function MdMthFmCntAyWithTopRmk(A As CodeModule, M) As FmCnt()
MdMthFmCntAyWithTopRmk = SrcMthFmCntAyWithTopRmk(MdSrc(A), M)
End Function

Function SrcMthFmCntAyWithTopRmk(A$(), MthNm) As FmCnt()
Dim FmIx&, ToIx&, IFm, Fm&
For Each IFm In AyNz(SrcMthNmIxAy(A, MthNm))
    Fm = IFm
    FmIx = SrcMthIxTopRmkFm(A, Fm)
    ToIx = SrcMthIxTo(A, Fm)
    PushObj SrcMthFmCntAyWithTopRmk, FmCnt(FmIx + 1, ToIx - FmIx + 1)
Next
End Function
Function SrcMthIxTopRmk$(A$(), MthIx&)
Dim O$(), J&, L$
Dim Fm&: Fm = SrcMthIxTopRmkFm(A, MthIx)
For J = Fm To MthIx - 1
    L = A(J)
    If FstChr(L) = "'" Then
        If L <> "'" Then
            PushI O, L
        End If
    End If
Next
SrcMthIxTopRmk = Join(O, vbCrLf)
End Function


Function SrcMthIxTopRmkFm&(A$(), MthIx&)
Dim M1&
    Dim J&
    For J = MthIx - 1 To 0 Step -1
        If IsCdLin(A(J)) Then
            M1 = J
            GoTo M1IsFnd
        End If
    Next
    M1 = -1
M1IsFnd:
Dim M2&
    For J = M1 + 1 To MthIx - 1
        If Trim(A(J)) <> "" Then
            M2 = J
            GoTo M2IsFnd
        End If
    Next
    M2 = MthIx
M2IsFnd:
SrcMthIxTopRmkFm = M2
End Function


Private Sub Z()
Z_MthFmCntAyWithTopRmk
End Sub
