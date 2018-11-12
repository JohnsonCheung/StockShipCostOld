Attribute VB_Name = "MIde_Ens_CSub_Brk_Mth"
Option Compare Database
Option Explicit
Function MdCSubBrkMthAy(A As CodeModule) As CSubBrkMth()
If MdIsNoLin(A) Then Exit Function
Dim Ix() As FTIx, Src$(), I, Nm$
Src = MdSrc(A)
Nm = MdNm(A)
Ix = SrcMthFTIxAy(Src)
For Each I In AyNz(Ix)
    PushObj MdCSubBrkMthAy, ZBrk(Nm, Src, CvFTIx(I))
Next
End Function

Private Function ZBrk(MdNm$, Src, MthFTIx As FTIx) As CSubBrkMth
Set ZBrk = New CSubBrkMth
Dim IFm&, ITo&, IsUsingCSub As Boolean, MthNm$
    IFm = MthFTIx.FmIx
    ITo = MthFTIx.ToIx
    IsUsingCSub = ZIsUsingCSub(Src, IFm, ITo)
    MthNm = LinMthNm(Src(IFm))
With ZBrk
    .MdNm = MdNm
    .MthNm = MthNm
    .IsUsingCSub = IsUsingCSub
    .NewCSub = ZNewCSub(MthNm)
    .NewLno = ZNewLno(Src, IFm, ITo)
    .OldLno = ZOldLno(Src, IFm, ITo)
    If .OldLno > 0 Then _
    .OldCSub = Src(.OldLno - 1)
    .NeedDlt = ZNeedDlt(IsUsingCSub, .NewCSub, .OldCSub)
    .NeedIns = ZNeedIns(IsUsingCSub, .NewCSub, .OldCSub)
End With
End Function

Private Function ZOldLno&(Src, IFm&, ITo&)
Dim J&
For J = IFm To ITo
    If HasPfx(Src(J), "Const CSub$") Then
        ZOldLno = J + 1
        Exit Function
    End If
Next
End Function

Private Function ZNewCSub$(MthNm$)
ZNewCSub = "Const CSub$ = CMod & """ & MthNm & """"
End Function

Private Function ZNewLno&(Src, IFm&, ITo&)
If IFm = ITo Then Exit Function
Dim J&, Fm&
Fm = ZNewLno1(Src, IFm, ITo) ' Ix after the MthDcl line
For J = Fm To ITo
    If IsCdLin(Src(J)) Then ZNewLno = J + 1: Exit Function
Next
Stop
End Function
Private Function ZNewLno1&(Src, IFm&, ITo&)
Dim J&
For J = IFm To ITo
    If Not HasSfx(Src(J), " _") Then
        ZNewLno1 = J + 1
        Exit Function
    End If
Next
Stop
End Function
Private Function ZNeedIns(IsUsingCSub As Boolean, NewCSub$, OldCSub$) As Boolean
If Not IsUsingCSub Then Exit Function
If NewCSub = OldCSub Then Exit Function
ZNeedIns = True
End Function

Private Function ZNeedDlt(IsUsingCSub As Boolean, NewCSub$, OldCSub$) As Boolean
If OldCSub = "" Then Exit Function
If IsUsingCSub Then
    ZNeedDlt = NewCSub <> OldCSub
Else
    ZNeedDlt = OldCSub <> ""
End If
End Function

Private Function ZIsUsingCSub(Src, IFm&, ITo&) As Boolean
Dim J&
For J = IFm To ITo
    If HasSubStrAy(Src(J), ZIsUsingCSub1) Then ZIsUsingCSub = True
Next
End Function
Private Function ZIsUsingCSub1() As String()
Static O$()
If Sz(O) = 0 Then
Const A$ = " CSub" & ","
Const B$ = "(CSub" & ","
O = ApSy(A, B)
End If
ZIsUsingCSub1 = O
End Function

