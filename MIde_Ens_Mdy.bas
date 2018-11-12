Attribute VB_Name = "MIde_Ens_Mdy"
Option Compare Database
Option Explicit
Const CMod$ = "MIde_Ens_Mdy."

Function DclRmvOpt$(A$)
Dim Ly$(): Ly = SplitCrLf(A)
Dim L, O$()
For Each L In AyNz(Ly)
    If Not HasPfx(L, "Option ") Then
        PushI O, L
    End If
Next
DclRmvOpt = JnCrLf(O)
End Function


Function MthLinEnsPrv$(A$)
If Not LinIsMth(A) Then Er "MthLinEnsPrv", "Given [Lin] is not MthLin", A
MthLinEnsPrv = "Private " & RmvMdy(A)
End Function

Function MthLinEnsPub$(A$)
If Not LinIsMth(A) Then Er "MthLinEnsPub", "Given [Lin] is not MthLin", A
MthLinEnsPub = RmvMdy(A)
End Function

Private Sub Z()
Z_MdMthEnsMdy
Z_MdMthEnsPub
End Sub
Sub EnsZPrv()
MdEnsZPrv CurMd
End Sub
Sub EnsPjZPrv()
PjEnsZPrv CurPj
End Sub

Sub PjEnsZPrv(A As VBProject)
Dim Md
For Each Md In AyNz(PjMdAy(A))
    MdEnsZPrv CvMd(Md)
Next
End Sub

Sub EnsZPub()
MdEnsZDashPub CurMd
End Sub

Sub MdEnsZPrv(A As CodeModule)
MdWhEnsPrv A, WhMth(WhMdy:="Pub", Nm:=WhNm("^Z.+"))
End Sub

Sub MdEnsZDashPub(A As CodeModule)
MdWhEnsPub A, WhMth(WhMdy:="Prv", Nm:=WhNm("^Z_"))
End Sub

Sub MdMthEnsFrd(A As CodeModule, MthNm$)
Const CSub$ = CMod & "MdMthEnsFrd"
If MdTy(A) <> vbext_ct_ClassModule Then Er CSub, "Given [Md]-[Ty] is not Class", MdNm(A), MdTyStr(A)
ZEnsMthMdy A, MthNm, "Friend"
End Sub

Sub MdMthEnsMdy(A As CodeModule, MthNm$, Mdy$)
Dim Lno
For Each Lno In AyNz(MdMthLnoAy(A, MthNm))
   ZEnsMdy A, Lno, Mdy
Next
End Sub

Sub MdMthEnsPrv(A As CodeModule, MthNm$)
ZEnsMthMdy A, MthNm, "Private"
End Sub

Sub MdMthEnsPub(A As CodeModule, MthNm$)
ZEnsMthMdy A, MthNm
End Sub

Sub MdWhEnsPrv(A As CodeModule, B As WhMth)
Dim Ix
For Each Ix In AyNz(SrcMthIxAy(MdSrc(A), B))
    ZEnsPrv A, Ix + 1
Next
End Sub

Sub MdWhEnsPub(A As CodeModule, B As WhMth)
Dim Ix
For Each Ix In SrcMthIxAy(MdSrc(A), B)
    ZEnsPub A, Ix + 1
Next
End Sub
Private Function ZNewLin$(OldLin$, Mdy$)
Const CSub$ = CMod & "ZNewLin"
Dim L$
    L = RmvMdy(OldLin)
    Select Case Mdy
    Case "Pub", "": ZNewLin = L
    Case "Prv":     ZNewLin = "Private " & L
    Case "Frd":     ZNewLin = "Friend " & L
    Case Else
        Er CSub, "Given parameter [Mdy] must be ['' Pub Prv Frd]", Mdy
    End Select
End Function

Private Sub ZEnsMdy(A As CodeModule, MthLno, Optional Mdy$)
Const CSub$ = CMod & "ZEnsMdy"
Dim OldLin$
    OldLin = A.Lines(MthLno, 1)
    If Not LinIsMth(OldLin) Then
       Er CSub, "Given [Md]-[MthLno]-[Lin] is not a method", MdNm(A), MthLno, OldLin
    End If
Dim NewLin$: NewLin = ZNewLin(OldLin, Mdy)
If OldLin = NewLin Then
   Debug.Print CSub; FmtQQ(": Same Mdy[?] in Lin[?]", Mdy, OldLin)
   Exit Sub
End If
MdRplLin A, MthLno, NewLin
Debug.Print CSub
Debug.Print FmtQQ("  Mdy[?] Of MthLno[?] of Md[?] is ensured", Mdy, MthLno, MdNm(A))
Debug.Print FmtQQ("  OldLin[?]", OldLin)
Debug.Print FmtQQ("  NewLin[?]", NewLin)
End Sub

Private Sub ZEnsMthMdy(A As CodeModule, MthNm$, Optional Mdy$)
Dim Lno
For Each Lno In AyNz(MdMthLnoAy(A, MthNm))
   ZEnsMdy A, Lno, Mdy
Next
End Sub

Private Function ZEnsPrv(A As CodeModule, MthLno)
ZEnsMdy A, MthLno, "Prv"
End Function

Private Function ZEnsPub(A As CodeModule, MthLno)
ZEnsMdy A, MthLno
End Function

Private Sub Z_MdMthEnsMdy()
Dim M As CodeModule
Dim MthNm$
Dim Mdy$
'--
Set M = CurMd
MthNm = "Z_A"
Mdy = "Prv"
GoSub Tst
Exit Sub
Tst:
    MdMthEnsMdy M, MthNm, Mdy
    Return
End Sub

Private Sub Z_MdMthEnsPub()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "YYA")
MdMthEnsPrv M, "ZZA": Ass MthDcl(M) = "Private Property Get ZZA()"
MdMthEnsPub M, "ZZA": Ass MthDcl(M) = "Property Get ZZA()"
End Sub

Function LinIsPrvZZDash(A) As Boolean
Dim L$: L = A
If ShfMdy(L) <> "Private" Then Exit Function
If ShfMthTy(L) <> "Sub" Then Exit Function
If Left(L, 3) <> "ZZ_" Then Exit Function
LinIsPrvZZDash = True
End Function

Function LinIsPubZDash(A) As Boolean
Dim L$: L = A
If Not AyHas(Array("", "Public"), ShfMdy(L)) Then Exit Function
If ShfMthTy(L) <> "Sub" Then Exit Function
LinIsPubZDash = Left(L, 2) = "Z_"
End Function

Function LinIsPubZZDash(A) As Boolean
Dim L$: L = A
If Not AyHas(Array("", "Public"), ShfMdy(L)) Then Exit Function
If ShfMthTy(L) <> "Sub" Then Exit Function
LinIsPubZZDash = Left(L, 3) = "ZZ_"
End Function

Private Sub ZZ_LinIsPrvZZDash()
Dim L
For Each L In CurSrc
    If LinIsPrvZZDash(L) Then
        Debug.Print L
    End If
Next
End Sub

Private Sub ZZ_LinIsPubZZDash()
Dim L
For Each L In CurSrc
    If LinIsPubZZDash(L) Then
        Debug.Print L
    End If
Next
End Sub

Sub MdMthSetMdy(A As CodeModule, MthNm$, Mdy$)
Ass IsMdy(Mdy)
Dim I&
    I = MdMthLno(A, MthNm)
Dim L$
    L = A.Lines(I, 1)
Dim Old$
    Old = LinMdy(L)
If Mdy = Old Then Exit Sub
Dim NewL$
    Dim B$
    If Mdy <> "" Then
        B = Mdy & " "
    Else
        B = Mdy
    End If
    NewL = B & L
With A
    .DeleteLines I, 1
    .InsertLines I, NewL
End With
End Sub

Sub MdMthSetPrv(A As CodeModule, MthNm$)
MdMthSetMdy A, MthNm, "Private"
End Sub

Sub MdMthSetPub(A As CodeModule, MthNm$)
MdMthSetMdy A, MthNm, ""
End Sub

Sub MthSetMdy(A As CodeModule, MthNm$, Mdy$)
Ass IsMdy(Mdy)
Dim I&
    I = MdMthLno(A, MthNm)
Dim L$
    L = MdLin(A, I)
Dim Old$
    Old = LinMdy(L)
If Mdy = Old Then Exit Sub
Dim NewL$
    Dim B$
    If Mdy <> "" Then
        B = Mdy & " "
    Else
        B = Mdy
    End If
    NewL = B & L
With A
    .DeleteLines I, 1
    .InsertLines I, NewL
End With
End Sub

Sub MthSetPrv(A As CodeModule, MthNm$)
MthSetMdy A, MthNm, "Private"
End Sub

Sub MthSetPub(A As CodeModule, MthNm$)
MthSetMdy A, MthNm, ""
End Sub
Sub Ens_Vbe_ZZDashPubMthAsPrivate()
VbeEnsZZDashPubMthAsPrivate CurVbe
End Sub


Sub PjEnsZDashMthAsPrv(A As VBProject)
ItrDo PjMdAy(A), "MdEnsZ3DMthAsPrivate"
End Sub

Sub PjEnsZZDashAsPrv(A As VBProject)

End Sub

Sub PjEnsZZDashAsPub(A As VBProject)
AyDo PjMdAy(A), "MdEnsZZDashAsPrv"
End Sub


Private Sub MdEnsZZDashAsPrv(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashAsPrv Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If LinIsPubZZDash(L) Then
        Debug.Print L
        By = MthLinEnsPrv(L)
        Debug.Print FmtQQ("MdEnsZZDashAsPrv: Md(?) Lin(?) is change to Private: [?]", DNm, J, By)
        'A.ReplaceLine J, By
    End If
Next
End Sub

Private Sub MdEnsZZDashPrvMthAsPub(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPub: Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If LinIsPrvZZDash(L) Then
        By = MthLinEnsPub(L)
        Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPub: Md(?) Lin(?) is change to Public: [?]", DNm, J, By)
        A.ReplaceLine J, By
    End If
Next
End Sub
