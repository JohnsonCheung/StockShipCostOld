Attribute VB_Name = "MIde_Mth_Nm"
Option Compare Database
Option Explicit

Function CurMdMthNy() As String()
CurMdMthNy = MdMthNy(CurMd)
End Function

Function CurMthDNm$()
Dim M$: M = CurMthNm
If M = "" Then Exit Function
CurMthDNm = CurMdDNm & "." & M
End Function

Function CurPjMthNy() As String()
CurPjMthNy = PjMthNy(CurPj)
End Function

Function CurPjMthNyWh(A As WhPjMth) As String()
Stop
'CurPjMthNy = PjMthNy(CurPj, CvPatn(MthPatn), MthExl, WhMdyAy, WhKdAy, MdPatn, MdExl, WhCmpTy)
End Function

Function CurVbeMthNy(Optional A As WhPjMth) As String()
CurVbeMthNy = VbeMthNy(CurVbe, A)
End Function

Function DDNmMth(MthDDNm$) As Mth
Dim M As CodeModule
Dim Nm$
Dim Ny$(): Ny = Split(MthDDNm, ".")
Select Case Sz(Ny)
Case 1: Nm = Ny(0): Set M = CurMd
Case 2: Nm = Ny(1): Set M = Md(Ny(0))
Case 3: Nm = Ny(2): Set M = PjMd(Pj(Ny(0)), Ny(1))
Case Else: Stop
End Select
Set DDNmMth = Mth(M, Nm)
End Function

Function FbMthNy(A) As String()
FbMthNy = VbeMthNy(FbAcs(A).Vbe)
End Function

Function LinMthNm$(A)
LinMthNm = MthNmBrkNm(LinMthNmBrk(A))
End Function

Function LinPrpNm$(A)
Dim L$
L = RmvMdy(A)
If ShfKd(L) <> "Property" Then Exit Function
LinPrpNm = TakNm(L)
End Function

Function MdDftMthNm$(Optional A As CodeModule, Optional MthNm$)
If MthNm = "" Then
   MdDftMthNm = MdCurMthNm(DftMd(A))
Else
   MdDftMthNm = A
End If
End Function

Function MdMthDDNy(A As CodeModule) As String()
MdMthDDNy = SrcMthDDNy(MdSrc(A))
End Function

Function MdMthNy(A As CodeModule, Optional B As WhMth) As String()
MdMthNy = SrcMthNy(MdBdyLy(A), B)
End Function

Function MthDDNyWh(A$(), B As WhMth) As String()
If IsNothing(B) Then
    MthDDNyWh = A
    Exit Function
End If
Dim N
For Each N In AyNz(A)
    If IsMthDDNmSel(N, B) Then
        PushI MthDDNyWh, N
    End If
Next
End Function

Function MthDNm_Nm$(A)
Dim Ay$(): Ay = Split(A, ".")
Dim Nm$
Select Case Sz(Ay)
Case 1: Nm = Ay(0)
Case 2: Nm = Ay(1)
Case 3: Nm = Ay(2)
Case Else: Stop
End Select
MthDNm_Nm = Nm
End Function

Function MthDotNTM$(MthDot$)
'MthDot is a string with last 3 seg as Mdy.Ty.Nm
'MthNTM is a string with last 3 seg as Nm:Ty.Mdy
Dim Ay$(), Nm$, Ty$, Mdy$
Ay = SplitDot(MthDot)
'AyAsg AyPop(Ay), Ay, Nm
'AyAsg AyPop(Ay), Ay, Ty
'AyAsg AyPop(Ay), Ay, Mdy
Push Ay, FmtQQ("?:?.?", Nm, Ty, Mdy)
MthDotNTM = JnDot(Ay)
End Function

Function MthNm$(A As Mth)
MthNm = A.Nm
End Function

Function MthNmMd(A$) As CodeModule '
Dim O As CodeModule
Set O = CurMd
If MdHasMth(O, A) Then Set MthNmMd = O: Exit Function
Dim N$
N = MthFul(A)
If N = "" Then
    Debug.Print FmtQQ("Mth[?] not found in any Pj")
    Exit Function
End If
Set MthNmMd = Md(N)
End Function

Function MthNy(Optional A As WhPjMth) As String()
MthNy = VbeMthNy(CurVbe, A)
End Function

Function SrcMthNy(A$(), Optional B As WhMth) As String()
SrcMthNy = AyWhDist(AyTakBefDot(MthDDNyWh(SrcMthDDNy(A), B)))
End Function


Private Sub Z_FbMthNy()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
'    For Each Fb In AppFbAy
        PushAy O, FbMthNy(Fb)
'    Next
    Brw O
    Return
X_BrwOne:
'    Brw FbMthNy(AppFbAy()(0))
    Return
End Sub

Private Sub Z_LinMthNm()
GoTo ZZ
Dim A$
A = "Function LinMthNm$(A)": Ept = "LinMthNm.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = LinMthNm(A)
    C
    Return
ZZ:
    Dim O$(), L, P, M
    For Each P In VbePjAy(CurVbe)
        For Each M In PjMdAy(CvPj(P))
            For Each L In MdBdyLy(CvMd(M))
                PushNonBlankStr O, LinMthNm(CStr(L))
            Next
        Next
    Next
    Brw O
End Sub

Private Sub Z_MdDftMthNm()
Dim I, Md As CodeModule
For Each I In PjMdAy(CurPj)
   MdShw CvMd(I)
   Debug.Print MdNm(Md), MdDftMthNm(Md)
Next
End Sub

Private Sub Z_SrcMthNy()
Brw SrcMthNy(CurSrc)
End Sub
Function Md_FunNy_OfPfx_ZZDash(A As CodeModule) As String()
Dim J%, O$(), L$, L1$, Is_ZFun As Boolean
For J = 1 To A.CountOfLines
    Is_ZFun = True
    L = A.Lines(J, 1)
    Select Case True
    Case HasPfx(L, "Private Sub ZZ_")
        Is_ZFun = True
        L1 = RmvPfx(L, "Sub ")
    Case HasPfx(L, "Private Sub ZZ_")
        Is_ZFun = True
        L1 = RmvPfx(L, "Sub ")
    Case Else:
        Is_ZFun = False
    End Select

    If Is_ZFun Then
        Push O, TakNm(L1)
    End If
Next
Md_FunNy_OfPfx_ZZDash = O
End Function

Function MdPubMthNy(A As CodeModule) As String()
Const CSub$ = CMod & "MdPubMthNy"
If MdTy(A) <> vbext_ct_StdModule Then Er CSub, "Given [Md]-[Ty] must have type Std", MdNm(A), MdTyStr(A)
MdPubMthNy = AyWhDist(SrcMthNy(MdSrc(A), WhMth("Pub")))
End Function



Private Sub Z()
Z_FbMthNy
Z_LinMthNm
Z_MdDftMthNm
Z_SrcMthNy
End Sub
