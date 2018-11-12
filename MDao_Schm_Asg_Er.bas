Attribute VB_Name = "MDao_Schm_Asg_Er"
Option Explicit
Option Compare Database
Private Type T: Lno As Integer: T As String: Fny() As String: Sk() As String:     End Type
Private Type F: Lno As Integer: E As String: LikT As String:  LikFny() As String: End Type
Private Type D: Lno As Integer: T As String: F As String:     Des As String:     End Type
Private Type E
    Lno As Integer
    E As String
    Ty As DAO.DataTypeEnum
    Req As Boolean
    ZLen As Boolean
    TxtSz As Byte
    Expr As String
    VRul As String
    Dft As String
    VTxt As String
End Type
Private A_Ly$(), O_Er$()
Private B_Tny$(), B_Eny$()
Private B_T() As T
Private B_D() As D
Private B_E() As E
Private B_F() As F

Private Sub AMain()
Dim X$()
MDao_Schm_Asg_Er.SchmLyEr X
End Sub

Private Sub Brk()
Dim X()  As Lnx
Dim XE() As Lnx
Dim XF() As Lnx
Dim XD() As Lnx
Dim XT() As Lnx
X = LyClnLnxAy(A_Ly)
XT = LnxAyWhRmvT1(X, "Tbl")
XD = LnxAyWhRmvT1(X, "Des")
XE = LnxAyWhRmvT1(X, "Ele")
XF = LnxAyWhRmvT1(X, "Fld")
B_E = BrkE(XE)
B_F = BrkF(XF)
B_D = BrkD(XD)
B_T = BrkT(XT)
B_Eny = AyTakT1(LnxAyLy(XE))
B_Tny = AyTakT1(LnxAyLy(XT))
O_Er = LnxAyT1Chk(X, "Des Ele Fld Tbl")
End Sub

Private Function BrkD(D() As Lnx) As D()
Dim J%
For J = 0 To UB(D)
    XPushD BrkD, BrkDLin(D(J))
Next
End Function

Private Function BrkDLin(D As Lnx) As D
Dim V$
With BrkDLin
    AyAsg Lin3TRst(D.Lin), .T, .F, V, .Des
    If V <> "|" Then Push O_Er, "..."
End With
End Function

Private Function BrkE(A() As Lnx) As E()
Dim J%
For J = 0 To UB(A)
    XPushE BrkE, BrkELin(A(J))
Next
End Function

Private Function BrkELin(ELin As Lnx) As E
Dim L$, Ty$
Dim E$  ' EleNm of the Ele-Lin
L = ELin.Lin
With BrkELin
    AyAsg ShfVal(L, "*Nm *Ty ?Req ?ZLen TxtSz VTxt Dft VRul Expr"), _
                     .E, Ty, .Req, .ZLen, .TxtSz, .VTxt, .Dft, .VRul, .Expr
    .Ty = DaoTy(Ty)
    If L <> "" Then
        Push O_Er, ErMsg_ExcessEleItm(ELin.Ix, L)
    End If
    If .Ty = 0 Then
        Push O_Er, ErMsg_TyEr(ELin.Ix, Ty)
    End If
End With
End Function

Private Function BrkF(A() As Lnx) As F()
Dim J%
For J = 0 To UB(A)
    XPushF BrkF, BrkFLin(A(J))
Next
End Function

Private Function BrkFLin(F As Lnx) As F
Dim LikFF$, A$, V$
With BrkFLin
    AyAsg Lin3TRst(F.Lin), .E, .LikT, V, A
    .LikFny = SslSy(LikFF)
End With
End Function

Private Function BrkT(A() As Lnx) As T()
If Sz(A) = 0 Then
    Push O_Er, ErMsg_NoTLin
    Exit Function
End If
Dim J%
For J = 0 To UB(A)
    XPushT BrkT, BrkTLin(A(J))
Next
End Function

Private Function BrkTLin(T As Lnx) As T
Dim A$, B$, C$, D$
BrkAsg T.Lin, "|", A, B
With BrkTLin
    .T = A
    B = Replace(B, "*", A)
    Brk1Asg B, "|", C, D
    If D = "" Then
        .Fny = SslSy(C)
    Else
        .Sk = SslSy(RmvPfx(C, A & " "))
        .Fny = SslSy(Replace(B, "|", " "))
    End If
    If Sz(.Fny) = 0 Then
        Push O_Er, "should have fields after |"
        Exit Function
    End If
    Dim Dup$()
    Dup = AyWhDup(.Fny)
    If Sz(Dup) > 0 Then
        Stop '
'       Push BrkTLin.Er, ErMsg_DupF(T.Ix + 1)
        Exit Function
    End If
End With
End Function

Private Function Er_DupE() As String()
Dim E
For Each E In AyNz(AyWhDup(B_Eny))
    Push Er_DupE, ErMsg_DupE(FndELnoAy(E), E)
Next
End Function

Private Function Er_DupT() As String()
Dim T
For Each T In AyNz(AyWhDup(Tny))
    Push Er_DupT, ErMsg_DupT(FndTLnoAy(T), T)
Next
End Function

Private Function Er_EzFLy_NotIn_Eny() As String()
Dim J%, O$(), Essl$
For J = 0 To XFUB(B_F)
    With B_F(J)
        If Not AyHas(B_Eny, .E) Then Push O, ErMsg_EzFLy_NotIn_Eny(.Lno, .E, Essl)
    End With
Next
Er_EzFLy_NotIn_Eny = O
End Function

Private Function Er_FzDLy_NotIn_TblFny() As String()
Dim J%, Fny1$()
For J = 0 To XDUB(B_D)
    With B_D(J)
        If Not AyHas(B_Tny, .T) Then GoTo Nxt
        Fny1 = FndFny(.T)
        If Not AyHas(Fny1, .F) Then
            Push Er_FzDLy_NotIn_TblFny, ErMsg_FzDLy_NotIn_TblFny(.Lno, .T, .F, JnSpc(Fny1))
        End If
    End With
Nxt:
Next
End Function

Private Function Er_TzDLy_NotIn_Tny() As String()
Dim Tssl$, J%
Tssl = JnSpc(Tny)
For J = 0 To XDUB(B_D)
    With B_D(J)
        If Not AyHas(Tny, .T) Then
            Push Er_TzDLy_NotIn_Tny, ErMsg_TzDLy_NotIn_Tny(.Lno, .T, Tssl)
        End If
    End With
Next
End Function

Private Function Er1(Fld) As String()
Const CSub$ = CMod & "Er1"
Const M$ = "[Fld] cannot of get a Fd by [Fld-Drs] and [Ele-Dic]"
'PushI Er1, FunMsgLin(CSub, M, Fld, Ele)
Stop '
End Function

Private Function Er3(Fld, Ele) As String()
Const Msg1$ = "[Fld] cannot of get a Fd by [Fld-Drs] and [Ele-Dic]"
PushI Er3, MsgLin(Msg1, Fld, Ele)
End Function

Private Function ErMsg$(Lno%, M$)
ErMsg = "--Lno" & Lno & ".  " & M
End Function

Private Function ErMsg_DupE$(LnoAy%(), E)
ErMsg_DupE = ErMsg1(LnoAy, FmtQQ("This E[?] is dup", E))
End Function

Private Function ErMsg_DupF$(Lno%, T$, Fny$())
ErMsg_DupF = ErMsg(Lno, FmtQQ("F[?] is dup in T[?]", JnSpc(Fny), T))
End Function

Private Function ErMsg_DupT$(LnoAy%(), T)
ErMsg_DupT = ErMsg1(LnoAy, FmtQQ("This T[?] is dup", T))
End Function

Private Function ErMsg_ExcessEleItm$(Lno%, L$)
ErMsg_ExcessEleItm = ErMsg(Lno, FmtQQ("Excess Ele Item[?]", L))
End Function

Private Function ErMsg_ExcessTxXTSz$(Lno%, Ty$)
ErMsg_ExcessTxXTSz = ErMsg(Lno, FmtQQ("Ty[?] is not Txt, it should not have TxtSz", Ty))
End Function

Private Function ErMsg_EzFLy_NotIn_Eny$(Lno%, E$, Essl$)
ErMsg_EzFLy_NotIn_Eny = ErMsg(Lno, FmtQQ("E[?] of is not in E-Lin[?]", E, Essl))
End Function

Private Function ErMsg_FldEleEr$(Lno%, E$, Essl$)
ErMsg_FldEleEr = ErMsg(Lno, FmtQQ("E[?] is invalid.  Valid E is [?]", E, Essl))
End Function

Private Function ErMsg_FzDLy_NotIn_TblFny$(Lno%, T$, F$, Fssl$)
ErMsg_FzDLy_NotIn_TblFny = ErMsg(Lno, FmtQQ("F[?] is invalid in T[?].  Valid F[?]", F, T, Fssl))
End Function

Private Function ErMsg_NoELin$()
ErMsg_NoELin = "No E-Line"
End Function

Private Function ErMsg_NoFLin$()
ErMsg_NoFLin = "No F-Line"
End Function

Private Function ErMsg_NoTLin$()
ErMsg_NoTLin = "No T-Line"
End Function

Private Function ErMsg_TblFldEr$(Lno%, T$, F$)
ErMsg_TblFldEr = ErMsg(Lno, FmtQQ("T[?] has invalid F[?], which cannot be found in any F-Lines"))
End Function

Private Function ErMsg_TyEr$(Lno%, Ty$)
ErMsg_TyEr = ErMsg(Lno, FmtQQ("Invalid DaoTy[?].  Valid Ty[?]", Ty, DaoTySsl))
End Function

Private Function ErMsg_TzDLy_NotIn_Tny$(Lno%, T$, Tssl$)
ErMsg_TzDLy_NotIn_Tny = ErMsg(Lno, FmtQQ("T[?] is invalid.  Valid T[?]", T, Tssl))
End Function

Private Function ErMsg1(LnoAy%(), M$)
ErMsg1 = "--" & Join(AyAddPfx(LnoAy, "Lno"), ".") & "  " & M
End Function

Private Function FldChk(F, EF As EF) As String()
If IsStdFld(F) Then Exit Function
Dim Ele$
Ele = FldChk1$(F, EF.F): If Ele = "" Then FldChk = Er1(F): Exit Function
If EleNmIsStd(Ele) Then Exit Function
If Not EF.E.Exists(Ele) Then FldChk = Er3(F, Ele)
End Function

Private Function FldChk1$(Fld, F As Drs) ' Return Ele$
Dim Dr
For Each Dr In AyNz(F.Dry)
    If Fld Like Dr(1) Then FldChk1 = Dr(0): Exit Function
Next
End Function

Private Function FndELnoAy(Ele) As Integer()
Dim J%
For J = 0 To UBound(B_E)
    If B_E(J).E = Ele Then
        Push FndELnoAy, B_E(J).Lno
    End If
Next
End Function

Private Function FndFny(Tbl) As String()
Dim J%
With FndT(Tbl)
    FndFny = .Fny
    If .T <> Tbl Then Stop
End With
End Function

Private Function FndT(Tbl) As T
Dim J%
For J = 0 To UBound(B_T)
    With B_T(J)
        If .T = Tbl Then FndT = B_T(J): Exit Function
    End With
Next
End Function

Private Function FndTLnoAy(Tbl) As Integer()
Dim J%
For J = 0 To XTUB(B_T)
    If B_T(J).T = Tbl Then
        PushI FndTLnoAy, B_T(J).Lno
    End If
Next
End Function

Function SchmLyEr(Ly$()) As String()
A_Ly = Ly
Brk
Er_DupT
Er_DupE
Er_TzDLy_NotIn_Tny
Er_FzDLy_NotIn_TblFny
Er_EzFLy_NotIn_Eny
End Function

Private Sub TdDefAss(A, EF As EF)
ChkAss TdDefChk(A, EF)
End Sub

Function TdDefChk(TdDef, EF As EF) As String() ' Chk may return Sz=0, But Er always Sz>0
Dim Fny$(), F, O$()
Fny = TdDefFny(Stru)
For Each F In Fny
    PushIAy O, FldChk(F, EF)
Next
If Sz(0) > 0 Then
    TdDefChk = AyAddAp("", O)
End If
End Function

Private Function XDSz%(A() As D): On Error Resume Next: XDSz = UBound(A) + 1: End Function

Private Function XDUB%(A() As D): XDUB = XDSz(A) - 1: End Function

Private Function XESz%(A() As E): On Error Resume Next: XESz = UBound(A) + 1: End Function

Private Function XEUB%(A() As E): XEUB = XESz(A) - 1: End Function

Private Function XFSz%(A() As F): On Error Resume Next: XFSz = UBound(A) + 1: End Function

Private Function XFUB%(A() As F): XFUB = XFSz(A) - 1: End Function

Private Sub XPushD(O() As D, M As D): Dim N&: N = XDSz(O): ReDim Preserve O(N): O(N) = M: End Sub

Private Sub XPushE(O() As E, M As E): Dim N&: N = XESz(O): ReDim Preserve O(N): O(N) = M: End Sub

Private Sub XPushF(O() As F, M As F): Dim N&: N = XFSz(O): ReDim Preserve O(N): O(N) = M: End Sub

Private Sub XPushT(O() As T, M As T): Dim N&: N = XTSz(O): ReDim Preserve O(N): O(N) = M: End Sub

Private Function XTSz%(A() As T): On Error Resume Next: XTSz = UBound(A) + 1: End Function

Private Function XTUB%(A() As T): XTUB = XTSz(A) - 1: End Function

Private Function ZZ_IsTItmEq(A As T, B As T) As Boolean
If A.T <> B.T Then Exit Function
If Not AyIsEq(A.Fny, B.Fny) Then Exit Function
ZZ_IsTItmEq = AyIsEq(A.Sk, B.Sk)
End Function

Private Sub Z()
Z_BrkTLin
MDao_Schm_Asg_Er:
End Sub

Private Sub Z_BrkTLin()
Dim Act As T
Dim Ept As T
Dim Emp As T
Dim EptEr$()
Dim TLnx As Lnx
Set TLnx = Lnx(999, "A")
Ept = Emp
Push EptEr, "should have a |"
GoSub Tst
'
TLnx.Lin = "A | B B"
Ept = Emp
Push EptEr, "dup fields[B]"
GoSub Tst
'
TLnx.Lin = "A | B B D C C"
Ept = Emp
Push EptEr, "dup fields[B C]"
GoSub Tst
'
TLnx.Lin = "A | * B D C"
Ept = Emp
With Ept
    .T = "A"
    .Fny = SslSy("A B D C")
End With
GoSub Tst
'
TLnx.Lin = "A | * B | D C"
Ept = Emp
With Ept
    .T = "A"
    .Fny = SslSy("A B D C")
    .Sk = SslSy("B")
End With
GoSub Tst
'
TLnx.Lin = "A |"
Ept = Emp
Push EptEr, "should have fields after |"
GoSub Tst
Exit Sub
Tst:
    Erase O_Er
    Act = BrkTLin(TLnx)
    Ass AyIsEq(O_Er, EptEr)
    Ass ZZ_IsTItmEq(Act, Ept)
    Return
End Sub
