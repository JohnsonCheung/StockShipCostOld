Attribute VB_Name = "MVb_Er_Msg"
Option Compare Database
Option Explicit
Function MsgNyAvLy(DotMsg$, Ny0, Av()) As String()
MsgNyAvLy = AyAddAp(DotMsgLy(DotMsg), NyAvLy(Ny0, Av, 4))
End Function


Function FunMsgNyAvLy(Fun$, DotMsg$, Ny0, Av()) As String()
FunMsgNyAvLy = AyAddAp(DotMsgLy(DotMsg, Fun), NyAvLy(Ny0, Av, 4))
End Function

Private Function DotMsgLy(DotMsg$, Optional Fun$) As String()
Dim A1$, A2$
    Brk1Asg DotMsg, ".", A1, A2
PushI DotMsgLy, A1 & ".  " & Fun
If A2 = "" Then Exit Function
PushIAy DotMsgLy, AyAddPfx(LinesWrap(A2), "    | ")
End Function

Function FunMsgNyApLy(Fun$, Msg$, Ny0, ParamArray Ap()) As String()
Dim Av(): Av = Ap
FunMsgNyApLy = FunMsgNyAvLy(Fun, Msg, Ny0, Av)
End Function

Sub FunMsgNyApDmp(Fun$, Msg$, Ny0, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgNyAvLy(Fun, Msg, Ny0, Av)
End Sub

Sub FunMsgAvBrw(A, Msg$, Av())
AyBrw FunMsgAvLy(A, Msg, Av)
End Sub

Function FunMsgAvLy(A, Msg$, Av()) As String()
Dim B$(), C$()
B = SplitVBar(Msg)
C = NyAvLy(CvSy(AyAdd(ApSy("Fun"), MsgNy(Msg))), CvAy(AyAdd(Array(A), Av)))
FunMsgAvLy = AyAdd(B, C)
End Function

Sub MsgApBrw(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAvBrw Msg, Av
End Sub

Sub MsgApDmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
AyDmp MsgAvLy(A, Av)
End Sub

Function MsgApLin$(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgApLin = MsgAvLin(A, Av)
End Function

Function MsgApLy(A$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
MsgApLy = MsgAvLy(A, Av)
End Function

Function NmvLy(Nm$, V) As String()
Dim Ly$(): Ly = VarLy(V)
Dim J%, S$
If Sz(Ly) = 0 Then
    PushI NmvLy, Nm & ": "
Else
    PushI NmvLy, Nm & ": " & Ly(0)
End If
S = Space(Len(Nm) + 2)
For J = 1 To UB(Ly)
    PushI NmvLy, S & Ly(J)
Next
End Function

Function NmvStr$(Nm$, V)
NmvStr = Nm & "=[" & VarStr(V) & "]"
End Function

Function NyAvLin$(A$(), Av())
Dim U&
U = UB(A)
If U = -1 Then Exit Function
Dim O$(), J%
For J = 0 To U
    Push O, NmvStr(A(J), Av(J))
Next
NyAvLin = Join(AyAddPfx(O, " | "))
End Function

Function NyAvLy(Ny0, Av(), Optional Indent%) As String()
Dim W%, O$(), J%, A1$(), A2$(), Ny$()
Ny = CvNy(Ny0)
W = AyWdt(Ny)
A1 = AyAlignL(Ny)
AyabSetSamMax A1, Av
For J = 0 To UB(A1)
    PushAy O, NmvLy(A1(J), Av(J))
Next
NyAvLy = AyAddPfx(O, Space(Indent))
End Function

Function NyAvScl$(A$(), Av())
Dim O$(), J%, X, Y
X = A
Y = Av
AyabSetSamMax X, Y
For J = 0 To UB(X)
    Push O, RmvSqBkt(X(J)) & "=" & VarStr(Y(J))
Next
NyAvScl = JnSemiColon(O)
End Function

Sub NyApDmp(Ny0, ParamArray Ap())
Dim Av(): Av = Ap
D NyAvLy(Ny0, Av, 0)
End Sub

Sub FunMsgAvLyDmp(A$, Msg$, Av())
D FunMsgAvLy(A, Msg, Av)
End Sub

Sub FunMsgAvLinDmp(A$, Msg$, Av())
D FunMsgAvLin(A, Msg, Av)
End Sub

Sub FunMsgApDmp(A$, MsgSV$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgAvLinDmp A, MsgSV, Av
End Sub

Sub MsgStop(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D MsgNyAvLy(Msg, MacroNy(Msg), Av)
End Sub

Sub MsgWhStop(Msg$, FF, ParamArray Ap())
Dim Av(): Av = Ap
D MsgNyAvLy(Msg, FF, Av)
End Sub

Sub Msg(Fun$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgNyAvLy(Fun, Msg, MacroNy(Msg), Av)
End Sub
Sub MsgObjPrp(Fun$, Msg$, Obj, PrpNy0)
MsgWh Fun, Msg, PrpNy0, ObjPrpAy(Obj, PrpNy0)
End Sub

Sub MsgWh(Fun$, Msg$, FF, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgNyAvLy(Fun, Msg, FF, Av)
End Sub

Function FunMsgAvLin$(Fun$, MacroStr$, Av())
End Function

Function FunMsgLin$(Fun$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgLin = FunMsgAvLin(Fun, Msg, Av)
End Function

Sub FunMsgLinDmp(Fun$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgAvLin(Fun, Msg, Av)
End Sub

Function FunMsgLy(A, Msg$, Av()) As String()
FunMsgLy = FunMsgAvLy(A, Msg, Av)
End Function

Sub FunMsgLyDmp(A, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgAvLy(A, Msg, Av)
End Sub

Function NmssAvLy(A$, Av()) As String()
NmssAvLy = NyAvLy(SslSy(A), Av)
End Function

Function NmssApLy(A$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
NmssApLy = NyAvLy(SslSy(A), Av)
End Function

Function MacroStrAvLy(A$, Av()) As String()
MacroStrAvLy = NyAvLy(MacroNy(A, OpnBkt:="["), Av)
End Function

Function FunMsgLines$(Fun$, MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
'ErMsgLines = ErMsgLinesByAv(Fun, MacroStr, Av)
End Function
Sub MsgAvBrw(A$, Av())
AyBrw MsgAvLy(A, Av)
End Sub

Function MsgAvLin$(A$, Av())
Dim B$(), C$
C = NyAvLin(MsgNy(A), Av)
MsgAvLin = EnsSfxDot(A) & C
End Function

Function MsgAvLy(A$, Av()) As String()
Dim B$(), C$()
B = SplitVBar(A)
C = AyTab(NyAvLy(MsgNy(A), Av))
MsgAvLy = AyAdd(B, C)
End Function

Sub MsgBrw(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAvBrw A, Av
End Sub

Sub MsgBrwStop(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAvBrw A, Av
Stop
End Sub

Sub MsgDmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
AyDmp MsgAvLy(A, Av)
End Sub

Function MsgLin$(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgLin = MsgAvLin(A, Av)
End Function

Function MsgLy(A$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
MsgLy = MsgAvLy(A, Av)
End Function

Function MsgNy(A) As String()
Dim O$(), P%, J%
O = Split(A, "[")
AyShf O
For J = 0 To UB(O)
    P = InStr(O(J), "]")
    O(J) = "[" & Left(O(J), P)
Next
MsgNy = O
End Function


Function VarStr$(V)
Select Case True
Case IsPrim(V):    VarStr = V
Case IsArray(V):   VarStr = AyLines(V)
Case IsNothing(V): VarStr = "*Nothing"
Case IsObject(V):  VarStr = "*Type[" & TypeName(V) & "]"
Case IsEmpty(V):   VarStr = "*Empty"
Case IsMissing(V): VarStr = "*Missing"
Case Else: Stop
End Select
End Function

Function VarLy(A) As String()
VarLy = SplitCrLf(VarLines(A))
End Function

Function VarLines$(A, Optional Lvl%)
Dim T$, S$, W%, I, O$(), Sep$
Select Case True
Case IsDic(A): VarLines = JnCrLf(DicFmt(CvDic(A)))
Case IsPrim(A): VarLines = A
Case IsLinesAy(A): VarLines = LinesAyLines(CvSy(A))
Case IsSy(A): VarLines = JnCrLf(A)
Case IsNothing(A): VarLines = "#Nothing"
Case IsEmpty(A): VarLines = "#Empty"
Case IsMissing(A): VarLines = "#Missing"
Case IsObject(A): VarLines = "#Obj(" & TypeName(A) & ")"
Case IsArray(A)
    If Sz(A) = 0 Then Exit Function
    For Each I In A
        PushI O, VarLines(I, Lvl + 1)
    Next
    If Lvl > 0 Then
        W = LinesAyWdt(O)
        Sep = LvlSep(Lvl)
        PushI O, StrDup(Sep, W)
    End If
    VarLines = JnCrLf(O)
Case Else
End Select
End Function

Function LvlSep$(Lvl%)
Select Case Lvl
Case 0: LvlSep = "."
Case 1: LvlSep = "-"
Case 2: LvlSep = "+"
Case 3: LvlSep = "="
Case 4: LvlSep = "*"
Case Else: LvlSep = Lvl
End Select
End Function
