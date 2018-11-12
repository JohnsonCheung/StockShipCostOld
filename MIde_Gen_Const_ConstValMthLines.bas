Attribute VB_Name = "MIde_Gen_Const_ConstValMthLines"
Option Compare Database
Option Explicit
Const CMod$ = "MIde_Gen_Const_ConstValMthLines."

Function ConstValMthLines$(ConstVal$, Nm$, IsPub As Boolean) _
' Return [MthLines] by [ConstVal$] and [Nm$]
Const CSub$ = CMod & "ConstValMthLines"
If ConstVal = "" Then Er CSub, "Given ConstVal is blank"
Dim A$()
Dim NChunk%
    A = SplitCrLf(ConstVal)
    NChunk = ZNChunk(Sz(A))
Dim O$()
    Dim J%
    For J = 0 To NChunk - 1
        PushI O, ZChunk(A, J)
    Next
    PushI O, ZLasLin(Nm, NChunk)
ConstValMthLines = ZMakeMth(JnCrLf(O), Nm, IsPub)
End Function

Private Function ZChunk$(ConstLy$(), IChunk%)
If Sz(ConstLy) = 0 Then Stop
Dim Ly$()
    Ly = AyMid(ConstLy, IChunk * 20, 20)
Dim O$()
    Dim L$, J&, U&
    U = UB(Ly)
    For J = 0 To U
        L = QuoteAsVb(Ly(J))
        Select Case True
        Case J = 0 And J = U: Push O, FmtQQ("Const A_?$ = ?", IChunk + 1, L)
        Case J = 0:           Push O, FmtQQ("Const A_?$ = ? & _", IChunk + 1, L)
        Case J = U:           Push O, "vbCrLf & " & L
        Case Else:            Push O, "vbCrLf & " & L & " & _"
        End Select
    Next
ZChunk = JnCrLf(O) & vbCrLf
End Function

Private Function ZLasLin$(Nm$, NChunk%)
Dim B$
    Dim O$(), J%
    For J = 1 To NChunk
        PushI O, "A_" & J
    Next
    B = Join(O, " & vbCrLf & ")
ZLasLin = Nm & " = " & B
End Function

Private Function ZMakeMth$(Lines$, Nm$, IsPub As Boolean)
Dim L1$, L2$
L1 = IIf(IsPub, "", "Private ") & "Function " & Nm & "$()" & vbCrLf
L2 = vbCrLf & "End Function"
ZMakeMth = vbCrLf & L1 & Lines & L2
End Function

Private Function ZNChunk%(Sz%)
ZNChunk = ((Sz - 1) \ 20) + 1
End Function

Private Sub Z()
Z_ConstValMthLines
End Sub

Private Sub Z_ConstValMthLines()
Const CSub$ = CMod & "Z_ConstValMthLines"
'GoSub Cas_Simple
GoSub Cas_Complex
'GoSub Cas_Complex1
Exit Sub
'--
Dim Nm$, ConstVal$, IsPub As Boolean
Dim IsEdt As Boolean, Cas$
Cas_Complex1:
    Cas = "Complex1"
    IsEdt = False
    Nm = "ZZ_B"
    ConstVal = TstItm(CurPjNm, CSub, Cas, "ConstVal", IsEdt)
    Ept = TstItm(CurPjNm, CSub, "Complex1", "Ept", IsEdt)
    IsPub = True
    GoTo Tst

Cas_Complex:
    IsEdt = True
    ConstVal = MdMthLines(CurMd, "ZChunk")
    StrBrw ConstVal
    Stop
    Nm = "ZZ_A"
    IsPub = True
    Ept = TstItm(CurPjNm, CSub, "Complex", "Ept", IsEdt)
    GoTo Tst
'
Cas_Simple:
    IsEdt = False
    Nm = "ZZ_A"
    ConstVal = "AAA"
    Ept = JnCrLf(Array("", _
        "Private Function ZZ_A$()", _
        "Const A_1$ = ""AAA""", _
        "", _
        "ZZ_A = A_1", _
        "End Function"))
    GoTo Tst
Tst:
    If IsEdt Then Return
    If ConstVal = "" Then Stop
    Act = ConstValMthLines(ConstVal, Nm, IsPub)
    'Brw Act: Stop
    C
    TstOk CSub, Cas
    Return
End Sub
