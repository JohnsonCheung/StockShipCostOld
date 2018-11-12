Attribute VB_Name = "MIde_Gen_Const_MthLinesConstVal"
Option Explicit
Option Compare Database
Const CMod$ = "MIde_Gen_Const_MthLinesConstVal."

Private Function AA$()
Const A_1$ = "Johnsonlskf;lskf" & _
vbCrLf & "ksjdflkjs dflkj" & _
vbCrLf & "lksdjf lskdfj" & _
vbCrLf & ""

AA = A_1
End Function

Function MthLinesConstVal$(MthLines)
'Return a string constant from the source code.  A reverse of [ConstValMthLines]
Dim O$, C
For Each C In AyNz(ZConstAy(MthLines))
    O = O & ZConst(C)
Next
MthLinesConstVal = O
End Function

Private Function ZConst$(C)
Dim I, O$(), A$, B$
For Each I In SplitCrLf(C)
    A = TakBetFstLas(I, """", """")
    B = Replace(A, """""", """")
    PushI O, B
Next
ZConst = JnCrLf(O)
End Function

Private Function ZConstAy(MthLines) As String()
Dim Ay$(), O$
O = MthLines
Lp:
    Ay = TakP123(O, "Const", vbCrLf & vbCrLf)
    If Sz(Ay) = 3 Then
        PushI ZConstAy, Ay(1)
        O = Ay(2)
        GoTo Lp
    End If
End Function

Sub Z_MthLinesConstVal()
Const CSub$ = CMod & "Z_MthLinesConstVal"
Dim IsEdt As Boolean, MthLines$, Cas$
GoSub Cas_Complex
GoSub Cas_Simple
Exit Sub
Cas_Complex:
    IsEdt = False
    Cas = "Complex"
    MthLines = TstItm(CurPjNm, CSub, Cas, "MthLines", IsEdt)
    Ept = TstItm(CurPjNm, CSub, Cas, "Ept", IsEdt)
    If IsEdt Then Return
    GoTo Tst
Cas_Simple:
    
    Return
Tst:
    Act = MthLinesConstVal(MthLines)
    'Brw Act: Stop
    C
    TstOk CSub, Cas
    Return
End Sub
