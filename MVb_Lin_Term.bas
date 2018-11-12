Attribute VB_Name = "MVb_Lin_Term"
Option Compare Database
Option Explicit
Const CMod$ = "MVb_Lin_Term."

Function LinTermAy(A) As String()
Dim L$, J%
L = A
While L <> ""
    J = J + 1: If J > 5000 Then Stop
    PushI LinTermAy, ShfTerm(L)
Wend
End Function

Function ShfT$(O)
ShfT = ShfTerm(O)
End Function

Function ShfX(O, X$) As Boolean
If LinT1(O) = X Then
    ShfX = True
    O = RmvT1(O)
End If
End Function

Private Function ShfTerm1$(O)
Dim A$
AyAsg BrkBkt(O, "["), A, ShfTerm1, O
End Function

Function ShfTerm$(O)
Dim A$
    A = LTrim(O)
If FstChr(A) = "[" Then ShfTerm = ShfTerm1(O): Exit Function
Dim P%
    P = InStr(A, " ")
If P = 0 Then
    ShfTerm = A
    O = ""
    Exit Function
End If
ShfTerm = Left(A, P - 1)
O = LTrim(Mid(A, P + 1))
End Function

Private Sub Z_ShfT()
Dim O$, OEpt$
O = " S   DFKDF SLDF  "
OEpt = "DFKDF SLDF  "
Ept = "S"
GoSub Tst
'
O = " AA BB "
Ept = "AA"
OEpt = "BB "
GoSub Tst
'
Exit Sub
Tst:
    Act = ShfT(O)
    C
    Ass O = OEpt
    Return
End Sub

Private Sub Z()
Z_ShfT
End Sub
