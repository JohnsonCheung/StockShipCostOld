Attribute VB_Name = "MVb_Str_Macro"
Option Compare Database
Option Explicit
Function MacroNy(A, Optional ExlBkt As Boolean, Optional OpnBkt$ = vbOpnSqBkt) As String()
'MacroStr-A is a with ..[xx].., this sub is to return all xx
Dim Q1$, Q2$
    Q1 = OpnBkt
    Q2 = ClsBkt(OpnBkt)
If Not HasSubStr(A, Q1) Then Exit Function

Dim Ay$(): Ay = Split(A, Q1)
Dim O$(), J%
For J = 1 To UB(Ay)
    Push O, TakBef(Ay(J), Q2)
Next
If Not ExlBkt Then
    O = AyAddPfxSfx(O, Q1, Q2)
End If
MacroNy = O
End Function
