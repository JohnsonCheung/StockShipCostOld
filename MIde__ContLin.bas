Attribute VB_Name = "MIde__ContLin"
Option Compare Database
Option Explicit
Const CMod$ = "MIde__ContLin."
Function MdContLin$(A As CodeModule, Lno)
Dim J&, L&
L = Lno
Dim O$: O = A.Lines(L, 1)
While LasChr(O) = "_"
    L = L + 1
    O = RmvLasChr(O) & A.Lines(L, 1)
Wend
MdContLin = O
End Function

Private Sub ZZ_SrcContLin()
Dim O$(3)
O(0) = "A _"
O(1) = "  B _"
O(2) = "C"
O(3) = "D"
Dim Act$: Act = SrcContLin(O, 0)
Ass Act = "A B C"
End Sub
Function SrcContLin$(A$(), Ix)
Const CSub$ = CMod & "SrcContLin"
If Ix <= -1 Then Exit Function
Dim J&, I$
Dim O$, IsCont As Boolean
For J = Ix To UB(A)
   I = A(J)
   O = O & LTrim(I)
   IsCont = HasSfx(O, " _")
   If IsCont Then O = RmvSfx(O, " _")
   If Not IsCont Then Exit For
Next
If IsCont Then Er CSub, "each lines {Src} ends with sfx _, which is impossible"
SrcContLin = O
End Function
