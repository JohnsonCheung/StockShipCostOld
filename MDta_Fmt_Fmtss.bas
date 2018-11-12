Attribute VB_Name = "MDta_Fmt_Fmtss"
Option Compare Database
Option Explicit
Function DryFmtss(A()) As String()
Dim W%(), Dr, O$()
W = DryWdt(A)
For Each Dr In AyNz(A)
    PushI O, DrFmtss(Dr, W)
Next
DryFmtss = O
'DrFmtss = DrFmt(A, W, " ")
End Function


Function DrFmtss$(A, W%())
Dim U%, J%
U = UB(A)
If U = -1 Then Exit Function
ReDim O$(U)
For J = 0 To U - 1
    O(J) = AlignL(A(J), W%(J))
Next
O(U) = A(U)
DrFmtss = JnSpc(O)
End Function

Function DrFmt$(Dr, Wdt%(), Optional Sep$ = " | ")
Dim UDr%
   UDr = UB(Dr)
Dim O$()
   Dim U1%: U1 = UB(Wdt)
   ReDim O(U1)
   Dim W, V
   Dim J%, V1$
   J = 0
   For Each W In Wdt
       If UDr >= J Then V = Dr(J) Else V = ""
       V1 = AlignL(V, W)
       O(J) = V1
       J = J + 1
   Next
If Sep = " | " Then
    DrFmt = "| " & Join(O, Sep) & " |"
Else
    DrFmt = Join(O, Sep)
End If
End Function

Function AyColBrkssDry(A, ColBrkss$) As Variant()
Dim Lin, Ay$()
Ay = SslSy(ColBrkss)
For Each Lin In AyNz(A)
    PushI AyColBrkssDry, LinBrkssDr(Lin, Ay)
Next
End Function

Sub DrsFmtssDmp(A As Drs)
D DrsFmtss(A)
End Sub

Function DrsFmtss(A As Drs) As String()
DrsFmtss = DryFmtss(CvAy(ItmAddAy(A.Fny, A.Dry)))
End Function

Sub DrsFmtssBrw(A As Drs)
Brw DrsFmtss(A)
End Sub

Function DryFmtssCell(A()) As Variant()
Dim Dr
For Each Dr In AyNz(A)
Stop
'    Push DryFmtssCell, DrFmtssCell(Dr) ' Fmtss(X)
Next
End Function
Function LinBrkssDr(Lin, BrkssAy$()) As String()
Dim Brk, P%, L$
L = Lin
For Each Brk In BrkssAy
    P = InStr(L, Brk)
    If P = 0 Then Exit For
    Push LinBrkssDr, Left(L, P - 1)
    L = Mid(L, P)
Next
Push LinBrkssDr, L
End Function

Function AyFmt(A, ColBrkss$) As String()
AyFmt = DryFmtss(AyColBrkssDry(A, ColBrkss))
End Function
