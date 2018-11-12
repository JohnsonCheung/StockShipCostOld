Attribute VB_Name = "MVb_Ay_Rmv"
Option Compare Database
Option Explicit

Function AyRmv3T(A) As String()
AyRmv3T = AyMapSy(A, "Rmv3T")
End Function

Function AyRmvFstChr(A) As String()
AyRmvFstChr = AyMapSy(A, "RmvFstChr")
End Function

Function AyRmvFstNonLetter(A) As String()
AyRmvFstNonLetter = AyMapSy(A, "RmvFstNonLetter")
End Function

Function AyRmvLasChr(A) As String()
AyRmvLasChr = AyMapSy(A, "RmvLasChr")
End Function

Function AyRmvPfx(A, Pfx) As String()
If Sz(A) = 0 Then Exit Function
Dim U&: U = UB(A)
Dim O$()
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = RmvPfx(A(J), Pfx)
Next
AyRmvPfx = O
End Function

Function AyRmvSngQRmk(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim X, O$()
For Each X In AyNz(A)
    If Not IsSngQRmk(CStr(X)) Then Push O, X
Next
AyRmvSngQRmk = O
End Function

Function AyRmvSngQuote(A$()) As String()
AyRmvSngQuote = AyMapSy(A, "RmvSngQuote")
End Function

Function AyRmvT1(A) As String()
Dim I
For Each I In AyNz(A)
    PushI AyRmvT1, RmvT1(I)
Next
End Function

Function AyRmvTT(A$()) As String()
AyRmvTT = AyMapSy(A, "RmvTT")
End Function
