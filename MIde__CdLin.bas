Attribute VB_Name = "MIde__CdLin"
Option Compare Database
Option Explicit
Function IsCdLin(A) As Boolean
Dim L$: L = Trim(A)
If A = "" Then Exit Function
If Left(A, 1) = "'" Then Exit Function
IsCdLin = True
End Function
Function AyWhCdLin(A) As String()
Dim L
For Each L In AyNz(A)
    If IsCdLin(L) Then
        PushI AyWhCdLin, L
    End If
Next
End Function
