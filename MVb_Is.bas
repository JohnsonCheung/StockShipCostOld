Attribute VB_Name = "MVb_Is"
Option Compare Database
Option Explicit

Function IsSqBktQuoted(A) As Boolean
IsSqBktQuoted = IsQuoted(A, "[", "]")
End Function
Function Limit(V, A, B)
Select Case V
Case V > B: Limit = B
Case V < A: Limit = A
Case Else: Limit = V
End Select
End Function
Function IsBet(V, A, B) As Boolean
If A > V Then Exit Function
If V > B Then Exit Function
IsBet = True
End Function
Function IsNBet(V, A, B) As Boolean
IsNBet = Not IsBet(V, A, B)
End Function

Function IsEmp(A) As Boolean
Select Case True
Case IsStr(A):    IsEmp = Trim(A) = ""
Case IsArray(A):  IsEmp = Sz(A) = 0
Case IsEmpty(A), IsNothing(A), IsMissing(A), IsNull(A): IsEmp = True
End Select
End Function

Function IsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
IsInAp = AyHas(Av, V)
End Function

Function IsInAy(V, Ay) As Boolean
IsInAy = AyHas(Ay, V)
End Function


Function IsInUCaseSy(A$, Sy$()) As Boolean
IsInUCaseSy = AyHas(Sy, UCase(A))
End Function
