Attribute VB_Name = "MVb_Fs_Pth_R"
Option Explicit
Option Compare Database
Private O$(), A_Spec$, A_Atr As FileAttribute ' Used in PthPthAyR/PthFfnAyR

Function PthEmpPthAyR(A) As String()
Dim I
For Each I In AyNz(PthPthAyR(A))
    If PthIsEmp(I) Then PushI PthEmpPthAyR, I
Next
End Function

Function PthEntAyR(A, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute) As String()
Erase O
A_Spec = FilSpec
A_Atr = Atr
PthEntAyR1 A
PthEntAyR = O
End Function

Private Function PthEntAyR1(A, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute) As String()
Dim P$(): P = PthPthAy(A, A_Spec, A_Atr)
If Sz(P) = 0 Then Exit Function
If Sz(O) Mod 1000 = 0 Then Debug.Print "PthPthAyR1: (Each 1000): " & A
Dim I
For Each I In P
    PushI O, I
    PushIAy O, PthFfnAy(A, A_Spec, A_Atr)
    PthFfnAyR1 I
Next
End Function

Function PthFfnAyR(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Erase O
A_Spec = Spec
A_Atr = Atr
PthFfnAyR1 A
PthFfnAyR = O
End Function

Private Sub PthFfnAyR1(A)
Dim P$(): P = PthPthAy(A, A_Spec, A_Atr)
If Sz(P) = 0 Then Exit Sub
If Sz(O) Mod 1000 = 0 Then Debug.Print "PthPthAyR1: (Each 1000): " & A
Dim I
For Each I In P
    PushIAy O, PthFfnAy(A, A_Spec, A_Atr)
    PthFfnAyR1 I
Next
End Sub

Function PthPthAyR(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Erase O
A_Spec = Spec
A_Atr = Atr
PthPthAyR1 A
PthPthAyR = O
End Function

Private Sub PthPthAyR1(A)
Dim P$(): P = PthPthAy(A, A_Spec, A_Atr)
If Sz(P) = 0 Then Exit Sub
If Sz(O) Mod 1000 = 0 Then Debug.Print "PthPthAyR1: (Each 1000): " & A
PushIAy O, P
Dim I
For Each I In P
    PthPthAyR1 I
Next
End Sub

Private Sub ZZ_PthEntAyR()
Dim A$(): A = PthEntAyR("C:\users\user\documents\")
Debug.Print Sz(A)
Stop
AyDmp A
End Sub

Private Sub Z_PthEntAyR()
Dim A$(): A = PthEntAyR("C:\users\user\documents\")
Debug.Print Sz(A)
Stop
AyDmp A
End Sub

Sub Z_PthFfnAyR()
D PthFfnAyR("C:\Users\User\Documents\WindowsPowershell\")
End Sub

Private Function AA$()
Const A_1$ = "kdjlf slkdfj skldjf" & _
vbCrLf & "sdflk jsd" & _
vbCrLf & "slkdfj"

AA = A_1
End Function

