Attribute VB_Name = "MVb_Fs_Pth_Rmv"
Option Explicit
Option Compare Database

Sub PthRmvAllEmpSubFdr(A)
Dim Ay$(), I
Lp:
    Ay = PthEmpPthAyR(A): If Sz(Ay) = 0 Then Exit Sub
    For Each I In Ay
        RmDir I
    Next
    GoTo Lp
End Sub

Sub PthRmvEmpSubDir(A$)
Dim I
For Each I In AyNz(PthPthAy(A))
   PthRmvIfEmp CStr(I)
Next
End Sub

Sub PthRmvIfEmp(A$)
If Not PthIsExist(A) Then Exit Sub
If PthIsEmp(A) Then Exit Sub
RmDir A
End Sub


