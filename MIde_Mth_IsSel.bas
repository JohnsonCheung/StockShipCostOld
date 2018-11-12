Attribute VB_Name = "MIde_Mth_IsSel"
Option Explicit
Option Compare Database

Function MdyIsSel(A$, MdyAy$()) As Boolean
If Sz(MdyAy) = 0 Then MdyIsSel = True: Exit Function
Dim Mdy
For Each Mdy In MdyAy
    If Mdy = "Public" Then
        If A = "" Then MdyIsSel = True: Exit Function
    End If
    If A = Mdy Then MdyIsSel = True: Exit Function
Next
End Function


Function MthNmBrkIsSel(MthNmBrk$(), B As WhMth) As Boolean
Select Case Sz(MthNmBrk)
Case 0: Exit Function
Case 3: MthNmBrkIsSel = Mth3NmIsSel(MthNmBrk(0), MthNmBrk(1), MthNmBrk(2), B)
Case Else: Stop
End Select
End Function

Function Mth3NmIsSel(MthNm$, Ty$, Mdy$, A As WhMth) As Boolean
If IsNothing(A) Then Mth3NmIsSel = True: Exit Function
If Not ItmIsSel(Mdy, A.InMdy) Then Exit Function
If Not ItmIsSel(MthTyKd(Ty), A.InKd) Then Exit Function
Mth3NmIsSel = IsNmSel(MthNm, A.Nm)
End Function

Function KdIsSel(Kd$, WhKd$) As Boolean
KdIsSel = ItmIsSel(Kd, CvWhKd(WhKd))
End Function
