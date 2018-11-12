Attribute VB_Name = "MVb_Ay__Sy"
Option Compare Database
Option Explicit
Function CvSy(A) As String()
Select Case True
Case IsEmpty(A) Or IsMissing(A)
Case IsSy(A): CvSy = A
Case IsArray(A): CvSy = AySy(A)
Case Else: CvSy = ApSy(CStr(A))
End Select
End Function


Function SyAddAp(ParamArray Ap()) As String()
Dim Av(): Av = Ap
Dim O$(), I
For Each I In Av
    If IsStr(I) Then
        Push O, I
    Else
        PushAy O, I
    End If
Next
End Function
Function Sy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
Sy = AySy(Av)
End Function


Function SyEptStmt$(A)
Dim O$(), I
Push O, "Ept =  EmpSy"
For Each I In AyNz(A)
    Push O, FmtQQ("Push Ept, ""?""", Replace(I, """", """"""))
Next
SyEptStmt = JnCrLf(O)
End Function


Function SyShow(XX$, Sy$()) As String()
Dim O$()
Select Case Sz(Sy)
Case 0
    Push O, XX & "()"
Case 1
    Push O, XX & "(" & Sy(0) & ")"
Case Else
    Push O, XX & "("
    PushAy O, Sy
    Push O, XX & ")"
End Select
SyShow = O
End Function
Function CvNy(Ny0) As String()
Select Case True
Case IsMissing(Ny0) Or IsEmpty(Ny0)
Case IsStr(Ny0): CvNy = SslSy(Ny0)
Case IsSy(Ny0): CvNy = Ny0
Case IsArray(Ny0): CvNy = AySy(Ny0)
Case Else: Stop
End Select
End Function

