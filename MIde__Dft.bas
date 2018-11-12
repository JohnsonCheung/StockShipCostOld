Attribute VB_Name = "MIde__Dft"
Option Compare Database
Option Explicit
Function DftMdByNm(MdNm$) As CodeModule
If MdNm = "" Then
    Set DftMdByNm = CurMd
Else
    Set DftMdByNm = Md(MdNm)
End If
End Function
Function DftMd(A As CodeModule) As CodeModule
If IsNothing(A) Then
   Set DftMd = CurMd
Else
   Set DftMd = A
End If
End Function

Function DftPj(A As VBProject) As VBProject
If IsNothing(A) Then
   Set DftPj = CurPj
Else
   Set DftPj = A
End If
End Function


Function DftMdyAy(A$) As String()
DftMdyAy = CvNy(A)
End Function

Function DftMth(MthDNm0$) As Mth
If MthDNm0 = "" Then
    Set DftMth = CurMth
    Exit Function
End If
Set DftMth = DDNmMth(MthDNm0)
End Function

Function DftPjByNm(PjNm$) As VBProject
If PjNm = "" Then
    Set DftPjByNm = CurPj
Else
    Set DftPjByNm = Pj(PjNm)
End If
End Function

Function DftFunByDDNm(MthDDNm0$) As Mth
If MthDDNm0 = "" Then
    Dim M As Mth
    Set M = CurMth
    If MthIsFun(M) Then
        Set DftFunByDDNm = M
    End If
Else
End If
Stop '
End Function
Function DftCmpTyAy(A) As vbext_ComponentType()
If IsLngAy(A) Then DftCmpTyAy = A
End Function

Private Sub Z_DftCmpTyAy()
Dim X() As vbext_ComponentType
DftCmpTyAy (X)
Stop
End Sub


Function DftMthNm$(MthNm0$)
If MthNm0 = "" Then
    DftMthNm = CurMthNm
    Exit Function
End If
DftMthNm = MthNm0
End Function



Private Sub Z()
Z_DftCmpTyAy
End Sub
