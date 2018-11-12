Attribute VB_Name = "MVb__Var"
Option Compare Database
Option Explicit

Function CvNothing(A)
If IsEmpty(A) Then Set CvNothing = Nothing: Exit Function
Set CvNothing = A
End Function

