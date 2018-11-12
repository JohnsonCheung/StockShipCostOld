Attribute VB_Name = "MVb_Fs_CurDir"
Option Compare Database
Option Explicit
Function CurFnAy(Optional Spec$ = "*") As String()
CurFnAy = PthFnAy(CurDir, Spec)
End Function

Function CurSubFdrAy(Optional Spec$ = "*") As String()
CurSubFdrAy = PthFdrAy(CurDir)
End Function
