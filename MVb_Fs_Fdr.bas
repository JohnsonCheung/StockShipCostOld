Attribute VB_Name = "MVb_Fs_Fdr"
Option Explicit
Option Compare Database

Sub FdrAss(A)
Const C$ = "\/:<>"
If HasChrList(A, C) Then Er CSub, "Fdr cannot has these char " & C, "Fdr Char", A, C
End Sub
