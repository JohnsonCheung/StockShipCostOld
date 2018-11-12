Attribute VB_Name = "MIde_Z_Md_Is"
Option Compare Database
Option Explicit
Function MdIsStd(A As CodeModule) As Boolean
MdIsStd = A.Parent.Type = vbext_ct_StdModule
End Function
Function MdIsCls(A As CodeModule) As Boolean
MdIsCls = A.Parent.Type = vbext_ct_ClassModule
End Function

Function MdIsNoLin(A As CodeModule) As Boolean
MdIsNoLin = A.CountOfLines = 0
End Function
