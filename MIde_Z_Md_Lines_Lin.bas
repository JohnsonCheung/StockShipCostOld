Attribute VB_Name = "MIde_Z_Md_Lines_Lin"
Option Compare Database
Option Explicit
Function MdLin$(A As CodeModule, L&)
MdLin = A.Lines(L, 1)
End Function
