Attribute VB_Name = "MIde_X_Mth_Is"
Option Compare Database
Option Explicit
Function MthIsFun(A As Mth) As Boolean
MthIsFun = MdIsStd(A.Md)
End Function

Function MthIsExist(A As Mth) As Boolean
MthIsExist = MdHasMth(A.Md, A.Nm)
End Function
