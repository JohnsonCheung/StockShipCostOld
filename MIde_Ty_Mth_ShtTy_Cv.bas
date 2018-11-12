Attribute VB_Name = "MIde_Ty_Mth_ShtTy_Cv"
Option Compare Database
Option Explicit

Function MthTyKd$(MthTy$)
Select Case MthTy
Case "Function", "Sub": MthTyKd = MthTy
Case "Property Get", "Property Let", "Property Set": MthTyKd = "Property"
End Select
End Function

Function IsMthTy(A$) As Boolean
IsMthTy = AyHas(MthTyAy, A)
End Function

Function IsMdy(A$) As Boolean
IsMdy = AyHas(MdyAy, A)
End Function

Function ShtMthTy$(MthTy)
Dim O$
Select Case MthTy
Case "Property Get": O = "Get"
Case "Property Set": O = "Set"
Case "Property Let": O = "Let"
Case "Function":     O = "Fun"
Case "Sub":          O = "Sub"
End Select
ShtMthTy = O
End Function

Function ShtMthKd$(MthKd)
Dim O$
Select Case MthKd
Case "Property": O = "Prp"
Case "Function": O = "Fun"
Case "Sub":      O = "Sub"
End Select
ShtMthKd = O
End Function

