Attribute VB_Name = "MIde_Ty_Mth_Ty_ConstAy"
Option Compare Database
Option Explicit

Const C_Enm$ = "Enum"
Const C_Prp$ = "Property"
Const C_Ty$ = "Type"
Const C_Fun$ = "Function"
Const C_Sub$ = "Sub"
Const C_Get$ = "Get"
Const C_Set$ = "Set"
Const C_Let$ = "Let"
Const C_Pub$ = "Public"
Const C_Prv$ = "Private"
Const C_Frd$ = "Friend"
Const C_PrpGet$ = C_Prp + " " + C_Get
Const C_PrpLet$ = C_Prp + " " + C_Let
Const C_PrpSet$ = C_Prp + " " + C_Set

Function PrpTyAy() As String()
Static X$()
If Sz(X) = 0 Then X = ApSy(C_Get, C_Set, C_Let)
PrpTyAy = X
End Function
Function MthTyAy() As String()
Static X$()
If Sz(X) = 0 Then X = ApSy(C_Fun, C_Sub, C_PrpGet, C_PrpLet, C_PrpSet)
MthTyAy = X
End Function

Function MdyAy() As String()
Static X$()
If Sz(X) = 0 Then X = ApSy(C_Pub, C_Prv, C_Frd, "")
MdyAy = X
End Function

Function ShtMdyAy() As String()
Static X$()
If Sz(X) = 0 Then X = ApSy("Pub", "Prv", "Frd", "")
ShtMdyAy = X
End Function
Function ShtMthKdAy() As String()
Static X$()
If Sz(X) = 0 Then X = ApSy("Fun", "Sub", "Prp")
ShtMthKdAy = X
End Function
Function ShtMthTyAy() As String()
Static X$()
If Sz(X) = 0 Then X = ApSy("Fun", "Sub", "Get", "Set", "Let")
ShtMthTyAy = X
End Function


Function MthKdAy() As String()
Static X$()
If Sz(X) = 0 Then X = ApSy(C_Fun, C_Sub, C_Prp)
MthKdAy = X
End Function

Function DclItmAy() As String()
Static X$()
If Sz(X) = 0 Then X = ApSy(C_Ty, C_Enm)
DclItmAy = X
End Function
