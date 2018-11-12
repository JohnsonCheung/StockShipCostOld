Attribute VB_Name = "MIde_Mth_Lin_Shf_and_Tak"
Option Compare Database
Option Explicit

Function ShfItmNy(A$, ItmNy0) As Variant()
ShfItmNy = AyShfItmNy(LinTermAy(A), ItmNy0)
End Function

Function ShfMthTy$(OLin)
Dim O$: O = TakMthTy(OLin)
If O = "" Then Exit Function
ShfMthTy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Sub ShfMthTyAsg(A, OMthTy, ORst$)
AyAsg ShfMthTy(A), OMthTy, ORst
End Sub

Function ShfAs(A) As Variant()
Dim L$
L = LTrim(A)
If Left(L, 3) = "As " Then ShfAs = Array(True, LTrim(Mid(L, 4))): Exit Function
ShfAs = Array(False, A)
End Function

Function ShfMdy$(OLin)
Dim O$
O = TakMdy(OLin): If O = "" Then Exit Function
ShfMdy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function ShfMthNmBrk(OLin) As String()
Dim B$()
ReDim B$(2)
B(2) = ShfMdy(OLin)
B(1) = ShfMthTy(OLin): If B(1) = "" Then Exit Function
B(0) = ShfNm(OLin)
ShfMthNmBrk = B
End Function

Function ShfKd$(OLin)
Dim T$: T = TakMthKd(OLin)
If T = "" Then Exit Function
ShfKd = T
OLin = LTrim(RmvPfx(OLin, T))
End Function

Function ShfMthSfx$(OLin)
ShfMthSfx = ShfChr(OLin, "#!@#$%^&")
End Function

Function ShfNm$(OLin)
Dim O$: O = TakNm(OLin): If O = "" Then Exit Function
ShfNm = O
OLin = RmvPfx(OLin, O)
End Function

Function ShfRmk(A) As String()
Dim L$
L = LTrim(A)
If FstChr(L) = "'" Then
    ShfRmk = ApSy(Mid(L, 2), "")
Else
    ShfRmk = ApSy("", A)
End If
End Function

Function TakMdy$(A)
TakMdy = TakPfxAySpc(A, MdyAy)
End Function

Function TakMthKd$(A)
TakMthKd = TakPfxAySpc(A, MthKdAy)
End Function

Function TakMthTy$(A)
TakMthTy = TakPfxAySpc(A, MthTyAy)
End Function

