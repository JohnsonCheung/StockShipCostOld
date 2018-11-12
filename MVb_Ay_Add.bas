Attribute VB_Name = "MVb_Ay_Add"
Option Compare Database
Option Explicit
Const CMod$ = "MVb_Ay_Add."

Function AyAdd1(A)
AyAdd1 = AyAddN(A, 1)
End Function

Function AyAdd(A, B)
AyAdd = A
PushAy AyAdd, B
End Function

Function AyAddAp(Ay, ParamArray Itm_or_Ay_Ap())
Const CSub$ = CMod & "AyAddAp"
Dim Av(): Av = Itm_or_Ay_Ap
If Not IsArray(Ay) Then ErWh CSub, "Fst parameter must be array", "Fst-Pm-TyeName", TypeName(Ay)
Dim I
AyAddAp = Ay
For Each I In Av
    If IsArray(I) Then
        PushAy AyAddAp, I
    Else
        Push AyAddAp, I
    End If
Next
End Function

Function AyAddFunCol(A, FunNm$) As Variant()
Dim X
For Each X In AyNz(A)
    PushI AyAddFunCol, Array(X, Run(FunNm, X))
Next
End Function

Function AyAddItm(A, Itm)
Dim O
O = A
Push O, Itm
AyAddItm = O
End Function

Function AyAddN(A, N)
AyAddN = AyCln(A)
Dim X
For Each X In AyNz(A)
    PushI AyAddN, X + N
Next
End Function

Private Sub Z_AyAdd()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyAdd(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
AyabEqChk Exp, Act
AyabEqChk Ay1, Array(1, 2, 2, 2, 4, 5)
AyabEqChk Ay2, Array(2, 2)
End Sub


Private Sub ZZ_AyAdd()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyAdd(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
AyIsEqAss Exp, Act
AyIsEqAss Ay1, Array(1, 2, 2, 2, 4, 5)
AyIsEqAss Ay2, Array(2, 2)
End Sub

Private Sub ZZ_AyAddPfx()
Dim A, Act$(), Pfx$, Exp$()
A = Array(1, 2, 3, 4)
Pfx = "* "
Exp = ApSy("* 1", "* 2", "* 3", "* 4")
GoSub Tst
Exit Sub
Tst:
Act = AyAddPfx(A, Pfx)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub

Private Sub ZZ_AyAddPfxSfx()
Dim A, Act$(), Sfx$, Pfx$, Exp$()
A = Array(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = ApSy("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = AyAddPfxSfx(A, Pfx, Sfx)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub

Function AyTab(A) As String()
AyTab = AyAddPfx(A, vbTab)
End Function

Private Sub ZZ_AyAddSfx()
Dim A, Act$(), Sfx$, Exp$()
A = Array(1, 2, 3, 4)
Sfx = "#"
Exp = ApSy("1#", "2#", "3#", "4#")
GoSub Tst
Exit Sub
Tst:
Act = AyAddSfx(A, Sfx)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub

Private Sub Z()
Z_AyAdd
End Sub
