Attribute VB_Name = "MVb_Ay_Is"
Option Compare Database
Option Explicit

Function AyIsAllEleEq(A) As Boolean
If Sz(A) = 0 Then AyIsAllEleEq = True: Exit Function
Dim J&
For J = 1 To UB(A)
    If A(0) <> A(J) Then Exit Function
Next
AyIsAllEleEq = True
End Function

Function AyIsAllEleHasVal(A) As Boolean
If Sz(A) = 0 Then Exit Function
Dim I
For Each I In A
    If IsEmp(I) Then Exit Function
Next
AyIsAllEleHasVal = True
End Function

Function AyIsAllEq(A) As Boolean
If Sz(A) <= 1 Then AyIsAllEq = True: Exit Function
Dim A0, J&
A0 = A(0)
For J = 2 To UB(A)
    If A0 <> A(0) Then Exit Function
Next
AyIsAllEq = True
End Function

Function AyIsAllStr(A) As Boolean
If IsSy(A) Then AyIsAllStr = True: Exit Function
Dim K
For Each K In AyNz(A)
    If Not IsStr(K) Then Exit Function
Next
AyIsAllStr = True
End Function

Sub AyIsEqAss(A, B)
If VarType(A) <> VarType(B) Then
    MsgWhStop "A & B are diff Type", "A-Ty B-Ty", TypeName(A), TypeName(B)
    Exit Sub
End If
If Not IsArray(A) Then MsgWhStop "A is not array", "A-Ty", TypeName(A)
If Not IsEqSzAy(A, B) Then MsgWhStop "Siz is dif", "A-Sz B-Sz", Sz(A), Sz(B)
Dim J&, X
For Each X In AyNz(A)
    IsEqAss X, B(J)
    J = J + 1
Next
End Sub
Function AyIsEq(A, B) As Boolean
Const CSub$ = CMod & "AyIsEq"
If VarType(A) <> VarType(B) Then Exit Function
If Not IsArray(A) Then Er CSub, "[A] is not array", A
If Not IsEqSzAy(A, B) Then Exit Function
Dim J&, X
For Each X In AyNz(A)
    If Not IsEq(X, B(J)) Then Exit Function
    J = J + 1
Next
AyIsEq = True
End Function

Function AyIsEqSz(A, B) As Boolean
AyIsEqSz = Sz(A) = Sz(B)
End Function

Sub AyIsEqSzAss(A, B)
Ass Not AyIsEqSz(A, B)
End Sub

Function AyIsLinesAy(A) As Boolean
If Not AyIsAllStr(A) Then Exit Function
Dim L
For Each L In AyNz(A)
    If HasCrLf(L) Then AyIsLinesAy = True: Exit Function
Next
End Function

Function AyIsSam(A, B) As Boolean
AyIsSam = DicIsEq(AyCntDic(A), AyCntDic(B))
End Function
