Attribute VB_Name = "MIde_Dcl_Const"
Option Explicit
Option Compare Database

Function ShfXConst(O) As Boolean
ShfXConst = ShfX(O, "Const")
End Function


Function MdHasConst(A As CodeModule, ConstNm$) As Boolean
Dim J%
For J = 1 To A.CountOfDeclarationLines
    If LinConstNm(A.Lines(J, 1)) = ConstNm Then MdHasConst = True: Exit Function
Next
End Function

Function LinConstNm$(A)
Dim L$: L = RmvMdy(A)
If ShfXConst(L) Then LinConstNm = TakNm(L)
End Function

