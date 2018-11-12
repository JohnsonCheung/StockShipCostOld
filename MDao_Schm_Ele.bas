Attribute VB_Name = "MDao_Schm_Ele"
Option Compare Database
Option Explicit
Public Const EleLblss$ = "*Ele *Ty ?Req ?AlwZLen Dft VTxt VRul TxtSz Expr"

Function FDicFmt(A As Dictionary) As String()
If IsNothing(A) Then PushI FDicFmt, "FDic is *Nothing": Exit Function
Dim K
For Each K In A.Keys
    PushI FDicFmt, K & " " & FdStr(A(K))
Next
End Function

Function EleDefFd(A) As DAO.Field2
Dim Ele$, TyStr$, Req As Boolean, AlwZLen As Boolean, Dft$, VTxt$, VRul$, TxtSz As Byte, Expr$
Dim L$: L = A
AyAsg ShfVal(L, EleLblss), _
    Ele, TyStr, Req, AlwZLen, Dft, VTxt, VRul, TxtSz, Expr
Set EleDefFd = NewFd( _
    Ele, DaoTy(TyStr), TxtSz, AlwZLen, Expr, Dft, Req, VRul, VTxt)
End Function

Private Sub Z_EleDefFd()
Dim A$, Act As DAO.Field2, Ept As DAO.Field2
A = "Int Req AlwZLen Dft=ABC TxtSz=10"
Set Ept = New DAO.Field
With Ept
    .Type = DAO.DataTypeEnum.dbInteger
'    .AllowZeroLength = True
    .DefaultValue = "ABC"
    .Required = True
    .Size = 10
End With
GoSub Tst
Exit Sub
Tst:
    Set Act = EleDefFd(A)
    If Not FdIsEq(Act, Ept) Then Stop
    Return
End Sub

Private Sub Z()
Z_EleDefFd
End Sub

