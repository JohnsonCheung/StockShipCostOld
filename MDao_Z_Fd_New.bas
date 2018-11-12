Attribute VB_Name = "MDao_Z_Fd_New"
Option Compare Database
Option Explicit
Function NewDefFd(FdDefLin) As DAO.Field2
Dim J%, F$, L$, T$, Ay$(), Sz%, Des$, Rq As Boolean, Ty As DAO.DataTypeEnum, AlwZ As Boolean, Dft$, VRul$, VTxt$, Expr$, Er$()
L = FdDefLin
F = ShfT(L)
T = ShfT(L)
Ty = DaoTy(T)
SclAsg L, VdtEleSclNmSsl, Rq, AlwZ, Sz, Dft, VRul, VTxt, Des, Expr
Dim O As New DAO.Field
With O
    .Name = F
    .DefaultValue = Dft
    .Required = Rq
    .Type = Ty
    If Ty = DAO.DataTypeEnum.dbText Then
        .Size = Sz
        .AllowZeroLength = AlwZ
    End If
    .ValidationRule = VRul
    .ValidationText = VTxt
End With
Set NewDefFd = O
End Function

Function NewFkFd(F) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = F
    .Type = dbLong
End With
Set NewFkFd = O
End Function

Function NewIdFd(T) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = T & "Id"
    .Type = dbLong
    .Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
    .Required = True
End With
Set NewIdFd = O
End Function

Function NewTxtFd(F, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional Req As Boolean, Optional VRul$, Optional VTxt$) As DAO.Field2
Set NewTxtFd = NewFd(F, dbText, TxtSz, ZLen, Expr, Dft, Req, VRul, VTxt)
End Function

Function NewFd(F, Optional Ty As DAO.DataTypeEnum = dbText, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional Req As Boolean, Optional VRul$, Optional VTxt$) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = F
    .Required = Req
    If Ty <> 0 Then .Type = Ty
    If Ty = dbText Then
        .Size = TxtSz
        .AllowZeroLength = ZLen
    End If
    If Expr <> "" Then
        CvFd2(O).Expression = Expr
    End If
    O.DefaultValue = Dft
End With
Set NewFd = O
End Function
