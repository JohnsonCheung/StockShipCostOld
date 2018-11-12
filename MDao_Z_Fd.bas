Attribute VB_Name = "MDao_Z_Fd"
Option Compare Database
Option Explicit
Function FdClone(A As DAO.Field2, FldNm) As DAO.Field2
Set FdClone = New DAO.Field
With FdClone
    .Name = FldNm
    .Type = A.Type
    .AllowZeroLength = A.AllowZeroLength
    .Attributes = A.Attributes
    .DefaultValue = A.DefaultValue
    .Expression = A.Expression
    .Required = A.Required
    .ValidationRule = A.ValidationRule
    .ValidationText = A.ValidationText
End With
End Function

Function FdEleScl$(A As DAO.Field2)
Dim Rq$, Ty$, TxtSz$, AlwZ$, Rul$, Dft$, VTxt$, Expr$, Des$
Des = AddLbl(FdDes(A), "Des")
Rq = BoolTxt(A.Required, "Req")
AlwZ = BoolTxt(A.AllowZeroLength, "AlwZ")
Ty = ShtDaoTy(A.Type)
If A.Type = DAO.DataTypeEnum.dbText Then TxtSz = BoolTxt(A.Type = dbText, "TxXTSz=" & A.Size)
Rul = AddLbl(A.ValidationText, "VTxt")
VTxt = AddLbl(A.ValidationRule, "VRul")
Expr = AddLbl(A.Expression, "Expr")
Dft = AddLbl(A.DefaultValue, "Dft")
FdEleScl = ApScl(Ty, TxtSz, Rq, AlwZ, Rul, VTxt, Dft, Expr)
End Function

Function FdScl$(A As DAO.Field2)
FdScl = A.Name & ";" & FdEleScl(A)
End Function

Function FdStr$(A As DAO.Field2)
Dim D$, R$, Z$, VTxt$, VRul, E$, S$
If A.Type = DAO.DataTypeEnum.dbText Then S = " TxtSz=" & A.Size
If A.DefaultValue <> "" Then D = " " & QuoteSqBktIfNeed("Dft=" & A.DefaultValue)
If A.Required Then R = " Req"
If A.AllowZeroLength Then Z = " AlwZLen"
If A.Expression <> "" Then E = " " & QuoteSqBktIfNeed("Expr=" & A.Expression)
If A.ValidationRule <> "" Then VRul = " " & QuoteSqBktIfNeed("VRul=" & A.ValidationRule)
If A.ValidationText <> "" Then VRul = " " & QuoteSqBktIfNeed("VTxt=" & A.ValidationText)
FdStr = A.Name & " " & ShtDaoTy(A.Type) & R & Z & S & VTxt & VRul & D & E
End Function

Function FdVal(A As DAO.Field)
FdVal = A.Value
End Function

Function FdSqlTy$(A As DAO.Field)
Stop '
End Function
Function FdIsEq(A As DAO.Field2, B As DAO.Field2) As Boolean
With A
    If .Name <> B.Name Then Exit Function
    If .Type <> B.Type Then Exit Function
    If .Required <> B.Required Then Exit Function
    If .AllowZeroLength <> B.AllowZeroLength Then Exit Function
    If .DefaultValue <> B.DefaultValue Then Exit Function
    If .ValidationRule <> B.ValidationRule Then Exit Function
    If .ValidationText <> B.ValidationText Then Exit Function
    If .Expression <> B.Expression Then Exit Function
    If .Attributes <> B.Attributes Then Exit Function
    If .Size <> B.Size Then Exit Function
End With
FdIsEq = True
End Function
Function CvFd(A) As DAO.Field
Set CvFd = A
End Function

Function CvFd2(A As DAO.Field) As DAO.Field2
Set CvFd2 = A
End Function
