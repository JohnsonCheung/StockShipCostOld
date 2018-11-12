Attribute VB_Name = "MDao__DefLines"
Option Explicit
Option Compare Database

Function TdDefLines$(A As DAO.TableDef)
TdDefLines = JnCrLf(TdDefLy(A))
End Function
Function TdDefLy(A As DAO.TableDef) As String()
TdDefLy = ApSy(TdDefLy1(A), IdxsDefLy(A.Indexes), FdsDefLy(A.Fields))
End Function
Function FdsDefLy(A As DAO.Fields) As String()
Dim F As DAO.Field
For Each F In A
    PushI FdsDefLy, FdDefLin(F)
Next
End Function
Function FdDefLin$(A As DAO.Field)
FdDefLin = FdStr(A)
End Function
Private Function TdDefLy1$(A As DAO.TableDef)
With A
TdDefLy1 = FmtQQ("Tbl;?;?", .Name, TdTyStr(.Attributes))
End With
End Function
Function IdxDefLin$(A As DAO.Index)
Dim X$, F$
With A
IdxDefLin = FmtQQ("Idx;?;?;?", .Name, X, F)
End With
End Function
Function IdxsDefLy(A As DAO.Indexes) As String()
Dim I As DAO.Index
For Each I In A
    PushI IdxsDefLy, IdxDefLin(I)
Next
End Function
