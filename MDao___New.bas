Attribute VB_Name = "MDao___New"
Option Compare Database
Option Explicit
Public Q$, CSub$
Public Const CMod$ = ""
Function LnkCol(Nm$, Ty As DAO.DataTypeEnum, ExtNm$) As LnkCol
Dim O As New LnkCol
Set LnkCol = O.Init(Nm, Ty, ExtNm)
End Function
