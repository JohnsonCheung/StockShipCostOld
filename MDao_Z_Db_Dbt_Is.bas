Attribute VB_Name = "MDao_Z_Db_Dbt_Is"
Option Explicit
Option Compare Database
Function DbtIsExist(A As Database, T) As Boolean
DbtIsExist = FbHasTbl(A.Name, T)
'DbtIsExist = Not A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Function DbtIsFbLnk(A As Database, T) As Boolean
DbtIsFbLnk = HasPfx(DbtCnStr(A, T), ";Database=")
End Function

Function DbtIsFxLnk(A As Database, T) As Boolean
DbtIsFxLnk = HasPfx(DbtCnStr(A, T), "Excel")
End Function

Function DbtIsSys(A As Database, T) As Boolean
DbtIsSys = A.TableDefs(T).Attributes And DAO.TableDefAttributeEnum.dbSystemObject
End Function

Function DbtIsXls(A As Database, T) As Boolean
DbtIsXls = HasPfx(DbtCnStr(A, T), "Excel")
End Function


