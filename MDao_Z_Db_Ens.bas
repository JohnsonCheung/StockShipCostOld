Attribute VB_Name = "MDao_Z_Db_Ens"
Option Explicit
Option Compare Database

Sub TdEns(A As DAO.TableDef)
DbTdEns CurDb, A
End Sub

Sub DbSchmEns(A As DAO.TableDef, SchmLines$)
DbTdEns CurDb, A
End Sub

Sub DbTdEns(A As Database, B As DAO.TableDef)
If HasTbl(B.Name) Then
    TdIsEqAss CurrentDb.TableDefs(B.Name), B
Else
    CurrentDb.TableDefs.Append B
End If
End Sub



