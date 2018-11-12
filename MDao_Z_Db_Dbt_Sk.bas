Attribute VB_Name = "MDao_Z_Db_Dbt_Sk"
Option Explicit
Option Compare Database

Sub DbtAssSk(A As Database, T)
Dim SkIdx As DAO.Index, I As DAO.Index
Set SkIdx = ItrFstNm(A.TableDefs(T).Indexes, "SecondaryKey")
Select Case True
Case Not IsNothing(SkIdx)
    If Not SkIdx.Unique Then
        Er CSub, "There is SecondaryKey idx, but it is not uniq", "Db Tbl Key", DbNm(A), T, IdxFny(SkIdx)
    End If
Case Else
    Set I = DbtFstUniqIdx(A, T)
    If Not IsNothing(I) Then
        Er CSub, "No SecondaryKey, but there is uniq idx, it should name as SecondaryKey", _
        "Db Tbl UniqKeyNm UniqKeyFld", _
        DbNm(A), T, I.Name, IdxFny(I)
    End If
End Select
End Sub

Function DbtFstUniqIdx(A As Database, T) As DAO.Index
Set DbtFstUniqIdx = ItrFstPrpTrue(A.TableDefs(T).Indexes, "Unique")
End Function

