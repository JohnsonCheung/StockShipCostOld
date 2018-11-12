Attribute VB_Name = "MDao_Lnk_Tbl"
Option Compare Database
Option Explicit
Function DbtLnk(A As Database, T, S$, Cn$) As String()
On Error GoTo X
Dim TT As New DAO.TableDef
DbttDrp A, T
With TT
    .Connect = Cn
    .Name = T
    .SourceTableName = S
    A.TableDefs.Append TT
End With
Exit Function
X:
Dim E$
E = Err.Description
DbtLnk = FunMsgNyApLy(CSub, "Cannot link", "Db Tbl SrcTbl CnStr Er", DbNm(A), T, S, Cn, E)
End Function

Function DbtLnkVbl$(A As Database, T)
Dim O$
O = DbtFxwLnkVbl(A, T): If O <> "" Then DbtLnkVbl = "LnkFx|" & O: Exit Function
O = DbtFbtLnkVbl(A, T): If O <> "" Then DbtLnkVbl = "LnkFb|" & O: Exit Function
DbtLnkVbl = "Lcl|" & A.Name & "|" & T
End Function

Function DbtFxwLnkVbl$(A As Database, T)
If DbtIsFxLnk(A, T) Then DbtFxwLnkVbl = DbtRawLnkVbl(A, T)
End Function

Function DbtFbtLnkVbl$(A As Database, T)
If DbtIsFbLnk(A, T) Then
    DbtFbtLnkVbl = DbtRawLnkVbl(A, T)
End If
End Function

Function DbtRawLnkVbl$(A As Database, T)
Dim Cn$, X$, Y$, Y1$
Cn = DbtCnStr(A, T)
X = TakBefOrAll(TakAft(Cn, "DATABASE="), ";")
Y = A.TableDefs(T).SourceTableName
Y1 = RmvSfx(Y, "$")
DbtRawLnkVbl = X & "|" & Y1
End Function
