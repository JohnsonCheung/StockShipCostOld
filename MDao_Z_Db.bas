Attribute VB_Name = "MDao_Z_Db"
Option Compare Database
Option Explicit
Const CMod$ = "MDao_Z_Db."
Function DbAddTd(A As Database, Td As DAO.TableDef) As DAO.TableDef
A.TableDefs.Append Td
Set DbAddTd = Td
End Function

Sub DbAddTmpTbl(A As Database)
DbAddTd CurDb, TmpTd
End Sub

Sub DbAppTd(A As Database, Td As DAO.TableDef)
A.TableDefs.Append Td
End Sub

Function DbChk(A As Database) As String()
Dim T$()
T = AySrt(DbTny(A))
DbChk = AyAlign1T(AyAddAp(DbChkPk(A, T), DbChkSk(A, T)))
End Function

Function DbChkPk(A As Database, Tny$()) As String()
DbChkPk = AyExlEmpEle(AyMapPXSy(Tny, "DbtMsgPk", A))
End Function

Function DbChkSk(A As Database, Tny$()) As String()
End Function

Function DbCnSy(A As Database) As String()
Dim T$(), S()
T = AyQuoteSqBkt(DbTny(A))
S = AyMapPX(T, "DbtCnStr", A)
DbCnSy = AyabNonEmpBLy(T, S)
End Function

Sub DbCrtQry(A As Database, Q, Sql)
If Not DbHasQry(A, Q) Then
    Dim QQ As New QueryDef
    QQ.Sql = Sql
    QQ.Name = Q
    A.QueryDefs.Append QQ
Else
    A.QueryDefs(Q).Sql = Sql
End If
End Sub

Sub DbCrtResTbl(A As Database)
DbtDrp A, "Res"
DoCmd.RunSQL "Create Table Res (ResNm Text(50), Att Attachment)"
End Sub

Sub DbCrtTbl(A As Database, T$, FldDclAy)
A.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Function DbDesy(A As Database) As String()
Dim T$(), D$()
T = DbTny(A)
DbDesy = AyExlEmpEle(AyMapPXSy(T, "DbtTblDes", A))
End Function

Sub DbDrpAllTmpTbl(A As Database)
DbttDrp A, DbTmpTny(A)
End Sub

Sub DbDrpQry(A As Database, Q)
If DbHasQry(A, Q) Then A.QueryDefs.Delete Q
End Sub

Function DbDrsNormSqy(A As Database, B As Drs, Tny$()) As String()

End Function

Function DbDs(A As Database, Optional Tny0, Optional DsNm$ = "Ds") As Ds
Dim DtAy() As Dt
    Dim U%, Tny$()
    Tny = DftTny(Tny0, A.Name)
    U = UB(Tny)
    ReDim DtAy(U)
    Dim J%
    For J = 0 To U
        Set DtAy(J) = DbtDt(A, Tny(J))
    Next
Set DbDs = Ds(DtAy, DftDbNm(DsNm, A))
End Function

Private Sub Z_DbDs()
Dim Db As Database, Tny0
Stop
ZZ1:
    Set Db = FbDb(SampleFb_Duty_Dta)
    Set Act = DbDs(Db)
    DsBrw CvDs(Act)
    Exit Sub
ZZ2:
    Tny0 = "Permit PermitD"
    Set Act = DbDs(CurDb, Tny0)
    Stop
End Sub

Sub DbEnsTmp1Tbl(A As Database)
If DbHasTbl(A, "Tmp1") Then Exit Sub
DbqRun A, "Create Table Tmp1 (AA Int, BB Text 10)"
End Sub

Function DbHasQry(A As Database, Q) As Boolean
DbHasQry = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type=5", Q))
End Function

Function DbHasTbl(A As Database, T) As Boolean
'DbHasTbl = CatHasTbl(FbCat(A.Name), T)
DbHasTbl = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type in (1,6)", T))
End Function

Function DbIsOk(A As Database) As Boolean
On Error GoTo X
DbIsOk = IsStr(A.Name)
Exit Function
X:
End Function

Sub DbKill(A As Database)
Dim F$
F = A.Name
A.Close
Kill F
End Sub

Function DbNm$(A As Database)
DbNm = ObjNm(A)
End Function

Function DbOupTny(A As Database) As String()
DbOupTny = DbqSy(A, "Select Name from MSysObjects where Name like '@*' and Type =1")
End Function

Sub DbBrw(A As Database)
FbBrw A.Name
End Sub
Function DbPth$(A As Database)
DbPth = FfnPth(A.Name)
End Function

Function DbQny(A As Database) As String()
DbQny = DbqSy(A, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Private Sub Z_DbQny()
AyDmp DbQny(CurDb)
End Sub

Function DbQryRs(A As Database, Qry) As DAO.Recordset
Set DbQryRs = A.QueryDefs(Qry).OpenRecordset
End Function

Sub DbReOpn(ODb As Database)
Dim Nm$
Nm = ODb.Name
ODb.Close
Set ODb = DAO.DBEngine.OpenDatabase(Nm)
End Sub

Sub DbResClr(A As Database, ResNm$)
A.Execute "Delete From Res where ResNm='" & ResNm & "'"
End Sub


Function DbScly(A As Database) As String()
DbScly = AySy(AyOfAy_Ay(AyMap(ItrMap(A.TableDefs, "TdScly"), "TdScly_AddPfx")))
End Function

Sub DbSetTDes(A As Database, TDes As Dictionary)
Stop '
'DbtDes(A, B.T) = B.Des
End Sub

Function DbSpecNy(A As DAO.Database) As String()
DbSpecNy = DbtfSy(A, "Spec", "SpecNm")
End Function

Sub DbSqyRun(A As Database, Sqy$())
Dim Q
For Each Q In AyNz(Sqy)
   A.Execute Q
Next
End Sub

Function DbSrcTny(A As Database) As String()
Dim S()
Dim T$()
T = AyQuoteSqBkt(DbTny(A))
S = AyMapPX(T, "DbtSrcTblNm", A)
DbSrcTny = AyabNonEmpBLy(T, S)
End Function

Function DbTmpTny(A As Database) As String()
DbTmpTny = AyWhPfx(DbTny(A), "#")
End Function

Function DbTnyADO(A As Database) As String()
DbTnyADO = FbTny(A.Name)
End Function

Function DbTny(A As Database) As String()
DbTny = DbTnyDAO(A)
End Function

Function DbTnyMSYS(A As Database) As String()
DbTnyMSYS = DbqSy(A, "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_*_Data'")
End Function

Function DbTnyDAO(A As Database) As String()
Dim T As TableDef, O$()
Dim X As DAO.TableDefAttributeEnum
X = DAO.TableDefAttributeEnum.dbHiddenObject Or DAO.TableDefAttributeEnum.dbSystemObject
For Each T In A.TableDefs
    Select Case True
    Case T.Attributes And X
    Case Else
        PushI DbTnyDAO, T.Name
    End Select
Next
End Function

Private Sub ZZ_DbQny()
AyDmp DbQny(FbDb(SampleFb_Duty_Dta))
End Sub

Private Sub Z()
Z_DbDs
Z_DbQny
End Sub
