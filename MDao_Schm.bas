Attribute VB_Name = "MDao_Schm"
Option Compare Database
Option Explicit

Sub AAA()
Z_DbSchmCrt
End Sub

Sub DbSchmCrt(A As Database, Schm$)
Dim E$(), TdDefAy$(), B As StruBase, I
SchmAsg Schm, _
    E, TdDefAy$(), B
If Sz(E) > 0 Then
    ErWh CSub, "There is error in Schm", "Db Schm Er", DbNm(A), Schm, E
End If
For Each I In TdDefAy
    DbTdDefCrt A, I, B
Next
End Sub

Sub DbtAddFF(A As Database, T, FF, EF As EF)
Dim F
For Each F In CvNy(FF)
    DbtEFAddFld A, T, EF, F
Next
End Sub

Sub DbTdDefCrt(A As Database, TdDef, B As StruBase)
Dim Td As DAO.TableDef
Dim Des$

Dim FDesDic As Dictionary
    Dim T$: T = LinT1(TdDef)
    Dim Fny$(): Fny = TdDefFny(TdDef)
    Dim Sk$(): Sk = TdDefSk(TdDef)
    Set Td = FndTd(T, Fny, Sk, B.EF)
    Des = MayDicVal(B.TDes, T)
    Set FDesDic = FndFDes(T, Fny, B.FDes, B.TFDes)
A.TableDefs.Append Td
If Des <> "" Then DbtDes(A, Td.Name) = Des
Dim F
For Each F In FDesDic.Keys
    DbtfDes(A, Td.Name, F) = FDesDic(F)
Next
End Sub

Sub DbTdDefEns(A As Database, TdDef, B As StruBase)
ChkAss TdDefChk(Stru, B.EF)
Dim S$
S = DbtStru(A, LinT1(Stru))
If S = "" Then
    DbTdDefCrt A, Stru, B
    Exit Sub
End If
If S = Stru Then Exit Sub
DbtReStru A, Stru, B.EF
End Sub

Sub DbtEFAddFld(A As Database, T, EF As EF, F)
If DbtHasFld(A, T, F) Then Exit Sub
A.TableDefs(T).Fields.Append LookupFd(F, T, EF)
End Sub

Private Function FndFDes(T$, Fny$(), FDes As Dictionary, TFDes As Dictionary) As Dictionary
Set FndFDes = New Dictionary
End Function

Private Function FndSkSql$(Stru, T)
Dim Sk$()
Sk = SslSy(RmvT1(Replace(TakBef(Stru, "|"), "*", T)))
If Sz(Sk) = 0 Then Exit Function
FndSkSql = CrtSkSql(T, Sk)
End Function

Private Function FndTd(T, Fny$(), Sk$(), EF As EF) As DAO.TableDef
'If AyHas(Fny, T & "Id") Then OPk = CrtPkSql(T) Else OPk = ""
'OSk = FndSkSql(Stru, T)
Dim FdAy() As DAO.Field2
Dim F
For Each F In Fny
    PushObj FdAy, LookupFd(F, T, EF)  '<===
Next
Set FndTd = NewTd(T, FdAy, Sk)
End Function

Private Sub Z_DbSchmCrt()
Dim Schm$, Db As Database
Set Db = TmpDb
Schm = _
         "Tbl A *Id *Nm | *Dte AATy Loc Expr Rmk" & _
vbCrLf & "Tbl B *Id AId *Nm | *Dte" & _
vbCrLf & "Fld Txt AATy" & _
vbCrLf & "Fld Loc Loc" & _
vbCrLf & "Fld Expr Expr" & _
vbCrLf & "Fld Mem Rmk" & _
vbCrLf & "Ele Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']" & _
vbCrLf & "Ele Expr Txt [Expr=Loc & 'abc']" & _
vbCrLf & "Des Tbl     A     AA BB " & _
vbCrLf & "Des Tbl     A     CC DD " & _
vbCrLf & "Des Fld     ANm   AA BB " & _
vbCrLf & "Des Tbl.Fld A.ANm TFDes-AA-BB"
GoSub Tst
Exit Sub
Tst:
    DbSchmCrt Db, Schm
    DbBrw Db
    Stop
    Return
End Sub

Private Sub Z_DbTdDefCrt()
Dim Td As DAO.TableDef
Dim B As DAO.Field2
GoSub X_Td
GoSub Tst
Exit Sub
Tst:
    TblDrp Td.Name
    Debug.Print ObjPtr(Td)
    CurDb.TableDefs.Append Td
    Debug.Print ObjPtr(Td)
    Set B = CurDb.TableDefs("#Tmp").Fields("B")
    CurDb.TableDefs("#Tmp").Fields("B").Properties.Append CurDb.CreateProperty(C_Des, dbText, "ABC")
    Return
X_Td:
    Dim FdAy() As DAO.Field2
    Set B = NewFd("B", dbInteger)
    PushObj FdAy, NewIdFd("#Tmp")
    PushObj FdAy, NewFd("A", dbInteger)
    PushObj FdAy, B
    Set Td = NewTd("#Tmp", FdAy, "A B")
    Return
End Sub

Private Function Z_DbTdDefCrt1() As DAO.TableDef
End Function

Private Sub Z()
Z_DbSchmCrt
End Sub
