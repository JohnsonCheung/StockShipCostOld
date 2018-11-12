Attribute VB_Name = "MDao_Att_Inf"
Option Explicit
Option Compare Database
Const CMod$ = "MDao_Att_Inf."

Function AttFfn$(A)
'Return Fst-Ffn-of-Att-A
AttFfn = RsMovFst(AttRs(A).AttRs)!FileName
End Function

Function AttFilCnt%(A)
AttFilCnt = DbAttFilCnt(CurDb, A)
End Function

Function AttFnAy(A) As String()
AttFnAy = DbAttFnAy(CurDb, "AA")
End Function

Function AttFny() As String()
AttFny = ItrNy(DbFstAttRs(CurDb).AttRs.Fields)
End Function

Function AttFstFn$(A)
AttFstFn = DbAttFstFn(CurDb, A)
End Function

Function AttHasOnlyOneFile(A$) As Boolean
AttHasOnlyOneFile = DbAttHasOnlyOneFile(CurDb, A)
End Function

Function AttIsOld(A$, Ffn$) As Boolean
Const CSub$ = CMod & "AttIsOld"
Dim TAtt As Date, TFfn As Date, AttIs$
TAtt = AttTim(A)
TFfn = FfnTim(Ffn)
AttIs = IIf(TAtt > TFfn, "new", "old")
Dim M$
M = "Att is " & AttIs
MsgWh CSub, M, "Att Ffn AttTim FfnTim AttIs-Old-or-New?", A, Ffn, TAtt, TFfn, AttIs
End Function

Function AttNy() As String()
AttNy = CurDbAttNy
End Function

Function AttSz(A) As Date
AttSz = TsfV("Att", A, "FilSz")
End Function

Function AttTim(A$) As Date
AttTim = TfkV("Att", "FilTim", A)
End Function

Function CurDbAttNy() As String()
CurDbAttNy = DbAttNy(CurDb)
End Function

Function DbAttFilCnt%(A As Database, Att)
'DbAttFilCnt = DbAttRs(A, Att).AttRs.RecordCount
DbAttFilCnt = AttRsFilCnt(DbAttRs(A, Att))
End Function

Function DbAttFnAy(A As Database, Att$) As String()
Dim T As DAO.Recordset ' AttTblRs
Dim F As DAO.Recordset ' AttFldRs
Set T = DbAttTblRs(A, Att)
If T.EOF And T.BOF Then Exit Function
Set F = T.Fields("Att").Value
DbAttFnAy = RsSy(F, "FileName")
End Function

Function DbAttFstFn(A As Database, Att)
DbAttFstFn = AttRsFstFn(DbAttRs(A, Att))
End Function

Function DbAttHasOnlyOneFile(A As Database, Att$) As Boolean
Debug.Print "DbAttHasOnlyFile: " & DbAttRs(A, Att).AttRs.RecordCount
DbAttHasOnlyOneFile = DbAttRs(A, Att).AttRs.RecordCount = 1
End Function

Function DbAttNy(A As Database) As String()
Q = "Select AttNm from Att order by AttNm": DbAttNy = RsSy(A.OpenRecordset(Q))
End Function

Function DbAttTblRs(A As Database, AttNm$) As DAO.Recordset
Set DbAttTblRs = A.OpenRecordset(FmtQQ("Select * from Att where AttNm='?'", AttNm))
End Function

Private Sub Z_AttFnAy()
FbCurDb SampleFb_ShpRate
D AttFnAy("AA")
ClsCurDb
End Sub
