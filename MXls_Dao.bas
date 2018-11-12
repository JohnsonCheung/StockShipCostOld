Attribute VB_Name = "MXls_Dao"
Option Compare Database
Option Explicit
Sub DbOupWb(A As Database, Wb As Workbook, OupNmSsl$)
'OupNm is used for Table-Name-@*, WsCdNm-Ws*, LoNm-Tbl*
Dim Ay$(), OupNm
Ay = SslSy(OupNmSsl)
WbVdtOupNy Wb, Ay
Dim T$
For Each OupNm In Ay
    T = "@" & OupNm
    DbtOupWb A, T, Wb, OupNm
Next
End Sub

Function DbtAtLo(A As Database, T, At As Range, Optional UseWc As Boolean) As ListObject
Dim N$, Q As QueryTable
N = TblNm_LoNm(T)
If UseWc Then
    Set Q = RgWs(At).ListObjects.Add(SourceType:=0, Source:=FbAdoCnStr(A.Name), Destination:=At).QueryTable
    With Q
        .CommandType = xlCmdTable
        .CommandText = T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = T
        .Refresh BackgroundQuery:=False
    End With
    Exit Function
End If
Set DbtAtLo = LoSetNm(RgLo(DbtRg(A, T, At)), N)
End Function

Function DbtLo(A As Database, T, At As Range) As ListObject
Set DbtLo = LoSetNm(SqLo(DbtSq(A, T), At), TblNm_LoNm(T))
End Function

Sub DbtOupWb(A As Database, T, Wb As Workbook, OupNm)
'OupNm is used for WsCdNm-Ws*, LoNm-Tbl*
Dim Ws As Worksheet
Set Ws = WbWsCd(Wb, "WsO" & OupNm)
DbtPutWs A, T, Ws
End Sub

Function DbtPutFx(A As Database, T, Fx$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As Workbook
Dim O As Workbook, Ws As Worksheet
Set O = FxWb(Fx)
Set Ws = WbWs(O, WsNm)
WsClrLo Ws
Stop ' LoNm need handle?
DbtPutWs A, T, WbWs(O, WsNm)
Set DbtPutFx = O
End Function

Sub DbtPutLo(A As Database, T, Lo As ListObject)
Dim Sq(), Drs As Drs, Rs As DAO.Recordset
Set Rs = DbtRs(A, T)
If Not AyIsEq(RsFny(Rs), LoFny(Lo)) Then
    Debug.Print "--"
    Debug.Print "Rs"
    Debug.Print "--"
    AyDmp RsFny(Rs)
    Debug.Print "--"
    Debug.Print "Lo"
    Debug.Print "--"
    AyDmp LoFny(Lo)
    Stop
End If
Sq = SqAddSngQuote(RsSq(Rs))
LoMin Lo
SqRg Sq, Lo.DataBodyRange
End Sub

Sub DbtPutWs(A As Database, T, Ws As Worksheet)
'Assume the WsCdNm is WsXXX and there will only 1 Lo with Name TblXXX
'Else stop
Dim Lo As ListObject
Set Lo = WsFstLo(Ws)

If Not HasPfx(Ws.CodeName, "WsO") Then Stop
If Ws.ListObjects.Count <> 1 Then Stop
If Mid(Lo.Name, 4) <> Mid(Ws.CodeName, 4) Then Stop
DbtPutLo A, T, Lo
End Sub

Function DbtRg(A As Database, T, At As Range) As Range
Set DbtRg = SqRg(DbtSq(A, T), At)
End Function

Function DbtRgByCn(A As Database, T, At As Range, Optional LoNm0$) As ListObject
If FstChr(T) <> "@" Then Stop
Dim LoNm$, Lo As ListObject
If LoNm0 = "" Then
    LoNm = "Tbl" & RmvFstChr(T)
Else
    LoNm = LoNm0
End If
Dim AtA1 As Range, CnStr, Ws As Worksheet
Set AtA1 = RgRC(At, 1, 1)
Set Ws = RgWs(At)
With Ws.ListObjects.Add(SourceType:=0, Source:=Array( _
        FmtQQ("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share D", A.Name) _
        , _
        "eny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Databa" _
        , _
        "se Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Je" _
        , _
        "t OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Com" _
        , _
        "pact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=" _
        , _
        "False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        ), Destination:=AtA1).QueryTable '<---- At
        .CommandType = xlCmdTable
        .CommandText = Array(T) '<-----  T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = LoNm '<------------ LoNm
        .Refresh BackgroundQuery:=False
    End With

End Function

Function DbtRplLo(A As Database, T, Lo As ListObject, Optional ReSeqSpec$) As ListObject
Set DbtRplLo = SqRplLo(DbtReSeqSq(A, T, ReSeqSpec), Lo)
End Function

Sub DbttFx(A As Database, Tny0, Fx$)
DbttWb(A, Tny0).SaveAs Fx
End Sub

Function DbttWb(A As Database, TT, Optional UseWc As Boolean) As Workbook
Dim O As Workbook
Set O = NewWb
Set DbttWb = WbAddDbtt(O, A, TT, UseWc)
WbWs(O, "Sheet1").Delete
End Function

Sub DbttWrtFx(A As Database, TT, Fx$)
DbttWb(A, TT).SaveAs Fx
End Sub


Sub TTFx(TT$, Fx$)
DbttFx CurDb, TT, Fx
End Sub

Function TblWs(T, Optional WsNm$ = "Data") As Worksheet
Set TblWs = LoWs(SqLo(TblSq(T), NewA1(WsNm)))
End Function


'
'Function TblLnkFx(T, Fx$, Optional WsNm$ = "Sheet1") As String()
'TblLnkFx = DbtLnkFx(CurDb, T, Fx, WsNm)
'End Function
'
'
'Function TblPutAt(T, At As Range) As Range
'Set TblPutAt = DbtRg(CurDb, T, At)
'End Function
'
'Function TblPutFx(T, Fx$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As Workbook
'Set TblPutFx = DbtPutFx(CurDb, T, Fx, WsNm, LoNm)
'End Function
'
'Function TblRg(T, At As Range) As Range
'Set TblRg = DbtRg(CurDb, T, At)
'End Function
'
'Sub DbLnkFx(A As Database, Fx$, WsNy0)
'Dim Ws
'For Each Ws In CvNy(WsNy0)
'   DbtLnkFx A, Fx, Ws
'Next
'End Sub
'
'Function DbInfWb(A As Database) As Workbook
'Set DbInfWb = DsWb(DbInfDs(A))
'End Function
'
'
'
'Function TTWb(TT0, Optional UseWc As Boolean) As Workbook
'Set TTWb = DbttWb(CurDb, TT0, UseWc)
'End Function
'
'Sub TTWbBrw(TT, Optional UseWc As Boolean)
'WbVis TTWb(TT, UseWc)
'End Sub
'
'Sub TtWrtFx(TT, Fx$)
'DbttWrtFx CurDb, TT, Fx
'End Sub
'
'

