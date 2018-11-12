Attribute VB_Name = "MDao_Lnk"
Option Compare Database
Option Explicit

Function ColLnkExpFny(A$()) As String()
ColLnkExpFny = AyTakT3(A)
End Function

Sub DbDrpLnkTbl(A As Database)
DbttDrp A, DbLnkTny(A)
End Sub

Function DbLnkSpecImp(A As Database, LnkSpec$()) As String()
Dim O$(), J%, T$(), L$(), W$(), U%
LnkSpecAyAsg LnkSpec, T, L, W
U = UB(LnkSpec)
For J = 0 To U
    PushAy O, DbtChkCol(A, T(J), L(J))
Next
If Sz(O) > 0 Then DbLnkSpecImp = O: Exit Function
For J = 0 To U
    DbtImpMap A, T(J), L(J), W(J)
Next
DbLnkSpecImp = O
End Function

Function FilLin_Msg$(A$)
Dim FilNm$, Ffn$, L$
Ffn = A
FilNm = ShfT(Ffn)
If FfnIsExist(Ffn) Then Exit Function
FilLin_Msg = FmtQQ("[?] file not found [?]", FilNm, Ffn)
End Function


Function DbLnkVbly(A As Database) As String()
DbLnkVbly = AyMapPXSy(DbTny(A), "DbtLnkVbl", A)
End Function

Sub DrpLnkTbl()
CurDbDrpLnkTbl
End Sub

Function LnkColIsEq(A As LnkCol, B As LnkCol) As Boolean
With A
    If .ExtNm <> B.ExtNm Then Exit Function
    If .Ty <> B.Ty Then Exit Function
    If .Nm <> B.Nm Then Exit Function
End With
LnkColIsEq = True
End Function

Sub LnkEdt()
SpnmEdt "Lnk"
End Sub

Sub LnkExp()
SpnmExp "Lnk"
End Sub

Private Function LnkFt$()
LnkFt = SpnmFt("Lnk")
End Function

Sub LnkImp()
SpnmImp "Lnk"
End Sub

Sub LnkSpecAyAsg(A$(), OTny$(), OLnkColVblAy$(), OWhBExprAy$())
Dim U%, J%
U = UB(A)
ReDim OTny(U)
ReDim OLnkColVblAy(U)
ReDim OWhBExprAy(U)
For J = 0 To U
    LSpecAsg A(J), OTny(J), OLnkColVblAy(J), OWhBExprAy(J)
Next
End Sub

Function LSpecLnkColVbl$(A)
Dim L$
LSpecAsg A, , L
LSpecLnkColVbl = L
End Function

Private Function NewLnkSpec(LnkSpec$) As LnkSpec
Dim Cln$():   Cln = LyCln(SplitCrLf(LnkSpec))
Dim AFx() As LnkAFil
Dim AFb() As LnkAFil
Dim ASw() As LnkASw

Dim FmFx() As LnkFmFil
Dim FmFb() As LnkFmFil
Dim FmIp() As String
Dim FmSw() As LnkFmSw
Dim FmWh() As LnkFmWh
Dim FmStu() As LnkFmStu

Dim IpFx() As LnkIpFil
Dim IpFb() As LnkIpFil
Dim IpS1() As String
Dim IpWs() As LnkIpWs

Dim StEle() As LnkStEle
Dim StExt() As LnkStExt
Dim StFld() As LnkStFld
    
    FmIp = SslSy(AyWhRmvTT(Cln, "FmIp", "|")(0))
'    FmFx = NewFmFil(AyWhRmvTT(Cln, "IpFx", "|"))
'    FmSw = NewFmSw(AyWhRmvT1(Cln, "IpSw"))
'    FmFb = NewFmFil(AyWhRmvTT(Cln, "IpFb", "|"))
'    FmWh = NewFmWh(AyWhRmvT1(Cln, "FmWh"))
    IpS1 = AyWhRmvTT(Cln, "IpS1", "|")
'    IpWs = NewIpWs(AyWhRmvTT(Cln, "IpWs", "|"))
Stop

With NewLnkSpec
    .AFx = AFx
    .AFb = AFb
    .ASw = ASw
    .FmFx = FmFx
    .FmFb = FmFb
    .FmIp = FmIp
    .FmSw = FmSw
    .FmStu = FmStu
    .FmWh = FmWh
    .IpFx = IpFx
    .IpFb = IpFb
    .IpS1 = IpS1
    .IpWs = IpWs
    .StEle = StEle
    .StExt = StExt
    .StFld = StFld
End With
End Function

Sub TTLnkFb(TT$, Fb$, Optional Fbtt)
DbttLnkFb CurDb, TT, Fb$, Fbtt
End Sub

Function DbInfLnkDt(A As Database) As Dt
Dim T, Dry(), C$
For Each T In DbTny(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Set DbInfLnkDt = Dt("Lnk", CvNy("Tbl Connect"), Dry)
End Function

