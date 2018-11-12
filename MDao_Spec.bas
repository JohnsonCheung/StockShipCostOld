Attribute VB_Name = "MDao_Spec"
Option Explicit
Option Compare Database


Sub DbImpSpec(A As Database, Spnm)
Const CSub$ = CMod & "DbImpSpec"
Dim Ft$
    Ft = SpnmFt(Spnm)
    
Dim NoCur As Boolean
Dim NoLas As Boolean
Dim CurOld As Boolean
Dim CurNew As Boolean
Dim SamTim As Boolean
Dim DifSz As Boolean
Dim SamSz As Boolean
Dim DifFt As Boolean
Dim Rs As DAO.Recordset
    
    Q = FmtQQ("Select SpecNm,Ft,Lines,Tim,Sz,LdTim from Spec where SpecNm = '?'", Spnm)
    Set Rs = CurDb.OpenRecordset(Q)
    NoCur = Not FfnIsExist(Ft)
    NoLas = RsAny(Rs)
    
    Dim CurT As Date, LasT As Date 'CurTim and LasTim
    Dim CurS&, LasS&
    Dim LasFt$, LdDTim$
    CurS = FfnSz(Ft)
    CurT = FfnTim(Ft)
    If Not NoLas Then
        With Rs
            LasS = Nz(Rs!Sz, -1)
            LasT = Nz(!Tim, 0)
            LasFt = Nz(!Ft, "")
            LdDTim = DteDTim(!LdTim)
        End With
    End If
    SamTim = CurT = LasT
    CurOld = CurT < LasT
    CurNew = CurT > LasT
    SamSz = CurS = LasS
    DifSz = Not SamSz
    DifFt = Ft <> LasFt
    

Const Imported$ = "***** IMPORTED ******"
Const NoImport$ = "----- no import -----"
Const NoCur______$ = "No Ft."
Const NoLas______$ = "No Last."
Const FtDif______$ = "Ft is dif."
Const SamTimSz___$ = "Sam tim & sz."
Const SamTimDifSz$ = "Sam tim & sz. (Odd!)"
Const CurIsOld___$ = "Cur is old."
Const CurIsNew___$ = "Cur is new."
Const C$ = "|[SpecNm] [Db] [Cur-Ft] [Las-Ft] [Cur-Tim] [Las-Tim] [Cur-Sz] [Las-Sz] [Imported-Time]."

Dim Dr()
Dr = Array(Spnm, Ft, FtLines(Ft), CurT, CurS, Now)
Select Case True
Case NoCur, SamTim:
Case NoLas: DrInsRs Dr, Rs
Case DifFt, CurNew: DrUpdRs Dr, Rs
Case Else: Stop
End Select

Dim Av()
Av = Array(Spnm, DbNm(A), Ft, LasFt, CurT, LasT, CurS, LasS, LdDTim)
Select Case True
Case NoCur:            FunMsgAvLinDmp CSub, NoImport & NoCur______ & C, Av
Case NoLas:            FunMsgAvLinDmp CSub, Imported & NoLas______ & C, Av
Case DifFt:            FunMsgAvLinDmp CSub, Imported & FtDif______ & C, Av
Case SamTim And SamSz: FunMsgAvLinDmp CSub, NoImport & SamTimSz___ & C, Av
Case SamTim And DifSz: FunMsgAvLinDmp CSub, NoImport & SamTimDifSz & C, Av
Case CurOld:           FunMsgAvLinDmp CSub, NoImport & CurIsOld___ & C, Av
Case CurNew:           FunMsgAvLinDmp CSub, Imported & CurIsNew___ & C, Av
Case Else: Stop
End Select
End Sub

Function SpecPth$()
SpecPth = PthEns(CurDbPth & "Spec\")
End Function

Sub SpecPthBrw()
PthBrw SpecPth
End Sub

Sub SpecPthClr()
PthClr SpecPth
End Sub

Function SpecSchmy() As String()
SpecSchmy = SplitCrLf(SpecSchmLines)
End Function

Sub DbEnsSpecTbl(A As Database)
'If Not DbHasTbl(A, "Spec") Then DbCrtSpecTbl A
End Sub

Sub SpecCrtTbl()
'DbCrtSpecTbl CurDb
End Sub

Sub SpecEnsTbl()
DbEnsSpecTbl CurDb
End Sub

Sub SpecExp()
SpecPthClr
Dim X
For Each X In AyNz(SpecNy)
    SpnmExp X
Next
End Sub

Function SpecNy() As String()
SpecNy = DbSpecNy(CurDb)
End Function


