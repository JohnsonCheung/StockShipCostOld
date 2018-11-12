Attribute VB_Name = "MDao_Att_Op_Imp"
Option Explicit
Option Compare Database
Const CMod$ = "MDao_Att_Op_Imp."

Sub AttImp(A$, FmFfn$)
DbAttImp CurDb, A, FmFfn
End Sub

Private Sub ZImp(A As AttRs, Ffn$)
Const CSub$ = CMod & "ZImp"
Dim F2 As Field2
Dim S&, T$
S = FfnSz(Ffn)
T = FfnDTim(Ffn)
Msg CSub, "[Att] is going to import [Ffn] with [Sz] and [Tim]", FdVal(A.TblRs!AttNm), Ffn, S, T
With A
    .TblRs.Edit
    With .AttRs
        If RsHasFldV(A.AttRs, "FileName", FfnFn(Ffn)) Then
            MsgDmp "Ffn is found in Att and it is replaced"
            .Edit
        Else
            MsgDmp "Ffn is not found in Att and it is imported"
            .AddNew
        End If
        Set F2 = !FileData
        F2.LoadFromFile Ffn
        .Update
    End With
    .TblRs.Fields!FilTim = FfnTim(Ffn)
    .TblRs.Fields!FilSz = FfnSz(Ffn)
    .TblRs.Update
End With
End Sub

Sub DbAttImp(A As Database, Att$, FmFfn$)
ZImp DbAttRs(A, Att), FmFfn
End Sub

Private Sub Z_AttImp()
Dim T$
T = TmpFt
StrWrt "sdfdf", T
AttImp "AA", T
Kill T
'T = TmpFt
'AttExpFfn "AA", T
'FtBrw T
End Sub
