Attribute VB_Name = "MDao_Att_Op_Exp"
Option Explicit
Option Compare Database
Const CMod$ = "MDao_Att_Op_Exp."

Function AttExp$(A$, ToFfn$)
'Exporting the only file in Att & Return ToFfn
AttExp = DbAttExp(CurDb, A, ToFfn)
End Function

Function AttExpFfn$(A$, AttFn$, ToFfn$)
AttExpFfn = DbAttExpFfn(CurDb, A, AttFn, ToFfn)
End Function

Function AttRsExp$(A As AttRs, ToFfn)
'Export the only File in {AttRs} {ToFfn}
Dim Fn$, Ext$, T$, F2 As DAO.Field2
With A.AttRs
    If FfnExt(!FileName) <> FfnExt(ToFfn) Then Stop
    Set F2 = !FileData
End With
F2.SaveToFile ToFfn
AttRsExp = ToFfn
End Function

Function DbAttExp$(A As Database, Att, ToFfn)
'Exporting the first File in Att.
'If no or more than one file in att, error
'If any, export and return ToFfn
Const CSub$ = CMod & "DbAttExp"
Dim N%
N = DbAttFilCnt(A, Att)
If N <> 1 Then
    ErWh CSub, "AttNm in Db should have a filecount of 1.  Cannot export.", _
        "AttNm Db FileCount ExpToFile", _
        Att, DbNm(A), N, ToFfn
End If
DbAttExp = AttRsExp(DbAttRs(A, Att), ToFfn)
MsgWh CSub, "Att is exported", "Att ToFfn FmDb", Att, ToFfn, DbNm(A)
End Function

Function DbAttExpFfn$(A As Database, Att$, AttFn$, ToFfn)
Const CSub$ = CMod & "DbAttExpFfn"
If FfnExt(AttFn) <> FfnExt(ToFfn) Then
    Er CSub, "AttFn & ToFfn are dif extension|" & _
        "To export an AttFn to ToFfn, their file extension should be same", _
        "AttFn-Ext ToFfn-Ext Db AttNm AttFn ToFfn", _
        FfnExt(AttFn), FfnExt(ToFfn), DbNm(A), Att, AttFn, ToFfn
End If
If FfnIsExist(ToFfn) Then
    Er CSub, "ToFfn exist, no over write", _
        "Db AttNm AttFn ToFfn", _
        DbNm(A), Att, AttFn, ToFfn
End If
Dim Fd2 As DAO.Field2
    Set Fd2 = DbAttExpFfn1(A, Att, AttFn$)

If IsNothing(Fd2) Then
    Er CSub, "In record of AttNm there is no given AttFn, but only Act-AttFnAy", _
        "Db Given-AttNm Given-AttFn Act-AttFny ToFfn", _
        DbNm(A), Att, AttFn, DbAttFnAy(A, Att), ToFfn
End If
Fd2.SaveToFile ToFfn
DbAttExpFfn = ToFfn
End Function
Private Function DbAttExpFfn1(A As Database, Att, AttFn$) As DAO.Field2
With DbAttRs(A, Att)
    With .AttRs
        .MoveFirst
        While Not .EOF
            If !FileName = AttFn Then
                Set DbAttExpFfn1 = !FileData
            End If
            .MoveNext
        Wend
    End With
End With
End Function
Private Sub ZZ_DbAttExpFfn()
Dim T$
T = TmpFx
DbAttExpFfn CurDb, "Tp", "TaxRateAlert(Template).xlsm", T
Debug.Assert FfnIsExist(T)
Kill T
End Sub
