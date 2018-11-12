Attribute VB_Name = "MDao_Att_AttRs"
Option Explicit
Option Compare Database

Function AttRs(A) As AttRs
AttRs = DbAttRs(CurDb, A)
End Function

Function AttRsAttNm$(A As AttRs)
AttRsAttNm = A.TblRs!AttNm
End Function

Function AttRsFilCnt%(A As AttRs)
AttRsFilCnt = RsNRec(A.AttRs)
End Function

Function AttRsFstFn$(A As AttRs)
With A.AttRs
    If .EOF Then
        If .BOF Then
            Msg CSub, "[AttNm] has no attachment files", AttRsAttNm(A)
            Exit Function
        End If
    End If
    .MoveFirst
    AttRsFstFn = !FileName
End With
End Function

Function DbAttRs(A As Database, Att) As AttRs
With DbAttRs
    Set .TblRs = A.OpenRecordset(FmtQQ("Select Att,FilTim,FilSz from Att where AttNm='?'", Att))
    If .TblRs.EOF Then
        A.Execute FmtQQ("Insert into Att (AttNm) values('?')", Att)
        Set .TblRs = A.OpenRecordset(FmtQQ("Select Att from Att where AttNm='?'", Att))
    End If
    Set .AttRs = .TblRs.Fields(0).Value
End With
End Function

Function DbFstAttRs(A As Database) As AttRs
With DbFstAttRs
    Set .TblRs = A.TableDefs("Att").OpenRecordset
    Set .AttRs = .TblRs.Fields("Att").Value
End With
End Function
