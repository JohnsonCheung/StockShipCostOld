Attribute VB_Name = "MDao_Z_Fds"
Option Compare Database
Option Explicit
Function FdsIsEq(A As DAO.Fields, B As DAO.Fields) As Boolean
Stop '
End Function

Sub FdsSetSqRow(A As DAO.Fields, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
DrSetSqRow FdsDr(A), Sq, R, NoTxtSngQ
End Sub

Function FdsCsv$(A As DAO.Fields)
FdsCsv = AyCsv(ItrVy(A))
End Function

Function FdsDr(A As DAO.Fields) As Variant()
Dim F As DAO.Field
For Each F In A
    PushI FdsDr, F.Value
Next
End Function

Function FdsFny(A As Fields) As String()
FdsFny = ItrNy(A)
End Function

Function FdsHasFld(A As DAO.Fields, F) As Boolean
FdsHasFld = ItrHasNm(A, F)
End Function

Function FdsKyDr(A As DAO.Fields, Ky0) As Variant()
Dim O(), K
For Each K In CvNy(Ky0)
    Push FdsKyDr, A(K).Value
Next
FdsKyDr = O
End Function

Function FdsVy(A As DAO.Fields, Optional Ky0) As Variant()
Select Case True
Case IsMissing(Ky0): FdsVy = ItrVy(A)
Case IsStr(Ky0):     FdsVy = FdsVyByKy(A, SslSy(Ky0))
Case IsSy(Ky0):      FdsVy = FdsVyByKy(A, CvSy(Ky0))
Case Else:           Stop
End Select
End Function

Function FdsVyByKy(A As DAO.Fields, Ky$()) As Variant()
Dim O(), J%, K
If Sz(Ky) = 0 Then
    FdsVyByKy = ItrVy(A)
    Exit Function
End If
ReDim O(UB(Ky))
For Each K In Ky
    O(J) = A(K).Value
    J = J + 1
Next
FdsVyByKy = O
End Function

Private Sub Z_FdsDr()
Dim Rs As DAO.Recordset, Dry()
Set Rs = FbDb(SampleFb_ShpRate).OpenRecordset("Select * from YMGRnoIR")
With Rs
    While Not .EOF
        Push Dry, FdsDr(Rs.Fields)
        .MoveNext
    Wend
    .Close
End With
Brw DryFmt(Dry)
End Sub

Private Sub Z_FdsVy()
Dim Rs As DAO.Recordset, Vy()
'Set Rs = CurDb.OpenRecordset("Select * from SkuB")
With Rs
    While Not .EOF
        Vy = FdsVy(Rs)
        Debug.Print JnComma(Vy)
        .MoveNext
    Wend
    .Close
End With
End Sub


Private Sub Z()
Z_FdsDr
Z_FdsVy
End Sub
