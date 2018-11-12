Attribute VB_Name = "MXls_Dao_Fb"
Option Compare Database
Option Explicit
Function FbOupTblWb(A$) As Workbook
Dim O As Workbook
Set O = NewWb
AyDoABX FbOupTny(A), "WbAddWc", O, A
ItrDo O.Connections, "WcAddWs"
WbRfh O, A
Set FbOupTblWb = O
End Function

Sub FbRplWbLo(Fb$, A As Workbook)
Dim I, Lo As ListObject, Db As Database
Set Db = FbDb(Fb)
For Each I In WbOupLoAy(A)
    Set Lo = I
    DbtRplLo Db, "@" & Mid(Lo.Name, 3), Lo
Next
Db.Close
Set Db = Nothing
End Sub

Sub FbWrtFx_zForExpOupTb(A$, Fx$)
FbOupTblWb(A).SaveAs Fx
End Sub
