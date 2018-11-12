Attribute VB_Name = "MXls__MinimizeLo"
Option Compare Database
Option Explicit
Sub LoMin(A As ListObject)
If FstTwoChr(A.Name) <> "T_" Then Exit Sub
Dim R1 As Range
Set R1 = A.DataBodyRange
If R1.Rows.Count >= 2 Then
    RgRR(R1, 2, R1.Rows.Count).EntireRow.Delete
End If
End Sub

Private Sub WsMinLo(A As Worksheet)
If A.CodeName = "WsIdx" Then Exit Sub
If FstTwoChr(A.CodeName) <> "Ws" Then Exit Sub
Dim L As ListObject
For Each L In A.ListObjects
    LoMin L
Next
End Sub

Function WbMinLo(A As Workbook) As Workbook
Dim Ws As Worksheet
For Each Ws In A.Sheets
    WsMinLo Ws
Next
Set WbMinLo = A
End Function

Sub FxMinLo(A)
WbClsNoSav WbSav(WbMinLo(FxWb(A)))
End Sub
