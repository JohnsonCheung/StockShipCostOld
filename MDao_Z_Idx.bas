Attribute VB_Name = "MDao_Z_Idx"
Option Compare Database
Option Explicit

Function CvIdx(A) As DAO.Index
Set CvIdx = A
End Function

Function IdxFny(A As DAO.Index) As String()
If IsNothing(A) Then Exit Function
IdxFny = ItrNy(A.Fields)
End Function

Function IdxIsEq(A As DAO.Index, B As DAO.Index) As Boolean
With A
Select Case True
Case .Name <> B.Name
Case .Primary <> B.Primary
Case .Unique <> B.Unique
Case Not AyIsEq(ItrNy(.Fields), ItrNy(B.Fields))
Case Else: IdxIsEq = True
End Select
End With
End Function

Function IdxIsSk(A As DAO.Index, T) As Boolean
If A.Name <> T Then Exit Function
IdxIsSk = A.Unique
End Function

Function IdxsIsEq(A As DAO.Indexes, B As DAO.Indexes) As Boolean
If A.Count <> B.Count Then Exit Function
If Not ItrIsEqNm(A, B) Then Exit Function
Dim I
For Each I In A
    If Not IdxIsEq(CvIdx(I), B(CvIdx(I).Name)) Then Exit Function
Next
End Function

Function IdxIsUniq(A As DAO.Index) As Boolean
If IsNothing(A) Then Exit Function
IdxIsUniq = A.Unique
End Function
