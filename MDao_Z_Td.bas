Attribute VB_Name = "MDao_Z_Td"
Option Compare Database
Option Explicit
Sub TdAddFdAy(A As DAO.TableDef, FdAy() As DAO.Field2)
Dim I
For Each I In FdAy
    A.Fields.Append I
Next
End Sub
Function TdTyStr$(A As DAO.TableDefAttributeEnum)
TdTyStr = A
End Function
Sub TdAddIdFld(A As DAO.TableDef)
A.Fields.Append NewIdFd(A.Name)
End Sub
Function TdFny(A As DAO.TableDef) As String()
TdFny = FdsFny(A.Fields)
End Function
Sub TdIsEqAss(A As DAO.TableDef, B As DAO.TableDef)
Dim A1$: A1 = TdDefLines(A)
Dim B1$: B1 = TdDefLines(B)
If A1 <> B1 Then Stop
End Sub
Function TdIsEq(A As DAO.TableDef, B As DAO.TableDef) As Boolean
With A
Select Case True
Case .Name <> B.Name
Case .Attributes <> B.Attributes
Case Not IdxsIsEq(.Indexes, B.Indexes)
Case Not FdsIsEq(.Fields, B.Fields)
Case Else: TdIsEq = True
End Select
End With
End Function

Sub TdAddLngFld(A As DAO.TableDef, FF)
TdAddFdAy A, ZFdAy(FF, dbLong)
End Sub

Private Function ZFdAy(FF, T As DAO.DataTypeEnum) As DAO.Field2()
Dim F
For Each F In CvNy(FF)
    PushObj ZFdAy, NewFd(F, T)
Next
End Function
Sub TdAddLngTxt(A As DAO.TableDef, FF)
TdAddFdAy A, ZFdAy(FF, dbText)
End Sub

Sub TdAddStamp(A As DAO.TableDef, F$)
A.Fields.Append NewFd(F, DAO.dbDate, Dft:="Now")
End Sub

Sub TdAddTxtFld(A As DAO.TableDef, FF0, Optional Sz As Byte = 255)
Dim F
For Each F In CvNy(FF0)
    A.Fields.Append NewFd(F, dbText, Sz)
Next
End Sub

Function TdFdScly(A As DAO.TableDef) As String()
Dim N$
N = A.Name & ";"
TdFdScly = AyAddPfx(ItrMapSy(A.Fields, "FdScl"), N)
End Function

Function TdScl$(A As DAO.TableDef)
TdScl = ApScl(A.Name, AddLbl(A.OpenRecordset.RecordCount, "NRec"), AddLbl(A.DateCreated, "CrtDte"), AddLbl(A.LastUpdated, "UpdDte"))
End Function

Function TdScly(A As DAO.TableDef) As String()
TdScly = AyAdd(Sy(TdScl(A)), TdFdScly(A))
End Function

Function TdScly_AddPfx(A) As String()
Dim O$(), U&, J&, X
U = UB(A)
If U = -1 Then Exit Function
ReDim O(U)
For Each X In AyNz(A)
    O(J) = IIf(J = 0, "Td;", "Fd;") & X
    J = J + 1
Next
TdScly_AddPfx = O
End Function
Function CvTd(A) As DAO.TableDef
Set CvTd = A
End Function
