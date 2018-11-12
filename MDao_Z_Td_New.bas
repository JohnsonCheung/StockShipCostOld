Attribute VB_Name = "MDao_Z_Td_New"
Option Explicit
Option Compare Database
Const CMod$ = "MDao_Z_Td_New."

Private Function CvIdxFds(A) As DAO.IndexFields
Set CvIdxFds = A
End Function

Private Function FdIsId(A As DAO.Field2, T) As Boolean
If A.Name <> T & "Id" Then Exit Function
If A.Attributes <> DAO.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If A.Type <> dbLong Then Exit Function
FdIsId = True
End Function

Function NewDefTd(TdDefLin, EleFldLikssLy$(), EleDic As Dictionary) As DAO.TableDef
Dim T$
Dim FdAy() As DAO.Field2
Dim Sk$()
NewDefTd1 TdDefLin, EleFldLikssLy, EleDic, _
    T, FdAy, Sk
Set NewDefTd = NewTd(T, FdAy, Sk)
End Function

Private Sub NewDefTd1(Lin, E$(), D As Dictionary, _
    OT$, OFdAy() As DAO.Field2, OSk$())
Dim L$, Fny$(), L1$, L2$
L = Lin
OT = ShfT(L)
L1 = Replace(L, "*", OT)
L2 = Replace(L1, "|", "")
Fny = SslSy(L2)
End Sub

Private Function NewPkIdx(T As DAO.TableDef) As DAO.Index
Const CSub$ = CMod & "NewPkIdx"
Dim O As New DAO.Index
O.Name = "PrimaryKey"
O.Primary = True
If Not ItrHasNm(T.Fields, T.Name & "Id") Then
    ErWh CSub, "Given Td does not have Id field", "Missing-Id-FldNm Td-Name Td-Fny", T.Name & "Id", T.Name, TdFny(T)
End If
CvIdxFds(O.Fields).Append NewIdFd(T.Name)
Set NewPkIdx = O
End Function

Private Function NewSkIdx(T As DAO.TableDef, Sk$()) As DAO.Index
Const CSub$ = CMod & "NewSkIdx"
Dim O As New DAO.Index
O.Name = "SecondaryKey"
O.Unique = True
If Not AyHasAy(TdFny(T), Sk) Then
    ErWh CSub, "Given Td does not contain all given-Sk", "Missing-Sk Td-Name Td-Fny Given-Sk", T.Name & "Id", AyMinus(Sk, TdFny(T)), T.Name, TdFny(T), Sk
End If
Dim I
For Each I In Sk
    CvIdxFds(O.Fields).Append NewFd(I)
Next
Set NewSkIdx = O
End Function

Function NewTd(T, FdAy() As DAO.Field2, Optional SkFny0) As DAO.TableDef
Dim O As New DAO.TableDef, F
O.Name = T
For Each F In FdAy
    O.Fields.Append F
Next
TdAddPkIdx O ' add Pk
TdAddSkIdx O, SkFny0 ' add Sk
Set NewTd = O
End Function

Private Sub TdAddPkIdx(A As DAO.TableDef)
'Any Pk Fields in A.Fields?, if no exit sub
If TdAddPkIdx1(A.Fields, A.Name) Then
    A.Indexes.Append NewPkIdx(A)
End If
End Sub

Private Function TdAddPkIdx1(A As DAO.Fields, T) As Boolean
'Return True if Fds-A has Id-Fld
Dim F As DAO.Field2
For Each F In A
    If FdIsId(F, T) Then TdAddPkIdx1 = True: Exit Function
Next
End Function

Private Sub TdAddSkIdx(A As DAO.TableDef, SkFny0)
Dim Sk$(): Sk = CvNy(SkFny0): If Sz(Sk) = 0 Then Exit Sub
A.Indexes.Append NewSkIdx(A, Sk)
End Sub

Private Function ZEle$(Fld, F$())

End Function

Private Function ZFd(Fld, F$(), E As Dictionary) As DAO.Field2
Set ZFd = ZStdFd(Fld): If Not IsNothing(ZFd) Then Exit Function
Set ZFd = E(ZEle(Fld, F))
End Function

Private Function ZFdAy(Fny$(), F$(), E As Dictionary) As DAO.Field2()
Dim Fld
For Each Fld In Fny
    PushObj ZFdAy, ZFd(Fld, F, E)
Next
End Function

Private Function ZStdFd(Fld) As DAO.Field2

End Function

Sub Z()
MDao_Z_Td_New:
End Sub
