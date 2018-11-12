Attribute VB_Name = "MDao_Att_Tbl"
Option Explicit
Option Compare Database
Private Const Nm$ = "!Att"
Function AttTd() As DAO.TableDef
Static X As DAO.TableDef
If IsNothing(X) Then
    ZEns
    Set X = CurDb.TableDefs(Nm)
End If
Set AttTd = X
End Function

Private Sub ZEns()
Select Case True
Case Not ZExist: ZCrt
Case Not ZSam:  Er CSub, "Att structure not expected", "Dif", ZDif
End Select
End Sub
Private Function ZDif() As String()

End Function
Private Function ZExist() As Boolean

End Function
Private Function ZSam() As Boolean
ZSam = MDao_Z_Td.TdIsEq(ZTd, CurrentDb.TableDefs("Att"))
End Function
Private Sub ZCrt()
CurrentDb.TableDefs.Append ZTd
End Sub
Private Function ZTd() As DAO.TableDef
Set ZTd = NewTd("Att", ZFdAy)
End Function
Private Function ZFdAy() As DAO.Field2()
ZFdAy = AyAddAp(ZFdAttNm, ZFdAtt, ZFdFilSz, ZFdFilTim)
End Function
Private Function ZFdAttNm() As DAO.Field2

End Function
Private Function ZFdAtt() As DAO.Field2

End Function
Private Function ZFdFilSz() As DAO.Field2

End Function
Private Function ZFdFilTim() As DAO.Field2

End Function

