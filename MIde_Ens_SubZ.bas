Attribute VB_Name = "MIde_Ens_SubZ"
Option Compare Database
Option Explicit

Sub EnsPjZ()
ZPjEns CurPj
PjEnsZPrv CurPj
End Sub

Sub EnsZ()
ZMdEns CurMd 'Ens SubZ()
MdEnsZPrv CurMd
End Sub
Private Sub Z_ZLinesEpt()
Dim A As CodeModule
'
Ept = Z_ZLinesEpt__Ept1
Set A = CurMd
GoSub Tst
Exit Sub
Tst:
    Act = ZLinesEpt(A)
    'StrBrw Act:Stop
    C
    Return
End Sub
Private Function ZLinesAct(A As CodeModule)
ZLinesAct = MdMthLines(A, "Z")
End Function

Private Function ZLinesEpt1$(MdNm$, ZMthNy$())
If Sz(ZMthNy) = 0 Then Exit Function
Dim O$()
PushI O, "Private Sub Z()"
PushIAy O, AySrt(ZMthNy)
PushI O, MdNm & ":"
PushI O, "End Sub"
ZLinesEpt1 = JnCrLf(O)
End Function

Private Function ZLinesEpt$(A As CodeModule)
ZLinesEpt = ZLinesEpt1(MdNm(A), MdMthNy(A, WhMth(Nm:=WhNm("^Z_"))))
End Function

Private Sub ZMdEns(A As CodeModule)
Dim Ept$
Ept = ZLinesEpt(A)
If ZLinesAct(A) = Ept Then Exit Sub
MdMthRmv A, "Z"
If Ept <> "" Then
    MdLinesApp A, vbCrLf & Ept
End If
End Sub

Private Sub ZPjEns(A As VBProject)
Dim I
For Each I In PjMdAy(A)
    Debug.Print MdNm(CvMd(I))
    ZMdEns CvMd(I)
Next
End Sub

Private Function Z_ZLinesEpt__Ept1$()
Const A_1$ = "Private Sub Z()" & _
vbCrLf & "Z_ZLinesEpt" & _
vbCrLf & "Z_ZLinesEpt__Ept1" & _
vbCrLf & "MIde_Ens_SubZ:" & _
vbCrLf & "End Sub"

Z_ZLinesEpt__Ept1 = A_1
End Function

