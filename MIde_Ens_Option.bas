Attribute VB_Name = "MIde_Ens_Option"
Option Compare Database
Option Explicit
Const OptExp$ = "Option Explicit"
Const OptCmpDb$ = "Option Database"
Sub PjRmvOptCmpDbLin(A As VBProject)
Dim I
For Each I In PjMdAy(A)
   MdRmvOptCmpDb CvMd(I)
Next
End Sub


Private Sub MdRmvOptCmpDb(A As CodeModule)
Dim I%: I = MdOptCmpDbLno(A)
If I = 0 Then Exit Sub
A.DeleteLines I
Debug.Print "MdRmvOptCmpDb: Option Compare Database at line " & I & " is removed"
End Sub


Sub EnsOptCmpDb(Optional MdNm$)
MdEns DftMdByNm(MdNm), OptCmpDb
End Sub

Sub EnsOptExp(Optional MdNm$)
MdEns DftMdByNm(MdNm), OptExp
End Sub

Sub EnsPjOptCmpDb(Optional PjNm$)
PjEns DftPjByNm(PjNm), OptCmpDb
End Sub

Sub EnsPjOptExp(Optional PjNm$)
PjEns DftPjByNm(PjNm), OptExp
End Sub

Sub EnsVbeOptExp()
Dim P As VBProject
For Each P In CurVbe.VBProjects
    PjEns P, OptExp
Next
End Sub

Private Function HasXXX(A As CodeModule, XXX$) As Boolean
Dim I
For Each I In AyNz(MdDclLy(A))
   If HasPfx(I, XXX) Then HasXXX = True: Exit Function
Next
End Function
Private Sub MdRmvEmpLinBetTwoOpt(A As CodeModule)
Const C = "Option"
If Not HasPfx(A.Lines(1, 1), C) Then Exit Sub
If Trim(A.Lines(2, 1)) <> "" Then Exit Sub
If Not HasPfx(A.Lines(3, 1), C) Then Exit Sub
A.DeleteLines 2, 1
Msg CSub, "Empty line between 2 option lines is removed (Md=" & MdNm(A) & ")"
End Sub
Private Sub MdEns(A As CodeModule, XXX$)
MdRmvEmpLinBetTwoOpt A
If HasXXX(A, XXX) Then Exit Sub
A.InsertLines 1, XXX
Debug.Print MdNm(A)
End Sub

Private Function MdOptCmpDbLno%(A As CodeModule)
Dim Ay$(): Ay = MdDclLy(A)
Dim J%
For J = 0 To UB(Ay)
    If HasPfx(Ay(J), "Option Compare Database") Then MdOptCmpDbLno = J + 1: Exit Function
Next
End Function

Private Sub PjEns(A As VBProject, XXX$)
Dim M
For Each M In PjMdAy(A)
    MdEns CvMd(M), OptExp
Next
End Sub

Private Sub Z()
Z_EnsOptExp
Z_EnsPjOptExp
MIde_Ens_Option:
End Sub

Private Sub Z_EnsOptExp()
EnsOptExp
End Sub

Private Sub Z_EnsPjOptExp()
EnsPjOptExp
End Sub
