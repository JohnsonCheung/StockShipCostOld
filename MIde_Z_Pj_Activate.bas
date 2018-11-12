Attribute VB_Name = "MIde_Z_Pj_Activate"
Option Compare Database
Option Explicit

Sub PjAct(A As VBProject)
Set CurVbe.ActiveVBProject = A
End Sub

Sub ActPj()
PjAct CurPj
End Sub

Sub PjNmAct(PjNm$)
Set CurVbe.ActiveVBProject = Pj(PjNm)
End Sub
