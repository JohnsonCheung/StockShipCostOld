Attribute VB_Name = "MIde_Z_Pj_Cur"
Option Compare Database
Option Explicit

Function CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Function

Sub CurPjAddMd(Nm$)
PjAddMod CurPj, Nm
End Sub

Sub CurPjDltMd(MdNm$)
PjDltMd CurPj, MdNm
End Sub

Function CurPjEnsMod(MdNm$) As CodeModule
Set CurPjEnsMod = PjEnsMod(CurPj, MdNm)
End Function

Function CurPjFfnAy() As String()
'PushIAy CurPjFfnAy, AppFbAy
PushIAy CurPjFfnAy, VbePjFfnAy(CurVbe)
End Function

Function CurPjFunPfxAy() As String()
CurPjFunPfxAy = PjFunPfxAy(CurPj)
End Function

Function CurPjMdAy(Optional A As WhMd) As CodeModule()
CurPjMdAy = PjMdAy(CurPj, A)
End Function

Function CurPjMthDot(Optional MdRe As RegExp, Optional ExlMd$, Optional WhMdyAy, Optional WhMthKd0$) As String()
Stop '
'CurPjMthDot = PjMthDot(CurPj, MdRe, ExlMd, WhMdyA, WhMthKd0)
End Function

Function CurPjNm$()
CurPjNm = CurPj.Name
End Function

Function CurPjPth$()
CurPjPth = PjPth(CurPj)
End Function
