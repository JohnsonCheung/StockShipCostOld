Attribute VB_Name = "MIde_Ens_CSub_Brk_Wb"
Option Compare Database
Option Explicit

Function PjCSubBrkWb(A As VBProject) As Workbook
Set PjCSubBrkWb = DsWb(ZDs(A))
End Function

Private Function ZAy1(A() As CSubBrk) As CSubBrkMd()
Dim I, M As CSubBrk
For Each I In AyNz(A)
    Set M = I
    PushObj ZAy1, M.MdBrk
Next
End Function

Private Function ZAy2(A() As CSubBrk) As CSubBrkMd()
Dim I, M As CSubBrk
For Each I In AyNz(A)
    Set M = I
    PushObjAy ZAy2, M.MthBrkAy
Next
End Function

Private Function ZDrs1(A() As CSubBrk) As Drs
Set ZDrs1 = OyPrpDrs(ZAy1(A), ZFny1)
End Function

Private Function ZDrs2(A() As CSubBrk) As Drs
Set ZDrs2 = OyPrpDrs(ZAy2(A), ZFny2)
End Function

Private Function ZDs(A As VBProject) As Ds
Dim DsNm$: DsNm = FmtQQ("CSubBrk:[PjNm=?] [PjFfn=?]", A.Name, PjFfn(A))
Set ZDs = Ds(ZDtAy(A), DsNm)
End Function

Private Function ZDtAy(A As VBProject) As Dt()
Dim Ay() As CSubBrk
    Ay = PjCSubBrkAy(A)
PushObj ZDtAy, DrsDt(ZDrs1(Ay), "MdBrk")
PushObj ZDtAy, DrsDt(ZDrs2(Ay), "MthBrk")
End Function

Private Function ZFny1() As String()
Dim X As New CSubBrkMd
ZFny1 = SslSy(X.Fldss)
End Function

Private Function ZFny2() As String()
Dim X As New CSubBrkMth
ZFny2 = SslSy(X.Fldss)
End Function

Private Sub Z_PjCSubBrkWb()
Dim A As VBProject
GoTo ZZ
ZZ:
    WbVis PjCSubBrkWb(CurPj)
End Sub

Private Sub Z()
Z_PjCSubBrkWb
End Sub
