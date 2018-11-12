Attribute VB_Name = "MIde_Ens_CSub_Brk"
Option Compare Database
Option Explicit
Function PjCSubBrkAy(A As VBProject) As CSubBrk()
Dim C As VBComponent
For Each C In A.VBComponents
    PushObj PjCSubBrkAy, MdCSubBrk(CvCmp(C).CodeModule)
Next
End Function

Function MdCSubBrk(A As CodeModule) As CSubBrk
Dim O As New CSubBrk
Dim Ay() As CSubBrkMth
Ay = MdCSubBrkMthAy(A)
Set MdCSubBrk = O.Init(ZMd(A, Ay), Ay)
End Function

Private Function ZIsUsingCSub(A() As CSubBrkMth) As Boolean
Dim I, M As CSubBrkMth
For Each I In AyNz(A)
    Set M = I
    If M.IsUsingCSub Then ZIsUsingCSub = True: Exit Function
Next
End Function

Private Function ZMd(A As CodeModule, B() As CSubBrkMth) As CSubBrkMd
Set ZMd = New CSubBrkMd
Dim IsUsingCSub As Boolean, DclLy$()
    DclLy = MdDclLy(A)
    IsUsingCSub = ZIsUsingCSub(B)
With ZMd
    .IsUsingCSub = IsUsingCSub
    .MdNm = MdNm(A)
    .NewCMod = ZNewCMod(.MdNm)
    .NewLno = ZNewLno(DclLy)
    .OldLno = ZOldLno(DclLy)
    If .OldLno > 0 Then _
    .OldCMod = DclLy(.OldLno - 1)
    .NeedDlt = ZNeedDlt(.NewLno, .OldLno, IsUsingCSub)
    .NeedIns = ZNeedIns(.NewCMod, .OldCMod, IsUsingCSub)
End With
End Function

Private Function ZNeedDlt(NewLno&, OldLno&, IsUsingCSub As Boolean) As Boolean
If IsUsingCSub Then Exit Function
If NewLno = 0 Then Exit Function
If OldLno = 0 Then Exit Function
ZNeedDlt = True
End Function

Private Function ZNeedIns(NewCMod$, OldCMod$, IsUsingCSub As Boolean) As Boolean
If Not IsUsingCSub Then Exit Function
If NewCMod = OldCMod Then Exit Function
ZNeedIns = True
End Function
Private Function ZNewLno&(DclLy$())
Dim J&, L
For Each L In AyNz(DclLy)
    J = J + 1
    If Not HasPfx(L, "Option ") Then ZNewLno = J: Exit Function
Next
ZNewLno = Sz(DclLy) + 1
End Function

Private Function ZNewCMod$(MdNm$)
ZNewCMod = "Const CMod$ = """ & MdNm & "."""
End Function

Private Function ZOldLno&(DclLy$())
Dim L, J&
For Each L In AyNz(DclLy)
    J = J + 1
    If HasPfx(L, "Const CMod") Then ZOldLno = J: Exit Function
Next
End Function

Private Sub Z_MdCSubBrk()
Dim A As CodeModule, Act As CSubBrk
'
Set A = CurMd
GoTo ZZ

ZZ:
    Set Act = MdCSubBrk(A)
    Stop
End Sub

Private Sub Z_PjCSubBrk()
Dim A As VBProject, Act() As CSubBrk
'
Set A = CurPj
GoTo ZZ

ZZ:
    Act = PjCSubBrkAy(A)
    Stop
End Sub


Private Sub Z()
Z_MdCSubBrk
Z_PjCSubBrk
End Sub

