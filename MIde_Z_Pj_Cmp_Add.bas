Attribute VB_Name = "MIde_Z_Pj_Cmp_Add"
Option Compare Database
Option Explicit

Function PjEnsCls(A As VBProject, ClsNm$) As CodeModule
Set PjEnsCls = PjEnsCmp(A, ClsNm, vbext_ct_ClassModule)
End Function

Function PjEnsCmp(A As VBProject, Nm, Optional Ty As vbext_ComponentType = vbext_ct_StdModule) As CodeModule
If Not PjHasCmp(A, Nm) Then
    PjCrtCmp A, Nm, Ty
End If
Set PjEnsCmp = A.VBComponents(Nm).CodeModule
End Function

Function PjEnsMod(A As VBProject, MdNm) As CodeModule
Set PjEnsMod = PjEnsCmp(A, MdNm, vbext_ct_StdModule)
End Function

Function PjEnsStd(A As VBProject, StdNm$) As CodeModule
Set PjEnsStd = PjEnsCmp(A, StdNm, vbext_ct_StdModule)
End Function

Function MdAddOptExpLin(A As CodeModule) As CodeModule
A.InsertLines 1, "Option Explicit"
Set MdAddOptExpLin = A
End Function
Function PjAddMod(A As VBProject, Nm) As CodeModule
Set PjAddMod = MdAddOptExpLin(PjAddCmp(A, Nm, vbext_ct_StdModule).CodeModule)
End Function


Sub PjCrtMd(A As VBProject, MdNm$)
PjCrtCmp A, MdNm, vbext_ct_StdModule
End Sub

Function PjAddCls(A As VBProject, Nm$) As CodeModule
Set PjAddCls = MdAddOptExpLin(PjAddCmp(A, Nm, vbext_ct_ClassModule).CodeModule)
End Function


Sub PjAddClsFmPj(A As VBProject, FmPj As VBProject, ClsNy0)
Dim I, ClsNy$(), ClsAy() As CodeModule
ClsNy = CvNy(ClsNy0)
For Each I In A
    MdCpy CvMd(I), A
Next
End Sub

Function PjAddCmp(A As VBProject, Nm, Ty As vbext_ComponentType) As VBComponent
If PjHasCmp(A, Nm) Then
    Er "PjAddCmp", "[Pj] already has [Cmp]", A.Name, Nm
End If
Set PjAddCmp = A.VBComponents.Add(Ty)
PjAddCmp.Name = Nm
End Function

Function PjAddCmpLines(A As VBProject, Nm, Ty As vbext_ComponentType, Lines$)
Dim O As VBComponent
Set O = PjAddCmp(A, Nm, Ty): If IsNothing(O) Then Stop
MdLinesApp O.CodeModule, Lines
Set PjAddCmpLines = O
End Function

Sub PjAddMdPfx(A As VBProject, B As WhMd, MdPfx$)
Dim Md As CodeModule, M
For Each M In AyNz(PjMdAy(A, B))
    Set Md = M
    MdRen Md, MdPfx & MdNm(Md)
Next
End Sub

Sub PjCrtCmp(A As VBProject, Nm, Ty As vbext_ComponentType)
Dim O As VBComponent
Set O = A.VBComponents.Add(Ty)
O.Name = Nm
End Sub


Sub AddCls(Nm$)
PjAddCmp CurPj, Nm, vbext_ComponentType.vbext_ct_ClassModule
End Sub

Sub AddFun(FunNm$)
'Des: Add Empty-Fun-Mth to CurMd
MdLinesApp CurMd, FmtQQ("Function ?()|End Function", FunNm)
MdMthGo CurMd, FunNm
End Sub

Sub AddSub(SubNm$)
MdLinesApp CurMd, FmtQQ("Sub ?()|End Sub", SubNm)
MdMthGo CurMd, SubNm
End Sub


Sub AddMod(Nm$)
PjAddMod CurPj, Nm
End Sub
