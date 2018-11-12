Attribute VB_Name = "MIde_Z_Pj_Cmp"
Option Compare Database
Option Explicit

Function PjClsAy(A As VBProject, Optional B As WhNm) As CodeModule()
PjClsAy = PjMdAy(A, WhMd("Cls"))
End Function

Function PjClsNy(A As VBProject, Optional B As WhNm) As String()
PjClsNy = PjCmpNy(A, WhMd("Cls", B))
End Function

Function PjCmp(A As VBProject, Nm) As VBComponent
Set PjCmp = A.VBComponents(Nm)
End Function

Function PjCmpAy(A As VBProject, Optional B As WhMd) As VBComponent()
Dim I
For Each I In AyNz(PjMdAy(A, B))
    PushObj PjCmpAy, CvMd(I).Parent
Next
End Function

Function PjCmpNy(A As VBProject, Optional B As WhMd) As String()
PjCmpNy = ItrNy(PjCmpAy(A, B))
End Function

Function PjFstMbr(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    Set PjFstMbr = Cmp.CodeModule
    Exit Function
Next
End Function

Function PjFstMd(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If Cmp.Type = vbext_ct_StdModule Then
        Set PjFstMd = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function PjHasCmp(A As VBProject, Nm) As Boolean
PjHasCmp = ItrHasNm(A.VBComponents, Nm)
End Function

Function PjHasCmpWhRe(A As VBProject, Re As RegExp) As Boolean
PjHasCmpWhRe = ItrHasNmWhRe(A.VBComponents, Re)
End Function

Function PjHasMd(A As VBProject, Nm) As Boolean
Dim T As vbext_ComponentType
If Not ItrHasNm(A.VBComponents, Nm) Then Exit Function
T = PjCmp(A, Nm).Type
If T = vbext_ct_StdModule Then PjHasMd = True: Exit Function
Debug.Print "PjHasMd: Pj(?) has Mbr(?), but it is not Md, but CmpTy(?)", PjNm(A), Nm, CmpTyStr(T)
End Function

Function PjHasNoStdClsMd(A As VBProject) As Boolean
Dim C As VBComponent
For Each C In A.VBComponents
    If C.Type = vbext_ComponentType.vbext_ct_ClassModule Then Exit Function
    If C.Type = vbext_ComponentType.vbext_ct_StdModule Then Exit Function
Next
PjHasNoStdClsMd = True
End Function

Function PjMdNy_With_TstSub(A As VBProject) As String()
Dim I As VBComponent
Dim O$()
For Each I In A.VBComponents
    If I.Type = vbext_ct_StdModule Then
        If MdHasTstSub(I.CodeModule) Then
            Push O, I.Name
        End If
    End If
Next
PjMdNy_With_TstSub = O
End Function

Function Pj_ClsNy_With_TstSub(A As VBProject) As String()
Dim I As VBComponent
Dim O$()
For Each I In A.VBComponents
    If I.Type = vbext_ct_ClassModule Then
        If MdHasTstSub(I.CodeModule) Then
            Push O, I.Name
        End If
    End If
Next
Pj_ClsNy_With_TstSub = O
End Function

Private Sub Z_PjClsNy()
AyDmp PjClsNy(CurPj)
End Sub

Private Sub Z_PjMdAy()
Dim O() As CodeModule
O = PjMdAy(CurPj)
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print MdNm(Md)
Next
End Sub

Private Sub Z_PjMdNy()
AyDmp PjMdNy(CurPj)
End Sub

Private Sub Z()
Z_PjClsNy
Z_PjMdAy
Z_PjMdNy
End Sub
