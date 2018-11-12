Attribute VB_Name = "MIde_Z_Pj_Cmp_Dlt"
Option Compare Database
Option Explicit
Sub PjDltMd(A As VBProject, MdNm$)
If Not PjHasMd(A, MdNm) Then Exit Sub
A.VBComponents.Remove A.VBComponents(MdNm)
End Sub

Sub MdDlt(A As CodeModule)
Dim M$, P$, Pj As VBProject
    M = MdNm(A)
    Set Pj = MdPj(A)
    P = Pj.Name
Debug.Print FmtQQ("MdDlt: Before Md(?) is deleted from Pj(?)", M, P)
A.Parent.Collection.Remove A.Parent
Debug.Print FmtQQ("MdDlt: After Md(?) is deleted from Pj(?)", M, P)
End Sub
Sub MdRmv(A As CodeModule)
Dim C As VBComponent: Set C = A.Parent
C.Collection.Remove C
End Sub

Sub PjRmvMdNmPfx(A As VBProject, Pfx$)
Dim I
For Each I In PjMdAy(A, WhMd(Nm:=WhNm("^" & Pfx)))
    MdRmvNmPfx CvMd(I), Pfx
Next
End Sub

Sub PjRmvMdPfx(A As VBProject, B As WhMd, MdPfx$)
Dim Md As CodeModule, M
For Each M In AyNz(PjMdAy(A, B))
    Set Md = M
    Md.Parent.Name = RmvPfx(MdNm(A), MdPfx)
Next
End Sub

