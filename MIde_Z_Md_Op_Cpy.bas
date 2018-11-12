Attribute VB_Name = "MIde_Z_Md_Op_Cpy"
Option Compare Database
Option Explicit
Sub CmpCpy(A As VBComponent, ToPj As VBProject, Optional Silent As Boolean)
Dim N$: N = A.Name
If PjHasCmp(ToPj, N) Then
    Er "CmpCpy", "[Cmp] of [Pj] already exists in [TarPj]", N, CmpPjNm(A), ToPj.Name
End If
If CmpIsCls(A) Then
    CmpCpy1 A, ToPj 'If ClassModule need to export and import due to the Public/Private class property can only the set by Export/Import
Else
    PjAddCmpLines ToPj, N, A.Type, LinesEndTrim(MdLines(A.CodeModule))
End If
If Not Silent Then Debug.Print FmtQQ("CmpCpy: Cmp(?) is copied from SrcPj(?) to TarPj(?).", A.Name, CmpPjNm(A), ToPj.Name)
End Sub
Sub MdCpy(A As CodeModule, ToPj As VBProject, Optional ShwMsg As Boolean)
CmpCpy A.Parent, ToPj, ShwMsg
End Sub
Private Sub CmpCpy1(A As VBComponent, ToPj As VBProject)
Dim T$: T = TmpFt(Fnn:=A.Name)
A.Export T
ToPj.VBComponents.Import T
Kill T
End Sub
