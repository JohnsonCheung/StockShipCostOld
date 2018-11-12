Attribute VB_Name = "MIde_Export"
Option Compare Database
Option Explicit
Const CMod$ = "MIde_Export."

Function MdExp(A As CodeModule)
MdExport A
End Function

Sub MdExport(A As CodeModule)
Dim F$: F = MdSrcFfn(A)
A.Parent.Export F
Debug.Print MdNm(A)
End Sub

Sub PjExp(A As VBProject)
PjExport A
End Sub

Sub PjExpRf(A As VBProject)
AyWrt PjRfLy(A), PjRfCfgFfn(A)
End Sub

Sub PjExpSrc(A As VBProject)
PjCpyToSrc A
PthClrFil PjSrcPth(A)
Dim Md As CodeModule, I
For Each I In PjModAy(A)
    Set Md = I
    MdExp Md
Next
End Sub

Sub PjExport(A As VBProject)
Debug.Print "PjExport: " & PjNm(A) & "-----------------------------"
Dim P$
    P = PjSrcPth(A)
    If P = "" Then
        Debug.Print FmtQQ("PjExport: Pj(?) does not have FileName", A.Name)
        Exit Sub
    End If
PthClrFil P 'Clr SrcPth ---
FfnCpyToPth A.FileName, P, OvrWrt:=True
'Export Mod
    Dim I
    For Each I In AyNz(PjMdAy(A)) ' Only Cls & Mod will be exported
        MdExport CvMd(I)  'Exp each md --
    Next
PjExpRf A
End Sub
