Attribute VB_Name = "MIde_Z_Component"
Option Compare Database
Option Explicit

Function CmpIsCls(A As VBComponent) As Boolean
CmpIsCls = A.Type = vbext_ct_ClassModule
End Function

Function CmpIsClsOrStd(A As VBComponent) As Boolean
Select Case A.Type
Case vbext_ct_ClassModule, vbext_ct_StdModule: CmpIsClsOrStd = True
End Select
End Function

Function CmpPjNm$(A As VBComponent)
CmpPjNm = A.Collection.Parent.Name
End Function

Sub CmpRmv(A As VBComponent)
A.Collection.Remove A
End Sub

Function CurCmp() As VBComponent
Set CurCmp = CurMd.Parent
End Function

Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function
