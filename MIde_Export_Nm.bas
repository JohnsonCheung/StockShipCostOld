Attribute VB_Name = "MIde_Export_Nm"
Option Compare Database
Option Explicit

Sub CurPjSrcPthBrw()
PthBrw PjSrcPth(CurPj)
End Sub

Function CurPjSrcPth$()
CurPjSrcPth = PjSrcPth(CurPj)
End Function

Function MdSrcExt$(A As CodeModule)
Dim O$
Select Case A.Parent.Type
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
MdSrcExt = O
End Function

Function MdSrcFfn$(A As CodeModule)
MdSrcFfn = PjSrcPth(MdPj(A)) & MdSrcFn(A)
End Function
