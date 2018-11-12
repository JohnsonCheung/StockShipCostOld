Attribute VB_Name = "MIde_Z_Pj_Rf"
Option Compare Database
Option Explicit

Function CurPjRfFfnAy() As String()
CurPjRfFfnAy = PjRfFfnAy(CurPj)
End Function

Function CurPjRfFmt() As String()
CurPjRfFmt = AyAlign2T(PjRfLy(CurPj))
End Function

Function CurPjRfLy() As String()
CurPjRfLy = PjRfLy(CurPj)
End Function

Sub PjCpyRf(A As VBProject, ToPj As VBProject)
PjAddRfFfnAy ToPj, PjRfFfnAy(A)
End Sub

Function EmpRfAy() As Reference()
End Function

Function CvPjRf(A) As VBIDE.Reference
Set CvPjRf = A
End Function


Function RfNy() As String()
RfNy = CurPjRfNy
End Function

Function CurPjRfNy() As String()
CurPjRfNy = PjRfNy(CurPj)
End Function

Sub PjAddRf(A As VBProject, RfNm, PjFfn)
If PjHasRf(A, RfNm) Then Exit Sub
A.References.AddFromFile PjFfn
End Sub

Sub PjRfBrw(A As VBProject)
AyBrw PjRfLy(A)
End Sub

Function PjRfCfgFfn$(A As VBProject)
PjRfCfgFfn = PjSrcPth(A) & "PjRf.Cfg"
End Function

Sub PjRfDmp(A As VBProject)
AyDmp PjRfLy(A)
End Sub

Function PjRfFfnAy(A As VBProject) As String()
PjRfFfnAy = ItrPrpSy(A.References, "FullPath")
End Function

Function PjRfLy(A As VBProject) As String()
Dim R As VBIDE.Reference
For Each R In A.References
Dim O$()
    PushI O, R.Name & " " & R.FullPath
Next
PjRfLy = AyAlign1T(O)
End Function

Function PjRfNmRfFfn$(A As VBProject, RfNm$)
PjRfNmRfFfn = PjPth(A) & RfNm & ".xlam"
End Function

Function PjRfNy(A As VBProject) As String()
PjRfNy = ItrNy(A.References)
End Function

Function RfFfn$(A As Reference)
On Error Resume Next
RfFfn = A.FullPath
End Function

Function RfLin$(A As VBIDE.Reference)
RfLin = A.Name & " " & QuoteSqBkt(A.FullPath) & " " & QuoteSqBkt(A.Description)
End Function

Function RfPth$(A As VBIDE.Reference)
On Error Resume Next
RfPth = A.FullPath
End Function

Function RfToStr$(A As VBIDE.Reference)
With A
   RfToStr = .Name & " " & RfPth(A)
End With
End Function

Sub PjAddRfFfnAy(A As VBProject, RfFfnAy$())
Dim F
For Each F In RfFfnAy
    If Not PjHasRfFfn(A, CStr(F)) Then
        A.References.AddFromFile F
    End If
Next
End Sub

Function PjHasRf(A As VBProject, RfNm)
Dim RF As VBIDE.Reference
For Each RF In A.References
    If RF.Name = RfNm Then PjHasRf = True: Exit Function
Next
End Function

Function PjHasRfFfn(A As VBProject, RfFfn) As Boolean
Dim R As Reference
For Each R In A.References
    If R.FullPath = RfFfn Then PjHasRfFfn = True: Exit Function
Next
End Function

Function PjHasRfNm(A As VBProject, RfNm$) As Boolean
Dim I, R As Reference
For Each I In A.References
    Set R = I
    If R.Name = RfNm Then PjHasRfNm = True: Exit Function
Next
End Function

Sub PjImpRf(A As VBProject, RfCfgPth$)
Dim B As Dictionary: Set B = FtDic(RfCfgPth & "PjRf.Cfg")
Dim K
For Each K In B.Keys
    PjAddRf A, K, B(K)
Next
End Sub
