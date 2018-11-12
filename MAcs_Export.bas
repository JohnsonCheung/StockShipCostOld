Attribute VB_Name = "MAcs_Export"
Option Explicit
Option Compare Database

Sub AcsExpFrm(A As Access.Application)
Dim Nm$, P$, I
P = AcsSrcPth(A)
For Each I In AcsFrmNy(A)
    SaveAsText acForm, Nm, P & Nm & ".Frm.Txt"
Next
End Sub

Function AcsFrmNy(A As Access.Application) As String()
AcsFrmNy = ItrNy(A.CodeProject.AllForms)
End Function

Function AcsSrcPth$(A As Access.Application)
AcsSrcPth = PjSrcPth(A.Vbe.ActiveVBProject)
End Function
