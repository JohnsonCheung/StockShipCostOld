Attribute VB_Name = "MVb_Tst_TstHomClr"
Option Explicit
Option Compare Database

Sub TstHomClr() ' Rmv-Empty-Pth Rmk-Pth-As-At
PthRmvAllEmpSubFdr TstHom
Ren_PjPth_AsAt
Ren_MdPth_AsAt
Ren_MthPth_AsAt
Ren_CasPth_AsAt
End Sub
Private Sub Ren_PjPth_AsAt()

End Sub
Private Sub Ren_MdPth_AsAt()

End Sub
Private Sub Ren_MthPth_AsAt()

End Sub
Private Sub Ren_CasPth_AsAt()
Ren CasPthAy
End Sub
Private Property Get CasPthAy() As String()

End Property
Private Sub Ren(PthAy)
Dim I
For Each I In AyNz(PthAy)
    PthRenAddPfx I, "@"
Next
End Sub
