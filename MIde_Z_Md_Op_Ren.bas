Attribute VB_Name = "MIde_Z_Md_Op_Ren"
Option Compare Database
Option Explicit
Sub RenMd(NewNm$)
CurMd.Name = NewNm
End Sub


Sub MdRen(A As CodeModule, NewNm$)
Dim Nm$: Nm = MdNm(A)
If NewNm = Nm Then
    Debug.Print FmtQQ("MdRen: Given Md-[?] name and NewNm-[?] is same", Nm, NewNm)
    Exit Sub
End If
If PjHasMd(MdPj(A), NewNm) Then
    Debug.Print FmtQQ("MdRen: Md-[?] already exist.  Cannot rename from [?]", NewNm, MdNm(A))
    Exit Sub
End If
MdCmp(A).Name = NewNm
Debug.Print FmtQQ("MdRen: Md-[?] renamed to [?] <==========================", Nm, NewNm)
End Sub

Private Sub Z_MdRen()
MdRen Md("A_Rs1"), "A_Rs"
End Sub


Sub PjRenMdByPfx(A As VBProject, FmMdPfx$, ToMdPfx$)
Dim CvNy$()
Dim Ny$()
'    Ny = PjMdNy(A, "^" & FmMdPfx)
    CvNy = AyMapAsgSy(Ny, "RplPfx", FmMdPfx, ToMdPfx)
Dim MdAy() As CodeModule
    Dim MdNm
    Dim Md As CodeModule
    For Each MdNm In Ny
        Set Md = PjMd(A, CStr(MdNm))
        PushObj MdAy, Md
    Next
Dim I%, U%
    For I = 0 To UB(CvNy)
        MdRen MdAy(I), CvNy(I)
    Next
End Sub

Private Sub Z_PjRenMdByPfx()
PjRenMdByPfx CurPj, "A_", ""
End Sub


Private Sub Z()
Z_MdRen
Z_PjRenMdByPfx
End Sub
