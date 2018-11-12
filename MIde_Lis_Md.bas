Attribute VB_Name = "MIde_Lis_Md"
Option Compare Database
Option Explicit
Function PjMdLisDt(A As VBProject, Optional B As WhMd) As Dt
Stop
End Function

Sub PjMdLisDtBrw(A As VBProject, Optional B As WhMd)
DtBrw PjMdLisDt(A, B)
End Sub

Sub PjMdLisDtDmp(A As VBProject, Optional B As WhMd)
DtDmp PjMdLisDt(A, B)
End Sub

Sub LisMd(Optional Patn$, Optional Exl$)
Dim A$()
    A = PjCmpNy(CurPj, WhMd("Std", WhNm(Patn, Exl)))
    A = AySrt(A)
    A = AyAddPfx(A, "ShwMbr """)
D A
End Sub
