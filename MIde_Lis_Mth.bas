Attribute VB_Name = "MIde_Lis_Mth"
Option Compare Database
Option Explicit

Sub LisMdMth(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$)
Dim Ny$(), M As WhMth
Set M = WhMth(WhMdy, WhKd, WhNm(MthPatn, MthExl))
Ny = MthDDNyWh(MdMthDDNy(CurMd), M)
D AyAddPfx(Ny, CurPjNm & ".")
End Sub

Function MdLisFny() As String()
MdLisFny = SplitSpc("PJ Md-Pfx Md Ty Lines NMth NMth-Pub NMth-Prv NTy NTy-Pub NTy-Prv NEnm NEnm-Pub NEnm-Prv")
End Function
