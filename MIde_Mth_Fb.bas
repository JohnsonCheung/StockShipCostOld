Attribute VB_Name = "MIde_Mth_Fb"
Option Compare Database
Option Explicit
Public Const MthLocFx$ = "C:\Users\User\Desktop\Vba-Lib-1\MthLoc.xlsx"
Public Const MthFb$ = "C:\Users\User\Desktop\Vba-Lib-1\Mth.accdb"
Public Const WrkFb$ = "C:\Users\User\Desktop\Vba-Lib-1\MthWrk.accdb"
Sub EnsMthTbl()
Dim A As Drs
Set A = CurPjFfnAyMthFullDrs
DrsRplDbt A, MthDb, "Mth"
End Sub

Sub EnsMthFb()
MthFbEns MthFb
End Sub

Function MthFbEns(A$) As Database
FbEns A
Dim Db As Database
Set Db = FbDb(A)
DbSchmEns Db, MthSchm
End Function

Function MthDb() As Database
Static A As Database, B As Boolean
If Not B Then
    B = True
    Set A = FbDb(MthFb)
End If
Set MthDb = A
End Function

Sub BrwMthFb()
FbBrw MthFb
End Sub

Private Function MthSchm$()
Const A_1$ = "Ele Nm  Md Pj" & _
vbCrLf & "Ele T50 MchStr" & _
vbCrLf & "ELe T10 MthPfx" & _
vbCrLf & "Ele Txt PjFfn Prm Ret LinRmk" & _
vbCrLf & "Ele T3  Ty Mdy" & _
vbCrLf & "Ele T4  MdTy" & _
vbCrLf & "Ele Lng Lno" & _
vbCrLf & "Ele Mem Lines TopRmk" & _
vbCrLf & "Tbl MthCache | PjFfn Md Nm Ty | Mdy Prm Ret LinRmk TopRmk Lines Lno Pj PjDte MdTy" & _
vbCrLf & "Tbl MthLoc   | Nm             | MthMchTy MchStr ToMdNm" & _
vbCrLf & "Tbl MthPfxMd | MthPfx         | MdNm" & _
vbCrLf & ""

MthSchm = A_1
End Function

