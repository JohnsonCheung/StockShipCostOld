Attribute VB_Name = "MVb_Fs_Pth"
Option Compare Database
Option Explicit
Public Const PthSep$ = "\"

Function Cd$(Optional A$)
If A = "" Then
    Cd = PthEnsSfx(CurDir)
    Exit Function
End If
ChDir A
Cd = PthEnsSfx(A)
End Function

Function CvPth$(ByVal A$)
If A = "" Then
    A = CurDir
End If
CvPth = PthEnsSfx(A)
End Function

Sub PthBrw(A$)
Shell FmtQQ("Explorer ""?""", A), vbMaximizedFocus
End Sub

Sub PthClr(A$)
FfnAyDltIfExist PthFfnAy(A)
End Sub

Sub PthClrFil(A$)
If Not PthIsExist(A) Then Exit Sub
Dim F
For Each F In AyNz(PthFfnAy(A))
   FfnDlt F
Next
End Sub
Function PthHasSfx(A) As Boolean
PthHasSfx = LasChr(A) = PthSep
End Function
Function PthEnsSfx$(A)
If PthHasSfx(A) Then
    PthEnsSfx = A
Else
    PthEnsSfx = A & PthSep
End If
End Function
Function PthEns$(A$)
If Not Fso.FolderExists(A) Then MkDir A
PthEns = PthEnsSfx(A)
End Function

Function PthEnsAll$(A$)
Dim Ay$(): Ay = Split(A, PthSep)
Dim J%, O$
O = Ay(0)
For J = 1 To UB(Ay)
    O = O & PthSep & Ay(J)
    PthEns O
Next
PthEnsAll = A
End Function

Function PthEntAy(A$, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute) As String()
Stop
'Erase O
End Function

Function PthFdr$(A$)
PthFdr = TakAftRev(RmvLasChr(A), "\")
End Function

Function PthFfnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthFfnAy = AyAddPfx(PthFnAy(A, Spec, Atr), A)
End Function

Function PthFnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Ass PthIsExist(A)
Dim O$()
Dim M$
M = Dir(A & Spec)
If Atr = 0 Then
    While M <> ""
       Push O, M
       M = Dir
    Wend
    PthFnAy = O
End If
Ass PthHasPthSfx(A)
While M <> ""
    If GetAttr(A & M) And Atr Then
        Push O, M
    End If
    M = Dir
Wend
PthFnAy = O
End Function

Function PthFxAy(A$) As String()
Dim O$(), B$
If Right(A, 1) <> "\" Then Stop
B = Dir(A & "*.xls")
Dim J%
While B <> ""
    J = J + 1
    If J > 1000 Then Stop
    If FfnExt(B) = ".xls" Then
        Push O, A & B
    End If
    B = Dir
Wend
PthFxAy = O
End Function

Function PthHasFil(A) As Boolean
PthHasFil = Fso.GetFolder(A).Files.Count > 0
End Function

Function PthHasPthSfx(A) As Boolean
PthHasPthSfx = LasChr(A) = "\"
End Function

Function PthHasFdr(A) As Boolean
PthHasFdr = Fso.GetFolder(A).SubFolders.Count > 0
End Function

Function PthIsEmp(A) As Boolean
If PthHasFil(A) Then Exit Function
If PthHasFdr(A) Then Exit Function
PthIsEmp = True
End Function

Function PthIsExist(A) As Boolean
PthIsExist = Fso.FolderExists(A)
End Function

Sub PthMovFilUp(A$)
Dim I, Tar$
Tar$ = PthUp(A)
For Each I In AyNz(PthFnAy(A))
    FfnMov CStr(I), Tar
Next
End Sub

Sub PthRenAddPfx(A, Pfx)
PthRen A, PthAddPfx(A, Pfx)
End Sub
Function PthAddPfx(A, Pfx)
With Brk2Rev(RmvSfx(A, PthSep), PthSep, NoTrim:=True)
    PthAddPfx = .S1 & PthSep & Pfx & .S2 & PthSep
End With
End Function
Sub PthRen(A, NewPth)
If PthIsExist(NewPth) Then ErWh CSub, "NewPth exist", "Pth NewPth", A, NewPth
If Not PthIsExist(A) Then ErWh CSub, "Pth not exist", "Pth NewPth", A, NewPth
Fso.GetFolder(A).Name = NewPth
End Sub

Function PthFdrAy(A, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
'PthFdrAy = ItrNy(Fso.GetFolder(A).SubFolders, Spec)
Ass PthIsExist(A)
Ass PthHasPthSfx(A)
Dim M$, X&
X = Atr Or vbDirectory
M = Dir(A & Spec, vbDirectory)
While M <> ""
    If InStr(M, "?") > 0 Then
        MsgWh CSub, "Unicode entry is skipped", "UniCode-Entry Pth Spec Atr", M, A, Spec, Atr
        GoTo Nxt
    End If
    If M = "." Then GoTo Nxt
    If M = ".." Then GoTo Nxt
    If GetAttr(A & M) And X Then
        PushI PthFdrAy, M
    End If
Nxt:
    M = Dir
Wend
End Function

Function PthPthAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthPthAy = AyAddPfxSfx(PthFdrAy(A, Spec, Atr), A, "\")
End Function

Function PthUp$(A, Optional Up% = 1)
Dim O$, J%
O = A
For J = 1 To Up
    O = PthUpOne(O)
Next
PthUp = O
End Function

Function PthUpOne$(A$)
PthUpOne = TakBefOrAllRev(RmvSfx(A, "\"), "\") & "\"
End Function

Private Sub ZZ_PthFxAy()
Dim A$()
A = PthFxAy(CurDir)
AyDmp A
End Sub

Private Sub ZZ_PthRmvEmpSubDir()
PthRmvEmpSubDir TmpPth
End Sub

