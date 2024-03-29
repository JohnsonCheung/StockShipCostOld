Attribute VB_Name = "MIde_Dcl_EnmAndTy"
Option Compare Database
Option Explicit
Const CMod$ = "MIde_Dcl_EnmAndTy."

Function DclEnmBdyLy(A$(), EnmNm$) As String()
Const CSub$ = CMod & "DclEnmBdyLy"
Dim B%: B = DclEnmIx(A, EnmNm): If B = -1 Then Er CSub, "No [EnmNm] in [Dcl]", EnmNm, A
Dim J%
For J = B To UB(A)
   PushI DclEnmBdyLy, A(J)
   If HasPfx(A(J), "End Enum") Then Exit Function
Next
Er CSub, "No End Enum in [Src] for [EnmNm]", A, EnmNm
End Function

Function DclEnmIx%(A$(), EnmNm$)
Dim J%, L
For Each L In AyNz(A)
    L = RmvMdy(L)
    If ShfXEnm(L) Then
        If TakNm(L) = EnmNm Then
            DclEnmIx = J
            Exit Function
        End If
    End If
    If LinIsMth(L) Then Exit For
    J = J + 1
Next
DclEnmIx = -1
End Function

Function DclEnmNy(A$()) As String()
Dim L
For Each L In AyNz(A)
   PushNonBlankStr DclEnmNy, LinEnmNm(L)
Next
End Function

Function DclHasTy(A$(), TyNm$) As Boolean
Dim L
For Each L In AyNz(A)
    If LinTyNm(L) = TyNm Then DclHasTy = True: Exit Function
Next
End Function

Function DclNEnm%(A$())
Dim L, O%
For Each L In AyNz(A)
   If IsEmnLin(L) Then O = O + 1
Next
DclNEnm = O
End Function

Function DclTyFmIx%(A$(), TyNm$)
Dim J%, L$
For J = 0 To UB(A)
   If LinTyNm(A(J)) = TyNm Then DclTyFmIx = J: Exit Function
Next
DclTyFmIx = -1
End Function

Function DclTyFTIx(A$(), TyNm$) As FTIx
Dim FmI&: FmI = DclTyFmIx(A, TyNm)
Dim ToI&: ToI = DclTyToIx(A, FmI)
Set DclTyFTIx = FTIx(FmI, ToI)
End Function

Function DclTyIxToIx%(A$(), TyIx%)
If 0 > TyIx Then DclTyIxToIx = -1: Exit Function
Dim O&
For O = TyIx + 1 To UB(A)
   If HasPfx(A(O), "End Type") Then DclTyIxToIx = O: Exit Function
Next
DclTyIxToIx = -1
End Function

Function DclTyLines$(A$(), TyNm$)
DclTyLines = JnCrLf(DclTyLy(A, TyNm))
End Function

Function DclTyLy(A$(), TyNm$) As String()
DclTyLy = AyWhFTIx(A, DclTyFTIx(A, TyNm))
End Function

Function DclTyNmIx&(A$(), TyNm)
Dim J%
For J = 0 To UB(A)
   If LinTyNm(A(J)) = TyNm Then DclTyNmIx = J: Exit Function
Next
DclTyNmIx = -1
End Function

Function DclTyNy(A$()) As String()
Dim L
For Each L In AyNz(A)
    PushNonBlankStr DclTyNy, LinTyNm(L)
Next
End Function

Private Function DclTyToIx%(A$(), FmIx)
If 0 > FmIx Then DclTyToIx = -1: Exit Function
Dim O&
For O = FmIx + 1 To UB(A)
   If HasPfx(A(O), "End Type") Then DclTyToIx = O: Exit Function
Next
DclTyToIx = -1
End Function

Function IsEmnLin(A) As Boolean
IsEmnLin = HasPfx(RmvMdy(A), "Enum ")
End Function

Function IsTyLin(A) As Boolean
IsTyLin = HasPfx(RmvMdy(A), "Type ")
End Function

Function LinEnmNm$(A)
Dim L$: L = RmvMdy(A)
If ShfXEnm(L) Then LinEnmNm = TakNm(L)
End Function

Function LinTyNm$(A)
Dim L$: L = RmvMdy(A)
If ShfT(L) Then LinTyNm = TakNm(L)
End Function

Function MdEnmBdyLy(A As CodeModule, EnmNm$) As String()
MdEnmBdyLy = DclEnmBdyLy(MdDclLy(A), EnmNm)
End Function

Function MdEnmMbrCnt%(A As CodeModule, EnmNm$)
MdEnmMbrCnt = Sz(MdEnmMbrLy(A, EnmNm))
End Function

Function MdEnmMbrLy(A As CodeModule, EnmNm$) As String()
MdEnmMbrLy = AyWhCdLin(MdEnmBdyLy(A, EnmNm))
End Function

Function MdEnmNy(A As CodeModule) As String()
MdEnmNy = DclEnmNy(MdDclLy(A))
End Function

Function MdNEnm%(A As CodeModule)
MdNEnm = DclNEnm(MdDclLy(A))
End Function

Function MdTyLCC(A As CodeModule, TyNm$) As LCC
Dim R&, C1&, C2&
R = MdTyLno(A, TyNm)
If R > 0 Then
    With SubStrPos(A.Lines(R, 1), TyNm)
        C1 = .FmIx
        C2 = .ToIx
    End With
End If
MdTyLCC = LCC(R, C1, C2)
End Function

Function MdTyLno$(A As CodeModule, TyNm$)
MdTyLno = -1
End Function

Function MdTyNm$(A As CodeModule)
MdTyNm = CmpTyStr(MdCmpTy(A))
End Function

Function MdTyNy(A As CodeModule) As String()
MdTyNy = AySrt(DclTyNy(MdDclLy(A)))
End Function

Function PjTyNy(A As VBProject, Optional TyNmPatn$ = ".", Optional MdNmPatn$ = ".") As String()
Dim I, Ny$(), O$()
For Each I In AyNz(PjMdAy(A, WhMd(Nm:=WhNm(MdNmPatn))))
    Ny = MdTyNy(CvMd(I))
    Ny = AyWhPatn(Ny, TyNmPatn)
    PushIAy O, AyAddPfx(Ny, MdNm(CvMd(I)) & ".")
Next
PjTyNy = AyQSrt(O)
End Function

Function ShfXEnm(O) As Boolean
ShfXEnm = ShfX(O, "Enum")
End Function

Function ShfXTy(O) As Boolean
ShfXTy = ShfX(O, "Type")
End Function

Private Sub Z()
Z_DclTyLines
End Sub

Private Sub Z_DclTyLines()
Debug.Print DclTyLines(MdDclLy(CurMd), "AA")
End Sub
