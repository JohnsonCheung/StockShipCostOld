Attribute VB_Name = "MIde_Z_Md"
Option Compare Database
Option Explicit
Const CMod$ = "MIde_Z_Md."

Function CvMd(A) As CodeModule
Set CvMd = A
End Function

Function Md(MdDNm) As CodeModule
Dim A1$(): A1 = Split(MdDNm, ".")
Select Case Sz(A1)
Case 1: Set Md = PjMd(CurPj, MdDNm)
Case 2: Set Md = PjMd(Pj(A1(0)), A1(1))
Case Else: Er "Md", "[MdDNm] should be XXX.XXX or XXX", MdDNm
End Select
End Function

Function Md_TstSub_Lno%(A As CodeModule)
Dim J%
For J = 1 To A.CountOfLines
    If LinIsTstSub(A.Lines(J, 1)) Then Md_TstSub_Lno = J: Exit Function
Next
End Function

Function MdAyWhInTy(A() As CodeModule, WhInTyAy0$) As CodeModule()
Dim TyAy() As vbext_ComponentType, Md
TyAy = CvWhCmpTy(WhInTyAy0)
Dim O() As CodeModule
For Each Md In A
    If AyHas(TyAy, CvMd(Md).Parent.Type) Then PushObj O, Md
Next
MdAyWhInTy = O
End Function

Function MdAyWhMdy(A() As CodeModule, CmpTyAy0) As CodeModule()
'MdAyWhMdy = AyWhPredXP(A, "MdIsInCmpAy", CvCmpTyAy(CmpTyAy0))
End Function

Function MdCanHasCd(A As CodeModule) As Boolean
Select Case MdTy(A)
Case _
    vbext_ComponentType.vbext_ct_StdModule, _
    vbext_ComponentType.vbext_ct_ClassModule, _
    vbext_ComponentType.vbext_ct_Document, _
    vbext_ComponentType.vbext_ct_MSForm
    MdCanHasCd = True
End Select
End Function

Sub MdClsWin(A As CodeModule)
A.CodePane.Window.Close
End Sub

Function MdCmp(A As CodeModule) As VBComponent
Set MdCmp = A.Parent
End Function

Function MdCmpTy(A As CodeModule) As vbext_ComponentType
MdCmpTy = A.Parent.Type
End Function

Sub MdCompare(A As CodeModule, B As CodeModule)
Dim A1 As Dictionary, B1 As Dictionary
    Set A1 = MdMthDic(A)
    Set B1 = MdMthDic(B)
DicCmpBrw A1, B1, MdDNm(A), MdDNm(B)
End Sub

Function MdDNm$(A As CodeModule)
MdDNm = MdPjNm(A) & "." & MdNm(A)
End Function

Function MdHasNoMth(A As CodeModule) As Boolean
Dim J&
For J = A.CountOfDeclarationLines + 1 To A.CountOfLines
    If LinIsMth(A.Lines(J, 1)) Then Exit Function
Next
MdHasNoMth = True
End Function

Function MdMthBrkAy(A As CodeModule, Optional B As WhMth) As Variant()
MdMthBrkAy = SrcMthBrkAy(MdSrc(A), B)
End Function
Function MdNm$(A As CodeModule)
MdNm = A.Parent.Name
End Function

Function MdNTy%(A As CodeModule)
MdNTy = SrcNTy(MdDclLy(A))
End Function

Function MdPj(A As CodeModule) As VBProject
Set MdPj = A.Parent.Collection.Parent
End Function

Function MdPjNm$(A As CodeModule)
MdPjNm = MdPj(A).Name
End Function

Sub MdReportSorting(A As CodeModule)
Dim Old$: Old = MdBdyLines(A)
Dim NewLines$: NewLines = MdSrtLines(A)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print MdNm(A), O
End Sub

Function MdResLy(A As CodeModule, ResNm$, Optional ResPfx$ = "ZZRes") As String()
Dim Z$()
    Z = MdMthBdyLy(A, ResPfx & ResNm)
    If Sz(Z) = 0 Then
        Er "MdResLy", "{MthNm} in {Md} is not found", ResPfx & ResNm, MdNm(A)
    End If
    Z = AyExlFstEle(Z)
    Z = AyExlLasEle(Z)
MdResLy = AyRmvFstChr(Z)
End Function

Function MdResStr$(A As CodeModule, ResNm$)
MdResStr = JnCrLf(MdResLy(A, ResNm))
End Function

Sub MdRmvFmCnt(A As CodeModule, FmCnt As FmCnt)
With FmCnt
    If .Cnt = 0 Then Exit Sub
    A.DeleteLines .FmLno, .Cnt
End With
End Sub

Sub MdRmvFmCntAy(A As CodeModule, FmCntAy() As FmCnt)
If Sz(FmCntAy) = 0 Then Exit Sub
Dim J%, M&
M = FmCntAy(0).FmLno
For J = 1 To UB(FmCntAy)
    If M > FmCntAy(J).FmLno Then Stop
    M = FmCntAy(J).FmLno
Next

For J = UB(FmCntAy) To 0 Step -1
    MdRmvFmCnt A, FmCntAy(J)
Next
End Sub

Sub MdRmvNmPfx(A As CodeModule, Pfx$)
Dim Nm$: Nm = MdNm(A): If Not HasPfx(Nm, Pfx) Then Exit Sub
MdRen A, RmvPfx(MdNm(A), Pfx)
End Sub

Sub MdRmvPfx(A As CodeModule, Pfx$)
MdRen A, RmvPfx(MdNm(A), Pfx)
End Sub

Sub MdRpl(A As CodeModule, NewMdLines$)
MdClr A
MdLinesApp A, NewMdLines
End Sub

Sub MdRplBdy(A As CodeModule, NewMdBdy$)
MdClrBdy A
MdLinesApp A, NewMdBdy
End Sub

Sub MdRplDclLy(A As CodeModule, DclLy$())
MdRmvDcl A
A.InsertLines 1, JnCrLf(DclLy)
End Sub

Sub MdRplLin(A As CodeModule, Lno, NewLin$)
With A
    .DeleteLines Lno
    .InsertLines Lno, NewLin
End With
End Sub

Sub MdSav(A As CodeModule)

End Sub

Sub MdShw(A As CodeModule)
A.CodePane.Show
End Sub

Function MdSrc(A As CodeModule) As String()
MdSrc = MdLy(A)
End Function

Function MdSrcFn$(A As CodeModule)
MdSrcFn = MdNm(A) & MdSrcExt(A)
End Function

Sub MdSrt(A As CodeModule)
Dim Nm$: Nm = MdNm(A)
Debug.Print "Sorting: "; AlignL(Nm, 20); " ";
Dim LinesN$: LinesN = MdSrtLines(A)
Dim LinesO$: LinesO = MdLines(A)
'Exit if same
    If LinesO = LinesN Then
        Debug.Print "<== Same"
        Exit Sub
    End If
'Delete
    Debug.Print FmtQQ("<--- Deleted (?) lines", A.CountOfLines);
    MdClr A, IsSilent:=True
'Add sorted lines
    A.AddFromString LinesN
    MdRmvEndBlankLin A
    Debug.Print "<----Sorted Lines added...."
End Sub

Function MdTopRmkMthLinesAy(A As CodeModule) As String()
MdTopRmkMthLinesAy = SrcDicTopRmkMthLinesAy(MdMthDic(A))
End Function

Function MdTy(A As CodeModule) As vbext_ComponentType
MdTy = A.Parent.Type
End Function

Function MdTyStr$(A As CodeModule)
MdTyStr = CmpTyStr(A.Parent.Type)
End Function

Sub Srt()
MdSrt CurMd
End Sub

Private Function ZZMd() As CodeModule
Set ZZMd = CurVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Function

Private Sub ZZ_MdDrs()
'DrsBrw MdDrs(Md("IdeFeature_EnsZZ_AsPrivate"))
End Sub

Private Sub ZZ_MdMthLno()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MdMthNy(A)
    For Each M In Ny
        DoEvents
        J = J + 1
        Push L, MdMthLno(A, CStr(M))
        If J Mod 150 = 0 Then
            Debug.Print J, Sz(Ny), "Z_MdMthLno"
        End If
    Next

    For Each Lno In L
        Push O, Lno & " " & A.Lines(Lno, 1)
    Next
AyBrw O
End Sub

Private Sub ZZ_MdSrt()
Dim Md As CodeModule
GoSub X0
Exit Sub
X0:
    Dim I
'    For Each I In PjMdAy(CurPjx)
        Set Md = I
        If MdNm(Md) = "Str_" Then
            GoSub Ass
        End If
'    Next
    Return
X1:

    Return
Ass:
    Debug.Print MdNm(Md); vbTab;
    Dim BefSrt$(), AftSrt$()
    BefSrt = MdLy(Md)
    AftSrt = SplitCrLf(MdSrtLines(Md))
    If JnCrLf(BefSrt) = JnCrLf(AftSrt) Then
        Debug.Print "Is Same of before and after sorting ......"
        Return
    End If
    If Sz(AftSrt) <> 0 Then
        If AyLasEle(AftSrt) = "" Then
            Dim Pfx
            Pfx = Array("There is non-blank-line at end after sorting", "Md=[" & MdNm(Md) & "=====")
            AyBrw AyAddAp(Pfx, AftSrt)
            Stop
        End If
    End If
    Dim A$(), B$(), II
    A = AyMinus(BefSrt, AftSrt)
    B = AyMinus(AftSrt, BefSrt)
    Debug.Print
    If Sz(A) = 0 And Sz(B) = 0 Then Return
    If Sz(AyExlEmpEle(A)) <> 0 Then
        Debug.Print "Sz(A)=" & Sz(A)
        AyBrw A
        Stop
    End If
    If Sz(AyExlEmpEle(B)) <> 0 Then
        Debug.Print "Sz(B)=" & Sz(B)
        AyBrw B
        Stop
    End If
    Return
End Sub

Private Sub ZZ_MdSrtLines()
StrBrw MdSrtLines(CurMd)
End Sub

Private Sub Z()
Z_MdCpy
Z_MdEndTrim
Z_MdEnmMbrCnt
Z_MdEnsPrpOnEr
Z_MdLinesApp
Z_MdLy
Z_MdMthBdyLines
Z_MdMthBdyLy
Z_MdMthDDNy
Z_MdMthLinCnt
Z_MdRmvFmCntAy
Z_MdRmvPrpOnEr
Z_MdTopRmkMthLinesAy
End Sub

Private Sub Z_MdCpy()
Dim A As CodeModule, ToPj As VBProject
'
Set ToPj = CurPj
Set A = Md("QDta.Dt")
GoSub Tst
Exit Sub
Tst:
    Dim N$
    N = MdNm(A)
    Stop
    MdCpy A, ToPj   '<====
    Ass PjHasMd(ToPj, N) = True
    PjDltMd ToPj, N
    Return
End Sub

Private Sub Z_MdEndTrim()
Dim M As CodeModule: Set M = Md("ZZModule")
MdLinesApp M, "  "
MdLinesApp M, "  "
MdLinesApp M, "  "
MdLinesApp M, "  "
MdEndTrim M, ShwMsg:=True
Ass M.CountOfLines = 15
End Sub

Private Sub Z_MdEnmMbrCnt()
Ass MdEnmMbrCnt(Md("Ide"), "AA") = 1
End Sub

Private Sub Z_MdEnsPrpOnEr()
MdEnsPrpOnEr ZZMd
End Sub

Private Sub Z_MdLinesApp()
Const MdNm$ = "Module1"
MdLinesApp CurMd, "'aa"
End Sub

Private Sub Z_MdLy()
AyBrw MdLy(CurMd)
End Sub

Private Sub Z_MdMthBdyLines()
Debug.Print Len(MdMthBdyLines(CurMd, "MdMthLines"))
Debug.Print MdMthBdyLines(CurMd, "MdMthLines")
End Sub

Private Sub Z_MdMthBdyLy()
Debug.Print Len(MdMthBdyLines(CurMd, "MdMthLines"))
Debug.Print MdMthBdyLines(CurMd, "MdMthLines")
End Sub

''======================================================================================
Private Sub Z_MdMthDDNy()
Dim Md1 As CodeModule
Set Md1 = Md("AAAMod")
Brw MdMthNy(Md1)
Brw MdMthDDNy(Md1)
End Sub

Private Sub Z_MdMthLinCnt()
Dim O$()
    Dim J%, M, L&, E&, A As CodeModule, Ny$()
    Set A = Md("Fct")
    Ny = MdMthNy(A)
    For Each M In Ny
        DoEvents
        L = MdMthLno(A, CStr(M))
        E = MdMthLinCnt(A, L) + L - 1
        Push O, Format(L, "0000 ") & A.Lines(L, 1)
        Push O, Format(E, "0000 ") & A.Lines(E, 1)
    Next
AyBrw O
End Sub

Private Sub Z_MdRmvFmCntAy()
Dim A() As FmCnt
A = MdMthFmCntAy(Md("Md_"), "XXX")
MdRmvFmCntAy Md("Md_"), A
End Sub

Private Sub Z_MdRmvPrpOnEr()
MdRmvPrpOnEr ZZMd
End Sub

Private Sub Z_MdTopRmkMthLinesAy()
Brw Jn(MdTopRmkMthLinesAy(CurMd), vbCrLf & "-----------------------------------------------" & vbCrLf)
End Sub
