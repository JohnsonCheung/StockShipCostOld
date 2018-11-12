Attribute VB_Name = "MIde_Z_Pj"
Option Compare Database
Option Explicit
Const CMod$ = "MIde_Z_Pj."
Function Pj(A) As VBProject
Set Pj = CurVbe.VBProjects(A)
End Function

Function IsPjNm(A) As Boolean
IsPjNm = AyHas(PjNy, A)
End Function

Function PjClsAndModNy(A As VBProject, Optional Patn$, Optional Exl$) As String()
PjClsAndModNy = PjCmpNy(A, WhMd("Cls Mod", WhNm(Patn, Exl)))
End Function

Function PjClsAndModAy(A As VBProject, Optional Patn$, Optional Exl$) As CodeModule()
PjClsAndModAy = PjModAy(A, WhMd("Cls Mod", WhNm(Patn, Exl)))
End Function

Function PjClsAndModCmpAy(A As VBProject, Optional Patn$, Optional Exl$) As VBComponent()
PjClsAndModCmpAy = PjCmpAy(A, WhMd("Cls Mod", WhNm(Patn, Exl)))
End Function
Function PjClsAndModCmpNy(A As VBProject, Optional Patn$, Optional Exl$) As String()
PjClsAndModCmpNy = OyNy(PjClsAndModCmpAy(A, Patn, Exl))
End Function

Sub PjCompile(A As VBProject)
PjGo A
AssCompileBtn PjNm(A)
With CompileBtn
    If .Enabled Then
        .Execute
        Debug.Print PjNm(A), "<--- Compiled"
    Else
        Debug.Print PjNm(A), "already Compiled"
    End If
End With
TileVBtn.Execute
SavBtn.Execute
End Sub

Sub PjCpyToSrc(A As VBProject)
FfnCpyToPth A.FileName, PjSrcPth(A), OvrWrt:=True
End Sub

Sub PjCpyToSrcPth(A As VBProject)
FfnCpyToPth A.FileName, PjSrcPth(A), OvrWrt:=True
End Sub

Function PjFfn$(A As VBProject)
On Error Resume Next
PjFfn = A.FileName
End Function

Function PjFfnApp(PjFfn) ' Return either Xls.Application (CurXls) or Acs.Application (Function-static)
Static Y As New Access.Application
Select Case True
Case IsFxa(PjFfn): FxaOpn PjFfn: Set PjFfnApp = CurXls
Case IsFb(PjFfn): Y.OpenCurrentDatabase PjFfn: Set PjFfnApp = Y
Case Else: Stop
End Select
End Function

Function PjFn$(A As VBProject)
PjFn = FfnFn(PjFfn(A))
End Function


Function CvPj(I) As VBProject
Set CvPj = I
End Function


Function PjFunPfxAy(A As VBProject) As String()
Dim Ay() As CodeModule: Ay = PjMdAy(A)
Dim Ay1(): Ay1 = AyMap(Ay, "MdFunPfx")
PjFunPfxAy = AyFlat(Ay1)
End Function


Sub PjImpSrcFfn(A As VBProject, SrcFfn)
A.VBComponents.Import SrcFfn
End Sub

Function PjIsUsrLib(A As VBProject) As Boolean
PjIsUsrLib = PjIsFxa(A)
End Function

Function PjMd(A As VBProject, Nm) As CodeModule
Set PjMd = PjCmp(A, Nm).CodeModule
End Function

Function PjMdAy(A As VBProject, Optional B As WhMd) As CodeModule()
If IsNothing(B) Then
    PjMdAy = ItrPrpInto(A.VBComponents, "CodeModule", PjMdAy)
    Exit Function
End If
Dim C
For Each C In AyNz(ItrWhNm(A.VBComponents, B.Nm))
    With CvCmp(C)
        If ItmIsSel(B.InCmpTy, .Type) Then
            PushObj PjMdAy, .CodeModule
        End If
    End With
Next
End Function

Sub PjMdDicApp(A As VBProject, MdDic As Dictionary)
Dim MdNm
For Each MdNm In MdDic.Keys
    PjEnsMod A, MdNm
    MdLinesApp PjMd(A, MdNm), MdDic(MdNm)
Next
End Sub

Function PjMdNy(A As VBProject, Optional B As WhMd) As String()
PjMdNy = PjCmpNy(A, B)
End Function

Function PjMdOpt(A As VBProject, Nm) As CodeModule
If Not PjHasMd(A, Nm) Then Exit Function
Set PjMdOpt = PjMd(A, Nm)
End Function

Function PjModAy(A As VBProject, Optional B As WhNm) As CodeModule()
PjModAy = PjMdAy(A, WhMd("Mod", B))
End Function

Function PjModClsNy(A As VBProject, Optional B As WhNm) As String()
PjModClsNy = PjCmpNy(A, WhMd("Mod Cls", B))
End Function

Function PjModNy(A As VBProject, Optional B As WhNm) As String()
PjModNy = PjCmpNy(A, WhMd("Mod", B))
End Function

Function PjMthKy(A As VBProject, Optional IsWrap As Boolean) As String()
PjMthKy = AyMapPXSy(PjMdAy(A), "MdMthKy", IsWrap)
End Function

Function PjMthKySq(A As VBProject) As Variant()
PjMthKySq = MthKy_Sq(PjMthKy(A, True))
End Function

Function PjMthKyWs(A As VBProject) As Worksheet
Set PjMthKyWs = WsVis(SqWs(PjMthKySq(A)))
End Function

Function PjMthLinDry(A As VBProject) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushAy PjMthLinDry, MdMthLinDry(CvMd(M))
Next
End Function

Function PjMthLinDryWP(A As VBProject) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushIAy PjMthLinDryWP, MdMthLinDryWP(CvMd(M))
Next
End Function

Private Sub Z_PjMthLinDry()
Dim A(): A = PjMthLinDry(CurPj)
Stop
End Sub

Function PjMthNy(A As VBProject, Optional B As WhMdMth) As String()
Dim Md As CodeModule, I, N$, Ny$()
N = A.Name & "."
For Each I In AyNz(PjMdAy(A, WhMdMthMd(B)))
    Set Md = I
    Ny = MthDDNyWh(MdMthDDNy(Md), B.Mth)
    Ny = AyAddPfx(Ny, N & MdNm(Md) & ".")
    PushAyNoDup PjMthNy, Ny
Next
End Function

Function PjNm$(A As VBProject)
PjNm = A.Name
End Function

Function PjNy() As String()
PjNy = ItrNy(CurVbe.VBProjects)
End Function

Function PjPatnLy(A As VBProject, Patn$) As String()
Dim I, Md As CodeModule, O$()
For Each I In PjMdAy(A)
   Set Md = I
   PushAy O, MdPatnLy(Md, Patn)
Next
PjPatnLy = O
End Function

Function PjPrpInfDt(A As VBProject) As Dt
End Function

Function PjPth$(A As VBProject)
PjPth = FfnPth(A.FileName)
End Function

Function PjReadRfCfg(A As VBProject) As String()
Const CSub$ = CMod & "PjReadRfCfg"
Dim B$: B = PjRfCfgFfn(A)
If Not FfnIsExist(B) Then Er CSub, "{Pj-Rf-Cfg-Fil} not found", B
PjReadRfCfg = FtLy(B)
End Function

Sub PjRmvRf(A As VBProject, RfNy0$)
AyDoPX CvNy(RfNy0), "PjRmvRf__X", A
PjSav A
End Sub

Private Sub PjRmvRf__X(A As VBProject, RfNm$)
If PjHasRfNm(A, RfNm) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfNm(?)", A.Name, RfNm)
    Exit Sub
End If
Dim RfFfn$: RfFfn = PjRfNmRfFfn(A, RfNm)
If PjHasRfFfn(A, RfFfn) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfFfnNm(?)", A.Name, RfFfn)
    Exit Sub
End If
A.References.AddFromFile RfFfn
End Sub

Sub PjSav(A As VBProject)
If FstChr(PjNm(A)) <> "Q" Then
    Debug.Print "PjSav: Project Name begin with Q, it is not saved: "; PjNm(A)
    Exit Sub
End If
If A.Saved Then
    Debug.Print FmtQQ("PjSav: Pj(?) is already saved", A.Name)
    Exit Sub
End If
Dim Fn$: Fn = PjFn(A)
If Fn = "" Then
    Debug.Print FmtQQ("PjSav: Pj(?) needs saved first", A.Name)
    Exit Sub
End If
PjAct A
If ObjPtr(CurPj) <> ObjPtr(A) Then Stop: Exit Sub
Dim B As CommandBarButton: Set B = SavBtn
If Not StrIsEq(B.Caption, "&Save " & Fn) Then Stop
B.Execute
If A.Saved Then Stop
Debug.Print FmtQQ("PjSav: Pj(?) is saved <---------------", A.Name)
End Sub

Function PjSrc(A As VBProject) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    PushAy PjSrc, MdSrc(C.CodeModule)
Next
End Function

Function PjSrcPth$(A As VBProject)
Dim P$:
P = FfnPth(A.FileName) & "Src\" & FfnFn(A.FileName) & "\"
PjSrcPth = PthEnsAll(P)
End Function

Sub PjSrcPthBrw(A As VBProject)
PthBrw PjSrcPth(A)
End Sub

Function PjTim(A As VBProject) As Date
PjTim = FfnTim(PjFfn(A))
End Function


Sub Pj_Gen_TstClass(A As VBProject)
If PjHasCmp(A, "Tst") Then
    CmpRmv PjCmp(A, "Tst")
End If
PjAddCls A, "Tst"
PjMd(A, "Tst").AddFromString Pj_TstClass_Bdy(A)
End Sub

Function Pj_TstClass_Bdy$(A As VBProject)
Dim N1$() ' All Class Ny with 'Friend Sub Z' method
Dim N2$()
Dim A1$, A2$
Const Q1$ = "Sub ?()|Dim A As New ?: A.Z|End Sub"
Const Q2$ = "Sub ?()|#.?.Z|End Sub"
N1 = Pj_ClsNy_With_TstSub(A)
A1 = SeedExpand(Q1, N1)
N2 = PjMdNy_With_TstSub(A)
A2 = Replace(SeedExpand(Q2, N2), "#", A.Name)
Pj_TstClass_Bdy = A1 & vbCrLf & A2
End Function


Private Sub ZZ_PjCompile()
PjCompile CurPj
End Sub

Private Sub ZZ_PjHasMd()
Ass PjHasMd(CurPj, "Drs") = False
Ass PjHasMd(CurPj, "A__Tool") = True
End Sub

Private Sub ZZ_PjSav()
PjSav CurPj
End Sub

Private Sub ZZ_PjSrtCmpRptWb()
Dim O As Workbook: Set O = PjSrtCmpRptWb(CurPj, Vis:=True)
Stop
End Sub

Private Sub Z_PjMdDicApp()
Dim MdDic As New Dictionary
Dim ToPj As VBProject: Set ToPj = TmpPj
PjMdDicApp ToPj, MdDic
End Sub

Private Sub Z()
Z_PjMdDicApp
Z_PjMthLinDry
End Sub
