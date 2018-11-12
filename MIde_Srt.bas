Attribute VB_Name = "MIde_Srt"
Option Compare Database
Option Explicit

Function CurMdSrtLines$()
CurMdSrtLines = MdSrtLines(CurMd)
End Function

Function LinMthSrtKey$(A)
Dim L$, Mdy$, Ty$, Nm$
L = A
Mdy = ShfMdy(L)
Ty = ShfMthTy(L): If Ty = "" Then Exit Function
Nm = TakNm(L)
LinMthSrtKey = MthNm3SrtKey(Mdy, Ty, Nm)
End Function

Function MdSrtLines$(A As CodeModule)
MdSrtLines = SrcSrtLines(MdSrc(A))
End Function

Function MdSrtLy(A As CodeModule) As String()
MdSrtLy = SrcSrtLy(MdSrc(A))
End Function

Function MthDDNmSrtKey$(A) ' MthDDNm is Nm.Ty.Mdy
If A = "*Dcl" Then MthDDNmSrtKey = "*Dcl": Exit Function
Dim B$(): B = SplitDot(A): If Sz(B) <> 3 Then Stop
Dim Mdy$, Ty$, Nm$
AyAsg B, Nm, Ty, Mdy
MthDDNmSrtKey = MthNm3SrtKey(Mdy, Ty, Nm)
End Function

Function MthNm3SrtKey$(Mdy$, Ty$, Nm$)
Dim P% 'Priority
    Select Case True
    Case HasPfx(Nm, "Init"): P = 1
    Case Nm = "Z":           P = 9
    Case HasPfx(Nm, "Z_"):   P = 8
    Case HasPfx(Nm, "ZZ_"):  P = 7
    Case HasPfx(Nm, "Z"):    P = 6
    Case Else:               P = 2
    End Select
MthNm3SrtKey = P & ":" & Nm & ":" & Ty & ":" & Mdy
End Function

Sub PjSrt(A As VBProject)
Dim M As CodeModule, I, Ay() As CodeModule
If Sz(Ay) = 0 Then Exit Sub
For Each I In Ay
    MdSrt CvMd(I)
Next
End Sub

Function SrcSrtDic(A$()) As Dictionary
Dim D As Dictionary, K
Set D = SrcDic(A)
Dim O As New Dictionary
    For Each K In D
        O.Add MthDDNmSrtKey(K), D(K)
    Next
Set SrcSrtDic = DicSrt(O)
End Function

Function SrcSrtLines$(A$())
SrcSrtLines = JnDblCrLf(SrcSrtDic(A).Items)
End Function

Function SrcSrtLy(A$()) As String()
SrcSrtLy = SplitCrLf(SrcSrtLines(A))
End Function

Private Sub ZZ_Dcl_BefAndAft_Srt()
Const MdNm$ = "VbStrRe"
Dim A$() ' Src
Dim B$() ' Src->Srt
Dim A1$() 'Src->Dcl
Dim B1$() 'Src->Src->Dcl
A = MdSrc(Md(MdNm))
B = SrcSrtLy(A)
A1 = SrcDclLy(A)
B1 = SrcDclLy(B)
Stop
End Sub

Private Sub ZZ_MthDDNmSrtKey()
GoSub X0
GoSub X1
Exit Sub
X0:
    Dim Ay1$(): Ay1 = SrcMthNy(CurSrc)
    Dim Ay2$(): Ay2 = AyMapSy(Ay1, "MthNmSrtKey")
    S1S2AyBrw AyabS1S2Ay(Ay2, Ay1)
    Return
X1:
    Const A$ = "YYA.Fun."
    Debug.Print MthDDNmSrtKey(A)
    Return
End Sub

Private Sub Z_LinMthSrtKey()
GoTo ZZ
Dim A$
'
Ept = "2:LinMthSrtKey:Function:": A = "Function LinMthSrtKey$(A)": GoSub Tst
Ept = "2:YYA:Function:":          A = "Function YYA()":            GoSub Tst
Exit Sub
Tst:
    Act = LinMthSrtKey(A)
    C
    Return
ZZ:
    Dim Ay1$(): Ay1 = SrcMthDclAy(CurSrc)
    Dim Ay2$(): Ay2 = AyMapSy(Ay1, "LinMthSrtKey")
    S1S2AyBrw AyabS1S2Ay(Ay2, Ay1)
End Sub

Private Sub Z_SrcSrtLy()
Brw SrcSrtLines(CurSrc)
End Sub

Private Sub Z()
Z_LinMthSrtKey
Z_SrcSrtLy
End Sub
