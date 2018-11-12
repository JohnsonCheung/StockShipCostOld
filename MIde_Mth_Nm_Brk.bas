Attribute VB_Name = "MIde_Mth_Nm_Brk"
Option Compare Database
Option Explicit

Sub MthBrkAsg(A As Mth, OMdy$, OMthTy$)
Dim L$
L = MthDcl(A)
OMdy = TakMdy(L)
OMthTy = LinMthTy(L)
End Sub

Function MthBrkAyDDNy(A() As Variant) As String()
MthBrkAyDDNy = DryJnDotSy(A)
End Function

Function MthBrkAyWhDup(A()) As Variant()
'MthBrk is Sy of Mdy Ty Nm
Dim Dry(): Dry = DryWhColInAy(A, 0, Array("", "Public")) '
MthBrkAyWhDup = DryWhColHasDup(Dry, 2)
End Function

Function MthNmBrkNm$(MthNmBrk$())
Select Case Sz(MthNmBrk)
Case 0:
Case 3: MthNmBrkNm = MthNmBrk(0)
Case Else: Stop
End Select
End Function

Function MthNmBrkAyWh(A() As Variant, B As WhMth) As Variant()
Dim Brk
For Each Brk In AyNz(A)
    If MthNmBrkIsSel(CvSy(Brk), B) Then PushI MthNmBrkAyWh, Brk
Next
End Function

Function LinMthNmBrk(A, Optional B As WhMth) As String()
Dim O$()
O = ShfMthNmBrk(CStr(A))
If MthNmBrkIsSel(O, B) Then LinMthNmBrk = O
End Function

Function MthNmBrkAyNy(A() As Variant) As String()
MthNmBrkAyNy = DryDistSy(A, 2)
End Function

Sub LinMthNmBrkAsg(A$, OMdy$, OTy$, ONm$)
Dim L$: L = A
OMdy = ShfMdy(L)
OTy = ShfMthTy(L)
ONm = TakNm(L)
End Sub
