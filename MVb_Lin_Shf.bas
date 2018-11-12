Attribute VB_Name = "MVb_Lin_Shf"
Option Compare Database
Option Explicit

Function ShfVal(OLin$, Lblss) As Variant()
'Return Ay, which is
'   Same sz as QLbl-term-cnt
'   Either string of boolean
'   Each element is corresponding to QLbl
'Return OLin
'   if the term match, it will removed from OLin
'QLbl term: is in *LLL ?LLL or LLL
'   *LLL is always first at beginning, mean the value OLin has not lbl
'   ?LLL means the value is in OLin is using LLL => it is true,
'   LLL  means the value in OLin is LLL=VVV
'OLin is
'   VVV VVV=LLL [VVV=L L]
Dim L, Ay$()
Ay = LinTermAy(OLin)
For Each L In AyNz(CvNy(Lblss))
    PushI ShfVal, ShfVal__ITM(Ay, L)
Next
OLin = JnSpc(AyQuoteSqBktIfNeed(Ay))
End Function

Private Function ShfVal__BOOL(OAy$(), Lbl) As Boolean
Dim J%, L$
L = RmvFstChr(Lbl)
For J = 0 To UB(OAy)
    If OAy(J) = L Then
        ShfVal__BOOL = True
        OAy = AyExlEleAt(OAy, J)
        Exit Function
    End If
Next
End Function

Private Function ShfVal__ITM(OAy$(), Lbl)
If Sz(OAy) = 0 Then ShfVal__ITM = "": Exit Function
'Return either Boolean or string
Select Case FstChr(Lbl)
Case "*": ShfVal__ITM = OAy(0): OAy = AyExlFstEle(OAy)
Case "?": ShfVal__ITM = ShfVal__BOOL(OAy, Lbl)
Case Else: ShfVal__ITM = ShfVal__STR(OAy, Lbl)
End Select
End Function

Private Function ShfVal__STR$(OAy$(), Lbl)
Dim J%
For J = 0 To UB(OAy)
    With Brk1(OAy(J), "=")
        If .S1 = Lbl Then
            ShfVal__STR = .S2
            OAy = AyExlEleAt(OAy, J)
            Exit Function
        End If
    End With
Next
End Function

Function ShfBktStr$(OLin$)
Dim O$
O = TakBetBkt(OLin): If O = "" Then Exit Function
ShfBktStr = O
OLin = TakAftBkt(OLin)
End Function

Function ShfChr$(OLin, ChrList$)
Dim F$, P%
F = FstChr(OLin)
P = InStr(ChrList, F)
If P > 0 Then
    ShfChr = Mid(ChrList, P, 1)
    OLin = Mid(OLin, 2)
    Exit Function
End If
End Function

Function ShfPfx(OLin, Pfx) As Boolean
If HasPfx(OLin, Pfx) Then
    OLin = RmvPfx(OLin, Pfx)
    ShfPfx = True
End If
End Function

Function ShfPfxSpc(OLin, Pfx) As Boolean
If HasPfxSpc(OLin, Pfx) Then
    OLin = Mid(OLin, Len(Pfx) + 2)
    ShfPfxSpc = True
End If
End Function

Private Sub Z_ShfVal()
GoSub Cas1
GoSub Cas2
GoSub Cas3
Exit Sub
Dim Lin$, Lblss, EptLin$
Cas1:
    Lin = "1 Req"
    EptLin = ""
    Lblss = "*XX ?Req"
    Ept = Array("1", True)
    GoTo Tst
Cas2:
    Lin = "A B C=123 D=XYZ"
    Lblss = "?B"
    Ept = Array(True)
    EptLin = "A C=123 D=XYZ"
    GoTo Tst
Cas3:
    Lin = "Txt VTxt=XYZ [Dft=A 1] VRul=123 Req"
    Lblss = "*Ty ?Req ?AlwZLen Dft VTxt VRul"
    Ept = Array("Txt", True, False, "A 1", "XYZ", "123")
    EptLin = ""
    GoTo Tst
Tst:
    Act = ShfVal(Lin, Lblss)
    C
    Ass Lin = EptLin
    Return
End Sub

Private Sub Z_ShfBktStr()
Dim A$, Ept1$
A$ = "(O$()) As X": Ept = "O$()": Ept1 = " As X": GoSub Tst
Exit Sub
Tst:
    Act = ShfBktStr(A)
    C
    Ass A = Ept1
    Return
End Sub

Private Function Z_ShfPfx()
Dim O$: O = "AA{|}BB "
Ass ShfPfx(O, "{|}") = "AA"
Ass O = "BB "
End Function


Private Sub Z()
Z_ShfVal
Z_ShfBktStr
Z_ShfPfx
Z_ShfVal
MVb_Lin_Shf:
End Sub
