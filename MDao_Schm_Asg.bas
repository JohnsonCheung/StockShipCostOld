Attribute VB_Name = "MDao_Schm_Asg"
Option Compare Database
Option Explicit
Private A_Ly$()

Sub Z_SchmAsg()
GoSub Cas1
GoSub Cas2
Exit Sub
Dim Schm$, EptEr$(), EptTdDefAy$(), EptB As StruBase
Cas2:
    GoTo Tst
Cas1:
Schm = _
         "Tbl A *Id *Nm | *Dte AATy Loc Expr Rmk" & _
vbCrLf & "Tbl B *Id AId *Nm | *Dte" & _
vbCrLf & "Fld Txt AATy" & _
vbCrLf & "Fld Loc Loc" & _
vbCrLf & "Fld Expr Expr" & _
vbCrLf & "Fld Mem Rmk" & _
vbCrLf & "Ele Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']" & _
vbCrLf & "Ele Expr Txt [Expr=Loc & 'abc']" & _
vbCrLf & "Des Tbl     A     AA BB " & _
vbCrLf & "Des Tbl     A     CC DD " & _
vbCrLf & "Des Fld     ANm   AA BB " & _
vbCrLf & "Des Tbl.Fld A.ANm TFDes-AA-BB"
EptEr = ApSy("")
EptTdDefAy = ApSy("")
Set EptB.EF.E = New Dictionary
Set EptB.EF.F = New Dictionary
Set EptB.FDes = New Dictionary
Set EptB.FDes = New Dictionary
Set EptB.TFDes = New Dictionary
GoTo Tst
Tst:
    Dim Er$(), TdDefAy$(), B As StruBase
    SchmAsg Schm, Er, TdDefAy, B
    Ass AyIsEq(Er, EptEr)
    Ass AyIsEq(TdDefAy, EptTdDefAy)
    StruBaseIsEqAss B, EptB
    
End Sub
Sub SchmAsg(Schm$, OEr$(), OTdDefAy$(), OStruBase As StruBase)
A_Ly = Split(Schm, vbCrLf)
OEr = SchmLyEr(A_Ly)
If Sz(OEr) > 0 Then Exit Sub
OTdDefAy = AyWhT1SelRst(A_Ly, "Tbl")
OStruBase = ZStruBase
End Sub

Private Function ZEDic() As Dictionary
'L in A_Ly is 'Fld' Ele FldLikss
'->
'EDic is FldLikss->Ele
Set ZEDic = DicSwapKV(LyDic(AyWhT1SelRst(A_Ly, "Fld")))
End Function

Private Function ZFDes() As Dictionary
Set ZFDes = LyDic(AyWhT1SelRst(A_Ly, "FDes"))
End Function

Private Function ZFDic() As Dictionary
Dim E
Set ZFDic = New Dictionary
For Each E In AyNz(AyWhT1SelRst(A_Ly, "Ele"))
    ZFDic.Add LinT1(E), EleDefFd(E)
Next
End Function

Private Function ZStruBase() As StruBase
With ZStruBase
    Set .EF.E = ZEDic
    Set .EF.F = ZFDic
    Set .TDes = ZTDes
    Set .FDes = ZFDes
    Set .TFDes = ZTFDes
End With
End Function

Private Function ZTDes() As Dictionary
Set ZTDes = LyDic(AyWhT1SelRst(A_Ly, "TDes"))
End Function

Private Function ZTFDes() As Dictionary
Dim L, T$, F$, Des$
Dim TF$ ' ttt.fff in the lin of "Des Tbl.Fld ttt.fff ....
Set ZTFDes = New Dictionary
For Each L In AyWhTTSelRst(A_Ly, "Des", "Tbl.Fld")
    LinTRstAsg L, TF, Des
    If Not HasSubStr(TF, ".") Then Stop
    ZTFDes.Add TF, Des
Next
End Function
