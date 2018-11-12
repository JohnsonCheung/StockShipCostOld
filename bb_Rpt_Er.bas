Attribute VB_Name = "bb_Rpt_Er"
Option Explicit
Option Compare Database

Function RptEr() As String()
RptEr = AyAddAp(MB52_8601_8701_Missing)
End Function
Private Function MB52_8601_8701_Missing() As String()
Const M$ = "MB52 file has no [Plant] of 8601 nor 8701"
Const Wh$ = "Plant in ('8601','8701')"
Dim Fx$
Fx = IFxMB52
DbtLnkFx W, "#A", Fx
If DbtNRec(W, "#A", Wh) = 0 Then
    MB52_8601_8701_Missing = _
        FunMsgNyApLy(CSub, M, "MB52-File", Fx)
End If
WDrp "#A"
End Function

