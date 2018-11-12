Attribute VB_Name = "bb_Rpt_Inp"
Option Explicit
Option Compare Database
Sub IMB52Opn(): FxOpn IFxMB52: End Sub
Sub IZHT1Opn(): FxOpn IFxZHT1: End Sub
Function IZHT1Fny() As String()
AyDmp DbtFny(W, ">ZHT1")
End Function
Function IFxMB52$()
IFxMB52 = PmFfn("MB52")
End Function
Function IFxUOM$()
IFxUOM = PmFfn("UOM")
End Function
Function IFxAy() As String()
IFxAy = ApSy(IFxMB52, IFxUOM, IFxZHT1)
End Function
Function IFxZHT1$()
IFxZHT1 = PmFfn("ZHT1")
End Function



Function PmStkYYMD$()
PmStkYYMD = Format(PmStkDte, "YYYY-MM-DD")
End Function

Function PmStkDte() As Date
Dim A$
A = Mid(PmVal("MB52Fn"), 6, 10)
PmStkDte = CDate(A)
End Function

