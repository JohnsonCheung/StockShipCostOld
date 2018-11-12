Attribute VB_Name = "MApp_GenOupFx"
Option Explicit
Option Compare Database
Function OupFx_Crt$(A$)
OupFx_Crt = AttExp("Tp", A)
End Function
Sub OupFx_Gen(OupFx$, Fb$, ParamArray WbFmtrAp())
Dim Av()
Av = WbFmtrAp
TpWrtFfn OupFx
WbFmt FxRfh(OupFx, Fb), Av
End Sub

