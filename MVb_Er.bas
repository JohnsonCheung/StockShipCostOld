Attribute VB_Name = "MVb_Er"
Option Compare Database
Option Explicit
'Calling this module functions will throw error
Sub ErWh(Fun$, Msg$, Ny0, ParamArray Ap())
Dim Av(): Av = Ap
ErWhAv Fun, Msg, Ny0, Av
End Sub

Sub ErWhAv(Fun$, MsgVbl$, Ny0, Av())
AyBrw FunMsgNyAvLy(Fun, MsgVbl, Ny0, Av)
RaiseErr
End Sub

Sub Er(Fun$, SqBktMacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
ErWhAv Fun, SqBktMacroStr, MacroNy(SqBktMacroStr), Av
End Sub

Sub ChkAss(Chk$())
If Sz(Chk) = 0 Then Exit Sub
AyBrw Chk
Stop
End Sub

Sub RaiseErr()
Err.Raise -1, , "Please check messages opened in notepad"
End Sub

