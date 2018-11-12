Attribute VB_Name = "MVb_Ay_Ap"
Option Compare Database
Option Explicit

Function ApAy(ParamArray Ap()) As Variant()
ApAy = Ap
End Function

Function ApDotLin$(ParamArray Ap())
Dim Av(): Av = Ap
ApDotLin = JnDot(Av)
End Function

Function ApDteAy(ParamArray Ap()) As Date()
Dim Av(): Av = Ap
ApDteAy = AyDteAy(Av)
End Function
Function ApJnDot$(ParamArray Ap())
Dim Av(): Av = Ap
ApJnDot = JnDot(Av)
End Function
Function ApJnDollar$(ParamArray Ap())
Dim Av(): Av = Ap
ApJnDollar = JnDollar(Av)
End Function
Function ApJnDblDollar$(ParamArray Ap())
Dim Av(): Av = Ap
ApJnDblDollar = JnDollar(Av)
End Function
Function ApJnPthSep$(ParamArray Ap())
Dim Av(): Av = Ap
ApJnPthSep = JnPthSep(Av)
End Function
Function ApIntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
ApIntAy = AyIntAy(Av)
End Function

Function ApInto(OInto, ParamArray AyAp())
Dim Av(): Av = AyAp
Dim Ay
ApInto = AyCln(OInto)
For Each Ay In Av
    PushObjAy ApInto, Ay
Next
End Function

Function ApLin$(ParamArray Ap())
Dim Av(): Av = Ap
ApLin = JnSpc(AyExlEmpEle(Av))
End Function

Function ApLines$(ParamArray Ap())
Dim Av(): Av = Ap
ApLines = JnCrLf(AyExlEmpEle(Av))
End Function

Function ApLngAy(ParamArray Ap()) As Long()
Dim Av(): Av = Ap
ApLngAy = AyLngAy(Av)
End Function

Function ApScl$(ParamArray Ap())
Dim Av(): Av = Ap
ApScl = JnSemiColon(AyExlEmpEle(Av))
End Function

Function ApSngAy(ParamArray Ap()) As Single()
Dim Av(): Av = Ap
ApSngAy = AySngAy(Av)
End Function

Function ApSy(ParamArray Itm_or_Ay_Ap()) As String()
Dim Av(): Av = Itm_or_Ay_Ap
Dim I
For Each I In Av
    If IsArray(I) Then
        PushIAy ApSy, I
    Else
        PushI ApSy, I
    End If
Next
End Function
