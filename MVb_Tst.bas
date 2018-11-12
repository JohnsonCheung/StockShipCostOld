Attribute VB_Name = "MVb_Tst"
Option Compare Database
Option Explicit
Public Act, Ept, Dbg As Boolean, Trc As Boolean

Sub C()
If Not IsEq(Act, Ept) Then
    ShwDbg
    D "=========================="
    D "Act"
    D Act
    D "---------------"
    D "Ept"
    D Ept
    Stop
End If
End Sub
Function TstHom$()
Static X$
If X = "" Then X = PthEns(CurPjPth & "TstRes\")
TstHom = X
End Function

Sub TstHomBrw()
PthBrw TstHom
End Sub

Sub PjTstPthBrw(PjNm$)
PthBrw PjTstPth(PjNm)
End Sub
Function PjTstPth$(PjNm$)
PjTstPth = PthEns(TstHom & PjNm & "\")
End Function

Sub TstItmEdt(PjNm$, CSub$, Cas$, Itm$)
FtBrw TstItmFt(PjNm, CSub, Cas, Itm)
End Sub

Private Function TstItmFt$(PjNm$, CSub$, Cas$, Itm$)
TstItmFt = FfnEnsPth(PjTstPth(PjNm) & ApJnPthSep(Replace(CSub, ".", PthSep), Cas, Itm) & ".txt")
End Function

Function TstItm$(PjNm$, CSub$, Cas$, Itm$, Optional IsEdt As Boolean)
If IsEdt Then
    TstItmEdt PjNm, CSub, Cas, Itm
    Exit Function
End If
TstItm = FtLines(TstItmFt(PjNm, CSub, Cas, Itm))
End Function

Sub TstOk(CSub$, Cas$)
Debug.Print "Tst OK | "; CSub; " | Case "; Cas
End Sub
