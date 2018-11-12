Attribute VB_Name = "MVb__Obj"
Option Compare Database
Option Explicit
Function ObjCompoundPrp$(Obj, PrpSsl$)
Dim Ny$(): Ny = SslSy(PrpSsl)
Dim O$(), I
For Each I In Ny
    Push O, CallByName(Obj, CStr(I), VbGet)
Next
ObjCompoundPrp = Join(O, "|")
End Function

Function ObjDr(A, PrpNy0) As Variant()
Dim PrpNy$(), U%, O(), J%
PrpNy = CvNy(PrpNy0)
U = UB(PrpNy)
ReDim O(U)
For J = 0 To U
    Asg ObjPrp(A, PrpNy(J)), O(J)
Next
ObjDr = O
End Function

Function ObjHasNmPfx(O, NmPfx$) As Boolean
ObjHasNmPfx = HasPfx(ObjNm(O), NmPfx)
End Function
Function ObjIsEq(A, B) As Boolean
ObjIsEq = ObjPtr(A) = ObjPtr(B)
End Function


Function ObjNm$(A)
If IsNothing(A) Then ObjNm = "#nothing#": Exit Function
On Error GoTo X
ObjNm = A.Name
Exit Function
X:
ObjNm = "#" & Err.Description & "#"
End Function

Function ObjPrp(A, P)
If IsNothing(A) Then MsgWh CSub, "Given object is nothing", "PrpNm", P: Exit Function
On Error GoTo X
Asg CallByName(A, P, VbGet), ObjPrp
Exit Function
X:
Dim E$
E = Err.Description
MsgWh CSub, "Error in getting Obj-Prp", "Obj-TypeName PrpNm Er", TypeName(A), P, E
End Function

Function ObjPrpAy(A, PrpNy0) As Variant()
If IsNothing(A) Then MsgWh CSub, "Given object is nothing", "PrpNy0", PrpNy0: Exit Function
Dim I
For Each I In CvNy(PrpNy0)
    Push ObjPrpAy, ObjPrp(A, I)
Next
End Function

Function ObjPrpDr(A, PrpNy0) As Variant()
ObjPrpDr = ObjPrpAy(A, PrpNy0)
End Function

Function ObjPrpPth(A, PrpPth$)
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim P$()
    P = Split(PrpPth, ".")
Dim O
    Dim J%, U%
    Set O = A
    U = UB(P)
    For J = 0 To U - 1      ' U-1 is to skip the last Pth-Seg
        Set O = CallByName(O, P(J), VbGet) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next

Asg CallByName(O, P(U), VbGet), ObjPrpPth ' Last Prp may be non-object, so must use 'Asg'
End Function

Function ObjStr$(A)
If Not IsObject(A) Then Stop
On Error GoTo X
ObjStr = A.ToStr: Exit Function
X: ObjStr = QuoteSqBkt(TypeName(A))
End Function

Private Sub ZZZ_ObjCompoundPrp()
Dim Act$: Act = ObjCompoundPrp(Excel.Application.Vbe.ActiveVBProject, "FileName Name")
Ass Act = "C:\Users\user\Desktop\Vba-Lib-1\QVb.xlam|QVb"
End Sub
