Attribute VB_Name = "MApp_Pm"
Option Explicit
Option Compare Database

Function PmOupPth$()
PmOupPth = PmVal("OupPth")
End Function

Function PmFfn$(A$)
PmFfn = PmPth(A) & PmFn(A)
End Function
Function PmPth$(A$)
PmPth = PthEnsSfx(PmVal(A & "Pth"))
End Function

Function PmFn$(A$)
PmFn = PmVal(A & "Fn")
End Function

Property Get PmVal$(Pm$)
PmVal = CurrentDb.TableDefs("Prm").OpenRecordset.Fields(Pm).Value
End Property

Property Let PmVal(Pm$, V$)
With CurrentDb.TableDefs("Prm").OpenRecordset
    .Edit
    .Fields(Pm).Value = V
    .Update
End With
End Property


