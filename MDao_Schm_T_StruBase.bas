Attribute VB_Name = "MDao_Schm_T_StruBase"
Option Explicit
Option Compare Database

Type EF
    E As Dictionary '     E As Dictionary 'Ele->Fd
    F As Dictionary '
End Type
Type StruBase
    EF As EF    'Ele FldLik
    TDes As Dictionary
    FDes As Dictionary
    TFDes As Dictionary
End Type

Sub StruBaseIsEqAss(A As StruBase, B As StruBase)
Stop '
End Sub
