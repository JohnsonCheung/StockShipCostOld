Attribute VB_Name = "MApp_Exp"
Option Explicit
Option Compare Database

Sub AppExp()
PthClr SrcPth
'SpecExp
'AppExpMd
'AppExpFrm
AppExpStru
End Sub

Sub AppExpStru()
StrWrt Stru, SrcPth & "Stru.txt"
End Sub

