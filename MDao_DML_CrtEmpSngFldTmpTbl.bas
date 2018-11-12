Attribute VB_Name = "MDao_DML_CrtEmpSngFldTmpTbl"
Option Compare Database
Option Explicit

Function DbtfCrtTmpTbl$(A As Database, T$, F$, TmpTbl$)
Const C$ = "Select [?] into [?] from [?] Where False"
Q = FmtQQ(C, F, TmpTbl, T)
A.Execute Q
End Function
