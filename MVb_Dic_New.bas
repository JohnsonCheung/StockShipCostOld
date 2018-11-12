Attribute VB_Name = "MVb_Dic_New"
Option Explicit
Option Compare Database
Function NewSyDic(T1SslLy$()) As Dictionary
Dim L, T$, Ssl$
Dim O As New Dictionary
For Each L In AyNz(T1SslLy)
    LinTRstAsg L, T, Ssl
    If O.Exists(T) Then
        O(T) = AyAdd(O(T), SslSy(Ssl))
    Else
        O.Add T, SslSy(Ssl)
    End If
Next
Set NewSyDic = O
End Function
