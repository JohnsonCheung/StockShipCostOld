Attribute VB_Name = "MVb_Dic_Ay"
Option Compare Database
Option Explicit
Function DicAyMge(A() As Dictionary) As Dictionary
'Assume there is no duplicated key in each of the dic in A()
Dim I
For Each I In AyNz(A)
    PushDic DicAyMge, CvDic(I)
Next
End Function

Function DicAyAdd(A() As Dictionary) As Dictionary
Dim O As New Dictionary, D
For Each D In A
    PushDic O, CvDic(D)
Next
Set DicAyAdd = O
End Function

Function DicAyDr(DicAy, K) As Variant()
Dim U%: U = UB(DicAy)
Dim O()
ReDim O(U + 1)
Dim I, Dic As Dictionary, J%
J = 1
O(0) = K
For Each I In DicAy
   Set Dic = I
   If Dic.Exists(K) Then O(J) = Dic(K)
   J = J + 1
Next
DicAyDr = O
End Function
