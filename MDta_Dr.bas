Attribute VB_Name = "MDta_Dr"
Option Compare Database
Option Explicit

Sub DrSetSqRow(A, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
Dim J%, I
If NoTxtSngQ Then
    For Each I In A
        J = J + 1
        Sq(R, J) = I
    Next
    Exit Sub
End If
For Each I In A
    J = J + 1
    If IsStr(I) Then
        Sq(R, J) = "'" & I
    Else
        Sq(R, J) = I
    End If
Next
End Sub
