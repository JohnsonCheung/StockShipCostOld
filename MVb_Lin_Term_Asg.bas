Attribute VB_Name = "MVb_Lin_Term_Asg"
Option Compare Database
Option Explicit

Sub Lin2TRstAsg(A, OT1, OT2, ORst)
AyAsg Lin2TRst(A), OT1, OT2, ORst
End Sub

Sub Lin3TRstAsg(A, OT1, OT2, OT3, ORst)
AyAsg Lin3TRst(A), OT1, OT2, OT3, ORst
End Sub

Sub LinTRstAsg(A, OT1, ORst)
AyAsg LinNTermRst(A, 1), OT1, ORst
End Sub

Sub LinTTAsg(A, O1, O2)
AyAsg LinTT(A), O1, O2
End Sub
