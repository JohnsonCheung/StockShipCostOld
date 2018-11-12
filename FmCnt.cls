VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_FmLno&, A_Cnt&
Property Get FmLno&()
FmLno = A_FmLno
End Property
Property Get Cnt&()
Cnt = A_Cnt
End Property
Friend Property Get Init(FmLno, Cnt) As FmCnt
If FmLno <= 0 Then Stop
If Cnt <= 0 Then Stop
A_FmLno = FmLno
A_Cnt = Cnt
Set Init = Me
End Property
