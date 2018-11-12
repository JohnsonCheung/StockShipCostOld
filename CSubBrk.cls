VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSubBrk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A_MthBrkAy() As CSubBrkMth
Private A_MdBrk As CSubBrkMd
Property Get Init(MdBrk As CSubBrkMd, MthBrkAy() As CSubBrkMth) As CSubBrk
A_MthBrkAy = MthBrkAy
Set A_MdBrk = MdBrk
Set Init = Me
End Property
Property Get MthBrkAy() As CSubBrkMth()
MthBrkAy = A_MthBrkAy
End Property
Property Get MdBrk() As CSubBrkMd
Set MdBrk = A_MdBrk
End Property
