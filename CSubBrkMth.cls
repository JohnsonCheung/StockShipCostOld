VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSubBrkMth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public MdNm$
Public MthNm$
Public IsUsingCSub As Boolean
Public NeedDlt As Boolean
Public NeedIns As Boolean
Public OldLno As Long
Public OldCSub As String
Public NewLno As Long
Public NewCSub As String
Public Property Get Fldss$()
Fldss = "MdNm MthNm IsUsingCSub NeedDlt NeedIns OldLno OldCSub NewLno NewCSub"
End Property
