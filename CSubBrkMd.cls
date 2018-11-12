VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSubBrkMd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public MdNm$
Public NeedDlt As Boolean
Public NeedIns As Boolean
Public OldLno%
Public NewLno%
Public NewCMod$
Public OldCMod$
Public IsUsingCSub As Boolean
Public Property Get Fldss$()
Fldss = "MdNm IsUsingCSub NeedDlt NeedIns OldLno OldCMod NewLno NewCMod"
End Property

