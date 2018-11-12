VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LnkCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Nm$, Ty As DAO.DataTypeEnum, ExtNm$
Friend Property Get Init(Nm, Ty As DAO.DataTypeEnum, ExtNm$)
Me.Nm = Nm
Me.Ty = Ty
Me.ExtNm = ExtNm
Set Init = Me
End Property
