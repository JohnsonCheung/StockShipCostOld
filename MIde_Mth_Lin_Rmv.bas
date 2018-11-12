Attribute VB_Name = "MIde_Mth_Lin_Rmv"
Option Compare Database
Option Explicit
Function RmvMdy$(A)
RmvMdy = LTrim(RmvPfxAySpc(A, MdyAy))
End Function

Function RmvMthTy$(A)
RmvMthTy = RmvPfxAySpc(A, MthTyAy)
End Function
