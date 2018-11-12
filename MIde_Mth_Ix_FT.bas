Attribute VB_Name = "MIde_Mth_Ix_FT"
Option Compare Database
Option Explicit
Function SrcMthFTIxAy(A$()) As FTIx()
Dim Ix
For Each Ix In AyNz(SrcMthIxAy(A))
    PushObj SrcMthFTIxAy, FTIx(Ix, SrcMthIxTo(A, Ix))
Next
End Function

Function SrcMthNmFTIxAy(A$(), MthNm) As FTIx()
Dim IxAy&(), F&, T&
IxAy = SrcMthNmIxAy(A, MthNm)
Dim J%
For J = 0 To UB(IxAy)
   F = IxAy(J)
   T = SrcMthLx_ToLx(A, F)
   Push SrcMthNmFTIxAy, FTIx(F, T)
Next
End Function
