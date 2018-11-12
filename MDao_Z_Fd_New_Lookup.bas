Attribute VB_Name = "MDao_Z_Fd_New_Lookup"
Option Compare Database
Option Explicit
Const CMod$ = "MDao_Z_Fd_New_Lookup."

Function LookupFd(F, T, EF As EF) As DAO.Field2
Const CSub$ = CMod & "LookupFd"
Dim O As DAO.Field2, Ele$
Set LookupFd = StdFldFd(F, T): If IsSomething(LookupFd) Then Exit Function
Ele = LookupEle(F, EF.E): If Ele = "" Then ErWh CSub, "Fld cannot lookup from EF", "T F EDic FDic", T, F, EF.E, EF.F
Set LookupFd = StdEleFd(Ele, F): If IsSomething(LookupFd) Then Exit Function
If Not EF.F.Exists(Ele) Then ErWh CSub, "F's Ele is found, this Ele cannot be found FDic", "F Ele-of-F FDic", F, Ele, EF.F
Set LookupFd = EF.F(Ele)
End Function

Private Function LookupEle$(Fld, E As Dictionary) ' Return Ele$
Dim LikAy
For Each LikAy In E.Keys
    If IsInLikAy(Fld, CvSy(LikAy)) Then LookupEle = E(LikAy): Exit Function
Next
End Function
