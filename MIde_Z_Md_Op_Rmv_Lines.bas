Attribute VB_Name = "MIde_Z_Md_Op_Rmv_Lines"
Option Compare Database
Option Explicit
Sub MdRmvBdy(A As CodeModule)
MdRmvFmCnt A, MdBdyFmCnt(A)
End Sub

Sub MdRmvDcl(A As CodeModule)
If A.CountOfDeclarationLines = 0 Then Exit Sub
A.DeleteLines 1, A.CountOfDeclarationLines
End Sub

Sub MdRmvEndBlankLin(A As CodeModule)
Dim J%
While A.CountOfLines > 1
    J = J + 1
    If J > 10000 Then Stop
    If Trim(A.Lines(A.CountOfLines, 1)) <> "" Then Exit Sub
    A.DeleteLines A.CountOfLines, 1
Wend
End Sub

Sub MdRmvFC(A As CodeModule, B() As FmCnt)
If Not FmCntAyIsInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub

Sub MdRmvLines(A As CodeModule)
If A.CountOfLines = 0 Then Exit Sub
A.DeleteLines 1, A.CountOfLines
End Sub

Sub MdClr(A As CodeModule, Optional IsSilent As Boolean)
With A
    If .CountOfLines = 0 Then Exit Sub
    If Not IsSilent Then Debug.Print FmtQQ("MdClr: Md(?) of lines(?) is cleared", MdNm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub

Sub MdClrBdy(A As CodeModule, Optional IsSilent As Boolean)
Stop
With A
    If .CountOfLines = 0 Then Exit Sub
    Dim N%, Lno%
        Lno = MdBdyFmLno(A)
        N = .CountOfLines - Lno + 1
    If N > 0 Then
        If Not IsSilent Then Debug.Print FmtQQ("MdClrBdy: Md(?) of lines(?) from Lno(?) is cleared", MdNm(A), N, Lno)
        .DeleteLines Lno, N
    End If
End With
End Sub

Function MdLno_Rmv(A As CodeModule, Lno)
If Lno = 0 Then Exit Function
MsgDmp "MdLno_Rmv: [Md]-[Lno]-[Lin] is removed", MdNm(A), Lno, A.Lines(Lno, 1)
A.DeleteLines Lno, 1
End Function


Sub MdEndTrim(A As CodeModule, Optional ShwMsg As Boolean)
If A.CountOfLines = 0 Then Exit Sub
Dim N$: N = MdDNm(A)
Dim J%
While Trim(A.Lines(A.CountOfLines, 1)) = ""
    If ShwMsg Then FunMsgLinDmp "MdEndTrim", "[LinNo] in [Md]", A.CountOfLines, N
    A.DeleteLines A.CountOfLines, 1
    If A.CountOfLines = 0 Then Exit Sub
    If J > 1000 Then Stop
    J = J + 1
Wend
End Sub


Sub MdFmCntDlt(A As CodeModule, B() As FmCnt)
If Not FmCntAyIsInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub
