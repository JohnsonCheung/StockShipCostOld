Attribute VB_Name = "MXls___Fun"
Option Compare Database
Option Explicit

Sub AyPutCol(A, At As Range)
Dim Sq()
Sq = AySqV(A)
RgReSz(At, Sq).Value = Sq
End Sub

Sub AyPutLoCol(A, Lo As ListObject, ColNm$)
Dim At As Range, C As ListColumn, R As Range
'AyDmp LoFny(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
AyPutCol A, At
End Sub

Sub AyPutRow(A, At As Range)
Dim Sq()
Sq = AySqH(A)
RgReSz(At, Sq).Value = Sq
End Sub

Function AyRgH(A, At As Range) As Range
Set AyRgH = SqRg(AySqH(A), At)
End Function

Function AyRgV(A, At As Range) As Range
Set AyRgV = SqRg(AySqV(A), At)
End Function

Function AyWs(A, Optional WsNm$) As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm)
SqRg AySqV(A), WsA1(O)
Set AyWs = O
End Function

Function AyabWs(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2", Optional LoNm$ = "AyAB") As Worksheet
Dim N&, AtA1 As Range, R As Range
N = Sz(A)
If N <> Sz(B) Then Stop
Set AtA1 = NewA1

AyRgH Array(N1, N2), AtA1
AyRgV A, AtA1.Range("A2")
AyRgV B, AtA1.Range("B2")
RgLo RgRCRC(AtA1, 1, 1, N + 1, 2)
Set AyabWs = AtA1.Parent
End Function


Function DicWb(A As Dictionary) As Workbook
'Assume each dic keys is name and each value is lines
'Prp-Wb is to create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Ass DicAllKeyIsNm(A)
Ass DicAllValIsStr(A)
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook
Set O = NewWb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        Set DicWb = O
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = LinesSqV(A(K))
Next
X: Set DicWb = O
End Function

Function DicWs(A As Workbook, Optional InclDicValTy As Boolean) As Worksheet
Set DicWs = DrsWs(DicDrs(A, InclDicValTy))
End Function

Function DicWsVis(A As Dictionary) As Worksheet
Dim O As Worksheet
   Set O = DicWs(A)
   WsVis O
Set DicWsVis = O
End Function

Function DtWs(A As Dt, Optional Vis As Boolean) As Worksheet
Dim O As Worksheet
Set O = NewWs(A.DtNm)
DrsLo DtDrs(A), WsA1(O)
Set DtWs = O
If Vis Then WsVis O
End Function

Function FmlNy(A$) As String()
FmlNy = MacroNy(A)
End Function

Sub LcSetTotLnk(A As ListColumn)
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = A.DataBodyRange
Set Ws = RgWs(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, R.Rows.Count + 1, 1)
Ws.Hyperlinks.Add Anchor:=R1, Address:="", SubAddress:=R2.Address
Ws.Hyperlinks.Add Anchor:=R2, Address:="", SubAddress:=R1.Address
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

Function LyWs(Ly$(), Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs()
AyRgV Ly, WsA1(O)
Set LyWs = O
End Function

Function MaxCol&()
Static C&, Y As Boolean
If Not Y Then
    Y = True
    C = IIf(CurXls.Version = "16.0", 16384, 255)
End If
MaxCol = C
End Function

Function MaxRow&()
Static R&, Y As Boolean
If Not Y Then
    Y = True
    R = IIf(CurXls.Version = "16.0", 1048576, 65535)
End If
MaxRow = R
End Function

Function N_SqH(N%) As Variant()
Dim O(), J%
ReDim O(1 To 1, 1 To N)
For J = 1 To N
    O(1, J) = J
Next
N_SqH = O
End Function

Function N_SqV(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
For J = 1 To N
    O(J, 1) = J
Next
N_SqV = O
End Function

Function N_ZerFill$(N, NDig%)
N_ZerFill = Format(N, String(NDig, "0"))
End Function

Private Function PjFfn$(A As VBProject)
On Error Resume Next
PjFfn = A.FileName
End Function

Function S1S2AyWs(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
Set S1S2AyWs = SqWs(S1S2AySq(A, Nm1, Nm2))
End Function

Private Sub ZZ_AyabWs()
Dim A, B
A = SslSy("A B C D E")
B = SslSy("1 2 3 4 5")
WsVis AyabWs(A, B)
Stop
End Sub

Private Sub ZZ_FbOupTblWb()
Dim W As Workbook
'Set W = FbOupTblWb(WFb)
WbVis W
Stop
W.Close False
Set W = Nothing
End Sub
