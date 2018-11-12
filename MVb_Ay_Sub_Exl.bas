Attribute VB_Name = "MVb_Ay_Sub_Exl"
Option Compare Database
Option Explicit

Function AyExlAtCnt(A, Optional At = 0, Optional Cnt = 1)
If Cnt <= 0 Then Stop
If Sz(A) = 0 Then AyExlAtCnt = A: Exit Function
If At = 0 Then
    If Sz(A) = Cnt Then
        AyExlAtCnt = AyCln(A)
        Exit Function
    End If
End If
Dim U&: U = UB(A)
If At > U Then Stop
If At < 0 Then Stop
Dim O: O = A
Dim J&
If IsObject(A(0)) Then
    For J = At To U - Cnt
        Set O(J) = O(J + Cnt)
    Next
Else
    For J = At To U - Cnt
        O(J) = O(J + Cnt)
    Next
End If
ReDim Preserve O(U - Cnt)
AyExlAtCnt = O
End Function

Function AyExlDDLin(A) As String()
AyExlDDLin = AyWhPredFalse(A, "LinIsDDLin")
End Function

Function AyExlDotLin(A) As String()
AyExlDotLin = AyWhPredFalse(A, "LinIsDotLin")
End Function

Function AyExlEle(A, Ele)
Dim Ix&: Ix = AyIx(A, Ele): If Ix = -1 Then AyExlEle = A: Exit Function
AyExlEle = AyExlEleAt(A, AyIx(A, Ele))
End Function

Function AyExlEleAt(Ay, Optional At = 0, Optional Cnt = 1)
AyExlEleAt = AyExlAtCnt(Ay, At, Cnt)
End Function

Function AyExlEleLik(A, Lik$) As String()
If Sz(A) = 0 Then Exit Function
Dim J&
For J = 0 To UB(A)
    If A(J) Like Lik Then AyExlEleLik = AyExlEleAt(A, J): Exit Function
Next
End Function

Function AyExlEmpEle(A)
Dim O: O = AyCln(A)
If Sz(A) > 0 Then
    Dim X
    For Each X In AyNz(A)
        PushNonEmp O, X
    Next
End If
AyExlEmpEle = O
End Function

Function AyExlEmpEleAtEnd(A)
Dim LasU&, U&
Dim O: O = A
For LasU = UB(A) To 0 Step -1
    If Not IsEmp(O(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase O
Else
    ReDim Preserve O(LasU)
End If
AyExlEmpEleAtEnd = O
End Function

Function AyExlFmTo(A, FmIx, ToIx)
Dim U&
U = UB(A)
If 0 > FmIx Or FmIx > U Then ErWh CSub, "[FmIx] is out of range", "Ay U FmIx ToIx", A, UB(A), FmIx, ToIx
If FmIx > ToIx Or ToIx > U Then ErWh CSub, "[ToIx] is out of range", "Ay U FmIx ToIx", A, UB(A), FmIx, ToIx
Dim O
    O = A
    Dim I&, J&
    I = 0
    For J = ToIx + 1 To U
        O(FmIx + I) = O(J)
        I = I + 1
    Next
    Dim Cnt&
    Cnt = ToIx - FmIx + 1
    ReDim Preserve O(U - Cnt)
AyExlFmTo = O
End Function

Function AyExlFstEle(A)
AyExlFstEle = AyExlEleAt(A)
End Function

Function AyExlFstNEle(A, N)
Dim O: O = A
ReDim O(N - 1)
Dim J&
For J = 0 To UB(A) - N
    O(J) = A(N + J)
Next
AyExlFstNEle = O
End Function

Function AyExlFTIx(A, B As FTIx)
With B
    AyExlFTIx = AyExlFmTo(A, .FmIx, .ToIx)
End With
End Function

Function AyExlIxAy(A, IxAy)
'IxAy holds index if A to be remove.  It has been sorted else will be stop
Ass AyIsSrt(A)
Ass AyIsSrt(IxAy)
Dim J&
Dim O: O = A
For J = UB(IxAy) To 0 Step -1
    O = AyExlEleAt(O, CLng(IxAy(J)))
Next
AyExlIxAy = O
End Function

Function AyExlLasEle(A)
AyExlLasEle = AyExlEleAt(A, UB(A))
End Function

Function AyExlLasNEle(A, Optional NEle% = 1)
Dim O: O = A
Select Case Sz(A)
Case Is > NEle:    ReDim Preserve O(UB(A) - NEle)
Case NEle: Erase O
Case Else: Stop
End Select
AyExlLasNEle = O
End Function

Function AyExlLik(A, Lik) As String()
Dim I
For Each I In AyNz(A)
    If Not I Like Lik Then PushI AyExlLik, I
Next
End Function

Function AyExlLikAy(A, LikAy$()) As String()
Dim I
For Each I In AyNz(A)
    If Not IsInLikAy(I, LikAy) Then Push AyExlLikAy, I
Next
End Function

Function AyExlLikss(A, Likss$) As String()
AyExlLikss = AyExlLikAy(A, SslSy(Likss))
End Function

Function AyExlLikssAy(A, LikssAy$()) As String()
If Sz(LikssAy) = 0 Then AyExlLikssAy = AySy(A): Exit Function
Dim Likss
For Each Likss In AyNz(A)
    If Not IsInLikss(X, Likss) Then PushI AyExlLikssAy, X
Next
End Function

Function AyExlNeg(A)
Dim I
AyExlNeg = AyCln(A)
For Each I In AyNz(A)
    If I >= 0 Then
        PushI AyExlNeg, I
    End If
Next
End Function

Function AyExlNEle(A, Ele, Cnt%)
If Cnt <= 0 Then Stop
AyExlNEle = AyCln(A)
Dim X, C%
C = Cnt
For Each X In AyNz(A)
    If C = 0 Then
        PushI AyExlNEle, X
    Else
        If X <> Ele Then
            Push AyExlNEle, X
        Else
            C = C - 1
        End If
    End If
Next
X:
End Function

Function AyExlOneTermLin(A) As String()
AyExlOneTermLin = AyWhPredFalse(A, "LinIsOneTermLin")
End Function

Function AyExlPfx(A, ExlPfx$) As String()
Dim I
For Each I In AyNz(A)
    If Not HasPfx(I, ExlPfx) Then PushI AyExlPfx, I
Next
End Function

Function AyExlT1Ay(A, ExlT1Ay0) As String()
'Exclude those Lin in Array-A its T1 in ExlT1Ay0
Dim Exl$(): Exl = CvNy(ExlT1Ay0): If Sz(Exl) = 0 Then Stop
Dim L
For Each L In AyNz(A)
    If Not AyHas(Exl, LinT1(L)) Then
        PushI AyExlT1Ay, L
    End If
Next
End Function

Private Sub Z()
Z_AyExlAtCnt
Z_AyExlEmpEleAtEnd
Z_AyExlFTIx
Z_AyExlFTIx1
Z_AyExlIxAy
End Sub

Private Sub Z_AyExlAtCnt()
Dim A()
A = Array(1, 2, 3, 4, 5)
Ept = Array(1, 4, 5)
GoSub Tst
'
Exit Sub

Tst:
    Act = AyExlAtCnt(A, 1, 2)
    C
    Return
End Sub

Private Sub Z_AyExlEmpEleAtEnd()
Dim A: A = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyExlEmpEleAtEnd(A)
Ass Sz(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub Z_AyExlFTIx()
Dim A
Dim FTIx1 As FTIx
Dim Act
A = SplitSpc("a b c d e")
Set FTIx1 = FTIx(1, 2)
Act = AyExlFTIx(A, FTIx1)
Ass Sz(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AyExlFTIx1()
Dim A
Dim Act
A = SplitSpc("a b c d e")
Act = AyExlFTIx(A, FTIx(1, 2))
Ass Sz(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AyExlIxAy()
Dim A(), IxAy
A = Array("a", "b", "c", "d", "e", "f")
IxAy = Array(1, 3)
Ept = Array("a", "c", "e", "f")
GoSub Tst
Exit Sub
Tst:
    Act = AyExlIxAy(A, IxAy)
    C
    Return
End Sub
