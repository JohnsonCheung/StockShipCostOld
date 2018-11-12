Attribute VB_Name = "MVb_Ay_Sub_Wh"
Option Compare Database
Option Explicit
Function SyWhFmTo(A$(), FmIx, ToIx) As String()
Dim J&
For J = FmIx To ToIx
    Push SyWhFmTo, A(J)
Next
End Function

Function AyWhFmTo(A, FmIx, ToIx)
If Sz(A) = 0 Then Exit Function
AyWhFmTo = AyCln(A)
FmIxToIxAss FmIx, ToIx, UB(A)
Dim J&
For J = FmIx To ToIx
    Push AyWhFmTo, A(J)
Next
End Function

Function AyWhIxAy(A, IxAy)
Dim U&
AyWhIxAy = AyCln(A)
Dim Ix
For Each Ix In AyNz(A)
    If 0 > Ix Or Ix > U Then Stop
    Push AyWhIxAy, A(Ix)
Next
End Function

Function AyWhPredXAP(A, PredXAP$, ParamArray Ap())
AyWhPredXAP = AyCln(A)
Dim I
Dim Av()
    Av = Ap
    Av = AyIns(Av)
For Each I In AyNz(A)
    Asg I, Av(0)
    If RunAv(PredXAP, Av) Then
        Push AyWhPredXAP, I
    End If
Next
End Function

Function AyWhEq3T(A, T1$, T2$, T3$) As String()
AyWhEq3T = AyWhPredXABC(A, "LinHas3T", T1, T2, T3)
End Function

Function AyWhDist(A)
AyWhDist = AyInto(AyCntDic(A).Keys, AyCln(A))
End Function

Function AyWhDup(A)
Dim D As Dictionary, I
AyWhDup = AyCln(A)
Set D = AyCntDic(A)
For Each I In AyNz(A)
    If D(I) > 1 Then
        Push AyWhDup, I
    End If
Next
End Function



Function AyWhFm(A, FmIx)
Dim O: O = A: Erase O
If 0 <= FmIx And FmIx <= UB(A) Then
    Dim J&
    For J = FmIx To UB(A)
        Push O, A(J)
    Next
End If
AyWhFm = O
End Function

Function AyWhFTIx(A, B As FTIx)
AyWhFTIx = AyWhFmTo(A, B.FmIx, B.ToIx)
End Function

Function AyWhFstNEle(Ay, N&)
Dim O: O = Ay
ReDim Preserve O(N - 1)
AyWhFstNEle = O
End Function

Function AyWhHasPfx(A, Pfx$) As String()
AyWhHasPfx = AyWhPfx(A, Pfx)
End Function

Function AyWhLik(A, Lik) As String()
Dim I
For Each I In AyNz(A)
    If I Like Lik Then PushI AyWhLik, I
Next
End Function

Function AyWhLikAy(A, LikAy$()) As String()
Dim I, Lik
For Each I In AyNz(A)
    For Each Lik In LikAy
        If I Like Lik Then
            PushI AyWhLikAy, I
            Exit For
        End If
    Next
Next
End Function

Function AyWhNm(A, B As WhNm) As String()
AyWhNm = AyExlLikssAy(AyWhRe(A, B.Re), B.ExlAy)
End Function

Function AyWhNotPfx(A, Pfx$) As String()
AyWhNotPfx = AyWhPredXPNot(A, "HasPfx", Pfx)
End Function

Function AyWhObjPred(A, Obj, Pred$)
Dim I, O, X
AyWhObjPred = AyCln(A)
For Each I In AyNz(A)
    X = CallByName(Obj, Pred, VbMethod, I)
    If X Then
        Push AyWhObjPred, I
    End If
Next
End Function

Function AyWhPatn(A, Patn$) As String()
If Sz(A) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then AyWhPatn = AySy(A): Exit Function
Dim X, R As RegExp
Set R = Re(Patn)
For Each X In AyNz(A)
    If R.Test(X) Then Push AyWhPatn, X
Next
End Function

Function AyWhPatnExl(A, Patn$, ExlLikss$) As String()
AyWhPatnExl = AyExlLikss(AyWhPatn(A, Patn), ExlLikss)
End Function

Function AyWhPatnIx(A, Patn$) As Long()
If Sz(A) = 0 Then Exit Function
Dim I, O&(), J&
Dim R As New RegExp
R.Pattern = Patn
For Each I In A
    If R.Test(I) Then Push O, J
    J = J + 1
Next
AyWhPatnIx = O
End Function

Function AyWhPfx(A, Pfx$) As String()
Dim I
For Each I In AyNz(A)
    If HasPfx(I, Pfx) Then PushI AyWhPfx, I
Next
End Function
Function AyWhPred(A, Pred$)
Dim X
AyWhPred = AyCln(A)
For Each X In AyNz(A)
    If Run(Pred, X) Then
        Push AyWhPred, X
    End If
Next
End Function

Function AyWhPredFalse(A, Pred$)
Dim X
AyWhPredFalse = AyCln(A)
For Each X In AyNz(A)
    If Not Run(Pred, X) Then
        Push AyWhPredFalse, X
    End If
Next
End Function

Function AyWhPredNot(A, Pred$)
AyWhPredNot = AyWhPredFalse(A, Pred)
End Function

Function AyWhPredXAB(Ay, XAB$, A, B)
Dim X
AyWhPredXAB = AyCln(Ay)
For Each X In AyNz(Ay)
    If Run(XAB, X, A, B) Then
        Push AyWhPredXAB, X
    End If
Next
End Function

Function AyWhPredXABC(Ay, XABC$, A, B, C)
Dim X
AyWhPredXABC = AyCln(Ay)
For Each X In AyNz(Ay)
    If Run(XABC, X, A, B, C) Then
        Push AyWhPredXABC, X
    End If
Next
End Function

Function AyWhPredXP(A, XP$, P)
Dim X
AyWhPredXP = AyCln(A)
For Each X In AyNz(A)
    If Run(XP, X, P) Then
        Push AyWhPredXP, X
    End If
Next
End Function

Function AyWhPredXPNot(A, XP$, P)
Dim X
AyWhPredXPNot = AyCln(A)
For Each X In AyNz(A)
    If Not Run(XP, X, P) Then
        Push AyWhPredXPNot, X
    End If
Next
End Function

Function AyWhRe(A, Re As RegExp) As String()
If IsNothing(Re) Then AyWhRe = AySy(A): Exit Function
Dim X
For Each X In AyNz(A)
    If Re.Test(X) Then PushI AyWhRe, X
Next
End Function

Function AyWhRmv3T(A, T1$, T2$, T3$) As String()
AyWhRmv3T = AyRmv3T(AyWhEq3T(A, T1, T2, T3))
End Function

Function AyWhRmvT1(A, T1$) As String()
AyWhRmvT1 = AyRmvT1(AyWhT1(A, T1))
End Function

Function AyWhRmvTT(A, T1$, T2$) As String()
AyWhRmvTT = AyRmvTT(AyWhTT(A, T1, T2))
End Function

Function AyWhSfx(A, Sfx$) As String()
Dim I
For Each I In AyNz(A)
    If HasSfx(I, Sfx) Then PushI AyWhSfx, I
Next
End Function

Function AyWhSingleEle(A)
Dim O: O = A: Erase O
Dim CntDry(): CntDry = AyCntDry(A)
If Sz(CntDry) = 0 Then
    AyWhSingleEle = O
    Exit Function
End If
Dim Dr
For Each Dr In CntDry
    If Dr(1) = 1 Then
        Push O, Dr(0)
    End If
Next
AyWhSingleEle = O
End Function

Function AyWhSng(A)
AyWhSng = AyMinus(A, AyWhDup(A))
End Function
Function AyWhSngEle(A)
'Return Set of Element as array in {Ay} having 2 or more element
Dim O: O = AyCln(A)
Dim K, D As Dictionary
Set D = AyCntDic(A)
For Each K In D.Keys
    If D(K) = 1 Then PushI O, K
Next
End Function

Function AyWhT1(A, V) As String()
AyWhT1 = AyWhPredXP(A, "LinHasT1", V)
End Function

Function AyWhT1InAy(A, Ay$()) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), L
For Each L In A
    If AyHas(Ay, LinT1(L)) Then Push O, L
Next
AyWhT1InAy = O
End Function

Function AyWhT1SelRst(A, T1) As String()
Dim L
For Each L In AyNz(A)
    If ShfTerm(L) = T1 Then PushI AyWhT1SelRst, L
Next
End Function
Function AyWhTTSelRst(A, T1, T2) As String()
Dim L, X1$, X2$, Rst$
For Each L In AyNz(A)
    Lin2TRstAsg L, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            PushI AyWhTTSelRst, Rst
        End If
    End If
Next
End Function

Function AyWhT2EqV(A$(), V) As String()
AyWhT2EqV = AyWhPredXP(A, "LinHasT2", V)
End Function

Function AyWhTT(A, T1$, T2$) As String()
AyWhTT = AyWhPredXAB(A, "LinHasTT", T1, T2)
End Function



Function AyWhNoEr(A, Msg$(), OEr$())
Dim J&
Erase OEr
If Not IsEqSzAy(A, Msg) Then Stop
For J = 0 To UB(A)
    If Msg(J) = "" Then Push AyWhNoEr, A(J) Else PushI OEr, Msg(J)
Next
End Function

Function AyWhDistT1(A) As String()
AyWhDistT1 = AyWhDist(AyTakT1(A))
End Function
