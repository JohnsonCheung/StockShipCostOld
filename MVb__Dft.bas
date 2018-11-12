Attribute VB_Name = "MVb__Dft"
Option Compare Database
Option Explicit
Function Dft(V, DftV)
If IsEmp(V) Then
   Dft = DftV
Else
   Dft = V
End If
End Function

Function DftDbNm$(DbNm0$, Db As Database)
If DbNm0 = "" Then
    DftDbNm = FfnFnn(Db.Name)
Else
    DftDbNm = DbNm0
End If
End Function

Function DftF0(A)
DftF0 = IIf(IsMissing(A), 0, A)
End Function

Function DftFb$(A$)
If A = "" Then
   Dim O$: O = TmpFb
   DAO.DBEngine.CreateDatabase(O, dbLangGeneral).Close
   DftFb = O
Else
   DftFb = A
End If
End Function

Function DftFx$(A$)
If A = "" Then
   Dim O$: O = TmpFx
   DftFx = O
Else
   DftFx = A
End If
End Function

Function DftLy(Ly0) As String()
If IsStr(Ly0) Then
   DftLy = SplitVBar(Ly0)
   Exit Function
End If
If IsArray(Ly0) Then
   DftLy = AySy(Ly0)
End If
End Function

Function DftPfxAy(PfxAyVbl0)
If IsArray(PfxAyVbl0) Then DftPfxAy = PfxAyVbl0: Exit Function
DftPfxAy = SplitVBar(PfxAyVbl0)
End Function

Function DftStr$(A, Dft)
DftStr = IIf(A = "", Dft, A)
End Function

Function DftTpLy(Tp0) As String()
Select Case True
Case IsStr(Tp0): DftTpLy = SplitVBar(Tp0)
Case IsSy(Tp0):  DftTpLy = Tp0
Case Else: Stop
End Select
End Function
