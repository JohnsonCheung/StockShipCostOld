Attribute VB_Name = "MIde_Ens_CSub"
Option Compare Database
Option Explicit
Const CMod$ = "MIde_Ens_CSub."

Sub EnsCSub()
ZMdEns CurMd
End Sub

Sub EnsPjCSub()
ZPjEns CurPj
End Sub

Private Sub ZMdEns(A As CodeModule)
With MdCSubBrk(A)
    ZMdEns1 A, .MthBrkAy  '<== MthBrk must first
    ZMdEns2 A, .MdBrk
End With
End Sub

Private Sub ZMdEns1(A As CodeModule, B() As CSubBrkMth) _
'sfdf _
'lksjdf
'lksdjf
Const CSub$ = CMod & "ZMdEns1"
Const Trace As Boolean = True
Dim J%
ZMdEns1a MdNm(A), B ' Ass if in sorting order
For J = UB(B) To 0 Step -1
    With B(J)
        If .MdNm = "MIde_Ens_CSub" And .MthNm = "ZMdEns1" Then GoTo Nxt
        If .NeedDlt Then
            If A.Lines(.OldLno, 1) <> .OldCSub Then
                ErWh CSub, "OldCSub not expected", _
                    "Md Mth OldLno ExpCSub ActCSub", _
                    MdNm(A), .MthNm, .OldLno, .OldCSub, A.Lines(.OldLno, 1)
            End If
            A.DeleteLines .OldLno         '<==
        End If
        If .NeedIns Then
            A.InsertLines .NewLno, .NewCSub
        End If
        '
        If .NeedDlt Or .NeedIns Then
            MsgWh CSub, "CSub is ensured", B(J), "MdNm MthNm NeedDlt OldLno OldCSub NeedIns NewLno NewCSub"
        End If
    End With
Nxt:
Next
End Sub

Private Sub ZMdEns1a(MdNm$, B() As CSubBrkMth)
Const CSub$ = CMod & "ZMdEns1a"
If Sz(B) = 0 Then Exit Sub
Dim L1&, L2&
L1 = B(0).OldLno
L2 = B(0).NewLno
Dim J%
For J = 1 To UB(B)
    With B(J)
        If .OldLno > 0 Then
            If L1 > .OldLno Then
                Er CSub, "[Md] has [J] with [Prv-OldLno] > [Cur-OldLno].  CSubBrkMthAy not in sorted order", _
                    MdNm, J, L1, .OldLno
            End If
            L1 = .OldLno
        End If
        If L2 > .NewLno Then
            Er CSub, "[Md] has [J] with [Prv-NewLno] > [Cur-NewLno].  CSubBrkMthAy not in sorted order", _
                MdNm, J, L2, .NewLno
        End If
        L2 = .NewLno
    End With
Next
End Sub

Private Sub ZMdEns2(A As CodeModule, B As CSubBrkMd)
Const CSub$ = CMod & "ZMdEns2"
With B
    If .NeedDlt Then
        If A.Lines(.OldLno, 1) <> .OldCMod Then
            ErWh CSub, "Md CMod is not as expected", _
                "Md LNo OldCMod Exptecd-CMod", _
                MdNm(A), .OldLno, A.Lines(.OldLno, 1), .OldCMod
        End If
        A.DeleteLines .OldLno         '<==
    End If
    If .NeedIns Then
        A.InsertLines .NewLno, .NewCMod
    End If
    '
    If .NeedDlt Or .NeedIns Then
        MsgObjPrp CSub, "CMod is Update", B, "NeedDlt NeedIns OldLno OldCMod NewLno NewCMod"
    End If
End With
End Sub

Private Sub ZPjEns(A As VBProject)
Dim I
For Each I In PjMdAy(A)
   ZMdEns CvMd(I)
Next
End Sub
