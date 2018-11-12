Attribute VB_Name = "MVb_Fs_Sel"
Option Explicit
Option Compare Database

Function FfnSel$(A, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
With Application.FileDialog(msoFileDialogFilePicker)
    .Filters.Clear
    .Title = Tit
    .AllowMultiSelect = False
    .Filters.Add "", FSpec
    .InitialFileName = A
    .ButtonName = BtnNm
    .Show
    If .SelectedItems.Count = 1 Then
        FfnSel = .SelectedItems(1)
    End If
End With
End Function


Function PthSel$(A$, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .InitialFileName = IIf(IsNull(A), "", A)
    .Show
    If .SelectedItems.Count = 1 Then
        PthSel = PthEnsSfx(.SelectedItems(1))
    End If
End With
End Function

Private Sub Z_PthSel()
GoTo ZZ
ZZ:
MsgBox FfnSel("C:\")
End Sub

