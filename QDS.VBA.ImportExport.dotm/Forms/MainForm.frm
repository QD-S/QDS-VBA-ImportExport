VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "QDS VBA Import/Export"
   ClientHeight    =   1830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal hKey&) As Long

Private Sub UserForm_Initialize()
    HasTypeFolderCheckBox.Value = VbaUtility.HasTypeFolder
    HasVbaFolderCheckBox.Value = VbaUtility.HasVbaFolder
    
    CodeTargetActiveWorkbookOptionButton.Value = True
    CodeTargetNameOptionButton_Change
End Sub

'////////////////////////////////////////////////////Command////////////////////////////////////////////////////

Private Sub ImportCommandButton_Click()
On Error GoTo ErrorHandle
    If CodeTargetAddInOptionButton.Value Then
        VbaUtility.ImportFilesOfAddIn True
    ElseIf CodeTargetActiveWorkbookOptionButton.Value Then
        VbaUtility.ImportFiles
    Else
        If InternalUtility.HasDocument(Documents, CodeTargetNameTextBox.Text) Then
            VbaUtility.ImportFiles Documents(CodeTargetNameTextBox.Text)
        Else
            MsgBox "'" + CodeTargetNameTextBox.Text + "' is not found.", , "Import"
        End If
    End If
    Exit Sub
ErrorHandle:
    ErrMsgBox
End Sub

Private Sub ExportCommandButton_Click()
On Error GoTo ErrorHandle
    If CodeTargetAddInOptionButton.Value Then
        VbaUtility.ExportFilesOfAddIn
    ElseIf CodeTargetActiveWorkbookOptionButton.Value Then
        VbaUtility.ExportFiles
    Else
        If InternalUtility.HasDocument(Documents, CodeTargetNameTextBox.Text) Then
            VbaUtility.ExportFiles Documents(CodeTargetNameTextBox.Text)
        Else
            MsgBox "'" + CodeTargetNameTextBox.Text + "' is not found.", , "Export"
        End If
    End If
    Exit Sub
ErrorHandle:
    ErrMsgBox
End Sub

'////////////////////////////////////////////////////Event////////////////////////////////////////////////////

Private Sub CodeTargetNameOptionButton_Change()
    If CodeTargetNameOptionButton.Value Then
        CodeTargetNameTextBox.Locked = False
        CodeTargetNameTextBox.BackColor = GetSystemColor(&H80000005)
    Else
        CodeTargetNameTextBox.Locked = True
        CodeTargetNameTextBox.BackColor = GetSystemColor(&H8000000F)
    End If
End Sub

Private Sub HasTypeFolderCheckBox_Change()
    VbaUtility.HasTypeFolder = HasTypeFolderCheckBox.Value
End Sub

Private Sub HasVbaFolderCheckBox_Change()
    VbaUtility.HasVbaFolder = HasVbaFolderCheckBox.Value
End Sub

'////////////////////////////////////////////////////Utility////////////////////////////////////////////////////

Private Sub ErrMsgBox()
    MsgBox "Error number: " & Err.Number & " " & Err.Description
End Sub

Function GetSystemColor(Value&)
    GetSystemColor = GetSysColor(Value And &HFF)
End Function
