VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "QDS VBA Import/Export"
   ClientHeight    =   1830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895
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
    HasTypeFolderCheckBox.Value = True
    HasVbaFolderCheckBox.Value = True
    
    CodeTargetActiveWorkbookOptionButton.Value = True
    CodeTargetNameOptionButton_Change
End Sub

'////////////////////////////////////////////////////Command////////////////////////////////////////////////////

Private Sub ImportCommandButton_Click()
On Error GoTo ErrorHandle
    If MsgBox("Do you import the VB component files?", vbYesNo, "Import") = vbYes Then
        If CodeTargetAddInOptionButton.Value Then
            VbaUtility.ImportFilesOfAddIn True
        ElseIf CodeTargetActiveWorkbookOptionButton.Value Then
            VbaUtility.ImportFiles , , HasTypeFolderCheckBox.Value, , CodeDirectoryFormat
        Else
            VbaUtility.ImportFiles Workbooks(CodeTargetNameTextBox.Text), , HasTypeFolderCheckBox.Value, , CodeDirectoryFormat
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
        VbaUtility.ExportFiles , , HasTypeFolderCheckBox.Value, CodeDirectoryFormat
    Else
        VbaUtility.ExportFiles Workbooks(CodeTargetNameTextBox.Text), , HasTypeFolderCheckBox.Value, CodeDirectoryFormat
    End If
    Exit Sub
ErrorHandle:
    ErrMsgBox
End Sub

Private Sub ArrangeCodeCommandButton_Click()
On Error GoTo ErrorHandle
    If CodeTargetAddInOptionButton.Value Then
        VbaUtility.ArrangeCodeOfAddIn
    ElseIf CodeTargetActiveWorkbookOptionButton.Value Then
        VbaUtility.ArrangeCode
    Else
        VbaUtility.ArrangeCode Workbooks(CodeTargetNameTextBox.Text)
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

'////////////////////////////////////////////////////Utility////////////////////////////////////////////////////

Private Property Get CodeDirectoryFormat()
    If HasVbaFolderCheckBox.Value Then CodeDirectoryFormat = VbaUtility.FileNameFormatText + ".vba" Else CodeDirectoryFormat = ""
End Property

Private Sub ErrMsgBox()
    MsgBox "Error number: " & Err.Number & " " & Err.Description
End Sub

Function GetSystemColor(Value&)
    GetSystemColor = GetSysColor(Value And &HFF)
End Function
