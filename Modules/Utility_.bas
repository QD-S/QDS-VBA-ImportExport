Attribute VB_Name = "Utility_"
Option Explicit

Private Const IsCommonVbComponent = True

' "UTF-8" is used if DefaultCharset is empty. You can set own char set like "Shift-JIS" in DefaultCharset.
Public Const DefaultCharset$ = ""

Public InternalUtility As New InternalUtility
Public VbaUtility As New VbaUtility

Sub OpenQdsVbaImportExportMainForm()
    MainForm.Show vbModeless
End Sub
