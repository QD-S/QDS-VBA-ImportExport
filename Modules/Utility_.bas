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

Sub QdsVbaImportActiveTarget()
    VbaUtility.ImportFiles
End Sub

Sub QdsVbaExportActiveTarget()
    VbaUtility.ExportFiles
End Sub

Sub ShowQdsVbaImportExportMainFormOnRibbon(control As IRibbonControl)
    MainForm.Show vbModeless
End Sub

Sub QdsVbaImportActiveTargetOnRibbon(control As IRibbonControl)
    VbaUtility.ImportFiles
End Sub

Sub QdsVbaExportActiveTargetOnRibbon(control As IRibbonControl)
    VbaUtility.ExportFiles
End Sub
