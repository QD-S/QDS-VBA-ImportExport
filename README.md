# QDS VBA Code Import/Export Tool
A simple Office add-in for VBA code import and export.

<p align="center">
  <img src="https://github.com/QD-S/QDS-VBA-ImportExport/blob/main/MD/MainForm.png">
</p>

## Set up

Here we provide an Office add-in to import and export VBA code.
QDS.VBA.ImportExport.xlam and QDS.VBA.ImportExport.dotm are add-ins for Excel and Word VBA code respectively. They use VBComponent to import and export code. Therefore, to use it, you have to enable the check of "Trust access to the VBA project object model" in the "Trust Center" of Excel and Word as shown below.

<p align="center">
  <img src="https://github.com/QD-S/QDS-VBA-ImportExport/blob/main/MD/ExcelTrustCenter.png">
  <img src="https://github.com/QD-S/QDS-VBA-ImportExport/blob/main/MD/WordTrustCenter.png">
</p>

## How to use

Open the add-in. Help is displayed as a tooltip.

### Export

1. Activate the office file you want to export the VBA code to.

1. Run OpenQdsVbaImportExportMainForm in the add-in's Utility_ module to display the MainForm.

1. Press the "Export" button. The VBA code will be exported to the same folder of the target office file. You can target non-active files by setting a file name in the Name text box.

### Import

1. Activate the office file where you want to import the VBA code.

1. Run OpenQdsVbaImportExportMainForm in the add-in's Utility_ module to display the MainForm.

1. Press the "Import" button. The VBA code will be imported into the same folder of the target office file. You can target non-active files by setting a file name in the Name text box.

### Folder structure (Check Box)
You can change the import/export folder structure by the following settings.

#### Type Folder (Check Box)

Output each file to the specified folder below.

| File Extension | Folder Name |
|:------------|:------------|
| cls | Classes |
| bas (Module) | Modules |
| bas (Sheet/Book) | Objects |
| frm | Forms |

#### VBA Folder (Check Box)

Export to a folder with a ".vba" suffix.

### Others

#### Arrange (Button)

Removes empty lines before and after VBA code.

#### AddIn (Option Button)

Export this add-in VBA code.

#### IsCommonVbComponent (Code)

Add the following line to the VBA code for importing and exporting to the upper folder. This allows you to share code between different files in the same folder.

```vb
Private Const IsCommonVbComponent = True
```

#### Charset (Code)

"UTF-8" is used if DefaultCharset is empty. You can set own char set like "Shift-JIS" in DefaultCharset of the Utility_ module.

```vb
Public Const DefaultCharset$ = ""
```

