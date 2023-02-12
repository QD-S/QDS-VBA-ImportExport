# QDS VBA Code Import/Export Tool
VBAコードのインポート・エクスポートを行う簡易オフィスアドイン

<p align="center">
  <img src="https://github.com/QD-S/QDS-VBA-ImportExport/blob/main/MD/MainForm.png">
</p>

## セットアップ

ここではVBAコードのインポートとエクスポートを行うOfficeアドインを提供します。
QDS.VBA.ImportExport.xlam及びQDS.VBA.ImportExport.dotmは、それぞれエクセル及びワードのVBAコードのためのアドインです。これらはVBComponentを使用しています。そのため、使用するにはエクセル及びワードそれぞれの"トラストセンター"で下記のように「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」のチェックを有効にする必要があります。

<p align="center">
  <img src="https://github.com/QD-S/QDS-VBA-ImportExport/blob/main/MD/ExcelTrustCenter.png">
  <img src="https://github.com/QD-S/QDS-VBA-ImportExport/blob/main/MD/WordTrustCenter.png">
</p>

:warning: $\textcolor{red}{日本語}$で使用する場合は、一番下のCharset (コード)を参照し、DefaultCharsetを"Shift-JIS"に変更してください。 

## 使用方法

対象のアドインを開きます。 ヘルプはツールチップとして表示されます。

### エクスポート

1. VBAコードをエクスポートしたいオフィスファイルをアクティブにしてください。

1. アドインのUtility_モジュールにあるOpenQdsVbaImportExportMainFormを実行して、MainFormを表示します。

1. 「Export」ボタンを押してください。対象オフィスファイルと同じフォルダにVBAコードがエクスポートされます。Nameを選択し、ファイル名を入力することで、アクティブ以外のファイルを対象にすることができます。

### インポート

1. VBAコードをインポートしたいオフィスファイルをアクティブにしてください。

1. アドインのUtility_モジュールにあるOpenQdsVbaImportExportMainForm関数を実行して、MainFormを表示します。

1. 「Import」ボタンを押してください。対象オフィスファイルと同じフォルダのVBAコードがインポートされます。Nameを選択し、ファイル名を入力することで、アクティブ以外のファイルを対象にすることができます。

### フォルダ構造 (チェックボックス)

以下の設定によりインポート・エクスポートのフォルダ構造を変更することができます。

#### Type Folder (チェックボックス)

それぞれのファイルについて、下記の指定されたフォルダに出力します。

| 拡張子 | フォルダ名 |
|:------------|:------------|
| cls | Classes |
| bas (Module) | Modules |
| bas (Sheet/Book) | Objects |
| frm | Forms |

#### VBA Folder (チェックボックス)

ファイル名に".vba"の接尾語がついたフォルダに出力します。

### その他

#### Arrange (ボタン)

VBAコードの前後の空の改行を削除します。

#### AddIn (オプションボタン)

このアドインVBAコードを出力します。

#### IsCommonVbComponent (コード)

以下の行をVBAコードに追加すると一つ上のフォルダにインポート・エクスポートします。これにより、同じフォルダ内の異なるファイルの間でコードを共有することができます。

```vb
Private Const IsCommonVbComponent = True
```

#### Charset (コード)

DefaultCharsetが空文字の場合、"UTF-8"が文字コードとして使用されます。　日本語の場合は、下記のようにUtility_のモジュールのDefaultCharsetに"Shift-JIS"を設定してください。

```vb
Public Const DefaultCharset$ = "Shift-JIS"
```
