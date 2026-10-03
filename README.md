# 一键拆表 | Excel Splitter | Excel 分割ツール

[中文](#中文) · [English](#english) · [日本語](#日本語)

## 中文

**一键拆表 V4.8** 是 Windows 桌面版 Excel 拆表工具。选择工作簿、工作表和拆分依据列，即可按列值生成新工作簿中的多个工作表，或导出多个独立的 `.xlsx` 文件。源文件不会被修改。

**[下载最新版 V4.8 EXE](./一键拆表-V4.8.exe)** · [使用方法](./一键拆表-V4.8-使用方法.txt) · [Python 源码](./一键拆表-V4.8-源码.py)

### 主要功能

- 两种拆分方式：保存为一个新工作簿中的多个工作表；或每个列值保存为一个独立文件。
- 可指定数据起始行；拆成多个文件时，默认保留源工作簿的其他工作表，也可以关闭该选项。
- 支持 `.xlsx`、`.xlsm`、`.xlsb`。读取 `.xlsb` 时，本机需要安装可通过 COM 调用的 WPS 表格或 Microsoft Excel。
- 输出为 `.xlsx`，因此 `.xlsm` / `.xlsb` 中的宏不会保留。旧版 `.xls` 请先另存为受支持的格式。

### 快速使用

1. 下载并运行 V4.8 EXE，选择 Excel 文件和要拆分的工作表。
2. 选择拆分依据列；如果有多行表头，填写第一行数据的行号。
3. 选择输出方式，点击“一键拆分”。

V4.8 修复了同名结果覆盖、保留其他工作表时命名区域丢失、相对路径超链接误报、选错文件后沿用旧数据，以及拆分表图片或图表丢失等问题。详细说明见 [V4.8 使用方法](./一键拆表-V4.8-使用方法.txt)。

## English

**Excel Splitter V4.8** is a Windows desktop tool that splits an Excel worksheet by the values in a selected column. Create multiple sheets in a new workbook or export one `.xlsx` file per value. The original workbook is left unchanged.

**[Download V4.8 for Windows](./一键拆表-V4.8.exe)** · [User guide (Chinese)](./一键拆表-V4.8-使用方法.txt) · [Python source code](./一键拆表-V4.8-源码.py)

### Features

- Choose between a new workbook with separate sheets and separate files for each column value.
- Set the first data row when your worksheet has multiple header rows. When exporting separate files, other sheets from the source workbook are included by default; you can turn this off.
- Supports `.xlsx`, `.xlsm`, and `.xlsb`. Reading `.xlsb` requires WPS Spreadsheets or Microsoft Excel installed locally with COM access.
- Output files use `.xlsx`, so macros from `.xlsm` or `.xlsb` are not retained. Convert legacy `.xls` files to a supported format first.

### Quick start

1. Download and run the V4.8 EXE. Select your Excel file and the worksheet to split.
2. Select the column to split by. If the sheet has multiple header rows, enter the row number of the first data row.
3. Choose an output mode and click **一键拆分** (Split).

V4.8 fixes output filename collisions, lost defined names when preserving other sheets, false errors on relative hyperlinks, reuse of previously selected data after an invalid file selection, and missing images or charts on split sheets. See the [V4.8 guide](./一键拆表-V4.8-使用方法.txt) for details.

## 日本語

**一键拆表 V4.8** は、選択した列の値ごとに Excel のワークシートを分割する Windows 用デスクトップツールです。結果を新しいブック内の複数シート、または値ごとの個別の `.xlsx` ファイルとして保存できます。元のブックは変更しません。

**[Windows 用 V4.8 をダウンロード](./一键拆表-V4.8.exe)** · [使い方（中国語）](./一键拆表-V4.8-使用方法.txt) · [Python ソースコード](./一键拆表-V4.8-源码.py)

### 主な機能

- 新しいブック内の複数シートに分割する方法と、列の値ごとに個別ファイルを作る方法を選べます。
- 見出しが複数行ある場合は、データの開始行を指定できます。個別ファイルへの出力では、元ブックの他のシートを既定で引き継ぎます。この設定はオフにできます。
- `.xlsx`、`.xlsm`、`.xlsb` に対応しています。`.xlsb` の読み込みには、COM から利用できる WPS Spreadsheets または Microsoft Excel のインストールが必要です。
- 出力形式は `.xlsx` のため、`.xlsm` / `.xlsb` のマクロは引き継がれません。旧形式の `.xls` は、先に対応形式で保存してください。

### 使い方

1. V4.8 の EXE をダウンロードして起動し、Excel ファイルと分割するシートを選びます。
2. 分割の基準にする列を選びます。見出しが複数行ある場合は、最初のデータ行の行番号を入力します。
3. 出力方法を選び、**一键拆分**（分割）をクリックします。

V4.8 では、同名ファイルの上書き、他のシートを残す際の定義名の欠落、相対パスのハイパーリンクに関する誤判定、無効なファイル選択後に以前のデータが使われる問題、分割対象シートの画像・グラフの欠落を修正しました。詳細は [V4.8 の説明書（中国語）](./一键拆表-V4.8-使用方法.txt) を参照してください。

## Versions / 历史版本 / 過去のバージョン

For everyday use, choose V4.8. / 日常使用建议选择 V4.8。 / 通常の利用には V4.8 をお勧めします。

| Version / 版本 | Download / 下载 | Guide / 说明 | Source / 源码 |
| --- | --- | --- | --- |
| **V4.8 (latest / 最新)** | [EXE](./一键拆表-V4.8.exe) | [Guide](./一键拆表-V4.8-使用方法.txt) | [Python](./一键拆表-V4.8-源码.py) |
| V4.7 | [EXE](./一键拆表-V4.7.exe) | [Guide](./一键拆表-V4.7-使用方法.txt) | [Python](./一键拆表-V4.7-源码.py) |
| V4.6 | [EXE](./一键拆表-V4.6.exe) | [Guide](./一键拆表-V4.6-使用方法.txt) | [Python](./一键拆表-V4.6-源码.py) |
| V4.5 | [EXE](./一键拆表-V4.5.exe) | [Guide](./一键拆表-V4.5-使用方法.txt) | [Python](./一键拆表-V4.5-源码.py) |
| V4.4 | [EXE](./一键拆表-V4.4.exe) · [RAR](./一键拆表-V4.4.rar) | [Guide](./一键拆表-V4.4-使用方法.txt) | [Python](./一键拆表-V4.4-源码.py) |
| V4.3 | [EXE](./一键拆表-V4.3.exe) · [RAR](./一键拆表-V4.3.rar) | [Guide](./一键拆表-V4.3-使用方法.txt) | — |
| V4.2 | [EXE](./一键拆表-V4.2.exe) | [Guide](./一键拆表-V4.2-使用方法.txt) | — |
| V4.1 | [EXE](./一键拆表-V4.1.exe) | [Guide](./一键拆表-V4.1-使用方法.txt) | — |
| V4 | [EXE](./一键拆表-V4.exe) | [Guide](./一键拆表-V4-使用方法.txt) | — |
| V3 | [EXE](./一键拆表-V3.exe) · [RAR](./一键拆表-V3.rar) | [Guide](./一键拆表-V3-使用方法.txt) | — |
| V2 | [RAR](./一键拆表-V2.rar) | — | — |

V4.8 EXE SHA-256:

```text
D0B74AB0D35917B53188FEF680887007A7BE3BB07046566987864F6234D7567E
```
