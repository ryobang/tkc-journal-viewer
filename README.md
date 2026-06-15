# TKC仕訳帳ビューアー (VSCode拡張)

TKC FX2（まいすたー）クラウドの仕訳データ、つまり読み込ませるExcel仕訳ファイル(.xlsx)とクラウドから出力した仕訳帳データ(.csv)を、Excelを開かずにVSCodeで仕訳帳の形で表示する読み取り専用のカスタムエディタ拡張です。

> TKC FX2（まいすたー）クラウド専用です。FX2の固定フォーマット（読み込み用xlsxの44列 / 仕訳帳データcsv）に合わせて作っており、他の会計ソフトの出力には対応していません。

記帳や巡回監査の前に、仕訳データの中身をざっと確かめたいとき用。Excelを起動するより速く、検索や月フィルタで目的の仕訳にすぐ飛べます。

## できること

- xlsx/xlsm/csv を右クリック →「TKC仕訳帳ビューアーで開く」で仕訳帳ビューに
- TKC読み込み用xlsx(44列) と FX2仕訳帳データcsv(仕訳帳データ) の両フォーマットに対応（列名の違いは自動吸収）
- 月日 / 伝票・証憑 / 借方・貸方の科目と金額 / 課税区分 / 取引先 / 摘要を一覧表示
- 補助コード・部門名・取引先T番号・消費税額・軽減税率(軽8%)も同じ行に表示
- 検索（科目名・取引先・摘要・伝票番号）と月フィルタ
- 複合仕訳を自動でグループ化して罫線で囲む
- 借方計と貸方計がずれていたら上部に貸借不一致の警告
- 表示中の仕訳をCSV出力（UTF-8 BOM付き）
- 文字サイズの拡大縮小、別モニターに出せる別ウィンドウ表示
- ファイルを書き換えると自動で再読込
- ライト / ダークテーマ対応

TKCの読み込み用フォーマットでないxlsx/csvを開いた場合は、その旨と検出したヘッダ一覧を表示します。csvの文字コードはUTF-8(BOM有無)・Shift_JISを自動判定します。

## インストール

PowerShellで1行:

```powershell
irm https://raw.githubusercontent.com/ryobang/tkc-journal-viewer/main/install.ps1 | iex
```

最新リリースの .vsix を取得して `code --install-extension` で入れます。事前にVSCodeの `code` コマンドがPATHに通っている必要があります（`Ctrl+Shift+P` →「Shell Command: Install 'code' command in PATH」）。

手動で入れる場合:

```powershell
code --install-extension tkc-journal-viewer-0.1.8.vsix
```

## 設定

- `tkcJournal.dateFormat` — 日付の表示形式（`MM/DD` / `M/D` / `YYYY-MM-DD`）
- `tkcJournal.zeroAmount` — ゼロ金額の表示（`blank` / `zero` / `dash`）

## 制限

- 読み取り専用。仕訳の編集はできません（元のxlsx/csvが正本）
- xlsxは1ファイルにつき先頭シートのみ対象
- 数式セルは計算結果が無いと0として扱う場合あり

## ライセンス

MIT
