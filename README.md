excl
====

これはexcelコントロール用のライブラリ

[![godoc](https://godoc.org/github.com/loadoff/excl?status.svg)](https://godoc.org/github.com/loadoff/excl)
[![CircleCI](https://circleci.com/gh/loadoff/excl.svg?style=svg)](https://circleci.com/gh/loadoff/excl)
[![go report](https://goreportcard.com/badge/github.com/loadoff/excl)](https://goreportcard.com/report/github.com/loadoff/excl)

## Description

基本的にもとのexcelファイルを破壊せずにデータの入力を行うためのライブラリです。
また大量のデータを扱う上でも優位になるように開発を行います。

## Usage

既存のExcelファイルを操作
```go
// Excelファイルを読み込み
w, _ := excl.Open("path/to/read.xlsx")
// シートを開く
s, _ := w.OpenSheet("Sheet1")
// 一行目を取得
r := s.GetRow(1)
// 1列目のセルを取得
c := r.GetCell(1)
// セルに10を出力
c.SetNumber("10")
// 2列目のセルにABCDEという文字列を出力
c = r.SetString("ABCDE", 2)
// シートを閉じる
s.Close()
// 保存
w.Save("path/to/new.xlsx")
```

新規Excelファイルを作成
```go
// 新規Excelファイルを作成
w, _ := excl.Create()
s, _ := w.OpenSheet("Sheet1")
s.Close()
w.Save("path/to/new.xlsx")
```

セルの書式の設定方法
```go
w, _ := excl.Open("path/to/read.xlsx")
s, _ := w.OpenSheet("Sheet1")
r := s.GetRow(1)
c := r.GetCell(1)
c.SetNumber("10000.00")
// 数値のフォーマットを設定する
c.SetNumFmt("#,##0.0")
// フォントの設定
c.SetFont(excl.Font{Size: 12, Color: "FF00FFFF", Bold: true, Italic: false,Underline: false})
// 背景色の設定
c.SetBackgroundColor("FFFF00FF")
// 罫線の設定
c.SetBorder(excl.Border{
	Left:   &excl.BorderSetting{Style: "thin", Color: "FFFFFF00"},
	Right:  &excl.BorderSetting{Style: "hair"},
	Top:    &excl.BorderSetting{Style: "dashDotDot"},
	Bottom: nil,
})
s.Close()
w.Save("path/to/new.xlsx")
```

グリッド線の表示非表示
```go
w, _ := excl.Open("path/to/read.xlsx")
s, _ := w.OpenSheet("Sheet1")
// シートのグリッド線を表示
s.ShowGridlines(true)
// シートのグリッド線を非表示
s.ShowGridlines(false)
s.Close()
w.Save("path/to/new.xlsx")
```

カラム幅の変更
```go
w, _ := excl.Open("path/to/read.xlsx")
s, _ := w.OpenSheet("Sheet1")
// 5番目のカラム幅を1.1に変更
s.SetColWidth(1.1, 5)
s.Close()
w.Save("path/to/new.xlsx")
```

計算式結果の更新が必要な場合はSetForceFormulaRecalculationを使用する
この関数を利用することでExcelを開いた際に結果が自動的に更新される
```go
w, _ := excl.Open("path/to/read.xlsx")
// 何か処理...
w.SetForceFormulaRecalculation(true)
w.Save("path/to/new.xlsx")
```

## Install

```bash
$ go get github.com/loadoff/excl
```

## Licence

[MIT](https://github.com/loadoff/excl/LICENCE)

## Author

[YuIwasaki](https://github.com/loadoff)
