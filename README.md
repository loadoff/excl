excl
====

これはexcelコントロール用のライブラリ

[![godoc](https://godoc.org/github.com/loadoff/excl?status.svg)](https://godoc.org/github.com/loadoff/excl)
[![go report](https://goreportcard.com/badge/github.com/loadoff/excl)](https://goreportcard.com/report/github.com/loadoff/excl)

## Description

基本的にもとのexcelファイルを破壊せずにデータの入力を行うためのライブラリです。
また大量のデータを扱う上でも優位になるように開発を行います。

## Usage

```go
// 読み込みExcelファイル、展開先、新規書き込み先を指定
w, _ := excl.NewWorkbook("path/to/read.xlsx", "path/to/expand", "path/to/write.xlsx")
// Execlブックを開く
w.Open()
// シートを開く
s, _ := w.OpenSheet("sheet1")
// 一行目を取得
r := s.GetRow(1)
// 1列目のセルを取得
c := r.GetCell(1)
// セルに10を出力
c.SetNumber("10")

// 2列目のセルにABCDEという文字列を出力
c = r.SetString("ABCDE", 2)

s.Close()
w.Close()
```

## Install

```bash
$ go get github.com/loadoff/excl
```

## Licence

[MIT](https://github.com/loadoff/excl/LICENCE)

## Author

[YuIwasaki](https://github.com/loadoff)
