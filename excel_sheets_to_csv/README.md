# excel_sheets_to_csv

# 概要
指定されたディレクトリ内にある Excel ファイルすべての全シートをcsv ファイルとして保存する。

## 使い方
	ruby excel_sheets_to_csv.rb .\sources .\out

## 解説
### 引数
.\sources Excel ファイルたちがあるディレクトリへのパス  
.\out csv ファイルを書き出すディレクトリへのパス  
### 出力
csv ファイル名は 	{Excel ファイル名}_{シート名}.csv となります。
sample1.xlsx の Sheet1 をもとにした場合、sample1_Sheet1.csv となります。
### 対象ファイル
.xlsx .xlsm .xlsb .xls





