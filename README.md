![version](https://img.shields.io/badge/version-18%2B-EB8E5F)

# 4d-tips-export-xlsx
Zipコマンドを活用してスプレッドシートを作成する例題

<img width="718" alt="xlsx" src="https://user-images.githubusercontent.com/1725068/182773631-866d88d1-7249-4ce6-8f29-4655c4bd9e19.png">

## 概要

あらかじめMicrosoft® Excelなどでスプレッドシートのテンプレートを作成しておき，セルの書式設定などを決めておきます。
  
```4d
$file:=Folder(fk resources folder).file("TEMPLATE.xlsx")
$temp:=xlsx_open ($file)
```
  
Zipコマンドで`.xlsx`ファイルを開き，セルの値を追加してゆきます。毎回，XMLをパースするのは面倒なので，オブジェクト型（プロパティ名はセルの位置）にセットしておき，数値・文字列・フォーミュラをそれぞれまとめて書き込みます。

**注意点**: エクセルは日付や時間の扱いが4Dとは違うので変換メソッドが必要です。

日付は基本的に1900年1月1日から数えますが，Microsoftの[有名なバグ](https://en.wikipedia.org/wiki/Year_1900_problem)で第`60`日（1900年2月29日）を計算に含める必要があります。

時間は`24`時間を`1`と数えるので，4Dの時間型（真夜中からの秒数）を`86400`で割った値になります。

文字列は，スプレッドシート内で簡易的にデータベース化されており，同じ文字列は番号で参照する必要があります。はじめに *sharedStrings.xml* ファイルを解析し，文字列と番号の関係を把握した上で，新しい文字列を追加してゆきます。
