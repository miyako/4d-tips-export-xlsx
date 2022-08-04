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

DOMコマンドのシンタックス（Xpath）はv18とv19で違う点にも注意が必要です。

```4d
//$si:=DOM Find XML element($sst;"si";$sis)//v19以降で標準Xpathを有効にしている場合
$si:=DOM Find XML element($sst;"sst/si";$sis)
```

フォーミュラは`E3+F3+G3`のようなセル参照でも，`SUM(H2:H3)`でも構いませんが，セルの書式と内容が合っていなければなりません。
  
また，`.xlsx`ファイルにはフォーミュラと一緒に計算結果も保存しておくことになっています（計算はセルを更新したタイミングで実行することが前提になっているため）。`fullCalcOnLoad`というプロパティを`1`に設定することで，ファイルを開いた直後にすべてのフォーミュラを再計算させることもできますが，Excel以外のアプリはこれに対応していないかもしれません。

## 例題

冒頭に挙げたスクリーンショットの結果を出力します。

```4d
/*
	
	あらかじめMicrosoft® Excelで作成した
	スプレッドシートのテンプレートを開きます。
	
	1行目
	通常はヘッダー行として使用します。
	列の幅やスタイルも設定することができます。
	
	2行目
	【重要】何も設定していないセルは保存時に省略されるので
	何らかの値をプレースホルダーとして入力するか
	デフォルト以外のフォーマットをかならず適用します。
	
	数値
	プレースホルダーとして0を入力するか
	セルの書式設定で桁区切りなどのフォーマットを設定します。
	
	文字列
	数値以外の文字をプレースホルダーとして入力するか
	セルの書式設定で右寄せなどのフォーマットを設定します。
	
	日付
	セルの書式設定で「yyyy/m/d;;; 」などのユーザー定義フォーマットを適用し
	値が0の場合には空の文字列が表示されるようにします。
	
	時刻
	エクセルの時刻（1900年の元旦の時刻）
	
*/

$file:=Folder(fk resources folder).file("TEMPLATE.xlsx")

$temp:=xlsx_open ($file)

$numericValues:=New object
$stringValues:=New object
$formulaValues:=New object

$stringValues["A2"]:="あいうえお"
$numericValues["B2"]:=convert_to_microsoft_date (!2022-04-14!)
$numericValues["C2"]:=convert_to_microsoft_time (?01:23:45?)
$numericValues["E2"]:=10000
$numericValues["F2"]:=9999
$numericValues["G2"]:=1
$formulaValues["H2"]:=New object("f";"E2+F2+G2";"v";10000+9999+1)

$stringValues["A3"]:="かきくけこ"
$numericValues["B3"]:=convert_to_microsoft_date (!2022-04-15!)
$numericValues["C3"]:=convert_to_microsoft_time (?01:32:54?)
$numericValues["E3"]:=1000
$numericValues["F3"]:=999
$numericValues["G3"]:=1
$formulaValues["H3"]:=New object("f";"E3+F3+G3";"v";1000+999+1)


$formulaValues["F4"]:=New object("f";"SUM(F2:F3)";"v";10000+1000)
$formulaValues["G4"]:=New object("f";"SUM(G2:G3)";"v";9999+999)
$formulaValues["H4"]:=New object("f";"SUM(H2:H3)";"v";10000+9999+1+1000+999+1)

/*
	テンプレートは2行だけなので
	値が表示されるようにシートを拡げます。
*/

xlsx_resize ($temp;1;4)

xlsx_set_cell_values_n ($temp;1;$numericValues)
xlsx_set_cell_values_t ($temp;1;$stringValues)
xlsx_set_cell_values_f ($temp;1;$formulaValues)

$XLSX:=xlsx_close ($temp)

$timestamp:=Split string(Replace string(Timestamp;":";"-";*);".")[0]
$folder:=Folder(fk user preferences folder).parent.folder(Folder(fk database folder).name).folder("export")
$folder.create()

$file:=$folder.file($timestamp+".xlsx")

$file.setContent($XLSX)

OPEN URL($file.platformPath)
```

## 関連情報

`.xlsx`を出力する方法はいろいろあります。

* [XL Plugin](https://www.pluggers.nl/product/xl-plugin/)
* [4D View Pro](https://doc.4d.com/4Dv18/4D/18/VP-EXPORT-DOCUMENT.301-4522260.ja.html)
* [2020年サミット - Justin Will](https://events.4d.com/summit2020/session/generate-pdfs-excel-files-and-ways-to-integrate-pre-post-scripts-through-quick-report/)
