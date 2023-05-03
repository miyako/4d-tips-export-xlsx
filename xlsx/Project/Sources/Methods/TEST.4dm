//%attributes = {"invisible":true}
//note: standard XPath is enabled

/*

あらかじめMicrosoft® Excelで作成した
スプレッドシートのテンプレートを開きます。

1行目
通常はヘッダー行として使用します。
列の幅やスタイルも設定することができます。

2行目
【重要】何も設定していないセルは保存時に省略されるので
何らかの値をプレースホルダーとして入力するか
デフォルト以外のフォーマットをかならず設定します。

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
エクセルの時刻（24時間を1として換算）

*/

$file:=Folder:C1567(fk resources folder:K87:11).file("TEMPLATE.xlsx")

$temp:=xlsx_open($file)

$numericValues:=New object:C1471
$stringValues:=New object:C1471
$formulaValues:=New object:C1471

$stringValues["A2"]:="あいうえお"
$numericValues["B2"]:=convert_to_microsoft_date(!2022-04-14!)
$numericValues["C2"]:=convert_to_microsoft_time(?01:23:45?)
$numericValues["E2"]:=10000
$numericValues["F2"]:=9999
$numericValues["G2"]:=1
$formulaValues["H2"]:=New object:C1471("f"; "E2+F2+G2"; "v"; 10000+9999+1)

$stringValues["A3"]:="かきくけこ"
$numericValues["B3"]:=convert_to_microsoft_date(!2022-04-15!)
$numericValues["C3"]:=convert_to_microsoft_time(?01:32:54?)
$numericValues["E3"]:=1000
$numericValues["F3"]:=999
$numericValues["G3"]:=1
$formulaValues["H3"]:=New object:C1471("f"; "E3+F3+G3"; "v"; 1000+999+1)


$formulaValues["F4"]:=New object:C1471("f"; "SUM(F2:F3)"; "v"; 10000+1000)
$formulaValues["G4"]:=New object:C1471("f"; "SUM(G2:G3)"; "v"; 9999+999)
$formulaValues["H4"]:=New object:C1471("f"; "SUM(H2:H3)"; "v"; 10000+9999+1+1000+999+1)

/*
テンプレートは2行だけなので
値が表示されるようにシートを拡げます。
*/

xlsx_resize($temp; 1; 4)

xlsx_set_cell_values_n($temp; 1; $numericValues)
xlsx_set_cell_values_t($temp; 1; $stringValues)
xlsx_set_cell_values_f($temp; 1; $formulaValues)

$XLSX:=xlsx_close($temp)

$timestamp:=Split string:C1554(Replace string:C233(Timestamp:C1445; ":"; "-"; *); ".")[0]
$folder:=Folder:C1567(fk user preferences folder:K87:10).parent.folder(Folder:C1567(fk database folder:K87:14).name).folder("export")
$folder.create()

$file:=$folder.file($timestamp+".xlsx")

$file.setContent($XLSX)

OPEN URL:C673($file.platformPath)