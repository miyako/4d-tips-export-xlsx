//%attributes = {"invisible":true}
C_OBJECT:C1216($1;$folder)
C_LONGINT:C283($2;$sheetIndex)
C_LONGINT:C283($3;$length)

$folder:=$1
$sheetIndex:=$2
$length:=$3

$sheet:=$folder\
.folder("xl")\
.folder("worksheets")\
.file(New collection:C1472("sheet";$sheetIndex;".xml")\
.join())

If ($sheet.exists)
	$xml:=$sheet.getText("utf-8";Document with LF:K24:22)
	$dom:=DOM Parse XML variable:C720($xml)
	If (OK=1)
		$dimension:=DOM Find XML element:C864($dom;"/worksheet/dimension")
		If (OK=1)
			
			$ref:=""
			DOM GET XML ATTRIBUTE BY NAME:C728($dimension;"ref";$ref)
			
			ARRAY LONGINT:C221($pos;0)
			ARRAY LONGINT:C221($len;0)
			
			If (Match regex:C1019("([A-Z]+)(\\d+):([A-Z]+)(\\d+)";$ref;1;$pos;$len))
				
				$firstCell:=Substring:C12($ref;$pos{1};$len{1})
				$firstRow:=Num:C11(Substring:C12($ref;$pos{2};$len{2}))
				$lastCell:=Substring:C12($ref;$pos{3};$len{3})
				$lastRow:=Num:C11(Substring:C12($ref;$pos{4};$len{4}))
				
				$rowsToAppend:=$length-$lastRow
				
				$ref:=New collection:C1472($firstCell;$firstRow;":";$lastCell;$length).join("")
				
				DOM SET XML ATTRIBUTE:C866($dimension;"ref";$ref)
				
				$sheetData:=DOM Find XML element:C864($dom;"/worksheet/sheetData")
				If (OK=1)
					ARRAY TEXT:C222($rows;0)
					  //$row:=DOM Find XML element($sheetData;"row";$rows)//v19以降で標準Xpathを有効にしている場合
					$row:=DOM Find XML element:C864($sheetData;"sheetData/row";$rows)
					If (OK=1)
						For ($i;1;Size of array:C274($rows))
							$row:=$rows{$i}
							For ($ii;1;DOM Count XML attributes:C727($row))
								DOM GET XML ATTRIBUTE BY INDEX:C729($row;$ii;$name;$stringValue)
								Case of 
									: ($name="r")
										
										If (2=Num:C11($stringValue))
											$ii:=MAXLONG:K35:2-1
											$i:=$ii
											
											$rowStyle:=New object:C1471
											
											For ($iii;1;DOM Count XML attributes:C727($row))
												DOM GET XML ATTRIBUTE BY INDEX:C729($row;$iii;$name;$stringValue)
												If ($name#"r")
													$rowStyle[$name]:=$stringValue
												End if 
											End for 
											
											ARRAY TEXT:C222($cols;0)
											  //$col:=DOM Find XML element($row;"c";$cols)//v19以降で標準Xpathを有効にしている場合
											$col:=DOM Find XML element:C864($row;"row/c";$cols)
											$styles:=New object:C1471
											
											For ($iii;1;Size of array:C274($cols))
												$col:=$cols{$iii}
												$style:=New object:C1471
												For ($ii;1;DOM Count XML attributes:C727($col))
													DOM GET XML ATTRIBUTE BY INDEX:C729($col;$ii;$name;$stringValue)
													If ($name="r")
														If (Match regex:C1019("([A-Z]+)(\\d+)";$stringValue;1;$pos;$len))
															$cell:=Substring:C12($stringValue;$pos{1};$len{1})
															$styles[$cell]:=$style
														End if 
													Else 
														$style[$name]:=$stringValue
													End if 
												End for 
											End for 
											
											For ($iii;3;2+$rowsToAppend)  //row
												
												$r:=String:C10($iii)
												
												$row:=DOM Create XML element:C865($sheetData;"row")
												
												For each ($name;$rowStyle)  //except r
													DOM SET XML ATTRIBUTE:C866($row;$name;$rowStyle[$name])
												End for each 
												DOM SET XML ATTRIBUTE:C866($row;"r";$r)
												
												For each ($name;$styles)
													
													$node:=DOM Create XML element:C865($row;"c")
													
													DOM SET XML ATTRIBUTE:C866($node;"r";$name+$r)
													
													$style:=$styles[$name]
													For each ($attribute;$style)
														DOM SET XML ATTRIBUTE:C866($node;$attribute;$style[$attribute])
													End for each 
												End for each 
												
											End for 
										End if 
								End case 
							End for 
						End for 
					End if 
				End if 
			End if 
		End if 
		
		$path:=$sheet.platformPath
		DOM EXPORT TO FILE:C862($dom;$path)
		DOM CLOSE XML:C722($dom)
		
	End if 
End if 