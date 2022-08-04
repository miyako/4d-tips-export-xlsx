//%attributes = {"invisible":true}
C_OBJECT:C1216($1;$folder)
C_LONGINT:C283($2;$sheetIndex)
C_OBJECT:C1216($3;$styles)

$folder:=$1
$sheetIndex:=$2
$styles:=$3

C_TEXT:C284($xml;$stringValue)

If ($folder.exists)
	
	$sheet:=$folder\
		.folder("xl")\
		.folder("worksheets")\
		.file(New collection:C1472("sheet";$sheetIndex;".xml")\
		.join())
	
	If ($sheet.exists)
		$xml:=$sheet.getText("utf-8";Document with LF:K24:22)
		$dom:=DOM Parse XML variable:C720($xml)
		If (OK=1)
			  //get cells
			$sheetData:=DOM Find XML element:C864($dom;"/worksheet/sheetData")
			If (OK=1)
				ARRAY TEXT:C222($rows;0)
				  //$row:=DOM Find XML element($sheetData;"row";$rows)  //v19以降で標準Xpathを有効にしている場合
				$row:=DOM Find XML element:C864($sheetData;"sheetData/row";$rows)
				If (OK=1)
					For ($i;1;Size of array:C274($rows))
						$row:=$rows{$i}
						
						For ($iii;1;DOM Count XML attributes:C727($row))
							DOM GET XML ATTRIBUTE BY INDEX:C729($row;$iii;$name;$stringValue)
						End for 
						
						ARRAY TEXT:C222($cs;0)
						  //$c:=DOM Find XML element($row;"c";$cs)//v19以降で標準Xpathを有効にしている場合
						$c:=DOM Find XML element:C864($row;"row/c";$cs)
						If (OK=1)
							For ($ii;1;Size of array:C274($cs))
								$c:=$cs{$ii}
								
								For ($iii;1;DOM Count XML attributes:C727($c))
									DOM GET XML ATTRIBUTE BY INDEX:C729($c;$iii;$name;$stringValue)
									Case of 
										: ($name="r")
											$cellRef:=$stringValue
											
											If (OB Is defined:C1231($styles;$cellRef))
												
												$s:=$styles[$cellRef]
												
												DOM SET XML ATTRIBUTE:C866($c;"s";$s)
												
											End if 
										Else 
											  //
									End case 
								End for 
							End for 
						End if 
					End for 
				End if 
			End if 
			$path:=$sheet.platformPath
			DOM EXPORT TO FILE:C862($dom;$path)
			DOM CLOSE XML:C722($dom)
		End if 
	End if 
	
End if 