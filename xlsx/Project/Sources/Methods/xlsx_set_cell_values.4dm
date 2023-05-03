//%attributes = {"invisible":true}
C_OBJECT:C1216($1; $folder)
C_LONGINT:C283($2; $sheetIndex)
C_OBJECT:C1216($3; $values)
C_TEXT:C284($4; $mode)

$folder:=$1
$sheetIndex:=$2
$values:=$3
$mode:=$4

C_TEXT:C284($xml; $stringValue)

If ($folder.exists)
	
	C_COLLECTION:C1488($sharedStringsCollection)
	C_BOOLEAN:C305($isSharedStringsOpen)
	
	If ($mode="t")
		
		$sharedStrings:=$folder\
			.folder("xl")\
			.file("sharedStrings.xml")
		
		If ($sharedStrings.exists)
			
			$sharedStringsCollection:=New collection:C1472
			
			$xml:=$sharedStrings.getText("utf-8"; Document with LF:K24:22)
			$domSharedStrings:=DOM Parse XML variable:C720($xml)
			If (OK=1)
				$isSharedStringsOpen:=True:C214
				$sst:=DOM Find XML element:C864($domSharedStrings; "/sst")
				C_LONGINT:C283($count; $uniqueCount)
				DOM GET XML ATTRIBUTE BY NAME:C728($sst; "count"; $count)
				DOM GET XML ATTRIBUTE BY NAME:C728($sst; "uniqueCount"; $uniqueCount)
				ARRAY TEXT:C222($sis; 0)
				$si:=DOM Find XML element:C864($sst; "si"; $sis)
				If (OK=1)
					For ($i; 1; Size of array:C274($sis))
						
						$si:=$sis{$i}
						ARRAY TEXT:C222($ts; 0)
						$t:=DOM Find XML element:C864($si; "t"; $ts)  //child
						
						If (OK=1)
							
							DOM GET XML ELEMENT VALUE:C731($t; $stringValue)
							$hash:=Generate digest:C1147($stringValue; MD5 digest:K66:1)
							$index:=$sharedStringsCollection.length  //0-based
							$sharedStringsCollection.push(New object:C1471(\
								"hash"; $hash; \
								"value"; $stringValue; \
								"index"; $index))
							
						Else 
							$sharedStringsCollection.push(Null:C1517)
						End if 
					End for 
				Else 
					//
				End if 
			End if 
		End if 
	End if 
	
	$sheet:=$folder\
		.folder("xl")\
		.folder("worksheets")\
		.file(New collection:C1472("sheet"; $sheetIndex; ".xml")\
		.join())
	
	If ($sheet.exists)
		$xml:=$sheet.getText("utf-8"; Document with LF:K24:22)
		$dom:=DOM Parse XML variable:C720($xml)
		If (OK=1)
			//get cells
			$sheetData:=DOM Find XML element:C864($dom; "/worksheet/sheetData")
			If (OK=1)
				ARRAY TEXT:C222($rows; 0)
				$row:=DOM Find XML element:C864($sheetData; "row"; $rows)
				If (OK=1)
					For ($i; 1; Size of array:C274($rows))
						$row:=$rows{$i}
						
						For ($iii; 1; DOM Count XML attributes:C727($row))
							DOM GET XML ATTRIBUTE BY INDEX:C729($row; $iii; $name; $stringValue)
						End for 
						
						ARRAY TEXT:C222($cs; 0)
						$c:=DOM Find XML element:C864($row; "c"; $cs)
						If (OK=1)
							For ($ii; 1; Size of array:C274($cs))
								$c:=$cs{$ii}
								
								For ($iii; 1; DOM Count XML attributes:C727($c))
									DOM GET XML ATTRIBUTE BY INDEX:C729($c; $iii; $name; $stringValue)
									Case of 
										: ($name="r")
											$cellRef:=$stringValue
											
											If (OB Is defined:C1231($values; $cellRef))
												
												Case of 
													: ($mode="f")
														
														$f:=DOM Find XML element:C864($c; "f")
														
														If (OK=0)
															$f:=DOM Create XML element:C865($c; "f")
														End if 
														
														DOM SET XML ELEMENT VALUE:C868($f; $values[$cellRef].f)
														
														$v:=DOM Find XML element:C864($c; "v")
														
														If (OK=0)
															$v:=DOM Create XML element:C865($c; "v")
														End if 
														
														DOM SET XML ELEMENT VALUE:C868($v; Num:C11($values[$cellRef].v))
														
													: ($mode="t") & ($isSharedStringsOpen)
														
														$v:=DOM Find XML element:C864($c; "v")
														
														If (OK=0)
															$v:=DOM Create XML element:C865($c; "v")
														End if 
														
														DOM SET XML ATTRIBUTE:C866($c; "t"; "s")  //shared string
														
														$stringValue:=$values[$cellRef]
														$hash:=Generate digest:C1147($stringValue; MD5 digest:K66:1)
														$find:=$sharedStringsCollection.query("hash === :1"; $hash)
														
														If ($find.length=0)
															
															$count:=$count+1
															$uniqueCount:=$uniqueCount+1
															$index:=$sharedStringsCollection.length
															
															$sharedStringsCollection.push(New object:C1471(\
																"hash"; $hash; \
																"value"; $stringValue; \
																"index"; $index))
															
															$si:=DOM Create XML element:C865($sst; "si")
															$t:=DOM Create XML element:C865($si; "t")
															DOM SET XML ELEMENT VALUE:C868($t; $stringValue)
															
															$phoneticPr:=DOM Create XML element:C865($si; "phoneticPr")
															DOM SET XML ATTRIBUTE:C866($phoneticPr; "fontId"; 1)
															
														Else 
															$index:=$sharedStringsCollection[($find[0].index)].index
														End if 
														
														DOM SET XML ELEMENT VALUE:C868($v; $index)
														
													: ($mode="n")
														
														$v:=DOM Find XML element:C864($c; "v")
														
														If (OK=0)
															$v:=DOM Create XML element:C865($c; "v")
														End if 
														
														DOM SET XML ELEMENT VALUE:C868($v; Num:C11($values[$cellRef]))
														
													Else 
														//
												End case 
												
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
			DOM EXPORT TO FILE:C862($dom; $path)
			DOM CLOSE XML:C722($dom)
		End if 
	End if 
	
	If ($isSharedStringsOpen)
		
		DOM SET XML ATTRIBUTE:C866($sst; "count"; $count; "uniqueCount"; $uniqueCount)
		
		$path:=$sharedStrings.platformPath
		DOM EXPORT TO FILE:C862($domSharedStrings; $path)
		DOM CLOSE XML:C722($domSharedStrings)
		
	End if 
	
End if 