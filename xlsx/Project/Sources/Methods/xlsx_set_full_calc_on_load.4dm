//%attributes = {"invisible":true}
C_OBJECT:C1216($1)

$folder:=$1

$workbook:=$folder\
.folder("xl")\
.file("workbook.xml")

If ($workbook.exists)
	
	C_OBJECT:C1216($sheetRefs)
	$sheetRefs:=New object:C1471
	
	$xml:=$workbook.getText("utf-8"; Document with LF:K24:22)
	$dom:=DOM Parse XML variable:C720($xml)
	
	If (OK=1)
		
		$calcId:=False:C215
		
		$calcPr:=DOM Find XML element:C864($dom; "/workbook/calcPr")
		If (OK=1)
			DOM SET XML ATTRIBUTE:C866($calcPr; "fullCalcOnLoad"; 1)
			For ($ii; 1; DOM Count XML attributes:C727($calcPr))
				DOM GET XML ATTRIBUTE BY INDEX:C729($calcPr; $ii; $name; $stringValue)
				If ($stringValue="calcId")
					$calcId:=True:C214
				End if 
				If ($calcId)
					DOM REMOVE XML ATTRIBUTE:C1084($dom; "calcId")
				End if 
			End for 
		End if 
		
		$path:=$workbook.platformPath
		DOM EXPORT TO FILE:C862($dom; $path)
		DOM CLOSE XML:C722($dom)
	End if 
End if 