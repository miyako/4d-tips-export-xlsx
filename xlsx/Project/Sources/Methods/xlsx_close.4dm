//%attributes = {"invisible":true}
C_OBJECT:C1216($1; $folder)
C_BLOB:C604($0)

$folder:=$1

If ($folder.exists)
	
	$file:=Folder:C1567(Temporary folder:C486; fk platform path:K87:2).file(Generate UUID:C1066+".xlsx")
	
	$source:=New object:C1471
	$source.files:=$folder.files().combine($folder.folders())
	$status:=ZIP Create archive:C1640($source; $file)
	
	If ($status.success)
		$0:=$file.getContent()
		$file.delete()
	End if 
	
End if 