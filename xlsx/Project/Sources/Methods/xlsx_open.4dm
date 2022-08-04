//%attributes = {"invisible":true}
C_OBJECT:C1216($1; $file; $0; $folder)

$file:=$1

If ($file.exists)
	
	$path:=Temporary folder:C486+Generate UUID:C1066
	$folder:=Folder:C1567($path; fk platform path:K87:2)
	$folder.create()
	
	If ($folder.exists)
		
		$folder:=ZIP Read archive:C1637($file).root.copyTo($folder)
		
	End if 
End if 

$0:=$folder