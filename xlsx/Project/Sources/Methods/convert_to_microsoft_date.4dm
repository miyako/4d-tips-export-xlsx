//%attributes = {"invisible":true}
C_DATE:C307($1; $date)
C_LONGINT:C283($0; $value)

$date:=$1

Case of 
	: ($date<!1900-01-01!)
		
		//out of range
		
	: (!1900-01-01!<$date) & ($date<!1900-03-01!)
		
		$value:=$date-!1899-12-31!
		
	: ($date>!1900-02-28!)
		
		$value:=$date-!1899-12-30!
		
	Else 
/*
		
Microsoft date 60 (29th February 1900) does not exist!
		
https://en.wikipedia.org/wiki/Year_1900_problem
		
*/
End case 

$0:=$value