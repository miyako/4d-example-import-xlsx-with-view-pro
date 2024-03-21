//%attributes = {"invisible":true}
C_OBJECT:C1216($1; $params; $event)

$params:=$1

$event:=FORM Event:C1606

Case of 
	: ($event.code=On Load:K2:1)
		
	: ($event.code=On End URL Loading:K2:47)
		
	: ($event.code=On VP Ready:K2:59)
		
		If ($params.src.exists)
			VP IMPORT DOCUMENT(This:C1470.area; $params.src.platformPath; $params)
		End if 
		
	: ($event.code=On VP Range Changed:K2:61)
		
	: ($event.code=On URL Loading Error:K2:48)
		
	: ($event.code=On Unload:K2:2)
		
End case 