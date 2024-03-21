//%attributes = {"invisible":true}
C_TEXT:C284($1; $area; $2; $path)
C_OBJECT:C1216($3; $params; $4; $status)

$area:=$1
$path:=$2

$params:=$3
$status:=$4

$params.that.result:=$status

If ($status.success)
	ACCEPT:C269
Else 
	CANCEL:C270
End if 