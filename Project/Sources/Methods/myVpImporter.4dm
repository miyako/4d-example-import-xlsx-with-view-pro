//%attributes = {"invisible":true,"preemptive":"capable"}
C_TEXT:C284($1; $area; $2; $path)
C_OBJECT:C1216($3; $params; $4; $status)

$area:=$1
$path:=$2

$params:=$3
$status:=$4

If ($status.success)
	$params.onSuccess.call($params)
	ACCEPT:C269
Else 
	$params.onError.call($params)
	CANCEL:C270
End if 