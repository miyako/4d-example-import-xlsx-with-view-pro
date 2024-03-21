//%attributes = {"invisible":true}
$xlsxFile:=Folder:C1567(fk resources folder:K87:11).file("test.xlsx")

$myVpConverterParams:=New object:C1471
$myVpConverterParams.src:=$xlsxFile
$myVpConverterParams.password:=""
$myVpConverterParams.formula:=Formula:C1597(myVpImporter)

$params:=New object:C1471
$params.area:="tempVpArea"
$params.onEvent:=Formula:C1597(onEvent($myVpConverterParams))
$params.autoQuit:=False:C215
$params.timeout:=10
$params.result:=Null:C1517

$myVpProcessorParams:=New object:C1471
$myVpProcessorParams.that:=$params
$myVpConverterParams.onSuccess:=Formula:C1597(myVpProcessor($myVpProcessorParams))
$myVpConverterParams.onError:=Formula:C1597(myVpProcessorError($myVpProcessorParams))
$myVpProcessorParams.formula:=Formula:C1597(myVpExporter)
$myVpProcessorParams.includeFormatInfo:=False:C215
$myVpProcessorParams.sheetIndex:=vk workbook:K89:4

C_VARIANT:C1683($result)
$result:=VP Run offscreen area($params)

If ($result#Null:C1517)  //タイムアウトの場合はNull
	
	ALERT:C41("Sheet1.A1の値は"+$result.spreadJS.sheets.Sheet1.data.dataTable["0"]["0"].value)
	
End if 
