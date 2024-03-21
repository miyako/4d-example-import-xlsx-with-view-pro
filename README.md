![version](https://img.shields.io/badge/version-18%2B-EB8E5F) ![](https://img.shields.io/badge/version-19%2B-5682DF) ![](https://img.shields.io/badge/version-20%2B-E23089)

オフスクリーンエリアを使用してXLSXを取り込む例題です。

クラスは敢えて使用していません。

セットアップ（`VP Run offscreen area`）とインポート（`VP IMPORT DOCUMENT`）は非同期処理，エクスポート（`VP Export to object`）は同期処理なのがポイントです。

<img src="https://github.com/miyako/4d-example-import-xlsx-with-view-pro/assets/1725068/a24e3cea-1765-4ff9-ab86-e62e63bf8490" style="width:400px;height:auto" />

```4d
$xlsxFile:=Folder(fk resources folder).file("test.xlsx")

$myVpConverterParams:=New object
$myVpConverterParams.src:=$xlsxFile
$myVpConverterParams.password:=""
$myVpConverterParams.formula:=Formula(myVpImporter)

$params:=New object
$params.area:="tempVpArea"
$params.onEvent:=Formula(onEvent($myVpConverterParams))
$params.autoQuit:=False
$params.timeout:=10
$params.result:=Null

$myVpProcessorParams:=New object
$myVpProcessorParams.that:=$params
$myVpConverterParams.onSuccess:=Formula(myVpProcessor($myVpProcessorParams))
$myVpConverterParams.onError:=Formula(myVpProcessorError($myVpProcessorParams))
$myVpProcessorParams.formula:=Formula(myVpExporter)
$myVpProcessorParams.includeFormatInfo:=False
$myVpProcessorParams.sheetIndex:=vk workbook

C_VARIANT($result)
$result:=VP Run offscreen area($params)

If ($result#Null)  //タイムアウトの場合はNull
	
	ALERT("Sheet1.A1の値は"+$result.spreadJS.sheets.Sheet1.data.dataTable["0"]["0"].value)
	
End if 
```
