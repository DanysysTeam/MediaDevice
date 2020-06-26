#include "..\MediaDevice.au3"

_Example()

Func _Example()
	Local $aDevices = _MD_DevicesList() ;Get all Devices
	_ArrayDisplay($aDevices,"Device List")

	Local $sDevice = _MD_DeviceGet("", 0, "", "", "Generic")
	ConsoleWrite("$sDevice: " & $sDevice & @CRLF)

	$sDevice = _MD_DeviceGet("", 1, "", "", "Generic") ;Get Second Generic Device
	ConsoleWrite("$sDevice: " & $sDevice & @CRLF)

	$sDevice = _MD_DeviceGet("", 0, "SanDisk", "", "Generic") ;Get First Device Matching Manufacturer and Device Type
	ConsoleWrite("$sDevice: " & $sDevice & @CRLF)

	$sDevice = _MD_DeviceGetFirst() ;Get first Device Found
    ConsoleWrite("$sDevice: " & $sDevice & @CRLF)
EndFunc   ;==>_Example1

