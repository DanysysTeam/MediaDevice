#include "..\MediaDevice.au3"

_Example()

Func _Example()
	Local $aDevices = _MD_DevicesList()
	_ArrayDisplay($aDevices,"Device List")
	Local $sDevice = _MD_DeviceGet() ;Get First Found Portable Device
	Local $oDevice = _MD_DeviceOpen($sDevice)
	Local $aDrives = _MD_DeviceGetDrives($oDevice)
	_ArrayDisplay($aDrives,"Storage Drives")
	_MD_DeviceClose($oDevice)
EndFunc   ;==>_Example1

