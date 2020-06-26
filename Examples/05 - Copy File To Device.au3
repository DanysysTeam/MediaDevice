#include "..\MediaDevice.au3"

_Example()

Func _Example()
	;Select a Device
	Local $aDevices = _MD_DevicesList() ;Get all Devices
	_ArrayDisplay($aDevices, "Device List", "", 0, Default, "DevicePnPID|FrienlyName|Manufacturer|Description|DeviceType")
	Local $iIndex = Number(InputBox("Select a Index Of a Device From The Previous shown List", "Set Device Index", 0))
	If $iIndex < 0 Or $iIndex > UBound($aDevices) Then Exit MsgBox(0, "Error", "Invalid Device Index")
	Local $sDevice = $aDevices[$iIndex][0]
	ConsoleWrite($sDevice & @CRLF)

	;Open Device
	Local $oDevice = _MD_DeviceOpen($sDevice)
	If Not IsObj($oDevice) Then Exit MsgBox(0, "Error", "Unable To Open Device")
	Local $aDrives = _MD_DeviceGetDrives($oDevice)
	If Not IsArray($aDrives) Then Exit MsgBox(0, "Error", "Unable To Get Device Drives")
	Local $sDrivePath = $aDrives[0][6] = "" ? $aDrives[0][0] : $aDrives[0][6] ;Get First Drive Path
	ConsoleWrite($sDrivePath & @CRLF)

	;Select File and Copy To Device
	Local $sLocalFilePath = FileOpenDialog("Select a File To Copy To " & $sDrivePath, @DesktopDir & "\", "File (*.*)", $FD_FILEMUSTEXIST)
	ConsoleWrite($sLocalFilePath & @CRLF)
	If $sLocalFilePath Then
		Local $sFileId = _MD_DeviceFileCopyTo($oDevice, $sLocalFilePath, $sDrivePath)
		ConsoleWrite("Copied FileId: " & $sFileId & @TAB & "@extended: " & @extended & @CRLF)
	Else
		MsgBox(0, "Error", "No file selected")
	EndIf


	_MD_DeviceClose($oDevice)
EndFunc   ;==>_Example


