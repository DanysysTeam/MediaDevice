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

	;Get Directories in Root Drive Selected
	;This is for avoiding to list all files in the Drive, So we get the a random Path
	Local $aFiles2D = _MD_DeviceFileListToArray($oDevice, $sDrivePath, True)
	If Not IsArray($aFiles2D) Then Exit MsgBox(0, "Error", "Unable To Get Directories")

    ;Get First File
	Local $sFirstFile = ""
	For $i = 0 To UBound($aFiles2D) - 1
		If Not $aFiles2D[$i][17] Then $sFirstFile = $aFiles2D[$i][18]
	Next
	;Copy File
	ConsoleWrite($sFirstFile & @CRLF)
	If $sFirstFile Then
	Local $sFilePath = _MD_DeviceFileCopyFrom($oDevice, $sFirstFile, @UserProfileDir & "\Desktop", True)
	ConsoleWrite("Copied FilePath: " & $sFilePath & @TAB & "@extended: " & @extended & @CRLF)
	EndIf


	_MD_DeviceClose($oDevice)
EndFunc   ;==>_Example


