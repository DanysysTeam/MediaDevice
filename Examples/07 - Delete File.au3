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

    ;Create, Rename and Delete
	Local $sFolderName = InputBox("Select a Name For a Folder to be Created", "Folder Name", "Folder_" & Random(1, 1000, 1))
	If $sFolderName Then
		$sFolderName = $sDrivePath & $sFolderName ; Create Full Path
		Local $sFolderRename = $sDrivePath & "AutoIt Rocks"
		Local $sFolderId = _MD_DeviceCreateDirectory($oDevice, $sFolderName)
		ConsoleWrite($sFolderId & @CRLF)
		MsgBox(0, "Information", "Folder will be Rename After this Message")
		Local $sFolderRenamedId = _MD_DeviceItemRename($oDevice, $sFolderName, "AutoIt Rocks")
		ConsoleWrite($sFolderRenamedId & @CRLF)
		MsgBox(0, "Information", "Folder will be Delete After this Message")
		_MD_DeviceItemDelete($oDevice, $sFolderRename)
	EndIf

	_MD_DeviceClose($oDevice)
EndFunc   ;==>_Example


