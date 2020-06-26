#include "..\MediaDevice.au3"

_Example()

Func _Example()
	Local $sDevice = _MD_DeviceGet() ;Get First Found Portable Device Same like _MD_DeviceGetFirst()
	Local $oDevice = _MD_DeviceOpen($sDevice)
	Local $aDrives= _MD_DeviceGetDrives($oDevice)
	Local $sDrivePath=$aDrives[0][0] ;Get First Drive

	;List Files
	Local $aFiles1D= _MD_DeviceFileListToArray($oDevice,$sDrivePath)
    _ArrayDisplay($aFiles1D,"Item ObjectIds")
    ;List With Deep Information
	Local $aFiles2D= _MD_DeviceFileListToArray($oDevice,$sDrivePath,True)
    _ArrayDisplay($aFiles2D,"Item ObjectIds Deep Information")

	;list Folders
	 $aFiles1D= _MD_DeviceDiretoryListToArray($oDevice,$sDrivePath)
    _ArrayDisplay($aFiles1D,"Item ObjectIds")
    ;List With Deep Information
	$aFiles2D= _MD_DeviceDiretoryListToArray($oDevice,$sDrivePath,True)
    _ArrayDisplay($aFiles2D,"Item ObjectIds Deep Information")

	_MD_DeviceClose($oDevice)
EndFunc   ;==>_Example1

