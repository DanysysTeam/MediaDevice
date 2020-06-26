#cs Copyright
    Copyright 2020 Danysys. <hello@danysys.com>

    Licensed under the MIT license.
    See LICENSE file or go to https://opensource.org/licenses/MIT for details.
#ce Copyright

#cs Information
    Author(s)......: DanysysTeam (Danyfirex & Dany3j)
    Description....: MediaDevice UDF Allows you to communicate with attached media and storage devices.
    Remarks........: The current implementation is designed for use with storage devices and mobile phones.
    Version........: 1.0.0
    AutoIt Version.: 3.3.14.5
	Thanks to .....:
					https://github.com/Bassman2/MediaDevices - Some ideas from that library.
#ce Information

#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 6 ;-w 5 -w 6
;~ #Tidy_Parameters=/tcb=-1 /sf /ewnl /reel /gd ;/sfc
#Region Include
#include-once
#include "WindowsPortableDevices.au3"
#include <Array.au3>
#EndRegion Include

; #CURRENT# =====================================================================================================================
; _MD_DeviceClose
; _MD_DeviceCreateDirectory
; _MD_DeviceDirectoryListToArrayRec
; _MD_DeviceDiretoryListToArray
; _MD_DeviceFileCopyFrom
; _MD_DeviceFileCopyTo
; _MD_DeviceFileListToArray
; _MD_DeviceFileListToArrayRec
; _MD_DeviceFindFile
; _MD_DeviceGet
; _MD_DeviceGetDrives
; _MD_DeviceGetFirst
; _MD_DeviceItemDelete
; _MD_DeviceItemRename
; _MD_DeviceOpen
; _MD_DevicesList
; _MD_GetFunctionalObjects
; _MD_GetItemPropertiesDefault
; _MD_GetItemProperty
; _MD_GetItemSingleProperties
; _MD_ItemIsContainer
; _MD_ItemIsFolder
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __MD_ChangePathLastNameWithNewName
; __MD_CopyStream
; __MD_CreateClientInformation
; __MD_CreateGlobalPropertiesDefault
; __MD_ErrorHandler
; __MD_FindItemByPath
; __MD_GetDrives
; __MD_GetLastDirectoryName
; __MD_GetStorageInfo
; __MD_ListRecursive
; __MD_RemoveLastDirectory
; __MD_RemoveLastDirectory2
; __MD_SHCreateStreamOnFileEx
; __MD_SlidePath
; ===============================================================================================================================

#Region Globals
Global Const $STGM_READ = 0x00000000
Global Const $STGM_WRITE = 0x00000001
Global Const $STGM_READWRITE = 0x00000002
Global Const $STGM_CREATE = 0x00001000

Global Enum $eMD_Device_PnPID, $eMD_DEVICE_FrienlyName, _
		$eMD_Device_Manufacturer, $eMD_Device_Description, $eMD_Device_DeviceType
Global $__oMD_ErrorHandler = ObjEvent("AutoIt.Error", "__MD_ErrorHandler")
Global $__goPropropertiesDefault = 0
Global $__gaFilesRec2D[0][0]
Global $__gaFilesRec2DBasic[0][0]
#EndRegion Globals

#Region UDF Funtions

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __MD_CopyStream
; Description ...: Copy Streams
; Syntax ........: __MD_CopyStream(Byref $oFileSourceStream, Byref $oFileDesStream, $iOptimalTransferSizeBytes)
; Parameters ....: $oFileSourceStream   - [in/out] an object.
;                  $oFileDesStream      - [in/out] an object.
;                  $iOptimalTransferSizeBytes- an integer value.
; Return values .: Total Written Bytes
; Author ........: Danysys
; Modified ......:
; Remarks .......:
; Related .......: _MD_DeviceFileCopyFrom, _MD_DeviceFileCopyTo
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __MD_CopyStream(ByRef $oFileSourceStream, ByRef $oFileDesStream, $iOptimalTransferSizeBytes)
	Local $hResult = 0
	Local $tBufferObjectData = DllStructCreate("byte Data[" & $iOptimalTransferSizeBytes & "]")

	Local $iTotalBytesRead = 0
	Local $iTotalBytesWritten = 0
	Local $iBytesRead = 0
	Local $iBytesWritten = 0

	Do
		$hResult = $oFileSourceStream.Read($tBufferObjectData, $iOptimalTransferSizeBytes, $iBytesRead)
		If SUCCEEDED($S_OK) Then
			$iTotalBytesRead += $iBytesRead
			$hResult = $oFileDesStream.Write($tBufferObjectData, $iBytesRead, $iBytesWritten)
			If SUCCEEDED($hResult) Then
				$iTotalBytesWritten += $iBytesWritten
			EndIf
		EndIf

	Until Not (SUCCEEDED($hResult) And $iBytesRead > 0)
	Return SetError(0, 0, $iTotalBytesWritten)
EndFunc   ;==>__MD_CopyStream

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __MD_CreateClientInformation
; Description ...: Create Default Client Information
; Syntax ........: __MD_CreateClientInformation([$sClientName = "AutoIt Sample Application"[, $sClientMajorVer = $CLIENT_MAJOR_VER[,
;                  $sClientMinorVer = $CLIENT_MINOR_VER[, $sClientRevision = $CLIENT_REVISION[, $iQualityOfService = $SECURITY_IMPERSONATION]]]]])
; Parameters ....: $sClientName         - [optional] a string value. Default is "AutoIt Sample Application".
;                  $sClientMajorVer     - [optional] a string value. Default is $CLIENT_MAJOR_VER.
;                  $sClientMinorVer     - [optional] a string value. Default is $CLIENT_MINOR_VER.
;                  $sClientRevision     - [optional] a string value. Default is $CLIENT_REVISION.
;                  $iQualityOfService   - [optional] an integer value. Default is $SECURITY_IMPERSONATION.
; Return values .: Success      - PortableDeviceValues Object
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: _MD_DeviceOpen
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __MD_CreateClientInformation($sClientName = "AutoIt Sample Application", $sClientMajorVer = $CLIENT_MAJOR_VER, _
		$sClientMinorVer = $CLIENT_MINOR_VER, $sClientRevision = $CLIENT_REVISION, $iQualityOfService = $SECURITY_IMPERSONATION)
	Local $oClientInfo = ObjCreateInterface($sCLSID_PortableDeviceValues, $sIID_IPortableDeviceValues, $sTag_IPortableDeviceValues)
	$oClientInfo.SetStringValue($WPD_CLIENT_NAME, $sClientName)
	$oClientInfo.SetUnsignedIntegerValue($WPD_CLIENT_MAJOR_VERSION, $sClientMajorVer)
	$oClientInfo.SetUnsignedIntegerValue($WPD_CLIENT_MINOR_VERSION, $sClientMinorVer)
	$oClientInfo.SetUnsignedIntegerValue($WPD_CLIENT_REVISION, $sClientRevision)
	$oClientInfo.SetUnsignedIntegerValue($WPD_CLIENT_SECURITY_QUALITY_OF_SERVICE, $iQualityOfService) ;SECURITY_IMPERSONATION
	Return $oClientInfo
EndFunc   ;==>__MD_CreateClientInformation

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __MD_CreateGlobalPropertiesDefault
; Description ...: Create Global Properties Object to use in _MD_GetItemPropertiesDefault
; Syntax ........: __MD_CreateGlobalPropertiesDefault(Byref $oContent)
; Parameters ....: $oContent            - [in/out] an object.
; Return values .: None
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: _MD_GetItemPropertiesDefault
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __MD_CreateGlobalPropertiesDefault(ByRef $oContent)
;~ 	If IsObj($__goPropropertiesDefault) Then Return ;Create Always to handle Multiples Device
	Local $pProperties = 0
	$oContent.Properties($pProperties)
	$__goPropropertiesDefault = ObjCreateInterface($pProperties, $sIID_IPortableDeviceProperties, $sTag_IPortableDeviceProperties)
EndFunc   ;==>__MD_CreateGlobalPropertiesDefault

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __MD_FindItemByPath
; Description ...: Find Device Item by Path and Return _MD_GetItemPropertiesDefault
; Syntax ........: __MD_FindItemByPath(Byref $oContent, $sLocation[, $sObjectId = $WPD_DEVICE_OBJECT_ID])
; Parameters ....: $oContent            - [in/out] an object.
;                  $sLocation           - a string value.
;                  $sObjectId           - [optional] a string value. Default is $WPD_DEVICE_OBJECT_ID.
; Return values .: Success      - 2D Array _MD_GetItemPropertiesDefault
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: _MD_DeviceFindFile
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __MD_FindItemByPath(ByRef $oContent, $sLocation, $sObjectId = $WPD_DEVICE_OBJECT_ID)
	Local $pDeviceObjectIDs = 0
	Local $hResult = $oContent.EnumObjects(0, $sObjectId, 0, $pDeviceObjectIDs)
	Local $oDeviceObjectIDs = ObjCreateInterface($pDeviceObjectIDs, $sIID_IEnumPortableDeviceObjectIDs, $sTag_IEnumPortableDeviceObjectIDs)
	Local $sObjID = ""
	Local $aProperties = 0
	Local $sOriginalFileName = ""
	Local $bFound = False
	While $oDeviceObjectIDs.Next(1, $sObjID, 0) = $S_OK
		$aProperties = _MD_GetItemPropertiesDefault($sObjID)
		$sOriginalFileName = $aProperties[3]
		If $sOriginalFileName = "" Then $sOriginalFileName = $aProperties[2] ;Use Object Name If Original File Name is Empty
		$sLocation = __MD_SlidePath($sLocation, $sOriginalFileName) ; $aProperties[3]=Original File Name
		If Not @error Then
			If $sLocation = "" Then ;Found
				$bFound = True
				ExitLoop
			EndIf
			$aProperties = __MD_FindItemByPath($oContent, $sLocation, $sObjID)
			$bFound = True
			ExitLoop
		Else
			ContinueLoop
		EndIf
	WEnd

	Return ($bFound ? $aProperties : 0)
EndFunc   ;==>__MD_FindItemByPath

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __MD_GetDrives
; Description ...: Get Drive/Storage Object of a Device Object
; Syntax ........: __MD_GetDrives(Byref $oCapabilities, Byref $oDeviceProperties)
; Parameters ....: $oCapabilities       - [in/out] an object.
;                  $oDeviceProperties   - [in/out] an object.
; Return values .: Success      - 2D Array Storages/Drives associated with the Device
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: _MD_DeviceGetDrives
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __MD_GetDrives(ByRef $oCapabilities, ByRef $oDeviceProperties)
	Local $aDrives = _MD_GetFunctionalObjects($oCapabilities, $WPD_FUNCTIONAL_CATEGORY_STORAGE)
	If IsArray($aDrives) Then
		Local $aStorageInfo = 0
		Local $aDrives2D[UBound($aDrives)][9]
		For $i = 0 To UBound($aDrives) - 1
			$aStorageInfo = __MD_GetStorageInfo($oDeviceProperties, $aDrives[$i])
			If IsArray($aStorageInfo) Then
				$aDrives2D[$i][0] = $aDrives[$i]
				$aDrives2D[$i][1] = $aStorageInfo[0]
				$aDrives2D[$i][2] = $aStorageInfo[1]
				$aDrives2D[$i][3] = $aStorageInfo[2]
				$aDrives2D[$i][4] = $aStorageInfo[3]
				$aDrives2D[$i][5] = $aStorageInfo[4]
				$aDrives2D[$i][6] = $aStorageInfo[5]
				$aDrives2D[$i][7] = $aStorageInfo[6]
				$aDrives2D[$i][8] = $aStorageInfo[7]
			EndIf
		Next
		Return $aDrives2D
	Else
		Return 0
	EndIf
EndFunc   ;==>__MD_GetDrives

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __MD_GetStorageInfo
; Description ...:
; Syntax ........: _MD_GetStorageInfo($oDeviceProperties, $sObjectId)
; Parameters ....: $oDeviceProperties   - an object.
;                  $sObjectId           - a string value.
; Return values .: Success      - 2D Array Storage Object Information
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: __MD_GetDrives
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __MD_GetStorageInfo($oDeviceProperties, $sObjectId)
	Local $oDeviceKeyCollection = ObjCreateInterface($sCLSID_PortableDeviceKeyCollection, $sIID_IPortableDeviceKeyCollection, $sTag_IPortableDeviceKeyCollection)
	$oDeviceKeyCollection.Add($WPD_STORAGE_TYPE)
	$oDeviceKeyCollection.Add($WPD_STORAGE_FILE_SYSTEM_TYPE)
	$oDeviceKeyCollection.Add($WPD_STORAGE_CAPACITY)
	$oDeviceKeyCollection.Add($WPD_STORAGE_FREE_SPACE_IN_BYTES)
	$oDeviceKeyCollection.Add($WPD_STORAGE_FREE_SPACE_IN_OBJECTS)
	$oDeviceKeyCollection.Add($WPD_STORAGE_DESCRIPTION)
	$oDeviceKeyCollection.Add($WPD_STORAGE_SERIAL_NUMBER)
	$oDeviceKeyCollection.Add($WPD_STORAGE_MAX_OBJECT_SIZE)
	$oDeviceKeyCollection.Add($WPD_STORAGE_CAPACITY_IN_OBJECTS)
	$oDeviceKeyCollection.Add($WPD_STORAGE_ACCESS_CAPABILITY)

	Local $pSupportedDeviceKeyCollection = 0
	$oDeviceProperties.GetSupportedProperties($sObjectId, $pSupportedDeviceKeyCollection)
	Local $oSupportedDeviceKeyCollection = ObjCreateInterface($pSupportedDeviceKeyCollection, $sIID_IPortableDeviceKeyCollection, $sTag_IPortableDeviceKeyCollection)

	Local $pValues = 0
	$oDeviceProperties.GetValues($sObjectId, $oDeviceKeyCollection(), $pValues)
	Local $oValues = ObjCreateInterface($pValues, $sIID_IPortableDeviceValues, $sTag_IPortableDeviceValues)

	Local $iStorageType = _WPD_ValueGetUnsignedInteger($oValues, $WPD_STORAGE_TYPE)
	Local $sStorageSytemType = _WPD_ValueGetString($oValues, $WPD_STORAGE_FILE_SYSTEM_TYPE)
	Local $fStorageCapacity = _WPD_ValueGetUnsignedLargeInteger($oValues, $WPD_STORAGE_CAPACITY)
	Local $fFreeSpaceInBytes = _WPD_ValueGetUnsignedLargeInteger($oValues, $WPD_STORAGE_FREE_SPACE_IN_BYTES)
	Local $fFreeSpaceInObjects = _WPD_ValueGetUnsignedLargeInteger($oValues, $WPD_STORAGE_FREE_SPACE_IN_OBJECTS)
	Local $sDescription = _WPD_ValueGetString($oValues, $WPD_STORAGE_DESCRIPTION)
	Local $sSerialNumber = _WPD_ValueGetString($oValues, $WPD_STORAGE_SERIAL_NUMBER)
	Local $fMaxObjectSize = _WPD_ValueGetUnsignedLargeInteger($oValues, $WPD_STORAGE_MAX_OBJECT_SIZE)
	Local $fCapacityInObjects = _WPD_ValueGetUnsignedLargeInteger($oValues, $WPD_STORAGE_CAPACITY_IN_OBJECTS)
	Local $iAccessCapability = _WPD_ValueGetUnsignedInteger($oValues, $WPD_STORAGE_CAPACITY_IN_OBJECTS)

	Local $aStorageInfo[] = [$iStorageType, $sStorageSytemType, $fStorageCapacity, $fFreeSpaceInBytes, $fFreeSpaceInObjects, _
			$sDescription, $sSerialNumber, $fMaxObjectSize, $fCapacityInObjects, $iAccessCapability]

	Return $aStorageInfo
EndFunc   ;==>__MD_GetStorageInfo

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __MD_ListRecursive
; Description ...:
; Syntax ........: __MD_ListRecursive(Byref $oContent, $sObjectId, $sDirFullPath, $iFullProperties, $iListOnlyFolder)
; Parameters ....: $oContent            - [in/out] an object.
;                  $sObjectId           - a string value.
;                  $sDirFullPath        - a string value.
;                  $iFullProperties     - an integer value.
;                  $iListOnlyFolder     - an integer value.
; Return values .: Item Count
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: _MD_DeviceFileListToArrayRec, _MD_DeviceDirectoryListToArrayRec
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __MD_ListRecursive(ByRef $oContent, $sObjectId, $sDirFullPath, $iFullProperties, $iListOnlyFolder)
	Local $pDeviceObjectIDs = 0
	Local $sObjID = $sObjectId
	Local $hResult = $oContent.EnumObjects(0, $sObjectId, 0, $pDeviceObjectIDs)
	Local $oDeviceObjectIDs = ObjCreateInterface($pDeviceObjectIDs, $sIID_IEnumPortableDeviceObjectIDs, $sTag_IEnumPortableDeviceObjectIDs)
	If _MD_ItemIsContainer($sObjID) Then $sDirFullPath &= "" & _MD_GetItemProperty($__goPropropertiesDefault, $sObjID, $WPD_OBJECT_ORIGINAL_FILE_NAME, _WPD_ValueGetString)
	Local $iFetched = 0
	Local $sOriginalFileName = ""
	$sDirFullPath &= (StringRight($sDirFullPath, 1) = "\" ? "" : "\")
	Local Static $iCount = 0
	Local $aProperties = 0
	Local $iBoundBasic = UBound($__gaFilesRec2D)
	Local $iBound = UBound($__gaFilesRec2DBasic)
	While $oDeviceObjectIDs.Next(1, $sObjID, $iFetched) = $S_OK
		If $iListOnlyFolder Then
			If Not _MD_ItemIsFolder($sObjID) Then ContinueLoop
		EndIf
		$sOriginalFileName = _MD_GetItemProperty($__goPropropertiesDefault, $sObjID, $WPD_OBJECT_ORIGINAL_FILE_NAME, _WPD_ValueGetString)
		If $iFullProperties Then
			$aProperties = _MD_GetItemPropertiesDefault($sObjID)
			$__gaFilesRec2D[$iCount][0] = $sObjID
			$__gaFilesRec2D[$iCount][1] = $aProperties[1]
			$__gaFilesRec2D[$iCount][2] = $aProperties[2]
			$__gaFilesRec2D[$iCount][3] = $aProperties[3]
			$__gaFilesRec2D[$iCount][4] = $aProperties[4]
			$__gaFilesRec2D[$iCount][5] = $aProperties[5]
			$__gaFilesRec2D[$iCount][6] = $aProperties[6]
			$__gaFilesRec2D[$iCount][7] = $aProperties[7]
			$__gaFilesRec2D[$iCount][8] = $aProperties[8]
			$__gaFilesRec2D[$iCount][9] = $aProperties[9]
			$__gaFilesRec2D[$iCount][10] = $aProperties[10]
			$__gaFilesRec2D[$iCount][11] = $aProperties[11]
			$__gaFilesRec2D[$iCount][12] = $aProperties[12]
			$__gaFilesRec2D[$iCount][13] = $aProperties[13]
			$__gaFilesRec2D[$iCount][14] = $aProperties[14]
			$__gaFilesRec2D[$iCount][15] = $aProperties[15]
			$__gaFilesRec2D[$iCount][16] = $aProperties[16]
			$__gaFilesRec2D[$iCount][17] = $aProperties[17]
			$__gaFilesRec2D[$iCount][18] = $sDirFullPath & $sOriginalFileName
			$iCount += 1
		Else
			$__gaFilesRec2DBasic[$iCount][0] = $sObjID
			$__gaFilesRec2DBasic[$iCount][1] = $sDirFullPath & $sOriginalFileName
			$iCount += 1
		EndIf

		If $iFullProperties Then
			If $iCount = $iBound Then ReDim $__gaFilesRec2D[$iBound + 1000][19]
		Else
			If $iCount = $iBoundBasic Then ReDim $__gaFilesRec2DBasic[$iBoundBasic + 1000][2]
		EndIf
		$iCount = __MD_ListRecursive($oContent, $sObjID, $sDirFullPath, $iFullProperties, $iListOnlyFolder)
	WEnd
	Local $iRetCount = $iCount
	$iCount = 0 ;Restore Count
	Return $iRetCount
EndFunc   ;==>__MD_ListRecursive

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceClose
; Description ...: Close Device Object
; Syntax ........: _MD_DeviceClose(Byref $oDevice)
; Parameters ....: $oDevice             - Device Object
; Return values .: None
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: _MD_DeviceOpen
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceClose(ByRef $oDevice)
	If IsObj($oDevice) Then $oDevice.Close()
EndFunc   ;==>_MD_DeviceClose

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceCreateDirectory
; Description ...: Create a Directory
; Syntax ........: _MD_DeviceCreateDirectory(Byref $oDevice, $sDirPath)
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sDirPath            - a string value.
; Return values .: Success      - ObjectId
;                  Failure      - ""
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceCreateDirectory(ByRef $oDevice, $sDirPath)
	Local $sObjectId = ""
	$sDirPath = StringRight($sDirPath, 1) = "\" ? StringLeft($sDirPath, StringLen($sDirPath) - 1) : $sDirPath
	Local $aFile = _MD_DeviceFindFile($oDevice, $sDirPath)
	If IsArray($aFile) Then Return SetError(1, 0, $aFile[0]) ;File Folder Already Exist
	Local $sDirName = __MD_GetLastDirectoryName($sDirPath)
	$sDirPath = __MD_RemoveLastDirectory($sDirPath)
	$aFile = _MD_DeviceFindFile($oDevice, $sDirPath)
	If Not IsArray($aFile) Then Return SetError(2, 0, "") ;Parent Folder Not Found
	Local $sParentId = $aFile[0]
	Local $oContent = _MPD_DeviceContent($oDevice)
	Local $oValues = ObjCreateInterface($sCLSID_PortableDeviceValues, $sIID_IPortableDeviceValues, $sTag_IPortableDeviceValues)

	$oValues.SetGuidValue($WPD_OBJECT_CONTENT_TYPE, $WPD_CONTENT_TYPE_FOLDER)
	$oValues.SetStringValue($WPD_OBJECT_NAME, $sDirName)
	$oValues.SetStringValue($WPD_OBJECT_ORIGINAL_FILE_NAME, $sDirName)
	$oValues.SetStringValue($WPD_OBJECT_PARENT_ID, $sParentId)

	Local $hResult = $oContent.CreateObjectWithPropertiesOnly($oValues(), $sObjectId)
	Return ($hResult = $S_OK ? SetError(0, 0, $sObjectId) : SetError(1, $hResult, ""))
EndFunc   ;==>_MD_DeviceCreateDirectory

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceDirectoryListToArrayRec
; Description ...:  List Directories Recursive
; Syntax ........: _MD_DeviceDirectoryListToArrayRec(Byref $oDevice, $sPath[, $iFullProperties = False])
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sPath               - a string value.
;                  $iFullProperties     - [optional] an integer value. Default is False.
; Return values .: Success      - Return 2D Array ObjectId + FullPath or _MD_GetItemPropertiesDefault + Full Path
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: _MD_DeviceFileListToArrayRec
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceDirectoryListToArrayRec(ByRef $oDevice, $sPath, $iFullProperties = False)
	Return _MD_DeviceFileListToArrayRec($oDevice, $sPath, $iFullProperties, True)
EndFunc   ;==>_MD_DeviceDirectoryListToArrayRec

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceDiretoryListToArray
; Description ...: List Directories
; Syntax ........: _MD_DeviceDiretoryListToArray(Byref $oDevice, $sPath[, $i2DArrayProperties = False])
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sPath               - a string value.
;                  $i2DArrayProperties  - [optional] an integer value. Default is False.
; Return values .: Success      -1D/2D Array ObjectId Item Properties
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: _MD_DeviceFileListToArray
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceDiretoryListToArray(ByRef $oDevice, $sPath, $i2DArrayProperties = False)
	Return _MD_DeviceFileListToArray($oDevice, $sPath, $i2DArrayProperties, True)
EndFunc   ;==>_MD_DeviceDiretoryListToArray

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceFileCopyFrom
; Description ...: Copy a File From Device
; Syntax ........: _MD_DeviceFileCopyFrom(Byref $oDevice, $sPathFrom, $sPathTo[, $bOverWrite = False])
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sPathFrom           - a string value.
;                  $sPathTo             - a string value.
;                  $bOverWrite          - [optional] a boolean value. Default is False.
; Return values .: Success      - File Path
;                               - @extended=TotalBytesWritten
;                  Failure      - "" or File Path + @Error
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceFileCopyFrom(ByRef $oDevice, $sPathFrom, $sPathTo, $bOverWrite = False)
	Local $sFileName = ""
	If FileExists($sPathTo) And StringInStr(FileGetAttrib($sPathTo), "D") Then
		$sFileName = __MD_GetLastDirectoryName($sPathFrom)
		$sPathTo = StringRight($sPathTo, 1) = "\" ? StringLeft($sPathTo, StringLen($sPathTo) - 1) : $sPathTo
		$sPathTo = $sPathTo & "\" & $sFileName
	Else
		$sFileName = __MD_GetLastDirectoryName($sPathTo)
		$sPathTo = __MD_RemoveLastDirectory($sPathTo) & $sFileName
	EndIf

	If FileExists($sPathTo) And $bOverWrite = False And StringInStr(FileGetAttrib($sPathTo), "A") Then
		Return SetError(1, 0, $sPathTo)      ; File Found No Ovewrite
	EndIf

	If FileExists($sPathTo) And $bOverWrite And StringInStr(FileGetAttrib($sPathTo), "A") Then
		FileDelete($sPathTo)
	EndIf

	Local $aFile = _MD_DeviceFindFile($oDevice, $sPathFrom)
	If Not IsArray($aFile) Or $aFile[17] = True Then Return SetError(2, 0, "") ; File Not Found or is Directory
	Local $sObjectId = $aFile[0]
	Local $oContent = _MPD_DeviceContent($oDevice)

	Local $pResources = 0
	Local $hResult = $oContent.Transfer($pResources)
	Local $oResources = ObjCreateInterface($pResources, $sIID_IPortableDeviceResources, $sTag_IPortableDeviceResources)

	Local $iOptimalTransferSizeBytes = 0
	Local $pFileSourceStream = 0
	$hResult = $oResources.GetStream($sObjectId, $WPD_RESOURCE_DEFAULT, $STGM_READ, $iOptimalTransferSizeBytes, $pFileSourceStream)
	Local $oFileSourceStream = ObjCreateInterface($pFileSourceStream, $sIID_IStream, $sTag_IStream)
	If Not IsObj($oFileSourceStream) Then Return SetError(3, 0, "")

	Local $pProperties = 0
	$oContent.Properties($pProperties)
	Local $oProperties = ObjCreateInterface($pProperties, $sIID_IPortableDeviceProperties, $sTag_IPortableDeviceProperties)

;~ 	Local $sOriginalFileName = _MD_GetItemProperty($oProperties, $sObjectId, $WPD_OBJECT_ORIGINAL_FILE_NAME, _WPD_ValueGetString)
	Local $sOriginalFileName = $aFile[3]
	Local $iFileFromSize = $aFile[6]
	Local $pFileDesStream = __MD_SHCreateStreamOnFileEx($sPathTo, BitOR($STGM_CREATE, $STGM_WRITE), $FILE_ATTRIBUTE_NORMAL)
	Local $oFileDesStream = ObjCreateInterface($pFileDesStream, $sIID_IStream, $sTag_IStream)
	If Not IsObj($oFileDesStream) Then Return SetError(3, 0, "")

	Local $iWrittenBytes = __MD_CopyStream($oFileSourceStream, $oFileDesStream, $iOptimalTransferSizeBytes)

	Return $iFileFromSize = $iWrittenBytes ? SetError(0, $iWrittenBytes, $sPathTo) : SetError(1, 0, "")
EndFunc   ;==>_MD_DeviceFileCopyFrom

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceFileCopyTo
; Description ...:
; Syntax ........: _MD_DeviceFileCopyTo(Byref $oDevice, $sPathFrom, $sPathTo[, $bOverWwrite = False])
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sPathFrom           - a string value.
;                  $sPathTo             - a string value.
;                  $bOverWrite         - [optional] a boolean value. Default is False.
; Return values .: Success      - ObjectId
;                               - @extended=TotalBytesWritten
;                  Failure      - ""
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceFileCopyTo(ByRef $oDevice, $sPathFrom, $sPathTo, $bOverWrite = False)
	If Not FileExists($sPathFrom) Or StringInStr(FileGetAttrib($sPathFrom), "D") Then Return SetError(1, 0, "") ; File Not Found or is Directory
	Local $aFile = _MD_DeviceFindFile($oDevice, $sPathTo)

	If $bOverWrite Then
		If IsArray($aFile) And $aFile[17] = False Then _MD_DeviceItemDelete($oDevice, $sPathTo)
	Else
		If IsArray($aFile) And ($aFile[17] = False And $aFile[16] = False) Then Return SetError(2, 0, $aFile[0]) ;File Already Exist
	EndIf

	Local $sFileName = ""
	If IsArray($aFile) And ($aFile[17] Or $aFile[16]) Then  ;Exist and Is Directory
		$sFileName = __MD_GetLastDirectoryName($sPathFrom)
	Else
		$sFileName = __MD_GetLastDirectoryName($sPathTo)
		$sPathTo = __MD_RemoveLastDirectory($sPathTo)
		$aFile = _MD_DeviceFindFile($oDevice, $sPathTo)
		If Not IsArray($aFile) Or Not ($aFile[17] Or $aFile[16]) Then Return SetError(3, 0, "") ; Not Found or is not Directory
	EndIf

	Local $sParentId = $aFile[0]
	Local $pFileSourceStream = __MD_SHCreateStreamOnFileEx($sPathFrom, $STGM_READ, $FILE_ATTRIBUTE_NORMAL)
	Local $oFileSourceStream = ObjCreateInterface($pFileSourceStream, $sIID_IStream, $sTag_IStream)
	If Not IsObj($oFileSourceStream) Then Return SetError(4, 0, "")

	Local $oContent = _MPD_DeviceContent($oDevice)
	;// Fills out the required properties for specific WPD content types.
	Local $oValuesTemp = ObjCreateInterface($sCLSID_PortableDeviceValues, $sIID_IPortableDeviceValues, $sTag_IPortableDeviceValues)
	$oValuesTemp.SetStringValue($WPD_OBJECT_PARENT_ID, $sParentId)
	Local $tSTATSTG = DllStructCreate($sTag_STATSTG)
	$oFileSourceStream.Stat($tSTATSTG, $STATFLAG_NONAME)

	Local $iSize = $tSTATSTG.cbSize
	$oValuesTemp.SetUnsignedLargeIntegerValue($WPD_OBJECT_SIZE, $iSize)
	$oValuesTemp.SetStringValue($WPD_OBJECT_NAME, $sFileName)
	$oValuesTemp.SetStringValue($WPD_OBJECT_ORIGINAL_FILE_NAME, $sFileName)
	$oValuesTemp.SetStringValue($WPD_OBJECT_PARENT_ID, $sParentId)

	Local $pFileDesStream = 0
	Local $iOptimalTransferSizeBytes = 0
	$oContent.CreateObjectWithPropertiesAndData($oValuesTemp(), $pFileDesStream, $iOptimalTransferSizeBytes, 0)
	Local $oFileDesStream = ObjCreateInterface($pFileDesStream, $sIID_IPortableDeviceDataStream, $sTag_IPortableDeviceDataStream)
	If Not IsObj($oFileDesStream) Then Return SetError(4, 0, "")

	Local $iWrittenBytes = __MD_CopyStream($oFileSourceStream, $oFileDesStream, $iOptimalTransferSizeBytes)

	If $iWrittenBytes = $iSize Then
		$oFileDesStream.Commit($STGC_DEFAULT)
	Else
		SetError(5, $iWrittenBytes, "")
	EndIf

	Local $sObjectId = ""
	$oFileDesStream.GetObjectID($sObjectId)

	Return SetError(0, $iWrittenBytes, $sObjectId)
EndFunc   ;==>_MD_DeviceFileCopyTo

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceFileListToArray
; Description ...: List Files/Directories
; Syntax ........: _MD_DeviceFileListToArray(Byref $oDevice, $sPath[, $i2DArrayProperties = False[, $iListOnlyFolder = False]])
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sPath               - a string value.
;                  $i2DArrayProperties  - [optional] an integer value. Default is False.
;                  $iListOnlyFolder     - [optional] an integer value. Default is False.
; Return values .: Success      -1D/2D Array ObjectId Item Properties
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceFileListToArray(ByRef $oDevice, $sPath, $i2DArrayProperties = False, $iListOnlyFolder = False)
	Local $aFile = _MD_DeviceFindFile($oDevice, $sPath)
	Local $sPreservePath = $sPath & (StringRight($sPath, 1) = "\" ? "" : "\") ;Force end with \
	If Not IsArray($aFile) Then Return SetError(1, 0, 0)
	Local $oContent = _MPD_DeviceContent($oDevice)
	__MD_CreateGlobalPropertiesDefault($oContent)
	Local $sObjectId = $aFile[0]
	Local $pDeviceObjectIDs = 0
	Local $hResult = $oContent.EnumObjects(0, $sObjectId, 0, $pDeviceObjectIDs)
	Local $oDeviceObjectIDs = ObjCreateInterface($pDeviceObjectIDs, $sIID_IEnumPortableDeviceObjectIDs, $sTag_IEnumPortableDeviceObjectIDs)
	Local $aProperties = 0
	Local $iBound = 10000
	Local $aFiles2D[$iBound][19]
	Local $aFiles[$iBound]
	Local $iCount = 0
	Local $iFetched = 0
	While $oDeviceObjectIDs.Next(1, $sObjectId, $iFetched) = $S_OK
		If $iListOnlyFolder Then
			If Not _MD_ItemIsFolder($sObjectId) Then ContinueLoop
		EndIf
		If $i2DArrayProperties Then
			$aProperties = _MD_GetItemPropertiesDefault($sObjectId)
			$aFiles2D[$iCount][0] = $sObjectId
			$aFiles2D[$iCount][1] = $aProperties[1]
			$aFiles2D[$iCount][2] = $aProperties[2]
			$aFiles2D[$iCount][3] = $aProperties[3]
			$aFiles2D[$iCount][4] = $aProperties[4]
			$aFiles2D[$iCount][5] = $aProperties[5]
			$aFiles2D[$iCount][6] = $aProperties[6]
			$aFiles2D[$iCount][7] = $aProperties[7]
			$aFiles2D[$iCount][8] = $aProperties[8]
			$aFiles2D[$iCount][9] = $aProperties[9]
			$aFiles2D[$iCount][10] = $aProperties[10]
			$aFiles2D[$iCount][11] = $aProperties[11]
			$aFiles2D[$iCount][12] = $aProperties[12]
			$aFiles2D[$iCount][13] = $aProperties[13]
			$aFiles2D[$iCount][14] = $aProperties[14]
			$aFiles2D[$iCount][15] = $aProperties[15]
			$aFiles2D[$iCount][16] = $aProperties[16]
			$aFiles2D[$iCount][17] = $aProperties[17]
			$aFiles2D[$iCount][18] = $sPreservePath & $aProperties[3]
		Else
			$aFiles[$iCount] = $sObjectId
		EndIf
		$iCount += 1
		If $iCount = $iBound Then
			$iBound += 1000
			If $i2DArrayProperties Then
				ReDim $aFiles2D[$iBound][19]
			Else
				ReDim $aFiles[$iBound]
			EndIf
		EndIf
	WEnd
	If Not $iCount Then Return SetError(1, 0, 0)

	If $i2DArrayProperties Then
		ReDim $aFiles2D[$iCount][19]
		Return $aFiles2D
	Else
		ReDim $aFiles[$iCount]
		Return $aFiles
	EndIf
EndFunc   ;==>_MD_DeviceFileListToArray

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceFileListToArrayRec
; Description ...: List Files Recursive
; Syntax ........: _MD_DeviceFileListToArrayRec(Byref $oDevice, $sPath[, $iFullProperties = False[, $iListOnlyFolder = False]])
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sPath               - a string value.
;                  $iFullProperties     - [optional] an integer value. Default is False.
;                  $iListOnlyFolder     - [optional] an integer value. Default is False.
; Return values .: Success      - Return 2D Array ObjectId + FullPath or _MD_GetItemPropertiesDefault + Full Path
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: __MD_ListRecursive
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceFileListToArrayRec(ByRef $oDevice, $sPath, $iFullProperties = False, $iListOnlyFolder = False)
	Local $aFile = _MD_DeviceFindFile($oDevice, $sPath)
	If Not IsArray($aFile) Or Not ($aFile[16] Or $aFile[17]) Then Return SetError(1, 0, 0) ;Is Not Folder or Storage
	$sPath &= (StringRight($sPath, 1) = "\" ? "" : "\") ;Force end with \
	Local $sBasePath = (StringInStr($sPath, "\", 2, 2) ? __MD_RemoveLastDirectory2($sPath) : $sPath)
	Local $sObjectId = $aFile[0]
	Local $oContent = _MPD_DeviceContent($oDevice)
	__MD_CreateGlobalPropertiesDefault($oContent)
	ReDim $__gaFilesRec2D[1000][19]
	ReDim $__gaFilesRec2DBasic[1000][2]
	Local $iCount = __MD_ListRecursive($oContent, $sObjectId, $sBasePath, $iFullProperties, $iListOnlyFolder)

	If $iFullProperties Then
		ReDim $__gaFilesRec2D[$iCount][19]
		Return ($iCount ? $__gaFilesRec2D : 0)
	Else
		ReDim $__gaFilesRec2DBasic[$iCount][2]
		Return ($iCount ? $__gaFilesRec2DBasic : 0)
	EndIf
EndFunc   ;==>_MD_DeviceFileListToArrayRec

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceFindFile
; Description ...: Find and Return a File From a Device by Path
; Syntax ........: _MD_DeviceFindFile(Byref $oDevice, $sPath)
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sPath               - Full File Path.
; Return values .: Success      - 2D Array _MD_GetItemPropertiesDefault
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceFindFile(ByRef $oDevice, $sPath)
	Local $oContent = _MPD_DeviceContent($oDevice)
	__MD_CreateGlobalPropertiesDefault($oContent)
	Local $aItem = __MD_FindItemByPath($oContent, $sPath)
	Return $aItem
EndFunc   ;==>_MD_DeviceFindFile

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceGet
; Description ...: Find/Match Media Device
; Syntax ........: _MD_DeviceGet([$sFriendlyName = ""[, $iInstance = 0[, $sManufacturer = ""[, $sDescription = ""[,
;                  $sDeviceType = ""]]]]])
; Parameters ....: $sFriendlyName       - [optional] Device FriendlyName String Default is "".
;                  $iInstance           - [optional] Device Instance Index. Default is 0.
;                  $sManufacturer       - [optional] Device Manufacturer String. Default is "".
;                  $sDescription        - [optional] Device Description  String. Default is "".
;                  $sDeviceType         - [optional] Device DeviceType  String. Default is "".
; Return values .: Success      - PnPDeviceID String
;                  Failure      - ""
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceGet($sFriendlyName = "", $iInstance = 0, $sManufacturer = "", $sDescription = "", $sDeviceType = "")
	If Not @NumParams Then Return _MD_DeviceGetFirst()
	Local $aFind = 0
	Local $iIndex = 0
	Local $aDevices = _MD_DevicesList()
	Local $sPnPDeviceID = ""
	If Not IsArray($aDevices) Then Return SetError(1, 0, $sPnPDeviceID)

	;Match All Parameters
	If $sFriendlyName <> "" And $sManufacturer <> "" And $sDescription <> "" And $sDeviceType <> "" Then
		$aFind = _ArrayFindAll($aDevices, $sFriendlyName, 0, 0, 0, 0, $eMD_DEVICE_FrienlyName)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
;~ 			_ArrayDisplay($aFind)
			$iIndex = Number($aFind[$iInstance])
			If $aDevices[$iIndex][$eMD_Device_Manufacturer] = $sManufacturer And _
					$aDevices[$iIndex][$eMD_Device_Description] = $sDescription And _
					$aDevices[$iIndex][$eMD_Device_DeviceType] = $sDeviceType Then

				$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
				Return SetError(0, $iIndex, $sPnPDeviceID)
			EndIf
			Return SetError(1, 0, $sPnPDeviceID) ;Not Match
		Else
			Return SetError(2, 0, $sPnPDeviceID) ;Not Found or Instance higher
		EndIf
	EndIf

	;Match FriendlyName and Instance
	If $sFriendlyName <> "" And $sManufacturer = "" And $sDescription = "" And $sDeviceType = "" Then
		$aFind = _ArrayFindAll($aDevices, $sFriendlyName, 0, 0, 0, 0, $eMD_DEVICE_FrienlyName)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
			Return SetError(0, $iIndex, $sPnPDeviceID)
		EndIf
	EndIf

	;Match FriendlyName, Manufacturer and Instance
	If $sFriendlyName <> "" And $sManufacturer <> "" And $sDescription = "" And $sDeviceType = "" Then
		$aFind = _ArrayFindAll($aDevices, $sFriendlyName, 0, 0, 0, 0, $eMD_DEVICE_FrienlyName)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			If $aDevices[$iIndex][$eMD_Device_Manufacturer] = $sManufacturer Then
				$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
				Return SetError(0, $iIndex, $sPnPDeviceID)
			EndIf
			Return SetError(1, 0, $sPnPDeviceID) ;Not Match
		EndIf
	EndIf

	;Match FriendlyName, Description and Instance
	If $sFriendlyName <> "" And $sManufacturer = "" And $sDescription <> "" And $sDeviceType = "" Then
		$aFind = _ArrayFindAll($aDevices, $sFriendlyName, 0, 0, 0, 0, $eMD_DEVICE_FrienlyName)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			If $aDevices[$iIndex][$eMD_Device_Description] = $sDescription Then
				$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
				Return SetError(0, $iIndex, $sPnPDeviceID)
			EndIf
			Return SetError(1, 0, $sPnPDeviceID)     ;Not Match
		EndIf
	EndIf

	;Match FriendlyName, DeviceType and Instance
	If $sFriendlyName <> "" And $sManufacturer = "" And $sDescription = "" And $sDeviceType <> "" Then
		$aFind = _ArrayFindAll($aDevices, $sFriendlyName, 0, 0, 0, 0, $eMD_DEVICE_FrienlyName)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			If $aDevices[$iIndex][$eMD_Device_DeviceType] = $sDeviceType Then
				$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
				Return SetError(0, $iIndex, $sPnPDeviceID)
			EndIf
			Return SetError(1, 0, $sPnPDeviceID)  ;Not Match
		EndIf
	EndIf

	;Match Manufacturer, Description, DeviceType and Instance
	If $sFriendlyName = "" And $sManufacturer <> "" And $sDescription <> "" And $sDeviceType <> "" Then
		$aFind = _ArrayFindAll($aDevices, $sManufacturer, 0, 0, 0, 0, $eMD_Device_Manufacturer)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			If $aDevices[$iIndex][$eMD_Device_Description] = $sDescription And _
					$aDevices[$iIndex][$eMD_Device_DeviceType] = $sDeviceType Then
				$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
				Return SetError(0, $iIndex, $sPnPDeviceID)
			EndIf
			Return SetError(1, 0, $sPnPDeviceID)  ;Not Match
		EndIf
	EndIf

	;Match Manufacturer, Description, and Instance
	If $sFriendlyName = "" And $sManufacturer <> "" And $sDescription <> "" And $sDeviceType = "" Then
		$aFind = _ArrayFindAll($aDevices, $sManufacturer, 0, 0, 0, 0, $eMD_Device_Manufacturer)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			If $aDevices[$iIndex][$eMD_Device_Description] = $sDescription Then
				$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
				Return SetError(0, $iIndex, $sPnPDeviceID)
			EndIf
			Return SetError(1, 0, $sPnPDeviceID)  ;Not Match
		EndIf
	EndIf

	;Match Manufacturer, DeviceType, and Instance
	If $sFriendlyName = "" And $sManufacturer <> "" And $sDescription = "" And $sDeviceType <> "" Then
		$aFind = _ArrayFindAll($aDevices, $sManufacturer, 0, 0, 0, 0, $eMD_Device_Manufacturer)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			If $aDevices[$iIndex][$eMD_Device_DeviceType] = $sDeviceType Then
				$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
				Return SetError(0, $iIndex, $sPnPDeviceID)
			EndIf
			Return SetError(1, 0, $sPnPDeviceID)  ;Not Match
		EndIf
	EndIf

	;Match Description, DeviceType, and Instance
	If $sFriendlyName = "" And $sManufacturer = "" And $sDescription <> "" And $sDeviceType <> "" Then
		$aFind = _ArrayFindAll($aDevices, $sDescription, 0, 0, 0, 0, $eMD_Device_Description)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			If $aDevices[$iIndex][$eMD_Device_DeviceType] = $sDeviceType Then
				$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
				Return SetError(0, $iIndex, $sPnPDeviceID)
			EndIf
			Return SetError(1, 0, $sPnPDeviceID)  ;Not Match
		EndIf
	EndIf

	;Match Manufacturer and Instance
	If $sFriendlyName = "" And $sManufacturer <> "" And $sDescription = "" And $sDeviceType = "" Then
		$aFind = _ArrayFindAll($aDevices, $sManufacturer, 0, 0, 0, 0, $eMD_Device_Manufacturer)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
			Return SetError(0, $iIndex, $sPnPDeviceID)
		EndIf
	EndIf

	;Match Description and Instance
	If $sFriendlyName = "" And $sManufacturer = "" And $sDescription <> "" And $sDeviceType = "" Then
		$aFind = _ArrayFindAll($aDevices, $sDescription, 0, 0, 0, 0, $eMD_Device_Description)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
			Return SetError(0, $iIndex, $sPnPDeviceID)
		EndIf
	EndIf

	;Match DeviceType and Instance
	If $sFriendlyName = "" And $sManufacturer = "" And $sDescription = "" And $sDeviceType <> "" Then
		$aFind = _ArrayFindAll($aDevices, $sDeviceType, 0, 0, 0, 0, $eMD_Device_DeviceType)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
			Return SetError(0, $iIndex, $sPnPDeviceID)
		EndIf
	EndIf

	;Match Instance
	If $sFriendlyName = "" And $sManufacturer = "" And $sDescription = "" And $sDeviceType = "" Then
		$aFind = _ArrayFindAll($aDevices, $sDeviceType, 0, 0, 0, 0, $eMD_Device_DeviceType)
		If IsArray($aFind) And UBound($aFind) > $iInstance Then
			$iIndex = Number($aFind[$iInstance])
			$sPnPDeviceID = $aDevices[$iIndex][$eMD_Device_PnPID]
			Return SetError(0, $iIndex, $sPnPDeviceID)
		EndIf
	EndIf

	Return $sPnPDeviceID
EndFunc   ;==>_MD_DeviceGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceGetDrives
; Description ...: Return List 2D Array of the Drives associated with a Device Object
; Syntax ........: _MD_DeviceGetDrives(Byref $oDevice)
; Parameters ....: $oDevice             - Device Object
; Return values .: Success      - 2D Array Drives Information
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: __MD_GetDrives
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceGetDrives(ByRef $oDevice)
	Local $oContent = _MPD_DeviceContent($oDevice)
	Local $oCapabilities = _MPD_DeviceCapabilities($oDevice)
	Local $oDeviceProperties = _MPD_DeviceProperties($oContent)
	Local $aDrives = __MD_GetDrives($oCapabilities, $oDeviceProperties)
	Return $aDrives
EndFunc   ;==>_MD_DeviceGetDrives

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceGetFirst
; Description ...: Return First Found Media Device
; Syntax ........: _MD_DeviceGetFirst()
; Parameters ....: None
; Return values .: Success      - PnPDeviceID String
;                  Failure      - ""
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceGetFirst()
	Local $aDevices = _MD_DevicesList()
	Return IsArray($aDevices) ? $aDevices[0][0] : ""
EndFunc   ;==>_MD_DeviceGetFirst

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceItemDelete
; Description ...: Delete a File/Folder
; Syntax ........: _MD_DeviceItemDelete(Byref $oDevice, $sPath[, $iRecursive = False])
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sPath               - a string value.
;                  $iRecursive          - [optional] an integer value. Default is False.
; Return values .: Success      - True
;                  Failure      - False
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceItemDelete(ByRef $oDevice, $sPath, $iRecursive = False)
	Local $aFile = _MD_DeviceFindFile($oDevice, $sPath)
	If Not IsArray($aFile) Then Return SetError(1, 0, "") ; File Not Found
	Local $sObjectId = $aFile[0]
	Local $oVariantCollection = ObjCreateInterface($sCLSID_IPortableDevicePropVariantCollection, $sIID_IPortableDevicePropVariantCollection, $sTag_IPortableDevicePropVariantCollection)
	Local $tPropVariant = _WPD_PropVariant()
	$tPropVariant.vt = 31 ;VT_LPWSTR
	Local $tObjectId = DllStructCreate('wchar Data[' & StringLen($sObjectId) + 1 & ']')
	$tObjectId.Data = $sObjectId
	DllStructSetData($tPropVariant, 5, DllStructGetPtr($tObjectId))

	Local $oContent = _MPD_DeviceContent($oDevice)
	$oVariantCollection.Add($tPropVariant)

	Local $pResult = 0
	Local $hResult = $oContent.Delete(($iRecursive ? $PORTABLE_DEVICE_DELETE_WITH_RECURSION : $PORTABLE_DEVICE_DELETE_NO_RECURSION), $oVariantCollection(), $pResult)
	Return SetError($hResult = $S_OK, $hResult, $hResult = $S_OK)
EndFunc   ;==>_MD_DeviceItemDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceItemRename
; Description ...: Rename a File
; Syntax ........: _MD_DeviceItemRename(Byref $oDevice, $sPath, $sName)
; Parameters ....: $oDevice             - [in/out] an object.
;                  $sPath               - a string value.
;                  $sName               - a string value.
; Return values .: Success      - ObjectId
;                  Failure      - ""
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceItemRename(ByRef $oDevice, $sPath, $sName)
	Local $sNewPath = __MD_ChangePathLastNameWithNewName($sPath, $sName)
	Local $aFile = _MD_DeviceFindFile($oDevice, $sNewPath)
	If IsArray($aFile) Then Return SetError(1, 0, $aFile[0]) ; File Already Exist
	$aFile = _MD_DeviceFindFile($oDevice, $sPath)
	If Not IsArray($aFile) Then Return SetError(1, 0, "") ; File Not Found
	Local $sObjectId = $aFile[0]
	Local $oContent = _MPD_DeviceContent($oDevice)
	Local $oDeviceProperties = _MPD_DeviceProperties($oContent)
	Local $oValues = _WPD_DevicePropertiesValues($oDeviceProperties, $sObjectId)
	Local $hResult = $oValues.SetStringValue($WPD_OBJECT_ORIGINAL_FILE_NAME, $sName)
	$oDeviceProperties.SetValues($sObjectId, $oValues(), 0)
	$aFile = _MD_DeviceFindFile($oDevice, $sNewPath)
	If Not IsArray($aFile) Then Return SetError(1, 0, "") ; File Not Found
	Local $sNewObjectId = $aFile[0]
	Return SetError($hResult = $S_OK, $hResult, $sNewObjectId)
EndFunc   ;==>_MD_DeviceItemRename

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DeviceOpen
; Description ...: Open a Device
; Syntax ........: _MD_DeviceOpen($sPnPDeviceID)
; Parameters ....: $sPnPDeviceID        - a string value.
; Return values .: Success      - PortableDevice Object
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DeviceOpen($sPnPDeviceID)
	Local $hResult = $E_FAIL
	Local $oClientInfo = __MD_CreateClientInformation()
	Local $oDevice = _WPD_CreateDevice()
	$hResult = $oDevice.Open($sPnPDeviceID, $oClientInfo())
	If $hResult = $S_OK Then Return $oDevice
EndFunc   ;==>_MD_DeviceOpen

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_DevicesList
; Description ...: Return 2D Array Of Media Devices
; Syntax ........: _MD_DevicesList()
; Parameters ....: None
; Return values .: Success      - 2D Array Devices Information
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_DevicesList()
	Return _WPDListDevices()
EndFunc   ;==>_MD_DevicesList

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_GetFunctionalObjects
; Description ...: Get Functional Objects
; Syntax ........: _MD_GetFunctionalObjects(Byref $oCapabilities, $tGUID)
; Parameters ....: $oCapabilities       - Device Capabilities object
;                  $tGUID               - WPD Functional Categories
; Return values .: Success      - 1D Array of ObjectIDs
;                  Failure      - 0
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_GetFunctionalObjects(ByRef $oCapabilities, $tGUID)
	Local $pPropVariantCollection = 0
	$oCapabilities.GetFunctionalObjects($tGUID, $pPropVariantCollection)
	Local $oPropVariantCollection = ObjCreateInterface($pPropVariantCollection, $sIID_IPortableDevicePropVariantCollection, $sTag_IPortableDevicePropVariantCollection)
	Local $iCount = 0
	$oPropVariantCollection.GetCount($iCount)
	Local $tProVariant = 0
	Local $aFuncionalObjects[$iCount]
	Local $iCountSuccess = 0
	Local $sString = ""
	For $i = 0 To $iCount - 1
		$tProVariant = _WPD_PropVariant()
		$oPropVariantCollection.GetAt($i, $tProVariant)
		$sString = __WPD_GetString($tProVariant)
		If $sString Then
			$aFuncionalObjects[$iCountSuccess] = $sString
			$iCountSuccess += 1
		EndIf
	Next
	ReDim $aFuncionalObjects[$iCountSuccess]
	Return UBound($aFuncionalObjects) ? $aFuncionalObjects : 0
EndFunc   ;==>_MD_GetFunctionalObjects

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_GetItemPropertiesDefault
; Description ...: Return 2D Array of Supported Properties (PortableDeviceProperties.GetSupportedProperties)
; Syntax ........: _MD_GetItemPropertiesDefault($sObjectId)
; Parameters ....: $sObjectId           - a string value.
; Return values .: 2D Array with Object Information
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: __MD_FindItemByPath
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_GetItemPropertiesDefault($sObjectId)
	Local $pKeyCollection = 0
	Local $pValues = 0
	$__goPropropertiesDefault.GetSupportedProperties($sObjectId, $pKeyCollection)
	$__goPropropertiesDefault.GetValues($sObjectId, $pKeyCollection, $pValues)
;~ 	Local $oKeyCollection = ObjCreateInterface($pKeyCollection, $sIID_IPortableDeviceKeyCollection, $sTag_IPortableDeviceKeyCollection)
	Local $oValues = ObjCreateInterface($pValues, $sIID_IPortableDeviceValues, $sTag_IPortableDeviceValues)

	Local $sCLSID = _WPD_ValueGetString($oValues, $WPD_OBJECT_CONTENT_TYPE)
	Local $sObjName = _WPD_ValueGetString($oValues, $WPD_OBJECT_NAME)
	Local $sObjOriginalName = _WPD_ValueGetString($oValues, $WPD_OBJECT_ORIGINAL_FILE_NAME)
	Local $sObjHintLocation = _WPD_ValueGetString($oValues, $WPD_OBJECT_HINT_LOCATION_DISPLAY_NAME)
	Local $sObjContainerId = _WPD_ValueGetString($oValues, $WPD_OBJECT_CONTAINER_FUNCTIONAL_OBJECT_ID)
	Local $sObjSize = _WPD_ValueGetUnsignedLargeInteger($oValues, $WPD_OBJECT_SIZE)
	Local $sDateCreated = _WPD_ValueGetUnsignedLargeInteger($oValues, $WPD_OBJECT_DATE_CREATED)
	Local $sDateModify = _WPD_ValueGetUnsignedLargeInteger($oValues, $WPD_OBJECT_DATE_MODIFIED)
	Local $sDateAuthored = _WPD_ValueGetUnsignedLargeInteger($oValues, $WPD_OBJECT_DATE_AUTHORED)
	Local $bCanDelete = _WPD_ValueGetBool($oValues, $WPD_OBJECT_CAN_DELETE)
	Local $bIsSystem = _WPD_ValueGetBool($oValues, $WPD_OBJECT_ISSYSTEM)
	Local $bIsHidden = _WPD_ValueGetBool($oValues, $WPD_OBJECT_ISHIDDEN)
	Local $bIsDRMProtected = _WPD_ValueGetBool($oValues, $WPD_OBJECT_IS_DRM_PROTECTED)
	Local $sParentId = _WPD_ValueGetString($oValues, $WPD_OBJECT_PARENT_ID)
	Local $sPercistentId = _WPD_ValueGetString($oValues, $WPD_OBJECT_PERSISTENT_UNIQUE_ID)
	Local $bIsStorage = _WPD_GUIDIsFunctionalObject($sCLSID)
	Local $bIsFolder = _WPD_GUIDIsFolder($sCLSID)
	$oValues = 0

	Local $aProperties[] = [$sObjectId, $sCLSID, $sObjName, $sObjOriginalName, $sObjHintLocation, $sObjContainerId, _
			$sObjSize, $sDateCreated, $sDateModify, $sDateAuthored, $bCanDelete, $bIsSystem, $bIsHidden, $bIsDRMProtected, $sParentId, $sPercistentId, $bIsStorage, $bIsFolder]

	Return $aProperties

EndFunc   ;==>_MD_GetItemPropertiesDefault

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_GetItemSingleProperties
; Description ...: Get ObjectId Property
; Syntax ........: _MD_GetItemSingleProperties($sObjectId, $pFunction, $tProperty)
; Parameters ....: $sObjectId           - a string value.
;                  $pFunction           - Type of Value to Query.
;                  $tProperty           - WPD_OBJECT GUID Property
; Return values .: _WPD_ValueGet* Value
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: _WPD_ValueGet*
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_GetItemSingleProperties($sObjectId, $pFunction, $tProperty)
	Local $pKeyCollection = 0
	Local $pValues = 0
	$__goPropropertiesDefault.GetSupportedProperties($sObjectId, $pKeyCollection)
	$__goPropropertiesDefault.GetValues($sObjectId, $pKeyCollection, $pValues)
;~ 	Local $oKeyCollection = ObjCreateInterface($pKeyCollection, $sIID_IPortableDeviceKeyCollection, $sTag_IPortableDeviceKeyCollection)
	Local $oValues = ObjCreateInterface($pValues, $sIID_IPortableDeviceValues, $sTag_IPortableDeviceValues)
	Local $sProperty = $pFunction($oValues, $tProperty)
	Return $sProperty
EndFunc   ;==>_MD_GetItemSingleProperties

Func _MD_ItemIsContainer($sObjectId)
	Local $sGUID = _MD_GetItemSingleProperties($sObjectId, _WPD_ValueGetString, $WPD_OBJECT_CONTENT_TYPE)
	Return _WPD_GUIDIsContainer($sGUID) Or _WPD_GUIDIsFolder($sGUID) ;_WPD_GUIDIsFunctionalObject
EndFunc   ;==>_MD_ItemIsContainer

; #FUNCTION# ====================================================================================================================
; Name ..........: _MD_ItemIsFolder
; Description ...: Check If an ObjectId is a Folder
; Syntax ........: _MD_ItemIsFolder($sObjectId)
; Parameters ....: $sObjectId           - a string value.
; Return values .: True/False
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _MD_ItemIsFolder($sObjectId)
	Local $sGUID = _MD_GetItemSingleProperties($sObjectId, _WPD_ValueGetString, $WPD_OBJECT_CONTENT_TYPE)
	Return _WPD_GUIDIsFolder($sGUID)
EndFunc   ;==>_MD_ItemIsFolder
#EndRegion UDF Funtions

#Region Utils

;Change Last Name Path
Func __MD_ChangePathLastNameWithNewName($sPath, $sName)
	$sPath = StringRight($sPath, 1) = "\" ? StringLeft($sPath, StringLen($sPath) - 1) : $sPath
	$sPath = __MD_RemoveLastDirectory($sPath) & $sName
	Return $sPath
EndFunc   ;==>__MD_ChangePathLastNameWithNewName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __MD_ErrorHandler
; Description ...: Object Error Handler
; Syntax ........: __MD_ErrorHandler($oError)
; Parameters ....: $oError              - an object.
; Return values .: None
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __MD_ErrorHandler($oError)
	; Do Nothing
EndFunc   ;==>__MD_ErrorHandler

;Get Name of Last Directory
Func __MD_GetLastDirectoryName($sDirPath)
	Return StringReplace(StringMid($sDirPath, StringInStr($sDirPath, "\", 2, -1)), "\", "", 0, 2)
EndFunc   ;==>__MD_GetLastDirectoryName

;Remove Last Directory
Func __MD_RemoveLastDirectory($sDirPath)
	Return StringLeft($sDirPath, StringInStr($sDirPath, "\", 2, -1))
EndFunc   ;==>__MD_RemoveLastDirectory

;Remove Last Directory Force Remove / at the end
Func __MD_RemoveLastDirectory2($sDirPath)
	Return __MD_RemoveLastDirectory(StringLeft($sDirPath, StringLen($sDirPath) - 1))
EndFunc   ;==>__MD_RemoveLastDirectory2

;Create File Stream
Func __MD_SHCreateStreamOnFileEx($sFilePath, $igrfMode, $idwAttributes, $bCreate = False)
	Local $aCall = DllCall("shlwapi.dll", "long", "SHCreateStreamOnFileEx", "wstr", $sFilePath, "dword", $igrfMode, "dword", $idwAttributes, "bool", $bCreate, "ptr", 0, "ptr*", 0)
	If $aCall[0] = $S_OK Then
		Return $aCall[6]
	EndIf
	Return 0
EndFunc   ;==>__MD_SHCreateStreamOnFileEx

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __MD_SlidePath
; Description ...: Jump to next backslash Path
; Syntax ........: __MD_SlidePath($sLocation, $sObjectName)
; Parameters ....: $sLocation           - a string value.
;                  $sObjectName         - a string value.
; Return values .: Success      - next backslash Path location
;                  Failure      - set @error=1
; Author ........: DanysysTeam
; Modified ......:
; Remarks .......:
; Related .......: __MD_FindItemByPath
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __MD_SlidePath($sLocation, $sObjectName)
	Local $sNewLocation = ""
	If StringMid($sLocation, 1, 1) = "\" Then $sLocation = StringMid($sLocation, 2)
	Local $iPos = StringInStr($sLocation, "\", 2)
	Local $sPathPart = $iPos ? StringMid($sLocation, 1, $iPos - 1) : $sLocation
	$sNewLocation = StringReplace($sLocation, $sPathPart, "", 1, 2)
	If $sObjectName = $sPathPart Then
		$sNewLocation = StringReplace($sLocation, $sPathPart, "", 1, 2)
		If StringLeft($sNewLocation, 1) = "\" Then $sNewLocation = StringRight($sNewLocation, StringLen($sNewLocation) - 1)
		Return SetError(0, 0, $sNewLocation)
	EndIf
	Return SetError(1, 0, $sLocation)
EndFunc   ;==>__MD_SlidePath

;Get Item Property Value
Func _MD_GetItemProperty(ByRef $oProperties, $sObjectId, $tProperty, $pFunction)
	Local $oDeviceKeyCollection = ObjCreateInterface($sCLSID_PortableDeviceKeyCollection, $sIID_IPortableDeviceKeyCollection, $sTag_IPortableDeviceKeyCollection)
	$oDeviceKeyCollection.Add($tProperty)
	Local $pValues = 0
	$oProperties.GetValues($sObjectId, $oDeviceKeyCollection(), $pValues)
	Local $oValues = ObjCreateInterface($pValues, $sIID_IPortableDeviceValues, $sTag_IPortableDeviceValues)
	Return $pFunction($oValues, $tProperty)
EndFunc   ;==>_MD_GetItemProperty

Func SUCCEEDED($hResult)
	Return ($hResult >= 0)
EndFunc   ;==>SUCCEEDED
#EndRegion Utils
