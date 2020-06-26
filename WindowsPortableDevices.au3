#cs Copyright
   	Copyright 2020 Danysys. <hello@danysys.com>

   	Licensed under the MIT license.
   	See LICENSE file or go to https://opensource.org/licenses/MIT for details.
#ce Copyright

#cs Information
   	Author(s)......: DanysysTeam (Danyfirex & Dany3j)
   	Description....: Windows Portable Devices UDF (Partial Implementation intended to be used with MediaDevice UDF)
   	Version........: 1.0.0
   	AutoIt Version.: 3.3.14.5
#ce Information

#include-once
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 6 ;-w 5 -w 6
;~ #Tidy_Parameters=/tcb=-1 /sf /ewnl /reel /gd ;/sfc
#Region Include
#include <WinAPICom.au3>
#include <StructureConstants.au3>
#include <WinAPIConv.au3>
#include <APIFilesConstants.au3>
#EndRegion Include

#Region PROVARIANT and PROPERTYKEY
Global Const $tagPROPERTYKEY = $tagGUID & ';DWORD pid'
Global Const $tagPROPVARIANT = _
		'USHORT vt;' & _ ;typedef unsigned short VARTYPE; - in WTypes.h
		'WORD wReserved1;' & _
		'WORD wReserved2;' & _
		'WORD wReserved3;' & _
		'PTR;PTR'  ;union, use the largest member (BSTRBLOB, which is 96-bit in x64)
#EndRegion PROVARIANT and PROPERTYKEY

#Region Common HRESULT
Global Const $S_OK = 0x00000000
Global Const $E_ABORT = 0x80004004
Global Const $E_ACCESSDENIED = 0x80070005
Global Const $E_FAIL = 0x80004005
Global Const $E_HANDLE = 0x80070006
Global Const $E_INVALIDARG = 0x80070057
Global Const $E_NOINTERFACE = 0x80004002
Global Const $E_NOTIMPL = 0x80004001
Global Const $E_OUTOFMEMORY = 0x8007000E
Global Const $E_POINTER = 0x80004003
Global Const $E_UNEXPECTED = 0x8000FFFF
#EndRegion Common HRESULT

#Region WPD Error Constants
;WPD error constants.
Global Const $E_WPD_DEVICE_ALREADY_OPENED = 0x802A0001 ;The device connection has already been opened by a prior call to IPortableDevice::Open.
Global Const $E_WPD_DEVICE_IS_HUNG = 0x802A0006 ;The device will no longer respond to input.
Global Const $E_WPD_DEVICE_NOT_OPEN = 0x802A0002 ;The device connection has not yet been opened by a call to IPortableDevice::Open.
Global Const $E_WPD_OBJECT_ALREADY_ATTACHED_TO_DEVICE = 0x802A0003 ;The interface object has already been attached to the device interface.
Global Const $E_WPD_OBJECT_ALREADY_ATTACHED_TO_SERVICE = 0x802A00CA ;The interface object has already been attached to the IPortableDeviceService interface.
Global Const $E_WPD_OBJECT_NOT_ATTACHED_TO_DEVICE = 0x802A0004 ;The interface object has not been attached to the device.
Global Const $E_WPD_OBJECT_NOT_ATTACHED_TO_SERVICE = 0x802A00CB ;The interface object has not been attached to the IPortableDeviceService interface. Typically, this is returned if the application tries to access methods of an attached interface, such as IPortableDeviceServiceCapabilities, after IPortableDevice::Close is called.
Global Const $E_WPD_OBJECT_NOT_COMMITED = 0x802A0005 ;IStream::Commit was never called when creating an object with data on a device.
Global Const $E_WPD_SERVICE_ALREADY_OPENED = 0x802A00C8 ;The service connection has already been opened by a prior call to IPortableDevice::Open.
Global Const $E_WPD_SERVICE_BAD_PARAMETER_ORDER = 0x802A00CC ;The method parameters for IPortableDeviceServiceMethods::Invoke or IPortableDeviceServiceMethods::InvokeAsync are not set in the correct order. The parameter must be set in the ordering specified by WPD_PARAMETER_ATTRIBUTE_ORDER.
Global Const $E_WPD_SERVICE_NOT_OPEN = 0x802A00C9 ;The service connection has not yet been opened by a call to IPortableDeviceService::Open.
Global Const $E_WPD_SMS_INVALID_RECIPIENT = 0x802A0064 ;The recipient specified for an SMS message is invalid.
Global Const $E_WPD_SMS_INVALID_MESSAGE_BODY = 0x802A0065 ;The body of a message specified for an SMS message is invalid.
Global Const $E_WPD_SMS_SERVICE_UNAVAILABLE = 0x802A0066 ;The SMS service is unavailable.

;Windows Media Rights Manager SDK error constants.
Global Const $NS_E_DRM_DEBUGGING_NOT_ALLOWED = 0xC00D2767 ;You cannot debug when accessing DRM-protected content.
Global Const $NS_E_NOT_LICENSED = 0xC00D00CD ;The content is not licensed.

;Standard Windows error codes  commonly-used with WPD
Global Const $ERROR_ACCESS_DENIED = 0x80070005 ;May be used to indicate that a read-only object or property cannot be modified or deleted. May be used to indicate that the object is being accessed outside its scope, for example a child object that falls outside the hierarchy of a device service. May be used to indicate that the application does not have the access (for example, if access control to devices is restricted by Group Policy) to send WPD commands to the device.
Global Const $ERROR_ARITHMETIC_OVERFLOW = 0x80070216 ;May be used to indicate that the number of elements in a data array has exceeded its limits (ULONGLONG).
Global Const $ERROR_BUSY = 0x800700AA ;May be used to indicate that the device is busy processing another operation. Applications should wait for that operation to complete before retrying.
Global Const $ERROR_CANCELLED = 0x800704C7 ;A command sent to the device has been aborted due to a cancellation, e.g. by calling one of the Cancel methods in the WPD API.
Global Const $ERROR_DATATYPE_MISMATCH = 0x8007070C ;May be used to indicate that an invalid data packet was received from the device.
Global Const $ERROR_DEVICE_IN_USE = 0x80070964 ;For an MTP/IP device, indicates that the connection has failed to initialize because the device is in use.
Global Const $ERROR_DEVICE_NOT_CONNECTE = 0x8007048F ;The device has been disconnected or unplugged.
Global Const $ERROR_DIR_NOT_EMPTY = 0x80070091 ;May be used to indicate that a non-recursive delete was called for an object with children. The application should use the recursive delete flag in IPortableDeviceContent::Delete.
Global Const $ERROR_EMPTY = 0x800710D2 ;May be used to indicate that the device failed to send any resource data when resource data was expected (e.g. a thumbnail or device icon). This usually indicates an Global Const $ERROR on the device.
Global Const $ERROR_FILE_NOT_FOUND = 0x80070002 ;May be used to indicate that the device has been disconnected or unplugged.
Global Const $ERROR_GEN_FAILURE = 0x8007001F ;May be used to indicate that the device has stopped responding (hung) or a general failure has occurred on the device. The device may need to be manually reset.
Global Const $ERROR_INVALID_DATA = 0x8007000D ;May be used to indicate that data sent to or received from the device cannot be parsed correctly. This may indicate a device-side or a transport Global Const $ERROR. If MTP vendor operations are sent to the device, this Global Const $ERROR may indicate that the specified operation parameters are not of the valid VARTYPE.
Global Const $ERROR_INVALID_DATATYPE = 0x8007070C ;May be used to indicate that the specified VARTYPE is invalid for a given property.
Global Const $ERROR_INVALID_FUNCTION = 0x80070001 ;A write request was made to a resource on the device that was opened in Read mode using IPortableDeviceResources::GetStream, or a read request was made to a resource opened for Write or Create.
Global Const $ERROR_INVALID_OPERATION = 0x800710DD ;A non-recursive delete is called for an object with children.
Global Const $ERROR_INVALID_PARAMETER = 0x80070057 ;The parameter supplied by the application is not valid.
Global Const $ERROR_INVALID_TIME = 0x8007076D ;May be used to indicate that a conversion of a datetime property has failed.
Global Const $ERROR_IO_DEVICE = 0x8007045D ;May be used to indicate that the device has stopped responding (hung). The device may need to be manually reset.
Global Const $ERROR_NOT_FOUND = 0x80070490 ;May be used to indicate that the device supports a property, but that property value is currently empty or uninitialized. May be used to indicate that the internal context for a long-running operation no longer exists, as the operation has completed or has been cancelled. Examples of such operations include bulk properties, object enumeration, transfer, and invoking device service methods. Applications should retry the operation from the beginning. May be used to indicate that the specified object does not exist. The child object may be outside of the device service hierarchy.
Global Const $ERROR_NOT_READY = 0x80070015 ;May be used to indicate that an operation is not initialized correctly. This usually indicates an internal Global Const $ERROR, or that the application is using a stale device handle. The application should retry the operation from the beginning, or reopen the device.
Global Const $ERROR_NOT_SUPPORTED = 0x80070032 ;May be used to indicate that a property or command is not supported by the device.
Global Const $ERROR_OPERATION_ABORTED = 0x800703E3 ;A command sent to the device has been aborted due to a manual cancellation, e.g. by calling one of the Cancel methods in the WPD API.
Global Const $ERROR_READ_FAULT = 0x8007001E ;May be used to indicate that the device is not sending the correct amount of data.
Global Const $ERROR_RESOURCE_NOT_AVAILABLE = 0x8007138E ;May be used to indicate that a resource (such as a thumbnail or an icon) is not present on the device.
Global Const $ERROR_SEM_TIMEOUT = 0x80070079 ;May be used to indicate that the device has stopped responding (hung). The device may need to be manually reset.
Global Const $ERROR_TIMEOUT = 0x800705B4 ;May be used to indicate that the device has stopped responding (hung). The device may need to be manually reset.
Global Const $ERROR_UNSUPPORTED_TYPE = 0x8007065E ;May be used to indicate that the specified format is not supported by the device.
Global Const $ERROR_WRITE_FAULT = 0x8007001D ;May be used to indicate that the application was unable to send the requested amount of data to the device.
Global Const $WSAETIMEDOUT = 0x8007274c ;For an MTP/IP device, indicates that the connection to the device has timed out. The device may need to be manually reconnected.
#EndRegion WPD Error Constants

#Region WPD defines
Global Const $WPD_DEVICE_OBJECT_ID = "DEVICE"
Global Const $WPD_CONTROL_FUNCTION_GENERIC_MESSAGE = 0x42
Global Const $PORTABLE_DEVICE_TYPE = "PortableDeviceType"
Global Const $PORTABLE_DEVICE_ICON = "Icons"
Global Const $PORTABLE_DEVICE_NAMESPACE_TIMEOUT = "PortableDeviceNameSpaceTimeout"
Global Const $PORTABLE_DEVICE_NAMESPACE_EXCLUDE_FROM_SHELL = "PortableDeviceNameSpaceExcludeFromShell"
Global Const $PORTABLE_DEVICE_NAMESPACE_THUMBNAIL_CONTENT_TYPES = "PortableDeviceNameSpaceThumbnailContentTypes"
Global Const $PORTABLE_DEVICE_IS_MASS_STORAGE = "PortableDeviceIsMassStorage"
Global Const $PORTABLE_DEVICE_DRM_SCHEME_WMDRM10_PD = "WMDRM10-PD"
Global Const $PORTABLE_DEVICE_DRM_SCHEME_PDDRM = "PDDRM"
#EndRegion WPD defines

#Region DELETE_OBJECT_OPTIONS
Global Const $PORTABLE_DEVICE_DELETE_NO_RECURSION = 0
Global Const $PORTABLE_DEVICE_DELETE_WITH_RECURSION = 1
#EndRegion DELETE_OBJECT_OPTIONS

#Region WPD_DEVICE_TYPES
Global Const $WPD_DEVICE_TYPE_GENERIC = 0
Global Const $WPD_DEVICE_TYPE_CAMERA = 1
Global Const $WPD_DEVICE_TYPE_MEDIA_PLAYER = 2
Global Const $WPD_DEVICE_TYPE_PHONE = 3
Global Const $WPD_DEVICE_TYPE_VIDEO = 4
Global Const $WPD_DEVICE_TYPE_PERSONAL_INFORMATION_MANAGER = 5
Global Const $WPD_DEVICE_TYPE_AUDIO_RECORDER = 6
#EndRegion WPD_DEVICE_TYPES

#Region WPD_DEVICE_TRANSPORTS
Global Const $WPD_DEVICE_TRANSPORT_UNSPECIFIED = 0
Global Const $WPD_DEVICE_TRANSPORT_USB = 1
Global Const $WPD_DEVICE_TRANSPORT_IP = 2
Global Const $WPD_DEVICE_TRANSPORT_BLUETOOTH = 3
#EndRegion WPD_DEVICE_TRANSPORTS

#Region WPD_STORAGE_TYPE_VALUES
Global Const $WPD_STORAGE_TYPE_UNDEFINED = 0
Global Const $WPD_STORAGE_TYPE_FIXED_ROM = 1
Global Const $WPD_STORAGE_TYPE_REMOVABLE_ROM = 2
Global Const $WPD_STORAGE_TYPE_FIXED_RAM = 3
Global Const $WPD_STORAGE_TYPE_REMOVABLE_RAM = 4
#EndRegion WPD_STORAGE_TYPE_VALUES

#Region Event Constants
;Device Interface Types.
Global Const $GUID_DEVINTERFACE_WPD = "{6AC27878-A6FA-4155-BA85-F98F491D4F33}" ;Identifies devices that appear in a normal WPD enumeration. Any device that registers this interface GUID will be enumerated when an application calls the IPortableDeviceManager::GetDevices method.
Global Const $GUID_DEVINTERFACE_WPD_PRIVATE = "{BA0C718F-4DED-49B7-BDD3-FABE28661211}" ;Identifies devices that will not appear during a normal WPD enumeration. Any device that registers this interface GUID will be enumerated only when an application calls the IPortableDeviceManager::GetPrivateDevices method.
Global Const $GUID_DEVINTERFACE_WPD_SERVICE = "{9EF44F80-3D64-4246-A6AA-206F328D1EDC}" ;Identifies services that support the WPD Services DDI. The WPD Class Extension component enables this device interface for WPD Services that use it. Clients use this PnP interface when registering for PnP device arrival messages for ALL WPD services. To register for specific categories of services, clients should use the service category or service implements GUID.

;Event Constants
Global Const $WPD_EVENT_DEVICE_CAPABILITIES_UPDATED = "{36885AA1-CD54-4DAA-B3D0-AFB3E03F5999}" ;Indicates that the device capabilities have changed. Clients should query the device again if they have made any decisions based on device capabilities.
Global Const $WPD_EVENT_DEVICE_REMOVED = "{E4CBCA1B-6918-48B9-85EE-02BE7C850AF9}" ;Sent when a driver for a device is being unloaded. This is typically a result of the device being unplugged. Clients should release the IPortableDevice interface they have open on the device specified in WPD_EVENT_PARAMETER_PNP_DEVICE_ID.
Global Const $WPD_EVENT_DEVICE_RESET = "{7755CF53-C1ED-44F3-B5A2-451E2C376B27}" ;Indicates that the device is about to be reset, and all connected clients should close their connection to the device.
Global Const $WPD_EVENT_NOTIFICATION = "{2BA2E40A-6B4C-4295-BB43-26322B99AEB2}" ;GUID that identifies all WPD driver events to the event subsystem. The driver uses this GUID when it queues an event with the IWdfDevice::PostEvent method. Applications never use this value.
Global Const $WPD_EVENT_OBJECT_ADDED = "{A726DA95-E207-4B02-8D44-BEF2E86CBFFC}" ;Indicates that a new object is available on the device.
Global Const $WPD_EVENT_OBJECT_REMOVED = "{BE82AB88-A52C-4823-96E5-D0272671FC38}" ;Sent after a previously existing object has been removed from the device.
Global Const $WPD_EVENT_OBJECT_TRANSFER_REQUESTED = "{8D16A0A1-F2C6-41DA-8F19-5E53721ADBF2}" ;Sent to request an application to transfer a particular object from the device. The object is usually a content object, for example, an image file.
Global Const $WPD_EVENT_OBJECT_UPDATED = "{1445A759-2E01-485D-9F27-FF07DAE697AB}" ;Sent after an object has been updated, so that any connected client should refresh its view of that object.
Global Const $WPD_EVENT_STORAGE_FORMAT = "{3782616B-22BC-4474-A251-3070F8D38857}" ;Indicates the progress of a format operation on a storage object.
#EndRegion Event Constants

;This category is for properties common to all objects whose functional category is WPD_FUNCTIONAL_CATEGORY_STORAGE.
#Region WPD_STORAGE_OBJECT_PROPERTIES_V1
Global Const $WPD_STORAGE_OBJECT_PROPERTIES_V1 = __DEFINE_GUID("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A}")
;~ [ VT_UI4 ] Indicates the type of storage e.g. fixed, removable etc.
Global Const $WPD_STORAGE_TYPE = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 2")
;~ //   [ VT_LPWSTR ] Indicates the file system type e.g. "FAT32" or "NTFS" or "My Special File System"
Global Const $WPD_STORAGE_FILE_SYSTEM_TYPE = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 3")
;~ //   [ VT_UI8 ] Indicates the total storage capacity in bytes.
Global Const $WPD_STORAGE_CAPACITY = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 4")
;~ //   [ VT_UI8 ] Indicates the available space in bytes.
Global Const $WPD_STORAGE_FREE_SPACE_IN_BYTES = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 5")
;~ //   [ VT_UI8 ] Indicates the available space in objects e.g. available slots on a SIM card.
Global Const $WPD_STORAGE_FREE_SPACE_IN_OBJECTS = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 6")
;~ //   [ VT_LPWSTR ] Contains a description of the storage.
Global Const $WPD_STORAGE_DESCRIPTION = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 7")
;~ //   [ VT_LPWSTR ] Contains the serial number of the storage.
Global Const $WPD_STORAGE_SERIAL_NUMBER = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 8")
;~ //   [ VT_UI8 ] Specifies the maximum size of a single object (in bytes) that can be placed on this storage.
Global Const $WPD_STORAGE_MAX_OBJECT_SIZE = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 9")
;~ //   [ VT_UI8 ] Indicates the total storage capacity in objects e.g. available slots on a SIM card.
Global Const $WPD_STORAGE_CAPACITY_IN_OBJECTS = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 10")
;~ //   [ VT_UI4 ] This property identifies any write-protection that globally affects this storage. This takes precedence over access specified on individual objects.
Global Const $WPD_STORAGE_ACCESS_CAPABILITY = __DEFINE_PROPERTYKEY("{01A3057A-74D6-4E80-BEA7-DC4C212CE50A} 11")
#EndRegion WPD_STORAGE_OBJECT_PROPERTIES_V1

#Region  WPD content types
;WPD_CONTENT_TYPE
;Requirements for Objects
Global Const $WPD_CONTENT_TYPE_ALL = __DEFINE_GUID("{80E170D2-1055-4A3E-B952-82CC4F8A8689}") ;This content type is only valid to use in certain query methods to indicate that you are interested in all device types; you cannot create an object of this type.If you are designing a custom object, it must support these properties, at minimum.
Global Const $WPD_CONTENT_TYPE_AUDIO = __DEFINE_GUID("{4AD2C85E-5E2D-45E5-8864-4F229E3C6CF0}") ;Object is an audio file, such as a WMA or MP3 file.
Global Const $WPD_CONTENT_TYPE_AUDIO_ALBUM = __DEFINE_GUID("{AA18737E-5009-48FA-AE21-85F24383B4E6}") ;Object is an audio album.
Global Const $WPD_CONTENT_TYPE_APPOINTMENT = __DEFINE_GUID("{0FED060E-8793-4B1E-90C9-48AC389AC631}") ;Object is an appointment in a calendar.
Global Const $WPD_CONTENT_TYPE_CALENDAR = __DEFINE_GUID("{A1FD5967-6023-49A0-9DF1-F8060BE751B0}") ;Object is a calendar.
Global Const $WPD_CONTENT_TYPE_CERTIFICATE = __DEFINE_GUID("{DC3876E8-A948-4060-9050-CBD77E8A3D87}") ;Object is a certificate that is used for authentication.
Global Const $WPD_CONTENT_TYPE_CONTACT = __DEFINE_GUID("{EABA8313-4525-4707-9F0E-87C6808E9435}") ;Object is personal contact data, such as a vCard file.
Global Const $WPD_CONTENT_TYPE_CONTACT_GROUP = __DEFINE_GUID("{346B8932-4C36-40D8-9415-1828291F9DE9}") ;Object represents a group of contacts. This object’s WPD_OBJECT_REFERENCES property contains a list of object identifiers for various WPD_CONTENT_TYPE_CONTACT objects.
Global Const $WPD_CONTENT_TYPE_DOCUMENT = __DEFINE_GUID("{680ADF52-950A-4041-9B41-65E393648155}") ;Object is a container for text, with or without formatting. Examples include Microsoft Word files and plain text files.
Global Const $WPD_CONTENT_TYPE_EMAIL = __DEFINE_GUID("{8038044A-7E51-4F8F-883D-1D0623D14533}") ;Object is an E-mail message.
Global Const $WPD_CONTENT_TYPE_FOLDER = __DEFINE_GUID("{27E2E392-A111-48E0-AB0C-E17705A05F85}") ;Object is a folder.
Global Const $WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT = __DEFINE_GUID("{99ED0160-17FF-4C44-9D98-1D7A6F941921}") ;Object is a functional object that represents device functionality.
Global Const $WPD_CONTENT_TYPE_GENERIC_FILE = __DEFINE_GUID("{0085E0A6-8D34-45D7-BC5C-447E59C73D48}") ;Object is a generic, physical file that does not fall into any of the other predefined content types for files.
Global Const $WPD_CONTENT_TYPE_GENERIC_MESSAGE = __DEFINE_GUID("{E80EAAF8-B2DB-4133-B67E-1BEF4B4A6E5F}") ;An object that describes its type as WPD_CONTENT_TYPE_GENERIC_MESSAGE represents a message, for example, SMS, e-mail, and so on.
Global Const $WPD_CONTENT_TYPE_IMAGE = __DEFINE_GUID("{ef2107d5-a52a-4243-a26b-62d4176d7603}") ;Object is a still image, such as a JPEG file.
Global Const $WPD_CONTENT_TYPE_IMAGE_ALBUM = __DEFINE_GUID("{75793148-15F5-4A30-A813-54ED8A37E226}") ;Object is an image album.
Global Const $WPD_CONTENT_TYPE_MEDIA_CAST = __DEFINE_GUID("{5E88B3CC-3E65-4E62-BFFF-229495253AB0}") ;Object is a media cast object. A media cast object can represent a container object that groups related content that is published online. For example, an RSS channel can be represented as a media cast object, and this object’s WPD_OBJECT_REFERENCES property contains a list of object identifiers that represent each item in the channel.
Global Const $WPD_CONTENT_TYPE_MEMO = __DEFINE_GUID("{9CD20ECF-3B50-414F-A641-E473FFE45751}") ;Object represents memo data, for example, a text note.
Global Const $WPD_CONTENT_TYPE_MIXED_CONTENT_ALBUM = __DEFINE_GUID("{00F0C3AC-A593-49AC-9219-24ABCA5A2563}") ;Object is an album of mixed media objects—for example, audio, image, and video files.
Global Const $WPD_CONTENT_TYPE_NETWORK_ASSOCIATION = __DEFINE_GUID("{031DA7EE-18C8-4205-847E-89A11261D0F3}") ;An object that describes its type as WPD_CONTENT_TYPE_NETWORK_ASSOCIATION represents an association between a host and a device.
Global Const $WPD_CONTENT_TYPE_PLAYLIST = __DEFINE_GUID("{1A33F7E4-AF13-48F5-994E-77369DFE04A3}") ;Object is a playlist.
Global Const $WPD_CONTENT_TYPE_PROGRAM = __DEFINE_GUID("{D269F96A-247C-4BFF-98FB-97F3C49220E6}") ;Object represents a file that can be run, for example, a script or an executable.
Global Const $WPD_CONTENT_TYPE_SECTION = __DEFINE_GUID("{821089F5-1D91-4DC9-BE3C-BBB1B35B18CE}") ;Object describes a section of data that is contained in another object. For example, a large audio file might best be described by a series of chapters. Each chapter could be a WPD_CONTENT_TYPE_SECTION object with its own chapter art, metadata, and so on, and whose data is a subset of the large audio file (For example, the first chapter is the first 10 minutes, the second chapter is the next 20 minutes, and so on).
Global Const $WPD_CONTENT_TYPE_TASK = __DEFINE_GUID("{63252F2C-887F-4CB6-B1AC-D29855DCEF6C}") ;Object is a task, such as an item in a to-do list.
Global Const $WPD_CONTENT_TYPE_TELEVISION = __DEFINE_GUID("{60A169CF-F2AE-4E21-9375-9677F11C1C6E}") ;Object is a television recording.
Global Const $WPD_CONTENT_TYPE_UNSPECIFIED = __DEFINE_GUID("{28D8D31E-249C-454E-AABC-34883168E634}") ;Object is a generic object that does not fall into the predefined WPD content types.
Global Const $WPD_CONTENT_TYPE_VIDEO = __DEFINE_GUID("{9261B03C-3D78-4519-85E3-02C5E1F50BB9}") ;Object is a video, such as a WMV or AVI file.
Global Const $WPD_CONTENT_TYPE_VIDEO_ALBUM = __DEFINE_GUID("{012B0DB7-D4C1-45D6-B081-94B87779614F}") ;Object is a video album.
Global Const $WPD_CONTENT_TYPE_WIRELESS_PROFILE = __DEFINE_GUID("{0BAC070A-9F5F-4DA4-A8F6-3DE44D68FD6C}") ;Object contains wireless network access information.

;WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT Functional Object Categories
Global Const $WPD_FUNCTIONAL_CATEGORY_ALL = __DEFINE_GUID("{2D8A6512-A74C-448E-BA8A-F4AC07C49399}") ;This functional category is valid only as a parameter for certain query functions (to indicate that all functional object types are acceptable), and is not a reported functional category by the driver.
Global Const $WPD_FUNCTIONAL_CATEGORY_AUDIO_CAPTURE = __DEFINE_GUID("{3F2A1919-C7C2-4A00-855D-F57CF06DEBBB}") ;The object encapsulates audio capture functionality on the device, for example, a voice recorder or other audio recording component.
Global Const $WPD_FUNCTIONAL_CATEGORY_DEVICE = __DEFINE_GUID("{08EA466B-E3A4-4336-A1F3-A44D2B5C438C}") ;The object encapsulates the device (that is, the top-most object of the device).
Global Const $WPD_FUNCTIONAL_CATEGORY_NETWORK_CONFIGURATION = __DEFINE_GUID("{48F4DB72-7C6A-4AB0-9E1A-470E3CDBF26A}") ;The object encapsulates network-configuration functionality for the device, for example, WiFi profiles or partnerships.
Global Const $WPD_FUNCTIONAL_CATEGORY_RENDERING_INFORMATION = __DEFINE_GUID("{08600BA4-A7BA-4A01-AB0E-0065D0A356D3}") ;he object describes the types of media files that the device is able to play.
Global Const $WPD_FUNCTIONAL_CATEGORY_SMS = __DEFINE_GUID("{0044A0B1-C1E9-4AFD-B358-A62C6117C9CF}") ;The object encapsulates short message service functionality (commonly called "text messaging") on the device.
Global Const $WPD_FUNCTIONAL_CATEGORY_STILL_IMAGE_CAPTURE = __DEFINE_GUID("{613CA327-AB93-4900-B4FA-895BB5874B79}") ;The object encapsulates still image capture functionality on a device such as a camera or camera attachment.
Global Const $WPD_FUNCTIONAL_CATEGORY_STORAGE = __DEFINE_GUID("{23F05BBC-15DE-4C2A-A55B-A9AF5CE412EF}") ;The object encapsulates physical file storage on the device.
Global Const $WPD_FUNCTIONAL_CATEGORY_VIDEO_CAPTURE = __DEFINE_GUID("{E23E5F6B-7243-43AA-8DF1-0EB3D968A918}") ;The object encapsulates video capture functionality on the device, for example, a video recorder component. An application uses this object to gain programmatic control.
#EndRegion  WPD content types

#Region WPD Formats
;Object Format GUIDs
Global Const $WPD_OBJECT_FORMAT_3GP = __DEFINE_GUID("{B9840000-AE6C-4804-98BA-C57B46965FE7}") ;3GP audio or video file.
Global Const $WPD_OBJECT_FORMAT_AAC = __DEFINE_GUID("{B9030000-AE6C-4804-98BA-C57B46965FE7}") ;Audio (AAC).
Global Const $WPD_OBJECT_FORMAT_ABSTRACT_CONTACT = __DEFINE_GUID("{BA060000-AE6C-4804-98BA-C57B46965FE7}") ;Generic format for contact group objects.
Global Const $WPD_OBJECT_FORMAT_ABSTRACT_MEDIA_CAST = __DEFINE_GUID("{BA0B0000-AE6C-4804-98BA-C57B46965FE7}") ;Abstract Media Cast
Global Const $WPD_OBJECT_FORMAT_AIFF = __DEFINE_GUID("{30070000-AE6C-4804-98BA-C57B46965FE7}") ;Audio (AIFF).
Global Const $WPD_OBJECT_FORMAT_ALL = __DEFINE_GUID("{C1F62EB2-4BB3-479C-9CFA-05B5F3A57B22}") ;Identifies all available formats.
Global Const $WPD_OBJECT_FORMAT_ASF = __DEFINE_GUID("{300C0000-AE6C-4804-98BA-C57B46965FE7}") ;Video Microsoft Advanced Streaming Format (ASF).
Global Const $WPD_OBJECT_FORMAT_ASXPLAYLIST = __DEFINE_GUID("{BA130000-AE6C-4804-98BA-C57B46965FE7}") ;Playlist (ASX).
Global Const $WPD_OBJECT_FORMAT_AUDIBLE = __DEFINE_GUID("{B9040000-AE6C-4804-98BA-C57B46965FE7}") ;Audio.
Global Const $WPD_OBJECT_FORMAT_AVI = __DEFINE_GUID("{300A0000-AE6C-4804-98BA-C57B46965FE7}") ;Video (AVI).
Global Const $WPD_OBJECT_FORMAT_BMP = __DEFINE_GUID("{38040000-AE6C-4804-98BA-C57B46965FE7}") ;Image (BMP, bitmap file).
Global Const $WPD_OBJECT_FORMAT_CIFF = __DEFINE_GUID("{38050000-AE6C-4804-98BA-C57B46965FE7}") ;Image (CIFF, Canon Camera Image File Format).
Global Const $WPD_OBJECT_FORMAT_DPOF = __DEFINE_GUID("{30060000-AE6C-4804-98BA-C57B46965FE7}") ;Text (Digital Print Order File).
Global Const $WPD_OBJECT_FORMAT_EXECUTABLE = __DEFINE_GUID("{30030000-AE6C-4804-98BA-C57B46965FE7}") ;Executable.
Global Const $WPD_OBJECT_FORMAT_EXIF = __DEFINE_GUID("{38010000-AE6C-4804-98BA-C57B46965FE7}") ;Image (Exchangeable File Format).
Global Const $WPD_OBJECT_FORMAT_FLAC = __DEFINE_GUID("{B9060000-AE6C-4804-98BA-C57B46965FE7}") ;Audio (FLAC).
Global Const $WPD_OBJECT_FORMAT_FLASHPIX = __DEFINE_GUID("{38030000-AE6C-4804-98BA-C57B46965FE7}") ;Image (Structured Storage Image Format).
Global Const $WPD_OBJECT_FORMAT_GIF = __DEFINE_GUID("{38070000-AE6C-4804-98BA-C57B46965FE7}") ;Image (GIF, Graphics Interchange Format).
Global Const $WPD_OBJECT_FORMAT_HTML = __DEFINE_GUID("{30050000-AE6C-4804-98BA-C57B46965FE7}") ;HTML.
Global Const $WPD_OBJECT_FORMAT_ICON = __DEFINE_GUID("{077232ED-102C-4638-9C22-83F142BFC822}") ;Windows icon (ICO).
Global Const $WPD_OBJECT_FORMAT_JFIF = __DEFINE_GUID("{38080000-AE6C-4804-98BA-C57B46965FE7}") ;Image (JPEG Interchange Format).
Global Const $WPD_OBJECT_FORMAT_JP2 = __DEFINE_GUID("{380F0000-AE6C-4804-98BA-C57B46965FE7}") ;Image (JPEG2000 Baseline File Format).
Global Const $WPD_OBJECT_FORMAT_JPX = __DEFINE_GUID("{38100000-AE6C-4804-98BA-C57B46965FE7}") ;Image (JPEG2000 Extended File Format).
Global Const $WPD_OBJECT_FORMAT_M3UPLAYLIST = __DEFINE_GUID("{BA110000-AE6C-4804-98BA-C57B46965FE7}") ;Playlist (M3U).
Global Const $WPD_OBJECT_FORMAT_MHT_COMPILED_HTML = __DEFINE_GUID("{BA840000-AE6C-4804-98BA-C57B46965FE7}") ;MHT Compiled HTML Document file format.
Global Const $WPD_OBJECT_FORMAT_MICROSOFT_EXCEL = __DEFINE_GUID("{BA850000-AE6C-4804-98BA-C57B46965FE7}") ;Microsoft Office Excel Document file format.
Global Const $WPD_OBJECT_FORMAT_MICROSOFT_POWERPOINT = __DEFINE_GUID("{BA860000-AE6C-4804-98BA-C57B46965FE7}") ;Microsoft Office PowerPoint Document file format.
Global Const $WPD_OBJECT_FORMAT_MICROSOFT_WFC = __DEFINE_GUID("{B1040000-AE6C-4804-98BA-C57B46965FE7}") ;Windows Connect Now file format.
Global Const $WPD_OBJECT_FORMAT_MICROSOFT_WORD = __DEFINE_GUID("{BA830000-AE6C-4804-98BA-C57B46965FE7}") ;Microsoft Office Word Document file format.
Global Const $WPD_OBJECT_FORMAT_MP2 = __DEFINE_GUID("{B9830000-AE6C-4804-98BA-C57B46965FE7}") ;Audio or Video file format (MP2).
Global Const $WPD_OBJECT_FORMAT_MP3 = __DEFINE_GUID("{30090000-AE6C-4804-98BA-C57B46965FE7}") ;Audio (MP3).
Global Const $WPD_OBJECT_FORMAT_MP4 = __DEFINE_GUID("{B9820000-AE6C-4804-98BA-C57B46965FE7}") ;MPEG4 video file.
Global Const $WPD_OBJECT_FORMAT_M4A = __DEFINE_GUID("{30ABA7AC-6FFD-4C23-A359-3E9B52F3F1C8}") ;MPEG4 audio file.
Global Const $WPD_OBJECT_FORMAT_MPEG = __DEFINE_GUID("{300B0000-AE6C-4804-98BA-C57B46965FE7}") ;Video (MPEG).
Global Const $WPD_OBJECT_FORMAT_MPLPLAYLIST = __DEFINE_GUID("{BA120000-AE6C-4804-98BA-C57B46965FE7}") ;Playlist (MPL).
Global Const $WPD_OBJECT_FORMAT_NETWORK_ASSOCIATION = __DEFINE_GUID("{B1020000-AE6C-4804-98BA-C57B46965FE7}") ;Network Association file format.
Global Const $WPD_OBJECT_FORMAT_OGG = __DEFINE_GUID("{B9020000-AE6C-4804-98BA-C57B46965FE7}") ;Audio (OCG).
Global Const $WPD_OBJECT_FORMAT_PCD = __DEFINE_GUID("{38090000-AE6C-4804-98BA-C57B46965FE7}") ;Image (PhotoCD Image Pac).
Global Const $WPD_OBJECT_FORMAT_PICT = __DEFINE_GUID("{380A0000-AE6C-4804-98BA-C57B46965FE7}") ;Image (Apple QuickDraw Image Format).
Global Const $WPD_OBJECT_FORMAT_PLSPLAYLIST = __DEFINE_GUID("{BA140000-AE6C-4804-98BA-C57B46965FE7}") ;Playlist (PLS).
Global Const $WPD_OBJECT_FORMAT_PNG = __DEFINE_GUID("{380B0000-AE6C-4804-98BA-C57B46965FE7}") ;Image (Portable Network Graphics).
Global Const $WPD_OBJECT_FORMAT_PROPERTIES_ONLY = __DEFINE_GUID("{30010000-AE6C-4804-98BA-C57B46965FE7}") ;This object has no data stream and is completely specified by properties.
Global Const $WPD_OBJECT_FORMAT_SCRIPT = __DEFINE_GUID("{30020000-AE6C-4804-98BA-C57B46965FE7}") ;Script (device specific format)
Global Const $WPD_OBJECT_FORMAT_TEXT = __DEFINE_GUID("{30040000-AE6C-4804-98BA-C57B46965FE7}") ;Text.
Global Const $WPD_OBJECT_FORMAT_TIFF = __DEFINE_GUID("{380D0000-AE6C-4804-98BA-C57B46965FE7}") ;Image (Tag Image File Format).
Global Const $WPD_OBJECT_FORMAT_TIFFEP = __DEFINE_GUID("{38020000-AE6C-4804-98BA-C57B46965FE7}") ;Image (Tag Image File Format for Electronic Photography).
Global Const $WPD_OBJECT_FORMAT_TIFFIT = __DEFINE_GUID("{380E0000-AE6C-4804-98BA-C57B46965FE7}") ;	Image (Tag Image File Format for Informational Technology).
Global Const $WPD_OBJECT_FORMAT_UNSPECIFIED = __DEFINE_GUID("{30000000-AE6C-4804-98BA-C57B46965FE7}") ;An undefined or unspecified object format on the device. This is used for objects that cannot be specified by the other defined WPD format codes.
Global Const $WPD_OBJECT_FORMAT_VCALENDAR1 = __DEFINE_GUID("{BE020000-AE6C-4804-98BA-C57B46965FE7}") ;vCalendar file format (vCalendar Version 1).
Global Const $WPD_OBJECT_FORMAT_ICALENDAR = __DEFINE_GUID("{BE030000-AE6C-4804-98BA-C57B46965FE7}") ;Icalendar
Global Const $WPD_OBJECT_FORMAT_VCALENDAR2 = __DEFINE_GUID("{icstatic-read-only-Guid-WPD_OBJECT_F}") ;ICALENDAR file format (vCalendar Version 2).
Global Const $WPD_OBJECT_FORMAT_VCARD2 = __DEFINE_GUID("{BB820000-AE6C-4804-98BA-C57B46965FE7}") ;vCard file format (vCard Version 2).
Global Const $WPD_OBJECT_FORMAT_VCARD3 = __DEFINE_GUID("{BB830000-AE6C-4804-98BA-C57B46965FE7}") ;vCard file format (vCard Version 3).
Global Const $WPD_OBJECT_FORMAT_WAVE = __DEFINE_GUID("{30080000-AE6C-4804-98BA-C57B46965FE7}") ;Audio file (WAV).
Global Const $WPD_OBJECT_FORMAT_WINDOWSIMAGEFORMAT = __DEFINE_GUID("{B8810000-AE6C-4804-98BA-C57B46965FE7}") ;Image.
Global Const $WPD_OBJECT_FORMAT_WMA = __DEFINE_GUID("{B9010000-AE6C-4804-98BA-C57B46965FE7}") ;Audio (WMA).
Global Const $WPD_OBJECT_FORMAT_WMV = __DEFINE_GUID("{B9810000-AE6C-4804-98BA-C57B46965FE7}") ;Video (WMV).
Global Const $WPD_OBJECT_FORMAT_WPLPLAYLIST = __DEFINE_GUID("{BA100000-AE6C-4804-98BA-C57B46965FE7}") ;Playlist (WPL).
Global Const $WPD_OBJECT_FORMAT_X509V3CERTIFICATE = __DEFINE_GUID("{B1030000-AE6C-4804-98BA-C57B46965FE7}") ;X.509 V3 Certificate file format.
Global Const $WPD_OBJECT_FORMAT_XML = __DEFINE_GUID("{BA820000-AE6C-4804-98BA-C57B46965FE7}") ;XML file format.
#EndRegion WPD Formats

#Region WPD_OBJECT_PROPERTIES_V1
;~ DEFINE_GUID( WPD_OBJECT_PROPERTIES_V1 , 0xEF6B490D, 0x5CD8, 0x437A, 0xAF, 0xFC, 0xDA, 0x8B, 0x60, 0xEE, 0x4A, 0x3C );
;~ //   [ VT_LPWSTR ] Uniquely identifies object on the Portable Device.
Global Const $WPD_OBJECT_ID = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 2")
;~ //   [ VT_LPWSTR ] Object identifier indicating the parent object.
Global Const $WPD_OBJECT_PARENT_ID = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 3")
;~ //   [ VT_BOOL ] Indicates whether the object should be hidden.
Global Const $WPD_OBJECT_ISHIDDEN = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 9")
;~ //   [ VT_BOOL ] Indicates whether the object represents system data.
Global Const $WPD_OBJECT_ISSYSTEM = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 10")
;~ //   [ VT_UI8 ] The size of the object data.
Global Const $WPD_OBJECT_SIZE = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 11")
;~ //   [ VT_BOOL ] Indicates whether the media data is DRM protected.
Global Const $WPD_OBJECT_IS_DRM_PROTECTED = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 17")
;~ //   [ VT_DATE ] Indicates the date and time the object was created on the device.
Global Const $WPD_OBJECT_DATE_CREATED = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 18")
;~ //   [ VT_DATE ] Indicates the date and time the object was modified on the device.
Global Const $WPD_OBJECT_DATE_MODIFIED = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 19")
;~ //   [ VT_DATE ] Indicates the date and time the object was authored (e.g. for music, this would be the date the music was recorded).
Global Const $WPD_OBJECT_DATE_AUTHORED = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 20")
;~ //   [ VT_LPWSTR ] Uniquely identifies the object on the Portable Device, similar to WPD_OBJECT_ID, but this ID will not change between sessions.
Global Const $WPD_OBJECT_PERSISTENT_UNIQUE_ID = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 5")
;~ //   [ VT_LPWSTR ] Contains the name of the file this object represents.
Global Const $WPD_OBJECT_ORIGINAL_FILE_NAME = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 12")
;~ //   [ VT_LPWSTR ] Indicates the Object ID of the closest functional object ancestor. For example, objects that represent files/folders under a Storage functional object, will have this property set to the object ID of the storage functional object.
Global Const $WPD_OBJECT_CONTAINER_FUNCTIONAL_OBJECT_ID = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 23")
;~ //   [ VT_LPWSTR ] If this object appears as a hint location, this property indicates the hint-specific name to display instead of the object name.
Global Const $WPD_OBJECT_HINT_LOCATION_DISPLAY_NAME = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 25")
;~ //   [ VT_BOOL ] Indicates whether the object can be deleted, or not.
Global Const $WPD_OBJECT_CAN_DELETE = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 26")
#EndRegion WPD_OBJECT_PROPERTIES_V1

#Region WPD_PROPERTYKEY
Global Const $WPD_CLIENT_NAME = __DEFINE_PROPERTYKEY("{204d9f0c-2292-4080-9f42-40664e70f859} 2")
Global Const $WPD_CLIENT_MAJOR_VERSION = __DEFINE_PROPERTYKEY("{204d9f0c-2292-4080-9f42-40664e70f859} 3")
Global Const $WPD_CLIENT_MINOR_VERSION = __DEFINE_PROPERTYKEY("{204d9f0c-2292-4080-9f42-40664e70f859} 4")
Global Const $WPD_CLIENT_REVISION = __DEFINE_PROPERTYKEY("{204d9f0c-2292-4080-9f42-40664e70f859} 5")
Global Const $WPD_CLIENT_SECURITY_QUALITY_OF_SERVICE = __DEFINE_PROPERTYKEY("{204d9f0c-2292-4080-9f42-40664e70f859} 8")
Global Const $WPD_CLIENT_DESIRED_ACCESS = __DEFINE_PROPERTYKEY("{204d9f0c-2292-4080-9f42-40664e70f859} 9")
Global Const $WPD_OBJECT_CONTENT_TYPE = __DEFINE_PROPERTYKEY("{ef6b490d-5cd8-437a-affc-da8b60ee4a3c} 7")
Global Const $WPD_OBJECT_NAME = __DEFINE_PROPERTYKEY("{EF6B490D-5CD8-437A-AFFC-DA8B60EE4A3C} 4")
#EndRegion WPD_PROPERTYKEY

#Region  Resource keys
;~ //   Represents the entire object's data. There can be only one default resource on an object.
Global Const $WPD_RESOURCE_DEFAULT = __DEFINE_PROPERTYKEY("{E81E79BE-34F0-41BF-B53F-F1A06AE87842} 0")
#EndRegion  Resource keys

Global $sTag_STATSTG = "ptr pwcsName;dword type;UINT64 cbSize;" & $tagFILETIME & ";" & $tagFILETIME & ";" & $tagFILETIME & ";dword grfMode;" & _
		"dword grfLocksSupported;" & $tagGUID & ";dword grfStateBits;dword reserved"

#Region STATFLAG
Global Enum $TATFLAG_DEFAULT = 0, _
		$STATFLAG_NONAME = 1, _
		$STATFLAG_NOOPEN = 2
#EndRegion STATFLAG

#Region STGC
Global Const $STGC_DEFAULT = 0
Global Const $STGC_OVERWRITE = 1
Global Const $STGC_ONLYIFCURRENT = 2
Global Const $STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE = 4
Global Const $STGC_CONSOLIDATE = 8
#EndRegion STGC

#Region Interfaces
Global Const $sCLSID_PortableDeviceManager = "{0af10cec-2ecd-4b92-9581-34f6ae0637f3}"
Global Const $sIID_IPortableDeviceManager = "{a1567595-4c2f-4574-a6fa-ecef917b9a40}"
Global Const $sTag_IPortableDeviceManager = "" & _
		"GetDevices hresult(ptr;dword*);" & _
		"RefreshDeviceList hersult();" & _
		"GetDeviceFriendlyName hresult(wstr;wstr;dword*);" & _
		"GetDeviceDescription hresult(wstr;wstr;dword*);" & _
		"GetDeviceManufacturer hresult(wstr;wstr;dword*);" & _
		"GetDeviceProperty hresult(wstr;wstr;ptr;dword*;dword*);" & _
		"GetPrivateDevices hresult(wstr;dword*)"

Global Const $sCLSID_PortableDevice = "{728a21c5-3d9e-48d7-9810-864848f0f404}"
Global Const $sIID_IPortableDevice = "{625e2df8-6392-4cf0-9ad1-3cfa5f17775c}"
Global Const $sTag_IPortableDevice = "" & _
		"Open hresult(wstr;ptr);" & _
		"SendCommand hresult(dword;ptr;ptr*);" & _
		"Content hresult(ptr*);" & _
		"Capabilities hresult(ptr*);" & _
		"Cancel hresult();" & _
		"Close hresult();" & _
		"Advise hresult(dword;ptr;ptr;wstr*);" & _
		"Unadvise hresult(wstr*);" & _
		"GetPnPDeviceID hresult(wstr*);"

Global Const $sCLSID_PortableDeviceValues = "{0c15d503-d017-47ce-9016-7b3f978721cc}"
Global Const $sIID_IPortableDeviceValues = "{6848f6f2-3155-4f86-b6f5-263eeeab3143}"
Global Const $sTag_IPortableDeviceValues = "" & _
		"GetCount hresult(dword*);" & _
		"GetAt hresult(dword;struct*;struct*);" & _
		"SetValue hresult(struct*;ptr);" & _
		"GetValue hresult(struct*;ptr*);" & _
		"SetStringValue hresult(struct*;wstr);" & _
		"GetStringValue hresult(struct*;wstr*);" & _
		"SetUnsignedIntegerValue hresult(struct*;ulong);" & _
		"GetUnsignedIntegerValue hresult(struct*;ulong*);" & _
		"SetSignedIntegerValue hresult(struct*;long);" & _
		"GetSignedIntegerValue hresult(struct*;long*);" & _
		"SetUnsignedLargeIntegerValue hresult(struct*;UINT64);" & _
		"GetUnsignedLargeIntegerValue hresult(struct*;UINT64*);" & _
		"SetSignedLargeIntegerValue hresult(struct*;INT64);" & _
		"GetSignedLargeIntegerValue hresult(struct*;INT64*);" & _
		"SetFloatValue hresult(struct*;float);" & _
		"GetFloatValue hresult(struct*;float*);" & _
		"SetErrorValue hresult(struct*;long);" & _
		"GetErrorValue hresult(struct*;long*);" & _
		"SetKeyValue hresult(struct*;struct*);" & _
		"GetKeyValue hresult(struct*;struct*);" & _
		"SetBoolValue hresult(struct*;bool);" & _
		"GetBoolValue hresult(struct*;bool*);" & _
		"SetIUnknownValue hresult(struct*;ptr);" & _
		"GetIUnknownValue hresult(struct*;ptr*);" & _
		"SetGuidValue hresult(struct*;struct*);" & _
		"GetGuidValue hresult(struct*;ptr*);" & _
		"SetBufferValue hresult(struct*;ptr);" & _
		"GetBufferValue hresult(struct*;ptr*);" & _
		"SetIPortableDeviceValuesValue hresult(struct*;ptr);" & _
		"SetIPortableDeviceValuesValue hresult(struct*;ptr*);" & _
		"SetIPortableDevicePropVariantCollectionValue hresult(struct*;ptr);" & _
		"GetIPortableDevicePropVariantCollectionValue hresult(struct*;ptr*);" & _
		"SetIPortableDeviceKeyCollectionValue hresult(struct*;ptr);" & _
		"GetIPortableDeviceKeyCollectionValue hresult(struct*;ptr*);" & _
		"SetIPortableDeviceValuesCollectionValue hresult(struct*;ptr);" & _
		"GetIPortableDeviceValuesCollectionValue hresult(struct*;ptr*);" & _
		"RemoveValue hresult(struct*);" & _
		"CopyValuesFromPropertyStore hresult(ptr);" & _
		"CopyValuesToPropertyStore hresult(ptr);" & _
		"Clear hresult();"

Global Const $sIID_IPortableDeviceContent = "{6a96ed84-7c73-4480-9938-bf5af477d426}"
Global Const $sTag_IPortableDeviceContent = "" & _
		"EnumObjects hresult(dword;wstr;ptr;ptr*);" & _
		"Properties hresult(ptr*);" & _
		"Transfer hresult(ptr*);" & _
		"CreateObjectWithPropertiesOnly hresult(ptr;wstr*);" & _
		"CreateObjectWithPropertiesAndData hresult(ptr;ptr*;dword*;ptr);" & _
		"Delete hresult(dword;ptr;ptr*);" & _
		"Cancel hresult();" & _
		"Move hresult(ptr;wstr;ptr);" & _
		"Copy hresult(ptr;wstr;ptr);"

Global Const $sIID_IPortableDeviceProperties = "{7f6d695c-03df-4439-a809-59266beee3a6}"
Global Const $sTag_IPortableDeviceProperties = "" & _
		"GetSupportedProperties hresult(wstr;ptr*);" & _
		"GetPropertyAttributes hresult(wstr;struct*;ptr*);" & _
		"GetValues hresult(wstr;ptr;ptr*);" & _
		"SetValues hresult(wstr;ptr;ptr*);" & _
		"Delete hresult(wstr;ptr*);" & _
		"Cancel hresult();"

Global Const $sIID_IPortableDevicePropertiesBulk = "{482b05c0-4056-44ed-9e0f-5e23b009da93}"
Global Const $sTag_IPortableDevicePropertiesBulk = "" & _
		"QueueGetValuesByObjectList hresult(ptr;ptr;ptr;ptr*);" & _
		"QueueGetValuesByObjectFormat hresult(struct*;wstr;dword;ptr;ptr;ptr*);" & _
		"Start hresult();" & _
		"Cancel hresult();"

Global Const $sIID_IEnumPortableDeviceObjectIDs = "{10ece955-cf41-4728-bfa0-41eedf1bbf19}"
Global Const $sTag_IEnumPortableDeviceObjectIDs = "" & _
		"Next hresult(ulong;wstr*;ulong*);" & _
		"Skip hresult(ulong);" & _
		"Clone hresult(ptr*);" & _
		"Cancel hresult();"

Global Const $sCLSID_PortableDeviceKeyCollection = "{de2d022d-2480-43be-97f0-d1fa2cf98f4f}"
Global Const $sIID_IPortableDeviceKeyCollection = "{dada2357-e0ad-492e-98db-dd61c53ba353}"
Global Const $sTag_IPortableDeviceKeyCollection = "" & _
		"GetCount hresult(dword);" & _
		"GetAt hresult(dword;struct*);" & _
		"Add hresult(struct*);" & _
		"Clear hresult();" & _
		"RemoveAt hresult(dword);"

Global Const $sIID_IPortableDevicePropertiesBulkCallback = "{9deacb80-11e8-40e3-a9f3-f557986a7845}"

Global Const $sCLSID_IPortableDevicePropVariantCollection = "{08a99e2f-6d6d-4b80-af5a-baf2bcbe4cb9}"
Global Const $sIID_IPortableDevicePropVariantCollection = "{89b2e422-4f1b-4316-bcef-a44afea83eb3}"
Global Const $sTag_IPortableDevicePropVariantCollection = "" & _
		"GetCount hresult(dword*);" & _
		"GetAt hresult(dword;struct*);" & _
		"Add hresult(struct*);" & _
		"GetType hresult(ptr);" & _
		"ChangeType hresult(ptr);" & _
		"Clear hresult();" & _
		"RemoveAt hresult(dword);"

Global Const $sIID_IPortableDeviceCapabilities = "{2c8c6dbf-e3dc-4061-becc-8542e810d126}"
Global Const $sTag_IPortableDeviceCapabilities = "" & _
		"GetSupportedCommands hresult(ptr*);" & _
		"GetCommandOptions hresult(struct*;ptr*);" & _
		"GetFunctionalCategories hresult(ptr*);" & _
		"GetFunctionalObjects hresult(struct*;ptr*);" & _
		"GetSupportedContentTypes hresult(struct*;ptr*);" & _
		"GetSupportedFormats hresult(struct*;ptr*);" & _
		"GetSupportedFormatProperties hresult(struct*;ptr*);" & _
		"GetFixedPropertyAttributes hresult(struct*;struct*;ptr*);" & _
		"Cancel hresult();" & _
		"GetSupportedEvents hresult(ptr*);" & _
		"GetEventOptions hresult(struct;ptr*);"

Global Const $sIID_IStream = "{0000000c-0000-0000-C000-000000000046}"
Global Const $sTag_IStream = "" & _
		"Read hresult(struct*;ULONG;ULONG*);" & _ ;IID_ISequentialStream 0c733a30-2a1c-11ce-ade5-00aa0044773d
		"Write hresult(struct*;ULONG;ULONG*);" & _ ;IID_ISequentialStream 0c733a30-2a1c-11ce-ade5-00aa0044773d
		"Seek hresult(INT64;dword);" & _
		"SetSize hresult(UINT64);" & _
		"CopyTo hresult(ptr;UINT64;UINT64*;UINT64*);" & _
		"Commit hresult(dword);" & _
		"Revert hresult();" & _
		"LockRegion hresult(UINT64;UINT64;dword);" & _
		"UnlockRegion hresult(UINT64;UINT64;dword);" & _
		"Stat hresult(struct*;dword);" & _
		"Clone hresult(ptr*);"

Global Const $sIID_IPortableDeviceDataStream = "{88e04db3-1012-4d64-9996-f703a950d3f4}"
Global Const $sTag_IPortableDeviceDataStream = $sTag_IStream & "" & _
		"GetObjectID hresult(wstr*);" & _
		"Cancel hresult();"

Global Const $sIID_IPortableDeviceResources = "{fd8878ac-d841-4d17-891c-e6829cdb6934}"
Global Const $sTag_IPortableDeviceResources = "" & _
		"GetSupportedResources hresult(wstr;ptr*);" & _
		"GetResourceAttributes hresult(wstr;struct*;ptr*);" & _
		"GetStream hresult(wstr;struct*;dword;dword*;ptr*);" & _
		"Delete hresult(wstr;ptr);" & _
		"Cancel hresult();" & _
		"CreateResource hresult(ptr;ptr*;dword*;wstr*);"

#EndRegion Interfaces

;Client Information
Global Const $CLIENT_NAME = "WPD Sample Application"
Global Const $CLIENT_MAJOR_VER = 1
Global Const $CLIENT_MINOR_VER = 0
Global Const $CLIENT_REVISION = 2

#Region Functions
Func __DEFINE_GUID($sGUID)
	Local $tGUID = _WinAPI_GUIDFromString($sGUID)
	Return $tGUID
EndFunc   ;==>__DEFINE_GUID

Func __DEFINE_PROPERTYKEY($sPKEY_Title)
	Local $tProperty = DllStructCreate($tagPROPERTYKEY)
	Local $aCall = DllCall("Propsys.dll", "long", "PSPropertyKeyFromString", "wstr", $sPKEY_Title, "struct*", $tProperty)
	Return $tProperty
EndFunc   ;==>__DEFINE_PROPERTYKEY

Func __WPD_CreateObjectIDString($sString)
	Local $tString = DllStructCreate("wchar Data[" & StringLen($sString) + 1 & "]")
	DllStructSetData($tString, 1, $sString)
	Return $tString
EndFunc   ;==>__WPD_CreateObjectIDString

Func __WPD_GetString(ByRef $tProVariant)
	If $tProVariant.vt = 31 Then ;31=VT_LPWSTR
		Local $tString = DllStructCreate('wchar Data[512]', DllStructGetData($tProVariant, 5))
		Return $tString.Data
	EndIf
	Return ""
EndFunc   ;==>__WPD_GetString

Func __WPD_PropertyCanonicalString(ByRef $tPropertyKey)
	Local $aCall = DllCall("Propsys.dll", "long", "PSGetNameFromPropertyKey", "struct*", $tPropertyKey, "ptr*", 0)
	If $aCall[0] = $S_OK Then
		Local $tString = DllStructCreate("wchar Data[512]", $aCall[2])
		Return $tString.Data
	EndIf
	Return ""
EndFunc   ;==>__WPD_PropertyCanonicalString

Func __WPD_PropertyString(ByRef $tPropertyKey)
	Local $tString = DllStructCreate("wchar Data[50]")
	Local $aCall = DllCall("Propsys.dll", "long", "PSStringFromPropertyKey", "struct*", $tPropertyKey, "ptr", DllStructGetPtr($tString), "uint", 50) ;PKEYSTR_MAX=50
	Return $tString.Data
EndFunc   ;==>__WPD_PropertyString

Func __WPDCompareGUIDStringStruct($sGUID1, $tGUID2)
	Local $sGUID2 = _WinAPI_StringFromGUID($tGUID2)
	Return $sGUID1 = $sGUID2
EndFunc   ;==>__WPDCompareGUIDStringStruct

Func __WPDGetDeviceDescription($oInterface, $sPnPDeviceID)
	Local $sString = ""
	$oInterface.GetDeviceDescription($sPnPDeviceID, $sString, 512)
	Return $sString
EndFunc   ;==>__WPDGetDeviceDescription

Func __WPDGetDeviceManufacturer($oInterface, $sPnPDeviceID)
	Local $sString = ""
	$oInterface.GetDeviceManufacturer($sPnPDeviceID, $sString, 512)
	Return $sString
EndFunc   ;==>__WPDGetDeviceManufacturer

Func __WPDGetFriendlyName($oInterface, $sPnPDeviceID)
	Local $sString = ""
	$oInterface.GetDeviceFriendlyName($sPnPDeviceID, $sString, 512)
	Return $sString
EndFunc   ;==>__WPDGetFriendlyName

Func __WPDGetGetDeviceProperty(ByRef $oDeviceManager, $sPnPDeviceID, $sDevicePropertyName, $sStructCast = "dword")
	Local $iSize = 0
	Local $iType = 0
	Local $hResult = $oDeviceManager.GetDeviceProperty($sPnPDeviceID, $sDevicePropertyName, 0, $iSize, $iType)
	If $hResult = $S_OK And $iSize Then
		Local $tBuffer = DllStructCreate("byte[" & $iSize & "]")
		Local $pBuffer = DllStructGetPtr($tBuffer, 1)
		Local $tCast = DllStructCreate($sStructCast, $pBuffer)
		$hResult = $oDeviceManager.GetDeviceProperty($sPnPDeviceID, $sDevicePropertyName, $pBuffer, $iSize, $iType)
		If $hResult = $S_OK Then
			Return DllStructGetData($tCast, 1)
		EndIf
	EndIf
	Return SetError(1, $hResult, "")
EndFunc   ;==>__WPDGetGetDeviceProperty

Func __WPDGetGetDeviceType(ByRef $oDeviceManager, $sPnPDeviceID)
	Local $iType = __WPDGetGetDeviceProperty($oDeviceManager, $sPnPDeviceID, "PortableDeviceType")
	Local $sType = "Unkown"
	If @error Then Return $sType
	Switch $iType
		Case $WPD_DEVICE_TYPE_GENERIC
			$sType = "Generic"
		Case $WPD_DEVICE_TYPE_CAMERA
			$sType = "Camera"
		Case $WPD_DEVICE_TYPE_MEDIA_PLAYER
			$sType = "Media Player"
		Case $WPD_DEVICE_TYPE_PHONE
			$sType = "Phone"
		Case $WPD_DEVICE_TYPE_VIDEO
			$sType = "Video"
		Case $WPD_DEVICE_TYPE_PERSONAL_INFORMATION_MANAGER
			$sType = "Personal Information Manager"
		Case $WPD_DEVICE_TYPE_AUDIO_RECORDER
			$sType = "Audio Recorder"
		Case Else
			$sType = ""
	EndSwitch
	Return $sType
EndFunc   ;==>__WPDGetGetDeviceType

Func _ListRecursive($sObjectId, ByRef $oContent)
	Local $pProperties = 0
	Local $pDeviceObjectIDs = 0
	Local $hResult = $oContent.Properties($pProperties)
	Local $oProperties = ObjCreateInterface($pProperties, $sIID_IPortableDeviceProperties, $sTag_IPortableDeviceProperties)
	$hResult = $oContent.EnumObjects(0, $sObjectId, 0, $pDeviceObjectIDs)
;~ 	_Debug("$hResult: " & $hResult)
;~ 	_Debug("$pDeviceObjectIDs: " & $pDeviceObjectIDs)
;~ 	_Debug("$pProperties: " & $pProperties)

	Local $oDeviceObjectIDs = ObjCreateInterface($pDeviceObjectIDs, $sIID_IEnumPortableDeviceObjectIDs, $sTag_IEnumPortableDeviceObjectIDs)
;~ 	__WDP_DebugInterface("IEnumPortableDeviceObjectIDs", $oDeviceObjectIDs)

	Local $sID = ""
	While $oDeviceObjectIDs.Next(1, $sID, 0) = $S_OK
;~ 		_ObjectPrintProperties($sID, $oProperties)
		_ListRecursive($sID, $oContent)
	WEnd
EndFunc   ;==>_ListRecursive

Func _MPD_DeviceCapabilities(ByRef $oDevice)
	Local $pObject = 0
	$oDevice.Capabilities($pObject)
	Return ObjCreateInterface($pObject, $sIID_IPortableDeviceCapabilities, $sTag_IPortableDeviceCapabilities)
EndFunc   ;==>_MPD_DeviceCapabilities

Func _MPD_DeviceContent(ByRef $oDevice)
	Local $pObject = 0
	$oDevice.Content($pObject)
	Return ObjCreateInterface($pObject, $sIID_IPortableDeviceContent, $sTag_IPortableDeviceContent)
EndFunc   ;==>_MPD_DeviceContent

Func _MPD_DeviceProperties(ByRef $oContent)
	Local $pObject = 0
	$oContent.Properties($pObject)
	Return ObjCreateInterface($pObject, $sIID_IPortableDeviceProperties, $sTag_IPortableDeviceProperties)
EndFunc   ;==>_MPD_DeviceProperties

Func _WPD_CreateDevice()
	Return ObjCreateInterface($sCLSID_PortableDevice, $sIID_IPortableDevice, $sTag_IPortableDevice)
EndFunc   ;==>_WPD_CreateDevice

Func _WPD_DevicePropertiesValues(ByRef $oDeviceProperties, $sObjectId, $pDeviceKeyCollection = 0)
	Local $pValues = 0
	$oDeviceProperties.GetValues($sObjectId, $pDeviceKeyCollection, $pValues)
	Return ObjCreateInterface($pValues, $sIID_IPortableDeviceValues, $sTag_IPortableDeviceValues)
EndFunc   ;==>_WPD_DevicePropertiesValues

Func _WPD_GUID($sGUID = "{00000000-0000-0000-0000-000000000000}")
	Return __DEFINE_GUID($sGUID)
EndFunc   ;==>_WPD_GUID

Func _WPD_GUIDIsContainer($sGUID)
	Return __WPDCompareGUIDStringStruct($sGUID, $WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT)
EndFunc   ;==>_WPD_GUIDIsContainer

Func _WPD_GUIDIsFolder($sGUID)
	Return __WPDCompareGUIDStringStruct($sGUID, $WPD_CONTENT_TYPE_FOLDER)
EndFunc   ;==>_WPD_GUIDIsFolder

Func _WPD_GUIDIsFunctionalObject($sGUID)
	Return __WPDCompareGUIDStringStruct($sGUID, $WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT)
EndFunc   ;==>_WPD_GUIDIsFunctionalObject

Func _WPD_PropertyKey($sPKEY_Title = "{00000000-0000-0000-0000-000000000000} 0")
	Return __DEFINE_PROPERTYKEY($sPKEY_Title)
EndFunc   ;==>_WPD_PropertyKey

Func _WPD_PropVariant()
	Return DllStructCreate($tagPROPVARIANT)
EndFunc   ;==>_WPD_PropVariant

Func _WPD_ValueGetBool(ByRef $oValues, $tPropertyKey)
	Local $bValue = ""
	Local $hResult = $oValues.GetBoolValue($tPropertyKey, $bValue)
	If $hResult <> $S_OK Then Return SetError(1, $hResult, "")
	Return $bValue
EndFunc   ;==>_WPD_ValueGetBool

Func _WPD_ValueGetString(ByRef $oValues, $tPropertyKey)
	Local $sValue = ""
	Local $hResult = $oValues.GetStringValue($tPropertyKey, $sValue)
	If $hResult <> $S_OK Then Return SetError(1, $hResult, "")
	Return $sValue
EndFunc   ;==>_WPD_ValueGetString

Func _WPD_ValueGetUnsignedInteger(ByRef $oValues, $tPropertyKey)
	Local $iValue = ""
	Local $hResult = $oValues.GetUnsignedIntegerValue($tPropertyKey, $iValue)
	If $hResult <> $S_OK Then Return SetError(1, $hResult, "")
	Return $iValue
EndFunc   ;==>_WPD_ValueGetUnsignedInteger

Func _WPD_ValueGetUnsignedLargeInteger(ByRef $oValues, $tPropertyKey)
	Local $fValue = ""
	Local $hResult = $oValues.GetUnsignedLargeIntegerValue($tPropertyKey, $fValue)
	If $hResult <> $S_OK Then Return SetError(1, $hResult, "")
	Return $fValue
EndFunc   ;==>_WPD_ValueGetUnsignedLargeInteger

Func _WPD_WriteObjectValues(ByRef $oValues)
	Return
	Local $iCount = 0
	$oValues.GetCount($iCount)
	Local $tPropertyKey = 0
	Local $tProVariant = 0

	For $i = 0 To $iCount - 1
		$tPropertyKey = _WPD_PropertyKey()
		$tProVariant = _WPD_PropVariant()
		$oValues.GetAt($i, $tPropertyKey, $tProVariant)
	Next
EndFunc   ;==>_WPD_WriteObjectValues

; #FUNCTION# ====================================================================================================================
; Name ..........: _WPDListDevices
; Description ...: Return a List of Windows Portables Devices
; Syntax ........: _WPDListDevices()
; Parameters ....: None
; Return values .: Success      - 2D Array  [n][4] ==> [n][0]PnPDeviceID, [n][1]DeviceFriendlyName, [n][2]DeviceManufacturer,[n][3]DeviceDescription ,[n][4]DeviceType
;                  Failure      - 0
; Author ........: Danyfirex
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _WPDListDevices()
	Local $aDevices[0][5]
	Local $iNumberOfDevices = 0
	Local $hResult = 0
	Local $taPnPDevicesID = 0
	Local $tPnPDeviceName = 0
	Local $oDeviceManager = ObjCreateInterface($sCLSID_PortableDeviceManager, $sIID_IPortableDeviceManager, $sTag_IPortableDeviceManager)
	If Not IsObj($oDeviceManager) Then Return SetError(1, 0, 0)

	$hResult = $oDeviceManager.RefreshDeviceList()
	$hResult = $oDeviceManager.GetDevices(Null, $iNumberOfDevices)
	If Not $iNumberOfDevices Then Return SetError(2, 0, 0)

	$taPnPDevicesID = DllStructCreate("ptr[" & $iNumberOfDevices & "]")
	ReDim $aDevices[$iNumberOfDevices][5]
	$hResult = $oDeviceManager.GetDevices(DllStructGetPtr($taPnPDevicesID), $iNumberOfDevices)
	If Not $hResult == $S_OK Or Not $iNumberOfDevices Then Return SetError(3, 0, 0)

	For $i = 0 To $iNumberOfDevices - 1
		$tPnPDeviceName = DllStructCreate("wchar[512]", DllStructGetData($taPnPDevicesID, 1, $i + 1))
		$aDevices[$i][0] = DllStructGetData($tPnPDeviceName, 1)
		$aDevices[$i][1] = StringStripWS(__WPDGetFriendlyName($oDeviceManager, $aDevices[$i][0]), 3)
		$aDevices[$i][2] = StringStripWS(__WPDGetDeviceManufacturer($oDeviceManager, $aDevices[$i][0]), 3)
		$aDevices[$i][3] = StringStripWS(__WPDGetDeviceDescription($oDeviceManager, $aDevices[$i][0]), 3)
		$aDevices[$i][4] = __WPDGetGetDeviceType($oDeviceManager, $aDevices[$i][0]) ;PORTABLE_DEVICE_TYPE
		_WinAPI_CoTaskMemFree(DllStructGetData($taPnPDevicesID, 1, $i + 1))
		$tPnPDeviceName = 0
	Next

	Return $aDevices
EndFunc   ;==>_WPDListDevices

Func _WPDPropVariantFromString($sString)
	Local $tPropVariant = _WPD_PropVariant()
	Local $tString = DllStructCreate('wchar Data[' & StringLen($sString) + 1 & ']')
	$tString.Data = $sString
	DllStructSetData($tPropVariant, 5, DllStructGetPtr($tString))
	DllStructSetData($tPropVariant, 1, 31)
	Return $tPropVariant
EndFunc   ;==>_WPDPropVariantFromString
#EndRegion Functions

