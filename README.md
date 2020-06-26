# AutoIt MediaDevice

[![Latest Version](https://img.shields.io/badge/Latest-v1.0.0-green.svg)]()
[![AutoIt Version](https://img.shields.io/badge/AutoIt-3.3.14.5-blue.svg)]()
[![MIT License](https://img.shields.io/github/license/mashape/apistatus.svg)]()
[![Made with Love](https://img.shields.io/badge/Made%20with-%E2%9D%A4-red.svg?colorB=e31b23)]()


AutoIt MediaDevice Library Allows you to communicate with attached media and storage devices.


## Features
* List Files and Folders.
* Delete File and Folders.
* Copy File From Device.
* Copy File To Device.


## Usage

##### Basic use:
```autoit

#include "..\MediaDevice.au3"

Local $sDevice = _MD_DeviceGet() ;Get First Found Portable Device
Local $oDevice = _MD_DeviceOpen($sDevice)
Local $aDrives = _MD_DeviceGetDrives($oDevice)
_ArrayDisplay($aDrives,"Storage Drives")
_MD_DeviceClose($oDevice)

```

##### More examples [here.](/Examples)


## Release History
See [CHANGELOG.md](CHANGELOG.md)


<!-- ## Acknowledgments & Credits -->


## License

Usage is provided under the [MIT](https://choosealicense.com/licenses/mit/) License.

Copyright © 2020, [Danysys.](https://www.danysys.com)