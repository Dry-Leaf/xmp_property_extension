A Windows property handler that allows reading data embedded in PNG and GIF files, and interpreting them as tags
in the same way natively supported by JPEG. 

Intended to be used in conjunction with [this utility](https://github.com/Dry-Leaf/tag-injector), which pulls
tag information from image boorus and inserts them into files.

Also works with the JXL property handler and thumbnailer implemented 
[here](https://github.com/saschanaz/jxl-winthumb).

## Installation

1. Download the latest dll from releases(only x86_64 has been tested)
2. Open a terminal window as administrator
3. Move to your download directory
4. `regsvr32 xmp_property_extension.dll`, or to uninstall, `regsvr32 /u xmp_property_extension.dll`

You may need to restart Windows Explorer in task manager a few times.

## Inspired by
* [FileMeta](https://github.com/Dijji/FileMeta)
* [jxl-winthumb](https://github.com/saschanaz/jxl-winthumb)
* [MantaPropertyExtension](https://github.com/sanje2v/MantaPropertyExtension)
