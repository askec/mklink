# mklink
Stand alone and Total Commander interface for Windows mklink (NTFS symbolic links, hardlink and junction points)

## Description

* VBScript to create symbolic, hardlinks or directory junctions via a rudimentary GUI that needs parameters.
* Compatible with modern Windows versions, including Windows 11.
* Supports multiple files/directory at once to one destination folder.
* Make hard link or directory junction instead of symbolic links if file or directory is on the same volume.

Note: This script requires the 'mklink_gui.hta' file to be in the same directory.

## Usage
cscript /nologo mklink.vbs "C:\Path\To\DestinationFolder" "C:\Path\To\ListFile.txt"

### Parameters
`%1`: Destination Folder: The folder where the new links will be created.  
`%2`: Source List File: A text file containing one source file or folder path per line.  

### Exit Codes
`0`: Success, no warnings.  
`1`: Script was cancelled by user.  
`2`: The base destination directory does not exist.  
`3`: Success, but with one or more warnings (e.g., source not found, destination exists).  

## Installation

1. Just download all the files from src folder. 
1. Copy all files in a folder anywhere to your disk.
1. Configure a command or a single button in Total Commander :

> Command: `cscript`  
> Parameters: `/noLogo "<path_of_mklink.vbs>\mklink.vbs" "%T" "%L"`  
> Icon file: `<path_of_mklink.vbs>\mklink.ico`  
> Tooltip: `Make NTFS Link`  

`<path_of_mklink.vbs>` is the path where you copied mklink.vbs.
