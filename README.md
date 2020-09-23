<div align="center">

## CheckFileVersion


</div>

### Description

Retrieve the version of a file (EXE/DLL etc). This code should be paste into a module and just called via CheckFileVersion('Path to the Exe or DLL').
 
### More Info
 
Path to the EXE or DLL file eg. "C:\Windows\Notepad.exe"

A Variant containing the version of the file eg. "4.10"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Riaan Aspeling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/riaan-aspeling.md)
**Level**          |Unknown
**User Rating**    |4.3 (47 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/riaan-aspeling-checkfileversion__1-1589/archive/master.zip)

### API Declarations

```
Option Explicit
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersionl As Integer   ' e.g. = &h0000 = 0
  dwStrucVersionh As Integer   ' e.g. = &h0042 = .42
  dwFileVersionMSl As Integer  ' e.g. = &h0003 = 3
  dwFileVersionMSh As Integer  ' e.g. = &h0075 = .75
  dwFileVersionLSl As Integer  ' e.g. = &h0000 = 0
  dwFileVersionLSh As Integer  ' e.g. = &h0031 = .31
  dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
  dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
  dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
  dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
  dwFileFlagsMask As Long    ' = &h3F for version "0.42"
  dwFileFlags As Long      ' e.g. VFF_DEBUG Or VFF_PRERELEASE
  dwFileOS As Long        ' e.g. VOS_DOS_WINDOWS16
  dwFileType As Long       ' e.g. VFT_DRIVER
  dwFileSubtype As Long     ' e.g. VFT2_DRV_KEYBOARD
  dwFileDateMS As Long      ' e.g. 0
  dwFileDateLS As Long      ' e.g. 0
End Type
```


### Source Code

```
'Example to use this function
'  MsgBox " Notepad's Version is " & CheckFileVersion("C:\Windows\Notepad.exe")
Public Function CheckFileVersion(FilenameAndPath As Variant) As Variant
On Error GoTo HandelCheckFileVersionError
  Dim lDummy As Long, lsize As Long, rc As Long
  Dim lVerbufferLen As Long, lVerPointer As Long
  Dim sBuffer() As Byte
  Dim udtVerBuffer As VS_FIXEDFILEINFO
  Dim ProdVer As String
  lsize = GetFileVersionInfoSize(FilenameAndPath, lDummy)
  If lsize < 1 Then Exit Function
  ReDim sBuffer(lsize)
  rc = GetFileVersionInfo(FilenameAndPath, 0&, lsize, sBuffer(0))
  rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
  MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
  '**** Determine Product Version number ****
  ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl)
  CheckFileVersion = ProdVer
  Exit Function
HandelCheckFileVersionError:
  CheckFileVersion = "N/A"
  Exit Function
End Function
```

