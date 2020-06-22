Option Explicit

Dim Files, WinDir

Set Files = CreateObject("Scripting.FileSystemObject")
Set WinDir = Files.GetSpecialFolder(0)

WScript.Echo "Windows Directory located at " & WinDir

Set WinDir = Files.GetSpecialFolder(1)
WScript.Echo "Windows System Directory located at " & WinDir
