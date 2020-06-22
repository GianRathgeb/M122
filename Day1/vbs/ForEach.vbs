Option Explicit
Dim Files, WinDir, AllFiles, AllFilesItems, Output

Set Files = CreateObject("Scripting.FileSystemObject")
Set WinDir = Files.GetSpecialFolder(0)
Set AllFiles = WinDir.Files


For Each AllFilesItems in AllFiles
    Output = Output & AllFilesItems.Path & ", "
Next

WScript.Echo "Files in " & WinDir & ":" & Chr(13) & Output
