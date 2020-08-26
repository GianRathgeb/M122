Option Explicit
Dim File, Lines

Set File = CreateObject("Scripting.FileSystemObject")
Set Lines = File.OpenTextFile("ForEach.vbs")
do
    WScript.Echo Lines.ReadLine
loop until Lines.AtEndOfStream
