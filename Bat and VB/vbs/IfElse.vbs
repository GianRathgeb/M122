Option Explicit

Dim Number

Number = InputBox("Please input a number:")
If IsNumeric(Number) Then
    MsgBox "Number: " & Number, 64
Else 
    MsgBox "That's not a number", 48
end if