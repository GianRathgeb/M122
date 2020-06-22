Option Explicit

Dim Number, Output

Number = InputBox("Please Input 1, 2 or 3")
If IsNumeric(Number) Then
    Select Case Number
        Case 1 Output = "number One"
        Case 2 Output = "nubmer Two"
        Case 3 Output = "number Three"
        Case Else Output = "None of the numbers inputed"
        End Select
    Else
        Output = "Input is not numberic"
    end If

MsgBox "You choosed " & Output & Chr(46)