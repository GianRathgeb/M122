Sub CleanData()
    ScreenUpdating = False
    

    Dim WS_Count As Integer
    Dim I As Integer

    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
    ' Begin the loop.
    For I = 1 To WS_Count
        
        Dim c As Range
        Dim str As String
            
        Dim wks As Worksheet
        Dim lngLastRow As Long, lngLastCol As Long, lngIdx As Long, _
            lngColCounter As Long
        Dim blnAllBlank As Boolean
        Dim UserInputSheet As String
    
        Set wks = Worksheets(ActiveWorkbook.Worksheets(I).Name)

        With wks
        'Now that our sheet is defined, we'll find the last row and last column
            lngLastRow = .Cells.Find(What:="*", LookIn:=xlFormulas, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlPrevious).Row
            lngLastCol = .Cells.Find(What:="*", LookIn:=xlFormulas, _
                                     SearchOrder:=xlByColumns, _
                                     SearchDirection:=xlPrevious).Column
            
            'Since we need to delete rows, we start from the bottom and move up
            For lngIdx = lngLastRow To 1 Step -1
        
                'Start by setting a flag to immediately stop checking
                'if a cell is NOT blank and initializing the column counter
                blnAllBlank = True
                lngColCounter = 2
        
                'Check cells from left to right while the flag is True
                'and the we are within the farthest-right column
                While blnAllBlank And lngColCounter <= lngLastCol
        
                    'If the cell is NOT blank, trip the flag and exit the loop
                    If .Cells(lngIdx, lngColCounter) <> "" Then
                        blnAllBlank = False
                    Else
                        lngColCounter = lngColCounter + 1
                    End If
        
                Wend
        
                'Delete the row if the blnBlank variable is True
                If blnAllBlank Then
                    .Rows(lngIdx).Delete
                End If
    
            Next lngIdx
            
        End With
        
        With wks
        'Now that our sheet is defined, we'll find the last row and last column
            lngLastRow = .Cells.Find(What:="*", LookIn:=xlFormulas, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlPrevious).Row
            lngLastCol = .Cells.Find(What:="*", LookIn:=xlFormulas, _
                                     SearchOrder:=xlByColumns, _
                                     SearchDirection:=xlPrevious).Column
            
            
            ' Delete fld in first line
            For Each c In .Range("A1:G1")
                If Left(c.Value, 3) = "fld" Then
                    c.Value = Replace(c.Value, "fld", "")
                End If
            Next c
            
            Dim strTotal
            Dim strRangeTotal
            Set strTotal = .Range("F" + CStr(lngLastRow + 1))
            strTotal.Formula2Local = "Total"
            strRangeTotal = "G2:G" + CStr(lngLastRow - 1)
            Set strTotal = .Range("G" + CStr(lngLastRow + 1))
            strTotal.Formula2Local = "=Summe(" + CStr(strRangeTotal) + ")"
            Worksheets(I).Activate
            Range("F" + CStr(lngLastRow + 1) + ":G" + CStr(lngLastRow + 1)).Select
            Selection.Font.Bold = True
           
            
        End With

    Next I
    ScreenUpdating = True
    
End Sub
