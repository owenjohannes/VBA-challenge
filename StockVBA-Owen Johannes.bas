Attribute VB_Name = "Module4"
Sub TickerColumnYearlyPercentVolume()
    
    Dim i As Long ' loop counter
    Dim j As Long ' loop counter for unique values
    Dim foundMatch As Boolean ' flag to track whether a value has been found in the array
    Dim uniqueArray() As Variant ' array for unique values
    
    ' determine the last row of data in the column
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' loop through the rows of data and store the values in the array
    ReDim myArray(1 To lastRow)
    For i = 1 To lastRow
        myArray(i) = Cells(i, "A").Value
    Next i
    
    ' Insert header in column I
    Range("L1").Value = "Ticker"
    
    ' loop through the values in the array to find unique values
    ReDim uniqueArray(1 To lastRow)
    j = 0 ' initialize the counter for unique values
    For i = 2 To lastRow
        foundMatch = False ' reset the flag
        ' loop through the previous values in the array to check for matches
        For k = 1 To j
            If myArray(i) = uniqueArray(k) Then
                foundMatch = True ' set the flag if a match is found
                Exit For ' exit the loop if a match is found
            End If
        Next k
        ' if no match is found, store the value in the unique array and increment the counter
        If Not foundMatch Then
            j = j + 1
            uniqueArray(j) = myArray(i)
        End If
    Next i
    
    ' resize the unique array to the actual number of unique values
    ReDim Preserve uniqueArray(1 To j)
    
    ' paste the unique values in Column L
    Range("L2").Resize(j, 1).Value = Application.Transpose(uniqueArray)
    
    Dim outputRow As Long
    Dim currentVal As String, currentMin As Long, currentMax As Long
    Dim currentCVal As Variant, currentFVal As Variant, currentTotalVal As Double
    Dim rng As Range
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    outputRow = 2

    Range("M1").Value = "Yearly Change"
    Range("N1").Value = "Percent Change"
    Range("O1").Value = "Total Value"
    
For i = 2 To lastRow
    If Cells(i, 1) <> Cells(i - 1, 1) Then
        currentVal = Cells(i, 1)
        currentMin = Cells(i, 2)
        currentMax = Cells(i, 2)
        currentCVal = Cells(i, 3).Value
        currentFVal = Cells(i, 6).Value
        currentTotalVal = 0 'initialize total value to zero
        
    End If
    
    If Cells(i, 2) < currentMin Then
        currentMin = Cells(i, 2)
        currentCVal = Cells(i, 3).Value
    End If
    
    If Cells(i, 2) > currentMax Then
        currentMax = Cells(i, 2)
        currentFVal = Cells(i, 6).Value
    End If
    
    currentTotalVal = currentTotalVal + Cells(i, 7).Value ' add value from column G to currentTotalVal
    
    If i = lastRow Or Cells(i + 1, 1) <> currentVal Then
        Cells(outputRow, 13).Value = currentFVal - currentCVal
        Cells(outputRow, 14).Value = ((currentFVal - currentCVal) / currentCVal)
        Cells(outputRow, 14).NumberFormat = "0.00%"
        
        Cells(outputRow, 15).Value = currentTotalVal ' print total value to column O
        Range("O2:O" & outputRow - 1).NumberFormat = "0"
        
        outputRow = outputRow + 1
        
        currentVal = ""
        currentMin = 0
        currentMax = 0
        currentTotalVal = 0 ' reset total value to zero for next group
    End If
Next i

     'Determine last row in Column M
      lastRow = Cells(Rows.Count, "M").End(xlUp).Row
    
    'Set range to Column M, starting from the second row
    Set rng = Range("M2:M" & lastRow)
    
    'Clear any existing conditional formatting
    rng.FormatConditions.Delete
    
    'Remove any interior color from the first row
    Range("M1").Interior.ColorIndex = xlNone
    
    'Apply new conditional formatting
    With rng.FormatConditions.Add(xlCellValue, xlGreater, "0")
        .Interior.Color = vbGreen
    End With
    With rng.FormatConditions.Add(xlCellValue, xlLess, "0")
        .Interior.Color = vbRed
    End With
End Sub




    
    
    
    
