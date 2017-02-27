Sub UnmergeTest()

    Dim thisCell, thisRow, thisCol, numRows, nextCellDown, alertsOn, mergeRange, rowHeight, prevCol, currentCell
    Dim values()
    Dim segments, x, i, a
    
        
    ' ------ get & format address for active cell. may not need later ------
    
    thisCell = ""
    thisCol = ""
    segments = Split(ActiveCell.Address, "$")
    For Each x In segments
        thisCell = thisCell + x
        If thisCol = "" Then
            thisCol = x
        Else
            thisRow = x
        End If
    Next
    
    rowHeight = ActiveCell.rowHeight
    
    
    nextCellDown = thisCol & (CInt(thisRow) + 1)
    
    
    ' ------ see how many rows in selected cell ------
    
    numRows = 0
    segments = Split(ActiveCell.Value, Chr(10))
    For Each x In segments
        ' change ubound of array, get individual values, increment counter
        numRows = numRows + 1
        ReDim Preserve values(numRows)
        values(numRows) = x
    Next
    
    ' for # of values-1 (return-separated) insert that many rows & insert each value in the new row
    
    numRows = numRows - 1
    If numRows < 1 Then Exit Sub
    
    For i = 1 To numRows
        Range(thisCol & (CInt(thisRow) + i)).EntireRow.Insert
        Range(thisCol & (CInt(thisRow) + i)).Value = values(i + 1)
    Next
    
    ' see if alerts are on; if so, turn them off temporarily
    alertsOn = False
    If Application.DisplayAlerts = True Then Application.DisplayAlerts = False: alertsOn = True
    
    ' then delete all but the 1st value from the original cell
    Range(thisCell).Value = values(1)
    
    ' then merge the preceeding cells (w/o alert showing)
    '   *** have to merge A10 through A12, B10 through B12, etc. up through the preceeding column ***
    
    'mergeRange = "A" & thisRow & ":" & thisCol & (CInt(thisRow) + numRows)
    
    ' get .Previous.Address, then keep going until previous = A
    
    prevCol = ""
    Set currentCell = ActiveCell
    
    Do Until prevCol = "A"
        ' get previous column
        segments = Split(currentCell.Previous.Address, "$")
        a = 0
        For Each x In segments
            If a = 1 Then prevCol = x
            a = a + 1
        Next
        
        Set currentCell = currentCell.Previous
        
        ' merge previous column
        mergeRange = prevCol & thisRow & ":" & prevCol & (CInt(thisRow) + numRows)
        Range(mergeRange).Merge
    Loop
    
    ' reset row heights
    For i = 0 To numRows
        Range(thisCol & (CInt(thisRow) + i)).rowHeight = (rowHeight) / (numRows + 1)
    Next
    
    If alertsOn Then Application.DisplayAlerts = True
    
    
End Sub
