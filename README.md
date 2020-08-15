# VBA-loopThroughSheetsAndFormat

## Start File
![image](https://user-images.githubusercontent.com/52837649/90320805-26b4a600-df12-11ea-89ca-7f89dbf08c90.png)

## Task
1. Extract state from each worksheet tab

2. Add the state to the first column of each worksheet

3. Convert the headers of each row to simply display the year

4. Convert the numbers to currency values for all cells

## Finished File
![image](https://user-images.githubusercontent.com/52837649/90320828-4d72dc80-df12-11ea-9e31-04a5751a46dc.png)

## Code
```
Sub addState()

For Each ws In Worksheets

    Dim wsName As String
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    wsName = ws.Name
    
    'split the State out of the worksheet name
    State = Split(wsName, "_")
    
    'add a new column
    ws.Range("A1").EntireColumn.Insert
    
    'add the word State to the first column header
    ws.Cells(1, 1).Value = "State"
    
    'add state to all rows
    ws.Range("A2:A" & lastRow) = State(0)
    
    'determine the last columne
    lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    For i = 3 To lastColumn
    
        yearHeader = ws.Cells(1, i).Value
        yearSplit = Split(yearHeader, " ")
        ws.Cells(1, i).Value = yearSplit(3)
        
    Next i
    
    For i = 2 To lastRow
        
        For j = 2 To lastColumn
        
            ws.Cells(i, j).Style = "Currency"
        
        Next j
        
    Next i
       
Next ws
        
End Sub
```

