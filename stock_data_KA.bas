Attribute VB_Name = "Module1"
Sub StockData()

' For loop to run the code on all worksheets
For Each ws In Worksheets

    ' Set an initial variable for holding the stock ticker symbol
    Dim ticker As String

    ' Create intital accumulator variable for total volume and set it to zero
    Dim total_volume As Double
    total_volume = 0
    
    ' Set variable for the last row in the worksheet
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set the variable for what row to start inserting data generated
    Dim data_cell As Long
    data_cell = 2
    
    ' For loop that will go through all the rows with data in the worksheet
    For i = 2 To last_row
    
        ' Add the volume in the volume column of data to the total volume accumulator
        total_volume = total_volume + ws.Cells(i, 7).Value
    
        ' If statement that checks if the next ticker symbol matches the current ticker symbol
        ' If they do not match, the statements execute
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            
            ' Sets the row equivalent to the data_cell variable in the J column to the ticker symbol
            ws.Range("J" & data_cell).Value = ticker
            
            ' Sets the row equivalent to the data_cell variable in the K column to the total volume value
            ws.Range("K" & data_cell) = total_volume
            
            ' Sets the total_volume to zero as it needs to reset for the next ticker symbol
            total_volume = 0
            
            ' Increments the value of the data_cell to log the data by one to insert data into a new row
            data_cell = data_cell + 1
            
        End If
        
    Next i
                        
Next ws

End Sub
