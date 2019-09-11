Attribute VB_Name = "Module2"
Sub stocks():
'Declaring the variables
    Dim Total_Vol As Double
    Dim inital_open As Double
    Dim close_val As Double
    Dim lastrow As Double
    Dim Ticker_name As String
    Dim largest_vol As Double
    
    largest_vol = 0
    Greatest_increase = 0
    Greatest_decrease = 0
    
    Dim index As Integer
    index = 0
    
    'go to each worksheet and do the same steps
    For Each ws In Worksheets
        
        'last row number of the data table on that sheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'reset total vol as 0 everytime its a new sheet
        Total_Vol = 0
        Final_Table = 2
        'set first value for intial opening price for each sheet
        initial_open = ws.Cells(2, 3).Value
        'get sheet name
        WorksheetName = ws.Name
        'MsgBox WorksheetName
        
        'Titles for result table
        ws.Cells(1, 11).Value = "Ticker_name"
        ws.Cells(1, 12).Value = "Total_Volume"
        ws.Cells(1, 13).Value = "Yearly_Change"
        ws.Cells(1, 14).Value = "Percent_Change"
        
        'process through each row
        For currRow = 2 To lastrow
            
            'If next ticker is not same that means current row is last row of the ticker data
            If ws.Cells(currRow + 1, 1).Value <> ws.Cells(currRow, 1).Value Then
                'get ticker name of the row and keep adding to total volume
                Ticker_name = ws.Cells(currRow, 1).Value
                Total_Vol = Total_Vol + ws.Cells(currRow, 7).Value
                
                'get closing value of the ticker's last row
                close_val = ws.Cells(currRow, 6).Value
                'Print values in output table
                ws.Cells(Final_Table, 11).Value = Ticker_name
                ws.Cells(Final_Table, 12).Value = Total_Vol
                ws.Cells(Final_Table, 13).Value = close_val - initial_open
                
                'Get Percent change when initial value is not 0
                'If initial value is 0, percent change is diplayed as 0
                If initial_open = 0 Then
                    ws.Cells(Final_Table, 14).Value = 0
                Else
                    ws.Cells(Final_Table, 14).Value = ws.Cells(Final_Table, 13).Value / initial_open
                End If

                ws.Cells(Final_Table, 14).NumberFormat = "0.00%"
                
                'Format color index for postive(green) and negative (red) for yearly change
                If ws.Cells(Final_Table, 14).Value > 0 Then
                    ws.Cells(Final_Table, 13).Interior.ColorIndex = 4
                Else
                    ws.Cells(Final_Table, 13).Interior.ColorIndex = 3
                End If
                
                'Reset Total_Vol, initial opening price and final table for next ticker iteration
                Total_Vol = 0
                initial_open = ws.Cells(currRow + 1, 3).Value
                Final_Table = Final_Table + 1
            Else
                'if the ticker value equals the next ticker, then keep adding Total_Vol
                Total_Vol = Total_Vol + ws.Cells(currRow, 7).Value
                'get new closeing price
                close_val = ws.Cells(currRow, 6).Value
            End If
        Next currRow
            'index keeps a check for each sheet
            If index = 0 Then
                'get largest volume for first sheet
                largest_vol = Application.WorksheetFunction.Max(ws.Columns("L"))
                'largest_volTickerRow = Application.WorksheetFunction.Match(largest_vol, Range("L:L"), 0)
                'Application.WorksheetFunction.Match(largest_vol, Range("L2:L300"), 1)

                'get greatest % increase and decrease stock and value
                Greatest_increase = Application.WorksheetFunction.Max(ws.Columns("N"))
                Greatest_decrease = Application.WorksheetFunction.Min(ws.Columns("N"))
            Else
                'get largest volume for next sheets
                worksheetMaxVal = Application.WorksheetFunction.Max(ws.Columns("L"))
                worksheetGreatest_increase = Application.WorksheetFunction.Max(ws.Columns("N"))
                worksheetGreatest_decrease = Application.WorksheetFunction.Min(ws.Columns("N"))
                
                'compare largest volume from first sheet with active sheet
                'reset largest volume when greater value if found
                If largest_vol < worksheetMaxVal Then
                    largest_vol = worksheetMaxVal
                End If
                
                'compare greatest % increase from first sheet with active sheet
                'reset greatest % increase when greater value if found
                If Greatest_increase < worksheetGreatest_increase Then
                    Greatest_increase = worksheetGreatest_increase
                End If
                
                'compare greatest % decrease from first sheet with active sheet
                'reset greatest % decrease when lesser value if found
                If Greatest_decrease > worksheetGreatest_decrease Then
                    Greatest_decrease = worksheetGreatest_decrease
                End If
            End If
                'increment to next sheet
                index = index + 1
                
            Next ws
            'print values for greatest total volume, % increase and % decrease
            Cells(2, 16).Value = "Greatest total volume"
            Cells(2, 17).Value = largest_vol
            Cells(3, 16).Value = "Greatest % increase"
            Cells(3, 17).Value = Greatest_increase
            Cells(3, 17).NumberFormat = "0.00%"
            Cells(4, 16).Value = "Greatest % Decrease"
            Cells(4, 17).Value = Greatest_decrease
            Cells(4, 17).NumberFormat = "0.00%"

End Sub
