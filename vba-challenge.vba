Sub GetInfo()

    'Create variables to hold information about the record holders
    dim Greatest_Incr as double, Greatest_Decr as double, Greatest_Vol as double
    dim Greatest_Incr_Tik as string, Greatest_Decr_Tik as string, Greatest_Vol_Tik as string
    
    For Each ws In Worksheets

        'Find out how many rows of data there are
        dim Endrow as double
        EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Create headers for a new table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Create dummy variable for the new table
        dim table_row as Double
        table_row = 2

        'Create variables for data to sum
        dim total_stock_v as Double 

        'Create variable to keep track of stock info at beginning of the year
        dim first as Double
        first = 2

        'Set Greatest Trackers to 0
        Greatest_Incr = 0
        Greatest_Incr_Tik = 0
        Greatest_Decr = 0
        Greatest_Decr_Tik = 0
        Greatest_Vol = 0
        Greatest_Vol_Tik = 0

        'Loop through Data
        Dim i As Double
        For i = 2 To EndRow

            If ws.Cells(i,1).value <> ws.cells(i+1, 1).value then 'If the next ticker is different

                total_stock_v = total_stock_v + ws.cells(i, 7).value     'Add in new stock volume
                ws.Cells(table_row, 9).Value = ws.Cells(i,1).Value    'Write ticker into table
                ws.Cells(table_row, 12).Value = total_stock_v      'Write total stock amount into table
                
                'Yearly change calculation
                ws.Cells(table_row, 10).Value =  ws.Cells(i, 6).Value - ws.Cells(first, 3).Value

                'Percent change calculation
                ws.Cells(table_row, 11).Value = (ws.Cells(i, 6).Value - ws.Cells(first, 3).Value)/(ws.Cells(first, 3).Value)

                'Format cells in table 
                if ws.Cells(table_row, 10).Value < 0 then  'If yearly change is negative
                    ws.Cells(table_row, 10).interior.colorindex = 3    'color cell red
                else    'if yearly change is positive
                    ws.Cells(table_row, 10).interior.colorindex = 4    'color cell green
                end if

                'Track greatest Increases and Decreases
                If ws.Cells(table_row, 11).Value > Greatest_Incr then
                    Greatest_Incr = ws.Cells(table_row, 11).Value
                    Greatest_Incr_Tik = ws.Cells(table_row, 9).Value

                ElseIf ws.Cells(table_row, 11).Value < Greatest_Decr then
                    Greatest_Decr = ws.Cells(table_row, 11).Value
                    Greatest_Decr_Tik = ws.Cells(table_row, 9).Value
                End if

                'Track the Greatest Volume
                If ws.Cells(table_row, 12).Value > Greatest_Vol then
                    Greatest_Vol = ws.Cells(table_row, 12).Value
                    Greatest_Vol_Tik = ws.Cells(table_row, 9).Value
                End If

                'Format percent change column
                ws.Cells(table_row, 11).NumberFormat = "0.00%"

                'Ready holder variables for new stock
                total_stock_v = 0 'reset total stock volume
                table_row = table_row + 1   'Ready next row of table
                first = i + 1   'change to first row of new stock

            Else 'If the next ticker is the same
                total_stock_v = total_stock_v + ws.Cells(i, 7).Value     'Add in new stockvolume

            End If
            
        Next i
        
    'Make Table for Greatest Values
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("P2").Value = Greatest_Incr_Tik
    ws.Range("Q2").Value = Greatest_Incr
    ws.Range("P3").Value = Greatest_Decr_Tik
    ws.Range("Q3").Value = Greatest_Decr
    ws.Range("P4").Value = Greatest_Vol_Tik
    ws.Range("Q4").Value = Greatest_Vol

    'Format percent change column
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    Next ws

End Sub