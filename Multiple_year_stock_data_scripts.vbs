Sub ClearContents()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Range("I:P").ClearContents
        ws.Cells.ClearFormats
    
    Next ws
End Sub

Sub WorksheetLoop()
    ' Loop code through all worksheets.  Information from https://excelchamps.com/vba/loop-sheets/
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ' Column A (1) of interest
        Dim column As Integer
        column = 1
        Dim i, lastrow As Long ' i as row, lastrow to store max row count of each worksheet
        lastrow = ws.Cells(Rows.Count, column).End(xlUp).Row
        
        Dim summary_row, lastrow_sum As Integer ' lastrow_sum to store row count of ticker summary table
        summary_row = 2
        
        Dim yearopen, yearclose, totalvol As Variant
        yearopen = ws.Cells(2, 3).Value ' cell C2; initial opening price of first ticker
        totalvol = 0 ' total volume for each ticker
        
        Dim yearchange As Variant
        Dim pctchange As Variant
    
        ' Set up headers for ticker summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
    
        For i = 2 To lastrow
            ' Conditional when next row value is different than that of the current row value
            ' (helps when data is sorted for column of interest)
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
            
                ' adds volume (col 7) of current row before displaying summary in columns I and L
                totalvol = totalvol + ws.Cells(i, 7).Value
                ws.Range("I" & summary_row).Value = ws.Cells(i, column).Value 'display ticker (sorted)
                ws.Range("L" & summary_row).Value = totalvol
                
                ' capture yearclose and calculate yearly change and percent change; display in columns J and K; number formatting
                yearclose = ws.Cells(i, 6).Value ' Date (B) column is also ascending sorted
                yearchange = yearclose - yearopen
                ws.Range("J" & summary_row).Value = yearchange
                ws.Range("J" & summary_row).NumberFormat = "0.00" ' format to 2 decimal places
                pctchange = yearchange / yearopen ' yearopen gets the same snapshot values for table summary
                ws.Range("K" & summary_row).Value = Format(pctchange, "0.00%") ' format to percent values
                
                ' conditional color formatting of columns J and K; no color cell change when variable = 0
                If yearchange > 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4 ' format cell fill to green
                ElseIf yearchange < 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3 ' format cell fill to red
                End If
                If pctchange > 0 Then
                    ws.Range("K" & summary_row).Interior.ColorIndex = 4 ' format cell fill to green
                ElseIf pctchange < 0 Then
                    ws.Range("K" & summary_row).Interior.ColorIndex = 3 ' format cell fill to red
                End If
            
                summary_row = summary_row + 1 'increment summary_row
                totalvol = 0 'reset total volume for next ticker
                yearopen = ws.Cells(i + 1, 3).Value 'gets next ticker's opening price (column A and B asc sorted)
            
            Else
                totalvol = totalvol + ws.Cells(i, 7).Value
            End If
        
        Next i
        
        column = 9  ' Set for Ticker column (I)
        lastrow_sum = ws.Cells(Rows.Count, column).End(xlUp).Row ' Get count of rows for summary table
            
        Dim ticker, tmaxpct, tminpct, tmaxvol As String
        Dim pct, vol, pctmax, pctmin, volmax As Variant
        pctmax = 0
        pctmin = 0
        volmax = 0
    
        For i = 2 To lastrow_sum
            ticker = ws.Cells(i, column).Value ' store ticker name of row
            pct = ws.Cells(i, column + 2).Value ' store percent value of row
            vol = ws.Cells(i, column + 3).Value ' store volume value of row
            
            If pct > pctmax Then
                pctmax = pct
                tmaxpct = ticker
            ElseIf pct = pctmax Then
                MsgBox ("There is another maximum") ' added these Elseif statements just in case; no hits when run
            End If
            
            If pct < pctmin Then
                pctmin = pct
                tminpct = ticker
            ElseIf pct = pctmin Then
                MsgBox ("There is another minimum") ' added these Elseif statements just in case; no hits when run
            End If
            
            If vol > volmax Then
                volmax = vol
                tmaxvol = ticker
            ElseIf vol = volmax Then
                MsgBox ("There is another max volume") ' added these Elseif statements just in case; no hits when run
            End If
        
        
        Next i
        
        ' Output variables and summary table
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(2, 15).Value = tmaxpct
        ws.Cells(3, 15).Value = tminpct
        ws.Cells(4, 15).Value = tmaxvol
        
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 16).Value = Format(pctmax, "0.00%")
        ws.Cells(3, 16).Value = Format(pctmin, "0.00%")
        ws.Cells(4, 16).Value = volmax

        
        ws.Columns("I:P").AutoFit 'autofit
    
        
    Next ws
End Sub

