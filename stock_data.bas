Attribute VB_Name = "Module1"
Sub stock_data()
    
    Dim ws As Worksheet
    Dim data As Variant
    Dim results As Variant
    Dim ticker As String
    Dim last_row As Long
    Dim i As Long
    Dim output_row As Long
    Dim quarterly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim first_open As Double
    Dim last_close As Double
    Dim current_ticker As String
    Dim next_ticker As String
    Dim first_row_of_ticker As Long
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_volume_ticker As String
    
    
    'Loop through each sheet
    For Each ws In ThisWorkbook.Sheets(Array("Q1", "Q2", "Q3", "Q4"))
    
        'Initialize variables for greatest increase, decrease,and volume
        greatest_increase = -999999
        greatest_decrease = 999999
        greatest_volume = 0
    
        'Find last row in column A for the current sheet
        last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
        'Load data in columns A to G, minus headers
        data = ws.Range("A2:G" & last_row).Value
        
        'Create array to store results incolumns I to L
        ReDim results(1 To last_row, 1 To 4)
         
        'Add the headers to columns I - L
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Set first row to check for ticker
        first_row_of_ticker = 1
        output_row = 1
        total_volume = 0
    
        'Loop through all rows in array to calculate quarterly change
        For i = 1 To UBound(data)
            current_ticker = data(i, 1)
            
            'Check if last row before accessing the next row
            If i < UBound(data) Then
                next_ticker = data(i + 1, 1)
            Else
                next_ticker = ""
            End If
            
            total_volume = total_volume + data(i, 7)
            
            If current_ticker <> next_ticker Or i = UBound(data) Then
                first_open = data(first_row_of_ticker, 3)
                last_close = data(i, 6)
                
                'Calculate quarterly change and round 2 decimal places
                quarterly_change = Round(last_close - first_open, 2)
                
                'Calculate percent change and round 2 decimal places
                If first_open <> 0 Then
                    percent_change = quarterly_change / first_open
                Else
                    percent_change = 0
                End If
                
                'Save the results in results array
                results(output_row, 1) = current_ticker
                results(output_row, 2) = quarterly_change
                results(output_row, 3) = percent_change
                results(output_row, 4) = total_volume
                
                'Apply color formatting to quarterly change
                If quarterly_change > 0 Then
                    ws.Cells(output_row + 1, 10).Interior.Color = vbGreen
                ElseIf quarterly_change < 0 Then
                    ws.Cells(output_row + 1, 10).Interior.Color = vbRed
                Else
                    ws.Cells(output_row + 1, 10).Interior.ColorIndex = xlNone
                End If
 
                'Check for greatest increase, decrease, and volume
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_row = output_row + 1
                End If
                
                If percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_row = output_row + 1
                End If
                
                If total_volume > greatest_volume Then
                    greatest_volume = total_volume
                    greatest_volume_row = output_row + 1
                End If
                
                'Move to next output row
                output_row = output_row + 1
                
                'Reset for next ticker
                first_row_of_ticker = i + 1
                total_volume = 0
            End If
        
        Next i
        
        'Write results array back to sheet (columns I to L)
        ws.Range("I2:L" & output_row + 1).Value = results
        
        'Format Percent Change column as percentages
        ws.Range("K2:K" & output_row + 1).NumberFormat = "0.00%"
     
        ' Write greatest increase, decrease, and total volume results to sheet Q1
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
        'Add "Ticker" and "Value" headers to columns P and Q
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Populate the tickers
        ws.Cells(2, 16).Value = ws.Cells(greatest_increase_row, 9).Value
        ws.Cells(3, 16).Value = ws.Cells(greatest_decrease_row, 9).Value
        ws.Cells(4, 16).Value = ws.Cells(greatest_volume_row, 9).Value
        
        'Populate the values
        ws.Cells(2, 17).Value = ws.Cells(greatest_increase_row, 11).Value
        ws.Cells(3, 17).Value = ws.Cells(greatest_decrease_row, 11).Value
        ws.Cells(4, 17).Value = ws.Cells(greatest_volume_row, 12).Value
        
        ' Format the values
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
        
    Next ws
        
End Sub



