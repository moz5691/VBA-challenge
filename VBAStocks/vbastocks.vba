Sub Stock_Report()

'To deal with large dataset, code should be efficient.
'The idea is check <prev> <current> <next> rows.
'If not <prev> == <current>, the beginning of new ticker
'if not <current> == <next>, the end of the current ticker.
'Using one for-loop for volume/change from all tickers
'Using one for-loop for best/worst performance tickers.
    
    Dim ws As Worksheet

    For Each ws In Worksheets

        'Create column labels for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        Dim ticker_symbol As String
        Dim total_stock_volume As Double
        Dim row_count As Long
        Dim year_start_price As Double
        Dim year_end_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim last_row As Long
 
        total_stock_volume = 0
        row_count = 2
        year_start_price = 0
        year_end_price = 0
        yearly_change = 0
        percent_change = 0
        
        'https://www.listendata.com/2013/05/excel-3-ways-to-extract-unique-values.html
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop to search through ticker symbols
        For i = 2 To last_row
            
            'if the current ticker is different from the previous, it is the beginning of new ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                year_start_price = ws.Cells(i, 3).Value
            End If

            'Total stock volume
            total_stock_volume = total_stock_volume + ws.Cells(i, 7)

            'if the current ticker is different from the next, it is the end of the current ticker
            'move all collected data and clean up
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                'Move ticker symbol to summary table
                ws.Cells(row_count, 9).Value = ws.Cells(i, 1).Value

                'Move total stock volume to the summary table
                ws.Cells(row_count, 12).Value = total_stock_volume

                'Year end price
                year_end_price = ws.Cells(i, 6).Value

                'Calculate yearly_change
                yearly_change = year_end_price - year_start_price
                ws.Cells(row_count, 10).Value = yearly_change

                'Format cell for "Yearly Change" and "Percent Change".
                If yearly_change >= 0 Then
                    'positive change - green background/black font
                    With ws.Cells(row_count, 10)  'Yearly Change column
                        .Interior.ColorIndex = 4
                        .Font.Color = vbBlue
                    End With
                    With ws.Cells(row_count, 11) 'Percent Change column
                        .Interior.ColorIndex = 4
                        .Font.Color = vbBlue
                    End With
                Else
                    'negative change - red background/white font
                    With ws.Cells(row_count, 10)  'Yearly Change column
                        .Interior.ColorIndex = 3
                        .Font.Color = vbWhite
                    End With
                    With ws.Cells(row_count, 11)  'Percent Change column
                        .Interior.ColorIndex = 3
                        .Font.Color = vbWhite
                    End With
                End If

                'Calculate percent_change.
                'if year_start_price or year_end_price is zero, we cannot calculate
                If year_start_price = 0 Or year_end_price = 0 Then
                    ws.Cells(row_count, 11).Value = 0
                Else
                    percent_change = yearly_change / year_start_price
                    With ws.Cells(row_count, 11)
                        .Value = percent_change
                        .NumberFormat = "0.00%"
                    End With
                End If

                'move to the next row
                row_count = row_count + 1

                'Clear all number variables. Important!!!
                total_stock_volume = 0
                year_start_price = 0
                year_end_price = 0
                yearly_change = 0
                percent_change = 0
                
            End If
        Next i

        'Top and bottom performance table
        'Titles
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        'reuse last_row variable to store last row of 'Ticker' column
        last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row


        Dim best_ticker As String
        Dim best_percent_change As Double
        Dim worst_ticker As String
        Dim worst_percent_change As Double
        Dim most_volume_ticker As String
        Dim most_volume_value As Double

        'first row of "Percent Change" column and "Total Stock Volume"
        best_percent_change = ws.Range("K2").Value
        worst_percent_change = ws.Range("K2").Value
        most_volume_value = ws.Range("L2").Value

        'loop through to compare to the next row.
        For j = 2 To last_row
            
            If ws.Cells(j, 11).Value > best_percent_change Then
                best_percent_change = ws.Cells(j, 11).Value
                best_ticker = ws.Cells(j, 9).Value
            End If

            If ws.Cells(j, 11).Value < worst_percent_change Then
                worst_percent_change = ws.Cells(j, 11).Value
                worst_ticker = ws.Cells(j, 9).Value
            End If

            If ws.Cells(j, 12).Value > most_volume_value Then
                most_volume_value = ws.Cells(j, 12).Value
                most_volume_ticker = ws.Cells(j, 9).Value
            End If

        Next j

        'copy to cells to create report.
        ws.Range("P2").Value = best_ticker
        With ws.Range("Q2")
            .Value = best_percent_change
            .NumberFormat = "0.00%"
            .Font.Color = vbBlue
        End With
        
        ws.Range("P3").Value = worst_ticker
        With ws.Range("Q3")
            .Value = worst_percent_change
            .NumberFormat = "0.00%"
            .Font.Color = vbRed
        End With
        
        ws.Range("P4").Value = most_volume_ticker
        With ws.Range("Q4")
            .Value = most_volume_value
            .NumberFormat = "#,###,##0"
        End With

        'autofit
        ws.Columns("I:Q").EntireColumn.AutoFit

    Next ws

End Sub


