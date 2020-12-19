Attribute VB_Name = "Module1"
Sub stock_market():

    For Each ws In Worksheets 'loops through all worksheets
    
        'define headings
            ws.Cells(1, 9) = "Ticker" 'prints word "Ticker" into cell
            ws.Cells(1, 10) = "Yearly Change" 'same as above
            ws.Cells(1, 11) = "Percent Change" 'same as above
            ws.Cells(1, 12) = "Total Stock Volume" 'same as above
                
        'declare variables
            'ticker variables
            Dim ticker As String
            'yearly changes and total volume variables
            Dim open_price As Double
            Dim close_price As Double
            Dim yearly_change As Double
            Dim total_stock_volume As Double
            'percent change variables
            Dim percent_change As Double
            'other variables
            Dim i As Long
            Dim last_row As Long
            Dim i_summary As Long
            
        'assign variables
            last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row 'defines last row of column1(all changes in data will reflect symbol change)
            i_summary = 2
            open_price = 0
            close_price = 0
            yearly_change = 0
            total_stock_volume = 0
            percent_change = 0
    
            open_price = ws.Cells(2, 3).Value 'open price starts
    
        'iterate through data
            For i = 2 To last_row
                
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            ticker = ws.Cells(i, 1).Value 'finds ticker symbol
                       'yearly change and percent change
                            close_price = ws.Cells(i, 6).Value 'finds close price
                            yearly_change = close_price - open_price
                    If open_price <> 0 Then
                        percent_change = (yearly_change / open_price) * 100
                    End If
                    
                        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                        
                        ws.Range("i" & i_summary).Value = ticker 'prints ticker symbol
                        ws.Range("j" & i_summary).Value = yearly_change 'prints yearly change = yearly_change
                
                'define color in col J
                    If (yearly_change > 0) Then
                        ws.Range("j" & i_summary).Interior.Color = vbGreen 'turn green if greater than 0
                    ElseIf (yearly_change <= 0) Then
                        ws.Range("j" & i_summary).Interior.Color = vbRed 'turns red if even or negative
                    End If
                    
                        ws.Range("k" & i_summary).Value = (CStr(percent_change) & "%") 'prints percent col k as %
                        ws.Range("l" & i_summary).Value = total_stock_volume 'prints total volume in col L
                            
                        i_summary = i_summary + 1
                        yearly_change = 0 'resets to 0
                        close_price = 0 'resets to 0
                        total_stock_volume = 0 'resets to 0
                        open_price = ws.Cells(i + 1, 3).Value
                    
                    Else
                        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                    
                    End If
                    
            Next i
    
    Next ws
    
End Sub




