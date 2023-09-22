# VBA-challenge
Code below
Sub Multi_year_stock()

    'loop through all the worksheets in the workbook
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
    

    Dim ticker As String
    Dim open_years_beginning As Double
    Dim close_years_end As Double
    Dim yearly_change As Double
    Dim volume As Long
    Dim years_percentage_change As Double
    Dim summary_table_row As Integer
    Dim lastrow As Long
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Long
    Dim ticker_greatest_increase As String
    Dim ticker_greatest_decrease As String
    Dim ticker_greatest_volume As String
    
    
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
   
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    summary_table_row = 2
    
     'header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greaest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
        
    
    For i = 2 To lastrow
        'Values for the summary
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        
        ticker = ws.Cells(i, 1).Value
        
        open_years_beginning = ws.Cells(i, 3).Value
        
        close_years_end = ws.Cells(i, 6).Value
        
        
        
        ' Check constraints before calculating yearly change
        yearly_change = close_years_end - open_years_beginning
        
        If open_years_beginning <> 0 Then
            
        years_percentage_change = (yearly_change / open_years_beginning)
        
        Else
        years_percentage_change = 0
        
        End If

                
        ws.Cells(summary_table_row, 9).Value = ticker
        ws.Cells(summary_table_row, 10).Value = yearly_change
        ws.Cells(summary_table_row, 11).Value = years_percentage_change
        ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        
        'Calculate total volume of stock
        ws.Cells(summary_table_row, 12).Value = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i - volume + 1, 7), ws.Cells(i, 7)))
        
        'Color coding displaing positive and negative yearly changes
            If years_percentage_change > 0 Then
            ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
        
            ElseIf years_percentage_change < 0 Then
            ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
            End If
            
            
          'get tickers and values for greatest increases and decreases
            
         If volume > greatest_volume Then
         greatest_volume = volume
         ticker_greatest_volume = ticker
         End If

         If years_percentage_change > greatest_increase Then
         greatest_increase = years_percentage_change
         ticker_max_increase = ticker
         End If

        If years_percentage_change < greatest_decrease Then
        greatest_decrease = years_percentage_change
        ticker_gretest_decrease = ticker
        End If
        
        
         ws.Cells(2, 16).Value = ticker_greatest_increase
         ws.Cells(3, 16).Value = ticker_greatest_decrease
         ws.Cells(4, 16).Value = ticker_greatest_volume

        ws.Cells(2, 17).Value = greatest_increase
        ws.Cells(3, 17).Value = greatest_decrease
        ws.Cells(4, 17).Value = greatest_volume
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
            
        summary_table_row = summary_table_row + 1
      
        
        volume = 0
        
       
        
        End If

        Next i
        Next ws
