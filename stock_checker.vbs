Sub alphabet_test()

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
   
        ' Declare variable for ticker
        Dim ticker As String
       
        ' Set initial variable for total stock volume
        Dim total_stock_volume As Double
        total_stock_volume = 0
       
        ' Set variable for summary table row to hold stock information
        Dim summary_table_row As Integer
        summary_table_row = 2
       
        ' Determine last row
        Dim last_row As Long
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Add column for ticker
        ws.Columns("I").Insert
       
        ' Name new column (ticker)
        ws.Cells(1, 9).Value = "ticker"
       
        ' Add total stock volume column
        ws.Columns("J").Insert
       
        ' Name new column (total_stock_volume)
        ws.Cells(1, 10).Value = "total_stock_volume"
       
        ' Add quarterly change column
        ws.Columns("K").Insert
       
        ' Name quarterly change column
        ws.Cells(1, 11).Value = "quarter_change"
       
        ' Add percentage change column
        ws.Columns("L").Insert
       
        ' Name percentage change column
        ws.Cells(1, 12).Value = "percent_change"
              
        ' autofit to display data
        ws.Columns("A:Q").AutoFit

        ' Loop through all tickers
        For i = 2 To last_row
           
            ' Check if we are still in the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           
                ' Set ticker name
                ticker = ws.Cells(i, 1).Value
           
                ' Set total stock volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
               
                ' Quarterly change
                Dim quarter_change As Double
                quarter_change = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
                
                ' Percentage change
                Dim percent_change As Double
                
                If ws.Cells(i, 3).Value <> 0 Then
                    percent_change = quarter_change / ws.Cells(i, 3).Value
                
                Else
                    percent_change = 0
                
            End If

               ' set variables for "quater_chage" color
               Dim rng As Range
               Dim cell As Range
                
               ' set range of the "quarter_change" column
               Set rng = ws.Range("K2:K2" & ws.Cells(ws.Rows.Count, "K").End(xlUp).Row)
                
               For Each cell In rng
                    
                  ' Set cell color to red for negative numbers
                   If cell.Value < 0 Then
                       cell.Interior.Color = RGB(255, 0, 0)
                        
                   ' Set cell color to green for positive numbers
                   ElseIf cell.Value > 0 Then
                       cell.Interior.Color = RGB(0, 255, 0)
                    
                   End If
                
                Next cell
                
                ' Print the ticker in the summary table
                ws.Cells(summary_table_row, 9).Value = ticker
           
                ' Print total stock volume in the summary table
                ws.Cells(summary_table_row, 10).Value = total_stock_volume

                ' Print quarterly change
                ws.Cells(summary_table_row, 11).Value = quarter_change
               
                ' Print percentage change
                ws.Cells(summary_table_row, 12).Value = percent_change
                
                ' Set the number format of the percentage change column to percentage
                ws.Columns("L").NumberFormat = "0.00%"
           
                ' Add one to the summary table row
                summary_table_row = summary_table_row + 1
               
                ' Reset the total stock volume
                total_stock_volume = 0
           
            Else
           
                ' Add to the total stock volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
               
            End If
        ' setting for loops to fill cells with greatest%, smallest%, and total volume
        Next i
        
        Dim hightestPercent As Double
        hightestPercent = 0
        
        Dim lowPercent As Double
        lowPercent = 0
        
        Dim highestVol As Double
        highestVol = 0
        
        tickerName = ""
        
        Dim vol As Double
        
        Dim percent As Double
        
        For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
            percent = Cells(i, "L").Value
            vol = Cells(i, "J").Value
            
            If percent > hightestPercent Then
            
                hightestPercent = percent
                
                tickerName = Cells(i, "I")
                
                Range("N2").Value = "Greatest % increase"
        
                Range("O2").Value = tickerName
                
                Range("P2").Value = percent
                
                Range("P2").NumberFormat = "0.00%"
            
            End If
            
            If percent < lowPercent Then
            
                lowPercent = percent
                
                tickerName = Cells(i, "I")
                
                Range("N3").Value = "Greatest % decrease"
                
                Range("O3").Value = tickerName
                
                Range("P3").Value = percent
                
                Range("P3").NumberFormat = "0.00%"
                
            End If
            
            If vol > highestVol Then
             
             highestVol = vol
             
             tickerName = Cells(i, "I")
             
             Range("N4").Value = "Highest Volume"
             
             Range("O4").Value = tickerName
             
             Range("P4").Value = vol
             
             Range("P4").NumberFormat = "0.00"
            
            End If
            
        Cells(1, 15).Value = "ticker"
        
        Cells(1, 16).Value = "value"
        
        Next i
 
    Next ws
       
End Sub