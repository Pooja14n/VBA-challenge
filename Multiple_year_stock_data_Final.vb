Sub Stock_data():

    
    ' loop through all the worksheets
    For Each ws In Worksheets
    
    
    ' use lastrow variable to count the number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    
    ' assign names to columns/rows where the results are to be stored
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    
    
    
    ' Create variables to hold the ticker name, total stock volume, opening price, closing price, yearly change, and percent change.
       
    Dim i As Long
    Dim j As Long
    Dim ticker_name As String
    Dim total_stock_volume As Double
    Dim opening As Double
    Dim closing As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    
    ' set the initial value for the total_stock_volume to 0
    total_stock_volume = 0
    
 
    ' keep track of the location for each stock symbol/ticker in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    j = 2
    
    
    ' loop through all the stock symbols/ticker
    For i = 2 To lastrow
    
    
    
        ' check if it is still within the same stock symbol/ticker or not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
            ticker_name = ws.Cells(i, 1).Value
            
            opening = ws.Cells(j, 3).Value
            closing = ws.Cells(i, 6).Value
            j = i + 1
            
            
            ' calculate the yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
            yearly_change = closing - opening
            
               
            ' calculate the percent change using the calculated yearly change formula
            percent_change = (yearly_change / opening) * 100
            
            ' calculate the total stock volume for each stock symbol/ticker
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                      
                      
            ' assign the ticker name, yearly change, percent change, and total stock volume to columns J, K, L, and M respectively.
            ws.Range("J" & summary_table_row).Value = ticker_name
            
            ws.Range("K" & summary_table_row).Value = yearly_change
            ws.Range("K" & summary_table_row).Style = "Currency"
            
            ws.Range("L" & summary_table_row).Value = WorksheetFunction.Round(percent_change, 2) & "%"
                                  
            ws.Range("M" & summary_table_row).Value = total_stock_volume
            
            
            
            ' Conditional Formatting: change the cell colors pertaining to Yearly Change and Percent Change columns, to green if yearly change is positive, else, change the color to red.
            If (yearly_change > 0) Then
        
                ws.Range("K" & summary_table_row).Interior.ColorIndex = 4
                ws.Range("L" & summary_table_row).Interior.ColorIndex = 4
                
            Else
                
                ws.Range("K" & summary_table_row).Interior.ColorIndex = 3
                ws.Range("L" & summary_table_row).Interior.ColorIndex = 3
        
            End If
            
                                    
            
            
            ' add one to the summary table row
            summary_table_row = summary_table_row + 1
      
      
            ' reset yearly change and the total stock volume
            yearly_change = 0
            total_stock_volume = 0
            

            Else
            
                ' add to the total stock volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
   
        End If
          
               
               
                                                                        
    Next i
        
        
        
        ' find the ticker with the greatest increase, greatest decrease, and the greatest total stock volume.
        
        ws.Range("R2") = "%" & WorksheetFunction.Max(ws.Range("L2:L" & lastrow)) * 100
                
        ws.Range("Q2") = ws.Cells(WorksheetFunction.Match(ws.Range("R2"), ws.Range("L2:L" & lastrow), 0) + 1, 10).Value
    
        ws.Range("R3") = "%" & WorksheetFunction.Min(ws.Range("L2:L" & lastrow)) * 100
        
        ws.Range("Q3") = ws.Cells(WorksheetFunction.Match(ws.Range("R3"), ws.Range("L2:L" & lastrow), 0) + 1, 10).Value
        
        ws.Range("R4") = WorksheetFunction.Max(ws.Range("M2:M" & lastrow))
        
        ws.Range("Q4") = ws.Cells(WorksheetFunction.Match(ws.Range("R4"), ws.Range("M2:M" & lastrow), 0) + 1, 10).Value
        
        
        
          
    Next ws
    

End Sub



