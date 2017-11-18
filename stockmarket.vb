Sub Stockmarket()

 Dim stockopen As Double
 Dim stockclose As Double
 Dim currentticker As Long
 Dim totalvolume As Double
 Dim lastRow As Long
 Dim ws As Worksheet
 
 
  '********************************************************************************
         
 ' loop throw all the worksheet
 
  For Each ws In ThisWorkbook.Worksheets
        
        'activate each worksheet at this workbook
        
        ws.Activate
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
         
        'For each worksheet Add the ticker, yearly change, percent change and total volume column's title
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly change"
        Cells(1, 11).Value = "Percent change"
        Cells(1, 12).Value = "Total Volume"
        
        'Initiate the variables
        
        stockopen = Cells(2, 3).Value
        currentticker = 2
        totalvolume = 0
        
        'check all the ticker
        
        For I = 2 To lastRow
            
            'If the ticker changed, need to calculate all new columns

            
            If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
               
               'Put the Ticker Symbol

                Cells(currentticker, 9).Value = Cells(I, 1).Value
               
               'calculate the Yearly change from what the stock opened the
               ' year at to what the closing price was.
               
                Cells(currentticker, 10).Value = stockclose - stockopen
               
               'conditional formatting that will highlight positive change
               'in green and negative chang in red
               
                If (Cells(currentticker, 10).Value > 0) Then
                    Cells(currentticker, 10).Interior.ColorIndex = 4
                   
                ElseIf Cells(currentticker, 10).Value < 0 Then
                       Cells(currentticker, 10).Interior.ColorIndex = 3
                   
                End If
               
                           
               'The percent change from the what it opened the year at to what it closed.

               If stockopen <> 0 Then
                    Cells(currentticker, 11).Value = Round(((100 * (stockclose - stockopen)) / stockopen), 2)
                    Cells(currentticker, 11).Style = "Percent"
               Else
                    Cells(currentticker, 11).Value = Round((100 * (stockclose - stockopen)), 2)
                    Cells(currentticker, 11).Style = "Percent"
                    
               End If
               
                Cells(currentticker, 12).Value = totalvolume
                totalvolume = 0
                stockopen = Cells(I + 1, 3).Value
                currentticker = currentticker + 1
                
               
            Else
                ' As long as the ticker the same, need to save the closing price of the ticker
                
                stockclose = Cells(I + 1, 6).Value
                
                'The total Volume of the stock

                totalvolume = totalvolume + Cells(I, 7).Value
                
            End If
            
        
        Next I
        
       
        
        '*****************************************************************************
         'Determines greatest increase, greatest decrease on percentage of change
         'and determine the greatest volume
         'find the right ticker related to each value
         
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest total volume"
        Cells(1, 17).Value = "Ticker"
        Cells(1, 18).Value = "Value"
        
        greatestIncrease = Cells(2, 11).Value
        greatestDecrease = Cells(2, 11).Value
        greatestTvolume = Cells(2, 12).Value
         
        For k = 2 To currentticker
            
            ' find the greatest increase
            If Cells(k, 11).Value > greatestIncrease Then
            
            
                greatestIncrease = Cells(k, 11).Value
                greatestIncreasecell = k
                
                If Cells(k, 12).Value > greatestTvolume Then
                   
                   greatestTvolume = Cells(k, 12).Value
                   greatestTvolumecell = k
                End If
                
            ' find the greatest decrease
            ElseIf Cells(k, 11).Value < greatestDecrease Then
            
                greatestDecrease = Cells(k, 11).Value
                greatestDecreasecell = k
                
                If Cells(k, 12).Value > greatestTvolume Then
                   
                   greatestTvolume = Cells(k, 12).Value
                   greatestTvolumecell = k
                End If
            
            
            ElseIf Cells(k, 12).Value > greatestTvolume Then
                   Cells(4, 17).Value = Cells(k, 9).Value
                   greatestTvolume = Cells(k, 12).Value
                   greatestTvolumecell = k
            
                
             Else
               k = k + 1
             End If
             
         Next k
         
         Cells(2, 18).Value = greatestIncrease
         Cells(3, 18).Value = greatestDecrease
         Cells(4, 18).Value = greatestTvolume
         Cells(2, 17).Value = Cells(greatestIncreasecell, 9).Value
         Cells(2, 17).Style = "Percent"
         Cells(3, 17).Value = Cells(greatestDecreasecell, 9).Value
         Cells(3, 17).Style = "Percent"
         Cells(4, 17).Value = Cells(greatestTvolumecell, 9).Value
            
             
        
            
   Next ws
    
End Sub
