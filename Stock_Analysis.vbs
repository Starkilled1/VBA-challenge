Sub Stock_Analysis():


    ' This for loop iterates all worksheets and runs the script
    For Each ws In Worksheets
    
        ' counter is used to position the value of the list of different ticker
        Dim counter As Integer
        ' declare a variable that will store the ticket name
        Dim ticketname As String
        ' declare volume variable that stores the sum of stock
        Dim volume As Double
        ' declare variables where the price are going to be stored
        Dim firstPrice As Double
        Dim lastPrice As Double
        ' declare variables where the max and min values are going to be stored
        Dim maxchange As Double
        Dim minchange As Double
        Dim maxvol As Double
    
        maxchange = 0
        minchange = 0
        maxvol = 0

        'function that finds the last value of the table
        lastrow = Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastrowJ = Cells(ws.Rows.Count, 10).End(xlUp).Row
        counter = 2
        volume = 0
        firstPrice = 0
        lastPrice = 0

        
        'this section just set the title of each column and rows
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"


        ' this loops throu all of the ticker column and finds differences and then prints the diferent
        ' tickets in the i column
        ' this loops also gets the total stock volume of the year per ticker
        ' and the yearly change
        For i = 2 To lastrow
        
            'this section compares if there is a difference between the actual value an the next one
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                ticketname = ws.Cells(i, 1).Value
                
                'this sums the last ticker
                volume = volume + ws.Cells(i, 7).Value
                
                ' Since this conditions only happens when the ticker is about to change we take the value of the last price
                lastPrice = ws.Cells(i, 6).Value
                
                ws.Cells(counter, 9).Value = ticketname
                ws.Cells(counter, 12).Value = volume
                ws.Cells(counter, 10).Value = lastPrice - firstPrice
                ws.Cells(counter, 11).Value = ((lastPrice - firstPrice) / firstPrice)
                ws.Cells(counter, 11).NumberFormat = "0.00%"
                
                counter = counter + 1
                
                'clear the variable when the new ticker it's counted
                volume = 0
                
                
            Else
            
                'this sums the volume of stocks of the rest of dates
                volume = volume + ws.Cells(i, 7).Value
            
            End If
            
            
            ' this condition evaluates all dates, when the current date is bigger than the next
            ' this means we got the another ticker so we store the first value at the openining
            If ws.Cells(i, 2).Value > ws.Cells(i + 1, 2).Value Then
            
                firstPrice = ws.Cells(i + 1, 3).Value
            
            'since this condition applies for ticker change, the first ticker is addressed in this condition
            
            ElseIf i = 2 Then
            
                firstPrice = ws.Range("C2").Value
            
            End If
            
            
        Next i
        
        For i = 2 To lastrowJ
        
            'conditional formating to change the cell color to red if negative and green if positive
            If ws.Cells(i, 10).Value > 0 Then
                
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
                
            ElseIf ws.Cells(i, 10).Value = 0 Then
                
                ws.Cells(i, 10).Interior.Color = RGB(125, 126, 133)
                
            Else
                
                ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
                    
            End If
            
            
            ' Conditions that loops throu the greatest percentage change increase and decrease and also biggest volume of stock
            If ws.Cells(i, 11).Value > maxchange Then
            
                maxchange = ws.Cells(i, 11).Value
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                ws.Range("Q2").NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 11).Value < minchange Then
            
                minchange = ws.Cells(i, 11).Value
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                ws.Range("Q3").NumberFormat = "0.00%"
                
            ElseIf ws.Cells(i, 12).Value > maxvol Then
            
                maxvol = ws.Cells(i, 12).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    Next ws
    
End Sub