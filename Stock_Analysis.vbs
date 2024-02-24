Sub Stock_Analysis():

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
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    lastrowJ = Cells(Rows.Count, 10).End(xlUp).Row
    counter = 2
    volume = 0
    firstPrice = 0
    lastPrice = 0

    
    'this section just set the title of each column and rows
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"


    ' this loops throu all of the ticker column and finds differences and then prints the diferent
    ' tickets in the i column
    ' this loops also gets the total stock volume of the year per ticker
    ' and the yearly change
    For i = 2 To lastrow
    
        'this section compares if there is a difference between the actual value an the next one
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            ticketname = Cells(i, 1).Value
            
            'this sums the last ticker
            volume = volume + Cells(i, 7).Value
            
            ' Since this conditions only happens when the ticker is about to change we take the value of the last price
            lastPrice = Cells(i, 6).Value
            
            Cells(counter, 9).Value = ticketname
            Cells(counter, 12).Value = volume
            Cells(counter, 10).Value = lastPrice - firstPrice
            Cells(counter, 11).Value = ((lastPrice - firstPrice) / firstPrice)
            Cells(counter, 11).NumberFormat = "0.00%"
            
            counter = counter + 1
            
            'clear the variable when the new ticker it's counted
            volume = 0
            
            
        Else
        
            'this sums the volume of stocks of the rest of dates
            volume = volume + Cells(i, 7).Value
        
        End If
        
        
        ' this condition evaluates all dates, when the current date is bigger than the next
        ' this means we got the another ticker so we store the first value at the openining
        If Cells(i, 2).Value > Cells(i + 1, 2).Value Then
        
            firstPrice = Cells(i + 1, 3).Value
        
        'since this condition applies for ticker change, the first ticker is addressed in this condition
        
        ElseIf i = 2 Then
        
            firstPrice = Range("C2").Value
        
        End If
        
        
    Next i
    
    For i = 2 To lastrowJ
    
        'conditional formating to change the cell color to red if negative and green if positive
        If Cells(i, 10).Value > 0 Then
            
            Cells(i, 10).Interior.Color = RGB(0, 255, 0)
            
        ElseIf Cells(i, 10).Value = 0 Then
            
            Cells(i, 10).Interior.Color = RGB(125, 126, 133)
            
        Else
            
            Cells(i, 10).Interior.Color = RGB(255, 0, 0)
                
        End If
        
        
        ' Conditions that loops throu the greatest percentage change increase and decrease and also biggest volume of stock
        If Cells(i, 11).Value > maxchange Then
        
            maxchange = Cells(i, 11).Value
            Range("P2").Value = Cells(i, 9).Value
            Range("Q2").Value = Cells(i, 11).Value
            Range("Q2").NumberFormat = "0.00%"
        
        ElseIf Cells(i, 11).Value < minchange Then
        
            minchange = Cells(i, 11).Value
            Range("P3").Value = Cells(i, 9).Value
            Range("Q3").Value = Cells(i, 11).Value
            Range("Q3").NumberFormat = "0.00%"
            
        ElseIf Cells(i, 12).Value > maxvol Then
        
            maxvol = Cells(i, 12).Value
            Range("P4").Value = Cells(i, 9).Value
            Range("Q4").Value = Cells(i, 12).Value
        
        End If
    
    Next i
    
End Sub