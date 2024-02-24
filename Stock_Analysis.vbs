Sub Stock_Analysis():

    ' counter is used to position the value of the list of different ticker
    Dim counter As Integer
    ' declare a variable that will store the ticket name
    Dim ticketname As String
    
    Dim volume As Double
    
    'function that finds the last value of the table
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    counter = 2
    volume = 0

    
    
    'this section just set the title of each column
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"



    ' this loops throu all of the ticker column and finds differences and then prints the diferent
    ' tickets in the i column
    
    'this loops also gets the total stock volume of the year per ticker
    
    For i = 2 To lastrow
    
        'this section compares if there is a difference between the actual value an the next one
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            ticketname = Cells(i, 1).Value
            
            'this sums the last ticker
            volume = volume + Cells(i, 7).Value
            
            Cells(counter, 9).Value = ticketname
            Cells(counter, 12).Value = volume
            
            counter = counter + 1
            
            'clear the variable when the new ticker it's counted
            volume = 0
            
        Else
        
            'this sums the volume of stocks of the rest of dates
            volume = volume + Cells(i, 7).Value
        
        End If
    Next i
    

    
    
    

End Sub