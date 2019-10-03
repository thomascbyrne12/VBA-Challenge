Sub VBAStocks():

    'Defining Variables
    Dim ticker As String
    Dim year As Long
    Dim oppri As Double
    Dim clopri As Double
    Dim a As Long
    Dim LastRow As Long
    Dim total As Long
    Dim i As Long
    Dim greattotal As Long
    Dim totalticker As String
    Dim greatpercent As Double
    Dim greatticker As String
    Dim lowpercent As Double
    Dim lowticker As String
    
    'Bonus Output Table
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    'For j = 1 To 3
    
        'Selects Worksheet
        Sheets(1).Select
    
        'Set-up for Output Table
        a = 2
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
        Range("K2:K" & LastRow).NumberFormat = "0.00%"
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"
    
        'Initializing Variable
        ticker = Cells(2, 1)
    
        'Determining End of Loop
        LastRow = Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow
        
            If Cells(i - 1, 1) <> ticker Then
            
                'Finds Opening Price
                oppri = Cells(i, 3)
            
                'Starts Total Volume
                'total = Cells(i, 7)
        
            ElseIf Cells(i + 1, 1) <> ticker Then
            
                'Finds Closing Price
                clopri = Cells(i, 6)
                'total = total + Cells(i, 7)
                
                'Prints Ticker and Price Change
                Cells(a, 9) = ticker
                Cells(a, 10) = clopri - oppri
            
                'Formats Cell with Colors
                If Cells(a, 10) < 0 Then
                    Cells(a, 10).Select
                    Selection.Interior.ColorIndex = 3
                Else
                    Cells(a, 10).Select
                    Selection.Interior.ColorIndex = 4
                End If
            
                'Prints Percent and Records Min/Max Values
                Cells(a, 11) = Cells(a, 10) / oppri
                If greatpercent < Cells(a, 11) Then
                    greatpercent = Cells(a, 11)
                    greatticker = ticker
                ElseIf lowpercent > Cells(a, 11) Then
                    lowpercent = Cells(a, 11)
                    lowticker = ticker
                End If
            
                'Prints Total and Records Great Value
                Cells(a, 12) = total
                If greattotal < total Then
                    greattotal = total
                    totalticker = total
                End If
            
                'Iterates Ticker Count and Changes Ticker
                ticker = Cells(i + 1, 1)
                total = 0
                a = a + 1
            
            Else
        
                'total = total + Cells(i, 7)
            
            End If
        
        Next i
    
    'Next j
    
    'Selects Initial Worksheet
    Sheets(1).Select
    
    'Prints Bonus Results
    Cells(2, 16) = greatticker
    Cells(2, 17) = greatpercent
    Cells(3, 16) = lowticker
    Cells(3, 17) = lowpercent
    Cells(4, 16) = totalticker
    Cells(4, 17) = greattotal
    
End Sub

