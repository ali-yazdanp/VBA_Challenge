Attribute VB_Name = "Module1"
Sub MultipleYearStockData()
    
    'Define "ws" as Worksheet
    
    Dim ws As Worksheet
    
    'Looping through each of the worksheets
    
    For Each ws In Worksheets
    
    'Defining variables
    
    Dim Ticker As String
    Dim ALastRow As Long
    Dim KLastRow As Long
    Dim LastRowValue As Long
    Dim PreviousPrice As Long
    Dim FirstPartTableRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim TotalStockVolume As Double
    Dim GreatestTotalStockVolume As Double
   
  
    'Setting default values for variables that are defined as counters
    
    PreviousPrice = 2
    FirstPartTableRow = 2
    TotalStockVolume = 0
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalStockVolume = 0
    
    'Setting the Column Headers for task 1 and for the Table in the bonus part
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    

    'Finding the last row for column A
    
    ALastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Loop through rows to find requested values in part 1 of the challenge
    
    For i = 2 To ALastRow

         'Finding Total Stock Volume for each row
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
        'Populating he Ticker column whenever the Ticker name changes in Column A
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'If Ticker name for next rowchanges then set the ticker name in Coloumn I for each Row
            
            Ticker = ws.Cells(i, 1).Value
                
            'Print the Ticker name in the column I
            
            ws.Range("I" & FirstPartTableRow).Value = Ticker

        'Print Total Stock Volume in the column L

            ws.Range("L" & FirstPartTableRow).Value = TotalStockVolume
        
            'Reset Total Stock Volume to populate the next row
            TotalStockVolume = 0
                
            'Setting the Yearly Opening Price
            OpeningPrice = ws.Range("C" & PreviousPrice)
                
            'Setting the Yearly Closing Price
            ClosingPrice = ws.Range("F" & i)
                
            'Setting the Value of Yearly Change
            
            YearlyChange = ClosingPrice - OpeningPrice
            ws.Range("J" & FirstPartTableRow).Value = YearlyChange
            
            'Changing format of Column J "$"
            ws.Range("J" & FirstPartTableRow).NumberFormat = "$0.00"

            'Determining Percent Change
            'if Yearly Opening Price is 0, then Percent Change is 0, Otherwise, it is equal to (Yearly Change/Yearly Opening Price)
            
            If OpeningPrice = 0 Then
                PercentChange = 0
            
                Else
                
                YearlyOpen = ws.Range("C" & PreviousPrice)
                PercentChange = YearlyChange / OpeningPrice
                        
            End If
                
            'Populating Column K with "PercentChange"
            ws.Range("K" & FirstPartTableRow).Value = PercentChange
            
            'Changing format of Column K "%"
            ws.Range("K" & FirstPartTableRow).NumberFormat = "0.00%"
                
    
            'Conditional Formatting, Green for +ve values and Red for -ve values
            
            If ws.Range("J" & FirstPartTableRow).Value >= 0 Then
            ws.Range("J" & FirstPartTableRow).Interior.ColorIndex = 4
                    
                Else
                ws.Range("J" & FirstPartTableRow).Interior.ColorIndex = 3
                
            End If
            
            'Moving by 1 for the First part's Table's Row
            
            FirstPartTableRow = FirstPartTableRow + 1
              
            'Setting the Previous Price

            PreviousPrice = i + 1
                
        End If
                
        'Moving to next row

        Next i

        'Finding the last row for column K
        
        KLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Loop through rows for Bonus table
        
        For i = 2 To KLastRow
            
            'Greatest % Increase
            
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If

            'Greatest % Decrease
            
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
                    
            End If

            'Greatest Total Volume
            
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                    
            End If

            Next i
            
        'Changing format of Q2 and Q3  "%"
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws

End Sub



