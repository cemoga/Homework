Sub EachSheet()
    Dim Wsht as Worksheet                                       ' Used to loop every worksheet
    For Each Wsht In Worksheets
        wsht.Activate
        Call Summarize
    Next Wsht
End Sub

Sub Summarize()

    Dim i, j As Long
    Dim LastRow As Long                                         ' Calculates the last row in column 1
    Dim Counter As Long                                         ' Counter for condition cell value equals next cell value
    Dim Ticker as String                                        ' It store the individual Ticker
    Dim TotalVolume As Double                                   ' It stores the total volume for Ticker
    Dim OpenPrice, ClosePrice As Double                         ' Variable that stores Opening for min and Closing for Max dates
    Dim YearlyChange, PercentChange as Double                   ' It stores the Changes in price
    Dim MaxPercIncrease, MaxPercDecrease, MaxVolume as Double    ' They store the maximun percentage increases and decreases
    Dim MaxPercIncrTicker, MaxPercDecrTicker, MaxVolumeTicker as String ' Stores the Ticker for maximun percentage increases, decreases and Max Volume

    Range("I1").Value = "Ticker"                                ' Header Titles for the results
    Range("J1").Value = "Yearly Change"                         ' Header Titles for the results
    Range("K1").Value = "Percent Change"                        ' Header Titles for the results
    Range("L1").Value = "Total Stock Volume"                    ' Header Titles for the results
'   Range("M1") = "Open Price"                                  ' Header Titles for the results
'   Range("N1") = "Close Price"                                 ' Header Titles for the results
    Range("O2").Value = "Greatest % Increase"                   ' Header Titles for the results
    Range("O3").Value = "Greatest % Decrease"                   ' Header Titles for the results
    Range("O4").Value = "Greatest Total Volume"                 ' Header Titles for the results
    Range("P1").Value = "Ticker"                                ' Header Titles for the results
    Range("Q1").Value = "Value"                                 ' Header Titles for the results

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row  ' Calculates the last row in column 1

    MaxPercIncrease = 0                                         ' Sets the initial value to challenge the data
    MaxPercDecrease = 1000                                       ' Sets the initial value to challenge the data
    MaxVolume = 0                                               ' Sets the initial value to challenge the data

    Counter = 2                                   ' It starts at 2, because row 1 contains headers                                    
     For i = 2 To LastRow                                           ' Loops through all data
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then          ' Cell value differs from next cell value
            Ticker = Cells(i, 1).Value                              ' It store the value of individual Tickers
            TotalVolume = TotalVolume + Cells(i, 7)                 ' Volume last data of an specific Ticker
            If TotalVolume > MaxVolume  then                        ' This part tests and stores the
                MaxVolume = TotalVolume                             ' Maximun Volume and its ticker
                MaxVolumeTicker = Ticker
            End If
            Cells(Counter, 9).Value = Ticker                        ' It lists individual Tickers 
            Cells(Counter, 12).Value = TotalVolume                  ' Prints the total volume cumulated
            ClosePrice = Cells(i, 6).Value                          ' Saves the Close Price when next data is different
'           Cells(Counter, 14).Value = ClosePrice                   ' Prints Close Price
'           Cells(Counter, 13).Value = OpenPrice                    ' Prints Open Price
            YearlyChange = ClosePrice - OpenPrice   
            Cells(Counter, 10).Value = YearlyChange                 ' It prints the Change in price in front of each ticker name
            if OpenPrice = 0 then                                   ' Division by 0 is indetermined
                PercentChange = 0
            Else
                PercentChange = ClosePrice / OpenPrice - 1 
            End If
            Cells(Counter, 11).Value = PercentChange                ' It prints the Change in price in front of each ticker name
            If YearlyChange <0 then 
                Cells(Counter, 10).Interior.ColorIndex = 3          ' Colors Red for negative changes
            ElseIf YearlyChange >=0 then 
                Cells(Counter, 10).Interior.ColorIndex = 4          ' Colors Green for Positive Changes
            End If
            Counter = Counter + 1                                   ' Keeps track of the number of Ticker names
            TotalVolume = 0                                         ' Resets Total Volume for Next Ticker
        Elseif Cells(i, 1).Value <> Cells(i - 1, 1).Value Then      ' Cell value differs from previous cell value
            OpenPrice = Cells(i, 3).Value
            TotalVolume = TotalVolume + Cells(i, 7)                 ' Adds volume for the Ticker
            If TotalVolume > MaxVolume  then                        ' This part tests and stores the
                MaxVolume = TotalVolume                             ' Maximun Volume and its ticker
                MaxVolumeTicker = Ticker
            End If
        Else
            TotalVolume = TotalVolume + Cells(i, 7)                 ' Adds volume for the Ticker
            If TotalVolume > MaxVolume  then                        ' This part tests and stores the
                MaxVolume = TotalVolume                             ' Maximun Volume and its ticker
                MaxVolumeTicker = Ticker
            End If
        End If    

        If PercentChange > MaxPercIncrease  then                    ' This part tests and stores the
            MaxPercIncrease = PercentChange                         ' Maximun Percentage Increase and its ticker
            MaxPercIncrTicker = Ticker
        End If

        If PercentChange < MaxPercDecrease  then                    ' This part tests and stores the
            MaxPercDecrease = PercentChange                         ' Maximun Percentage Decrease and its ticker
            MaxPercDecrTicker = Ticker
        End If

    Next i

    Range("Q2").Value = MaxPercIncrease                               ' Prints the Max Percentage Increase
    Range("P2").Value = MaxPercIncrTicker                             ' Prints the Ticker with the Max Percentage Increase

    Range("Q3").Value = MaxPercDecrease                               ' Prints the Max Percentage Decrease
    Range("P3").Value = MaxPercDecrTicker                             ' Prints the Ticker with the Max Percentage Decrease
 
    Range("Q4").Value = MaxVolume                                     ' Prints the Max Volume
    Range("P4").Value = MaxVolumeTicker                               ' Prints the Ticker with the Max Volume

    ActiveSheet.UsedRange.EntireColumn.AutoFit                        ' It auto resizes all the columns in the sheet to make the data fit               
    Range("A1").Select
End Sub