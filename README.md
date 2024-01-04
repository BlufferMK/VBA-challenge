# VBA-challenge
Module 2 VBA Challenge
      Sub WorksheetLoop()
        'this  looping through worksheet script is from     https://support.microsoft.com/en-au/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
        
         Dim WS_Count As Integer
         Dim J As Integer
         Dim count As Integer
         Dim Tick_Num As Integer
        'count is used to keep track of the number of unique ticker symbols
        'Tick_Num counts up the number of stocks to loop through for the summary table
        'J is for looping through worksheets
        'WS_Count is for counting worksheets to loop through

         ' Set WS_Count equal to the number of worksheets in the active workbook.
         WS_Count = ActiveWorkbook.Worksheets.count

        'Declare a variable to store the row count and variables for highest percent change, biggest decrease, highest volume
        Dim rowCount As Long
        Dim Vol As Double
        Dim Highest As Double
        Dim Lowest As Double
        Dim High_Vol As Double
        'Vol is for summing each volume
        'Highest is for storing the highest percent change for the sheet
        'Lowest is for storing the highest percent drop for the sheet
        'High_Vol is for storing the maximum volume traded amount
        
        'List is the array that will include the ticker symbols
         Dim List(1000) As String
         Dim OpenP As Double
         'for storing the opening price for the year for each stock
         Dim CloseP As Double
         'for storing the closing price for the year for each stock
       
        Dim ws As Worksheet
        
       
       
         ' Begin the loop through SHEETS.
         For J = 1 To WS_Count

            'subroutine for each sheet: check the ticker symbols for unique values and create an array with the symbols and show them in column I

             Set ws = ActiveSheet
             'This works only on the active sheet
    
             'Count the rows in the used range of the worksheet
             rowCount = ws.UsedRange.Rows.count
    

             'add titles to the first row
             Cells(1, 9).Value = "Ticker"
             Cells(1, 10).Value = "Yearly Change"
             Cells(1, 11).Value = "Percent Change"
             Cells(1, 12).Value = "Total Stock Volume"
             Cells(1, 14).Value = "Year Open"
             Cells(1, 15).Value = "Year Close"

             'Now the looping code for checking through the sheet
             'count is initialized to zero
             count = 0

             For I = 2 To rowCount

                 'Check if the current cell is NOT the same as the previous ticker symbol
                 If Cells(I, 1).Value <> Cells(I - 1, 1).Value Then
                 'Add the ticker symbol to our list: the count is used for indexing
                     List(count) = Cells(I, 1).Value
                 'capture the opening price for the year
                     OpenP = Cells(I, 3).Value
                 'set the value in column N to the opening price for the year
                     Cells(count + 2, 14).Value = OpenP
            
                 'put the ticker symbol in our list of ticker symbols in column I
                     Cells(count + 2, 9).Value = List(count)

                 'capture the year end price and subtract the opening price and add the yearly change value to the
                 'correct cell.  The final row isn't collected, that's further below

                    If count > 0 Then
                 'set the value in column O to the end price for the year but NOT the first time that the Previous ticker symbol doesnt match
                        Cells(count + 1, 15).Value = Cells(I - 1, 6).Value
                 'calculate and show the amount of change
                        Cells(count + 1, 10).Value = Cells(count + 1, 15).Value - Cells(count + 1, 14).Value

                 'Caluclate and show the percent change for the year

                        Cells(count + 1, 11).Value = Cells(count + 1, 10).Value / Cells(count + 1, 14).Value
                        Cells(count + 1, 11).NumberFormat = "0.00%"
                     End If
            
                'increase the counter by 1 so that the next string added to the list is indexed correctly
                     count = count + 1
                 End If

        
                 'capture the final value for the year of last alphabetical stock ticker and do calculations
                 If I = rowCount Then
                     Cells(count + 1, 15).Value = Cells(I - 1, 6).Value
                 'set the value in column O to the end price for the year
                     Cells(count + 1, 15).Value = Cells(I - 1, 6).Value

                 'calculate the yearly change
                     Cells(count + 1, 10).Value = Cells(count + 1, 15).Value - Cells(count + 1, 14).Value
   
                 'calculate and show percent change for the year
                     Cells(count + 1, 11).Value = Cells(count + 1, 10).Value / Cells(count + 1, 14).Value
                     Cells(count + 1, 11).NumberFormat = "0.00%"
                 End If

            Next I
           
           'at this point, all of the ticker symbols, and open and closing prices have been found and set
           'counting volume subroutine
    
            Vol = 0
            count = 2
            I = 2
           'For each row of the table…loop through the volumes and sum them for an individual ticker symbol
            For I = 2 To rowCount
               'check to see if the ticker value is the same as the CURRENT ticker value
                If Cells(I, 1).Value = Cells(count, 9).Value Then
                'if it's the current ticker value, add up the volume
                    Vol = Vol + Cells(I, 7).Value
                
                'the else happens when looping through the rows reaches the NEXT ticker symbol
                Else
                  'if it's NOT the current ticker value, set the total volume column value
                    Cells(count, 12).Value = Vol
                   'restart summing the volume with the first day of the year for the new stock
                    Vol = Cells(I, 7).Value
                   'increase counter to properly index the sums in our new table
                    count = count + 1

                End If
            Next I

            'Finishes the column with the final total volume since there are no more rows and the else wasn't triggered
             Cells(count, 12).Value = Vol
 
       'Use the current count to identify the number of tickers on this sheet
        Tick_Num = count
    

       'format interior colors based on positive or negative change
    
        For I = 2 To rowCount
           'negative values get shaded in red
            If Cells(I, 10).Value < 0 Then
                Cells(I, 10).Interior.ColorIndex = 3
           'positive values get shaded green
            ElseIf Cells(I, 10).Value > 0 Then
                Cells(I, 10).Interior.ColorIndex = 4
            End If

        Next I
    
        'script to build the comparison table showing greatest change and volume
        Range("Q2").Value = "Greatest % Change"
        Range("Q3").Value = "Greatest Neg. % Change"
        Range("Q4").Value = "Greatest Volume"
    
        Range("R1").Value = "Ticker"
        Range("S1").Value = "Value"
        
    
        Highest = 0
        Lowest = 0
        High_Vol = 0
    
       'loop through new table to find maximum change, minimum change, and volume
        For I = 2 To Tick_Num - 1
    
           'highest neg. change
            If Cells(I, 11).Value < Lowest Then
                Lowest = Cells(I, 11).Value
               'set the ticker in the new table
                Range("R3").Value = Cells(I, 9).Value
            
            End If
           'highest change
            If Cells(I, 11).Value > Highest Then
                Highest = Cells(I, 11).Value
                Range("R2").Value = Cells(I, 9).Value
            
            End If
            'highest volume
            If Cells(I, 12).Value > High_Vol Then
                High_Vol = Cells(I, 12).Value
                Range("R4").Value = Cells(I, 9).Value
            End If
        
        Next I
    
       'show values in the table
        Range("S2").Value = Highest
        Range("S2").NumberFormat = "0.00%"
        Range("S3").Value = Lowest
        Range("S3").NumberFormat = "0.00%"
        Range("S4").Value = High_Vol


    Next J

End Sub
