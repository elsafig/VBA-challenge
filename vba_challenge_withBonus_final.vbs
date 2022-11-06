Sub final_VBAChallenge()


'Create variables to hold values for calculations
Dim openingPrice As Double

Dim closingPrice As Double
closingPrice = 0

Dim totalStockVolume As LongLong
totalStockVolume = 0

Dim yearlyChange As Double
yearlyChange = 0

'Create variables to hold current ticker id
Dim ticker As String

'create variables
Dim ws As Worksheet

'Cycle through worksheets in workbook
For Each ws In ThisWorkbook.Worksheets

'Print headers in  new columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
        
'Populate openingPrice variable with value in <open> col
openingPrice = ws.Cells(2, 3).Value

'Create variables to mark outputRow
Dim outputRow As Double
outputRow = 2

'Determine the Last Row of worksheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Create a For loop from currentRow to last row
    For currentRow = 2 To LastRow
    
    'Create variable for readability
    Dim nextRow As Double
    nextRow = currentRow + 1
    
        'check if value does not match previous using col A <ticker> as the conditional
        If ws.Cells(nextRow, 1).Value <> ws.Cells(currentRow, 1).Value Then
        
            'populate ticker variable and print to new ticker col
            ticker = ws.Cells(currentRow, 1).Value

            'populate closingPrice variable with value in <close> col
            closingPrice = ws.Cells(currentRow, 6).Value

            'populate ticker column with ticker variable
            ws.Cells(outputRow, 9).Value = ticker

            'calculate and populate Yearly Change by subtracting openingPrice - closingPrice
            yearlyChange = closingPrice - openingPrice
            ws.Cells(outputRow, 10).Value = yearlyChange

        ' create if loop to conditionally format Yearly Change so that if yearlyChange >= 0
                If yearlyChange >= 0 Then
                    'then columns is filled green,
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4

                    'else filled red
                    Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3

                    'end if loop
                    End If
                    
        'if loop to take into account that cannot divide by 0
                    If openingPrice > 0 Then
                    'calculate and populate Percent Change by dividing yearlyChange/openingPrice
                    ws.Cells(outputRow, 11).Value = yearlyChange / openingPrice
            'else populate with 0 end if
                    Else
                    ws.Cells(outputRow, 11).Value = 0
                End If

        'format output as percentage
            ws.Cells(outputRow, 11).NumberFormat = "0.00%"

            'add value in <vol> volume to totalStockVolume
            totalStockVolume = totalStockVolume + ws.Cells(currentRow, 7).Value

            'populate Total Stock Volume into column
            ws.Cells(outputRow, 12).Value = totalStockVolume

            'reset totalStockVolume = 0
            totalStockVolume = 0

            'repopulate with new opening price
            openingPrice = ws.Cells(nextRow, 3).Value

            'add to outputRow = outputRow +1
            outputRow = outputRow + 1

            'else
            Else

            'add value in <vol> volume to totalStockVolume
            totalStockVolume = totalStockVolume + ws.Cells(currentRow, 7).Value

        End If
        
    Next currentRow
    
Next ws

'-------------------------------------------------------------- BONUS CODE -------------------------------------------------------------
'create variables to hold values
Dim percentIncreaseTick As String
Dim percentIncreaseValue As Double
percentIncreaseValue = 0

Dim percentDecreaseTick As String
Dim percentDecreaseValue As Double
percentDecreaseValue = 0

Dim maxTotalVolTick As String
Dim maxTotalVolValue As Double

'Cycle through worksheets in workbook
For Each ws In ThisWorkbook.Worksheets

'Get last row variable
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Cycle through rows i in worksheets
    For i = 2 To LastRow
    
'testing if value trumps current variable value - if value in Cells(i, 11).Value > percentIncreaseValue then store in percentIncreaseValue and store Cells(i, 9).Value in percentIncreaseTick
        If ws.Cells(i, 11).Value > percentIncreaseValue Then
        percentIncreaseValue = ws.Cells(i, 11).Value
        percentIncreaseTick = ws.Cells(i, 9).Value
        'else if value in Cells(i, 11).Value < percentDecreaseValue then store in percentDecreaseValue and store Cells(i, 9).Value in percentDecreaseTick
         ElseIf ws.Cells(i, 11).Value < percentDecreaseValue Then
        percentDecreaseValue = ws.Cells(i, 11).Value
        percentDecreaseTick = ws.Cells(i, 9).Value
        Else
        End If
        
'testing if value trumps current variable value if value in Cells(i, 12).Value > maxTotalVolValue and store Cells(i, 9).Value in maxTotalVolTick
       If ws.Cells(i, 12).Value > maxTotalVolValue Then
       maxTotalVolValue = ws.Cells(i, 12).Value
       maxTotalVolTick = ws.Cells(i, 9).Value
       Else
       End If
        
    'end for loops
    Next i
    
'print headers in  new columns
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

    'populate Cells(2,16) ticker with percentIncreaseTick
ws.Cells(2, 16).Value = percentIncreaseTick

'populate Cells(2,17).Value = percentIncreaseValue
ws.Cells(2, 17).Value = percentIncreaseValue

'format cell as percentage
ws.Cells(2, 17).NumberFormat = "0.00%"

'populate Cells(3,16).Value=percentDecreaseTick
ws.Cells(3, 16).Value = percentDecreaseTick

'populate Cells(3,17).Value=percentDecreaseValue
ws.Cells(3, 17).Value = percentDecreaseValue

'format cell as percentage
ws.Cells(3, 17).NumberFormat = "0.00%"

'populate Cells(4,16).value = maxTotalVolTick
ws.Cells(4, 16).Value = maxTotalVolTick

'populate Cells(4,17).value =  maxTotalVolValue
ws.Cells(4, 17).Value = maxTotalVolValue

'reset values for next sheet
percentIncreaseValue = 0
percentDecreaseValue = 0
maxTotalVolValue = 0

Next ws



End Sub