Option Explicit

'Return the openning price based on the range for a given ticker
Function GetYearOpenningPrice(ByRef tRange As Range, ByRef dRange As Range) As Double

    Dim openningPrice As Double
    Dim minDate As Variant

    minDate = Application.WorksheetFunction.Min(dRange)
    openningPrice = Application.WorksheetFunction.Index(tRange, Application.WorksheetFunction.Match(minDate, dRange, 0), 3)
    GetYearOpenningPrice = openningPrice
    
End Function

'Return the closing price based on the range for a given ticker
Function GetYearClosingPrice(ByRef tRange As Range, ByRef dRange As Range) As Double

    Dim maxDate As Variant
    maxDate = Application.WorksheetFunction.Max(dRange)
    GetYearClosingPrice = Application.WorksheetFunction.Index(tRange, Application.WorksheetFunction.Match(maxDate, dRange, 0), 6)
    
End Function

'Return the yearly change between the openningPrice and closingPrice
Function GetYearlyChange(ByVal openningPrice As Double, ByVal closingPrice As Double) As Double
        
        GetYearlyChange = closingPrice - openningPrice
    
End Function

'Return the yearly change as a percentage, ensuring no divide by zero errors occur via an openning price of 0
Function GetPercentageChange(ByVal yearlyChange As Double, ByVal openningPrice As Double) As Double

    'if the openning price is 0, return 0, otheriwse calculate the percentage
    If openningPrice = 0 Then
        GetPercentageChange = 0
    Else
        GetPercentageChange = (yearlyChange / openningPrice)
    End If
    
End Function

'Return the colour index based on the yearlyChange
Function GetFormattingColourIndex(ByVal yearlyChange As Double) As Integer

    'if the change is positive return green.
    If yearlyChange > 0 Then
        GetFormattingColourIndex = 4
    'If negative, return red
    ElseIf yearlyChange < 0 Then
        GetFormattingColourIndex = 3
    'if 0, return yellow
    Else
        GetFormattingColourIndex = 6
    End If
    
End Function

'display and format the calculated stats
Sub DisplayWorkSheetResults(ByRef ws As Worksheet, ByRef data() As Variant)

    Dim lastCol, startingDisplayCol, tickerRow, i As Integer
    
    tickerRow = 2
    
    'get the last column for the existing data
    lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'add a spacer column
    startingDisplayCol = lastCol + 2
    
    'build the results grid header
    With ws.Cells(1, startingDisplayCol)
        .Value = "Ticker"
        .EntireColumn.HorizontalAlignment = xlLeft
    End With
        
    With ws.Cells(1, startingDisplayCol + 1)
        .Value = "Yearly Change"
    End With
        
    With ws.Cells(1, startingDisplayCol + 2)
        .Value = "Percent Change"
        .EntireColumn.NumberFormat = "0.00%"
    End With
    
    With ws.Cells(1, startingDisplayCol + 3)
        .Value = "Total Stock Volume"
    End With
  
    'process the data array onto the sheet
    For i = LBound(data) To UBound(data)
        ws.Cells(tickerRow, startingDisplayCol).Value = data(i)(0)
      
        With ws.Cells(tickerRow, startingDisplayCol + 1)
            .Value = data(i)(1)
            .Interior.ColorIndex = GetFormattingColourIndex(.Value)
         End With
         
        ws.Cells(tickerRow, startingDisplayCol + 2).Value = data(i)(2)
        ws.Cells(tickerRow, startingDisplayCol + 3).Value = data(i)(3)
    
        tickerRow = tickerRow + 1
    Next i
    
    ws.Cells(1, startingDisplayCol).EntireColumn.AutoFit
    ws.Cells(1, startingDisplayCol + 1).EntireColumn.AutoFit
    ws.Cells(1, startingDisplayCol + 2).EntireColumn.AutoFit
    ws.Cells(1, startingDisplayCol + 3).EntireColumn.AutoFit

End Sub

'BONUS - Get the stock with the greatest % increase
Function GetGreatestIncrease(ByRef statsRange As Range, ByRef percentageRange As Range) As Variant()

    Dim arTickerRecord(2) As Variant
    arTickerRecord(0) = "Greatest % Increase"
    arTickerRecord(2) = Application.WorksheetFunction.Max(statsRange.Columns(3))
    arTickerRecord(1) = Application.WorksheetFunction.Index(statsRange, Application.WorksheetFunction.Match(arTickerRecord(2), percentageRange, 0), 1)
    GetGreatestIncrease = arTickerRecord

End Function

'BONUS - Get the stock with the greatest % decrease
Function GetGreatestDecrease(ByRef statsRange As Range, ByRef percentageRange As Range) As Variant()

    Dim arTickerRecord(2) As Variant
    arTickerRecord(0) = "Greatest % Decrease"
    arTickerRecord(2) = Application.WorksheetFunction.Min(statsRange.Columns(3))
    arTickerRecord(1) = Application.WorksheetFunction.Index(statsRange, Application.WorksheetFunction.Match(arTickerRecord(2), percentageRange, 0), 1)
    GetGreatestDecrease = arTickerRecord
    
End Function

'BONUS - Get the stock with the greatest total volume
Function GetGreatestTotalVolume(ByRef statsRange As Range, ByRef volumeRange As Range) As Variant()

    Dim arTickerRecord(2) As Variant
    arTickerRecord(0) = "Greatest Total Volume"
    arTickerRecord(2) = Application.WorksheetFunction.Max(statsRange.Columns(4))
    arTickerRecord(1) = Application.WorksheetFunction.Index(statsRange, Application.WorksheetFunction.Match(arTickerRecord(2), volumeRange, 0), 1)
    GetGreatestTotalVolume = arTickerRecord

End Function

'BONUS - Show the Bonus Results
Sub DisplayBonusResults(ByRef ws As Worksheet, ByRef data() As Variant)

    Dim lastCol, startingDisplayCol, tickerRow, i As Integer
    
    tickerRow = 2
    
    'get the last column for the existing data
    lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'add 2 spacer columns
    startingDisplayCol = lastCol + 3
    
    'build the results grid header
    With ws.Cells(1, startingDisplayCol)
        .Value = ""
        .EntireColumn.HorizontalAlignment = xlLeft
    End With
        
    With ws.Cells(1, startingDisplayCol + 1)
        .Value = "Ticker"
        .EntireColumn.HorizontalAlignment = xlLeft
    End With
        
    With ws.Cells(1, startingDisplayCol + 2)
        .Value = "Value"
    End With
    
    'process the data array onto the sheet
    For i = LBound(data) To UBound(data)
        ws.Cells(tickerRow, startingDisplayCol).Value = data(i)(0)
        ws.Cells(tickerRow, startingDisplayCol + 1).Value = data(i)(1)
        ws.Cells(tickerRow, startingDisplayCol + 2).Value = data(i)(2)
        
        If tickerRow = 2 Or tickerRow = 3 Then
            ws.Cells(tickerRow, startingDisplayCol + 2).NumberFormat = "0.00%"
        End If
    
        tickerRow = tickerRow + 1
    Next i
    
    ws.Cells(1, startingDisplayCol).EntireColumn.AutoFit
    ws.Cells(1, startingDisplayCol + 1).EntireColumn.AutoFit
    ws.Cells(1, startingDisplayCol + 2).EntireColumn.AutoFit

End Sub

'BONUS - Process the bonus stats for display
Function ProcessBonusStats(ByRef ws As Worksheet) As Variant()

    Dim lastCol, lastRow, startingDisplayCol, tickerRow, i As Integer
    Dim statsRange As Range
    Dim percentageRange As Range
    Dim volumeRange As Range
    Dim arData(2) As Variant

    tickerRow = 2
    
    'get the last column for the existing data
    lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(Rows.Count, lastCol).End(xlUp).Row
      
    'get the range for the stats
    Set statsRange = ws.Range(ws.Cells(tickerRow, lastCol - 3), ws.Cells(lastRow, lastCol))
    Set percentageRange = statsRange.Columns(3)
    Set volumeRange = statsRange.Columns(4)
    
    arData(0) = GetGreatestIncrease(statsRange, percentageRange)
    arData(1) = GetGreatestDecrease(statsRange, percentageRange)
    arData(2) = GetGreatestTotalVolume(statsRange, volumeRange)
    
    ProcessBonusStats = arData
    
End Function

'Main sub to run the program
Sub KickOf()
    Dim rw As Range
    Dim wrkSheetRange As Range
    Dim tickerRange As Range
    Dim dateRange As Range
    Dim wrkSheet As Worksheet
    Dim tickerSummaryRow As Integer
    Dim strDate As String
    Dim ticker, nextTicker As String
    Dim dblOpenningPrice As Double
    Dim dblClosingPrice As Double
    Dim dblYearlyChange As Double
    Dim dblVolume As Double
    Dim dblTotalVolume As Double
    Dim rangeCount As Double
    Dim tickerStartRange As Double
    Dim tickerEndRange As Double
    Dim summaryData() As Variant
    Dim summaryRecord(3) As Variant
    Dim isFirstRowheader As Boolean
        
    For Each wrkSheet In ActiveWorkbook.Worksheets
        
        'activate the worksheet so its ready to be updated
        wrkSheet.Activate
        
        'ensure the worksheet is sorted with all the tickers together, by date in ascending order
        With wrkSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wrkSheet.Range("a1"), Order:=xlAscending
            .SortFields.Add Key:=wrkSheet.Range("b1"), Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
       
        'set the range to use and omit the first row because its headers
        Set wrkSheetRange = wrkSheet.UsedRange
        Set wrkSheetRange = wrkSheetRange.Offset(1, 0).Resize(wrkSheetRange.Rows.Count - 1, wrkSheetRange.Columns.Count)
        
        rangeCount = wrkSheetRange.Rows.Count
    
        'reset the summary row counter
        tickerSummaryRow = 0
        
        'init tickerStartRange
        tickerStartRange = 2
                        
        'process the records to calculate the yearly change, the percent change and total stock volume
        For Each rw In wrkSheetRange.Rows
            ticker = wrkSheet.Cells(rw.Row, 1).Value
            strDate = wrkSheet.Cells(rw.Row, 2).Value
            dblVolume = wrkSheet.Cells(rw.Row, 7).Value
            nextTicker = wrkSheet.Cells(rw.Row + 1, 1).Value
            
            If ticker = nextTicker Then
                dblTotalVolume = dblTotalVolume + dblVolume
                
            ElseIf ticker <> "" Then
                'Get the row number as the last number in the current ticker set
                tickerEndRange = rw.Row
                             
                'update the total volume with the last volume for the given ticker
                dblTotalVolume = dblTotalVolume + dblVolume
                
                'create a range for just that ticker and for the date for openning and closing calculations
                Set tickerRange = wrkSheetRange.Range("A" & tickerStartRange & ":G" & tickerEndRange).Offset(-1, 0)
                Set dateRange = wrkSheetRange.Range("B" & tickerStartRange & ":B" & tickerEndRange).Offset(-1, 0)
                              
                'Debug.Print "dateRange set as : " + ("B" & tickerStartRange & ":B" & tickerEndRange)
                'Debug.Print "dateRange address is : " + dateRange.Address
                
                'update the summary array size
                ReDim Preserve summaryData(tickerSummaryRow)

                'get the metrics for stats
                dblOpenningPrice = GetYearOpenningPrice(tickerRange, dateRange)
                dblClosingPrice = GetYearClosingPrice(tickerRange, dateRange)
                dblYearlyChange = GetYearlyChange(dblOpenningPrice, dblClosingPrice)
                
                summaryRecord(0) = ticker
                summaryRecord(1) = dblYearlyChange
                summaryRecord(2) = GetPercentageChange(dblYearlyChange, dblOpenningPrice)
                summaryRecord(3) = dblTotalVolume
                summaryData(tickerSummaryRow) = summaryRecord
               
               'reset vars
                dblOpenningPrice = 0
                dblClosingPrice = 0
                dblYearlyChange = 0
                dblTotalVolume = 0
                
                'update the ticker start range - add 1 to the row and set it for the first number in the next ticker set
                tickerStartRange = rw.Row + 1
                
                'update the ticker summary row
                tickerSummaryRow = tickerSummaryRow + 1
            End If
            
        Next rw
                
         'show the results for the sheet
        DisplayWorkSheetResults wrkSheet, summaryData
               
        'BONUS - show the bonus results for the sheet
        DisplayBonusResults wrkSheet, ProcessBonusStats(wrkSheet)
        
        'clean up for the next sheet
        Erase summaryRecord()
        Erase summaryData()
        ReDim summaryData(0)
        
    Next wrkSheet
    
End Sub
