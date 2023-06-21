Attribute VB_Name = "Module1"
Sub AddYearlySummaries()
    
    'Declare variables
    Dim sheet As Worksheet
    
    'Loop Through Every Sheet In Workbook
    For Each sheet In ThisWorkbook.Worksheets
        AddYearlySummaryColumns sheet
        AddGreatestColumns sheet
    Next sheet
    
End Sub
Sub AddYearlySummaryColumns(sheet As Worksheet)
    
    'Declare variables
    Dim lastRow As Long
    Dim currentTickerName As String
    Dim firstRowWithTickerName As Long
    Dim lastRowWithTickerName As Long
    Dim tickerSumColNum As Integer
    Dim yearlyChColNum As Integer
    Dim percentChColNum As Integer
    Dim totalStkColNum As Integer
    Dim lastRowTickerSumm As Long
    Dim totalStkRangeString As String
    
    
    'Initialize Column Numbers
    tickerSumColNum = 9
    yearlyChColNum = 10
    percentChColNum = 11
    totalStkColNum = 12
    
    
    'Add Column Headers
    sheet.Cells(1, tickerSumColNum).Value = "Ticker"
    sheet.Cells(1, yearlyChColNum).Value = "Yearly Change"
    sheet.Cells(1, percentChColNum).Value = "Percent Change"
    sheet.Cells(1, totalStkColNum).Value = "Total Stock Volume"
    
    'Autofit Column Headers
    sheet.Columns("I:L").AutoFit
    
    
    'Get Last Row of the Column
    lastRow = LastRowInColumn(sheet, 1)
    
    'Get First Ticker Name
    currentTickerName = sheet.Cells(2, 1).Value

    'Set firstRowWithTickerName
    firstRowWithTickerName = 2
    
        
    'Get Last Row With Ticker Name
    lastRowWithTickerName = LastRowOfString(sheet, firstRowWithTickerName, 1, currentTickerName)
    
    'Set the last populated row in Summary Table which is initially 1
    lastRowTickerSumm = 1
    
    'Loop To Create Every Row In Ticker Summary Table
    Do While True
        'Set Name in Summary Table
        sheet.Cells(lastRowTickerSumm + 1, tickerSumColNum).Value = currentTickerName
        
        'Change format of Yearly Change Cell
        sheet.Range("j" & lastRowTickerSumm + 1).NumberFormat = "#,##0.00"
        
        'Set Yearly Change
        sheet.Cells(lastRowTickerSumm + 1, yearlyChColNum).Value = GetYearlyChange(sheet, firstRowWithTickerName, lastRowWithTickerName)
        
        'Set Cell Background Color Of Yearly Change
        If sheet.Cells(lastRowTickerSumm + 1, yearlyChColNum).Value < 0 Then
            sheet.Cells(lastRowTickerSumm + 1, yearlyChColNum).Interior.ColorIndex = 3
        ElseIf sheet.Cells(lastRowTickerSumm + 1, yearlyChColNum).Value > 0 Then
            sheet.Cells(lastRowTickerSumm + 1, yearlyChColNum).Interior.ColorIndex = 4
        End If
        
        'Change format of Percent Change Cell
        sheet.Range("k" & lastRowTickerSumm + 1).NumberFormat = "0.00%"
        
        
        'Set Percent Change
        sheet.Cells(lastRowTickerSumm + 1, percentChColNum).Value = GetPercentChange(sheet, firstRowWithTickerName, sheet.Cells(lastRowTickerSumm + 1, yearlyChColNum).Value)
        
        'Get Total Stock Volumne
        totalStkRangeString = "=SUM(G" & firstRowWithTickerName & ":G" & lastRowWithTickerName & ")"
        sheet.Range("L" & lastRowTickerSumm + 1).Formula2 = totalStkRangeString
        
        'Set New firstRowWithTickerName
        firstRowWithTickerName = lastRowWithTickerName + 1
        
        'Test If End Of Worksheet Has Been Reached
        If firstRowWithTickerName > lastRow Then Exit Do
        
        'Set New currentTickerName
        currentTickerName = sheet.Cells(firstRowWithTickerName, 1).Value
        
        'Set New lastRowWithTickerName
        lastRowWithTickerName = LastRowOfString(sheet, firstRowWithTickerName, 1, currentTickerName)
        
        'Set New lastRowTickerSumm
        lastRowTickerSumm = lastRowTickerSumm + 1

     Loop
    
    
    
End Sub
Function LastRowInColumn(sheet As Worksheet, colNum As Integer) As Long

    'Last row statement from class
    LastRowInColumn = sheet.Cells(Rows.Count, colNum).End(xlUp).Row
    
End Function
Function LastRowOfString(sheet As Worksheet, startingRow As Long, colNum As Integer, str As String) As Long
    
    'Declare variables
    Dim lastRow As Long
    
    'Find Last Row In Column
    lastRow = LastRowInColumn(sheet, colNum)
    
    'Search for last row same str in it
    For currRow = startingRow To lastRow
        If sheet.Cells(currRow, colNum).Value <> str Then
            LastRowOfString = currRow - 1
            Exit Function
        End If
    Next currRow
    
    'The last row of column is the last row with string in it
    LastRowOfString = lastRow
    
End Function
Function GetYearlyChange(sheet As Worksheet, firstRowWithTickerName As Long, lastRowWithTickerName As Long) As Double
    
    'Declare variables
    Dim openingPrice As Double
    Dim closingPrice As Double
    
    'Assign opening and closing values for Ticker Name
    openingPrice = sheet.Cells(firstRowWithTickerName, 3).Value
    closingPrice = sheet.Cells(lastRowWithTickerName, 6).Value
    
    
    'Calculate And Return Value
    GetYearlyChange = closingPrice - openingPrice
    
End Function
Function GetPercentChange(sheet As Worksheet, firstRowWithTickerName As Long, yearlyChange As Double) As Double
    
    'Declare variables
    Dim openingPrice As Double
    
    openingPrice = sheet.Cells(firstRowWithTickerName, 3).Value
    'MsgBox yearlyChange & openingPrice
    
    'Calculate And Return Value
    GetPercentChange = yearlyChange / openingPrice
    
End Function
Sub AddGreatestColumns(sheet As Worksheet)
    
    'Declare variables
    Dim percentRange As String
    Dim totalStkRange As String
    Dim tickerSummaryRange As String
    Dim targetForXLookUp As String
    Dim lastRowOfPercentChange As Long
    
    'Last Of Percent Change Column
    lastRowOfPercentChange = LastRowInColumn(sheet, 11)
    
    'Percent Change Range For Building Formula2
    percentRange = "K2:K" & lastRowOfPercentChange
    
    'Total Stock Volume Range For Building Formula2
    totalStkRange = "L2:L" & lastRowOfPercentChange
    
    'Ticker Summary Range For Building Formula2
    tickerSummaryRange = "i2:i" & lastRowOfPercentChange


    'Add Column Headers
    sheet.Range("P1").Value = "Ticker"
    sheet.Range("Q1").Value = "Value"
    
    'Add Greatest Row Headers
    sheet.Range("O2").Value = "Greatest % Increase"
    sheet.Range("O3").Value = "Greatest % Decrease"
    sheet.Range("O4").Value = "Greatest Total Volume"
    
    'Set Percent Change Cells Q2 and Q3 to percent
    sheet.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Set Total Stock Volume Cell Q4 to scientific
    sheet.Range("Q4").NumberFormat = "0.00E+00"
    
    
    '------------------------
    'Find And Set Greatest Increase Value
    sheet.Range("Q2").Formula2 = "=MAX(" & percentRange & ")"
    
    'Get Greatest Increase Ticker Name
    targetForXLookUp = "Q2"
    sheet.Range("P2").Formula2 = "=XLOOKUP(" & targetForXLookUp & ", " & percentRange & ", " & tickerSummaryRange & ")"
    
    
    '------------------------
    'Find And Set Greatest Decrease Value
    sheet.Range("Q3").Formula2 = "=MIN(" & percentRange & ")"
    
    'Set Greatest Decrease Ticker Name
    targetForXLookUp = "Q3"
    sheet.Range("P3").Formula2 = "=XLOOKUP(" & targetForXLookUp & ", " & percentRange & ", " & tickerSummaryRange & ")"
    
    
    '------------------------
    'Find And Set Greatest Total Stock Volume
    sheet.Range("Q4").Formula2 = "=MAX(" & totalStkRange & ")"
    
    'Set Greatest Decrease Ticker Name
    targetForXLookUp = "Q4"
    sheet.Range("P4").Formula2 = "=XLOOKUP(" & targetForXLookUp & ", " & totalStkRange & ", " & tickerSummaryRange & ")"
    
    
    
    
    
    'Autofit Columns Of Greatest Table
    sheet.Columns("O:Q").AutoFit
    
    
End Sub


