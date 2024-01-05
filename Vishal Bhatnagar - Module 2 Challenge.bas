Attribute VB_Name = "Module1"
Sub Stocks()

Dim ws As Worksheet

' Loop through all sheets in the workbook
For Each ws In ThisWorkbook.Sheets

' Define headings for columns

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

' Set an initial variable for holding the Ticker Symbol
Dim LastRow As Double
Dim i As Double
Dim Ticker As String

' Set an initial variable for holding the Yearly Change
Dim YearlyChange As Double

' Set an initial variable for holding the Percent Change and Format output
Dim PercentChange As Double

' Set an initial variable for holding the Total Stock Volumn
Dim TotalStockVolume As Double

' Tracking location of each Ticker Symbol in the Summary Table
Dim SummaryTable As Integer
SummaryTable = 2

Dim ConditionalFormating As Range
Dim Color As FormatCondition

Dim NextTicker As String
Dim PreviousTicker As String
Dim OpeningPrice As Double
Dim ClosingPrice As Double

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double
Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestVolumeTicker As String

GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all Ticker Symbols
For i = 2 To LastRow

    'Set Ticker Symbol Value
     Ticker = ws.Cells(i, 1).Value
     NextTicker = ws.Cells(i + 1, 1).Value
     PreviousTicker = ws.Cells(i - 1, 1).Value
     
     ' Calculate Total Stock Volume
     TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
     
     ' Finds first row of data for each unique Ticker symbol
    If PreviousTicker <> Ticker Then
       OpeningPrice = ws.Cells(i, 3).Value
    
    ' Check for Ticker Symbol in the total list of symbols
    ElseIf NextTicker <> Ticker Then
                
        ClosingPrice = ws.Cells(i, 6).Value
        
        ' Print Ticker Symbol in Summary table
        ws.Range("I" & SummaryTable).Value = Ticker
        
        ' Calculate Yearly Change
        YearlyChange = ClosingPrice - OpeningPrice
       
        ' Print Yearly Change in Summary Table
        ws.Range("J" & SummaryTable).Value = YearlyChange
    
    ' Calculate Percent Change
        PercentChange = YearlyChange / OpeningPrice
        
        ' Finding Greatest % Increase
        If PercentChange > GreatestIncrease Then
            GreatestIncrease = PercentChange
            
            GreatestIncreaseTicker = Ticker
            
        End If
                         
         ' Finding Greatest % Decrease
         If PercentChange < GreatestDecrease Then
            GreatestDecrease = PercentChange
            
            GreatestDecreaseTicker = Ticker
            
         End If
         
         ' Finding Greatest Total Volume
         If TotalStockVolume > GreatestVolume Then
            GreatestVolume = TotalStockVolume
            
            GreatestVolumeTicker = Ticker
            
          End If
          
        ' Print Percent Change in Summary Table
        ws.Range("K" & SummaryTable).Value = PercentChange
        
        ' Calculate Formatted Percent Change
        ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
        
        ' Define Conditional Formating Range
        Set ConditionalFormating = ws.Range("J2:K" & SummaryTable)
        
        ' Conditional formating for Positive Change
        Set Color = ConditionalFormating.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        Color.Interior.Color = RGB(0, 255, 0)
        
        ' Conditional formating for Negative Change
        Set Color = ConditionalFormating.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        Color.Interior.Color = RGB(255, 0, 0)
                      
        ' Print Total Stock Volume in Summary Table
        ws.Range("L" & SummaryTable).Value = TotalStockVolume
        
        ' Populate next cell in Summary Table
        SummaryTable = SummaryTable + 1
        
        ' Reset Values
        PercentChange = 0
        TotalStockVolume = 0
        YearlyChange = 0
   
    End If
   
Next i

ws.Cells(2, 16).Value = GreatestIncreaseTicker
ws.Cells(2, 17).Value = GreatestIncrease
ws.Cells(2, 17).NumberFormat = "0.00%"

ws.Cells(3, 16).Value = GreatestDecreaseTicker
ws.Cells(3, 17).Value = GreatestDecrease
ws.Cells(3, 17).NumberFormat = "0.00%"

ws.Cells(4, 16).Value = GreatestVolumeTicker
ws.Cells(4, 17).Value = GreatestVolume

Next ws

End Sub





