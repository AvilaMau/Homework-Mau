'Locate and read the data on the ticker column, identify where the ticker changes names and compare the price between the opening price and closing price to identify the difference and wirte it in a new column at the end of the current data set'
'Locate and read the data'
'identify where the ticket number changes'
'substract the closing price on the last day from the opening price on the first day'
'show it on a new cell'
'calculate the percentage difference on a new column'
'add total stock volume from equal ticker name'
'Write a for loop to lock over every row'

Public Const tickerIndex As Integer = 1
Public Const dateIndex As Integer = 2
Public Const openIndex As Integer = 3
Public Const highIndex As Integer = 4
Public Const lowIndex As Integer = 5
Public Const closeIndex As Integer = 6
Public Const volumeIndex As Integer = 7

Public Const summaryTickerIndex As Integer = 9
Public Const summaryYearlyChangeIndex As Integer = 10
Public Const summaryPercentChangeIndex As Integer = 11
Public Const summaryTotalVolumeIndex As Integer = 12

Public Const greatestSummaryLabel As Integer = 15
Public Const greatestSummaryTickerIndex As Integer = 16
Public Const greatestSummaryValueIndex As Integer = 17

Type Ticker
 name As String
 date As String
 open As Double
 high As Double
 low As Double
 close As Double
 volume As Double
End Type

Type TickerSummary
    name As String
    yearlyChange As Double
    percentChange As Double
    totalStockVolume As Double
End Type

Function PopulateTicker(row As Long) As Ticker
    Dim Ticker As Ticker
    Ticker.name = Cells(row, tickerIndex)
    Ticker.date = Cells(row, dateIndex)
    Ticker.open = Cells(row, openIndex)
    Ticker.high = Cells(row, highIndex)
    Ticker.low = Cells(row, lowIndex)
    Ticker.close = Cells(row, closeIndex)
    Ticker.volume = Cells(row, volumeIndex)
    PopulateTicker = Ticker
End Function

Sub PrintSummary(Summary As TickerSummary, row As Integer)
    Cells(row, summaryTickerIndex) = Summary.name
    Cells(row, summaryYearlyChangeIndex) = Summary.yearlyChange
    Cells(row, summaryPercentChangeIndex) = Summary.percentChange
    Cells(row, summaryTotalVolumeIndex) = Summary.totalStockVolume
    
    Cells(row, summaryYearlyChangeIndex).Interior.ColorIndex = 4
    
        If Cells(row, summaryYearlyChangeIndex) < 0 Then
            
            Cells(row, summaryYearlyChangeIndex).Interior.ColorIndex = 3
                
        End If
    
    Cells(row, summaryPercentChangeIndex).NumberFormat = "0.00%"
    
End Sub

Sub PrintSummaryHeader(row As Integer)
    Cells(1, summaryTickerIndex) = "Ticker"
    Cells(1, summaryYearlyChangeIndex) = "Yearly Change"
    Cells(1, summaryPercentChangeIndex) = "Percent Change"
    Cells(1, summaryTotalVolumeIndex) = "Total Stock Volume"
End Sub

Sub PrintGreatestTickerSummaries(greatestIncrease As TickerSummary, greatestDecrease As TickerSummary, greatestTotalVolume As TickerSummary)
    Cells(1, greatestSummaryTickerIndex) = "Ticker"
    Cells(1, greatestSummaryValueIndex) = "Value"
    
    Cells(2, greatestSummaryLabel) = "Greatest % Increase"
    Cells(2, greatestSummaryTickerIndex) = greatestIncrease.name
    Cells(2, greatestSummaryValueIndex) = greatestIncrease.percentChange
    Cells(2, greatestSummaryValueIndex).NumberFormat = "0.00%"
    
    Cells(3, greatestSummaryLabel) = "Greatest % Decrease"
    Cells(3, greatestSummaryTickerIndex) = greatestDecrease.name
    Cells(3, greatestSummaryValueIndex) = greatestDecrease.percentChange
    Cells(3, greatestSummaryValueIndex).NumberFormat = "0.00%"
    
    Cells(4, greatestSummaryLabel) = "Greatest Total Volume"
    Cells(4, greatestSummaryTickerIndex) = greatestTotalVolume.name
    Cells(4, greatestSummaryValueIndex) = greatestTotalVolume.totalStockVolume
End Sub

    
Sub Test()
    Dim tickerRow As Long
    Dim summaryRow As Integer
    tickerRow = 2
    summaryRow = 2
    
    Call PrintSummaryHeader(summaryRow - 1)
    
    Dim FirstTicker As Ticker
    Dim TickerSummary As TickerSummary
    Dim GreatestIncreaseSummary As TickerSummary
    Dim GreatestDecreaseSummary As TickerSummary
    Dim GreatestTotalVolumeSummary As TickerSummary
    Dim CurrentTicker As Ticker
        
    While Cells(tickerRow, tickerIndex) <> ""
    
        CurrentTicker = PopulateTicker(tickerRow)
    
        If CurrentTicker.name <> FirstTicker.name Then
            FirstTicker = PopulateTicker(tickerRow)
            TickerSummary.name = FirstTicker.name
            TickerSummary.totalStockVolume = 0
        End If
        TickerSummary.totalStockVolume = TickerSummary.totalStockVolume + CurrentTicker.volume
        
        tickerRow = tickerRow + 1
    
        If Cells(tickerRow, tickerIndex) <> FirstTicker.name Then
            TickerSummary.yearlyChange = CurrentTicker.close - FirstTicker.open
            TickerSummary.percentChange = TickerSummary.yearlyChange / FirstTicker.open
            
            If TickerSummary.yearlyChange > GreatestIncreaseSummary.yearlyChange Then
                GreatestIncreaseSummary = TickerSummary
            End If
            
            If TickerSummary.yearlyChange < GreatestDecreaseSummary.yearlyChange Then
                GreatestDecreaseSummary = TickerSummary
            End If
            
            If TickerSummary.totalStockVolume > GreatestTotalVolumeSummary.totalStockVolume Then
                GreatestTotalVolumeSummary = TickerSummary
            End If
            
            Call PrintSummary(TickerSummary, summaryRow)
            
            summaryRow = summaryRow + 1
        End If
    
    Wend
    
    Call PrintGreatestTickerSummaries(GreatestIncreaseSummary, GreatestDecreaseSummary, GreatestTotalVolumeSummary)

End Sub

