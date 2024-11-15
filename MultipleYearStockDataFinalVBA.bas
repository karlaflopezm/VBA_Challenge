Attribute VB_Name = "Module1"
Option Explicit

Sub AnalyzeStockData()

'Define variables

Dim i As Long
Dim j As Long
Dim ws As Worksheet

Dim LastRow As Long
Dim currentTicker As String
Dim previousTicker As String
Dim openingPrice As Double
Dim closingPrice As Double
Dim quarterlyChange As Double
Dim percentageChange As Double
Dim totalVolume As Double
Dim outputRow As Long
Dim DateCell As Date
 Dim currentQuarter As String

Dim start As Long
Dim find_value As Long

Dim quaterly_change_last_row As Long


Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Double
Dim tickerGreatestIncrease As String
Dim tickerGreatestDecrease As String
Dim tickerGreatestVolume As String

Dim rng As Range
 
'Worksheet Loop

    For Each ws In Worksheets
    

'Name header for columns that will show output information'

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"

    MsgBox "Worksheet: " & ws.Name
    
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
'Initiate variables

        outputRow = 2
        start = 2
        totalVolume = 0
          
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
    
' Loop through all tickers in worksheet

                 
                For i = 2 To LastRow
            
                                 
                    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
                        currentTicker = ws.Cells(i, 1).Value
                        ws.Range("I" & outputRow).Value = currentTicker
                                                
                        totalVolume = totalVolume + ws.Cells(i, 7).Value
                        
                                                
                         If totalVolume = 0 Then
                            Range("I" & outputRow).Value = Cells(i, 1).Value
                            Range("J" & outputRow).Value = 0
                            Range("K" & outputRow).Value = "%" & 0
                            Range("L" & outputRow).Value = 0

                        Else
                        
                        If Cells(start, 3) = 0 Then
                            For find_value = start To i
                                If Cells(find_value, 3).Value <> 0 Then
                                    start = find_value
                            Exit For
                        End If
                     Next find_value
                End If
                                                
                                                
                        openingPrice = ws.Cells(start, 3).Value
                        closingPrice = ws.Cells(i, 6).Value
                        quarterlyChange = closingPrice - openingPrice
                        ws.Range("J" & outputRow).Value = quarterlyChange
                        
                        
                        
                        percentageChange = (quarterlyChange / openingPrice)
                        ws.Range("L" & outputRow).Value = totalVolume
                        ws.Range("K" & outputRow).Value = percentageChange
                        ws.Range("K" & outputRow).NumberFormat = "0.00%"
                        
                        
                        
                        start = i + 1
                        
                    ' Coloring Code
                                     
                    
                    Set rng = ws.Range("J2  : J" & ws.Cells(ws.Rows.Count, "J").End(xlUp).Row)
                    
                    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
                    rng.FormatConditions(rng.FormatConditions.Count).Interior.Color = RGB(0, 255, 0)
                    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
                    rng.FormatConditions(rng.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)
 

                    End If
                        
                        
                        outputRow = outputRow + 1
                        totalVolume = 0
                        quarterlyChange = 0
                        
                              Else
                              
                        totalVolume = totalVolume + Cells(i, 7).Value
                        
                        
                End If
                         
                        ' Set Ticker, Value, Greatest %, Increase, % Decrease, and Total volume headers
                   
                        ws.Cells(1, 16).Value = "Ticker"
                        ws.Cells(1, 17).Value = "Value"
                        ws.Cells(2, 15).Value = "Greatest % Increase"
                        ws.Cells(3, 15).Value = "Greatest % Decrease"
                        ws.Cells(4, 15).Value = "Greatest Total Volume"
                        
                          If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) Then
                                
                                ws.Cells(2, 16).Value = Cells(i, 9).Value
                                ws.Cells(2, 17).Value = Cells(i, 11).Value
                                ws.Cells(2, 17).NumberFormat = "0.00%"
                        
                        ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) Then
                                
                                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                                ws.Cells(3, 17).NumberFormat = "0.00%"
                                
                    ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow)) Then
                    
                                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                                
                                
                                End If
    
                
              Next i
       

    Next ws
    
                        
    
    
    End Sub


