Attribute VB_Name = "Módulo1"
Option Explicit


Sub stockMarket_Summary()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        ' set the position in the sheet
        
         ws.Activate
         ws.Select
        
         Call stockMarket_Summary_perSheet
         Call grandSummary_perSheet
                 
         Columns("L:L").EntireColumn.AutoFit
         Columns("P:P").EntireColumn.AutoFit
         
    Next ws

End Sub

Private Sub writeStockSummary(pRow, pTicket, pOpeningValue, pClosingValue, pTotalVolume)
    
    ' start writing summary in this column
    Dim column As Integer
    column = 9
    
    If pRow = 1 Then
        
        ' write header
        
        Cells(pRow, column).Value = "Ticker"
        Cells(pRow, column + 1).Value = "Yearly Change"
        Cells(pRow, column + 2).Value = "Percent Change"
        Cells(pRow, column + 3).Value = "Total Stock Volume"
    
    Else
        
        Cells(pRow, column).Value = pTicket
        Cells(pRow, column + 1).Value = (pClosingValue - pOpeningValue)
        Cells(pRow, column + 2).Value = (pClosingValue - pOpeningValue) / pOpeningValue
        Cells(pRow, column + 3).Value = pTotalVolume
        
        ' format and color cells
        
        If (pClosingValue - pOpeningValue) < 0 Then
        
            Cells(pRow, column + 1).Interior.ColorIndex = 3
        
        Else
            
            Cells(pRow, column + 1).Interior.ColorIndex = 4
            
        End If
                    
        Cells(pRow, column + 1).NumberFormat = "00.00"
        Cells(pRow, column + 2).NumberFormat = "00.00%"
        
    End If
    
    
End Sub

Private Sub grandSummary_perSheet()
    
    Dim bestTicker As String
    Dim greatestIncrease As Double
    Dim worstTicker As String
    Dim greatestDecrease As Double
    Dim bestVolTicker As String
    Dim greatestVolume As Double
    Dim i As Integer
        
    bestTicker = ""
    worstTicker = ""
    bestVolTicker = ""
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
        
    ' print header
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Range("P2:P3").NumberFormat = "00.00%"
    
    ' calculate and print Grand Summary
    
    For i = 2 To Cells(Rows.Count, 9).End(xlUp).Row
        
        If Cells(i, 11).Value > greatestIncrease Then
        
            greatestIncrease = Cells(i, 11).Value
            bestTicker = Cells(i, 9).Value
            
        End If
        
        If Cells(i, 11).Value < greatestDecrease Then
        
            greatestDecrease = Cells(i, 11).Value
            worstTicker = Cells(i, 9).Value
        
        End If
        
        If Cells(i, 12).Value > greatestVolume Then
        
            greatestVolume = Cells(i, 12).Value
            bestVolTicker = Cells(i, 9).Value
            
        End If
        
    Next i
        
        Cells(2, 15).Value = bestTicker
        Cells(2, 16).Value = greatestIncrease
        Cells(3, 15).Value = worstTicker
        Cells(3, 16).Value = greatestDecrease
        Cells(4, 15).Value = bestVolTicker
        Cells(4, 16).Value = greatestVolume
     
    
End Sub

Private Sub stockMarket_Summary_perSheet()

    Dim ticket As String
    Dim openingValue As Double
    Dim closingValue As Double
    Dim totalVolume As Double
    Dim initialRow As Integer
    Dim summaryRow As Integer
    Dim i As Double
    
    initialRow = 2
    summaryRow = 1
    
    closingValue = 0
    totalVolume = 0
    openingValue = 0
        
    ticket = Cells(initialRow, 1).Value
    openingValue = Cells(initialRow, 3).Value
    totalVolume = Cells(initialRow, 7).Value
    
    Call writeStockSummary(pRow:=summaryRow, pTicket:=ticket, pOpeningValue:=openingValue, pClosingValue:=closingValue, pTotalVolume:=totalVolume)
    summaryRow = summaryRow + 1
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        ' Change stock
        
            closingValue = Cells(i, 6).Value
            totalVolume = totalVolume + Cells(i, 7).Value
            
            Call writeStockSummary(pRow:=summaryRow, pTicket:=ticket, pOpeningValue:=openingValue, pClosingValue:=closingValue, pTotalVolume:=totalVolume)
                        
            summaryRow = summaryRow + 1
            ticket = Cells(i + 1, 1).Value
            openingValue = Cells(i + 1, 3).Value
            totalVolume = Cells(i + 1, 7).Value
            
        Else
        
        ' Same stock
        
            totalVolume = totalVolume + Cells(i, 7).Value
            
        End If
    
    Next i


End Sub
    

