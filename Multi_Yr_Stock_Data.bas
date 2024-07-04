Attribute VB_Name = "Module1"
Sub VBA_Challenge()

'_______________________________________________________________________________________________

'---Defining Variables---
'_______________________________________________________________________________________________

'---Define ws a Worksheet and executing a loop through worksheets A to F---
    Dim ws As Worksheet
    For Each ws In Worksheets

'---Defining variables for the Ticker, Quarterly Change, Percent Change and Total Stock Volume---
    Dim Ticker_Name As String
    Dim Quart_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim i As Long
    Dim Great_Inc As Double
    Dim Great_Dec As Double
    Dim Greal_Vol As Double
    Dim Inc_match As String
    Dim Dec_match As String
    
'---Working Summary table to collect values for each quarter and storing from row 2---
    Dim Summary_Table As Integer
    Summary_Table = 2
    
'---Defining variables to populate the Summary table---
    Dim Open_Price As Double
    Dim Close_Price As Double
'_______________________________________________________________________________________________

'---Columns Headers and Initial Values---
'_______________________________________________________________________________________________

'---Column headers for Summary_Tables in each worksheet---
    ws.Range("J1").Value = "Ticker"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("K1").Value = "Quarterly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("R1").Value = "Value"
    
'---Starting Values for each Variable---
    Open_Price = ws.Cells(2, 3).Value
    Close_Price = 0
    Quart_Change = 0
    Percent_Change = 0
    Total_Volume = 0

'_______________________________________________________________________________________________

'---Summary table for the Ticker, Quarterly Change, Percent Change and Total Stock Volume---
'_______________________________________________________________________________________________
    
    '---Find the last row in each worksheet---
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    '---Loop through all the stocks from the row 2 to the last row for each sheet---
    For i = 2 To LastRow
    
        '---Add to the Total Stock Volume for each ticker---
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
        '---Identify all same ticker in this quarter and set it in the Ticker variable---
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker_Name = ws.Cells(i, 1).Value
                        
            '---Set a Closing price variable for each ticker---
            Close_Price = ws.Cells(i, 6).Value
        
            '---Calculate the Price difference of the Closing Price from the Open Price---
            Quart_Change = Close_Price - Open_Price
            
            '---Calculate the Percent Difference of the Closing Price---
                
            Percent_Change = (Close_Price - Open_Price) / Open_Price
                 
                    
            '---Print the Ticker name in the Summary Table under Ticker header---
            ws.Range("J" & Summary_Table).Value = Ticker_Name
            
            '---Print the Quarterly Change value for the given ticker under the Quarterly Change header---
            ws.Range("K" & Summary_Table).Value = Quart_Change
            
            '---Print the Percent Change value for the given ticker under the Percent Change header and change to % format---
            ws.Range("L" & Summary_Table).Value = Percent_Change
            ws.Range("L:L").NumberFormat = "0.00%"
            
            '---Print the Total Stock Volume value for the given ticker under the Total Volume column---
            ws.Range("M" & Summary_Table).Value = Total_Volume
            
            '---Set conditional format to Quarterly Change values; Red if the change is -ve and Green if the change is +ve---
                If ws.Range("K" & Summary_Table).Value >= 0 Then
                  
                  ws.Range("K" & Summary_Table).Interior.ColorIndex = 4
                  
                  ElseIf ws.Range("K" & Summary_Table).Value < 0 Then
                  
                  ws.Range("K" & Summary_Table).Interior.ColorIndex = 3
                
                
        End If
        
            '---Add 1 to the summary table count---
            Summary_Table = Summary_Table + 1
                             
            
            '---Reset the Closing Price, Price_Diff and Per_Diff to 0 to calculate the values for a new ticker---
            Close_Price = 0
            
            Quart_Change = 0
            
            Percent_Change = 0
            
            Total_Volume = 0
            
            Ticker_Name = ws.Cells(i + 1, 1).Value
            
            '---Open price for the next Ticker---
            Open_Price = ws.Cells(i + 1, 3).Value
          
        End If
        
    Next i
    
'_______________________________________________________________________________________________

'---Summary table for the Greatest % Increase, Decrease, and Total Volume---
'_______________________________________________________________________________________________

    '---Search for the maximum value under Percent Change---
    Great_Inc = ws.Application.WorksheetFunction.Max(ws.Range("L:L"))
    
    '---Change the formatting for the Greatest % increase to %---
    ws.Range("R2").NumberFormat = "0.00%"
                
   '---Print for the maximum value in the Greatest % increase---
    ws.Range("R2") = Great_Inc
    
    '---Find the corresponding Ticker to the Greatest % Increase---
    Inc_match = WorksheetFunction.Match(ws.Range("R2").Value, ws.Range("L:L"), 0)
    ws.Range("Q2").Value = ws.Range("J" & Inc_match)
    
'________________________________________________________________________________________________

    '---Search for the minimum value under Percent Change---
    Great_Dec = ws.Application.WorksheetFunction.Min(ws.Range("L:L"))
    
    '---Change the formatting for the Greatest % decrease to %---
    ws.Range("R3").NumberFormat = "0.00%"
                
   '---Print for the maximum value in the Greatest % decrease---
    ws.Range("R3") = Great_Dec
    
    '---Find the corresponding Ticker to the Greatest % decrease---
    Dec_match = WorksheetFunction.Match(ws.Range("R3").Value, ws.Range("L:L"), 0)
    ws.Range("Q3").Value = ws.Range("J" & Dec_match)
    
'________________________________________________________________________________________________

    '---Search for the Greatest Total Volume---
    Greal_Vol = ws.Application.WorksheetFunction.Max(ws.Range("M:M"))
           
   '---Print for the Greatest Total Volume---
    ws.Range("R4") = Greal_Vol
    
    '---Find the corresponding Ticker to Greatest Total Volume---
    Greal_Vol = WorksheetFunction.Match(ws.Range("R4").Value, ws.Range("M:M"), 0)
    ws.Range("Q4").Value = ws.Range("J" & Greal_Vol)
        
    Next ws
     
        
End Sub


