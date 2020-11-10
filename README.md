# VBA-Challenge

## Summary of Project
This repository contains a vba script and screenshots of a completed bootcamp assignment. 

The assignment required me to loop through a year of stock ticker data the included opening price, high/low price, closing price and volume for each day throughout the course of 1 year. The workbook contained 3 sheets with one year of data on each sheet. 

We were tasked with summarizing each unique stock ticker with the **yearly $ change, yearly % change and total annual volume.** 

In addition, a bonus assignment was offered to identify the stock with the **largest % increase, largest % decrease and the highest total annual volume.** 

## Pictures of Results

### 2016 Screenshot of summary table and bonus table: 

![2016 Screenshot](/2016-pt-1.png)

### 2015 Screenshot of summary table and bonus table: 

![2015 Screenshot](/2015-pt-1.png)

### 2014 Screenshot of summary table and bonus table: 

![2014 Screenshot](/2014-pt-1.png)

### VBA Code Syntax Below:
```vba
Sub stocktickers():
'Setup to loop through worksheets in workbook

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

ws.Activate

'Set variable for annual open price

Dim op As Double

'Set variable for year-end close price

Dim cl As Double

'Set variable to track annual volume

Dim vol As LongLong

'Set variable to hold Stock Ticker Symbol

Dim stock As String

'Set Variable for % change of yearly price

Dim percentage_change As Double

'Set variable to track row in summary table

Dim summary_row As Integer

'Set initial summary row and year-end close price

summary_row = 2

'Determine last row of stock ticker worksheet

lr = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set annaul open price for first stock ticker

op = ws.Cells(2, 3).Value

'Loop through stock tickers looking for change in symbol

For i = 2 To lr
  
'If values are different then

    If ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Set stock ticker
        stock = ws.Cells(i, 1).Value
         
        'Add final Volume to volume tracker
        vol = vol + ws.Cells(i, 7).Value
            
        'Set closing stock price
        cl = ws.Cells(i, 6).Value
             
        'Enter stock ticker in summary table
        ws.Range("I" & summary_row).Value = stock
            
        'Enter yearly change for stock in summary table
        ws.Range("J" & summary_row).Value = cl - op
          
        'Enter total annual volume in summary table
        ws.Range("L" & summary_row).Value = vol
             
'Calculate percentage change, accounting for DIV/0 potential
            
        If (op = 0 And cl = 0) Then
            
            percentage_change = 0
        
        ElseIf (op = 0 And cl > 0) Then
            
            percentage_change = 1
        
        ElseIf (op = 0 And cl < 0) Then
            
            percentage_change = -1
        
        Else
            
            percentage_change = (cl - op) / op
        
        End If
            
        ' Enter % change on summary table

        ws.Range("K" & summary_row).Value = percentage_change
            
        'Move to next row in Summary Table
             
        summary_row = summary_row + 1
             
        'Reset Volume
             
        vol = 0
             
        'Set Open price of next stock
             
        op = ws.Cells(i + 1, 3).Value
    
    Else
    
        'If stock tickers are the same, add volume to volume tracker
        
        vol = vol + ws.Cells(i, 7).Value
        
    End If
    
Next i

'Set Conditional Color Formatting for Positive or Negative Change

'Define last row for summary table

lrs = ws.Cells(Rows.Count, 9).End(xlUp).Row

For r = 2 To lrs

'If cells are greater than 0, color green

If ws.Cells(r, 10) > 0 Then

    ws.Cells(r, 10).Interior.ColorIndex = 4

'If cells are less than/equal to 0, color red

    Else:
        
        ws.Cells(r, 10).Interior.ColorIndex = 3
    
    End If
    
    Next r

'Define variables for bonus largest increase, largest decrease, highest volume

Dim percent_max As Double
Dim percent_min As Double
Dim vol_max As LongLong

'Find min/max values for percentage change and max value for Volume

percent_max = Application.WorksheetFunction.Max(ws.Range("K2:K" & lrs))
percent_min = Application.WorksheetFunction.Min(ws.Range("K2:K" & lrs))
vol_max = Application.WorksheetFunction.Max(ws.Range("L2:L" & lrs))

'Enter values in summary table

ws.Range("P2").Value = percent_max
ws.Range("P3").Value = percent_min
ws.Range("P4").Value = vol_max

'Find associated stock tickers for greatest increase, greatest decrease, highest volume

For r = 2 To lrs

If ws.Cells(r, 11) = percent_max Then
    
    ws.Cells(2, 15) = ws.Cells(r, 9)

ElseIf ws.Cells(r, 11) = percent_min Then
    
    ws.Cells(3, 15) = ws.Cells(r, 9)
   
ElseIf ws.Cells(r, 12) = vol_max Then
    
    ws.Cells(4, 15) = ws.Cells(r, 9)
    
    End If
  
  Next r

'Set headings for worksheet summary tables

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'Set formatting for numbers & column widths

ws.Range("K2:K" & lrs).NumberFormat = "0.00%"
ws.Range("P2:P3").NumberFormat = "0.00%"
ws.Range("P4").NumberFormat = "#,###0"
ws.Range("L2:L" & lrs).NumberFormat = "#,###0"
ws.Range("G2:G" & lr).NumberFormat = "#,###0"
ws.Range("B2:B" & lr).NumberFormat = "####-##-##"
Columns("A:P").EntireColumn.AutoFit

Next


End Sub

```
