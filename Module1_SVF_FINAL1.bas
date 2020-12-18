Attribute VB_Name = "Module1"
Sub stock_ticker()

' Variables
Dim ticker_name As String 'Ticker List summary
Dim yearly_change As Double ' Yearly Change summary
Dim percent_change As Double 'Percent Change summary
Dim total_volume As LongLong 'Total Stock Volume, set to 0
        total_volume = 0
Dim ticker_list As Integer 'Container for Ticker List summary table, row 2
        ticker_list = 2
Dim last_row As Long '
Dim open_price As Double
Dim close_price As Double
Dim ws As Worksheet

'Label Cell Headers
Cells(1, 9).Value = "Ticker List"
Cells(1, 10).Value = "Yearly Change"

Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"


open_price = Range("C2").Value 'Container to hold open price value for FIRST stock only; loop below will catch all others EXCEPT for this first instance.


' Determine the Last Row
last_row = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To last_row
    
    
'--Loops all tickers------------------------------
'If next to prior tickers don't match
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Takes ticker name and adds it to the ticker total variable
    ticker_name = Cells(i, 1).Value
    total_volume = total_volume + Cells(i, 7).Value
    
'Yearly change calc
yearly_change = close_price - open_price

'Percent change calc
If open_price = 0 And close_price <> 0 Then
percent_change = -100
ElseIf close_price <> open_price Then
percent_change = ((close_price - open_price)) / open_price
Else: percent_change = 0
End If

'Adds ticker to ticker list summary - what to put where
    Range("I" & ticker_list).Value = ticker_name

'Adds volume to total stock volume summary
    Range("L" & ticker_list).Value = total_volume
    
'Adds yearly change to total stock volume summary
    Range("J" & ticker_list).Value = yearly_change
    
'Adds yearly change to total stock volume summary
    Range("K" & ticker_list).Value = percent_change

'Adds one to the summary table row - add to next row
      ticker_list = ticker_list + 1
      
'Resets ticker for next volume summary
    total_volume = 0

open_price = Cells(i + 1, 3).Value


Else

'Add ticker to total stock volume summary
total_volume = total_volume + Cells(i, 7).Value
close_price = Cells(i + 1, 6).Value


End If

  Next i

'-----Formatting percent change
Columns("K").NumberFormat = "0.00%"


'-------Yearly change red/green conditional shading
Dim Cell_Color As Range
Dim Rng As Range
Set Rng = Range("J2", Range("J2").End(xlDown))
For Each Cell_Color In Rng
If Cell_Color.Value < 0 Then
Cell_Color.Interior.Color = vbRed
Else
Cell_Color.Interior.Color = vbGreen
End If
Next Cell_Color



End Sub

