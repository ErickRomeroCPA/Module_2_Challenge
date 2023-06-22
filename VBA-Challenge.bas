VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks():

For Each ws In Worksheets

'create variables
Dim open_price_counter As Integer
Dim summary_table_row As Integer
Dim yearly_change As Double
Dim total_volume As Variant
Dim last_row As Integer
Dim max As Double
Dim min As Double
Dim greatest_total_volume As Variant
Dim lr As Integer

'store variables
open_price_counter = 2
summary_table_row = 2
total_volume = 0
last_row = Cells(Rows.Count, 1).End(xlUp).Row


'create column headers
Range("i1").Value = "Ticker"
Range("j1").Value = "Yearly Change"
Range("k1").Value = "Percent Change"
Range("l1").Value = "Total Stock Volume"

Range("n2").Value = "Greatest % Increase"
Range("n3").Value = "Greatest % Decrease"
Range("n4").Value = "Greatest Total Volume"
Range("n2").Value = "Greatest % Increase"
Range("o1").Value = "Ticker"
Range("p1").Value = "Value"

'create conditional loop to calculate yearly change, percent change &  total stock volume
For Z = 2 To last_row

    If Cells(Z + 1, 1).Value <> Cells(Z, 1).Value Then
        
        Range("i" & summary_table_row).Value = Cells(Z, 1).Value
        
        yearly_change = Range("f" & Z).Value - Range("c" & open_price_counter).Value
        
        total_volume = total_volume + Cells(Z, 7).Value
        
        Range("j" & summary_table_row).Value = yearly_change
        
        Range("k" & summary_table_row).Value = FormatPercent(yearly_change)
        
        Range("l" & summary_table_row).Value = total_volume
        
        summary_table_row = summary_table_row + 1
        
        open_price_counter = Z + 1
        
        total_volume = 0

'store total volume in counter when ticker symbols are the same
    Else
    
        total_volume = total_volume + Cells(Z, 7).Value

'end conditional loop
    End If
    
Next Z

'store variables
lr = Cells(Rows.Count, 10).End(xlUp).Row

max = Application.WorksheetFunction.max(Range("k2" & ":" & "k" & lr))

min = Application.WorksheetFunction.min(Range("k2" & ":" & "k" & lr))

greatest_total_volume = Application.WorksheetFunction.max(Range("l2" & ":" & "l" & lr))

'storing values in cells

Range("p2").Value = FormatPercent(max)

Range("p3").Value = FormatPercent(min)
 
Range("p4").Value = greatest_total_volume

'creating conditional loop to calculate greatest % increase, greatest % decrease, & greatest total volume & their associated ticker symbols

For Z = 2 To lr
    
    If Cells(Z, 11) = max Then
        
        Range("o2").Value = Cells(Z, 9).Value
    
    ElseIf Cells(Z, 11) = min Then
        
        Range("o3").Value = Cells(Z, 9).Value
    
    End If
    
    If Cells(Z, 12) = greatest_total_volume Then
        
        Range("o4").Value = Cells(Z, 9).Value

'end conditional loop
    End If
    
Next Z

Next ws

End Sub

