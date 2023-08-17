Attribute VB_Name = "Challenge_VBA"
Sub Challenge_2_vba():
Dim ws As Worksheet
' Loop through all sheets in the workbook
For Each ws In ThisWorkbook.Sheets
ws.Activate
        
Dim ticker As String
Dim open_col As Double
Dim close_col As Double
Dim yearly_change As Double
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim firstRow As Double
Dim precentage_change As Long
Dim total_stock As Double

'summary table values
Dim max_ticker As String
Dim min_ticker As String
Dim volume_ticker As String
'summary table values
Dim max_value As Double
Dim min_value As Double
Dim max_volume As Double




' Keep track of the location in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
firstRow = 2
total_stock = 0

'inital values for summary tables
max_value = 0
min_value = 0
max_volume = 0



'start loop
For i = 2 To lastRow

'cycle through
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the ticker
ticker = Cells(i, 1).Value


' Print the ticker in the Summary Table
Range("I" & Summary_Table_Row).Value = ticker

 
'close value
close_col = Cells(i, 6).Value
  
'open values
open_col = Cells(firstRow, 3).Value
      
'formula for yearly change
yearly_change = close_col - open_col
  
' Print the yearly change in the Summary Table
Range("J" & Summary_Table_Row).Value = yearly_change



    
'calculate percentage change
percentage_change = (yearly_change / open_col)


'find max value from percentage change
If percentage_change > max_value Then
max_value = percentage_change
max_ticker = ticker
End If



'find min value from percentage change
If percentage_change < min_value Then
min_value = percentage_change
min_ticker = ticker
End If


' Add to the total stock
total_stock = total_stock + Cells(i, 7).Value

' find the total stock max volume
If total_stock > max_volume Then
max_volume = total_stock
volume_ticker = ticker
End If


'imput the values into the cell
Cells(2, 16).Value = max_ticker
Cells(2, 17).Value = max_value
Cells(3, 16).Value = min_ticker
Cells(3, 17).Value = min_value
Cells(4, 16).Value = volume_ticker
Cells(4, 17).Value = max_volume
'format to percentage
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"
                      
' Print the percentage change in the Summary Table (Column K)
Range("K" & Summary_Table_Row).Value = percentage_change


'format to percentage
Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

' Print the Brand Amount to the Summary Table
Range("L" & Summary_Table_Row).Value = total_stock

'conditinal formating for yearly change
If yearly_change > 0 Then
Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
Else
Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
 End If

'conditinal formating for percentage change
If percentage_change > 0 Then
Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
Else
Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
 End If

' Reset the Total stock
total_stock = 0

'move onto next cell
firstRow = i + 1
      
' Add one to the ticker row
  Summary_Table_Row = Summary_Table_Row + 1
    
' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Total total stock
      total_stock = total_stock + Cells(i, 7).Value


  
  End If
  
 Next i
'name header ticker
 ws.Cells(1, 9).Value = "Ticker"
 'name header Yearly change
 ws.Cells(1, 10).Value = "Yearly Change"
  'name header Total stock
 ws.Cells(1, 12).Value = "Total stock values"
 'name the header for precentagechange
 ws.Cells(1, 11).Value = "Precentage Change"
 ' Populate the summary information on the same worksheet
Cells(1, 15).Value = "Metric"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
    
' Autofit to display data
ws.Columns("I:Q").AutoFit

    
    Next ws
MsgBox ("Fixes Complete")

End Sub



