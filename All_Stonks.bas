Attribute VB_Name = "Module1"
Sub All_Stonks()

'Code credit for running code on multiple sheets: https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Stonks
    Next
    Application.ScreenUpdating = True
End Sub
Sub Stonks():

'Insert headers and format cell width for headers in summary table

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Columns("J").ColumnWidth = 13
Cells(1, 11).Value = "Percent Change"
Columns("K").ColumnWidth = 14
Cells(1, 12).Value = "Total Stock Volume"
Columns("L").ColumnWidth = 17

'Format cell width for later columns
Columns("O").ColumnWidth = 20
Columns("Q").ColumnWidth = 17

'Set variables for ticker and the total stock value

Dim Ticker As String
Dim Stock_Total As Double
Stock_Total = 0

'Set var for starting and ending row numbers
Dim r_start As Double
Dim r_end As Double
r_start = 2
r_end = 2

'Set var for calculating the yearly and percent change

Dim change_y As Double
Dim change_p As Double
change_y = 0
change_p = 0

'Set variable so each ticker, change value and vol total goes to a unique row in the summary table
Dim Sum_row As Double
Sum_row = 2

'Using a countif function on the first column to determine how many entries are in the table.
'This way the loop will run for as many entries as needed, even if the total number of entries change on a different sheet.
Dim entries As Double
entries = Application.WorksheetFunction.CountIf(Range("A:A"), "<>")


' Loop through all stocks to collect the ticker, change values and total volume value for each ticker
  For i = 2 To entries

'If the ticker value is different from the cell above it, this establishes the first row of data for a ticker.
'We set the ticker value, the row value for our opening price, and add the first amount to the stock volume.

    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
      Ticker = Cells(i, 1).Value
      r_start = i
      Stock_Total = Stock_Total + Cells(i, 7).Value
      
'If the ticker value is different from the cell below it, we've found the last row of that stock.
'we add to the total volume and set the row number for our closing value
'We then perform all the math needed to calculate the change from open to close and the % change
'We then print our ticker, the change values and total stock volume to our summary table
'We set the new row number for our next stock we will record in the summary table and reset the volume total to zero
      
      ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Stock_Total = Stock_Total + Cells(i, 7).Value
      r_end = i
      
      change_y = Cells(r_end, 6) - Cells(r_start, 3)
      change_p = change_y / Cells(r_start, 3)
      Range("I" & Sum_row).Value = Ticker
      Range("J" & Sum_row).Value = change_y
      Range("K" & Sum_row).Value = change_p
      Range("L" & Sum_row).Value = Stock_Total
      Sum_row = Sum_row + 1
      Stock_Total = 0

    ' In all other instances, we add to the total which will only happen inbetween the first and last instance of the stock.
    Else
      Stock_Total = Stock_Total + Cells(i, 7).Value

    End If

  Next i

'Formatting for columns with yearly and % changes

'Assignment instructions did not give a color for "no change" values so I am categorizing them as green since they are not negative

Dim y_entries As Double
y_entries = Application.WorksheetFunction.CountIf(Range("J:J"), "<>")

For y = 2 To y_entries

    If Cells(y, 10) >= 0 Then
        Cells(y, 10).Interior.ColorIndex = 4
    Else
    Cells(y, 10).Interior.ColorIndex = 3
    
    End If
    
Next y
 
'Formating for the columns
Range("J:J").NumberFormat = "0.00"
Range("K:K").NumberFormat = "0.00%"



'****Part 2, added summary table, headers and axis labeling

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'setting values to count the number of instances in column K and L, tehcnically just one of these could be used for both but in case there is something wrong with one column of data it counts values for the data it later needs to summarize.

Dim p_entries As Double
p_entries = Application.WorksheetFunction.CountIf(Range("K:K"), "<>")

Dim v_entries As Double
v_entries = Application.WorksheetFunction.CountIf(Range("L:L"), "<>")

'Setting variables to find the min and max in the list and store the row number to later pull the Ticker

Dim max_p As Double
Dim min_p As Double
Dim max_v As Double
Dim r1 As Integer
Dim r2 As Integer
Dim r3 As Integer

max_p = 0
min_p = 0
max_v = 0
r1 = 0
r2 = 0
r3 = 0

'Two seperate loops to find the min and max percent changes and the largest stock volume
'The first loop is for the two percent change values, the second loop is for the max volume.

For a = 2 To p_entries

    If Cells(a, 11) >= max_p Then
        max_p = Cells(a, 11)
        r1 = a
    
    ElseIf Cells(a, 11) < min_p Then
        min_p = Cells(a, 11)
        r2 = a
    
    End If
Next a

For b = 2 To v_entries

    If Cells(b, 12) >= max_v Then
        max_v = Cells(b, 12)
        r3 = b

    End If
Next b

'Print and format the values

Cells(2, 16).Value = Cells(r1, 9)
Cells(2, 17).Value = max_p
Cells(3, 16).Value = Cells(r2, 9)
Cells(3, 17).Value = min_p
Cells(4, 16).Value = Cells(r3, 9)
Cells(4, 17).Value = max_v

'Format the final values

Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"
Cells(4, 17).NumberFormat = "0.00E+0"


End Sub
