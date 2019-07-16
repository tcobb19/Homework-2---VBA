Sub StockSum()

'define variables'

Dim lRow As Long
Dim ticker As String
Dim totalvolume As Variant
Dim rowvolume As Long
Dim itemcount As Long
Dim firstrow As Long
Dim lastrow As Long
Dim firstprice As Double
Dim lastprice As Double
Dim ratio As Variant

ticker = " "
itemcount = 0
totalvolume = 0
rowvolume = 0
firstrow = 2
lastrow = 2

'determine # of rows in worksheet'

lRow = Cells(Rows.Count, 1).End(xlUp).Row

'identify each unique item in ticket column'

For i = 2 To lRow + 1

    'check if the stock in the row is a different from the last'
    If Not Cells(i, 1).Value = ticker Then
    
    'Return the total volume, yearly price change, and percent yearly price change of the previous stock in the results table'
    firstprice = Cells(firstrow, 3).Value
    lastprice = Cells(lastrow, 6).Value
    Cells(itemcount + 1, 9).Value = ticker
    Cells(itemcount + 1, 10).Value = lastprice - firstprice
    Cells(itemcount + 1, 10).NumberFormat = "0.00"
    'Conditional color cells in the yearly price change, green if positive and red if negative'
    If Cells(itemcount + 1, 10).Value > 0 Then
        Cells(itemcount + 1, 10).Interior.ColorIndex = 4
    Else: Cells(itemcount + 1, 10).Interior.ColorIndex = 3
    End If
    
    'calculate and return percentage change'
    ratio = Round(lastprice / firstprice, 3)
    Cells(itemcount + 1, 11).Value = ratio
    Cells(itemcount + 1, 11).NumberFormat = "0.00%"
    'return total volume'
    Cells(itemcount + 1, 12).Value = totalvolume
    
    'Proceed to next row'
    itemcount = itemcount + 1
    
    'record the first row in which the new stock appears'
    firstrow = i
    lastrow = i
    'change the value to the new stock'
    ticker = Cells(i, 1).Value
    totalvolume = Cells(i, 7).Value
    
    'if the same stock as previous row, add row volume to total volume'
    'add the row volume to the total'
    Else: rowvolume = Cells(i, 7).Value
    totalvolume = totalvolume + rowvolume
    
    'increment the last row counter'
    lastrow = lastrow + 1
    End If

Next i

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 10).Interior.ColorIndex = 2
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

End Sub

