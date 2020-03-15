Attribute VB_Name = "Module1"
Sub process_stock()
    
    Dim StockBook As Worksheet
    Dim TableData, UniqueTicker, SummaryData As Range
    Dim startprice, endprice As Double


'Turn off Screen Updating to speed up execution
    Application.ScreenUpdating = False


For Each StockBook In ThisWorkbook.Worksheets
'StockBook is a Range that contains all ticker values in the table
'TableData is a Range that contains the entire table
    Set TableData = StockBook.Range("A1", StockBook.Range("A1").End(xlToRight).End(xlDown))
    Set Performers = StockBook.Range("O1:Q3")

    
    
    
    
    
'Use Excel's Sort and Filter function to put the Unique ticker values in Column I _
 and assign the range of Unique Values to UniqueTicker
 
    TableData.Sort Header:=xlYes, Key1:=TableData.Columns(1), Order1:=xlAscending, Key2:=TableData.Columns(2), Order2:=xlAscending
    TableData.Columns(1).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=StockBook.Range("I1"), Unique:=True
    Set UniqueTicker = StockBook.Range("I2", StockBook.Range("I2").End(xlDown))
    
'Populate Headers and Summary Labels
    StockBook.Range("J1") = "Yearly Change"
    StockBook.Range("K1") = "Percent Change"
    StockBook.Range("L1") = "Total Volume"
    Performers.Cells(1, 1) = "Greatest % Increase"
    Performers.Cells(2, 1) = "Greatest % Decrease"
    Performers.Cells(3, 1) = "Greatest Total Volume"


    
'Enabling Filtering on the Full Table.  Filtered versions of the table will be used to perform operations on the Tickers
    TableData.AutoFilter
    
    
'For each Stock Ticker Apply a Filter.
'Since the list has been sorted by Date the Starting Price will be in the first row _
 and the Ending Price will be in the last row
 
    For Each stock In UniqueTicker
        'Apply the Filter
        TableData.AutoFilter Field:=1, Criteria1:=stock.Value
        
        'Get the Starting Price
        startprice = TableData.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 6)
        
        'Get the Ending Price
        endprice = TableData.Cells(1, 6).End(xlDown)
        
        'Compute the Total Difference and Populate and Format the Table
        stock.Offset(0, 1) = endprice - startprice
        If stock.Offset(0, 1) > 0 Then
            stock.Offset(0, 1).Interior.ColorIndex = 4
        Else
            stock.Offset(0, 1).Interior.ColorIndex = 3
        End If
        
        'Compute the Percentage Change and Populate the Table
        If startprice <> 0 Then
            stock.Offset(0, 2) = stock.Offset(0, 1) / startprice
        Else
            stock.Offset(0, 2) = 0
        End If
        
        
        'Compute the Total Volume and Populate the Table
        stock.Offset(0, 3) = WorksheetFunction.Sum(TableData.Offset(1).SpecialCells(xlCellTypeVisible).Columns(7))

        
    Next stock

'Disable the Filter

    StockBook.AutoFilterMode = False


    
'BONUS Performance Table

'Set a Range to the Newly Created Table
    Set SummaryData = StockBook.Range("I1").CurrentRegion
'Format Percentage Change Column as a Percent
    SummaryData.Columns(3).NumberFormat = "0.00%"
    
'Get the Maximum and Minimum % Change Stocks
'Similar to above I have sorted the data and retrieved the first and last rows to get Minimum and Maximum Respectively
'This only needs to be done once I am making the assumption that for any given sampling of stocks (i.e Alphabetically or By Year)
'The minimum will be the largest decrease and the maximum will be the largest increase
    SummaryData.Sort Key1:=SummaryData.Columns(3), Order1:=xlAscending, Header:=xlYes
    Performers.Cells(1, 2) = SummaryData(SummaryData.Rows.Count, 3)
    Performers.Cells(1, 2).NumberFormat = "0.00%"
    Performers.Cells(2, 2) = SummaryData(2, 3)
    Performers.Cells(2, 2).NumberFormat = "0.00%"
    
'Same for Max Volume
    SummaryData.Sort Key1:=SummaryData.Columns(4), Order1:=xlDescending, Header:=xlYes
    Performers.Cells(3, 2) = SummaryData(2, 4)
    
'Put the Summary Table back in Alphabetical Order
    SummaryData.Sort Key1:=SummaryData.Columns(1), Order1:=xlAscending, Header:=xlYes
Next

    Application.ScreenUpdating = True

End Sub

Sub process_stock2()
    Dim StockBook As Worksheet
    Dim T_index As Integer
    Dim TableData, SumVolume As Variant
    Dim startprice, endprice As Double
    Dim Tickers(4096, 3) As String
    Dim SummaryData, TableData_Rng, Performers As Range

For Each StockBook In ThisWorkbook.Worksheets

'StockBook is a Range that contains all ticker values in the table
'TableData is a Range that contains the entire table
Set SummaryData = StockBook.Range("I1:L1")
SummaryData.Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Volume")

Set TableData_Rng = StockBook.Range("A1", StockBook.Range("A1").End(xlToRight).End(xlDown))
TableData_Rng.Sort Header:=xlYes, Key1:=TableData_Rng.Columns(1), Order1:=xlAscending, Key2:=TableData_Rng.Columns(2), Order2:=xlAscending


'Much Faster than process_stock2 put all the table data in an array and work from memory
TableData = TableData_Rng.Resize(TableData_Rng.Rows.Count + 1, TableData_Rng.Columns.Count).Value
T_index = 0
SumVolume = 0
For i = 2 To UBound(TableData, 1) - 1
    
    SumVolume = SumVolume + CVar(TableData(i, 7))
    
    If TableData(i, 1) <> TableData(i - 1, 1) Then
        startprice = CDbl(TableData(i, 6))
        
    Else
        If TableData(i, 1) = TableData(i + 1, 1) Then
        Else
            endprice = CDbl(TableData(i, 6))
            Tickers(T_index, 0) = TableData(i, 1)
            Tickers(T_index, 1) = endprice - startprice
            If startprice <> 0 Then
                Tickers(T_index, 2) = Tickers(T_index, 1) / startprice
            Else
                Tickers(T_index, 2) = 0
            End If
            Tickers(T_index, 3) = SumVolume
            T_index = T_index + 1
            SumVolume = 0
        End If
    End If
    
Next i
SummaryData.Resize(UBound(Tickers, 1)).Offset(1, 0).Value = Tickers




'Set a Range to the Newly Created Table
    Set SummaryData = StockBook.Range("I1").CurrentRegion
    Set Performers = StockBook.Range("O1:Q3")

    
'Format Percentage Change Column as a Percent
    SummaryData.Columns(2).Value = SummaryData.Columns(2).Value
    SummaryData.Columns(3).Value = SummaryData.Columns(3).Value
    SummaryData.Columns(3).NumberFormat = "0.00%"
    SummaryData.Columns(4).Value = SummaryData.Columns(4).Value

    For Each Difference In SummaryData.Columns(2).Cells
        If IsNumeric(Difference.Value) Then
            If Difference.Value > 0 Then
                Difference.Interior.ColorIndex = 4
            Else
                Difference.Interior.ColorIndex = 3
            End If
        End If
    Next Difference

'BONUS Performance Table
    Performers.Cells(1, 1) = "Greatest % Increase"
    Performers.Cells(2, 1) = "Greatest % Decrease"
    Performers.Cells(3, 1) = "Greatest Total Volume"
    
'Get the Maximum and Minimum % Change Stocks
'Similar to above I have sorted the data and retrieved the first and last rows to get Minimum and Maximum Respectively
'This only needs to be done once I am making the assumption that for any given sampling of stocks (i.e Alphabetically or By Year)
'The minimum will be the largest decrease and the maximum will be the largest increase
    SummaryData.Sort Key1:=SummaryData.Columns(3), Order1:=xlAscending, Header:=xlYes
    
'Populate the max from the last row of data in the table
    Performers.Cells(1, 2) = SummaryData(SummaryData.Rows.Count, 1)
    Performers.Cells(1, 3) = SummaryData(SummaryData.Rows.Count, 3)
'Populate the min from the first row of data in the table
    Performers.Cells(2, 2) = SummaryData(2, 1)
    Performers.Cells(2, 3) = SummaryData(2, 3)

'Format as a percentage
    Performers.Cells(1, 3).NumberFormat = "0.00%"
    Performers.Cells(2, 3).NumberFormat = "0.00%"
    
'Same Procedure as above to retrieve Max Volume
    SummaryData.Sort Key1:=SummaryData.Columns(4), Order1:=xlDescending, Header:=xlYes
    Performers.Cells(3, 2) = SummaryData(2, 1)
    Performers.Cells(3, 3) = SummaryData(2, 4)
    
'Put the Summary Table back in Alphabetical Order
    SummaryData.Sort Key1:=SummaryData.Columns(1), Order1:=xlAscending, Header:=xlYes
    
'Set all Objects and Variables to Null before moving on to the next book
    T_index = 0
    SumVolume = 0
    startprice = 0
    endprice = 0
    Erase Tickers
    Set TableData = Nothing
    Set SummaryData = Nothing
    Set TableData_Rng = Nothing
    Set Performers = Nothing
Next StockBook
End Sub
