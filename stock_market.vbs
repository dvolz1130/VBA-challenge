Attribute VB_Name = "stock_market"
Sub stock_market()

    ' Variable to hold the ticker name
    Dim ticker_name As String

    ' Keep track of the location for each ticker name in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Variables to hold beginning year open, end of year close, and Yearly_change
    Dim begin_year_open As Double
    Dim end_year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_vol As Double
    yearly_change = 0
    percent_change = 0
    total_vol = 0
    
    ' Set begin_year_open variable
    begin_year_open = Range("C2").Value
    'MsgBox ("begin year open is " & begin_year_open)
    
    ' Variables to hold last row and last column
    Dim LRow As Double
    Dim LCol As Double
    
    LRow = Cells(Rows.Count, 1).End(xlUp).Row
    LCol = Cells(1, Columns.Count).End(xlToLeft).Column

    ' Put Headers and keep location for Ticker symbol, Yearly change, Percent change, and Total stock volume
    TS = LCol + 2
    YC = LCol + 3
    PC = LCol + 4
    TSV = LCol + 5
    Cells(1, TS).Value = "Ticker Symbol"
    Cells(1, TS).Font.Bold = True
    Cells(1, YC).Value = "Yearly Change"
    Cells(1, YC).Font.Bold = True
    Cells(1, PC).Value = "Percent Change"
    Cells(1, PC).Font.Bold = True
    Cells(1, TSV).Value = "Total Stock Volume"
    Cells(1, TSV).Font.Bold = True

    For x = 2 To LRow

        'Checking that ticker names are different, if so, set ticker_name
        If Cells(x + 1, 1) <> Cells(x, 1) Then
            
            ' set ticker name
            ticker_name = Cells(x, 1).Value
        
            ' Print the ticker name in the Ticker symbol column
            Cells(Summary_Table_Row, TS).Value = ticker_name
            
            ' Add to total_vol
            total_vol = total_vol + Range("G" & x).Value
            
            'Set end_year_close variable
            end_year_close = Range("F" & x).Value
            ' MsgBox ("end year close is " & end_year_close)
            
            'Getting Yearly change of stock
            If begin_year_open < end_year_close Then
            
                ' If begin_year_open is less, its a postive change
                yearly_change = end_year_close - begin_year_open
                ' MsgBox (yearly_change)
                Cells(Summary_Table_Row, YC).Value = yearly_change
                Cells(Summary_Table_Row, YC).Interior.ColorIndex = 4
                
            ElseIf begin_year_open > end_year_close Then
            
                ' If begin_year_open is greater, its a negitive change
                yearly_change = end_year_close - begin_year_open
                ' MsgBox (yearly_change)
                Cells(Summary_Table_Row, YC).Value = yearly_change
                Cells(Summary_Table_Row, YC).Interior.ColorIndex = 3
                
            Else
                
                ' Begin year open and end year close are equal
                Cells(Summary_Table_Row, YC).Value = 0
            
            ' Set yearly change back to 0
            yearly_change = 0
            
            End If
            
            ' Getting Percent changed
            percent_change = (end_year_close - begin_year_open) / end_year_close
            Cells(Summary_Table_Row, PC).Value = percent_change
            Columns(PC).NumberFormat = "0.00%"
            
            ' Add total volume to Total Stock Volume column
            Cells(Summary_Table_Row, TSV).Value = total_vol
            
            ' Add 1 to the summary table
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset total_vol
            total_vol = 0
            
            ' Reset begin_year_open variable to new open value
            begin_year_open = Range("C" & (x + 1)).Value
            ' MsgBox ("begin year open is " & begin_year_open)
            
        ' If the ticker symbols are the same, need to add to total_vol
        Else
            total_vol = total_vol + Range("G" & x).Value
        End If
                
    Next x
    
' Auto fit all new columns
Columns(TS).EntireColumn.AutoFit
Columns(YC).EntireColumn.AutoFit
Columns(PC).EntireColumn.AutoFit
Columns(TSV).EntireColumn.AutoFit

End Sub
