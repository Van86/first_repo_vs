Sub WallStreet():

Worksheet_count = ActiveWorkbook.Worksheets.Count

   'For ws = 1 To Worksheet_count
       Worksheets(ws).Activate

  ' Set an initial variable for holding the Ticker
        Dim Ticker As String

  ' Set an initial variable for holding the total volume
        Dim Total_Volume As Double
    
        Dim Stock_Table_Row As Integer
    
        Dim lastRow As Long
    
    'Moderate Double
        Dim yearChange As Double
        Dim percentChange As Double
        Dim yearClose As Double
        Dim yearOpen As Double
        Dim yearOpenRow As Long
    
    'Formulas
    'yearChange=yearClose -yearOpen
    'Ticker A hasyearOpenincells(2,3)
    'Ticker  as yearClose in Cells(263,6)
    'percentChange =Round((yearChange/yearOpen)*100,2)
    
    yearClose = Cells(I, 6).Value
    yearOpen = Cells(yearOpenRow, 3).Value
    yearChange = yearsOpen - yearClose
    
    'Assigning values to variables
    Stock_Table_Row = 2
    yearOpenRow = 2
     
    'Count row
    lastRow = Cells(Rows.Count, "A").End(xiup).Row
    
    'set headers for summary total
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    
    


  'Loop through all ticker symbols
        'For I = 2 To lastRow

    'Check if we are still within the same ticker, if it is not...
            'If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      'Set the Ticker
            'Ticker = Cells(I, 1).Value

      'Add to the Total Volume
        Total_Volume = Volume + Cells(I, 7).Value

      'Print the Ticker in the Summary Table
        Range("I" & Stock_Table_Row).Value = Ticker

      'Print the Brand Amount to the Summary Table
        Range("J" & Stock_Table_Row).Value = Total_Volume
        
        'Print Yearly Change in Summary Table
        
        Range("K" & Stock_Table_Row).Value = YearlyChange
        
        Calculate Percent Change
            If yearOpen = 0 And yearChange = 0 Then
                percentChange = 0
            ElseIf yearChange = 0 And yearOpen <> 0 Then
            percentChange = 0
            
            ElseIf yearChange <> 0 And yearOpen Then
            
            percentChange = 1E+99
            Else: percentChange = Round((yearChange / yearOpen) * 100, 2)
    
            
            End If
        
                percentChange = Round((yearChange / yearOpen) * 100, 2)
        
        Range("L" & Stock_Table_Row).Value = percentChange

      ' Add one to the summary table row
        Stock_Table_Row = Stock_Table_Row + 1
      
      ' Reset the Volume
        Total_Volume = 0
        
        'Reset Next Year
            yearOpenRow = I + 1

    If the cell immediately following a row is the same ticker symbol..
            'Else

      If ticker is equal then we are adding to the volume
            Total_Volume = Total_Volume + Cells(I, 7).Value

            End If

    Next I

Next ws


End Sub
