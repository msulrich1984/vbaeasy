sub byws
'to run script per worksheet rather than one-by-one
    For each ws in Multiple_year_stock_data

  '     Set an initial variable for holding the ticker
        Dim ticker As String

      ' Set an initial variable for holding the volume by ticker
        Dim volume As long
        volume = 0

    ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        Dim lastrow as long
        'Since there's a lot of entries, I'm assuming this is needed.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    ' Loop through all tickers
    For i = 2 To lastrow

        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'Set the ticker name
            ticker = Cells(i, 1).Value

          ' Add to the volume by ticker
            volume = ticker + Cells(i, 7).Value

          ' Print the ticker in the Summary Table
          Range("I" & Summary_Table_Row).Value = ticker

        ' Print the volume to the Summary Table
            Range("J" & Summary_Table_Row).Value = volume

          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the volume
        volume = 0

        ' If the cell immediately following a row is the same brand...
        Else

      '     Add to the Brand Total
        volume = volume + Cells(i, 3).Value

    End If

  Next i
Next ws 
End Sub
