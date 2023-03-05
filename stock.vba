Sub stock()

'Declare variables
    Dim opening_price As Double
    Dim closing_price As Double
    Dim cumul_vol As LongLong
    Dim ticker_name As String
    Dim row_count As Long
    Dim output_row As Integer
    Dim headings_array As Variant
    Dim CondFormatRange As Range
    Dim Current As Worksheet
    
'Declare variables for bonus
    Dim max_gain As Double
    Dim max_loss As Double
    Dim max_vol As LongLong
    
    headings_array = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
'Label columns P and Q and cells O2:O4 for bonus
    With Sheets(1)
    .Range("P1:Q1").Value = Array("Ticker", "Value")
    .Range("O2:O4").Value = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
    .Range("Q2:Q3").NumberFormat = "0.00%"
    End With

For Each Current In Worksheets
    Current.Select

'Label columns I:L and conditional format column J
    Range("I1:L1").Value = headings_array
    Set CondFormatRange = Columns("J")
    With CondFormatRange
        .FormatConditions.Delete
        .FormatConditions.Add xlCellValue, xlLess, 0
        .FormatConditions.Add xlCellValue, xlGreater, 0
        .FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        .FormatConditions(2).Interior.Color = RGB(0, 255, 0)
    End With
    Range("J1").FormatConditions.Delete
    Columns("K").NumberFormat = "0.00%"
    
'set row counters
    row_count = 2
    output_row = 2

'Loop until next row is empty
    Do Until IsEmpty(Cells(row_count, 1))

    'set ticker starting values
        ticker_name = Cells(row_count, 1)
        opening_price = Cells(row_count, 3)
        cumul_vol = 0

    'Loop while ticker name is the same
        Do While Cells(row_count, 1) = ticker_name
        cumul_vol = cumul_vol + Cells(row_count, 7)
        row_count = row_count + 1
        Loop

    'After ticker name changes
        closing_price = Cells(row_count - 1, 6).Value
        Cells(output_row, 9).Value = ticker_name
        Cells(output_row, 10).Value = opening_price - closing_price
        Cells(output_row, 11).Value = (opening_price - closing_price) / opening_price
        Cells(output_row, 12).Value = cumul_vol

    'Increment output_row
        output_row = output_row + 1

    Loop
    
    'Autofit column width
        Current.Columns("A:Q").AutoFit
        
    'Get current maxs and min for bonus
        max_gain = Application.WorksheetFunction.Max(Columns("K"))
        max_loss = Application.WorksheetFunction.Min(Columns("K"))
        max_vol = Application.WorksheetFunction.Max(Columns("L"))
        
    'Update table with overall maxs and min for bonus
        If max_gain > Sheets(1).Cells(2, 17).Value Then
            Sheets(1).Cells(2, 17).Value = max_gain
            Sheets(1).Cells(2, 16).Value = Range("I" & WorksheetFunction.Match(max_gain, Current.Range("K:K"), 0)).Value
        End If
        If max_loss < Sheets(1).Cells(3, 17).Value Then
            Sheets(1).Cells(3, 17).Value = max_loss
            Sheets(1).Cells(3, 16).Value = Range("I" & WorksheetFunction.Match(max_loss, Current.Range("K:K"), 0)).Value
        End If
        If max_vol > Sheets(1).Cells(4, 17).Value Then
            Sheets(1).Cells(4, 17).Value = max_vol
            Sheets(1).Cells(4, 16).Value = Range("I" & WorksheetFunction.Match(max_vol, Current.Range("L:L"), 0)).Value
        End If
        
Next Current
        
        Sheets(1).Columns("A:Q").AutoFit
    
End Sub
