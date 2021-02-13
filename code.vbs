Sub stock456()
Dim i As Long
Dim size456 As Double 'row count
Dim ticker456() As String
Dim open456 As Double
Dim close456 As Double
Dim vol456 As Double
Dim count456 As Long
Dim WS_Count As Integer
Dim II As Integer
'------------------------------------------------------------------------------
' Set WS_Count equal to the number of worksheets in the activeworkbook.
'------------------------------------------------------------------------------
WS_Count = ActiveWorkbook.Worksheets.Count
For II = 1 To WS_Count
'------------------------------------------------------------------------------
'headings
'------------------------------------------------------------------------------
    ActiveWorkbook.Worksheets(II).Range("I1").Value = "Ticker"
    ActiveWorkbook.Worksheets(II).Range("J1").Value = "Yearly Change"
    ActiveWorkbook.Worksheets(II).Range("K1").Value = "Percentage Change"
    ActiveWorkbook.Worksheets(II).Range("L1").Value = "Total Stock Volume"
    ActiveWorkbook.Worksheets(II).Range("P1").Value = "Ticker"
    ActiveWorkbook.Worksheets(II).Range("Q1").Value = "Value"
    ActiveWorkbook.Worksheets(II).Range("O2").Value = "Greatest % Increase"
    ActiveWorkbook.Worksheets(II).Range("O3").Value = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(II).Range("O4").Value = "Greatest Total Volume"
'------------------------------------------------------------------------------
'find the size of the new matrix of the yearly change and volume
'------------------------------------------------------------------------------
    size456 = WorksheetFunction.CountIfs(ActiveWorkbook.Worksheets(II).Range("B:B"), ">20140000")
    ReDim ticker456(size456) As String
'------------------------------------------------------------------------------
'initial values
'------------------------------------------------------------------------------
    ticker456(1) = ActiveWorkbook.Worksheets(II).Cells(2, 1).Value
    open456 = ActiveWorkbook.Worksheets(II).Cells(2, 3).Value
    vol456 = ActiveWorkbook.Worksheets(II).Cells(2, 7).Value
    count456 = 1
'------------------------------------------------------------------------------
'loop to find values and fill out the matrix of the yearly change and volume
'------------------------------------------------------------------------------
    For i = 2 To size456 Step 1
        ticker456(i) = ActiveWorkbook.Worksheets(II).Cells(i + 1, 1).Value
        If ticker456(i) = ticker456(i - 1) Then
            vol456 = vol456 + ActiveWorkbook.Worksheets(II).Cells(i + 1, 7).Value
            If open456 = 0 Then 'take care of divide by zero error later
                open456 = ActiveWorkbook.Worksheets(II).Cells(i + 1, 3).Value
            End If
            If i = size456 Then
                close456 = ActiveWorkbook.Worksheets(II).Cells(i + 1, 6).Value
                GoTo 1000
            End If
        Else
            close456 = ActiveWorkbook.Worksheets(II).Cells(i, 6).Value
1000
            count456 = count456 + 1
            ActiveWorkbook.Worksheets(II).Cells(count456, 9).Value = ticker456(i - 1)
            ActiveWorkbook.Worksheets(II).Cells(count456, 10).Value = close456 - open456
            If open456 <> 0 Then
                ActiveWorkbook.Worksheets(II).Cells(count456, 11).Value = (close456 - open456) / open456
                If ActiveWorkbook.Worksheets(II).Cells(count456, 11).Value < 0 Then
                    ActiveWorkbook.Worksheets(II).Cells(count456, 11).Interior.Color = RGB(255, 0, 0)
                Else
                    ActiveWorkbook.Worksheets(II).Cells(count456, 11).Interior.Color = RGB(0, 255, 0)
                End If
            Else
                ActiveWorkbook.Worksheets(II).Cells(count456, 11).Value = "N/A"
            End If
            open456 = ActiveWorkbook.Worksheets(II).Cells(i + 1, 3).Value
            ActiveWorkbook.Worksheets(II).Cells(count456, 12).Value = vol456
            vol456 = ActiveWorkbook.Worksheets(II).Cells(i + 1, 7).Value
        End If
    Next i
'------------------------------------------------------------------------------
'Use Excel Built-in functions to find max and min
'------------------------------------------------------------------------------
    ActiveWorkbook.Worksheets(II).Range("Q2").Value = WorksheetFunction.Max(ActiveWorkbook.Worksheets(II).Range("K:K"))
    ActiveWorkbook.Worksheets(II).Range("Q3").Value = WorksheetFunction.Min(ActiveWorkbook.Worksheets(II).Range("K:K"))
    ActiveWorkbook.Worksheets(II).Range("Q4").Value = WorksheetFunction.Max(ActiveWorkbook.Worksheets(II).Range("L:L"))
'------------------------------------------------------------------------------
'loop to find tickers that matches the max and min
'------------------------------------------------------------------------------
    For i = 2 To count456 Step 1
        If ActiveWorkbook.Worksheets(II).Cells(i, 11).Value = ActiveWorkbook.Worksheets(II).Range("Q2").Value Then
            ActiveWorkbook.Worksheets(II).Range("P2").Value = ActiveWorkbook.Worksheets(II).Cells(i, 9).Value
        End If
        If ActiveWorkbook.Worksheets(II).Cells(i, 11).Value = ActiveWorkbook.Worksheets(II).Range("Q3").Value Then
            ActiveWorkbook.Worksheets(II).Range("P3").Value = ActiveWorkbook.Worksheets(II).Cells(i, 9).Value
        End If
        If ActiveWorkbook.Worksheets(II).Cells(i, 12).Value = ActiveWorkbook.Worksheets(II).Range("Q4").Value Then
            ActiveWorkbook.Worksheets(II).Range("P4").Value = ActiveWorkbook.Worksheets(II).Cells(i, 9).Value
        End If
    Next i
    MsgBox (ActiveWorkbook.Worksheets(II).Name & " calculation is done")
Next II

End Sub
