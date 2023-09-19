
Private Sub Impute_Click()


' -----------------------
' Define variables
' -----------------------
Dim i As Integer, nr As Integer


' -------------------------------------------------------------
' count number of non-blank rows in date column (ie. column 5)
' -------------------------------------------------------------
i = 14 'set the row-index for 1st datapoint
Do While Cells(i, 5) <> ""

    i = i + 1 'update row-pointer
    
Loop

'set the last non-blank row index
nr = i - 1


' ------------------------------------------------------------
' impute if missing or zero
' ------------------------------------------------------------
For i = 14 To nr

    'check if any value is missing or zero
    'Note: if yes, replace all values by zero
    If Cells(i, 6) = "" _
        Or Cells(i, 6) = 0 _
        Or Cells(i, 7) = "" _
        Or Cells(i, 7) = 0 _
    Then
        Cells(i, 6) = 0 'impute injection rate
        Cells(i, 7) = 0 'impute bottom hole pressure [bar]
        Cells(i, 8) = 0 'impute bottom hole pressure [psig]
    End If
    
    
    
Next i

MsgBox "Imputation completed."

End Sub
