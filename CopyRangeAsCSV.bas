Sub CopyRangeAsCSV()
' Excel - VBA (untested on Windows)
' Copy selected range as CSV-formatted text

    Dim rng As Range
    Dim row As Range
    Dim cell As Range
    Dim csvData As String
    Dim csvRowData As String
    Dim clipboard As MSForms.DataObject
    
    ' Check if a range is selected
    If TypeName(Selection) = "Range" Then
        Set rng = Selection
    Else
        MsgBox "Please select a range to copy as CSV."
        Exit Sub
    End If
    
    ' Loop through each row and cell in the selected range
    For Each row In rng.Rows
        csvRowData = ""
        
        For Each cell In row.Cells
            ' Append cell value to csvRowData
            csvRowData = csvRowData & cell.Value & ","
        Next cell
        
        ' Remove the last comma from the row data
        csvRowData = Left(csvRowData, Len(csvRowData) - 1)
        
        ' Append the row data to csvData
        csvData = csvData & csvRowData & vbLf
    Next row
    
    ' Copy the CSV data to the clipboard
    Set clipboard = New MSForms.DataObject
    clipboard.SetText csvData
    clipboard.PutInClipboard
    
    MsgBox "Copied selected range as CSV-formatted text."

End Sub
