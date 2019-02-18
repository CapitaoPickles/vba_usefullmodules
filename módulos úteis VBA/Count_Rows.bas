Attribute VB_Name = "Módulo2"
Function count_rows(wb As Workbook, ws As Worksheet, bla As Range) As Long
    
    Dim i, j, nrows As Long
    Dim cell As Range
    nrows = 0
    
    For Each cell In bla
    If Application.WorksheetFunction.CountBlank(cell) > 0 Then
    
    Exit For
    
    End If
    
    nrows = nrows + 1
  
    Next cell
    
    count_rows = nrows

End Function

