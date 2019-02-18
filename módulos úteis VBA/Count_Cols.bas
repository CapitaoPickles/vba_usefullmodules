Attribute VB_Name = "Módulo3"
Function count_cols(wb As Workbook, ws As Worksheet, id As Range)

    Dim i, j, ncols As Long
    Dim cell As Range
    nrows = 0
    
    For Each cell In id
    If Application.WorksheetFunction.CountBlank(cell) > 0 Then
    
    Exit For
    
    End If
    
    ncols = ncols + 1
  
    Next cell
    
    count_cols = ncols

End Function
