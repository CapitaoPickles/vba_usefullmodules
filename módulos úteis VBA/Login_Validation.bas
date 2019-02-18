Attribute VB_Name = "Módulo4"
Function login_page(user As Variant, password As Variant) As Boolean

    Dim valid As Boolean
    valid = False
    
    If IsError(Application.VLookup(user, ThisWorkbook.Worksheets("aut_page").Range("A:B"), 2, 0)) Then GoTo FIM
    
    
    If Application.WorksheetFunction.VLookup(user, ThisWorkbook.Worksheets("aut_page").Range("A:B"), 2, 0) = password Then
        
        valid = True
    Else
        valid = False
        
    End If

FIM:
    login_page = valid
    

End Function

