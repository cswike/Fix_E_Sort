Attribute VB_Name = "Fix_E_Sort"
Sub Fix_E_Sort()
' Allows part #s with the format [number]E[number] to be sorted as alphanumeric strings instead of numbers.
'
' Excel reads part #s like "9E12" as a numeric value using scientific notation.
' This is a problem when sorting a list,
' as #s like "9E12" will be sorted as numbers rather than alphanumerically -
' eg a sorted list might look like:
' {1, 900, 9E12, 9A12, 9B12, 9C12, 9D12, 9F12}
' where we really want the 9E12 to come after 9D12.
'
' This macro adds a helper column and appends a single parenthesis ( after the E for part #s with this format.
' Since ( is a special character, it gets sorted before any numbers or letters,
' and this also forces Excel to read it as a string instead of a number in E-notation.
' Also, ( doesn't have any mathematical operations assigned to it, so it should be safe to use.
'
' Once the macro finishes, just sort by the helper column and your list should be sorted alphanumerically like you wanted.
' You can delete the helper column when your list is sorted the way you want.
'

    Dim strVal As String
    Dim sortCol As String
    Dim sortColNo As Long
    Dim lRow As Long
    Dim E_loc As Long
    
    ' Get column with data to be sorted
    sortCol = InputBox("Enter the column you need to sort (column letter)", "Column to sort", "A")
    If sortCol = "" Then Exit Sub
    While Not IsLetter(sortCol)
        sortCol = InputBox("Error: not a letter. Enter the column you need to sort (column letter)", "Column to sort", "A")
        If sortCol = "" Then Exit Sub
    Wend
    
    ' Get column # and make sure column is not too large
    sortColNo = GetColNo(sortCol)
    If sortColNo = 0 Then
        MsgBox "Sorry, can't sort columns larger than ZZ. Please move your data column or select a different column to sort."
        Exit Sub
    End If
    
    ' Check for headers
    headerYN = MsgBox("Does your data have headers?", vbYesNo)
    If headerYN = vbYes Then fRow = 2 Else fRow = 1
    
    ' Get last row
    lRow = Cells(Rows.Count, sortColNo).End(xlUp).Row
    
    ' Insert helper column
    Columns(1).Insert
    Columns(1).NumberFormat = "@"
    If headerYN = vbYes Then Range("A1").Formula = "helper column"
    
    ' Add 1 to sort column to account for new helper row
    sortColNo = sortColNo + 1
    
    ' Loop through each cell in data column
    For Each c In Range(Cells(fRow, sortColNo), Cells(lRow, sortColNo))
    
        strVal = CStr(c.Formula)
        
        If c.Formula Like "*[0-9]E[0-9]*" Then
            E_loc = InStr(1, strVal, "E", vbTextCompare)
            Cells(c.Row, 1).Formula = Left(strVal, E_loc) & "(" & Mid(strVal, E_loc + 1)
        Else
            Cells(c.Row, 1).Formula = c.Formula
        End If
        
    Next c
    
    MsgBox "Done. Use the 'helper column' (A) to sort your data alphanumerically. You can delete column A when finished."

End Sub

Function IsLetter(r As String) As Boolean
    Dim x As String
    Dim Counter As Integer
    For Counter = 1 To Len(r)
        x = UCase(Mid(r, Counter, 1))
        IsLetter = Asc(x) > 64 And Asc(x) < 91
        If IsLetter = False Then Exit For
    Next
End Function

Function GetColNo(Col As String) As Integer
' Returns 0 if column is greater than ZZ (even though Excel can go up to XFD)
' ...but really, isn't 702 columns enough for you?
    
    If Len(Col) = 1 Then
        GetColNo = Asc(UCase(Col)) - 64
    ElseIf Len(Col) = 2 Then
        Dim L1 As Integer
        Dim L2 As Integer
        L1 = Asc(UCase(Left(Col, 1))) - 64
        L2 = Asc(UCase(Right(Col, 1))) - 64
        GetColNo = (26 * L1) + L2
    Else
        GetColNo = 0
    End If

End Function
