# fix_date
fix date such as 108.09.19 to excel recognizable date


Sub fix_date()

    Dim fix_range As Range
    Dim cell As Range
    Dim myyear As Integer
    Dim mymonth As Integer
    Dim mydate As Integer

On Error GoTo errorhandle


'asking the user for the range they wish to fix
    Set fix_range = Application.InputBox("please select range.", "Range Selection", , , , , , 8)

'loop though the range to fix each cell

    For Each cell In fix_range
        
        myyear = Left(cell.Value, WorksheetFunction.Search(".", cell.Value) - 1) + 1911 ' make it west year
        mymonth = Mid(cell.Value, WorksheetFunction.Search(".", cell.Value) + 1, 2)
        mydate = Right(cell.Value, 2)
        
        cell = myyear & "/" & mymonth & "/" & mydate
        fix_range.NumberFormatLocal = "[$-zh-TW]e/m/d;@"
        
    Next cell
    
    Exit Sub
    
errorhandle:

        Select Case Err.Number
        
            Case 424
                Exit Sub
            Case Else
                MsgBox "Unfortunately, an error occurred."
                
        End Select
        
    
End Sub
