Attribute VB_Name = "Module3"
Sub extractFileName()

    'This macro is to extract filename under the same directory/folder.
    Dim fs As Object
    Dim f, path, filetype As String
    Dim i As Integer
    
    On Error Resume Next
    Err.Clear
    old_file = Cells(3, "A") 'old file
    new_file = Cells(3, "B") 'new file
    filetype = Cells(3, "C") 'file format
    path = Cells(1, 2) & "\" 'folder destination
    i = 1
    j = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row + 1 'last row of sheet
    'myfile = Dir(path & old_file)
    Range(Cells(3, 1), Cells(20, 4)).ClearContents 'Clear the content
    
    
    f = Dir(path)
    While f <> ""
        Cells(i + 2, "A") = f
        'extract the file format
        Cells(i + 2, "C") = Right(f, Len(f) - InStr(1, f, ".") + 1)
        i = i + 1
        f = Dir
        'Make the columns same width
        Columns("a:a").AutoFit
        
    Wend
    'Save workbook
    ThisWorkbook.Save
    MsgBox ("DONE") 'Message when completion
    
    
If Err.Number <> 0 Then
    MsgBox Err.Description
Else
End If

    
    
End Sub





