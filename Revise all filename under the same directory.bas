Attribute VB_Name = "Module2"
Sub renamefile()

    'This macro is to rename the files.

    Dim fs As Object 'filesystem
    Dim old_file, new_file, path, filetype As String
    Dim i As Integer
    

    On Error Resume Next
    Err.Clear
    old_file = Cells(3, "A") 
    new_file = Cells(3, "B") 
    filetype = Cells(3, "C") 
    path = Cells(1, 2) & "\" 'location
    i = 1 'counter
    j = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row + 1 'last row of sheet
    Range(Cells(3, "D"), Cells(3, "D").End(xlDown)).Clear  'Clean the data from the last usage
    
    Do
        Set fs = CreateObject("scripting.filesystemobject")
        If fs.fileexists(Trim(path & Cells(i + 2, "A"))) Then
            'Rename the filename based on the B2 and C2 cells
            Name path & Cells(i + 2, "A") As path & Cells(i + 2, "B") & Cells(i + 2, "c")
            'Show the status in D2 cells
            Cells(i + 2, "D") = "Done"
            
                    
        Else
            'Return status if the file doesn't exist
            Cells(i + 2, "D") = "The file doesn't exist."
        End If
    i = i + 1 'Loop counts +1
    Loop Until i = j
    
    ThisWorkbook.Save
    MsgBox ("DONE")
    
    
If Err.Number <> 0 Then
    MsgBox Err.Description
Else
End If

    
    
End Sub




