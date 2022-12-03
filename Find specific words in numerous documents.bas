Attribute VB_Name = "Module1"
Sub Button1_Click()

    'This macro is to get the package codes in different documents and collect them into the main database.
    Dim doc_path
    Dim pkg_code
    Dim i, j, k
    

    i = 91 'Start row
    j = 189 'End row
    Do Until i = j
        'Judge the file format
        If Right(Cells(2 + i, 4), 3) = "lsx" Then
            Application.ScreenUpdating = False
            'Define the specific cells
            k = Workbooks("PACKAGE CODE").Sheets(1).Cells(2 + i, 4).Value
            Cells(2 + i, 4).Hyperlinks(1).Follow
            'Automatically send keys to open the doc
            Application.SendKeys ("{LEFT}")
            Application.SendKeys ("{ENTER}")
            'Find the start of data
            Cells(9, 3).Select
            Selection.Copy
            'Activate the main database
            Windows("PACKAGE CODE").Activate
            'Select the destination and paste
            Cells(2 + i, 8).Select
            ActiveSheet.Paste
            'Close the source document without saving it
            Workbooks(k).Close SaveChanges:=False
            Application.ScreenUpdating = True
            
        Else
            
        End If
        
        i = i + 1
        
    
    Loop
    
     
    
    
    
    
    
    
    
    
    
End Sub
