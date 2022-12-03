Attribute VB_Name = "Module2"
Sub Button1_Click()

    'This macro is to get the address of the online documents.
    
    Dim i 'Start row
    Dim mypath, url
    Dim ie As Object
    
        
    For i = 2 To 158
        url = Cells(i, 2).Hyperlinks(1).Address 'get the address in the hidden hyperlinks
        applicatoin.DisplayStatusBar = True
        Application.ScreenUpdating = False
        
        Set ie = CreateObject("internetexplorer.application")
        
        With ie
            .Visible = True 'open IE explorer
            .navigate url
        End With
        
        
        
    Next
    
End Sub
