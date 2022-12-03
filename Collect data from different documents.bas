Attribute VB_Name = "Module1"
Sub Button1_Click()

    'This macro is to get the data from different documents and collect them to the main database.
    
    Dim doc_path
    Dim pkg_code
    Dim i, j, k, p, o
    Dim SearchStr
    
    o = 49  'start row
    p = o - 2 'start row - the top 2 row(titles)
    j = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row + 1  'find the last row of this sheet
    
    Do Until p = j
        'Define the string to search in
        SearchStr = Cells(2 + p, 6).Value
        If Right(Cells(2 + p, 6), 3) = "lsx" Then
            If CInt(InStr(1, SearchStr, "-620-")) > 0 Then
                Application.ScreenUpdating = False
                'Define the destination
                k = Workbooks("Main Database").Sheets("620").Cells(2 + p, 6).Value
                'The link to open documents
                Cells(2 + p, 6).Hyperlinks(1).Follow
                'Send keys and agree to open the doc
                Application.SendKeys ("{LEFT}")
                Application.SendKeys ("{ENTER}")
                'Activate the open doc
                Worksheets("Information").Activate
                'Select the data in doc
                Range(Cells(3, 3), Cells(3, 3).End(xlDown)).Select
                Selection.Copy
                'Activate the main Database workbook
                Windows("Main Database").Activate
                'Activate the sheet
                Worksheets("620").Activate
                'Paste the data to destination
                Cells(2 + p, 14).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                Workbooks(k).Close SaveChanges:=False
                Application.ScreenUpdating = True
                
            
            
            ElseIf CInt(InStr(1, SearchStr, "-610-")) > 0 Then
                Application.ScreenUpdating = False
                k = Workbooks("Main Database").Sheets("610").Cells(2 + p, 6).Value
                Cells(2 + p, 6).Hyperlinks(1).Follow
                Application.SendKeys ("{LEFT}")
                Application.SendKeys ("{ENTER}")
                Worksheets("Information").Activate
                Range(Cells(3, 3), Cells(3, 3).End(xlDown)).Select
                Selection.Copy
                Windows("Main Database").Activate
                Worksheets("610").Activate
                Cells(2 + p, 14).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                Workbooks(k).Close SaveChanges:=False
                Application.ScreenUpdating = True
            
            
            ElseIf CInt(InStr(1, SearchStr, "-600-")) > 0 Then
            
                
                Application.ScreenUpdating = False
                k = Workbooks("Main Database").Sheets("600").Cells(2 + p, 6).Value
                Cells(2 + p, 6).Hyperlinks(1).Follow
                Application.SendKeys ("{LEFT}")
                Application.SendKeys ("{ENTER}")
                
                If Sheets(1).Name = "Information" Then
                    Worksheets("Information").Activate
                    Range(Cells(3, 3), Cells(3, 3).End(xlDown)).Select
                    Selection.Copy
                    Windows("Main Database").Activate
                    Worksheets("600").Activate
                    Cells(2 + p, 14).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=True
                    Workbooks(k).Close SaveChanges:=False
                    Application.ScreenUpdating = True
                    
                'If the documents have 2 Information sheets, copy the data of the sheets individually
                ElseIf Sheets(1).Name = "Information #1" Then
                    Worksheets("Information #1").Activate
                    Range(Cells(3, 3), Cells(3, 3).End(xlDown)).Select
                    Selection.Copy
                    Windows("Main Database").Activate
                    Worksheets("600").Activate
                    Cells(2 + p, 15).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=True
                        
                    Workbooks(k).Activate
                    Worksheets("Information #2").Activate
                    Range(Cells(3, 3), Cells(3, 3).End(xlDown)).Select
                    Selection.Copy
                    Windows("Main Database").Activate
                    Worksheets("600").Activate
                    Cells(2 + p, 55).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=True
                    Workbooks(k).Close SaveChanges:=False
                    Application.ScreenUpdating = True
                    
                ElseIf Sheets(1).Name = "Information #2" Then
                    Worksheets("Information #1").Activate
                    Range(Cells(3, 3), Cells(3, 3).End(xlDown)).Select
                    Selection.Copy
                    Windows("Main Database").Activate
                    Worksheets("600").Activate
                    Cells(2 + p, 55).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=True
                        
                    Workbooks(k).Activate
                    Worksheets("Information #2").Activate
                    Range(Cells(3, 3), Cells(3, 3).End(xlDown)).Select
                    Selection.Copy
                    Windows("Main Database").Activate
                    Worksheets("600").Activate
                    Cells(2 + p, 15).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=True
                    Workbooks(k).Close SaveChanges:=False
                    Application.ScreenUpdating = True
                End If
                
                
                
                
                
                
            End If
            
            
            
        Else
            
        End If
        
        p = p + 1
        'Auto save
        Workbooks("Main Database").Save
    
    Loop
    
     
    

End Sub
