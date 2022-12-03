Attribute VB_Name = "Module2"
Sub insertmarking()

    'This macro is to insert the actual marking into the blank spec.
    
    Dim ss, main As String
    Dim c, e, j, i, s, r, t, sr, er, q, w, y, row_sap, col_sap As Integer
    Dim row_cus, col_cus As Integer
    Dim myfile, sap, cus_part, top_line, bottom_line, max_digit
    Dim fs As Object
    Dim cell As Object
    Dim msg, style, title, response As Variant
    
    
    
    'message to confirm to run macro
    msg = "Please confirm the entries, if correct, please click 'Yes'. Press 'No' to revise your entries."
    style = vbYesNo + vbQuestion + vbDefaultButton1 + vbApplicationModal
    title = "Warning"
    
    ss = "Search System"
    main = "Axxxxx All Marking(with formula)"
    e = ThisWorkbook.ActiveSheet.name
    c = Cells(2, 2) & "\"
    sr = Cells(1, 4)
    er = Cells(2, 4)
    response = MsgBox(msg, style, title)
    s = 0
    
    Application.ScreenUpdating = False
    
    'PROCEDURE: 
    'Remove duplicate customer part number
    'Confirm if the document exists
    'Open doc
    'Check if it has 1 page or 2 pages of marking data
    'Judge how many lines of marking
    'Get the data from database
    'Add the customer part number and marking in the table of doc
    'If it has bottomside marking, judge how many lines of marking in bottomside
    'Save and close the doc
    

    On Error Resume Next
    Err.Clear

    If response = vbYes Then
    
        Set fs = CreateObject("scripting.filesystemobject")
        'Remove duplicate customer part number
        Range(Cells(sr, 1), Cells(er, 1)).EntireRow.RemoveDuplicates Columns:=Array(1, 5), Header:=xlNo
        er = Cells(sr, 1).End(xlDown).Row
        For i = sr To er
            'Confirm if the document exists
            If fs.fileexists(Trim(c & Cells(i, 1) & "-Rev" & Cells(i, 2) & ".xlsx")) Then
                Workbooks(main).Activate
                Worksheets(ss).Activate
                myfile = Cells(i, 1) & "-Rev" & Cells(i, 2)
                sap = Cells(i, 3).Value
                cus_part = Cells(i, 5).Value
                top_line = Cells(i, 45).Value
                bottom_line = Cells(i, 46).Value
                max_digit = Cells(i, 47).Value
                Application.DisplayAlerts = False
                'Open doc
                Workbooks.Open c & myfile & ".xlsx"
                Application.DisplayAlerts = True
                '=====================TopSideMarking Judgement===========================
                 If Worksheets(3).name = "Top Side Marking" Then
                        Workbooks(myfile).Activate
                        
                        '==============if topline=3==============
                        q = 0
                        w = 0
                        If top_line = 3 Then
                        'find the customer part no. in the spec(navigation)
                            Worksheets("Top Side Marking").Activate
                             For Each cell In Range("E6:N100")
                                    If cell.Value = cus_part Then
                                       cell.Select
                                       
                                    End If
                                  
                            Next
                            
                            
                            'Add 2 rows
                             Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(2, 0)).EntireRow.Insert
                             
                             'activecell=vlookup(custom part no., marking doc
                             '44=top side marking columns of start
                             Do
                                ActiveCell.Offset(w, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 44 + q, False)
                                ActiveCell.Offset(w + 1, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 56 + q, False)
                                ActiveCell.Offset(w + 2, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 68 + q, False)
                                q = q + 1
            
                            Loop Until q = max_digit
                            
                    '==============if topline=4 ==============
                        ElseIf top_line = 4 Then
                            Worksheets("Top Side Marking").Activate
                            For Each cell In Range("E6:N100")
                                    If cell.Value = cus_part Then
                                       cell.Select
                                       
                                    End If
                                  
                            Next
                            
                            'Add 3 rows
                            Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(3, 0)).EntireRow.Insert
                            'activecell=vlookup(custom part no., marking doc
                            Do
                                ActiveCell.Offset(w, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 44 + q, False)
                                ActiveCell.Offset(w + 1, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 56 + q, False)
                                ActiveCell.Offset(w + 2, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 68 + q, False)
                                ActiveCell.Offset(w + 3, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 80 + q, False)
                             
                                q = q + 1
            
                            Loop Until q = max_digit
                               
                               
                       '==============if topline=5==============
                        ElseIf top_line = 5 Then
                            Worksheets("Top Side Marking").Activate
                            For Each cell In Range("E6:N100")
                                    If cell.Value = cus_part Then
                                       cell.Select
                                       
                                    End If
                                  
                            Next
                            Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(4, 0)).EntireRow.Insert
                            Do
                                ActiveCell.Offset(w, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 44 + q, False)
                                ActiveCell.Offset(w + 1, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 56 + q, False)
                                ActiveCell.Offset(w + 2, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 68 + q, False)
                                ActiveCell.Offset(w + 3, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 80 + q, False)
                                ActiveCell.Offset(w + 4, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 92 + q, False)
                                
                                q = q + 1
            
                            Loop Until q = max_digit
                        End If
                        
                        'Select the marking range and make it centered
                        Range(ActiveCell.Offset(0, -1), ActiveCell.Offset(top_line - 1, max_digit * -1)).Select
                        With Selection
                             .HorizontalAlignment = xlCenter
                        End With
                        Range("b2").Select
                        
                        '=====================BottomSideMarking===========================
                        
                        If Worksheets(4).name = "Bottom Side Marking" Then
                    
                            ' r & t = counter
                            
                            r = 0
                            t = 0
                        
                            
                            'Judge line of marking
                            If bottom_line = 3 Then
                                 Workbooks(myfile).Activate
                                 Worksheets("Bottom Side Marking").Activate
                                 For Each cell In Range("E6:N100")
                                    If cell.Value = cus_part Then
                                       cell.Select
                                       
                                    End If
                                  
                                Next
                                 
                                
                    
                                 Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(2, 0)).EntireRow.Insert
                                 Do
                                    ActiveCell.Offset(t, max_digit * -1 + r).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 104 + r, False)
                                    ActiveCell.Offset(t + 1, max_digit * -1 + r).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 116 + r, False)
                                    ActiveCell.Offset(t + 2, max_digit * -1 + r).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 128 + r, False)
                                    r = r + 1
                
                                Loop Until r = max_digit
                            
                            End If
                            
                            Range(ActiveCell.Offset(0, -1), ActiveCell.Offset(bottom_line - 1, max_digit * -1)).Select
                            With Selection
                                 .HorizontalAlignment = xlCenter
                            End With
                            Range("b2").Select
                            
                            
                            
                         
                        End If
                        
                    '==================if the doc only have "Marking" sheet======================
                    'if topline=3
                    ElseIf Worksheets(3).name = "Marking" Then
                         q = 0
                        w = 0
                        If top_line = 3 Then
                            Worksheets("Marking").Activate
                             For Each cell In Range("E6:N100")
                                    If cell.Value = cus_part Then
                                       cell.Select
                                       
                                    End If
                                  
                            Next
                            
                            
                            
                             Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(2, 0)).EntireRow.Insert
                             Do
                                ActiveCell.Offset(w, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 44 + q, False)
                                ActiveCell.Offset(w + 1, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 56 + q, False)
                                ActiveCell.Offset(w + 2, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 68 + q, False)
                                q = q + 1
            
                            Loop Until q = max_digit
                            
                    '==============if topline=4 ==============
                        ElseIf top_line = 4 Then
                            Worksheets("Marking").Activate
                            For Each cell In Range("E6:N100")
                                    If cell.Value = cus_part Then
                                       cell.Select
                                       
                                    End If
                                  
                            Next
                            
                            
                            Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(3, 0)).EntireRow.Insert
                            
                            Do
                                ActiveCell.Offset(w, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 44 + q, False)
                                ActiveCell.Offset(w + 1, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 56 + q, False)
                                ActiveCell.Offset(w + 2, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 68 + q, False)
                                ActiveCell.Offset(w + 3, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 80 + q, False)
                             
                                q = q + 1
            
                            Loop Until q = max_digit
                               
                               
                       '==============if topline=5==============
                        ElseIf top_line = 5 Then
                            Worksheets("Marking").Activate
                            For Each cell In Range("E6:N100")
                                    If cell.Value = cus_part Then
                                       cell.Select
                                       
                                    End If
                                  
                            Next
                                
                                Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(4, 0)).EntireRow.Insert
                                Do
                                    ActiveCell.Offset(w, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 44 + q, False)
                                    ActiveCell.Offset(w + 1, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 56 + q, False)
                                    ActiveCell.Offset(w + 2, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 68 + q, False)
                                    ActiveCell.Offset(w + 3, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 80 + q, False)
                                    ActiveCell.Offset(w + 4, max_digit * -1 + q).Value = Application.WorksheetFunction.VLookup(cus_part, Workbooks(main).Sheets(ss).Range("$E:$EL"), 92 + q, False)
                                    
                                    q = q + 1
                
                                Loop Until q = max_digit
                            End If
                        
                        Range(ActiveCell.Offset(0, -1), ActiveCell.Offset(top_line - 1, max_digit * -1)).Select
                        With Selection
                             .HorizontalAlignment = xlCenter
                        End With
                        Range("b2").Select
                    
                    End If
                    Sheets("Information").Activate          'Activate the first sheet of doc
                    Application.DisplayAlerts = False
                    ActiveWorkbook.Save
                    ActiveWorkbook.Close (True)             'Save and close doc
                    Application.DisplayAlerts = True
                    s = s + i
            
            Else
            'if the doc doesn't exist
                MsgBox (myfile & "is not existed in the" & c)
            End If
       
        
            
    
        
        Next
        MsgBox ("Process is done")
    Else
        
    End If
    
    
    
    
If Err.Number <> 0 Then
    MsgBox Err.Description
Else

End If
End Sub
