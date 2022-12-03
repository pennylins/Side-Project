Attribute VB_Name = "Module2"
Sub create600()

    'This macro is to create documents with data from database, embed the PDF in the document, and insert the customer part number.
    
    Dim sr, er, i, s As Integer
    Dim b, c, d, e As String
    Dim msg, style, title, response As Variant
    Dim fs As Object
    Dim bd As String
    Dim name As String
    Dim mk As String
    Dim j As Integer
    Dim cell As Object
    Dim basedie As String
    Dim pkgtype As String
    Dim sap_part(30, 30)
    Dim q As Integer
    Dim w As Integer
    Dim rw_sap As String
    Dim col_sap As String
    Dim u As Integer
    Dim x As Integer
    Dim cus_part As Variant
    Dim rw_cus As String
    Dim dat As String
    Dim doc_no As String
    Dim map As String
    Dim filetype As String
   
    'ReDim sap_part(30, 30)
    msg = "Please confirm the entries, if correct, please click 'Yes'. Press 'No' to revise your entries."
    style = vbYesNo + vbQuestion + vbDefaultButton1 + vbApplicationModal
    title = "Warning"
    b = ThisWorkbook.name
    c = Cells(2, 6) & "\"
    name = Cells(3, 6)
    e = ThisWorkbook.ActiveSheet.name
    s = 0
    sr = Cells(4, 6)
    er = Cells(5, 6)
    response = MsgBox(msg, style, title)
        
    If response = vbYes Then
        Range(Cells(sr, 2), Cells(er, 2)).Delete Shift:=xlToLeft
        Range(Cells(sr, 1), Cells(er, 1)).EntireRow.RemoveDuplicates Columns:=Array(1, 10), Header:=xlNo
        er = Cells(sr, 1).End(xlDown).Row
        For i = sr To er
          Set fs = CreateObject("scripting.filesystemobject")
         'Check if this doc exist
          If fs.fileexists(Trim(c & Cells(i, 10) & "-Rev" & Cells(i, 11) & ".xlsx")) Then
           'if yes, go to next row and skip re-generate doc
            GoTo flag1
          Else
            If Not fs.fileexists(Trim(Cells(i, 10) & ".xlsx")) Then
                Workbooks(b).Activate
                Worksheets(e).Activate
                bd = c & Cells(i, 48).Value
                mk = Cells(i, 49).Value
                basedie = Cells(i, 4).Value
                pkgtype = Cells(i, 5).Value
                dat = Cells(i, 12).Value
                doc_no = Cells(i, 10).Value
                map = c & Cells(i, 50).Value
                Range(Cells(i, 9), Cells(i, 50)).Select
                Selection.Copy
                template = Cells(i, 8).Value
                Application.DisplayAlerts = False
                Workbooks.Open c & Cells(i, 8) & ".xlsx"
                ''ActiveWorkbook.AutoSaveOn = False
                Application.DisplayAlerts = True
                Cells(2, 3).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                Cells(5, 3).Value = FormatDateTime(Cells(5, 3).Value, vbShortDate)
                Worksheets("Revision History").Cells(3, 2).Value = "A"
                Worksheets("Revision History").Cells(3, 5).Value = name
                Worksheets("Revision History").Cells(3, 4).NumberFormatLocal = "[$-en-GB]d mmmm yyyy;@"
                Worksheets("Revision History").Cells(3, 4).Value = FormatDateTime(dat, vbShortDate)
                Worksheets("Revision History").Cells(3, 3).Value = "NEW SPECIFICATION"
                If Worksheets(1).Cells(43, 3).Value <> "" Then
                    Worksheets("Bonding Diagram").Cells(3, 2).Value = "internal path" & Worksheets(1).Cells(41, 3).Value & vbCrLf & "internal path" & Worksheets(1).Cells(43, 3).Value
                Else
                    Worksheets("Bonding Diagram").Cells(3, 2).Value = "internal path" & Worksheets(1).Cells(41, 3).Value
                End If
                'identify filetype
                If CInt(InStr(1, bd, ".pdf")) > 0 Then
                        filetype = "Acrobat Reader DC.exe"
                ElseIf CInt(InStr(1, bd, ".dwg")) > 0 Then
                        filetype = "Launch dwgviewr.exe"
                End If
                'insert BD file
                If fs.fileexists(Trim(bd)) Then
                    Worksheets("Bonding Diagram").Activate
                    Cells(3, 3).Select
                    
                    ActiveSheet.OLEObjects.Add _
                    Filename:=bd, _
                    Link:=False, _
                    DisplayAsIcon:=True, _
                    IconFileName:=filetype, _
                    IconIndex:=0, _
                    IconLabel:=bd
                    
                End If
                
                filetype = ""
                If CInt(InStr(1, bd, ".pdf")) > 0 Then
                        filetype = "Acrobat Reader DC.exe"
                ElseIf CInt(InStr(1, bd, ".dwg")) > 0 Then
                        filetype = "Launch dwgviewr.exe"
                End If

                'Insert 2nd BD file. Use column of map file to store 2nd BD filename
                If fs.fileexists(Trim(map)) Then
                    
                    Cells(3, 4).Select
                    ActiveSheet.OLEObjects.Add _
                    Filename:=map, _
                    Link:=False, _
                    DisplayAsIcon:=True, _
                    IconFileName:=filetype, _
                    IconIndex:=0, _
                    IconLabel:=map
                End If
                Worksheets("Information").Activate

                Application.DisplayAlerts = True
                'insert marking template
                If fs.fileexists(Trim(c & mk)) Then
                    'get sap_part
                     Workbooks("6xx document summary_Macro").Activate
                     w = 1
                     x = 1
                     '1/28: change key value to "Document No." only
                     For Each cell In Range("J" & sr & ":" & "J" & er)
                         If cell.Value = doc_no Then
                            ''If cell.Offset(0, 1).Value = pkgtype Then
                                'ReDim Preserve sap_part(w, x)
                                 sap_part(w, 1) = cell.Offset(0, -8).Value
                                 sap_part(w, 2) = cell.Offset(0, -9).Value
                                'sap_part = Range(cell.Offset(0, -2), cell.Offset(0, -3)).Value
                                 w = w + 1
                            ''End If
                         End If
                     Next
                    'open marking template
                    Application.DisplayAlerts = False
                    Workbooks.Open c & mk
                    ''ActiveWorkbook.AutoSaveOn = False
                    Application.DisplayAlerts = True
                    j = 1
                    'deal with duplication of sap_part
                    'Call Get_sap_part(sap_part, w)
                    Do
                        'search "Assembly SAP Material Number"
                        Workbooks(mk).Activate
                        Worksheets(j).Activate
                        For Each cell In Range("B:B")
                            If cell.Value = "Assembly SAP Material Number" Then
                                rw_sap = cell.Row
                                y = 1
                                Do
                                    If cell.Offset(0, y).Value <> "Customer Part Number" Then
                                        y = y + 1
                                    Else
                                        rw_cus = y
                                    End If
                                Loop Until cell.Offset(0, y).Value = "Customer Part Number"
                                ''col_sap = cell.Column
                                q = 1
                                Do
                                    cell.Offset(q, 0) = sap_part(q, 1)
                                    cell.Offset(q, y) = sap_part(q, 2)
                                    q = q + 1
                                Loop Until q = w
                            End If
                        Next
                        Columns("B:B").Select
                        With Selection.Font
                            .name = "Calibri"
                            .Size = 11
                        End With
                        Worksheets(j).Copy After:=Workbooks(template).Sheets(2)
                        j = j + 1
                    Loop Until j = Workbooks(mk).Worksheets.Count + 1
                    Application.DisplayAlerts = False
                    Workbooks(mk).Close
                    Application.DisplayAlerts = True
                    Workbooks(template).Activate
                    Worksheets("Information").Activate
                End If
                Range(Cells(40, 3), Cells(46, 3)).Delete
                'Worksheets(1).Cells(2, 3).Select
                'Set font and row autofit for specific sheet
                u = 1
                For u = 1 To ActiveWorkbook.Sheets.Count
                    If Worksheets(u).name = "Information" Or Worksheets(u).name = "Revision History" Then
                        Worksheets(u).Activate
                        Columns("A:F").Select
                        With Selection.Font
                            .name = "Calibri"
                            .Size = 11
                        End With
                        With Selection
                            .VerticalAlignment = xlTop
                            .Orientation = 0
                            .AddIndent = False
                            .ShrinkToFit = False
                            .ReadingOrder = xlContext
                        End With
                        'Columns.AutoFit
                        Selection.EntireRow.AutoFit
                    End If
                    ''Worksheets(u).Range("B10").Select
                    Worksheets(u).Activate
                    ActiveSheet.Range("B3").Select
                Next
                Worksheets(1).Activate
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs c & Cells(3, 3) & "-Rev" & Cells(4, 3) & ".xlsx"
                ActiveWorkbook.Close (False)
                s = s + i
            Else
                Cells(i, 8).Font.Color = RGB(255, 0, 0)
                MsgBox "template: '" & Cells(i, 8) & "' doesn't exist! Please check the filename or create a template!", vbOKCancel, "Note"
            End If
          End If

flag1:
        Next i
        i = ""
        
    End If
    MsgBox "Row" & sr & "-" & er & " have been processed!"

End Sub
Public Function extractFileName(filePath)

    For i = Len(filePath) To 1 Step -1
        If Mid(filePath, i, 1) = "\" Then
        extractFileName = Mid(filePath, i + 1, Len(filePath) - i + 1)
        Exit Function
        End If
    Next

End Function
Public Function Get_sap_part(sap_part, w)
    
Dim i As Integer
Dim h As Integer
Dim sh_cnt As String
Dim j As Integer

sh_cnt = ActiveSheet.Index
Sheets.Add After:=ActiveSheet
Sheets(sh_cnt + 1).name = "temp"

i = 1
j = 1
Do
    Range("A" & i) = sap_part(i, j)
    Range("B" & i) = sap_part(i, j + 1)
    i = i + 1
    j = 1
Loop Until i = w + 1


Range("A1:" & "B" & i).Select
Selection.RemoveDuplicates Columns:=1

i = 1
j = 1
Do
    'ReDim Preserve sap_part(i, j + 1)
    sap_part(i, j) = Range("A" & i).Value
    sap_part(i, j + 1) = Range("B" & i).Value
    i = i + 1
    j = 1
Loop Until Range("A" & i).Value = ""

w = i - 1



Application.DisplayAlerts = False
ActiveSheet.Delete
Application.DisplayAlerts = True
    
End Function


