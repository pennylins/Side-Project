Attribute VB_Name = "Module3"
Sub change600()

' This macro is to Update docs and revise the content automatically based on the database.
' Penny created in 2020
' Claire modified on 2021/3/3
' Penny modified on 2021/5/10
    
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
    Dim rev As String
    Dim r As Integer
    Dim change_detail As String
    Dim sht_marking()

    msg = "Please confirm the entries, if correct, please click 'Yes'. Press 'No' to revise your entries."
    style = vbYesNo + vbQuestion + vbDefaultButton1 + vbApplicationModal
    title = "Warning"
    b = ThisWorkbook.name
    c = Cells(2, 2) & "\"
    name = Cells(3, 2)
    e = ThisWorkbook.ActiveSheet.name
    s = 0
    sr = Cells(4, 2)
    er = Cells(5, 2)
    response = MsgBox(msg, style, title)
        
    If response = vbYes Then
        If Not Cells(sr, 9) = "Assembly Specification" Then
            Range(Cells(sr, 2), Cells(er, 2)).Delete Shift:=xlToLeft
            Range(Cells(sr, 1), Cells(er, 1)).EntireRow.RemoveDuplicates Columns:=Array(1, 10), Header:=xlNo
            er = Cells(sr, 1).End(xlDown).Row
        End If
        For i = sr To er
          Set fs = CreateObject("scripting.filesystemobject")
         'Check if this doc exist
          If fs.fileexists(Trim(c & Cells(i, 10) & "-Rev" & Cells(i, 11) & ".xlsx")) Then
           'if yes, go to next row and skip to re-generate doc
            GoTo flag1
          Else
            If Not fs.fileexists(Trim(Cells(i, 10) & ".xlsx")) Then
                Workbooks(b).Activate
                Worksheets(e).Activate
                change_detail = Cells(i, 47).Value
                bd = c & Cells(i, 48).Value
                mk = Cells(i, 49).Value
                basedie = Cells(i, 4).Value
                pkgtype = Cells(i, 5).Value
                dat = Cells(i, 12).Value
                doc_no = Cells(i, 10).Value
                map = c & Cells(i, 50).Value
                rev = Cells(i, 11).Value
                Range(Cells(i, 9), Cells(i, 50)).Select
                Selection.Copy
                template = Cells(i, 8).Value
                Application.DisplayAlerts = False
                Workbooks.Open c & template & ".xlsx"
                ''ActiveWorkbook.AutoSaveOn = False
                Application.DisplayAlerts = True
                Cells(2, 3).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                'Calculate the row numbers of column "Rev" in revision history
                r = 1
                Do
                    r = r + 1
                Loop Until Worksheets("Revision History").Range("B" & r) = ""
                'Fill up Revision History sheet
                Worksheets("Revision History").Cells(r, 2).Value = rev
                Worksheets("Revision History").Cells(r, 5).Value = name
                Worksheets("Revision History").Cells(r, 4).NumberFormatLocal = "[$-en-GB]d mmmm yyyy;@"
                Worksheets("Revision History").Cells(r, 4).Value = FormatDateTime(dat, vbShortDate)
                'Penny: put change_detail to cells(i,47) to replace the inputbox (2021/05/10 Penny)
                Worksheets("Revision History").Cells(r, 3).Value = change_detail
                Worksheets("Bonding Diagram").Activate
                Worksheets("Bonding Diagram").Shapes.SelectAll
                Selection.Delete
                'Fill up Bonding Diagram sheet
                If Worksheets(1).Cells(43, 3).Value <> "" Then
                    Worksheets("Bonding Diagram").Cells(3, 2).Value = "Internal path" & Worksheets(1).Cells(41, 3).Value & vbCrLf & "Internal path" & Worksheets(1).Cells(43, 3).Value
                Else
                    Worksheets("Bonding Diagram").Cells(3, 2).Value = "Internal path" & Worksheets(1).Cells(41, 3).Value
                End If
                Range("D4").Select
'                'identify filetype
                If CInt(InStr(1, bd, ".pdf")) > 0 Then
                        filetype = "Acrobat Reader DC.exe"
                ElseIf CInt(InStr(1, bd, ".dwg")) > 0 Then
                        filetype = "Launch dwgviewr.exe"
                End If
                'insert BD file
                If fs.fileexists(Trim(bd)) Then
                    'Worksheets("Bonding Diagram").Activate
                    Cells(3, 3).Select
                    'Application.DisplayAlerts = False
                    ActiveSheet.OLEObjects.Add _
                    Filename:=bd, _
                    Link:=False, _
                    DisplayAsIcon:=True, _
                    IconFileName:=filetype, _
                    IconIndex:=0, _
                    IconLabel:=bd
                Else
                    MsgBox ("Couldn't find the correct bonding diagram in " & doc_no & ", please check if the file is opened or missing.")
                End If
                filetype = ""
                If CInt(InStr(1, bd, ".pdf")) > 0 Then
                        filetype = "Acrobat Reader DC.exe"
                ElseIf CInt(InStr(1, bd, ".dwg")) > 0 Then
                        filetype = "Launch dwgviewr.exe"
                End If
                'Insert 2nd BD file. Use column of map file to store 2nd BD filename
                If fs.fileexists(Trim(map)) Then
                    'Worksheets("Bonding Diagram").Activate
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
                     Workbooks("6xx document summary_Macro").Activate
                     w = 1
                     x = 1
                     '1/28: change key value to "Document No." only
                     For Each cell In Range("J" & sr & ":" & "J" & er)
                         If cell.Value = doc_no Then
                            'If cell.Offset(0, 1).Value = pkgtype Then
                                'ReDim Preserve sap_part(w, x)
                                 sap_part(w, 1) = cell.Offset(0, -8).Value
                                 sap_part(w, 2) = cell.Offset(0, -9).Value
                                'sap_part = Range(cell.Offset(0, -2), cell.Offset(0, -3)).Value
                                 w = w + 1
                            'End If
                         End If
                     Next
                    'open marking template
                    Application.DisplayAlerts = False
                    Workbooks.Open c & mk
                    ''ActiveWorkbook.AutoSaveOn = False
                    'Application.DisplayAlerts = True
                    j = 1
                    'Deal with duplication of sap_part
                    'Call Get_sap_part(sap_part, w)
                    p = 3
                    Do
                        If Workbooks(template).Sheets(p).name = "Top Side Marking" Or Workbooks(template).Sheets(p).name = "Bottom Side Marking" Or Workbooks(template).Sheets(p).name = "Marking" Then
                            ReDim Preserve sht_marking(p - 2)
                            sht_marking(p - 2) = Workbooks(template).Sheets(p).name
                            p = p + 1
                        End If
                    Loop Until Workbooks(template).Sheets(p).name = "Revision History"
                    Do
                        Workbooks(template).Worksheets(sht_marking(p - 3)).Delete
                            p = p - 1
                    Loop Until Workbooks(template).Worksheets(3).name = "Revision History"
                    Do
                        'search "Assembly SAP Material Number"
                        Workbooks(mk).Activate
                        'Worksheets(j).Select
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
                        Workbooks(mk).Activate
                        Workbooks(mk).Worksheets(j).Copy After:=Workbooks(template).Sheets(2)
                        j = j + 1
                    Loop Until j = Workbooks(mk).Worksheets.Count + 1
                    Workbooks(mk).Close
                    Application.DisplayAlerts = True
                    Workbooks(template).Activate
                    Worksheets("Information").Activate
                End If
                Range(Cells(40, 3), Cells(46, 3)).Delete
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
                Worksheets(u).Activate
                ActiveSheet.Range("B3").Select
                Next
                ' fix the width of column in Revision History (2021/05/10 Penny)
                Worksheets("Revision History").Range("C:C").ColumnWidth = 45.09
                Worksheets("Revision History").Range("C:C").WrapText = True
                Worksheets(1).Activate
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs c & Cells(3, 3) & "-Rev" & Cells(4, 3) & ".xlsx"
                'ActiveWorkbook.SaveAs c & Cells(3, 3) & ".xlsx"
                ActiveWorkbook.Close (False)
                s = s + i
            Else
                 Cells(i, 8).Font.Color = RGB(255, 0, 0)
                MsgBox "template: '" & Cells(i, 8) & "' doesn't exist! Please check the filename or create a template!", vbOKCancel, "Note"
            End If
          End If
        Erase sap_part
flag1:
        Next i
        i = ""
    End If
    MsgBox "Row" & sr & "-" & er & " have been processed!"
End Sub










