Attribute VB_Name = "Module1"
Sub delold()

    'This macro is to remove specific words in different documents.

    Dim f  
    Dim i      'counter
    Dim mypath 'my path
    Dim fname  'filename
    Dim j      'Start row


    mypath = Application.ThisWorkbook.Path & "\" 'the path of this workbook
    Application.ScreenUpdating = False
    i = 1 'counter

    'Clean the range
    Columns("A:B").Select
    Selection.ClearContents
    Range("A1") = "Files"
    Range("B1") = "Clear"
    
    'Return the file name
    f = Dir(mypath)
    While f <> ""
        Cells(i + 1, "A") = f
        i = i + 1
        f = Dir
    Wend

For j = 2 To i
    Cells(j, 1).Select
    'Judge if the activecell is 600 or not
    If InStr(ActiveCell, "-600-") > 0 Then
        fname = ActiveCell.Value
        'open the doc
        Workbooks.Open Filename:=mypath & fname
        'Activate Marking sheet and select A1
        Sheets("Marking").Select
        Range("A1").Select
        'Find the Customer Part number in this sheet
        Cells.Find(What:="Customer Part Number", After:=ActiveCell, LookIn:= _
            xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, MatchByte:=False, SearchFormat:=False).Activate
        'Select the range which needs to be deleted
        ActiveCell.Offset(1, 0).Range("A1").Select
        Range(Selection, Selection.End(xlDown)).Select
        'Delete the range
        Selection.ClearContents
        'Back to the first sheet(page)
        Sheets("Information").Select
        'Save and close doc.
        ActiveWorkbook.Save
        ActiveWindow.Close

    'Show the status
    Cells(j, 2) = "OK"


    End If
Next j
'Message when completion
MsgBox ("Done")
    
End Sub



