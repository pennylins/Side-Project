Attribute VB_Name = "Module2"
Sub move()

    'This macro is to move documents to another folder.

    Dim a, j, i As Integer
    Dim c, ox, op As String
    Dim fs As Object
    Dim d As Date
    
    
    
    a = ThisWorkbook.name
    'find the last row of the worksheet
    j = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row + 1
    c = "internal path1"
    ox = "internal path2"
    op = "internal path3"
    
    On Error Resume Next
    Err.Clear
    'Define the top row style
    Range("b:b").Clear
    Range("b1").Value = "Consequence"
    Range("b1").Select
    Selection.Interior.Color = RGB(0, 0, 0)
    Selection.Font.Color = RGB(255, 255, 255)

    'Check if the documents exist, if yes, move the document to another folder; if no, put comment in the cells.    
    Application.ScreenUpdating = False
    Set fs = CreateObject("scripting.filesystemobject")
    i = 2
    Do
        If fs.fileexists(Trim(c & Cells(i, 1))) Then
            If CInt(InStr(1, Cells(i, 1), ".xlsx")) > 0 Or CInt(InStr(1, Cells(i, 1), ".dwg")) > 0 Then
                fs.movefile c & Cells(i, 1), ox & Cells(i, 1)
                Cells(i, 1).Offset(, 1).Value = "OK"
            End If
        Else
            Cells(i, 1).Offset(, 1).Value = "The file isn't existed in 'Controlled Document' folder"
            
        End If
        Range("b1").EntireColumn.AutoFit
        i = i + 1
    Loop Until i = j
    Application.ScreenUpdating = True
    
    'pop out message box when completed
    MsgBox ("Process is done")
    'auto save
    d = Now()
    Range("c1").Value = "Last Save: "
    Range("d1").Value = d
    Range("d1").EntireColumn.AutoFit
    ThisWorkbook.Save
    
    
    
    
If Err.Number <> 0 Then
    MsgBox Err.Description
Else

End If
     

End Sub
