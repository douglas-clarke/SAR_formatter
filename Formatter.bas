Attribute VB_Name = "Formatter"
Public Source As String
Public SourceObj As Workbook
Public GenObj As Workbook
Public Cancel As Boolean
Public FolderFile As Object
Public SourceFolderPath As String
Public FS As Object

Option Explicit

' This section turns on high overhead operations
Sub AppTrue()
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

' This section increases performance by turning off high overhead operations
Sub AppFalse()
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

End Sub

Sub SARFormatter()
    Call AppFalse ' Turn off overhead processes
    Set GenObj = Workbooks(ActiveWorkbook.Name) ' set interface workbook as object
    Cancel = False ' Set default value for variable for error handling
    
here: Source = Cells(3, 3)
    On Error GoTo Error ' error handling
    Workbooks.Open Filename:=Source, ReadOnly:=False 'open file
    On Error GoTo 0
    
    Call MainLoop
    
    GoTo Skip ' skip error handling
Error:        ' error handling for errant file/folder name
    Call ErrorSub
    If Cancel = False Then
        GoTo Skip
    End If
    GoTo here
Skip:

EndLine:

Call AppTrue ' Turn on overhead processes
End Sub

Sub MSARFormatter() ' called if running all files in a folder
        Call AppFalse ' Turn off overhead processes
        Dim FolderFile As Object
        Dim SourceFolderPath As String
        Dim FS As Object
        Set GenObj = Workbooks(ActiveWorkbook.Name)  ' set interface workbook as object
        Cancel = False  ' Set default value for variable for error handling
        Set FS = CreateObject("Scripting.FileSystemObject")
       
here:   SourceFolderPath = Cells(3, 3)
        'On Error GoTo Error ' error handling
        Dim b As Integer
        b = 1
    
        For Each FolderFile In FS.GetFolder(SourceFolderPath).Files ' work through each file in folder
            Workbooks.Open Filename:=FolderFile, ReadOnly:=False    ' open next file
            On Error GoTo 0
            Call MainLoop
            If b = 1 Then
            MkDir ActiveWorkbook.path & " autocompleted"
            End If
            ActiveWorkbook.SaveAs ActiveWorkbook.path & " autocompleted" & "\" & ActiveWorkbook.Name
            ActiveWorkbook.Close 'close workbook
            b = b + 1
        Next FolderFile ' move to next file
        GoTo Skip ' skip error handling
Error:        ' error handling for errant file/folder name
        Call ErrorSub
        If Cancel = False Then
            GoTo Skip
        End If
        GoTo here
Skip:

EndLine:

Call AppTrue ' Turn on overhead processes
End Sub

Sub MainLoop()
    Set SourceObj = Workbooks(ActiveWorkbook.Name)
    
    Dim Sheet As Object
    Dim Col As Integer
    Dim Row As Integer
    
    Dim SheetCount As Integer
    SheetCount = ActiveWorkbook.Sheets.Count ' find number of sheets in workbook
    
    Dim x As Integer ' define counter
        
    For x = 1 To SheetCount ' cycle through sheets
        Set Sheet = SourceObj.Sheets(x)
        Sheet.Activate
        
        Col = FindLastCol ' find number of columns
        Row = FindLastRow ' find number of rows
        
        Range(Cells(Row, 1), Cells(1, Col)).Select ' select all
        Selection.MergeCells = False
        Selection.WrapText = False
        Selection.Columns.AutoFit
        Selection.Rows.AutoFit
        
        Dim i As Integer
        Dim j As Integer
        Dim a As Integer
        
        For i = 1 To Col ' cycle through columns and delete if empty
            Range(Cells(1, i), Cells(Row, i)).Select
            If WorksheetFunction.CountBlank(Selection) = Row Then
                Columns(i).Select
                Selection.Delete
                i = i - 1
                Col = Col - 1
            End If
            If i = Col Then
            Exit For
            End If
        Next
    
        For j = 1 To Row ' cycle through rows and delete if empty
            Range(Cells(j, 1), Cells(j, Col)).Select
            If WorksheetFunction.CountBlank(Selection) = Col Then
                Rows(j).Select
                Selection.Delete
                i = i - 1
                Row = Row - 1
            End If
            If i = Row Then
            Exit For
            End If
        Next
        i = 1
        j = 1
        For j = 1 To Row
            For i = 1 To Col
                Cells(j, i) = Trim(Cells(j, i))
                If Right(Cells(j, i), 2) = "--" Then
                    Cells(j, i) = Left(Cells(j, i), Len(Cells(j, i)) - 2)
                End If
                If Trim(Cells(j, i)) = "--" Or Trim(Cells(j, i)) = "-- --" Then
                    Cells(j, i) = ""
                End If
                On Error Resume Next
                Cells(j, i).NumberFormat = "0.0"
                On Error GoTo 0
            Next
        Next
    Next
    
    j = 0
    For j = 1 To 8
        Sheets(1).Cells(1, 1).EntireColumn.Insert
    Next
        
    GenObj.Sheets(3).UsedRange.Copy
    SourceObj.Sheets(1).Activate
    Range(Cells(1, 1), Cells(1, 1)).PasteSpecial
    
    If SheetCount = 2 Then
        j = 0
        For j = 1 To 4
            Sheets(2).Cells(1, 1).EntireColumn.Insert
        Next
        GenObj.Sheets(4).UsedRange.Copy
        SourceObj.Sheets(2).Activate
        Range(Cells(1, 1), Cells(1, 1)).PasteSpecial xlPasteColumnWidths
        Range(Cells(1, 1), Cells(1, 1)).PasteSpecial xlPasteAll
    Else
        j = 0
        For j = 1 To 4
            Sheets(3).Cells(1, 1).EntireColumn.Insert
        Next
        GenObj.Sheets(4).UsedRange.Copy
        SourceObj.Sheets(3).Activate
        Range(Cells(1, 1), Cells(1, 1)).PasteSpecial xlPasteColumnWidths
        Range(Cells(1, 1), Cells(1, 1)).PasteSpecial xlPasteAll
    End If
        
End Sub

' Find last row on sheet
Function FindLastRow() As Integer
    If WorksheetFunction.CountA(Cells) > 0 Then
        FindLastRow = Cells.Find(What:="*", After:=[A1], _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
         
End Function

' Find last Column on Source
Function FindLastCol()
    If WorksheetFunction.CountA(Cells) > 0 Then
        'Search for any entry, by searching backwards by Columns.
        FindLastCol = Cells.Find(What:="*", After:=[A1], _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
End Function

Sub ErrorSub()
    MsgBox "File or folder not found, Please Browse for source"
    Application.Run "Sheet1.Browse_Click" ' browse for new file/folder
End Sub
