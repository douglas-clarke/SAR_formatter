Attribute VB_Name = "Template_Generation"
Option Explicit
Option Base 1

Sub GenerateSarVarTemplate()
    Call AppFalse
    Dim GenObj As Object
    Set GenObj = Workbooks(ActiveWorkbook.Name)
    If Sheets(1).RunOne = True Then
        MsgBox ("To choose a folder select the" & " ""Run Folder"" " & "option" & _
        ", then click browse")
        Exit Sub
    End If
    If Cells(3, 3) = "" Then
        MsgBox "Please select a folder"
        Exit Sub
    End If
    Dim counter
    counter = 0
    SourceFolderPath = Cells(3, 3) ' Public in module1
    Set FS = CreateObject("Scripting.FileSystemObject") ' Public in module1
    ReDim Years(1) As Integer
    ReDim Files(1) As String
    Dim x As Integer
    x = 1
    Dim FilePath As String
    
    For Each FolderFile In FS.GetFolder(SourceFolderPath).Files
        If counter = 0 Then
            FilePath = ActiveWorkbook.path
        End If
        counter = counter + 1
        ReDim Preserve Files(x)
        ReDim Preserve Years(x)
        Files(x) = FolderFile
        Dim i As Integer
        Dim j As String
        For i = 4 To 100
            If IsNumeric(Left(Trim(Right(Files(x), i)), 4)) Then
                Years(x) = Left(Trim(Right(Files(x), i)), 4)
            End If
        Next
        x = x + 1
    Next
    
    Workbooks.Add
    Dim SarVar As Object
    Set SarVar = ActiveWorkbook
    Dim SheetCount
    SheetCount = UBound(Years)
    x = 0
    
    SarVar.Sheets.Add Count:=SheetCount - 3
    
    For x = 1 To SheetCount
        GenObj.Sheets(2).UsedRange.Copy
        SarVar.Sheets(x).Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths
        SarVar.Sheets(x).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
        Sheets(x).Name = CStr(Years(x))
    Next
            
    ActiveWorkbook.SaveAs FilePath & "\" & "SAR Variances.xlsx"
    GenObj.Sheets(1).Cells(9, 3) = ActiveWorkbook.path & "\" & ActiveWorkbook.Name
    ActiveWorkbook.Close 'close workbook
    
    Call AppTrue
End Sub
