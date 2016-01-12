Attribute VB_Name = "Transfer_Data"
Option Explicit
Option Base 1
Sub DataTransfer()
        Call AppFalse ' Turn off overhead processes
        
        Dim FolderFile As Object
        Dim SourceFolderPath As String
        Dim FS As Object
        Set GenObj = Workbooks(ActiveWorkbook.Name) ' Set interface workbook as object
        Set FS = CreateObject("Scripting.FileSystemObject")

        'Cancel = False ' Set default value for variable for error handling
        
        SourceFolderPath = Cells(8, 3)
        
        Dim SAR As Object
        Workbooks.Open Filename:=Cells(9, 3)
        Set SAR = ActiveWorkbook
        
        ' On Error GoTo Error ' error handling
        Dim x As Integer
        x = 1
        Dim i As Integer
        Dim Year As Integer
        Dim SheetCount As Integer

        For Each FolderFile In FS.GetFolder(SourceFolderPath).Files ' work through each file in folder
            Set FolderFile = FolderFile
            Workbooks.Open Filename:=FolderFile
            SheetCount = ActiveWorkbook.Sheets.Count
            Sheets(x).Range(Cells(5, 3), Cells(5, 6)).Copy
            SAR.Sheets(x).Activate
            Cells(4, 8).PasteSpecial
            
            x = x + 1
        Next FolderFile
        GoTo Skip ' skip error handling
        
Error:  ' error handling for errant file/folder name
        Call ErrorSub
        If Cancel = False Then
        GoTo Skip
        End If
Skip:
EndLine:
        Call AppTrue
End Sub
