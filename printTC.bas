Attribute VB_Name = "printTC"
Sub A06_printTC()
    Dim folderPath As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    If fd.Show = -1 Then
        folderPath = fd.SelectedItems(1) & "\"
        Call ListFilesInSubFolder(folderPath)
    Else
        MsgBox "Папка не выбрана.", vbExclamation
    End If

    Set fd = Nothing
End Sub

Sub ListFilesInSubFolder(ByVal folderPath As String)
    Dim fso As Object
    Dim wb As Workbook
    Dim file As Object
    Dim folder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    Dialog = MsgBox("ДА - Печать ПЕРВОЙ страницы всех книг в папке, НЕТ - Печать всех книг целиком в папке", vbYesNo)
    
    For Each file In folder.Files
        Application.ScreenUpdating = False
        Set wb = Workbooks.Open(file.Path)
        If Dialog = vbYes Then
            wb.Worksheets(1).PrintOut
        Else
            wb.PrintOut
        End If
        wb.Close SaveChanges:=False

    Next file

    Set fso = Nothing
    Set file = Nothing
    Set folder = Nothing
    Set wb = Nothing
End Sub
