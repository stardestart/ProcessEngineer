Attribute VB_Name = "partName"
Sub A09_partName()
    Dim folderPath As String
    Dim fd As FileDialog
    Dim NewWb As Workbook
    Dim i As Integer

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    Set NewWb = Workbooks.Add(xlWBATWorksheet)
    NewWb.Sheets(1).range("A:A").ColumnWidth = 2
    NewWb.Sheets(1).range("B:B").ColumnWidth = 20
    NewWb.Sheets(1).range("C:C, D:D").ColumnWidth = 30
    NewWb.Sheets(1).range("E:E").ColumnWidth = 40
    NewWb.Sheets(1).range("F:F").ColumnWidth = 10
    NewWb.Sheets(1).range("G:G").ColumnWidth = 150
    NewWb.Sheets(1).range("B6:G6").Borders.LineStyle = xlContinuous
    NewWb.Sheets(1).range("B6:G6").Borders.Weight = xlThick
    NewWb.Sheets(1).range("B6:G6").Interior.Color = RGB(255, 255, 158)
    NewWb.Sheets(1).range("B6:G6").Font.Bold = True
    NewWb.Sheets(1).range("B6:G6").Font.Name = "Modern H Medium"
    NewWb.Sheets(1).Rows(6).RowHeight = 30
    NewWb.Sheets(1).range("B:B, C:C, D:D, E:E, F:F, G:G").VerticalAlignment = xlCenter
    NewWb.Sheets(1).range("B:B, C:C, D:D, E:E, F:F, G:G").HorizontalAlignment = xlCenter
    NewWb.Sheets(1).range("B6").value = "PART"
    NewWb.Sheets(1).range("C6").value = "PART NAME"
    NewWb.Sheets(1).range("D6").value = "No PROCESS"
    NewWb.Sheets(1).range("E6").value = "OPERATION NAME"
    NewWb.Sheets(1).range("F6").value = "MODEL"
    NewWb.Sheets(1).range("G6").value = "FOLDER"
    
    i = 7

    If fd.Show = -1 Then
        folderPath = fd.SelectedItems(1) & "\"
        Call ListFilesInSubFolder(folderPath, NewWb, i)
    Else
        MsgBox "Папка не выбрана.", vbExclamation
    End If

    Set fd = Nothing
End Sub

Sub ListFilesInSubFolder(ByVal folderPath As String, ByVal NewWb As Workbook, ByRef i As Integer)
    Dim fso As Object
    Dim wb As Workbook
    Dim file As Object
    Dim folder As Object
    Dim row As Integer
    Dim row2 As Integer
    Dim row3 As Integer
    Dim i2 As Integer
    Dim part As String
    Dim partName As String
    Dim NoProc As String
    Dim operation As String
    Dim model As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    i2 = i
    part = ThisWorkbook.Sheets(1).range("E93").value
    
    If part = "Q" Then
        row = 34
        row2 = 44
        row3 = row
        partName = "U"
        NoProc = "AC53"
        operation = "S53"
        model = "AD56"
    ElseIf part = "O" Then
        row = 22
        row2 = 35
        row3 = row
        partName = "R"
        NoProc = "AC49"
        operation = "P46"
        model = "M44"
    ElseIf part = "H" Then
        row = 32
        row2 = 40
        row3 = row
        partName = "L"
        NoProc = "N49"
        operation = "J46"
        model = "G46"
    Else: MsgBox "not correct value"
    End If
    For Each file In folder.Files
        Application.ScreenUpdating = False
        Set wb = Workbooks.Open(file.Path)

        Do While row < row2
            If Not IsEmpty(wb.Sheets(1).range(part & row).value) Then
                NewWb.Sheets(1).range("B" & i2).value = wb.Sheets(1).range(part & row).value
                NewWb.Sheets(1).range("C" & i2).value = wb.Sheets(1).range(partName & row).value
    
                NewWb.Sheets(1).range("D" & i2).value = wb.Sheets(1).range(NoProc).value
                NewWb.Sheets(1).range("E" & i2).value = wb.Sheets(1).range(operation).value
                NewWb.Sheets(1).range("F" & i2).value = wb.Sheets(1).range(model).value
                NewWb.Sheets(1).Hyperlinks.Add Anchor:=NewWb.Sheets(1).range("G" & i2), _
                    Address:=file.Path, _
                    TextToDisplay:=file.Path
                row = row + 1
                i2 = i2 + 1
            Else: row = row + 1
            End If
        Loop
        
        row = row3
        wb.Close SaveChanges:=False
    Next file
    
    NewWb.SaveAs CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & folder.Name & "_partNum_" & Format(Date, "dd-mm-yy") & ".slx", FileFormat:=51
    NewWb.Close SaveChanges:=False
    Set fso = Nothing
    Set file = Nothing
    Set folder = Nothing
    Set wb = Nothing
    Set NewWb = Nothing
End Sub
