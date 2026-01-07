Attribute VB_Name = "OTS"
Sub A08_OTS()
    Dim folderPath As String
    Dim fd As FileDialog
    Dim NewWb As Workbook
    Dim i As Integer

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    Set NewWb = Workbooks.Add(xlWBATWorksheet)
    NewWb.Sheets(1).range("A:A").ColumnWidth = 2
    NewWb.Sheets(1).range("B:B").ColumnWidth = 14
    NewWb.Sheets(1).range("C:C").ColumnWidth = 30
    NewWb.Sheets(1).range("D:D, E:E").ColumnWidth = 62
    NewWb.Sheets(1).range("F:F, G:G, H:H, I:I").ColumnWidth = 20
    NewWb.Sheets(1).range("B6:I6").Borders.LineStyle = xlContinuous
    NewWb.Sheets(1).range("B6:I6").Borders.Weight = xlThick
    NewWb.Sheets(1).range("B6:I6").Interior.Color = RGB(255, 255, 158)
    NewWb.Sheets(1).range("B6:I6").Font.Bold = True
    NewWb.Sheets(1).range("B6:I6").Font.Name = "Modern H Medium"
    NewWb.Sheets(1).Rows(6).RowHeight = 30
    NewWb.Sheets(1).range("B:B, C:C, D:D, E:E, F:F, G:G, H:H, I:I").VerticalAlignment = xlCenter
    NewWb.Sheets(1).range("B:B, C:C, D:D, E:E, F:F, G:G, H:H, I:I").HorizontalAlignment = xlCenter
    NewWb.Sheets(1).range("B6").value = "SYSTEM"
    NewWb.Sheets(1).range("C6").value = "OPERATION №"
    NewWb.Sheets(1).range("D6").value = "OPERATION NAME"
    NewWb.Sheets(1).range("E6").value = "OPERATION NAME (RUSSIAN)"
    NewWb.Sheets(1).range("F6").value = "TYPE"
    NewWb.Sheets(1).range("G6").value = "LINE"
    NewWb.Sheets(1).range("H6").value = "STATION"
    NewWb.Sheets(1).range("I6").value = "OPERATION TIME"
    
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
    Dim subFolder As Object
    Dim Num As String
    Dim ValueNum As String
    Dim OperName As String
    Dim ValueOperName As String
    Dim OperNameRus As String
    Dim ValueOperNameRus As String
    Dim Time As String
    Dim ValueTime As Double
    Dim Spec As String
    Dim ValueSpec As String
    Dim Sys As String
    Dim ValueSys As String
    

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    For Each subFolder In folder.SubFolders
        ListFilesInSubFolder subFolder.Path, NewWb, i
        For Each file In subFolder.Files
            If Right(file.Name, 4) = ".lnk" Then
                Application.ScreenUpdating = False
                Set wb = Workbooks.Open(file.Path)
            
                Sys = ThisWorkbook.Sheets(1).range("D71").value
                ValueSys = wb.Sheets(1).range(Sys).value
            
                Num = ThisWorkbook.Sheets(1).range("F71").value
                ValueNum = wb.Sheets(1).range(Num).value
            
                OperName = ThisWorkbook.Sheets(1).range("H71").value
                ValueOperName = wb.Sheets(1).range(OperName).value
            
                OperNameRus = ThisWorkbook.Sheets(1).range("I71").value
                ValueOperNameRus = wb.Sheets(1).range(OperNameRus).value
            
                Time = ThisWorkbook.Sheets(1).range("G71").value
                ValueTime = wb.Sheets(1).range(Time).value
            
                Spec = ThisWorkbook.Sheets(1).range("E71").value
                ValueSpec = wb.Sheets(1).range(Spec).value
            
                wb.Close SaveChanges:=False
                NewWb.Sheets(1).range("B" & i).value = ValueSys
                NewWb.Sheets(1).range("B" & i & ":" & "I" & i).Borders.LineStyle = xlContinuous
                NewWb.Sheets(1).range("C" & i).value = ValueNum
                NewWb.Sheets(1).range("D" & i).value = ValueOperName
                NewWb.Sheets(1).range("E" & i).value = ValueOperNameRus
                NewWb.Sheets(1).range("F" & i).value = ValueSpec
                NewWb.Sheets(1).range("I" & i).value = ValueTime
            
                NewWb.Sheets(1).range("G" & i).value = folder.Name
                NewWb.Sheets(1).range("H" & i).value = subFolder.Name
                NewWb.Sheets(1).Rows(i).RowHeight = 30
                i = i + 1
            
            End If
        Next file
    Next subFolder
    Set fso = Nothing
    Set file = Nothing
    Set folder = Nothing
    Set wb = Nothing
    Set subFolder = Nothing
End Sub
