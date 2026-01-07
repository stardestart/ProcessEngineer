Attribute VB_Name = "lnkCreate2"
Sub A04_lnkCreate2()
    Dim folderPath As String
    Dim fd As FileDialog
    
    Worksheets("lnkCreate").TextBox1.Text = ""
    
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
    Dim searchRange As range
    Dim lookupRange As range
    Dim cell As range
    Dim foundCell As range
    Dim shell As Object
    Dim shortcut As Object
    Dim counter As Integer
    Dim formattedNumber As String
    Dim file As Object
    
    counter = 1
    Set searchRange = ThisWorkbook.Sheets("lnkCreate").range("D3:AK3")
    Set lookupRange = ThisWorkbook.Sheets("lnkCreate").range("C4:C" & ThisWorkbook.Sheets("lnkCreate").Cells(Rows.Count, "C").End(xlUp).row)
    For Each cell In searchRange
    If Not IsEmpty(cell.value) Then
        Set foundCell = lookupRange.Find(What:=cell.value, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                formattedNumber = Format(counter, "00")
                Set file = CreateObject("Scripting.FileSystemObject").GetFile(range("A" & cell + 3).Hyperlinks(1).Address)
                Set shell = CreateObject("WScript.Shell")
                Set shortcut = shell.CreateShortcut(folderPath & formattedNumber & " " & file.Name & ".lnk")
                shortcut.targetPath = file.Path
                shortcut.Save
                Set shortcut = Nothing
                Set shell = Nothing
                Set file = Nothing
                counter = counter + 1
            End If
        End If
    Next cell
    ThisWorkbook.Sheets("lnkCreate").range("D3:AK3").ClearContents
    Worksheets("lnkCreate").TextBox1.Activate
End Sub
