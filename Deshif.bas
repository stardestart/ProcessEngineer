Sub A03_Deshif()
    Dim folderPath As String
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    If fd.Show = -1 Then
        folderPath = fd.SelectedItems(1) & "\"
        Call ListFilesInSubFolder(folderPath, 4)
    Else
        MsgBox "Ïàïêà íå âûáðàíà.", vbExclamation
    End If
    
    Set fd = Nothing
End Sub

Sub ListFilesInSubFolder(ByVal folderPath As String, ByRef row As Integer)
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim shell As Object
    Dim fileExtension As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.Files
        fileExtension = LCase(fso.GetExtensionName(file.Name))
        If fileExtension = "xlsm" Or fileExtension = "xls" Or fileExtension = "xlsx" Then
            Dim wb As Workbook
            Dim NewWb As Workbook
            Dim ws As Worksheet
            Application.ScreenUpdating = False
            Set wb = Workbooks.Open(file)
            Set NewWb = Workbooks.Add(xlWBATWorksheet)
        
            For Each ws In wb.Sheets
                ws.Copy After:=NewWb.Sheets(NewWb.Sheets.Count)
            Next ws
        
            If NewWb.Sheets.Count > 1 Then
                Application.DisplayAlerts = False
                NewWb.Sheets(1).Delete
                Application.DisplayAlerts = True
            End If
        
            wb.Close SaveChanges:=False
            NewWb.SaveAs Left(file, InStrRev(file, ".")) & "slx", FileFormat:=51
            NewWb.Close SaveChanges:=False
            Application.Wait Now + TimeValue("0:00:01")
            shell.Run "cmd /c del """ & file.Path & """", 0, True
            Set wb = Nothing
            Set NewWb = Nothing
            Set ws = Nothing
        
            ElseIf fileExtension = "pptx" Or fileExtension = "ppt" Then
                Dim pptApp As Object
                Dim pptPres As Object
                Dim newPres As Object
                Dim sld As Object
                Dim pres As Object
                Set pptApp = CreateObject("PowerPoint.Application")
                Set pptPres = pptApp.Presentations.Open(file.Path)
                Set newPres = pptApp.Presentations.Add
                For Each sld In pptPres.Slides
                    sld.Copy
                    newPres.Slides.Paste
                Next sld
                newPres.SaveAs Left(file, InStrRev(file, ".")) & "pptm"
                For Each pres In pptApp.Presentations
                    pres.Close
                Next pres
                pptApp.Quit
                Application.Wait Now + TimeValue("0:00:01")
                shell.Run "cmd /c del """ & file.Path & """", 0, True
                Set pptApp = Nothing
                Set pptPres = Nothing
                Set newPres = Nothing
                Set sld = Nothing
                Set pres = Nothing
        
            ElseIf fileExtension = "docx" Or fileExtension = "doc" Then
                Dim wdApp As Object
                Dim docSource As Object
                Dim docTarget As Object
                Set wdApp = CreateObject("Word.Application")
                wdApp.Visible = False
                Set docSource = wdApp.Documents.Open(file.Path)
                Set docTarget = wdApp.Documents.Add
                docSource.Content.Select
                wdApp.Selection.Copy
                docTarget.Content.Select
                wdApp.Selection.HomeKey Unit:=6
                wdApp.Selection.Paste
                docTarget.SaveAs Left(file, InStrRev(file, ".")) & "cod"
                docSource.Close False
                docTarget.Close True
                wdApp.Quit
                Application.Wait Now + TimeValue("0:00:01")
                shell.Run "cmd /c del """ & file.Path & """", 0, True
                Set wdApp = Nothing
                Set docSource = Nothing
                Set docTarget = Nothing
            Else: MsgBox "no file"
        End If
    Next file
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    Set shell = Nothing
    MsgBox "Done"
End Sub
