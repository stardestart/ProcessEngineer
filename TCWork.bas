Attribute VB_Name = "TCWork"
Sub A07_TCWork()
    Dim folderPath As String
    Dim fd As FileDialog
    Dim folderPath2 As String
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    If fd.Show = -1 Then
        folderPath = fd.SelectedItems(1) & "\"
    Else
        MsgBox "Папка не выбрана.", vbExclamation
        Exit Sub
    End If
    If fd.Show = -1 Then
        folderPath2 = fd.SelectedItems(1) & "\"
    Else
        MsgBox "Папка не выбрана.", vbExclamation
        Exit Sub
    End If
    Call ListFilesInSubFolder(folderPath, folderPath2)
    Set fd = Nothing
    Set fd2 = Nothing
End Sub

Sub ListFilesInSubFolder(ByVal folderPath As String, ByVal folderPath2 As String)
    Dim pic As String
    Dim NumProc As String
    Dim model As String
    Dim oper As String
    Dim pict As Picture
    Dim file As Object
    Dim fso As Object
    Dim folder As Object
    Dim wb As Workbook
    Dim NewWb As Workbook
    Dim transl As Workbook
    Dim range1 As range
    Dim range12 As String
    Dim range2 As range
    Dim range22 As String
    Dim cell1 As range
    Dim cell2 As range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim xPic As Shape
    Dim xRg As range
    Dim xPicRg As range
    Dim counter As Integer
    Dim NumOper As Integer
    Dim row As Long
    Dim formattedNumber As String
    Dim first As String

    pic = ThisWorkbook.Sheets("TCWork").range("A58").value
    NumProc = ThisWorkbook.Sheets("TCWork").range("P104").value
    model = ThisWorkbook.Sheets("TCWork").range("G101").value
    oper = ThisWorkbook.Sheets("TCWork").range("J101").value
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    counter = 2
    row = 4
    Set transl = Workbooks.Add(xlWBATWorksheet)
    transl.Sheets(1).range("A:A,B:B,C:C,D:D,E:E").ColumnWidth = 100
    transl.Sheets(1).range("A3:E3").Borders.LineStyle = xlContinuous
    transl.Sheets(1).range("A3:E3").Borders.Weight = xlThick
    transl.Sheets(1).range("A3:E3").Interior.Color = RGB(255, 255, 158)
    transl.Sheets(1).range("A3:E3").Font.Bold = True
    transl.Sheets(1).range("A3:E3").Font.Name = "Modern H Medium"
    transl.Sheets(1).range("A3").value = "FILE"
    transl.Sheets(1).range("B3").value = "OPERATION"
    transl.Sheets(1).range("C3").value = "translation OPERATION"
    transl.Sheets(1).range("D3").value = "OPERATION NAME"
    transl.Sheets(1).range("E3").value = "translation OPERATION NAME"
    transl.Sheets(1).Name = "transl"
    Application.ScreenUpdating = False
    
    For Each file In folder.Files
    
        Set wb = Workbooks.Open(file)
        
        NumOper = ThisWorkbook.Sheets("TCWork").range("K72").value - ThisWorkbook.Sheets("TCWork").range("K61").value + 1
        
        lastRow = wb.Sheets(1).Cells(wb.Sheets(1).Rows.Count, 1).End(xlUp).row
        lastCol = wb.Sheets(1).range(ThisWorkbook.Sheets("TCWork").range("S56").value & ":" & ThisWorkbook.Sheets("TCWork").range("S56").value).Column
        Set range1 = wb.Sheets(1).range(wb.Sheets(1).Cells(1, 1), wb.Sheets(1).Cells(1, lastCol))
        
        For Each cell1 In range1
            range12 = range12 + CStr(cell1.value)
        Next cell1
        
        For i = 2 To lastRow
            Set range2 = wb.Sheets(1).range(wb.Sheets(1).Cells(i, 1), wb.Sheets(1).Cells(i, lastCol))
            
            For Each cell2 In range2
                range22 = range22 + CStr(cell2.value)
            Next cell2
            
            If range12 = range22 Then
                formattedNumber = Format(counter, "00")
                Set NewWb = Workbooks.Add(xlWBATWorksheet)
                wb.Sheets(1).Copy After:=NewWb.Sheets(NewWb.Sheets.Count)
                Application.DisplayAlerts = False
                NewWb.Sheets(1).Delete
                Application.DisplayAlerts = True
                NewWb.Sheets(1).Unprotect
                NewWb.Sheets(1).Select
                
                For Each xPic In NewWb.Sheets(1).Shapes
                    Set xRg = range(NewWb.Sheets(1).Cells(1, 1), NewWb.Sheets(1).Cells(i - 1, lastCol))
                    Set xPicRg = range(xPic.TopLeftCell.Address & ":" & xPic.BottomRightCell.Address)
                    If Not Intersect(xRg, xPicRg) Is Nothing Then
                        xPic.Delete
                    End If
                Next xPic
                
                NewWb.Sheets(1).range(NewWb.Sheets(1).Cells(1, 1), NewWb.Sheets(1).Cells(i - 1, lastCol)).Select
                Selection.Locked = False
                Selection.Delete Shift:=xlUp
                
                For Each xPic In NewWb.Sheets(1).Shapes
                    Set xRg = range(NewWb.Sheets(1).Cells(i, 1), NewWb.Sheets(1).Cells(lastRow, lastCol))
                    Set xPicRg = range(xPic.TopLeftCell.Address & ":" & xPic.BottomRightCell.Address)
                    If Not Intersect(xRg, xPicRg) Is Nothing Then
                        xPic.Delete
                    End If
                Next xPic
                
                NewWb.Sheets(1).range(NewWb.Sheets(1).Cells(i, 1), NewWb.Sheets(1).Cells(lastRow, lastCol)).Select
                Selection.Locked = False
                Selection.Delete Shift:=xlUp
                
                NewWb.Sheets.Add Before:=Sheets(1)
                Sheets(1).Name = "a"
        
                NewWb.Sheets(1).range("A:A,G:G,H:H").ColumnWidth = 3
                NewWb.Sheets(1).range("B:B").ColumnWidth = 20
                NewWb.Sheets(1).range("D:D,E:E,F:F").ColumnWidth = 10
                NewWb.Sheets(1).range("C:C").ColumnWidth = 30
                NewWb.Sheets(1).range("I:I").ColumnWidth = 17
                NewWb.Sheets(1).range("J:J").ColumnWidth = 23
                NewWb.Sheets(1).range("K:K").ColumnWidth = 7
                NewWb.Sheets(1).range("N:N").ColumnWidth = 6
                NewWb.Sheets(1).range("M:M,O:O,P:P,Q:Q,R:R").ColumnWidth = 8
                ThisWorkbook.Sheets("TCWork").range("A1:R52").Copy
                NewWb.Sheets(1).Paste
        
                Application.PrintCommunication = True
                NewWb.Sheets(1).PageSetup.PrintArea = "$A$1:$R$52"
                ActiveWindow.View = xlPageBreakPreview
                With ActiveSheet.PageSetup
                    .LeftHeader = ""
                    .CenterHeader = ""
                    .RightHeader = ""
                    .LeftFooter = ""
                    .CenterFooter = ""
                    .RightFooter = ""
                    .LeftMargin = Application.InchesToPoints(0)
                    .RightMargin = Application.InchesToPoints(0)
                    .TopMargin = Application.InchesToPoints(0)
                    .BottomMargin = Application.InchesToPoints(0)
                    .HeaderMargin = Application.InchesToPoints(0)
                    .FooterMargin = Application.InchesToPoints(0)
                    .PrintHeadings = False
                    .PrintGridlines = False
                    .PrintComments = xlPrintNoComments
                    .PrintQuality = 600
                    .CenterHorizontally = True
                    .CenterVertically = True
                    .Orientation = xlLandscape
                    .Draft = False
                    .PaperSize = xlPaperA4
                    .FirstPageNumber = xlAutomatic
                    .Order = xlDownThenOver
                    .BlackAndWhite = False
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = 1
                    .PrintErrors = xlPrintErrorsDisplayed
                    .OddAndEvenPagesHeaderFooter = False
                    .DifferentFirstPageHeaderFooter = False
                    .ScaleWithDocHeaderFooter = True
                    .AlignMarginsHeaderFooter = True
                    .EvenPage.LeftHeader.Text = ""
                    .EvenPage.CenterHeader.Text = ""
                    .EvenPage.RightHeader.Text = ""
                    .EvenPage.LeftFooter.Text = ""
                    .EvenPage.CenterFooter.Text = ""
                    .EvenPage.RightFooter.Text = ""
                    .FirstPage.LeftHeader.Text = ""
                    .FirstPage.CenterHeader.Text = ""
                    .FirstPage.RightHeader.Text = ""
                    .FirstPage.LeftFooter.Text = ""
                    .FirstPage.CenterFooter.Text = ""
                    .FirstPage.RightFooter.Text = ""
                    .PrintTitleRows = ""
                    .PrintTitleColumns = ""
                End With
        
        
                NewWb.Sheets(2).range(pic).Copy
                NewWb.Sheets(1).range("A3:F35").Select
                NewWb.Sheets(1).Pictures.Paste
                Set pict = NewWb.Sheets(1).Pictures(NewWb.Sheets(1).Pictures.Count)
        
                If range("A3:F35").Width / range("A3:F35").Height > (pict.Width / pict.Height) Then
                    pict.Width = range("A3:F35").Height * (pict.Width / pict.Height)
                    pict.Height = range("A3:F35").Height
                    pict.Top = range("A3:F35").Top + (range("A3:F35").Height - range("A3:F35").Height) / 2
                    pict.Left = range("A3:F35").Left + (range("A3:F35").Width - range("A3:F35").Height * (pict.Width / pict.Height)) / 2
                Else
                    pict.Width = range("A3:F35").Width
                    pict.Height = range("A3:F35").Width / (pict.Width / pict.Height)
                    pict.Top = range("A3:F35").Top + (range("A3:F35").Height - range("A3:F35").Width / (pict.Width / pict.Height)) / 2
                    pict.Left = range("A3:F35").Left + (range("A3:F35").Width - range("A3:F35").Width) / 2
                End If
        
                NewWb.Sheets(1).range("B46").value = NewWb.Sheets(2).range(NumProc).value
                NewWb.Sheets(1).range("N49").value = ThisWorkbook.Sheets("TCWork").range("N104").value + "_" + NewWb.Sheets(2).range(NumProc).value + "-" + formattedNumber
                
                NewWb.Sheets(1).range("N49").Replace What:="/", Replacement:="%", LookAt:=xlPart, SearchOrder _
                    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                NewWb.Sheets(1).range("N49").Replace What:="\", Replacement:="%", LookAt:=xlPart, SearchOrder _
                    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
                NewWb.Sheets(1).range("D46, G50, J50, K50").value = Date
                NewWb.Sheets(1).range("G46").value = NewWb.Sheets(2).range(model).value
                NewWb.Sheets(1).range("J46").value = NewWb.Sheets(2).range(oper).value
        
                NewWb.SaveAs folderPath2 & NewWb.Sheets(1).range("N49").value & ".slx", FileFormat:=51
                transl.Sheets(1).range("B" & row & ":" & "B" & row + NumOper - 1).value = NewWb.Sheets(2).range(ThisWorkbook.Sheets("TCWork").range("J61").value & ThisWorkbook.Sheets("TCWork").range("K61").value & ":" & ThisWorkbook.Sheets("TCWork").range("J61").value & ThisWorkbook.Sheets("TCWork").range("K72").value).value
                transl.Sheets(1).range("A" & row).value = folderPath2 & NewWb.Sheets(1).range("N49").value & ".slx"
                transl.Sheets(1).range("D" & row).value = NewWb.Sheets(1).range("J46").value
                row = row + NumOper
                NewWb.Close SaveChanges:=False
                
                wb.Sheets(1).Unprotect
                wb.Sheets(1).Select
                
                For Each xPic In wb.Sheets(1).Shapes
                    Set xRg = range(wb.Sheets(1).Cells(i, 1), wb.Sheets(1).Cells(i + i - 2, lastCol))
                    Set xPicRg = range(xPic.TopLeftCell.Address & ":" & xPic.BottomRightCell.Address)
                    If Not Intersect(xRg, xPicRg) Is Nothing Then
                        xPic.Delete
                    End If
                Next
                
                wb.Sheets(1).range(wb.Sheets(1).Cells(i, 1), wb.Sheets(1).Cells(i + i - 2, lastCol)).Select
                Selection.Locked = False
                Selection.Delete Shift:=xlUp
                
                counter = counter + 1
                i = i - 1
                first = "-01"
            End If
            range22 = ""
        Next i
        
        wb.Sheets(1).Unprotect
        wb.Sheets(1).Select
        wb.Sheets.Add Before:=Sheets(1)
        Sheets(1).Name = "a"
        
        wb.Sheets(1).range("A:A,G:G,H:H").ColumnWidth = 3
        wb.Sheets(1).range("B:B").ColumnWidth = 20
        wb.Sheets(1).range("D:D,E:E,F:F").ColumnWidth = 10
        wb.Sheets(1).range("C:C").ColumnWidth = 30
        wb.Sheets(1).range("I:I").ColumnWidth = 17
        wb.Sheets(1).range("J:J").ColumnWidth = 23
        wb.Sheets(1).range("K:K").ColumnWidth = 7
        wb.Sheets(1).range("N:N").ColumnWidth = 6
        wb.Sheets(1).range("M:M,O:O,P:P,Q:Q,R:R").ColumnWidth = 8
        ThisWorkbook.Sheets("TCWork").range("A1:R52").Copy
        wb.Sheets(1).Paste
        
        Application.PrintCommunication = True
        wb.Sheets(1).PageSetup.PrintArea = "$A$1:$R$52"
        ActiveWindow.View = xlPageBreakPreview
        With ActiveSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0)
            .RightMargin = Application.InchesToPoints(0)
            .TopMargin = Application.InchesToPoints(0)
            .BottomMargin = Application.InchesToPoints(0)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = True
            .CenterVertically = True
            .Orientation = xlLandscape
            .Draft = False
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
        
        
        wb.Sheets(2).range(pic).Copy
        wb.Sheets(1).range("A3:F25").Select
        wb.Sheets(1).Pictures.Paste
        Set pict = wb.Sheets(1).Pictures(wb.Sheets(1).Pictures.Count)
        
        If range("A3:F35").Width / range("A3:F35").Height > (pict.Width / pict.Height) Then
            pict.Width = range("A3:F35").Height * (pict.Width / pict.Height)
            pict.Height = range("A3:F35").Height
            pict.Top = range("A3:F35").Top + (range("A3:F35").Height - range("A3:F35").Height) / 2
            pict.Left = range("A3:F35").Left + (range("A3:F35").Width - range("A3:F35").Height * (pict.Width / pict.Height)) / 2
        Else
            pict.Width = range("A3:F35").Width
            pict.Height = range("A3:F35").Width / (pict.Width / pict.Height)
            pict.Top = range("A3:F35").Top + (range("A3:F35").Height - range("A3:F35").Width / (pict.Width / pict.Height)) / 2
            pict.Left = range("A3:F35").Left + (range("A3:F35").Width - range("A3:F35").Width) / 2
        End If
        
        wb.Sheets(1).range("B46").value = wb.Sheets(2).range(NumProc).value
        wb.Sheets(1).range("N49").value = ThisWorkbook.Sheets("TCWork").range("N104").value + "_" + wb.Sheets(2).range(NumProc).value + first
        
        wb.Sheets(1).range("N49").Replace What:="/", Replacement:="%", LookAt:=xlPart, SearchOrder _
                    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        wb.Sheets(1).range("N49").Replace What:="\", Replacement:="%", LookAt:=xlPart, SearchOrder _
                    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                    
        wb.Sheets(1).range("D46, G50, J50, K50").value = Date
        wb.Sheets(1).range("G46").value = wb.Sheets(2).range(model).value
        wb.Sheets(1).range("J46").value = wb.Sheets(2).range(oper).value
        
        Application.CutCopyMode = False
        wb.SaveAs folderPath2 & wb.Sheets(1).range("N49").value & ".slx", FileFormat:=51
        transl.Sheets(1).range("B" & row & ":" & "B" & row + NumOper - 1).value = wb.Sheets(2).range(ThisWorkbook.Sheets("TCWork").range("J61").value & ThisWorkbook.Sheets("TCWork").range("K61").value & ":" & ThisWorkbook.Sheets("TCWork").range("J61").value & ThisWorkbook.Sheets("TCWork").range("K72").value).value
        transl.Sheets(1).range("A" & row).value = folderPath2 & wb.Sheets(1).range("N49").value & ".slx"
        transl.Sheets(1).range("D" & row).value = wb.Sheets(1).range("J46").value
        row = row + NumOper
        wb.Close SaveChanges:=False
        counter = 2
        range12 = ""
        range22 = ""
        first = ""
    Next file
    
    Dim vbComp As Object
    Dim NewCode As String
    Dim btn As Button

    NewCode = "Sub transl()" & vbCrLf _
    & "Dim fso As Object" & vbCrLf _
    & "Dim wb As Workbook" & vbCrLf _
    & "Dim NumOper As Integer" & vbCrLf _
    & "Dim row As Long" & vbCrLf _
    & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf _
    & "row = 4" & vbCrLf _
    & "NumOper = " & NumOper & vbCrLf _
    & "Do While ThisWorkbook.Sheets(1).range(""A"" & row).value <> 0" & vbCrLf _
    & "Application.ScreenUpdating = False" & vbCrLf _
    & "Set wb = Workbooks.Open(ThisWorkbook.Sheets(1).range(""A"" & row).value)" & vbCrLf _
    & "wb.Worksheets(1).range(""J6:J"" & NumOper + 5).value = ThisWorkbook.Sheets(1).range(""C"" & row & "":"" & ""C"" & row + NumOper - 1).value" & vbCrLf _
    & "wb.Worksheets(1).range(""J47"").value = ThisWorkbook.Sheets(1).range(""E"" & row).value" & vbCrLf _
    & "row = row + NumOper" & vbCrLf _
    & "wb.Close SaveChanges:=True" & vbCrLf _
    & "Loop" & vbCrLf _
    & "Set wb = Nothing" & vbCrLf _
    & "Set fso = Nothing" & vbCrLf _
    & "End Sub"

    With transl.VBProject.VBComponents
        Set vbComp = .Add(1)
    End With

    With vbComp.CodeModule
        .AddFromString NewCode
    End With
    
    Set btn = transl.Sheets(1).Buttons.Add(6, 4.5, 70, 25)
    btn.Characters.Text = "Run"
    btn.OnAction = "'" & transl.FullName & "'!transl"
    
    MsgBox "Выполните перевод и нажмите кнопку Run / Make a translation and press the button Run", vbExclamation
    transl.Activate
    
    Set pict = Nothing
    Set file = Nothing
    Set fso = Nothing
    Set folder = Nothing
    Set wb = Nothing
    Set NewWb = Nothing
    Set range1 = Nothing
    Set range2 = Nothing
    Set cell1 = Nothing
    Set cell2 = Nothing
    Set xPic = Nothing
    Set xRg = Nothing
    Set xPicRg = Nothing
    Set vbComp = Nothing
    Set btn = Nothing
End Sub
