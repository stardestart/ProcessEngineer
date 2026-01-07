VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lnkCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub TextBox1_Change()
    Dim ws As Worksheet
    Dim searchTerm As String
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("lnkCreate")
    searchTerm = Me.TextBox1.Text
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    
    If lastRow >= 3 Then
        If searchTerm <> "" Then
            ws.range("B3:B" & lastRow).AutoFilter Field:=1, Criteria1:="*" & searchTerm & "*"
        End If
    Else
        MsgBox "Нет данных для фильтрации.", vbExclamation
    End If
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim range As range
        Set range = ThisWorkbook.Sheets("lnkCreate").range("B4:B" & ThisWorkbook.Sheets("lnkCreate").Cells(ThisWorkbook.Sheets("lnkCreate").Rows.Count, 2).End(xlUp).row).SpecialCells(xlCellTypeVisible).Cells(1, 1)
        ThisWorkbook.Sheets("lnkCreate").Cells(range.row, 3).Select
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As range)
    Dim n As Integer
    Dim value As Integer
    If Target.Cells.Count = 1 And Target.Column = 3 And Target.row >= 4 Then
        value = Target.value
        n = 4
        Do While Not IsEmpty(Cells(3, n))
            n = n + 1
        Loop
        ThisWorkbook.Sheets("lnkCreate").Cells(3, n).value = value
        ThisWorkbook.Sheets("lnkCreate").Hyperlinks.Add Anchor:=Cells(3, n), Address:="", SubAddress:="B" & Target.row
        Target.Interior.Color = RGB(255, 0, 0)
        ThisWorkbook.Sheets("lnkCreate").Cells(Target.row, 2).Interior.Color = RGB(255, 0, 0)
        ThisWorkbook.Sheets("lnkCreate").Cells(Target.row, 1).Interior.Color = RGB(255, 0, 0)
        Worksheets("lnkCreate").TextBox1.Text = ""
        Worksheets("lnkCreate").TextBox1.Activate
    End If
End Sub
