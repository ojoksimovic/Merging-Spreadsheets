Attribute VB_Name = "Module1"
Sub MergeUnitSubmissions()
Dim Path As String
Path = "C:\Users\"
Filename = Dir(Path & "*.xlsx")

Application.DisplayAlerts = False

Do While Filename <> ""

Workbooks.Open Filename:=Path & Filename, ReadOnly:=True

LastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

Range("A7", Cells(27, 35)).Copy Workbooks("SGS University Wide Awards Master").Sheets(1).Range("A65536").End(xlUp).Offset(1, 0)
Workbooks(Filename).Close
Filename = Dir()
Loop

Application.DisplayAlerts = True

End Sub

Sub DeleteSampleAndBlanks()

Range("A1", "A9999").Select
For Each Cell In Selection
   If Cell.Value = "SAMPLE ONLY" Then
      Rows(Cell.Row).ClearContents
   End If
Next
End Sub

Sub SeleteBlanks()

For Each Cell In Range("A6", "A9999")
    If Cell.Value = "" Then
    Cell.Activate
    ActiveCell.EntireRow.Delete
    End If
Next
End Sub

