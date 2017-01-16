Sub createSheet(name As String)
Dim dup As Boolean
dup = False

For Each sheetData In ActiveWorkbook.Worksheets
    If sheetData.name = name Then
        sheetData.Activate
        dup = True
        MsgBox "[" & name & "]" & "ª×èÍªÕ·¹Õé«éÓ ÊÃéÒ§äÁèä´é"
        Exit For
    End If
Next
If Not dup Then
        ThisWorkbook.Worksheets.Add
        ActiveSheet.name = name
End If
End Sub
