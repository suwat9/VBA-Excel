Sub ListStyles()
    Dim oSt As Style
    Dim oCell As Range
    Dim lCount As Long
    Dim oStylesh As Worksheet
    
    deleteSheet ("Config - Styles")
    ThisWorkbook.Worksheets.Add
    ActiveSheet.name = "Config - Styles"
    Set oStylesh = ThisWorkbook.Worksheets("Config - Styles")
    With oStylesh
        lCount = oStylesh.UsedRange.Rows.Count + 1
        For Each oSt In ThisWorkbook.Styles
            On Error Resume Next
            Set oCell = Nothing
            Set oCell = Intersect(oStylesh.UsedRange, oStylesh.Range("A:A")).Find(oSt.name, _
                oStylesh.Range("A1"), xlValues, xlWhole, , , False)
            If oCell Is Nothing Then
            lCount = lCount + 1
            .Cells(lCount, 1).Style = oSt.name
            .Cells(lCount, 1).Value = oSt.NameLocal
            .Cells(lCount, 2).Style = oSt.name
            End If
        Next
    End With
End Sub
