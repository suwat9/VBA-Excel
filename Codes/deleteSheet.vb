Sub deleteSheet(name As String)
For Each sheetData In ActiveWorkbook.Worksheets
    If sheetData.name = name Then
        Application.DisplayAlerts = False           'ไม่เตือนก่อนลบชีท
        sheetData.Delete
        Application.DisplayAlerts = True            'เตือนก่อนลบชีท
        Exit For
    End If
Next
End Sub
