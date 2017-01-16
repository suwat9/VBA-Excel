Sub ClearStyle(name As String)
Dim mpStyle As Style
For Each mpStyle In ActiveWorkbook.Styles
If mpStyle.Value = name Then
mpStyle.Delete
End If
Next mpStyle
End Sub
