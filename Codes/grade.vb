Function myGrade(ByVal score As Integer) As String
If score >= 80 Then
myGrade = "A"
ElseIf score >= 70 Then
myGrade = "B"
ElseIf score >= 60 Then
myGrade = "C"
ElseIf score >= 50 Then
myGrade = "D"
Else
myGrade = "E"
End If
End Function
Sub grade1()
With ThisWorkbook.Worksheets("grade")
For i = 1 To 20
s = .Cells(2 + i, 3).Value
gr = myGrade(s)
.Cells(2 + i, 4) = gr
Next i
End With
End Sub
