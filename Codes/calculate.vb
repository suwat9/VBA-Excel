Sub calculate()
Dim data As Range
Set data = ActiveSheet.Range("A1:A5")
data.Value = 10

Range("A6").Select
Selection.Formula = "=sum(a1:a5)"
'Move Next Right Cell
Selection.Offset(0, 1).Select
Selection.Formula = "=average(a1:a5)"
'Copy & Paste formula
Range("a6:b6").Copy Range("C6")
End Sub
