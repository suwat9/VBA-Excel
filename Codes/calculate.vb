Sub calculate()
Dim data As Range
Set data = ActiveSheet.Range("A1:A5")
data.Value = 10

Range("A6").Select
Selection.Formula = "=sum(a1:a5)"
'ค่า Offset (R, C) โดยที่ R ใช้บอกแถวที่จะขยับไป และ C ใช้บอกสดมภ์ที่จะขยับไป จำนวนบวกใช้ขยับค่าให้มากขึ้น จำนวนลบใช้ขยับค่าให้น้อยลง
'C มีค่าเป็น +1 เป็นการขยับสดมภ์ให้มากขึ้น คือขยับสดมภ์ไปทางขวา
Selection.Offset(0, 1).Select
Selection.Formula = "=average(a1:a5)"
'Copy & Paste formula
Range("a6:b6").Copy Range("C6")
End Sub
