Sub Multi_FindReplace()
'PURPOSE: Find & Replace a list of text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim sht As Worksheet
Dim fndList As Variant
Dim rplcList As Variant
Dim x As Long

fndList = Array("Canada", "United States", "Mexico")
rplcList = Array("CAN", "USA", "MEX")

'Loop through each item in Array lists
  For x = LBound(fndList) To UBound(fndList)
    'Loop through each worksheet in ActiveWorkbook
      For Each sht In ActiveWorkbook.Worksheets
        sht.Cells.Replace What:=fndList(x), Replacement:=rplcList(x), _
          LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
          SearchFormat:=False, ReplaceFormat:=False
      Next sht
  
  Next x

End Sub
