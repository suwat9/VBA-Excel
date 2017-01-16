Sub Arabic2Thai()
    ClearStyle ("ArabicToThai")
    ActiveWorkbook.Styles.Add name:="ArabicToThai"
    With ActiveWorkbook.Styles("ArabicToThai")
        .IncludeNumber = True
        .IncludeFont = True
        .IncludeAlignment = True
        .IncludeBorder = True
        .IncludePatterns = True
        .IncludeProtection = True
    End With
    ActiveWorkbook.Styles("ArabicToThai").NumberFormat = "[$-D07041E]t#,##0.00"
    'Code แสดงตัวอย่าง การเลือกตำแหน่งเซลที่ต้องการเปลี่ยนจากเลขอาระบิคเป็นเลขไทย ในตัวอย่างเลือก A1
    Range("A1").Select
    Range("A1").Value = 999.99
    'กำหนดให้เซลสไตล์เป็น ArabicToThai ซึ่งเป็นการแสดงเลขไทย
    Selection.Style = "ArabicToThai"
End Sub
