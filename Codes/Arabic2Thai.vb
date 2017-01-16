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
    'àÅ×Í¡µÓáË¹è§à«Å·ÕèµéÍ§¡ÒÃà»ÅÕèÂ¹¨Ò¡àÅ¢ÍÒÃÐºÔ¤à»ç¹àÅ¢ä·Â ã¹µÑÇÍÂèÒ§àÅ×Í¡ A1
    Range("A1").Select
    Range("A1").Value = 999.99
    '¡ÓË¹´ãËéà«ÅÊäµÅìà»ç¹ ArabicToThai «Öè§à»ç¹¡ÒÃáÊ´§àÅ¢ä·Â
    Selection.Style = "ArabicToThai"
End Sub
