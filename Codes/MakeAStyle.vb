Sub MakeAStyle()
    ClearStyle ("Demo1")
    ActiveWorkbook.Styles.Add name:="Demo1"
    With ActiveWorkbook.Styles("Demo1")
        .IncludeNumber = True
        .IncludeFont = True
        .IncludeAlignment = True
        .IncludeBorder = True
        .IncludePatterns = True
        .IncludeProtection = True
    End With
    With ActiveWorkbook.Styles("Demo1").Font
        .name = "Arial Narrow"
        .Size = 11
        .Bold = False
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .Strikethrough = False
        .Color = -16776961
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveWorkbook.Styles("Demo1")
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
    End With
End Sub
