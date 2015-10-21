' This macro is used in the exported CSV files
' It will format the colours and layout automatically
' Copy the code into MS Excel VBA Worksheet/Module and run
Sub UniformlyFormatWorksheets()

    Dim WS_Count As Integer
    Dim i As Integer

    WS_Count = ActiveWorkbook.Worksheets.Count

    For i = 1 To WS_Count
        
        DoEvents
        ActiveWorkbook.Worksheets(i).Select
    
        ' Heading Setup
        Columns("A:A").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$A$9999").AutoFilter Field:=1, Criteria1:= _
            "GraphicName"
        Cells.Select
        Selection.Font.Bold = True
        ActiveSheet.Range("$A$1:$A$9999").AutoFilter Field:=1
    
        ' Highlight Errors
        With Application.ReplaceFormat.Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599963377788629
            .PatternTintAndShade = 0
        End With
        Cells.Replace What:="error", Replacement:="error", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=True
            
        ' Highlight NIU
        With Application.ReplaceFormat.Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        Cells.Replace What:="NIU", Replacement:="NIU", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=True
    
        ' Highlight not exist
        With Application.ReplaceFormat.Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Application.ReplaceFormat.Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Cells.Replace What:="not exist", Replacement:="not exist", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=True
            
        ' Add Borders
        Cells.Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThin
        End With
        Range("A1").Select
            
        ' Page Setup
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
        Application.PrintCommunication = True
        ActiveSheet.PageSetup.PrintArea = ""
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .LeftHeader = "&""-,Bold""&18&E&A"
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = "&D"
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.7)
            .RightMargin = Application.InchesToPoints(0.7)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .PrintQuality = 300
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlPortrait
            .Draft = False
            .PaperSize = xlPaperLetter
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 100
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
        Application.PrintCommunication = True
        
        ' Autofit Cells
        Cells.Select
        Cells.EntireColumn.AutoFit
        
        ' Print Setup
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
        Application.PrintCommunication = True
        ActiveSheet.PageSetup.PrintArea = ""
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .LeftHeader = "&""-,Bold""&18&E&A"
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = "&D"
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.7)
            .RightMargin = Application.InchesToPoints(0.7)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintSheetEnd
            .PrintQuality = 300
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlLandscape
            .Draft = False
            .PaperSize = xlPaperA3
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
        Application.PrintCommunication = True
    
    Next i
  
    MsgBox Finished

    
End Sub
