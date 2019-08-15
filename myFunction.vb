Imports Excel = Microsoft.Office.Interop.Excel

Module myFunction

    Public app As Excel.Application
    Public wbook As Excel.Workbook
    Public sheet As Excel.Worksheet
    Public sheet2 As Excel.Worksheet

    Public Function WorkBook(app As Object, path As String) As Excel.Workbook

        'Dim app As Excel.Application

        'path = "\\erp\TEMP\App\Программа Export\Октябрь-2017.xlsx"

        Try
            app = GetObject(, "Excel.Application")

        Catch ex As Exception
            app = CreateObject("Excel.Application")
        End Try

        app.Visible = True

        Return app.Workbooks.Add(path)

    End Function

    Public Sub SheetSettings()

        sheet.Columns("A:A").ColumnWidth() = 6
        sheet.Columns("B:B").ColumnWidth() = 57
        sheet.Columns("C:C").ColumnWidth() = 15
        sheet.Columns("D:D").ColumnWidth() = 11.14
        sheet.Columns("E:E").ColumnWidth() = 7
        sheet.Columns("F:F").ColumnWidth() = 11.14
        sheet.Columns("G:G").ColumnWidth() = 11.14
        sheet.Columns("H:H").ColumnWidth() = 11.14

        sheet.Columns("I:I").ColumnWidth() = 11.14
        sheet.Columns("J:J").ColumnWidth() = 11.14
        sheet.Columns("K:K").ColumnWidth() = 11.14
        sheet.Columns("L:L").ColumnWidth() = 11.14
        sheet.Columns("M:M").ColumnWidth() = 11.14
        sheet.Columns("N:N").ColumnWidth() = 11.14

        Dim rngBB As Excel.Range = sheet.Columns("B:B")
        With rngBB
            .WrapText = True        '    .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
        End With

        

    End Sub


    Public Function tableCreateListObject(cell As Excel.Range) As Excel.ListObject

        Dim r1 As Excel.Range = app.Range(cell, cell.End(Excel.XlDirection.xlToRight))
        Dim r2 As Excel.Range = app.Range(cell, cell.End(Excel.XlDirection.xlDown))

        Dim table1 As Excel.Range = app.Range(r1, r2)

        Dim tableObject As Excel.ListObject = sheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, table1)
        tableObject.Name = "table1"

        Return tableObject

    End Function

    Public Sub sortTable(tableObject As Excel.ListObject, column As Integer)
        tableObject.Range.Sort( _
        Key1:=tableObject.ListColumns(column).Range, Order1:=Excel.XlSortOrder.xlAscending, _
        Key2:=tableObject.ListColumns(column).Range, Order2:=Excel.XlSortOrder.xlAscending, _
        Orientation:=Excel.XlSortOrientation.xlSortColumns, _
        Header:=Excel.XlYesNoGuess.xlYes)
    End Sub

    Public Sub tableHeaderColor(cell As Excel.Range)
        Dim r3 As Excel.Range = sheet.Range(cell, cell.End(Excel.XlDirection.xlToRight))

        With r3.Interior
            .Pattern = Excel.Constants.xlSolid
            .PatternColorIndex = Excel.Constants.xlAutomatic
            .ThemeColor = Excel.XlThemeColor.xlThemeColorLight1
            .TintAndShade = 0.499984740745262
            .PatternTintAndShade = 0
        End With

        With r3.Font
            .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .TintAndShade = 0
        End With

    End Sub

    Public Sub RenameRange(rngCount As Excel.Range, columnRename As Excel.Range)

        app.ScreenUpdating = False

        Dim rCount As Integer
        rCount = rngCount.Count

        Dim currentCell As Excel.Range

        For i = 1 To rCount
            currentCell = columnRename.Offset(i, 0)

            Dim rep As String = ""
            rep = Replace(currentCell.Value, ".", "")

            currentCell.Value = rep

            currentCell.NumberFormat = "0"
            currentCell.HorizontalAlignment = Excel.Constants.xlLeft

        Next

        app.ScreenUpdating = True
    End Sub

    Public Sub PageSettings()
        With sheet.PageSetup

            '.LeftMargin = 0.196850393700787
            '.RightMargin = 0.196850393700787
            '.TopMargin = 0.196850393700787
            '.BottomMargin = 0.196850393700787

            .Orientation = Excel.XlPageOrientation.xlLandscape
            .Zoom = 95

            .LeftMargin = 0.196850393700787
            .RightMargin = 0.196850393700787
            .TopMargin = 0.393700787401575
            .BottomMargin = 0.393700787401575

            .CenterHorizontally = True
            .CenterVertically = True

            .HeaderMargin = 0
            .FooterMargin = 0

            .PaperSize = Excel.XlPaperSize.xlPaperA3

        End With
    End Sub

End Module
