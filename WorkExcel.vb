Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Text

Public Class WorkExcel
    Private app As Excel.Application
    Private wbook As Excel.Workbook
    Public sheet As Excel.Worksheet
    Private path_ As String

    ''' <summary>
    ''' Ячейка с которой начинается таблица на листе
    ''' </summary>
    ''' <remarks></remarks>
    Dim cellFirst As Excel.Range

    ''' <summary>
    ''' Перенаименование и сортировка по этому столбцу
    ''' </summary>
    ''' <remarks></remarks>
    Dim columnRename As Excel.Range

    ''' <summary>
    ''' Количество строк документа без шапки
    ''' </summary>
    ''' <remarks></remarks>
    Dim rngCount As Excel.Range

    Dim tableObject As Excel.ListObject

    Public Shared ReadOnly Property PathDirectoryOSV As String
        Get
            'Return "\\erpdb\Updates\_App540\OSV"
            'Return "\\fs\TMP$\540\OSV"
            Return "\\erpdb\TEMP\OSV"
            'Return "D:"
        End Get

    End Property

    ''' <summary>
    ''' Конструктор, получает файл и начальные настройки для работы с ним
    ''' <para>
    ''' fileName => Начальная ячейка, где находиться ячейка 
    ''' </para>
    ''' 
    ''' <para>
    ''' cellFirst => Начальная ячейка, где находиться таблица (по умолчанию = ячейка "A5")
    ''' </para>
    ''' 
    ''' <para>
    ''' columnRename => Ячейка там, где находится столбец в котором размещен старый шифр и который требуется изменить (по умолчанию = ячейка "C5")
    ''' </para>
    ''' 
    ''' </summary>
    ''' <param name="fileName">Путь файла там, где запускается программа</param>
    ''' <param name="cellFirst">Начальная ячейка, где находиться таблица</param>
    ''' <remarks></remarks>
    Public Sub New(fileName As String,
                   Optional ByVal cellFirst As String = "A6")

        path_ = Environment.CurrentDirectory
        'Dim path_s As String = Path.GetFullPath(Path.Combine(path_, "..\..\" & fileName))
        Dim path_s As String = Path.GetFullPath(Path.Combine(path_, "..\..\" & fileName))

        Try
            app = GetObject(, "Excel.Application")
        Catch ex As Exception
            app = CreateObject("Excel.Application")
        End Try

        If app.ScreenUpdating = False Then app.ScreenUpdating = True

        wbook = app.Workbooks.Add(path_s)
        sheet = wbook.ActiveSheet

        Me.cellFirst = sheet.Range(cellFirst)


    End Sub

    Sub Visible(visible_ As Boolean)
        app.Visible = visible_
    End Sub

    

    'sheet.Columns("A:A").ColumnWidth() = 6


    ''' <summary>
    ''' 
    ''' Настройка листа.
    ''' 
    ''' <para>
    ''' 1. Настройка ширины колонок
    ''' </para>
    ''' 
    ''' <para>
    ''' 2. Настройка столбца 'Наименование' - перенос по словам и определенной ширины
    ''' </para>
    ''' 
    ''' </summary>
    ''' 
    ''' <remarks></remarks>
    Public Sub SheetSettings()

        sheet.Range("A1").Value = "Оборотно-сальдова відомість"
        sheet.Range("A2").Value = ""
        sheet.Range("A3").Value = ""
        sheet.Range("A4").Value = ""
        sheet.Range("A5").Value = ""

        '"Відповідальний:                                               __________          "

        sheet.Columns("A:A").ColumnWidth() = 6
        sheet.Columns("B:B").ColumnWidth() = 6
        sheet.Columns("C:C").ColumnWidth() = 57
        sheet.Columns("D:D").ColumnWidth() = 15
        sheet.Columns("E:E").ColumnWidth() = 11.14
        sheet.Columns("F:F").ColumnWidth() = 7
        sheet.Columns("G:G").ColumnWidth() = 11.14
        sheet.Columns("H:H").ColumnWidth() = 11.14

        sheet.Columns("I:I").ColumnWidth() = 11.14
        sheet.Columns("J:J").ColumnWidth() = 11.14
        sheet.Columns("K:K").ColumnWidth() = 11.14
        sheet.Columns("L:L").ColumnWidth() = 11.14
        sheet.Columns("M:M").ColumnWidth() = 11.14
        sheet.Columns("N:N").ColumnWidth() = 11.14
        sheet.Columns("O:O").ColumnWidth() = 11.14

        Dim rngCC As Excel.Range = sheet.Columns("C:C")
        With rngCC
            .WrapText = True        '    .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
        End With
    End Sub

    ''' <summary>
    ''' Столбец в котором размещен старый шифр и который требуется изменить (по умолчанию = столбец "C5")
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RenameRange(Optional ByVal columnRenameString As String = "D")

        Me.columnRename = sheet.Range(columnRenameString & cellFirst.Row)

        app.ScreenUpdating = False


        rngCount = app.Range(app.Selection, app.Selection.End(Excel.XlDirection.xlDown))

        Dim rCount As Integer
        rCount = rngCount.Count

        Dim currentCell As Excel.Range

        For i = 1 To rCount
            currentCell = columnRename.Offset(i, 0)

            currentCell.Value = Replace(currentCell.Value, ".", "")

            currentCell.NumberFormat = "0"
            currentCell.HorizontalAlignment = Excel.Constants.xlLeft

        Next

        app.ScreenUpdating = True
    End Sub

    ''' <summary>
    ''' Создание объекта таблице в книге Excel (по умолчанию объект таблица = "table1")
    ''' </summary>
    ''' <param name="objectName">Имя объекта Таблицы</param>
    ''' <remarks></remarks>
    Public Sub tableCreateListObject(Optional ByVal objectName = "table1")

        Dim r1 As Excel.Range = app.Range(cellFirst, cellFirst.End(Excel.XlDirection.xlToRight))
        Dim r2 As Excel.Range = app.Range(cellFirst, cellFirst.End(Excel.XlDirection.xlDown))

        Dim table1 As Excel.Range = app.Range(r1, r2)

        tableObject = sheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, table1)
        tableObject.Name = objectName

    End Sub

    ''' <summary>
    ''' Сортировка столбца происходит по столбцу там, где обозначается старый номенклатурный номер (по умолчанию = столбец №3)
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub sortTable()
        'If (Not column = 3) Then column = columnRename.Column

        Dim column = Me.columnRename.Column

        tableObject.Range.Sort( _
        Key1:=tableObject.ListColumns(column).Range, Order1:=Excel.XlSortOrder.xlAscending, _
        Key2:=tableObject.ListColumns(column).Range, Order2:=Excel.XlSortOrder.xlAscending, _
        Orientation:=Excel.XlSortOrientation.xlSortColumns, _
        Header:=Excel.XlYesNoGuess.xlYes)
    End Sub

    ''' <summary>
    ''' Выравнивание строк по содержимому
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub EntireRowAutoFit()

        app.ScreenUpdating = False

        cellFirst.Activate()

        Dim rngCount As Excel.Range
        rngCount = app.Range(app.Selection, app.Selection.End(Excel.XlDirection.xlDown))

        Dim rCount As Integer
        rCount = rngCount.Count

        For i = cellFirst.Row + 1 To rngCount.Count
            Dim sRow As String = i & ":" & i
            sheet.Rows(sRow).EntireRow.AutoFit()
        Next

        app.ScreenUpdating = True


    End Sub

    Public Sub tableHeaderColor()
        Dim r3 As Excel.Range = sheet.Range(cellFirst, cellFirst.End(Excel.XlDirection.xlToRight))

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

    ''' <summary>
    ''' Создание нумереции строк
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateCounterRow()

        app.ScreenUpdating = False

        cellFirst.Activate()

        cellFirst.Value = "№ п/п"

        Dim rngCount As Excel.Range
        rngCount = app.Range(app.Selection, app.Selection.End(Excel.XlDirection.xlDown))

        Dim rCount As Integer
        rCount = rngCount.Count

        Dim currentCell As Excel.Range

        For i = 1 To rCount - 1
            currentCell = cellFirst.Offset(i, 0)
            currentCell.Value = i
        Next

        sheet.Range("A" & (6 + rCount + 2)).Value = "Відповідальний:                                               __________          "

        app.ScreenUpdating = True
    End Sub

    ''' <summary>
    ''' Настройка страницы печати (А3, отступы, масштаб = 95%)
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PageSettings()
        With sheet.PageSetup

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







End Class
