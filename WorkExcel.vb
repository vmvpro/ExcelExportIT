Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Text

Public Class WorkExcel
    Private app_ As Excel.Application
    Private wbook_ As Excel.Workbook
    Public sheet_ As Excel.Worksheet
    Private path_ As String

    ''' <summary>
    ''' Ячейка с которой начинается таблица на листе
    ''' </summary>
    ''' <remarks></remarks>
    Dim cellFirst_ As Excel.Range
    Public Property CellFirst As Excel.Range
        Get
            Return cellFirst_
        End Get
        Set(value As Excel.Range)
            cellFirst_ = value
        End Set
    End Property

    ''' <summary>
    ''' Перенаименование и сортировка по этому столбцу
    ''' </summary>
    ''' <remarks></remarks>
    Dim columnRename_ As Excel.Range
    Public Property ColumnRename As Excel.Range
        Get
            Return columnRename_
        End Get
        Set(value As Excel.Range)
            columnRename_ = value
        End Set
    End Property



    ''' <summary>
    ''' Количество строк документа без шапки
    ''' </summary>
    ''' <remarks></remarks>
    Dim rowCount_ As Excel.Range
    Public Property RowCount As Excel.Range
        Get
            Return rowCount_
        End Get
        Set(value As Excel.Range)
            rowCount_ = value
        End Set
    End Property

    Dim tableObject As Excel.ListObject

    Public Shared ReadOnly Property PathDirectoryNetwork As String
        Get
            Return "\\erpdb\TEMP\OSV\Files"
        End Get
    End Property

    Public Shared ReadOnly Property PathDirectoryLocal As String
        Get
            Dim path_s As String = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "..\..\Files"))
            Return path_s
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
    Public Sub New(fileName As String, cellFirst As String)

        path_ = Environment.CurrentDirectory
        'Dim path_s As String = Path.GetFullPath(Path.Combine(path_, "..\..\" & fileName))
        Dim path_s As String = Path.GetFullPath(Path.Combine(path_, "..\..\" & fileName))

        Try
            app_ = GetObject(, "Excel.Application")
        Catch ex As Exception
            app_ = CreateObject("Excel.Application")
        End Try

        If app_.ScreenUpdating = False Then app_.ScreenUpdating = True

        wbook_ = app_.Workbooks.Add(path_s)
        sheet_ = wbook_.ActiveSheet

        Me.cellFirst_ = sheet_.Range(cellFirst)

    End Sub

    Private ReadOnly Property App As Excel.Application
        Get
            Return app_
        End Get
    End Property

    Private ReadOnly Property WorkBook As Excel.Workbook
        Get
            Return wbook_
        End Get
    End Property

    Private ReadOnly Property ActiveSheet As Excel.Worksheet
        Get
            Return sheet_
        End Get
    End Property

    Public Sub New(fileName As String)

        path_ = Environment.CurrentDirectory
        'Dim path_s As String = Path.GetFullPath(Path.Combine(path_, "..\..\" & fileName))
        Dim path_s As String = Path.Combine(path_, "..\..\" & fileName)

        Try
            app_ = GetObject(, "Excel.Application")
        Catch ex As Exception
            app_ = CreateObject("Excel.Application")
        End Try

        If app_.ScreenUpdating = False Then app_.ScreenUpdating = True

        wbook_ = app_.Workbooks.Add(path_s)
        sheet_ = wbook_.ActiveSheet

    End Sub

    Sub Visible(visible_ As Boolean)
        app_.Visible = visible_
    End Sub

    Public Sub SaveExcel(fileName As String)

        Dim path_ As String = Path.Combine(DirectoryExcel(), fileName & ".xlsx")
        Dim fi As New FileInfo(path_)
        If fi.Exists Then
            fi.Delete()
        End If

        wbook_.SaveAs(path_, Excel.XlFileFormat.xlWorkbookDefault)
    End Sub

    Private Function DirectoryExcel() As String

        Dim directoryExcel_ As New DirectoryInfo(Path.Combine(PathDirectoryNetwork, Environment.UserName, "Excel"))
        directoryExcel_.Create()

        Return directoryExcel_.FullName

    End Function

    Public Sub SavePdf(fileName As String)
        Dim pathSave = Path.Combine(DirectoryPdf(), fileName & ".pdf")
        wbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pathSave)
    End Sub

    Private Function DirectoryPdf() As String

        Dim directoryPdf_ As New DirectoryInfo(Path.Combine(PathDirectoryNetwork, Environment.UserName, "Pdf"))
        directoryPdf_.Create()

        Return directoryPdf_.FullName

    End Function




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

        sheet_.Range("A1").Value = "Оборотно-сальдова відомість"
        sheet_.Range("A2").Value = ""
        sheet_.Range("A3").Value = ""
        sheet_.Range("A4").Value = ""
        sheet_.Range("A5").Value = ""

        sheet_.Columns("A:A").ColumnWidth() = 6
        sheet_.Columns("B:B").ColumnWidth() = 6
        sheet_.Columns("C:C").ColumnWidth() = 57
        sheet_.Columns("D:D").ColumnWidth() = 15
        sheet_.Columns("E:E").ColumnWidth() = 11.14
        sheet_.Columns("F:F").ColumnWidth() = 7
        sheet_.Columns("G:G").ColumnWidth() = 11.14
        sheet_.Columns("H:H").ColumnWidth() = 11.14

        sheet_.Columns("I:I").ColumnWidth() = 11.14
        sheet_.Columns("J:J").ColumnWidth() = 11.14
        sheet_.Columns("K:K").ColumnWidth() = 11.14
        sheet_.Columns("L:L").ColumnWidth() = 11.14
        sheet_.Columns("M:M").ColumnWidth() = 11.14
        sheet_.Columns("N:N").ColumnWidth() = 11.14
        sheet_.Columns("O:O").ColumnWidth() = 11.14

        Dim rngCC As Excel.Range = sheet_.Columns("C:C")
        With rngCC
            .WrapText = True        '    .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
        End With
    End Sub

    ''' <summary>
    ''' Столбец в котором размещен старый шифр и который требуется изменить (по умолчанию = столбец "D6")
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RenameRange(Optional ByVal columnRenameString As String = "D")

        Me.columnRename_ = sheet_.Range(columnRenameString & cellFirst_.Row)

        app_.ScreenUpdating = False

        rowCount_ = app_.Range(app_.Selection, app_.Selection.End(Excel.XlDirection.xlDown))

        Dim rCount As Integer
        rCount = rowCount_.Count

        Dim currentCell As Excel.Range

        For i = 1 To rCount
            currentCell = columnRename_.Offset(i, 0)

            currentCell.Value = Replace(currentCell.Value, ".", "")

            currentCell.NumberFormat = "0"
            currentCell.HorizontalAlignment = Excel.Constants.xlLeft

        Next

        app_.ScreenUpdating = True
    End Sub

    ''' <summary>
    ''' Создание объекта таблице в книге Excel (по умолчанию объект таблица = "table1")
    ''' </summary>
    ''' <param name="objectName">Имя объекта Таблицы</param>
    ''' <remarks></remarks>
    Public Sub tableCreateListObject(Optional ByVal objectName = "table1")

        Dim r1 As Excel.Range = app_.Range(cellFirst_, cellFirst_.End(Excel.XlDirection.xlToRight))
        Dim r2 As Excel.Range = app_.Range(cellFirst_, cellFirst_.End(Excel.XlDirection.xlDown))

        Dim table1 As Excel.Range = app_.Range(r1, r2)

        tableObject = sheet_.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, table1)
        tableObject.Name = objectName

    End Sub

    ''' <summary>
    ''' Сортировка столбца происходит по столбцу там, где обозначается старый номенклатурный номер (по умолчанию = столбец №3)
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub sortTable()

        Dim column = Me.columnRename_.Column

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

        app_.ScreenUpdating = False

        cellFirst_.Activate()

        Dim rngCount As Excel.Range
        rngCount = app_.Range(app_.Selection, app_.Selection.End(Excel.XlDirection.xlDown))

        Dim rCount As Integer
        rCount = rngCount.Count

        For i = cellFirst_.Row + 1 To rngCount.Count
            Dim sRow As String = i & ":" & i
            sheet_.Rows(sRow).EntireRow.AutoFit()
        Next

        app_.ScreenUpdating = True


    End Sub

    Public Sub tableHeaderColor()
        Dim r3 As Excel.Range = sheet_.Range(cellFirst_, cellFirst_.End(Excel.XlDirection.xlToRight))

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

        app_.ScreenUpdating = False

        cellFirst_.Activate()

        cellFirst_.Value = "№ п/п"

        Dim rngCount As Excel.Range = app_.Range(app_.Selection, app_.Selection.End(Excel.XlDirection.xlDown))

        Dim rCount As Integer = rngCount.Count

        Dim currentCell As Excel.Range

        For i = 1 To rCount - 1
            currentCell = cellFirst_.Offset(i, 0)
            currentCell.Value = i
        Next

        sheet_.Range("A" & (6 + rCount + 2)).Value = "Відповідальний:                                               __________          "

        app_.ScreenUpdating = True
    End Sub

    ''' <summary>
    ''' Настройка страницы печати (А3, отступы, масштаб = 95%)
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PageSettings()
        With sheet_.PageSetup

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
