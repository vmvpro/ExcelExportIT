Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Text

Public Class frmMain

    Dim dt As New DataTable
    Private Sub btnCreateOSV_Click(sender As Object, e As EventArgs) Handles btnCreateOSV.Click
        Label2.Text = ""

        If (cbo_MonthOSV.Text = String.Empty) Then
            MessageBox.Show("Выберите месяц по которому формируете Оборотно-сальдовую ведомость", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim s As String


        Try
            If CheckBox1.Checked Then
                Dim result As DialogResult = MessageBox.Show("Формирование отчетов по всем складам будет длительное время, желаете продолжить?", "Оповещение", MessageBoxButtons.YesNo, MessageBoxIcon.Information)

                If result = Windows.Forms.DialogResult.Yes Then
                    s = DataTableTBRL()
                    If s.Length > 0 Then
                        MessageBox.Show(s, "Возможные ошибки!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Else
                    Exit Sub
                End If
            Else
                Me.Visible = False



                MySub(ceh2, cbo_MonthOSV.Text)

                Me.Visible = True

                MessageBox.Show("Формирование завершенно успешно." + vbNewLine +
                                "Для перехода к папке сохранения, нажмите на кнопку открыть!", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Exit Sub
            End If
        Catch ex As Exception
            Label2.Text = "По " & ceh2 & " ничего не найдено!!!"
            Debug.WriteLine(ex.Message & Environment.NewLine + "По " & ceh2 & " ничего не найдено!!!")

            'If app.ScreenUpdating = False Then app.ScreenUpdating = True

            wbook.Close(False)
            app.WindowState = Excel.XlWindowState.xlMinimized
            Me.Visible = True

            Exit Sub

        End Try


    End Sub

    Function DataTableTBRL() As String
        Dim sb As New StringBuilder

        Me.Visible = False
        For Each row As DataRow In dt.Rows

            Try
                MySub(row("it").ToString(), cbo_MonthOSV.Text)
            Catch ex As Exception
                sb.Append(row("it").ToString() & Environment.NewLine & ex.Message + Environment.NewLine)
            End Try

        Next

        Me.Visible = True

        MessageBox.Show("Формирование завершенно успешно." + vbNewLine +
                                "Для перехода к папке сохранения, нажмите на кнопку 'Открыть файлы'.", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Return sb.ToString

    End Function


    Sub MySub(ceh As String, cboMonthOSV As String)

        Dim ofd As New OpenFileDialog()
        ofd.Title = "Выберите файл Excel c оборотно-сальдовой ведомостью для сохранения в PDF"

        If (Environment.UserName = "Vetal") Then '  Or (Environment.UserName = "mww54001") 
            ofd.InitialDirectory = "d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\TempExcel\"
        Else
            'ofd.InitialDirectory = "\\erp\TEMP\Оборотно-сальдовая ведомость (ОСВ)\Excel\"
            ' ofd.InitialDirectory = WorkExcel.PathDirectoryOSV & "\Excel\"
            ofd.InitialDirectory = WorkExcel.PathDirectoryOSV & "\"
        End If

        ofd.Filter = "Файлы формата Excel ( *.xlsx; *.xls; *.xlsm;) | *.xlsm; *.xlsx; *.xls; "

        Dim path As String = ""

        If (Environment.UserName = "mww54001_") Then
            path = Environment.CurrentDirectory & "\Files\" & cboMonthOSV & ".xlsx"
            'MessageBox.Show(path)
        Else

            'path = WorkExcel.PathDirectoryOSV & "\Files\" & cboMonthOSV & ".xlsx"
            'path = Environment.CurrentDirectory & "\Files\" & cboMonthOSV & ".xlsx"
            'path = "\\erpdb\PERSONAL\OSV\Files\" & cboMonthOSV & ".xlsx"

            'path = Environment.CurrentDirectory & "\Files\" & cboMonthOSV & ".xlsx"

            path = "\\erpdb\TEMP\OSV" & "\Files\" & cboMonthOSV & ".xlsx"

        End If

        'path = Environment.CurrentDirectory & "\Вся оборотка (Ver. 3.1).xlsx"
        '\\erp\TEMP\App\Программа Export

        Try
            app = GetObject(, "Excel.Application")
        Catch ex As Exception
            app = CreateObject("Excel.Application")
        End Try

        If app.ScreenUpdating = False Then app.ScreenUpdating = True

        app.Visible = True

        wbook = app.Workbooks.Add(path)


        sheet = wbook.ActiveSheet


        Dim rngBB As Excel.Range
        rngBB = sheet.Columns("D:D")
        '=============================

        Dim rngAA As Excel.Range

        app.ScreenUpdating = False
        '

        Try
            sheet.ShowAllData()
        Catch ex As Exception
            'app.ScreenUpdating = True
        End Try


        'sheet.Range("table1").
        Dim rngFilter As Object = sheet.Range("table1").AutoFilter(Field:=5, Criteria1:=ceh)

        Dim rngRow As Excel.Range = sheet.Range("table1").Find(ceh)
        Dim rngFirst As Excel.Range = sheet.Cells(rngRow.Row, 1)

        rngFirst.Select()

        'rngFirst.Select()

        Dim rngB As Excel.Range = rngFirst.Offset(0, 1)
        rngB.Select()


        'For Each rngRow As Excel.Range In rngFilter
        'rngRow
        'Next


        app.ScreenUpdating = True

        sheet.Range("A6").Select()
        sheet.Range("A7").Activate()

        'Dim rngTable As Excel.Range



        For k = 1 To 4
            app.Range("A" & k).Value = ""
        Next

        'sheet.Range("table1").AutoFilter(4, "Склад 1")

        Dim rngCount As Excel.Range
        'Dim rngFirst As Excel.Range = sheet.Range("B6")

        'rngFirst.Offset(1, 0).Select()

        'sheet.Range("B6").Select()

        'rngCount = app.Range(app.Selection, app.Selection.End(Excel.XlDirection.xlDown))

        'Dim rCount As Integer
        'rCount = rngCount.Count

        'Dim rngEnd As Excel.Range
        'rngEnd = sheet.Range("B" & rCount + 6 + 3)
        'rngEnd.Value = ""

        rngAA = sheet.Columns("A:A")
        'rngAA.ColumnWidth = 4.6

        Dim rng As Excel.Range
        Dim rng2 As Excel.Range

        sheet.Range("A1").Value = "Оборотно-сальдова відомість"
        sheet.Range("A2").Value = "За рахунками: 20, 22, 281"
        sheet.Range("A3").Value = cbo_MonthOSV.Text
        sheet.Range("A4").Value = ceh

        'Dim rngCeh As Excel.Range = sheet.Range("A4")
        'rngCeh.Value = ceh

        'rng = sheet.Range("A6")
        'rng.Value = "№"

        'rngFirst.Value = 1
        'For i = 0 To rCount
        '    rng2 = rngFirst.Offset(i + 1, 0)
        '    rng2.Value += i + 2
        'Next

        sheet.Columns("E:E").Hidden = True
        'Selection.EntireColumn.Hidden = True

        'Dim rngBB As Excel.Range

        'rngBB = sheet.Columns("B:B")

        'rngBB.HorizontalAlignment = Excel.Constants.xlCenter
        'rngBB.VerticalAlignment = Excel.Constants.xlCenter

        '============================================================

        'With sheet.PageSetup

        '.LeftMargin = 0.196850393700787
        '.RightMargin = 0.196850393700787
        '.TopMargin = 0.196850393700787
        '.BottomMargin = 0.196850393700787

        '.LeftMargin = 0.196850393700787
        '.RightMargin = 0.196850393700787
        '.TopMargin = 0.393700787401575
        '.BottomMargin = 0.393700787401575

        '.CenterHorizontally = True
        '.CenterVertically = True

        '.HeaderMargin = 0
        '.FooterMargin = 0

        '.PaperSize = Excel.XlPaperSize.xlPaperA3

        '.Zoom = 100

        '.Orientation = Excel.XlPageOrientation.xlLandscape

        'End With
        sheet.Range("A6").Select()

        Dim fileName As String '= InputBox("Введите номер склада (Пример: 1306)")
        fileName = ceh

        app.ScreenUpdating = True

        'sheet.Range("A7").Activate()
        '==========================================================
        If (Environment.UserName = "Vetal") Then

            Try
                Dim fi As New FileInfo("d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\TempExcel\" & fileName & "_" & cboMonthOSV & ".xls")
                If fi.Exists Then
                    fi.Delete()
                End If
                wbook.SaveAs("d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\TempExcel\" & fileName & "_" & cboMonthOSV, Excel.XlFileFormat.xlWorkbookDefault)
            Catch ex As Exception

                Dim fi As New FileInfo("d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\TempExcel\" & fileName & "_" & cboMonthOSV & ".xls")
                If fi.Exists Then
                    fi.Delete()
                End If

                wbook.SaveAs("d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\TempExcel\" & fileName & "_" & cboMonthOSV, Excel.XlFileFormat.xlExcel8)
            End Try

            wbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, "d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\PDF\" & fileName & "_" & cboMonthOSV & ".pdf")
        Else
            Dim di1 As New DirectoryInfo(WorkExcel.PathDirectoryOSV & "\" & Environment.UserName & "\Excel")
            'Dim di1 As New DirectoryInfo("D:" & "\" & Environment.UserName & "\Excel")
            Try

                di1.Create()
                Dim str As String = di1.FullName & "\" & fileName & "_" & cboMonthOSV & ".xlsx"
                Dim fi As New FileInfo(str)
                If fi.Exists Then
                    fi.Delete()
                End If

                wbook.SaveAs(str, Excel.XlFileFormat.xlWorkbookDefault)
            Catch ex As Exception
                Dim fi As New FileInfo(WorkExcel.PathDirectoryOSV & "\Excel\" & fileName & "_" & cboMonthOSV & ".xlsx")
                If fi.Exists Then
                    fi.Delete()
                End If

                'wbook.SaveAs("\\erp\TEMP\Оборотно-сальдовая ведомость (ОСВ)\Excel\" & fileName & "_Отформатирован", Excel.XlFileFormat.xlExcel8)
            End Try
            Dim di2 As New DirectoryInfo(WorkExcel.PathDirectoryOSV & "\" & Environment.UserName & "\PDF")
            di2.Create()

            wbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, di2.FullName & "\" & fileName & "_" & cboMonthOSV & ".pdf")
        End If

        '==========================================================

        wbook.Close()

        If (app.Workbooks.Count = 0) Then app.Quit()

        'MsgBox("Файлы успешно сформированы!!!", , ceh2)

        'Me.Close()


    End Sub




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'CheckBox1.Visible = False
        Label2.Text = ""

        If (Environment.UserName = "mww54001") Or (Environment.UserName = "Vetal") Then
            btnSettings.Visible = True
            Me.Height = Me.Height + 30
        Else
            btnSettings.Visible = False
        End If


        'Dim path_ = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "\Files\Months.dat"))
        'Dim path_ = Path.Combine(Environment.CurrentDirectory, "Files\Months.dat")
        Dim path_ = Path.Combine("\\erpdb\TEMP\OSV", "Files\Months.dat")

        Dim files As New StreamReader(path_)
        Dim filesArray = File.ReadLines(path_).ToArray

        'TextBox1.Text = txtHelp.ReadToEnd

        cbo_MonthOSV.Items.Add("")

        'While Not files.EndOfStream
        '    cbo_MonthOSV.Items.Add(files.ReadLine())
        'End While

        For i As Int32 = 0 To filesArray.Length - 1
            cbo_MonthOSV.Items.Add(filesArray(i))
        Next


        'cbo_MonthOSV.Items.Add("2017_Октябрь")
        'cbo_MonthOSV.Items.Add("2017_Ноябрь")
        'cbo_MonthOSV.Items.Add("2018_Октябрь")
        'cbo_MonthOSV.Items.Add("ОСВ (Распоряжение)")
        'cbo_MonthOSV.Items.Add("2019_Январь")
        'cbo_MonthOSV.Items.Add("2019_Февраль")
        'cbo_MonthOSV.Items.Add("2019_Март")
        'cbo_MonthOSV.Items.Add("2019_Июнь")

        cbo_MonthOSV.SelectedIndex = 0

    End Sub

    Private Sub cbo_MonthOSV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_MonthOSV.SelectedIndexChanged
        LoadComboBox(cbo_MonthOSV.Text)
    End Sub

    Dim ceh2 As String = ""
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'ceh = ComboBox1.Text
        Dim ceh As String = ""
        ceh = DirectCast(ComboBox1.SelectedItem, System.Data.DataRowView).Row.ItemArray(0).ToString
        ceh2 = ceh

    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ComboBox1.Enabled = Not CheckBox1.Checked
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim di As New DirectoryInfo(WorkExcel.PathDirectoryOSV & "\" & Environment.UserName & "\PDF")
        di.Create()

    End Sub

    Private Sub btnOpenCurrentDirectory_Click(sender As Object, e As EventArgs) Handles btnOpenCurrentDirectory.Click, btnOpenFiles.Click
        System.Diagnostics.Process.Start("explorer", WorkExcel.PathDirectoryOSV & "\" & Environment.UserName)
    End Sub

    Private Sub btnSettings_Click(sender As Object, e As EventArgs) Handles btnSettings.Click

        ' Файл находится в папке там где запускается исходник (не в папке bin\)
        Dim excel_ As New WorkExcel("2019_Июль_Origin.xlsx", "A6")


        excel_.Visible(True)



        ' Переименовать столбец (Переименование столбца со старым кодом )
        excel_.RenameRange("D")

        ' Загловок таблицы сделать цветным
        excel_.tableHeaderColor()

        ' Создание в книге объекта таблица
        excel_.tableCreateListObject()

        ' Сортировка столбца в таблице (старый шифр)
        excel_.sortTable()

        ' Выравнивание строк по содержимому
        excel_.EntireRowAutoFit()

        ' Создание нумереции строк
        excel_.CreateCounterRow()

        ' Настроить колонки листа
        excel_.SheetSettings()

        ' Настройка страницы печати (А3, отступы, масштаб = 95%)
        excel_.PageSettings()

        ' Порядковый номер оборотки после этого зделать
        'Dim cell_ As Range
        'Dim currentCell As Range

        'cell_ = Range("A5")

        'For i = 1 To 65520
        '    currentCell = cell_.Offset(i, 0)
        '    currentCell.Value = i

        'Next

        ' ВНИМАНИЕ!!!
        ' Не забыть пересохранить файл Ексель, 
        ' так как файл Excel после сортировки таким образом меняет структуру
        ' после которой невозможно открыть файл программным образом

    End Sub

    Sub LoadComboBox(fileName As String)

        dt = New DataTable()

        Dim dc1 As DataColumn = New DataColumn("it")
        Dim dc2 As DataColumn = New DataColumn("ceh")

        dt.Columns.AddRange({dc1, dc2})

        If (fileName = "ОСВ (Распоряжение)") Then
            LoadTable.NewLoadTable(dt)
        Else
            LoadTable.LoadTable(dt)
        End If

        ComboBox1.DataSource = dt
        ComboBox1.DisplayMember = "ceh"
        ComboBox1.ValueMember = "it"

        ComboBox1.SelectedIndex = 0


    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        System.Diagnostics.Process.Start("explorer", WorkExcel.PathDirectoryOSV & "\Files")
    End Sub
End Class
