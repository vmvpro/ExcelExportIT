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

        Dim path As String = ""

        Dim excel_ As New WorkExcel("2019_Июль_Origin.xlsx")

        path = WorkExcel.PathDirectoryOSV & "\Files\" & cboMonthOSV & ".xlsx"


        Try
            app = GetObject(, "Excel.Application")
        Catch ex As Exception
            app = CreateObject("Excel.Application")
        End Try

        If app.ScreenUpdating = False Then app.ScreenUpdating = True

        app.Visible = True

        wbook = app.Workbooks.Add(path)

        sheet = wbook.ActiveSheet

        '=============================

        app.ScreenUpdating = False
        '

        Try
            sheet.ShowAllData()
        Catch ex As Exception
            'app.ScreenUpdating = True
        End Try

        Dim rngFilter As Object = sheet.Range("table1").AutoFilter(Field:=5, Criteria1:=ceh)

        Dim rngRow As Excel.Range = sheet.Range("table1").Find(ceh)
        Dim rngFirst As Excel.Range = sheet.Cells(rngRow.Row, 1)

        rngFirst.Select()

        Dim rngB As Excel.Range = rngFirst.Offset(0, 1)
        rngB.Select()

        app.ScreenUpdating = True

        sheet.Range("A6").Select()
        sheet.Range("A7").Activate()

        For k = 1 To 5
            app.Range("A" & k).Value = ""
        Next


        Dim rngAA As Excel.Range = sheet.Columns("A:A")

        sheet.Range("A1").Value = "Оборотно-сальдова відомість"
        sheet.Range("A2").Value = "За рахунками: 20, 22, 281"
        sheet.Range("A3").Value = cbo_MonthOSV.Text
        sheet.Range("A4").Value = ceh

        sheet.Columns("E:E").Hidden = True
        
        '============================================================

        Dim fileName As String = ceh

        app.ScreenUpdating = True

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
        Dim path_ = Path.Combine(WorkExcel.PathDirectoryOSV, "Files\Months.dat")

        Dim files As New StreamReader(path_)
        Dim filesArray = File.ReadLines(path_).ToArray

        cbo_MonthOSV.Items.Add("")

        For i As Int32 = 0 To filesArray.Length - 1
            cbo_MonthOSV.Items.Add(filesArray(i))
        Next

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
