Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Text

Public Class frmMain

    'Dim app As Excel.Application
    'Private wbook As Excel.Workbook
    'Public sheet As Excel.Worksheet


    Dim dt As New DataTable

    Private Sub btnCreateOSV_Click(sender As Object, e As EventArgs) Handles btnCreateOSV.Click
        Label2.Text = ""

        If (cboMonth.Text = String.Empty) Then
            MessageBox.Show("Выберите месяц по которому формируете Оборотно-сальдовую ведомость", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim s As String


        Try
            If chkIsAllCeh.Checked Then
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

                MySub(ceh2, cboMonth.Text)

                Me.Visible = True

                MessageBox.Show("Формирование завершенно успешно." + vbNewLine +
                                "Для перехода к папке сохранения, нажмите на кнопку открыть!", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Exit Sub
            End If
        Catch ex As Exception
            Label2.Text = "По " & ceh2 & " ничего не найдено!!!"
            Debug.WriteLine(ex.Message & Environment.NewLine + "По " & ceh2 & " ничего не найдено!!!")

            'wbook.Close(False)
            'app.WindowState = Excel.XlWindowState.xlMinimized
            Me.Visible = True

            Exit Sub

        End Try


    End Sub

    Function DataTableTBRL() As String
        Dim sb As New StringBuilder

        Me.Visible = False
        For Each row As DataRow In dt.Rows
            Try
                MySub(row("it").ToString(), cboMonth.Text)
            Catch ex As Exception
                sb.Append(row("it").ToString() & Environment.NewLine & ex.Message + Environment.NewLine)
            End Try
        Next

        Me.Visible = True

        MessageBox.Show("Формирование завершенно успешно." + vbNewLine +
                                "Для перехода к папке сохранения, нажмите на кнопку 'Открыть файлы'.", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Return sb.ToString

    End Function


    Sub MySub(ceh As String, monthOSV As String)

        Dim pathFile As String = Path.Combine(WorkExcel.PathDirectoryNetwork, monthOSV & ".xlsx")

        Dim excel_ As New WorkExcel(pathFile)

        excel_.Visible(True)

        '=============================

        'excel_.ScreenUpdating(False)

        excel_.AutoFilter(ceh)

        excel_.WriteHeaderCells(cboMonth.Text, ceh)

        excel_.ColumnHiddenCeh("E:E")

        'excel_.ScreenUpdating(True)

        '============================================================

        Dim fileName As String = ceh & "_" & monthOSV

        excel_.SaveExcel(fileName)
        excel_.SavePdf(fileName)

        '==========================================================

        excel_.WorkBookClose()

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
        Dim path_ = Path.Combine(WorkExcel.PathDirectoryNetwork, "Months.dat")

        Dim files As New StreamReader(path_)
        Dim filesArray = File.ReadLines(path_).ToArray

        cboMonth.Items.Add("")

        For i As Int32 = 0 To filesArray.Length - 1
            cboMonth.Items.Add(filesArray(i))
        Next

        cboMonth.SelectedIndex = 0

    End Sub

    Private Sub cbo_MonthOSV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMonth.SelectedIndexChanged

        cboCeh.DataSource = LoadComboBox()
        cboCeh.DisplayMember = "ceh"
        cboCeh.ValueMember = "it"

        cboCeh.SelectedIndex = 0

    End Sub

    Dim ceh2 As String = ""
    Private Sub cbo_Ceh_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCeh.SelectedIndexChanged

        Dim ceh As String = ""
        ceh = DirectCast(cboCeh.SelectedItem, System.Data.DataRowView).Row.ItemArray(0).ToString
        ceh2 = ceh

    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles chkIsAllCeh.CheckedChanged
        cboCeh.Enabled = Not chkIsAllCeh.Checked
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim di As New DirectoryInfo(WorkExcel.PathDirectoryNetwork & "\" & Environment.UserName & "\PDF")
        di.Create()

    End Sub

    Private Sub btnOpenCurrentDirectory_Click(sender As Object, e As EventArgs) Handles btnOpenCurrentDirectory.Click, btnOpenFiles.Click
        System.Diagnostics.Process.Start("explorer", WorkExcel.PathDirectoryNetwork & "\" & Environment.UserName)
    End Sub

    Private Sub btnSettings_Click(sender As Object, e As EventArgs) Handles btnSettings.Click

        
        SettingsSheetExcel.Run()
        

        ' ВНИМАНИЕ!!!
        ' Не забыть пересохранить файл Ексель, 
        ' так как файл Excel после сортировки таким образом меняет структуру
        ' после которой невозможно открыть файл программным образом
        '
        ' Т.е. файл  сначала закрыть, потом открыть и на все всплывающие окна ответить 
        ' положительно, затем следует пересохранить с тем же именем.

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        System.Diagnostics.Process.Start("explorer", WorkExcel.PathDirectoryNetwork & "\Files")
    End Sub

    Private Sub Vetal(fileName As String, monthOSV As String)
        Try
            Dim fi As New FileInfo("d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\TempExcel\" & fileName & "_" & monthOSV & ".xls")
            If fi.Exists Then
                fi.Delete()
            End If
            'wbook.SaveAs("d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\TempExcel\" & fileName & "_" & monthOSV, Excel.XlFileFormat.xlWorkbookDefault)
        Catch ex As Exception

            Dim fi As New FileInfo("d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\TempExcel\" & fileName & "_" & monthOSV & ".xls")
            If fi.Exists Then
                fi.Delete()
            End If

            'wbook.SaveAs("d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\TempExcel\" & fileName & "_" & monthOSV, Excel.XlFileFormat.xlExcel8)
        End Try

        'wbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, "d:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelExportIT\PDF\" & fileName & "_" & monthOSV & ".pdf")
    End Sub
End Class
