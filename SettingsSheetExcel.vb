
Public Class SettingsSheetExcel
    Public Shared Sub Run()
        ' Файл находится в папке там где запускается исходник (не в папке bin\)
        Dim excel_ As New WorkExcel("2019_Июль_Origin.xlsx", "A6")

        excel_.Visible(True)

        ' Настройка ширины колонок
        excel_.ColumnsWidth()

        ' Переименовать столбец (Переименование столбца со старым кодом )
        excel_.RenameColumn("D")

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
        excel_.ColumnsWidth()

        'Авто-ширина столбца ресурс (перенос по словам и определенной ширины)
        excel_.AutoWidthColumnResources("C:C")

        ' Настройка страницы печати (А3, отступы, масштаб = 95%)
        excel_.PageSettings()

        ' ВНИМАНИЕ!!!
        ' Не забыть пересохранить файл Ексель, 
        ' так как файл Excel после сортировки таким образом меняет структуру
        ' после которой невозможно открыть файл программным образом
    End Sub

End Class
