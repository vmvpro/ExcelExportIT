
Public Class SettingsSheetExcel
    Public Shared Sub Run()
        ' Файл находится в папке там где запускается исходник (не в папке bin\)
        Dim excel_ As New WorkExcel("test_ceh05.xlsx", "A6")

        excel_.Visible(True)

        'Заполнение заголовка, очистка ячеек перед таблицей и в конце таблицы ответственного
        excel_.ClearHeaderCells()

        ' Настройка ширины колонок
        excel_.ColumnsWidth()

        ' Переименовать столбец (Переименование столбца со старым кодом )
        excel_.ColumnEditingOldResources("D")

        ' Загловок таблицы сделать цветным
        excel_.TableHeaderColor()

        ' Создание в книге объекта таблица
        excel_.TableCreateListObject()

        ' Сортировка столбца в таблице (старый шифр)
        excel_.SortTable()

        ' Выравнивание строк по содержимому
        excel_.EntireRowAutoFit()

        ' Создание нумереции строк
        excel_.CreateCounterRow()

        ' Настроить колонки листа
        excel_.ColumnsWidth()

        ' Настройки формата столбца (перенос по словам и определенной ширины)
        excel_.AutoHeightColumnResources("C")

        ' Настройка страницы печати (А3, отступы, масштаб = 95%)
        excel_.PageSettings()

        ' ВНИМАНИЕ!!!
        ' Не забыть пересохранить файл Ексель, 
        ' так как файл Excel после сортировки таким образом меняет структуру
        ' после которой невозможно открыть файл программным образом
        '
        ' Т.е. файл  сначала закрыть, потом открыть и на все всплывающие окна ответить 
        ' положительно, затем следует пересохранить с тем же именем.
    End Sub

End Class
