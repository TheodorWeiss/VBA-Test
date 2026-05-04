Лучше перейти на более надёжный вариант: VBS напрямую открывает Excel и запускает макрос, без Workbook_Open и ENV.

1. Создай файл, например на Desktop

run_excel_update.vbs

Внутрь:

Dim xl
Dim wb
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open("C:\Pfad\DeineDatei.xlsm")
xl.Run "'" & wb.Name & "'!AutoUpdateByScheduler"
wb.Save
wb.Close False
xl.Quit
Set wb = Nothing
Set xl = Nothing

Замени путь:

C:\Pfad\DeineDatei.xlsm

на реальный путь к файлу.

2. В Excel убери временно Workbook_Open

Или оставь пустым:

Private Sub Workbook_Open()
End Sub

Макрос AutoUpdateByScheduler оставь в обычном модуле.

3. В Aufgabenplanung

Aktion → Programm/Skript:

wscript.exe

Argumente hinzufügen:

"C:\Users\...\Desktop\run_excel_update.vbs"

4. Проверь вручную

Двойной клик по run_excel_update.vbs.

Если всё хорошо, Excel должен:
открыться → запустить макрос → обновить сводные → сохранить → закрыться.

Это чище, чем Workbook_Open: при обычном открытии файл не висит, а по расписанию макрос запускается напрямую.
