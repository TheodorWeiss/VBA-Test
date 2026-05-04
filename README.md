Окей, делаем нормальную схему:

Планировщик → cmd.exe → ставит ENV → открывает Excel-файл → Workbook_Open видит ENV → запускает макрос → обновляет → сохраняет → логирует → закрывает Excel.

1. В ThisWorkbook

Private Sub Workbook_Open()
    If Environ("RUN_AUTOMATION") = "1" Then
        Call AutoUpdateByScheduler
    End If
End Sub

2. В обычный VBA-модуль

Sub AutoUpdateByScheduler()
    On Error GoTo ErrorHandler
    Dim logPath As String
    Dim startTime As Date
    startTime = Now
    logPath = ThisWorkbook.Path & "\scheduler_log.txt"
    Call WriteLog(logPath, "START update")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    ThisWorkbook.Save
    Call WriteLog(logPath, "SUCCESS update. Duration: " & Format(Now - startTime, "hh:nn:ss"))
CleanExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Quit
    Exit Sub
ErrorHandler:
    Call WriteLog(logPath, "ERROR: " & Err.Number & " - " & Err.Description)
    Resume CleanExit
End Sub
Sub WriteLog(logPath As String, message As String)
    Dim f As Integer
    f = FreeFile
    Open logPath For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #f
End Sub

3. В Aufgabenplanung → Aktion

Programm/Skript:

cmd.exe

Argumente hinzufügen:

/c set RUN_AUTOMATION=1 && start "" "C:\Pfad\DeineDatei.xlsm"

Пример:

/c set RUN_AUTOMATION=1 && start "" "C:\Users\Theo\Documents\Report.xlsm"

Так мы не ищем EXCEL.EXE, а открываем сам файл через ассоциацию Windows.

4. Настройки задачи

Во вкладке Allgemein:

Nur ausführen, wenn der Benutzer angemeldet ist

Галочку Mit höchsten Privilegien ausführen лучше снять.

5. Retry в планировщике

Открой задачу → Eigenschaften → вкладка Einstellungen.

Поставь:

Falls Aufgabe fehlschlägt, Neustart alle: 5 Minuten

и:

Neustartversuche: 3

По-немецки это может быть примерно:

Bei Fehler alle 5 Minuten neu starten
Maximal 3 Neustartversuche

6. Проверка

1. Сохрани файл как .xlsm
2. Закрой Excel полностью
3. В Aufgabenplanung нажми правой кнопкой по задаче → Ausführen
4. Проверь рядом с Excel-файлом файл:

scheduler_log.txt

Там должно появиться что-то вроде:

2026-05-04 07:00:01 | START update
2026-05-04 07:01:34 | SUCCESS update. Duration: 00:01:33

Главный плюс: при обычном ручном открытии файл не будет обновляться, потому что RUN_AUTOMATION не равен 1.
