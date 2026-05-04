Значит ENV и Workbook_Open уже работают. Проблема теперь внутри макроса: он либо не доходит до RefreshAll, либо RefreshAll не обновляет именно эти сводные.

Сначала поставь диагностический макрос вместо текущего:

Sub AutoUpdateByScheduler()
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\scheduler_log.txt"
    Call WriteLog(logPath, "START macro")
    Call WriteLog(logPath, "RUN_AUTOMATION=" & Environ("RUN_AUTOMATION"))
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Call WriteLog(logPath, "Before RefreshAll")
    ThisWorkbook.RefreshAll
    Call WriteLog(logPath, "After RefreshAll")
    Application.CalculateUntilAsyncQueriesDone
    Call WriteLog(logPath, "After CalculateUntilAsyncQueriesDone")
    ThisWorkbook.Save
    Call WriteLog(logPath, "SAVED")
    Exit Sub
ErrorHandler:
    Call WriteLog(logPath, "ERROR: " & Err.Number & " - " & Err.Description)
End Sub
Sub WriteLog(logPath As String, message As String)
    Dim f As Integer
    f = FreeFile
    Open logPath For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #f
End Sub

И в ThisWorkbook пока так:

Private Sub Workbook_Open()
    MsgBox "Workbook_Open работает. RUN_AUTOMATION=" & Environ("RUN_AUTOMATION")
    If Environ("RUN_AUTOMATION") = "1" Then
        Call AutoUpdateByScheduler
    End If
End Sub

После запуска проверь рядом с файлом scheduler_log.txt.

Если лога нет, значит AutoUpdateByScheduler не вызывается — тогда, скорее всего, макрос лежит не в обычном модуле или называется иначе.

Если лог есть и доходит до SAVED, но сводные не меняются — тогда ThisWorkbook.RefreshAll не обновляет твои PivotTables. Тогда надо будет обновлять их явно через:

Dim ws As Worksheet
Dim pt As PivotTable
For Each ws In ThisWorkbook.Worksheets
    For Each pt In ws.PivotTables
        pt.PivotCache.Refresh
    Next pt
Next ws

Начни с диагностического лога — он сразу покажет, где именно обрыв.



Нет, вручную создавать не нужно. Макрос должен сам создать scheduler_log.txt.

Если файл не появляется вообще, значит почти наверняка проблема в одном из этих мест:

1. AutoUpdateByScheduler не вызывается
2. WriteLog лежит не в обычном модуле
3. у Excel нет права писать в папку, где лежит файл
4. ThisWorkbook.Path пустой или неудобный, например файл открыт из Teams/SharePoint/временной папки

Сделай самый простой тест.

1. В обычный модуль вставь только это

Sub TestLog()
    Dim logPath As String
    logPath = Environ("USERPROFILE") & "\Desktop\scheduler_log_test.txt"
    Dim f As Integer
    f = FreeFile
    Open logPath For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | TEST"
    Close #f
    MsgBox "Log written to: " & logPath
End Sub

2. Запусти вручную из Excel

Alt + F8 → TestLog → Ausführen

Если на Desktop появился scheduler_log_test.txt, значит запись работает.

Тогда в основном макросе лучше временно заменить строку:

logPath = ThisWorkbook.Path & "\scheduler_log.txt"

на:

logPath = Environ("USERPROFILE") & "\Desktop\scheduler_log.txt"

Так мы уберём проблему с правами/SharePoint/сетевой папкой.

3. Если TestLog тоже не создаёт файл

Тогда Excel/политики безопасности блокируют запись через VBA, но это уже отдельная проблема.

Сначала проверь именно TestLog через Alt + F8.
