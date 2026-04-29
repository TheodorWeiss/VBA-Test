Да, для иерархии лучше использовать VisibleItemsList, а не CurrentPage.

Попробуй такой макрос:

Sub SetCurrentYearAndWeekForSelectedPivots()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pivotNames As Variant
    Dim p As Variant
    
    Dim currentYear As Long
    Dim currentWeek As Long
    Dim mdxWeek As String
    Dim changedCount As Long
    Dim skippedCount As Long
    currentYear = Year(Date)
    currentWeek = WorksheetFunction.ISOWeekNum(Date)
    pivotNames = Array("PivotTable1", "PivotTable2", "PivotTable3", "PivotTable4", "PivotTable9", "PivotTable12")
    ' ВАЖНО: этот MDX-путь может потребовать корректировки под твой куб
    mdxWeek = "[Datum].[KalenderWoche].&[" & currentYear & "].&[" & currentWeek & "]"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    For Each ws In ThisWorkbook.Worksheets
        For Each p In pivotNames
            Set pt = Nothing
            Set pf = Nothing
            On Error Resume Next
            Set pt = ws.PivotTables(CStr(p))
            On Error GoTo 0
            If Not pt Is Nothing Then
                On Error Resume Next
                Set pf = pt.PivotFields("[Datum].[KalenderWoche].[KalenderWoche]")
                If pf Is Nothing Then Set pf = pt.PivotFields("Datum.KalenderWoche")
                On Error GoTo 0
                If Not pf Is Nothing Then
                    On Error Resume Next
                    ' Если уже стоит нужная неделя — ничего не делаем
                    If Not IsEmpty(pf.VisibleItemsList) Then
                        If UBound(pf.VisibleItemsList) >= 0 Then
                            If pf.VisibleItemsList(0) = mdxWeek Then
                                skippedCount = skippedCount + 1
                                GoTo NextPivot
                            End If
                        End If
                    End If
                    Err.Clear
                    pf.VisibleItemsList = Array(mdxWeek)
                    If Err.Number = 0 Then
                        changedCount = changedCount + 1
                    Else
                        skippedCount = skippedCount + 1
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
NextPivot:
        Next p
    Next ws
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Готово." & vbCrLf & _
           "Год: " & currentYear & vbCrLf & _
           "Неделя: " & currentWeek & vbCrLf & _
           "Обновлено сводных: " & changedCount & vbCrLf & _
           "Пропущено: " & skippedCount
End Sub

Если не сработает, почти наверняка проблема только в этой строке:

mdxWeek = "[Datum].[KalenderWoche].&[" & currentYear & "].&[" & currentWeek & "]"

Тогда нужно один раз получить точный MDX-ключ выбранной недели через диагностический макрос.
