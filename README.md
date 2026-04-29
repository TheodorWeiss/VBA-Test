Sub SetCurrentYearWeek_CorrectMDX()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pivotNames As Variant
    Dim p As Variant

    Dim currentYear As Long
    Dim currentWeek As Long
    Dim yearWeekKey As Long
    Dim mdxValue As String

    Dim changedCount As Long
    Dim skippedCount As Long

    currentYear = Year(Date)
    currentWeek = WorksheetFunction.ISOWeekNum(Date)

    ' Ключ как в кубе: 202617
    yearWeekKey = currentYear * 100 + currentWeek

    mdxValue = "[Datum].[KalenderWoche].[Woche].&[" & yearWeekKey & "]"

    pivotNames = Array("PivotTable1", "PivotTable2", "PivotTable3", "PivotTable4", "PivotTable9", "PivotTable12")

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
                Set pf = pt.PivotFields("[Datum].[KalenderWoche].[Woche]")
                If pf Is Nothing Then Set pf = pt.PivotFields("Datum.KalenderWoche")
                On Error GoTo 0

                If Not pf Is Nothing Then

                    On Error Resume Next

                    ' Проверка: уже стоит нужная неделя?
                    If Not IsEmpty(pf.VisibleItemsList) Then
                        If UBound(pf.VisibleItemsList) >= 0 Then
                            If pf.VisibleItemsList(0) = mdxValue Then
                                skippedCount = skippedCount + 1
                                GoTo NextPivot
                            End If
                        End If
                    End If

                    Err.Clear
                    pf.VisibleItemsList = Array(mdxValue)

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

    MsgBox "Готово 🚀" & vbCrLf & _
           "Год: " & currentYear & vbCrLf & _
           "Неделя: " & currentWeek & vbCrLf & _
           "Ключ: " & yearWeekKey & vbCrLf & _
           "Обновлено: " & changedCount & vbCrLf & _
           "Пропущено: " & skippedCount

End Sub
