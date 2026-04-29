Sub SetCurrentYearWeek_CorrectMDX_v2()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pivotNames As Variant
    Dim p As Variant
    Dim yearWeekKey As Long
    Dim mdxValue As String
    Dim changedCount As Long
    Dim skippedCount As Long
    Dim errText As String

    yearWeekKey = Year(Date) * 100 + WorksheetFunction.ISOWeekNum(Date)
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
                On Error GoTo 0

                If Not pf Is Nothing Then

                    On Error Resume Next
                    Err.Clear

                    pf.VisibleItemsList = Array(mdxValue)

                    If Err.Number = 0 Then
                        pt.RefreshTable
                        changedCount = changedCount + 1
                    Else
                        errText = errText & pt.Name & ": " & Err.Description & vbCrLf
                        Err.Clear
                        skippedCount = skippedCount + 1
                    End If

                    On Error GoTo 0

                Else
                    skippedCount = skippedCount + 1
                    errText = errText & pt.Name & ": поле Woche не найдено" & vbCrLf
                End If

            End If

        Next p
    Next ws

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Готово" & vbCrLf & _
           "Ключ: " & yearWeekKey & vbCrLf & _
           "Обновлено: " & changedCount & vbCrLf & _
           "Пропущено: " & skippedCount & vbCrLf & vbCrLf & _
           errText

End Sub
