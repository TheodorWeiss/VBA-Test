Sub DiagnoseWeekFilterToSheet()

    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim outWs As Worksheet
    Dim r As Long
    Dim i As Long

    Set pt = ActiveSheet.PivotTables("PivotTable1")

    Set outWs = ThisWorkbook.Worksheets.Add
    outWs.Name = "Pivot_Diagnose"

    r = 1
    outWs.Cells(r, 1).Value = "PivotTable"
    outWs.Cells(r, 2).Value = pt.Name
    r = r + 2

    For Each pf In pt.PivotFields
        If InStr(1, pf.Name, "KalenderWoche", vbTextCompare) > 0 _
           Or InStr(1, pf.Caption, "KalenderWoche", vbTextCompare) > 0 _
           Or InStr(1, pf.Name, "Woche", vbTextCompare) > 0 Then

            outWs.Cells(r, 1).Value = "FIELD NAME"
            outWs.Cells(r, 2).Value = pf.Name
            r = r + 1

            outWs.Cells(r, 1).Value = "CAPTION"
            outWs.Cells(r, 2).Value = pf.Caption
            r = r + 1

            outWs.Cells(r, 1).Value = "SOURCE NAME"
            outWs.Cells(r, 2).Value = pf.SourceName
            r = r + 1

            outWs.Cells(r, 1).Value = "CUBE FIELD"
            outWs.Cells(r, 2).Value = pf.CubeField.Name
            r = r + 2

            outWs.Cells(r, 1).Value = "VisibleItemsList"
            r = r + 1

            On Error Resume Next
            For i = LBound(pf.VisibleItemsList) To UBound(pf.VisibleItemsList)
                outWs.Cells(r, 1).Value = pf.VisibleItemsList(i)
                r = r + 1
            Next i
            On Error GoTo 0

            r = r + 1
            outWs.Cells(r, 1).Value = "Visible PivotItems"
            r = r + 1

            On Error Resume Next
            For Each pi In pf.PivotItems
                If pi.Visible Then
                    outWs.Cells(r, 1).Value = pi.Name
                    outWs.Cells(r, 2).Value = pi.Caption
                    r = r + 1
                End If
            Next pi
            On Error GoTo 0

            r = r + 2

        End If
    Next pf

    outWs.Columns.AutoFit

End Sub
