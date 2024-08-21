Sub AuswertungErstellen()
    Dim wsFragenkatalog As Worksheet
    Dim wsHilfstabelle As Worksheet
    Dim wsAuswertung As Worksheet
    Dim tblFragenkatalog As ListObject
    Dim tblAuswertung As ListObject
    Dim lastRowHilfstabelle As Long
    Dim i As Long, j As Long, k As Long, l As Long
    Dim frage As String
    Dim antwort As String
    Dim gewichtung As Double
    Dim antwortKategorie As String
    Dim kategorie As String
    
    ' Set worksheets
    Set wsFragenkatalog = ThisWorkbook.Sheets("Fragenkatalog")
    Set wsHilfstabelle = ThisWorkbook.Sheets("Hilfstabelle Antworten")
    Set wsAuswertung = ThisWorkbook.Sheets("Auswertung")
    
    ' Get the table "Fragenkatalog" and "AuswertungKategorien"
    Set tblFragenkatalog = wsFragenkatalog.ListObjects("Fragenkatalog")
    Set tblAuswertung = wsAuswertung.ListObjects("AuswertungKategorien")
    
    ' Get the last row of Hilfstabelle
    lastRowHilfstabelle = wsHilfstabelle.Cells(wsHilfstabelle.Rows.Count, "A").End(xlUp).Row
    
    Debug.Print "Last Row in Hilfstabelle: " & lastRowHilfstabelle
    
    ' Check if the rows in Fragenkatalog are found
    Debug.Print "Fragenkatalog Rows Count: " & tblFragenkatalog.ListRows.Count
    Debug.Print "Auswertung Rows Count: " & tblAuswertung.ListRows.Count
    
    ' Initialize Auswertung table
    For i = 1 To tblAuswertung.ListRows.Count
        Debug.Print "Initializing Auswertung, Row: " & i
        For j = 2 To tblAuswertung.ListColumns.Count
            tblAuswertung.DataBodyRange(i, j).Value = 0
        Next j
    Next i
    
    ' Loop through each question in Fragenkatalog
    For i = 1 To tblFragenkatalog.ListRows.Count
        frage = tblFragenkatalog.DataBodyRange(i, 2).Value
        gewichtung = tblFragenkatalog.DataBodyRange(i, 3).Value
        antwort = tblFragenkatalog.DataBodyRange(i, 5).Value
        kategorie = tblFragenkatalog.DataBodyRange(i, 1).Value
        
        Debug.Print "Verarbeite Frage: " & frage & ", Gewichtung: " & gewichtung & ", Antwort: " & antwort & ", Kategorie: " & kategorie
        
        ' Find matching answer in Hilfstabelle and get the associated tool
        For j = 2 To lastRowHilfstabelle
            If wsHilfstabelle.Cells(j, 1).Value = antwort Then
                antwortKategorie = wsHilfstabelle.Cells(j, 3).Value
                
                Debug.Print "Gefundene Antwort in Hilfstabelle: " & antwort & ", Kategorie: " & antwortKategorie
                
                ' Find the row and column in AuswertungKategorien
                For k = 1 To tblAuswertung.ListRows.Count
                    If tblAuswertung.DataBodyRange(k, 1).Value = kategorie Then
                        For l = 2 To tblAuswertung.ListColumns.Count
                            If tblAuswertung.HeaderRowRange(1, l).Value = antwortKategorie Then
                                Debug.Print "Kategorie gefunden: " & kategorie & ", Spalte: " & l
                                tblAuswertung.DataBodyRange(k, l).Value = tblAuswertung.DataBodyRange(k, l).Value + gewichtung
                                Debug.Print "Neuer Wert in Auswertung: " & tblAuswertung.DataBodyRange(k, l).Value
                            End If
                        Next l
                    End If
                Next k
            End If
        Next j
    Next i
    
    ' Set the worksheet variable to the desired sheet
    Set ws = ThisWorkbook.Sheets("Auswertung")
    
    ' Activate the worksheet
    ws.Activate
    
    ' Call the CreateColumnChartOnAuswertung macro
    Call CreateColumnChartOnAuswertung
    ' Call the CreatePieChartOnAuswertung macro
    Call CreatePieChartOnAuswertung
    ' Call the WriteMaxAnsatzToC20 macro
    Call WriteMaxAnsatzToC20
    
    If wsAuswertung Is Nothing Then
        MsgBox "Das Arbeitsblatt 'Auswertung' wurde nicht gefunden!"
    Else
        wsAuswertung.Activate
    End If
    
    MsgBox "Auswertung wurde erfolgreich erstellt!"
End Sub

Sub CreateColumnChartOnAuswertung()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As chart
    
    ' Set the worksheet to "Auswertung"
    Set ws = ThisWorkbook.Sheets("Auswertung")
    
    ' Add a new chart object, positioning it next to the table
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(1, 5).Left, Width:=500, Top:=ws.Cells(1, 5).Top, Height:=300)
    Set chart = chartObj.chart
    
    ' Set the data range for the chart
    chart.SetSourceData Source:=ws.Range("A1:D6")
    
    ' Set the chart type to column chart
    chart.ChartType = xlColumnClustered
    
    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Punktzahl nach Kategorien"
    
    ' Set axis titles
    chart.Axes(xlCategory, xlPrimary).HasTitle = True
    chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Kategorien"
    chart.Axes(xlValue, xlPrimary).HasTitle = True
    chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Punktzahl"
    
    ' Add data labels
    Dim series As series
    For Each series In chart.SeriesCollection
        series.HasDataLabels = True
    Next series
    
    ' Set the chart legend position
    chart.Legend.Position = xlLegendPositionBottom
End Sub

Sub CreatePieChartOnAuswertung()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As chart
    
    ' Set the worksheet to "Auswertung"
    Set ws = ThisWorkbook.Sheets("Auswertung")
    
    ' Add a new chart object, positioning it below the column chart
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(1, 5).Left, Width:=500, Top:=ws.Cells(1, 5).Top + 320, Height:=300)
    Set chart = chartObj.chart
    
    ' Set the data range for the chart
    chart.SetSourceData Source:=ws.Range("A10:B13")
    
    ' Set the chart type to pie chart
    chart.ChartType = xlPie
    
    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Gesamtpunktzahl nach Werkzeug"
    
    ' Add data labels
    Dim series As series
    For Each series In chart.SeriesCollection
        series.HasDataLabels = True
    Next series
    
    ' Set the chart legend position
    chart.Legend.Position = xlLegendPositionRight
End Sub
Sub WriteMaxAnsatzToC20()
    Dim ws As Worksheet
    Dim maxValue As Double
    Dim maxAnsatz As String
    Dim i As Integer
    Dim lastRow As Integer
    
    ' Set the worksheet to "Auswertung"
    Set ws = ThisWorkbook.Sheets("Auswertung")
    
    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Initialize the maximum value
    maxValue = -1
    
    ' Loop through the rows to find the maximum value
    For i = 12 To lastRow
        If ws.Cells(i, 2).Value > maxValue Then
            maxValue = ws.Cells(i, 2).Value
            maxAnsatz = ws.Cells(i, 1).Value
        End If
    Next i
    
    ' Write the max Ansatz to cell C20
    ws.Range("C20").Value = maxAnsatz
    
    ' Format the cell C20
    With ws.Range("C20")
        .Font.Color = RGB(255, 0, 0) ' Red color
        .Font.Size = 14 ' Font size 20
        .Font.Name = "Arial" ' Font Arial
    End With
End Sub
