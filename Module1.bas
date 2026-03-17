Attribute VB_Name = "Module1"
Sub automate()

    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    
    Set wsData = ActiveSheet
    
    With wsData.Rows("1:1")
        .Font.Bold = True
        .Interior.Color = RGB(50, 50, 50)
        .Font.Color = RGB(255, 255, 255)
    End With
    wsData.Cells.EntireColumn.AutoFit
    
  
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Summary_Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsReport = Sheets.Add(After:=wsData)
    wsReport.Name = "Summary_Report"
    
  
    wsReport.Range("A1").Value = "Region"
    wsReport.Range("B1").Value = "Count of Orders"
    wsReport.Range("A1:B1").Font.Bold = True
    

    wsData.Range("M:M").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=wsReport.Range("A2"), Unique:=True
    
    
    Dim i As Long
    For i = 2 To wsReport.Cells(Rows.Count, 1).End(xlUp).Row
        wsReport.Cells(i, 2).Value = Application.WorksheetFunction.CountIf(wsData.Range("M:M"), wsReport.Cells(i, 1).Value)
    Next i
    
    MsgBox "Report tayar hai! 'Summary_Report' sheet check karein.", vbInformation
    
    
End Sub

Sub TEST()

    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    
    ws.Cells(1, 19).Value = "Tax (10%)" '
    ws.Cells(1, 19).Font.Bold = True
    
   
    Dim i As Long
    For i = 2 To lastRow
        
    ws.Cells(i, 19).Value = ws.Cells(i, 18).Value * 0.1
    Next i
    
    MsgBox "Tax calculation complete!", vbInformation
End Sub

Sub CopyHighValueOrders()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long, i As Long, destRow As Long
    
    Set wsSource = ActiveSheet
    
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("High Value Orders").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
   
    Set wsDest = Sheets.Add(After:=wsSource)
    wsDest.Name = "High Value Orders"
    
    
    wsSource.Rows(1).Copy Destination:=wsDest.Rows(1)
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    destRow = 2
    
    
    For i = 2 To lastRow
        If wsSource.Cells(i, 18).Value > 1000 Then
            wsSource.Rows(i).Copy Destination:=wsDest.Rows(destRow)
            destRow = destRow + 1
        End If
    Next i
    
    wsDest.Columns.AutoFit
    MsgBox "High Value Orders are in another sheet ", vbInformation
End Sub

Sub CreatePivotReport()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    Set wsData = ActiveSheet
    
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Pivot_Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    
    Set wsPivot = Sheets.Add(After:=wsData)
    wsPivot.Name = "Pivot_Report"
    
    
    Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)
        
   
    Set pt = pc.CreatePivotTable(TableDestination:=wsPivot.Range("A3"), _
        TableName:="SalesPivot")
        
    
    With pt
        .PivotFields("Category").Orientation = xlRowField
        .AddDataField .PivotFields("Sales"), "Total Sales", xlSum
        .RowAxisLayout xlTabularRow
    End With
    
    MsgBox "Pivot Table tayar hai!", vbInformation
End Sub

