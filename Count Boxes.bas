Attribute VB_Name = "Module1"
Sub CreateAndFormatSheets()

    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet
    Dim lastRow As Long, maxValue As Long, i As Long, splitData As Variant
    
    Set ws1 = ThisWorkbook.Sheets(1)
    Set ws2 = ThisWorkbook.Sheets.Add(After:=ws1)
    ws2.Name = "Sheet2"
    
    lastRow = ws1.Cells(ws1.Rows.Count, "L").End(xlUp).Row
    ws1.Range("L1:L" & lastRow).Copy Destination:=ws2.Range("A1")
    
    For i = 1 To lastRow
        If InStr(ws2.Cells(i, 1).Value, "-") > 0 Then
            splitData = Split(ws2.Cells(i, 1).Value, "-")
            ws2.Cells(i, 1).Value = Trim(splitData(0))
            ws2.Cells(i, 2).Value = Trim(splitData(1))
        End If
    Next i
    
    For i = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row To 1 Step -1
        If IsEmpty(ws2.Cells(i, 2).Value) Then ws2.Rows(i).Delete
    Next i
    
    ws2.Range("A1:B" & ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row).Sort Key1:=ws2.Range("A1"), Header:=xlYes
    
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ws2)
    ws3.Name = "Sheet3"
    
    maxValue = 0
    lastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row

    For i = 1 To lastRow
        If ws2.Cells(i, 1).Value < 700 And ws2.Cells(i, 1).Value > maxValue Then
            maxValue = ws2.Cells(i, 1).Value
        End If
    Next i
    
    For i = 1 To maxValue
        ws3.Cells(i, 1).Value = i - 1
        ws3.Cells(i, 2).FormulaArray = "=MAX(IF(Sheet2!A:A=A" & i & ", Sheet2!B:B))"
    Next i
    
    ws3.Range("A1").Value = "Cells"
    ws3.Range("B1").Value = "Count"

    Set ws4 = ThisWorkbook.Sheets.Add(After:=ws3)
    ws4.Name = "Summary"
    
    ws3.UsedRange.Copy
    ws4.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    lastRow = ws4.Cells(ws4.Rows.Count, 2).End(xlUp).Row
    ws4.Range("A1:B" & lastRow).Sort Key1:=ws4.Range("B1"), Order1:=xlDescending, Header:=xlYes
    ws4.Cells(2, 3).Value = 0
    ws4.Cells(1, 3).Value = "Boxes Approximately"
    For i = 2 To maxValue
        If ws4.Cells(i, 2).Value >= 6 Then
            ws4.Cells(i, 2).Interior.Color = RGB(220, 20, 60)
            ws4.Cells(2, 3).Value = ws4.Cells(2, 3).Value + 1
                ' ws4.Cells(i, 2).Font.Color = RGB(30, 230, 190)
        End If
        If ws4.Cells(i, 2).Value = 5 Then
            ws4.Cells(i, 2).Interior.Color = RGB(255, 140, 0)
            ws4.Cells(2, 3).Value = ws4.Cells(2, 3).Value + 1
                ' ws4.Cells(i, 2).Font.Color = RGB(30, 230, 190)
        End If
        If ws4.Cells(i, 2).Value = 4 Then
            ws4.Cells(i, 2).Interior.Color = RGB(255, 215, 0)
            ws4.Cells(2, 3).Value = ws4.Cells(2, 3).Value + 1
                ' ws4.Cells(i, 2).Font.Color = RGB(30, 230, 190)
        End If
    Next i
    If ws4.Cells(2, 3).Value >= 20 Then
            ws4.Cells(2, 3).Interior.Color = RGB(255, 127, 80)
                ' ws4.Cells(i, 2).Font.Color = RGB(30, 230, 190)
        End If
    Application.DisplayAlerts = False
    ws2.Delete: ws3.Delete
    Application.DisplayAlerts = True
    
End Sub
