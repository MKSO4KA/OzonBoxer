Sub CreateAndFormatSheets()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet
    Dim lastRow As Long, maxValue As Long, i As Long
    Dim data As Variant, resultData() As Variant
    Dim countDict As Object

    Application.DisplayAlerts = False
    ' ???????? ???? ??????, ????? ???????
    If ThisWorkbook.Sheets.count >= 2 Then
        For i = ThisWorkbook.Sheets.count To 2 Step -1
            ThisWorkbook.Sheets(i).Delete
        Next i
    End If
    Application.DisplayAlerts = True

    Set ws1 = ThisWorkbook.Sheets(1)
    Set ws2 = ThisWorkbook.Sheets.Add(After:=ws1)
    ws2.Name = "Sheet2"

    lastRow = ws1.Cells(ws1.Rows.count, "L").End(xlUp).Row
    data = ws1.Range("L1:L" & lastRow).value

    ' ?????????? ??????? ??? ???????? ????????
    Set countDict = CreateObject("Scripting.Dictionary")

    ' ????????? ?????? ? ?????????? ???????
    For i = 1 To lastRow
        If InStr(data(i, 1), "-") > 0 And Not data(i, 1) Like "*[a-zA-Z]*" Then
            Dim splitData As Variant
            splitData = Split(data(i, 1), "-")
            Dim key As Long: key = CLng(Trim(splitData(0)))
            Dim value As Long: value = CLng(Trim(splitData(1)))

            If key < 700 Then
                If Not countDict.Exists(key) Then
                    countDict(key) = value
                Else
                    countDict(key) = Application.WorksheetFunction.Max(countDict(key), value)
                End If
                maxValue = Application.WorksheetFunction.Max(maxValue, key)
            End If
        End If
    Next i

    ' ?????????? Sheet3
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ws2)
    ws3.Name = "Sheet3"
    
    ReDim resultData(1 To maxValue, 1 To 2)

    For i = 1 To maxValue
        resultData(i, 1) = i - 1
        If countDict.Exists(i - 1) Then
            resultData(i, 2) = countDict(i - 1)
        Else
            resultData(i, 2) = 0
        End If
    Next i

    ws3.Range("A1:B" & maxValue).value = resultData

    ' ?????????
    ws3.Range("A1").value = "Cells"
    ws3.Range("B1").value = "Count"

    ' ???????? Summary Sheet
    Set ws4 = ThisWorkbook.Sheets.Add(After:=ws3)
    ws4.Name = "Summary"
    
    ws3.UsedRange.Copy
    ws4.Range("A1").PasteSpecial Paste:=xlPasteValues

    lastRow = ws4.Cells(ws4.Rows.count, 2).End(xlUp).Row
    ws4.Range("A1:B" & lastRow).Sort Key1:=ws4.Range("B1"), Order1:=xlDescending, Header:=xlYes

    ' ??????? ??????? ? ??????????? ?????
    Dim boxCount As Long: boxCount = 0
    For i = 2 To lastRow
        Select Case ws4.Cells(i, 2).value
            Case Is >= 6
                ws4.Cells(i, 2).Interior.Color = RGB(220, 20, 60)
                boxCount = boxCount + 1
            Case 5
                ws4.Cells(i, 2).Interior.Color = RGB(255, 140, 0)
                boxCount = boxCount + 1
            Case 4
                ws4.Cells(i, 2).Interior.Color = RGB(255, 215, 0)
                boxCount = boxCount + 1
        End Select
    Next i

    ws4.Cells(2, 3).value = boxCount
    ws4.Cells(1, 3).value = "Boxes Approximately"

    If boxCount >= 20 Then
        ws4.Cells(2, 3).Interior.Color = RGB(255, 127, 80)
    End If

    ' ???????? ????????? ??????
    Application.DisplayAlerts = False
    ws2.Delete: ws3.Delete
    Application.DisplayAlerts = True

End Sub

