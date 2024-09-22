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
    Dim ind As Long: ind = 0
    ' ????????? ?????? ? ?????????? ???????
    For i = 1 To lastRow
        If InStr(data(i, 1), "-") > 0 And Not data(i, 1) Like "*[a-zA-Zа-яА-Я]*" Then
            Dim splitData As Variant
            splitData = Split(data(i, 1), "-")
            Dim key As Long: key = CLng(Trim(splitData(0)))
            Dim value As Long: value = CLng(Trim(splitData(1)))
            
            If key < 700 Then
                If Not countDict.Exists(key) Then
                    countDict(key) = 1
                Else
                    countDict(key) = countDict(key) + 1
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
Sub DeleteLists()
    Application.DisplayAlerts = False
    If ThisWorkbook.Sheets.count > 2 Then
        For i = ThisWorkbook.Sheets.count To 3 Step -1
            ThisWorkbook.Sheets(i).Delete
        Next i
    End If
    Application.DisplayAlerts = True
End Sub
Sub SearchElement()
    Dim arr As Variant
    Dim SearchElement As Variant
    Dim foundCell As Range
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, c As Long
    Dim dataRange As Range
    Dim dataArr As Variant

    ' Set the search element (e.g., 44)
    SearchElement = InputBox("Enter the element to search for:", "Search Element")

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SearchElement).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Set the worksheet to search (e.g., "Sheet1")
    Set ws = ThisWorkbook.Worksheets(2)

    ' Set the range to search (e.g., column A)
    Set foundCell = ws.Range("A:A").Find(what:=SearchElement, lookat:=xlWhole)

    ' Check if the element was found
    If foundCell Is Nothing Then
        MsgBox "Element " & SearchElement & " not found"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Worksheets(2))
    ws.Name = SearchElement

    lastRow = ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).Rows.count, 12).End(xlUp).Row
    Set dataRange = ThisWorkbook.Sheets(1).Range(ThisWorkbook.Sheets(1).Cells(1, 12), ThisWorkbook.Sheets(1).Cells(lastRow, 12))
    
    ' Load data into an array for faster processing
    dataArr = dataRange.value
    c = 0

    For i = 1 To UBound(dataArr, 1)
        ' Check if the cell value contains the search value
        If dataArr(i, 1) Like SearchElement & "-*" Then
            c = c + 1
            arr = Split(dataArr(i, 1), "-")
            With ws.Cells(arr(1) + 1, 1)
                .NumberFormat = "@"
                .value = dataArr(i, 1)
            End With
            With ws.Cells(arr(1) + 1, 2)
                .value = ThisWorkbook.Sheets(1).Cells(i, 7).value ' Adjusted to -5 offset
                .NumberFormat = "_-*  #,##0.00 ???"
            End With
            With ws.Cells(arr(1) + 1, 3)
                .value = ThisWorkbook.Sheets(1).Cells(i, 5).value ' Adjusted to -7 offset
                .NumberFormat = "00000000-0000"
            End With
            With ws.Cells(arr(1) + 1, 4)
                .NumberFormat = "@"
                .value = Split(ThisWorkbook.Sheets(1).Cells(i, 9).value, " ")(0) ' Adjusted to -3 offset
            End With
            With ws.Cells(arr(1) + 1, 5)
                .NumberFormat = "@"
                .value = "Отправлен"
                .Interior.Color = RGB(255, 127, 80)

                If Not ThisWorkbook.Sheets(1).Cells(i, 13) Like "*[0-9]*" Then
                    .Interior.Color = RGB(255, 255, 0)
                    .value = Split(ThisWorkbook.Sheets(1).Cells(i, 3).value, " ")(0)
                ElseIf ThisWorkbook.Sheets(1).Cells(i, 13) Like "*[0-9]*" And Not ThisWorkbook.Sheets(1).Cells(i, 11) Like "*[0-9]*" Then
                    .Interior.Color = RGB(220, 20, 60)
                    .value = Split(ThisWorkbook.Sheets(1).Cells(i, 3).value, " ")(0)
                End If
            End With
            
            If foundCell.Offset(0, 1) = c Then Exit For
        End If
    Next i

    ' Remove empty rows in reverse order
    For i = ws.Cells(ws.Rows.count, 4).End(xlUp).Row To 2 Step -1
        If ws.Cells(i, 1).value = "" Then ws.Rows(i).Delete Shift:=xlUp
    Next i

    ' Set headers and autofit columns
    ws.Cells(1, 1).Resize(1, 5).value = Array("Item", "Cost", "User Identifier", "Payment Method", "Status")
    ws.Cells(1, 6).value = c
    ws.Columns("B:E").AutoFit
End Sub


