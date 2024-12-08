Attribute VB_Name = "OzonHelper"
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
    ws2.name = "Sheet2"

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
    ws3.name = "Sheet3"
    maxValue = maxValue + 1
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
    ws4.name = "Summary"
    
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
Function IsSearchElementValid(SearchElement As Variant) As Boolean
    If Len(SearchElement) < 8 Then
        IsSearchElementValid = True
    Else
        IsSearchElementValid = False
    End If
End Function
Function InsertCodeIntoSheet(name As String)
    Dim ws As Worksheet
    Dim code As String
    Dim vbComp As Object

    ' ??????? ??? ?????, ???? ?? ?????? ???????? ???
    Set ws = ThisWorkbook.Sheets(name) ' ???????? ?? ??? ?????? ?????

    ' ???, ??????? ????? ???????? ? ?????? ?????
    code = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & _
           "    Dim chromePath As String" & vbCrLf & _
           "    Dim url As String" & vbCrLf & vbCrLf & _
           "    ' ??????? ???? ? ???????????? ????? Google Chrome" & vbCrLf & _
           "    chromePath = ""C:\Program Files\Google\Chrome\Application\chrome.exe""" & vbCrLf & vbCrLf & _
           "    ' ?????????, ??????? ?? ?????? ? ??????? ?" & vbCrLf & _
           "    If Not Intersect(Target, Me.Columns(""H"")) Is Nothing Then" & vbCrLf & _
           "        ' ???????? URL ?? ????????? ??????" & vbCrLf & _
           "        url = Target.Value" & vbCrLf & vbCrLf & _
           "        ' ????????? URL ? Google Chrome" & vbCrLf & _
           "        Shell """" & chromePath & """" & "" --new-tab """" & url & """" , vbNormalFocus" & vbCrLf & _
           "    End If" & vbCrLf & _
           "End Sub"

    ' ???????? ????????? VBA ??? ?????
    Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.CodeName)

    ' ????????? ??? ? ?????? ?????
    vbComp.CodeModule.InsertLines 1, code
End Function
Sub AddButtonInCell(cell As Range, ws As Worksheet)
    Dim btn As Object
    Dim value As String
    value = "123"
    btnName = "Button_" & CStr(cell.Row)
    ' ??????? ??????, ???? ??? ??? ??????????
    On Error Resume Next
    ws.Shapes("").Delete ' ?????????? ??????????? ??? ??????
    On Error GoTo 0
    
    ' ????????? ????? ??????
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, cell.Left, cell.Top, cell.Width, cell.Height)
    With btn
        .name = btnName
        .OnAction = "RunMyMacro"
    End With
End Sub

Sub RunMyMacro()
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    Dim ws As Worksheet
    Dim chromePath As String
    chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    With regEx
        .Global = True
        .IgnoreCase = True
        .Pattern = "[^0-9]"
    End With
    Set ws = ThisWorkbook.Sheets(ActiveSheet.name)
    url = ws.Cells(regEx.Replace(CStr(Application.Caller), ""), 15)
    'MsgBox "Button Pressed. Value: " & url
    Shell """" & chromePath & """ --new-tab """ & url & """", vbNormalFocus
End Sub
Function FindNumber(SearchElement As Variant, ListName As String, Optional ByVal c As Long = 1) As Variant
    Dim arr As Variant
    Dim foundCell As Range
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dataRange As Range
    Dim dataArr As Variant
    Set ws = ThisWorkbook.Worksheets(2)
    Dim url As String
    url = "https://turbo-pvz.ozon.ru/search?filter=%7B%22search%22%3A%22"

    ' Set the range to search (e.g., column A)
    Set foundCell = ws.Range("A:A").Find(what:=SearchElement, lookat:=xlWhole)
    
    ' Check if the element was found
    If foundCell Is Nothing Then
        If Len(SearchElement) < 6 Then
            MsgBox "Element " & SearchElement & " not found"
            Exit Function
        End If
    End If
    Set ws = ThisWorkbook.Worksheets(CStr(ListName))
    

    lastRow = ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).Rows.count, 12).End(xlUp).Row
    Set dataRange = ThisWorkbook.Sheets(1).Range(ThisWorkbook.Sheets(1).Cells(1, 12), ThisWorkbook.Sheets(1).Cells(lastRow, 12))
    
    ' Load data into an array for faster processing
    dataArr = dataRange.value

    For i = 1 To UBound(dataArr, 1)
        ' Check if the cell value contains the search value
        If dataArr(i, 1) Like SearchElement & "-*" Then
            arr = Split(dataArr(i, 1), "-")
            With ws.Cells(c + 1, 1)
                .NumberFormat = "@"
                .value = dataArr(i, 1)
            End With
            With ws.Cells(c + 1, 2)
                .value = ThisWorkbook.Sheets(1).Cells(i, 7).value ' Adjusted to -5 offset
                .NumberFormat = "_-*  #,##0.00 ???"
            End With
            With ws.Cells(c + 1, 3)
                .value = ThisWorkbook.Sheets(1).Cells(i, 5).value ' Adjusted to -7 offset
                .NumberFormat = "00000000-0000"
            End With
            With ws.Cells(c + 1, 4)
                .NumberFormat = "@"
                .value = Split(ThisWorkbook.Sheets(1).Cells(i, 9).value, " ")(0) ' Adjusted to -3 offset
            End With
            With ws.Cells(c + 1, 5)
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
            
                
            
            Dim exampleString As String
            exampleString = ThisWorkbook.Sheets(1).Cells(i, 4).value ' ????????????, ??? i ??? ?????????
            
    ' ????????? ?????? ?? ????????
            splitValues = Split(exampleString, " ")
    
    ' ???????? ?????????? ?????????
            maxElements = UBound(splitValues) + 1
    
    ' ? ??????????? ?? ?????????? ????????? ????????? ????????
            Select Case maxElements
                Case 0
            ' ?????? ?? ??????
                Case 1
            ' ?????? ??, ??? ??????? ??? ?????? ????????
                    ws.Hyperlinks.Add Anchor:=ws.Cells(c + 1, 6), _
                      Address:="", _
                      SubAddress:="'" & ThisWorkbook.Sheets(1).name & "'!" & ThisWorkbook.Sheets(1).Cells(i, 4).Address, _
                      TextToDisplay:=splitValues(0)
                      url = url & Trim(splitValues(0)) & "%22" & "%7D"
                Case 2
            ' ?????????? ??? ???????? ????? ??????
                    result = splitValues(0) & " " & splitValues(1)
                    ws.Hyperlinks.Add Anchor:=ws.Cells(c + 1, 6), _
                      Address:="", _
                      SubAddress:="'" & ThisWorkbook.Sheets(1).name & "'!" & ThisWorkbook.Sheets(1).Cells(i, 4).Address, _
                      TextToDisplay:=result
                      url = url & Trim(splitValues(UBound(splitValues))) & "%22" & "%7D"
                Case Else
            ' ?????????? ??? ???????? ????? ??????
                    result = splitValues(0) & " " & splitValues(1) & " " & splitValues(2)
                    ws.Hyperlinks.Add Anchor:=ws.Cells(c + 1, 6), _
                      Address:="", _
                      SubAddress:="'" & ThisWorkbook.Sheets(1).name & "'!" & ThisWorkbook.Sheets(1).Cells(i, 4).Address, _
                      TextToDisplay:=result
                    url = url & Trim(splitValues(UBound(splitValues))) & "%22" & "%7D"
            End Select
            With ws.Cells(c + 1, 15)
                .NumberFormat = "@"
                .value = CStr(url)
                .Font.Color = RGB(255, 255, 255)
            End With
            ws.Columns("H").ColumnWidth = 20
            AddButtonInCell ws.Cells(c + 1, 8), ws
            
            
            c = c + 1
            
            
            
            'ws.Hyperlinks.Add Anchor:=ws.Cells(c, 7), _
            '    Address:=url, _
            '    TextToDisplay:="На сайт!"
            'InsertCodeIntoSheet (CStr(ListName))
            If foundCell.Offset(0, 1) = c - 1 Then
            Exit For
            End If
            
        End If
        url = "https://turbo-pvz.ozon.ru/search?filter=%7B%22search%22%3A%22"
    Next i
    FindNumber = c
End Function
Function Trim(inputVar As Variant) As String
    Dim regEx As Object
    Dim inputString As String
    inputString = CStr(inputVar)
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Убираем пробелы
    inputString = Replace(inputString, " ", "")
    
    ' Настраиваем регулярное выражение для удаления недопустимых символов
    With regEx
        .Global = True
        .IgnoreCase = True
        ' Разрешаем только цифры и букву 'i' в нижнем регистре
        .Pattern = "[^0-9i-]"
    End With
    
    ' Заменяем все недопустимые символы на пустую строку
    Trim = regEx.Replace(inputString, "")
    
    ' Удаляем все заглавные 'I'
    Trim = Replace(Trim, "I", "")
End Function
Sub searchNumber(ByVal SearchElements As Variant, Optional ByVal ListName As Variant = "")
    Dim SearchElement, StartElement As Variant
    Dim i, c, d As Long
    Dim ws As Worksheet
    Dim count As Long
    c = 0
    count = 1
    If Len(SearchElements(0)) > 6 Then
        d = FindNumByUserId(SearchElements(0))
    Else
        d = SearchElements(UBound(SearchElements))
    End If
    
    ' Set the search element (e.g., 44)
    If ListName = "" Then
        
        If Len(SearchElements(0)) > Len(d) Then
            SearchElement = d
        Else
            SearchElement = SearchElements(0)
        End If
        ListName = SearchElement
        count = FindNumber(SearchElement, CStr(ListName))
    Else
        For i = SearchElements(0) To d
            
            SearchElement = SearchElements(c)
            c = c + 1
            count = FindNumber(SearchElement, CStr(ListName), count)
        Next i
    End If
    ' Set the worksheet to search (e.g., "Sheet1")
    
    Set ws = ThisWorkbook.Worksheets(CStr(ListName))
    ' Remove empty rows in reverse order
    'For i = ws.Cells(ws.Rows.count, 4).End(xlUp).Row To 2 Step -1
    '    If ws.Cells(i, 1).value = "" Then ws.Rows(i).Delete Shift:=xlUp
    'Next i
    
    ' Set headers and autofit columns
    ws.Cells(1, 1).Resize(1, 8).value = Array("Item", "Cost", "User Identifier", "Payment Method", "Status", "Name", SearchElement, count - 1)
    ws.Cells(1, 8).Interior.Color = RGB(50, 220, 110)
    ws.Cells(1, 8).Font.Size = 30
    ws.Cells(1, 7).Font.Size = 30
    ws.Cells(1, 7).Interior.Color = RGB(220, 220, 100)
    ws.Columns("B:G").AutoFit
    ws.Rows.RowHeight = 30
    ws.Cells.VerticalAlignment = xlVAlignCenter
    ws.Cells.HorizontalAlignment = xlHAlignCenter
    ws.Columns("H").ColumnWidth = 20
End Sub
Function FindNumByUserId(SearchElement As Variant) As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean
    Dim searchPattern As String
    Dim lValue As String
    Dim result As Variant
    
    ' Устанавливаем рабочий лист
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' Получаем последний заполненный ряд в столбце E
    lastRow = ws.Cells(ws.Rows.count, 5).End(xlUp).Row
    
    ' Создаем шаблон для поиска (с нулем и без)
    searchPattern = Left(SearchElement, InStr(SearchElement & "-", "-") - 1)
    
    ' Ищем значение в столбце E
    found = False
    For i = 1 To lastRow
        ' Проверяем наличие совпадения с нулем и без
        If ws.Cells(i, 5).value Like searchPattern & "*" Or _
           ws.Cells(i, 5).value Like "0" & searchPattern & "*" Then
            ' Если найдено совпадение, ищем значение в столбце L
            lValue = ws.Cells(i, 12).value ' Столбец L
            If InStr(lValue, "-") > 0 Then
                ' Извлекаем значение до дефиса
                result = Left(lValue, InStr(lValue, "-") - 1)
                found = True
                Exit For
            End If
        End If
    Next i
    
    ' Если не найдено значение, возвращаем сообщение
    If Not found Then
        FindNumByUserId = -1
    Else
        FindNumByUserId = result
    End If
End Function
Function CreateArrayFromStroke(Stroke As String) As Variant
    Dim start As Integer, finish As Integer
    Dim arr() As Integer
    
    ' ????????? ?????? ?? ?????? ? ?????
    start = CInt(Split(Stroke, "-")(0))
    finish = CInt(Split(Stroke, "-")(1))
    
    ' ???????? ?????? ????? ?? ?????? ?? ?????
    ReDim arr(finish - start)
    For i = 0 To UBound(arr)
        arr(i) = start + i
    Next i
    
    CreateArrayFromStroke = arr
End Function
Sub SearchElement2()
    DeleteLists
    Dim Element, result As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim maxValue As Long
    maxValue = Application.WorksheetFunction.Max(ThisWorkbook.Worksheets(2).Range("A2:A" & ThisWorkbook.Worksheets(2).Cells(ThisWorkbook.Worksheets(2).Rows.count, "A").End(xlUp).Row))
    For i = 1 To maxValue - (maxValue Mod 7) Step 7
        Element = i & "-" & i + 6
    If IsSearchElementValid(Element) Then
        
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Worksheets(Element).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Worksheets(2))
        ws.name = Element
        If Element Like "*-*" Then
            searchNumber CreateArrayFromStroke(CStr(Element)), CStr(Element)
        Else
            searchNumber (Array(Element))
        End If
        
        ws.Cells(1, 7).NumberFormat = "@"
        ws.Cells(1, 7) = CStr(Element)
        ws.Columns("B:G").AutoFit
        ws.Columns("H").ColumnWidth = 20
    Else
        result = FindNumByUserId(Element)
        If result = -1 Then
            MsgBox "User on ID " & Element & " not found"
            End
        End If
        
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Worksheets(Element).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Worksheets(2))
        ws.name = Element
        searchNumber (Array(Element))
    End If
    Next i
End Sub
Sub SortSheetsByFirstNumberInNameSkippingFirstTwo()
    Dim ws As Worksheet
    Dim sheetNames As Collection
    Dim sortedNames() As String
    Dim i As Long
    Dim j As Long
    Dim temp As String
    Dim n As Long

    ' ??????? ????????? ??? ???????? ???? ??????
    Set sheetNames = New Collection
    
    ' ????????? ????? ?????? ? ?????????, ??????? ? ????????
    For i = 3 To ThisWorkbook.Sheets.count
        sheetNames.Add ThisWorkbook.Sheets(i).name
    Next i
    
    ' ???????? ?????????? ?????? ??? ??????????
    n = sheetNames.count
    
    ' ??????? ?????? ??? ???????? ??????????????? ????
    ReDim sortedNames(1 To n)
    
    ' ???????? ????? ? ??????
    For i = 1 To n
        sortedNames(i) = sheetNames(i)
    Next i
    
    ' ?????????? ????????? ?? ??????? ????? ? ????? ?????
    For i = 1 To n - 1
        For j = i + 1 To n
            If GetFirstNumber(sortedNames(i)) > GetFirstNumber(sortedNames(j)) Then
                ' ?????? ??????? ?????
                temp = sortedNames(i)
                sortedNames(i) = sortedNames(j)
                sortedNames(j) = temp
            End If
        Next j
    Next i
    
    ' ?????????? ????? ? ??????????????? ???????
    For i = 1 To n
        ThisWorkbook.Sheets(sortedNames(i)).Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    Next i
    ActivateSheetIfExists (1)
    ActivateSheetIfExists (3)
End Sub
Function ActivateSheetIfExists(sheetId As Integer)
    Dim ws As Worksheet
    
    ' ?????????, ?????????? ?? ?????? ????
    If ThisWorkbook.Sheets.count >= sheetId Then
        ' ????????????? ?????? ???? ??? ????????
        Set ws = ThisWorkbook.Sheets(sheetId)
        ws.Activate
    End If
End Function
Function GetFirstNumber(sheetName As String) As Double
    Dim numbers() As String
    Dim num As Double
    
    ' ????????? ??? ????? ?? ??????? "-" ? ????? ?????? ?????
    If InStr(sheetName, "-") > 0 Then
        numbers = Split(sheetName, "-")
        num = Val(numbers(0))
    Else
        num = Val(sheetName)
    End If
    
    GetFirstNumber = num
End Function
Sub SearchElement()
    Dim Element, result As Variant
    Dim ws As Worksheet
    Element = InputBox("Enter the element to search for:", "Search Element")
    If IsSearchElementValid(Element) Then
        
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Worksheets(Element).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Worksheets(2))
        ws.name = Element
        If Element Like "*-*" Then
            searchNumber CreateArrayFromStroke(CStr(Element)), CStr(Element)
        Else
            searchNumber (Array(Element))
        End If
        
        ws.Cells(1, 7).NumberFormat = "@"
        ws.Cells(1, 7) = CStr(Element)
        ws.Columns("B:G").AutoFit
        ws.Columns("H").ColumnWidth = 20
    Else
        result = FindNumByUserId(Element)
        If result = -1 Then
            MsgBox "User on ID " & Element & " not found"
            End
        End If
        
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Worksheets(result).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Worksheets(2))
        ws.name = result
        searchNumber (Array(result))
        ws.Cells(1, 7).NumberFormat = "@"
        ws.Cells(1, 7) = CStr(result)
        ws.Columns("B:G").AutoFit
        ws.Columns("H").ColumnWidth = 20
    End If
End Sub
