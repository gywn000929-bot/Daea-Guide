Attribute VB_Name = "modShipmentSummary"
Option Explicit

Public Sub STOCK_통합집계()

    Dim ws As Worksheet, resWs As Worksheet
    Dim wb As Workbook
    Dim dict As Object
    Dim lastRow As Long, i As Long, rowIdx As Long
    Dim customerCode As String, factory As String
    Dim itemCode As String, itemName As String
    Dim qty As Double, totalQty As Double
    Dim key As String, k As Variant, data As Variant
    Dim pass As Long
    Dim baseName As String, sheetName As String
    Dim suffix As Long
    Dim sh As Object, nameExists As Boolean

    On Error GoTo Failed

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "STOCK 원본 시트를 선택해주세요.", vbExclamation
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set wb = ws.Parent

    ' 원본 열 확인
    If Trim(CStr(ws.Cells(1, "Y").Value2)) <> "출고처코드" Or _
       Trim(CStr(ws.Cells(1, "U").Value2)) <> "거래처 품목코드" Or _
       Trim(CStr(ws.Cells(1, "I").Value2)) <> "출고수량" Then

        MsgBox "원본 열을 확인해주세요." & vbCrLf & _
               "Y열: 출고처코드" & vbCrLf & _
               "U열: 거래처 품목코드" & vbCrLf & _
               "I열: 출고수량", vbExclamation
        Exit Sub
    End If

    lastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, "C").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "U").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "Y").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "I").End(xlUp).Row)

    If lastRow < 2 Then
        MsgBox "집계할 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow

        If IsError(ws.Cells(i, "Y").Value2) Or _
           IsError(ws.Cells(i, "U").Value2) Or _
           IsError(ws.Cells(i, "C").Value2) Then

            Err.Raise vbObjectError + 701, , _
                "공장·품목코드·품목명에 오류 셀이 있습니다. 행: " & i
        End If

        customerCode = Trim(CStr(ws.Cells(i, "Y").Value2))
        itemCode = Trim(CStr(ws.Cells(i, "U").Value2))
        itemName = Trim(CStr(ws.Cells(i, "C").Value2))

        ' 원본 마지막 합계행 제외
        If customerCode = "" And itemCode = "" Then
            Select Case UCase(itemName)
                Case "합계", "총 합계", "총합계", "계", "TOTAL"
                    GoTo NextRow
            End Select
        End If

        ' 완전히 빈 행 제외
        If customerCode = "" And itemCode = "" And itemName = "" Then
            If Not IsError(ws.Cells(i, "I").Value2) Then
                If Len(Trim(CStr(ws.Cells(i, "I").Value2))) = 0 Then
                    GoTo NextRow
                End If
            End If
        End If

        If IsError(ws.Cells(i, "I").Value2) Then
            Err.Raise vbObjectError + 702, , _
                "출고수량에 오류 셀이 있습니다. 행: " & i
        End If

        If Not IsNumeric(ws.Cells(i, "I").Value2) Then
            Err.Raise vbObjectError + 703, , _
                "출고수량이 숫자가 아닙니다. 행: " & i
        End If

        If itemCode = "" And itemName = "" Then
            Err.Raise vbObjectError + 704, , _
                "수량은 있지만 품목코드와 품목명이 없습니다. 행: " & i
        End If

        ' Y열 출고처코드로 공장 구분
        Select Case customerCode
            Case "1010"
                factory = "F1"
            Case "1060"
                factory = "F2"
            Case "1150"
                factory = "F5"
            Case Else
                If customerCode = "" Then
                    factory = "확인필요"
                Else
                    factory = "확인필요(" & customerCode & ")"
                End If
        End Select

        qty = CDbl(ws.Cells(i, "I").Value2)

        ' 공장 + 품목코드 + 품목명 기준 집계
        key = Len(factory) & ":" & factory & _
              Len(itemCode) & ":" & itemCode & itemName

        If dict.exists(key) Then
            data = dict(key)
            data(3) = data(3) + qty
            dict(key) = data
        Else
            dict.Add key, Array(factory, itemCode, itemName, qty)
        End If

        totalQty = totalQty + qty

NextRow:
    Next i

    If dict.Count = 0 Then
        MsgBox "집계할 상세 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    ' 결과 시트 이름 중복 방지
    baseName = "STOCK_통합_" & Format(Now, "hhmmss")
    sheetName = baseName

    Do
        nameExists = False

        For Each sh In wb.Sheets
            If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
                nameExists = True
                Exit For
            End If
        Next sh

        If Not nameExists Then Exit Do

        suffix = suffix + 1
        sheetName = baseName & "_" & suffix
    Loop

    Set resWs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    resWs.Name = sheetName

    resWs.Range("A1:D1").value = _
        Array("FAC", "ITEM CD", "Part Name", "Shipped Q'ty")

    resWs.Columns("B:C").NumberFormat = "@"
    rowIdx = 2

    ' 품목코드가 없는 항목 먼저 출력
    For pass = 1 To 2
        For Each k In dict.keys

            data = dict(k)

            If (pass = 1 And data(1) = "") Or _
               (pass = 2 And data(1) <> "") Then

                resWs.Cells(rowIdx, 1).Value2 = data(0)
                resWs.Cells(rowIdx, 2).Value2 = data(1)
                resWs.Cells(rowIdx, 3).Value2 = data(2)
                resWs.Cells(rowIdx, 4).Value2 = data(3)

                rowIdx = rowIdx + 1
            End If

        Next k
    Next pass

    resWs.Cells(rowIdx, 3).Value2 = "TOTAL"
    resWs.Cells(rowIdx, 4).Value2 = totalQty

    With resWs.Range("A1:D" & rowIdx)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Color = vbBlack
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(200, 200, 200)
        .RowHeight = 18
    End With

    With resWs.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
        .RowHeight = 24
    End With

    With resWs.Range("A" & rowIdx & ":D" & rowIdx)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    resWs.Range("D2:D" & rowIdx).NumberFormat = "#,##0"
    resWs.Columns("A").HorizontalAlignment = xlCenter
    resWs.Columns("A:D").AutoFit

    If resWs.Columns("A").ColumnWidth < 8 Then _
        resWs.Columns("A").ColumnWidth = 8

    If resWs.Columns("B").ColumnWidth < 20 Then _
        resWs.Columns("B").ColumnWidth = 20

    If resWs.Columns("C").ColumnWidth < 28 Then _
        resWs.Columns("C").ColumnWidth = 28

    If resWs.Columns("D").ColumnWidth < 16 Then _
        resWs.Columns("D").ColumnWidth = 16

    MsgBox "STOCK 통합집계 완료" & vbCrLf & _
           "총 출고수량: " & Format(totalQty, "#,##0"), vbInformation
    Exit Sub

Failed:
    MsgBox Err.Description, vbExclamation

End Sub
