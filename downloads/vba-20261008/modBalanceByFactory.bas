Attribute VB_Name = "modBalanceByFactory"
Sub Balance()

    Dim ws As Worksheet, resWs As Worksheet
    Dim lastRow As Long, i As Long, rowIdx As Long
    Dim dict As Object, key As String
    Dim itemCode As String, itemName As String, qty As Double
    Dim custCode As Variant, custMapped As String
    Dim k As Variant, items As Variant
    Dim totalQty As Double

    Set ws = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    totalQty = 0

    ' 수주처 + 품목코드 + 품목명 기준 집계
    For i = 2 To lastRow

        custCode = ws.Cells(i, "B").Value
        itemCode = Trim(CStr(ws.Cells(i, "E").Value))
        itemName = Trim(CStr(ws.Cells(i, "G").Value))

        Select Case CStr(custCode)
            Case "1150"
                custMapped = "F5"
            Case "1010"
                custMapped = "F1"
            Case "1060"
                custMapped = "F2"
            Case "1050"
                custMapped = "삼도"
            Case Else
                custMapped = CStr(custCode)
        End Select

        If IsNumeric(ws.Cells(i, "N").Value) Then
            qty = CDbl(ws.Cells(i, "N").Value)
        Else
            qty = 0
        End If

        key = custMapped & "|" & itemCode & "|" & itemName

        If itemName <> "" Then
            If dict.Exists(key) Then
                dict(key) = dict(key) + qty
            Else
                dict.Add key, qty
            End If

            totalQty = totalQty + qty
        End If

    Next i

    ' 결과 시트 생성
    Set resWs = Sheets.Add
    resWs.Name = "미납수량합계_" & Format(Now, "hhmmss")

    resWs.Range("A1:D1").Value = _
        Array("수주처", "거래처 품목코드", "품목명", "미납수량 합계")

    resWs.Columns("B:C").NumberFormat = "@"
    rowIdx = 2

    ' 품목코드가 없는 항목 먼저 출력
    For Each k In dict.Keys

        items = Split(k, "|")

        If items(1) = "" Then
            resWs.Cells(rowIdx, 1).Value = items(0)
            resWs.Cells(rowIdx, 2).Value = ""
            resWs.Cells(rowIdx, 3).Value = items(2)
            resWs.Cells(rowIdx, 4).Value = dict(k)
            rowIdx = rowIdx + 1
        End If

    Next k

    ' 품목코드가 있는 항목 출력
    For Each k In dict.Keys

        items = Split(k, "|")

        If items(1) <> "" Then
            resWs.Cells(rowIdx, 1).Value = items(0)
            resWs.Cells(rowIdx, 2).Value = items(1)
            resWs.Cells(rowIdx, 3).Value = items(2)
            resWs.Cells(rowIdx, 4).Value = dict(k)
            rowIdx = rowIdx + 1
        End If

    Next k

    ' 총합계
    resWs.Cells(rowIdx, 3).Value = "총 합계"
    resWs.Cells(rowIdx, 4).Value = totalQty

    ' 전체 표 서식
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

    ' 머리글
    With resWs.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
        .RowHeight = 24
    End With

    ' 합계행
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

    resWs.Calculate

    MsgBox "미납수량 발란스 정리가 완료되었습니다."

End Sub
