Attribute VB_Name = "modInboundSummary"
Option Explicit

Public Sub 입고관리()

    Dim ws As Worksheet, resWs As Worksheet
    Dim wb As Workbook
    Dim factoryDicts As Object, dict As Object
    Dim factories As Variant, fac As Variant
    Dim lastRow As Long, i As Long, rowIdx As Long
    Dim customerCode As String, factory As String
    Dim itemCode As String, itemName As String
    Dim key As String, k As Variant, data As Variant
    Dim qty As Double, totalQty As Double, grandTotal As Double
    Dim detailCount As Long, pass As Long
    Dim resultName As String
    Dim sh As Object

    On Error GoTo Failed

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "입고 원본 시트를 선택해주세요.", vbExclamation
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set wb = ws.Parent

    ' 원본 열 확인
    If Trim(CStr(ws.Cells(1, "B").Value2)) <> "수주처" Or _
       Trim(CStr(ws.Cells(1, "AQ").Value2)) <> "수주처 품목코드" Or _
       Trim(CStr(ws.Cells(1, "K").Value2)) <> "입고수량" Then

        MsgBox "입고 원본의 열을 확인해주세요." & vbCrLf & _
               "B열: 수주처" & vbCrLf & _
               "AQ열: 수주처 품목코드" & vbCrLf & _
               "G열: 품목명" & vbCrLf & _
               "K열: 입고수량", vbExclamation
        Exit Sub
    End If

    factories = Array("F1", "F2", "F5")

    ' 기존 결과 시트 보호
    For Each fac In factories

        resultName = "입고_" & CStr(fac)

        For Each sh In wb.Sheets
            If StrComp(sh.Name, resultName, vbTextCompare) = 0 Then
                MsgBox resultName & " 시트가 이미 있습니다." & vbCrLf & _
                       "기존 시트 이름을 변경한 후 다시 실행해주세요.", _
                       vbExclamation
                Exit Sub
            End If
        Next sh

    Next fac

    Set factoryDicts = CreateObject("Scripting.Dictionary")

    For Each fac In factories
        Set dict = CreateObject("Scripting.Dictionary")
        factoryDicts.Add CStr(fac), dict
    Next fac

    lastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "AQ").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "G").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "K").End(xlUp).Row)

    ' 전체 데이터 검증 및 집계
    ' 검증이 끝나기 전에는 결과 시트를 만들지 않음
    For i = 2 To lastRow

        If IsError(ws.Cells(i, "B").Value2) Or _
           IsError(ws.Cells(i, "AQ").Value2) Or _
           IsError(ws.Cells(i, "G").Value2) Then

            Err.Raise vbObjectError + 710, , _
                "수주처·품목코드·품목명에 오류 셀이 있습니다." & _
                vbCrLf & "원본 행: " & i
        End If

        customerCode = Trim(CStr(ws.Cells(i, "B").Value2))
        itemCode = Trim(CStr(ws.Cells(i, "AQ").Value2))
        itemName = Trim(CStr(ws.Cells(i, "G").Value2))

        ' 원본 합계행 제외
        If customerCode = "" And itemCode = "" Then
            Select Case UCase(itemName)
                Case "합계", "총 합계", "총합계", "계", "TOTAL"
                    GoTo NextInboundRow
            End Select
        End If

        ' 완전히 빈 행 제외
        If customerCode = "" And itemCode = "" And itemName = "" Then
            If Not IsError(ws.Cells(i, "K").Value2) Then
                If Len(Trim(CStr(ws.Cells(i, "K").Value2))) = 0 Then
                    GoTo NextInboundRow
                End If
            End If
        End If

        If IsError(ws.Cells(i, "K").Value2) Then
            Err.Raise vbObjectError + 711, , _
                "입고수량에 오류 셀이 있습니다." & _
                vbCrLf & "원본 행: " & i
        End If

        If Not IsNumeric(ws.Cells(i, "K").Value2) Then
            Err.Raise vbObjectError + 712, , _
                "입고수량이 숫자가 아닙니다." & _
                vbCrLf & "원본 행: " & i
        End If

        If itemCode = "" And itemName = "" Then
            Err.Raise vbObjectError + 713, , _
                "수량은 있지만 품목코드와 품목명이 없습니다." & _
                vbCrLf & "원본 행: " & i
        End If

        ' B열 수주처로 공장 구분
        Select Case UCase(customerCode)
            Case "1010", "F1"
                factory = "F1"
            Case "1060", "F2"
                factory = "F2"
            Case "1150", "F5"
                factory = "F5"
            Case Else
                Err.Raise vbObjectError + 714, , _
                    "공장을 구분할 수 없는 수주처입니다." & vbCrLf & _
                    "원본 행: " & i & vbCrLf & _
                    "B열 값: [" & customerCode & "]" & vbCrLf & _
                    "1010 / 1060 / 1150인지 확인해주세요." & vbCrLf & _
                    "결과 시트는 생성하지 않았습니다."
        End Select

        qty = CDbl(ws.Cells(i, "K").Value2)
        Set dict = factoryDicts(factory)

        ' 공장별 딕셔너리 안에서 품목코드 + 품목명으로 집계
        key = Len(itemCode) & ":" & itemCode & itemName

        If dict.Exists(key) Then
            data = dict(key)
            data(2) = data(2) + qty
            dict(key) = data
        Else
            dict.Add key, Array(itemCode, itemName, qty)
        End If

        grandTotal = grandTotal + qty
        detailCount = detailCount + 1

NextInboundRow:
    Next i

    If detailCount = 0 Then
        MsgBox "집계할 입고 상세 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    ' F1 / F2 / F5 결과 시트 생성
    For Each fac In factories

        Set dict = factoryDicts(CStr(fac))
        totalQty = 0

        Set resWs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        resWs.Name = "입고_" & CStr(fac)

        resWs.Range("A1:D1").Value = _
            Array("FAC", "수주처 품목코드", "품목명", "입고수량")

        ' 품목코드 앞자리 0 보존
        resWs.Columns("B:C").NumberFormat = "@"
        rowIdx = 2

        ' 코드 없는 항목 먼저, 코드 있는 항목 다음 출력
        For pass = 1 To 2
            For Each k In dict.Keys

                data = dict(k)

                If (pass = 1 And data(0) = "") Or _
                   (pass = 2 And data(0) <> "") Then

                    resWs.Cells(rowIdx, 1).Value2 = CStr(fac)
                    resWs.Cells(rowIdx, 2).Value2 = data(0)
                    resWs.Cells(rowIdx, 3).Value2 = data(1)
                    resWs.Cells(rowIdx, 4).Value2 = data(2)

                    totalQty = totalQty + data(2)
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

    Next fac

    MsgBox "공장별 입고 정리가 완료되었습니다." & vbCrLf & _
           "생성 시트: 입고_F1 / 입고_F2 / 입고_F5" & vbCrLf & _
           "전체 입고수량: " & Format(grandTotal, "#,##0"), _
           vbInformation
    Exit Sub

Failed:
    MsgBox Err.Description, vbExclamation

End Sub
