Attribute VB_Name = "modPurchaseOutstanding"
Option Explicit

Public Sub 발주미납_집계()

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

    Dim headerRow As Long, cols As Variant
    On Error GoTo Failed

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "발주미납 원본 시트를 선택해주세요.", vbExclamation
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set wb = ws.Parent

    If Not FindHeaders(ws, Array("수주번호", "수주처 품목코드", "품목명", "미납수량"), headerRow, cols) Then
        Err.Raise vbObjectError + 700, , "발주 미납 원본 시트를 선택하세요. 필요한 머리글: 수주번호/수주처 품목코드/품목명/미납수량"
    End If

    lastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, cols(2)).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, cols(1)).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, cols(0)).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, cols(3)).End(xlUp).Row)

    If lastRow < 2 Then
        MsgBox "집계할 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    For i = headerRow + 1 To lastRow

        If IsError(ws.Cells(i, cols(0)).Value2) Or _
           IsError(ws.Cells(i, cols(1)).Value2) Or _
           IsError(ws.Cells(i, cols(2)).Value2) Then

            Err.Raise vbObjectError + 701, , _
                "공장·품목코드·품목명에 오류 셀이 있습니다. 행: " & i
        End If

        customerCode = Trim(CStr(ws.Cells(i, cols(0)).Value2))
        itemCode = Trim(CStr(ws.Cells(i, cols(1)).Value2))
        itemName = Trim(CStr(ws.Cells(i, cols(2)).Value2))

        If customerCode = "" And itemCode = "" And itemName = "" Then
            If InStr(1, CStr(ws.Cells(i, 4).Value2), "합계", vbTextCompare) > 0 Then GoTo NextRow
        End If
        ' 원본 마지막 합계행 제외
        If customerCode = "" And itemCode = "" Then
            Select Case UCase(itemName)
                Case "합계", "총 합계", "총합계", "계", "TOTAL"
                    GoTo NextRow
            End Select
        End If

        ' 완전히 빈 행 제외
        If customerCode = "" And itemCode = "" And itemName = "" Then
            If Not IsError(ws.Cells(i, cols(3)).Value2) Then
                If Len(Trim(CStr(ws.Cells(i, cols(3)).Value2))) = 0 Then
                    GoTo NextRow
                End If
            End If
        End If

        If IsError(ws.Cells(i, cols(3)).Value2) Then
            Err.Raise vbObjectError + 702, , _
                "미납수량에 오류 셀이 있습니다. 행: " & i
        End If

        If Not IsNumeric(ws.Cells(i, cols(3)).Value2) Then
            Err.Raise vbObjectError + 703, , _
                "미납수량이 숫자가 아닙니다. 행: " & i
        End If

        If itemCode = "" And itemName = "" Then
            Err.Raise vbObjectError + 704, , _
                "수량은 있지만 품목코드와 품목명이 없습니다. 행: " & i
        End If

        ' 수주번호 앞자리 F1/F2/F5로 공장 구분
        Select Case UCase$(Left$(customerCode, 2))
            Case "F1": factory = "F1"
            Case "F2": factory = "F2"
            Case "F5": factory = "F5"
            Case Else
                Err.Raise vbObjectError + 705, , "수주번호에서 공장을 구분할 수 없습니다. 행: " & i & " / " & customerCode
        End Select

        qty = CDbl(ws.Cells(i, cols(3)).Value2)

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
    baseName = "발주미납_통합_" & Format(Now, "hhmmss")
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
        Array("FAC", "수주처 품목코드", "품목명", "미납잔량")

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

    If rowIdx > 3 Then
        resWs.Range("A1:D" & rowIdx - 1).Sort Key1:=resWs.Range("A2"), Order1:=xlAscending, Key2:=resWs.Range("B2"), Order2:=xlAscending, Key3:=resWs.Range("C2"), Order3:=xlAscending, header:=xlYes
    End If
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

    MsgBox "발주미납 통합집계 완료" & vbCrLf & _
           "총 미납수량: " & Format(totalQty, "#,##0"), vbInformation
    Exit Sub

Failed:
    MsgBox Err.Description, vbExclamation

End Sub

Private Function HKey(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(160), "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, "△", "")
    s = Replace(s, "▲", "")
    HKey = s
End Function
Private Function FindHeaders(ByVal ws As Worksheet, ByVal labels As Variant, ByRef hr As Long, ByRef cols As Variant) As Boolean
    Dim r As Long, c As Long, j As Long, n As Long, hits As Long, found() As Long
    n = UBound(labels) - LBound(labels) + 1
    For r = 1 To 10
        ReDim found(0 To n - 1)
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            For j = 0 To n - 1
                If HKey(ws.Cells(r, c).Value2) = HKey(labels(j)) Then found(j) = c
            Next j
        Next c
        hits = 0
        For j = 0 To n - 1
            If found(j) > 0 Then hits = hits + 1
        Next j
        If hits = n Then
            hr = r: cols = found: FindHeaders = True: Exit Function
        End If
    Next r
End Function
Private Function NewSheetName(ByVal wb As Workbook, ByVal base As String) As String
    Dim s As String, i As Long, sh As Object, exists As Boolean
    s = base
    Do
        exists = False
        For Each sh In wb.Sheets
            If StrComp(sh.Name, s, vbTextCompare) = 0 Then exists = True: Exit For
        Next sh
        If Not exists Then NewSheetName = s: Exit Function
        i = i + 1: s = base & "_" & i
    Loop
End Function
