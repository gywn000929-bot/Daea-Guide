Attribute VB_Name = "modPurchaseOrderSummary"
Option Explicit

Public Sub 발주잔량_수주별_정리()
    Dim ws As Worksheet, wb As Workbook, outWs As Worksheet
    Dim labels As Variant, cols(0 To 4) As Long, hr As Long
    Dim r As Long, c As Long, j As Long, hits As Long, lastRow As Long, n As Long
    Dim po As String, shortPO As String, orderNo As String, item As String, key As String
    Dim dict As Object, poMap As Object, data As Variant, k As Variant
    Dim q As Double, bal As Double, totalQ As Double, totalBal As Double
    Dim baseName As String, outName As String, suffix As Long, sh As Object, exists As Boolean
    On Error GoTo Failed
    If TypeName(ActiveSheet) <> "Worksheet" Then Err.Raise vbObjectError + 801, , "발주잔량 원본 시트를 선택하세요."
    Set ws = ActiveSheet
    Set wb = ws.Parent
    labels = Array("발주번호", "수주번호", "품목명", "발주수량", "미납수량")
    For r = 1 To 10
        Erase cols
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            For j = 0 To 4
                If HeaderKey(ws.Cells(r, c).Value2) = labels(j) Then cols(j) = c
            Next j
        Next c
        hits = 0
        For j = 0 To 4
            If cols(j) > 0 Then hits = hits + 1
        Next j
        If hits = 5 Then hr = r: Exit For
    Next r
    If hr = 0 Then Err.Raise vbObjectError + 802, , "발주번호 / 수주번호 / 품목명 / 발주수량 / 미납수량 머리글을 찾지 못했습니다. 원본 시트에서 실행하세요."
    For j = 0 To 4
        n = ws.Cells(ws.Rows.Count, cols(j)).End(xlUp).Row
        If n > lastRow Then lastRow = n
    Next j
    Set dict = CreateObject("Scripting.Dictionary")
    Set poMap = CreateObject("Scripting.Dictionary")
    For r = hr + 1 To lastRow
        po = TextValue(ws.Cells(r, cols(0)).Value2, r)
        orderNo = TextValue(ws.Cells(r, cols(1)).Value2, r)
        item = TextValue(ws.Cells(r, cols(2)).Value2, r)
        If po = "" And orderNo = "" Then
            If IsTotalRow(ws, r) Then GoTo NextRow
            If item = "" And Len(TextValue(ws.Cells(r, cols(3)).Value2, r)) = 0 And Len(TextValue(ws.Cells(r, cols(4)).Value2, r)) = 0 Then GoTo NextRow
        End If
        If Len(po) < 5 Or item = "" Then Err.Raise vbObjectError + 803, , "발주번호 또는 품목명이 비어 있거나 발주번호가 5자리 미만입니다. 원본 행: " & r
        shortPO = Right$(po, 5)
        If poMap.Exists(shortPO) Then
            If poMap(shortPO) <> po Then Err.Raise vbObjectError + 804, , "서로 다른 발주번호의 끝 5자리가 같습니다: " & poMap(shortPO) & " / " & po & ". 원본을 확인해주세요."
        Else
            poMap.Add shortPO, po
        End If
        q = QuantityValue(ws.Cells(r, cols(3)).Value2, r)
        bal = QuantityValue(ws.Cells(r, cols(4)).Value2, r)
        key = Len(shortPO) & ":" & shortPO & Len(orderNo) & ":" & orderNo & Len(item) & ":" & item
        If dict.Exists(key) Then
            data = dict(key)
            data(3) = data(3) + q
            data(4) = data(4) + bal
            dict(key) = data
        Else
            dict.Add key, Array(shortPO, orderNo, item, q, bal)
        End If
        totalQ = totalQ + q
        totalBal = totalBal + bal
NextRow:
    Next r
    If dict.Count = 0 Then Err.Raise vbObjectError + 805, , "집계할 상세 데이터가 없습니다."
    baseName = "발주잔량_수주별"
    outName = baseName
    Do
        exists = False
        For Each sh In wb.Sheets
            If sh.Name = outName Then exists = True: Exit For
        Next sh
        If Not exists Then Exit Do
        suffix = suffix + 1
        outName = baseName & "_" & suffix
    Loop
    Set outWs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    outWs.Name = outName
    outWs.Range("A1:E1").Value = Array("품목명", "수주번호", "주문서 번호", "발주수량", "미납수량")
    outWs.Columns("A:C").NumberFormat = "@"
    r = 2
    For Each k In dict.Keys
        data = dict(k)
        outWs.Cells(r, 1).Value2 = data(2)
        outWs.Cells(r, 2).Value2 = data(1)
        outWs.Cells(r, 3).Value2 = data(0)
        outWs.Cells(r, 4).Value2 = data(3)
        outWs.Cells(r, 5).Value2 = data(4)
        r = r + 1
    Next k
    If r > 3 Then outWs.Range("A1:E" & (r - 1)).Sort Key1:=outWs.Range("A2"), Order1:=xlAscending, Key2:=outWs.Range("B2"), Order2:=xlAscending, Key3:=outWs.Range("C2"), Order3:=xlAscending, Header:=xlYes
    outWs.Cells(r, 1).Value2 = "합계"
    outWs.Cells(r, 4).Value2 = totalQ
    outWs.Cells(r, 5).Value2 = totalBal
    With outWs.Range("A1:E" & r)
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Color = vbBlack
        .VerticalAlignment = xlCenter
        .WrapText = False
        .RowHeight = 23
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(190, 190, 190)
        .Borders.Weight = xlThin
    End With
    With outWs.Range("A1:E1")
        .Font.Bold = True
        .Interior.Pattern = xlNone
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With
    With outWs.Range("A" & r & ":E" & r)
        .Font.Bold = True
        .Interior.Color = 14524132
    End With
    outWs.Columns("A").ColumnWidth = 45
    outWs.Columns("B").ColumnWidth = 22
    outWs.Columns("C").ColumnWidth = 14
    outWs.Columns("D:E").ColumnWidth = 15
    outWs.Range("A2:B" & r).ShrinkToFit = True
    outWs.Range("D2:E" & r).NumberFormat = "#,##0.########;-#,##0.########;0"
    outWs.Range("A2:C" & (r - 1)).HorizontalAlignment = xlCenter
    outWs.Range("D2:E" & r).HorizontalAlignment = xlRight
    ' 같은 품목명이 연속된 행의 품목명 셀을 병합
    Dim groupStart As Long, groupEnd As Long, nextRow As Long, itemLabel As String
    groupStart = 2
    Do While groupStart < r
        itemLabel = CStr(outWs.Cells(groupStart, 1).Value2)
        groupEnd = groupStart
        Do While groupEnd + 1 < r
            If CStr(outWs.Cells(groupEnd + 1, 1).Value2) <> itemLabel Then Exit Do
            groupEnd = groupEnd + 1
        Loop
        If groupEnd > groupStart Then
            ' 동일한 값만 병합하므로 첫 셀 외의 중복 표시를 먼저 비움
            outWs.Range(outWs.Cells(groupStart + 1, 1), outWs.Cells(groupEnd, 1)).ClearContents
            With outWs.Range(outWs.Cells(groupStart, 1), outWs.Cells(groupEnd, 1))
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If
        groupStart = groupEnd + 1
    Loop
    With outWs.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .BlackAndWhite = True
        .PrintTitleRows = "$1:$1"
        .PrintArea = outWs.Range("A1:E" & r).Address
        .CenterHorizontally = True
        .CenterFooter = "&P / &N"
    End With
    MsgBox "발주잔량 정리 완료: " & outName & vbCrLf & "발주수량: " & Format(totalQ, "#,##0.########") & vbCrLf & "미납수량: " & Format(totalBal, "#,##0.########"), vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

Private Function HeaderKey(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    HeaderKey = Replace(Replace(Replace(Replace(Replace(Trim$(CStr(v)), " ", ""), vbCr, ""), vbLf, ""), vbTab, ""), ChrW(160), "")
    HeaderKey = Replace(Replace(HeaderKey, "△", ""), "▲", "")
    If HeaderKey = "주문서번호" Then HeaderKey = "발주번호"
End Function
Private Function TextValue(ByVal v As Variant, ByVal rowNo As Long) As String
    If IsError(v) Then Err.Raise vbObjectError + 806, , "오류 셀이 있습니다. 원본 행: " & rowNo
    TextValue = Trim$(CStr(v))
End Function
Private Function QuantityValue(ByVal v As Variant, ByVal rowNo As Long) As Double
    If IsError(v) Then Err.Raise vbObjectError + 807, , "수량 오류 셀이 있습니다. 원본 행: " & rowNo
    If Not IsNumeric(v) Then Err.Raise vbObjectError + 808, , "수량이 숫자가 아닙니다. 원본 행: " & rowNo
    QuantityValue = CDbl(v)
End Function
Private Function IsTotalRow(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim c As Long, s As String
    For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        s = HeaderKey(ws.Cells(r, c).Value2)
        Select Case UCase$(s)
            Case "합계", "총합계", "계", "TOTAL", "GRANDTOTAL"
                IsTotalRow = True
                Exit Function
        End Select
    Next c
End Function