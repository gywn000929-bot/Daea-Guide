Attribute VB_Name = "modKetInventoryCompare"
Sub KET_이전_이후()
    Dim wsPrev As Worksheet, wsAfter As Worksheet, wsRes As Worksheet
    Dim lastRowA As Long, lastRowR As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim itemName As String, newQty As Variant
    Dim found As Boolean
    Const FLAGCOL As Long = 100   ' 신규 표시용 임시 보조열(CV)
    Const KEEPCOL As Long = 101   ' 이후에 존재 표시용 임시 보조열(CW)

    ' 1. 시트 확인
    On Error Resume Next
    Set wsPrev = Sheets("이전")
    Set wsAfter = Sheets("이후")
    On Error GoTo 0
    If wsPrev Is Nothing Or wsAfter Is Nothing Then
        MsgBox "'이전'과 '이후' 시트가 필요합니다.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 2. 결과 시트 생성 ('이전' 구조 복제)
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("업데이트결과").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    wsPrev.Copy After:=Sheets(Sheets.Count)
    Set wsRes = ActiveSheet
    wsRes.Name = "업데이트결과"

    lastCol = wsRes.Cells(1, wsRes.Columns.Count).End(xlToLeft).Column
    lastRowA = wsAfter.Cells(wsAfter.Rows.Count, "A").End(xlUp).Row

    ' 3. 이후 시트 순회 → 같은 품목은 이후 값으로 교체(KEEP 표시), 신규는 추가
    For i = 2 To lastRowA
        itemName = Trim(CStr(wsAfter.Cells(i, 1).value))
        If itemName = "" Then GoTo NextItem
        newQty = wsAfter.Cells(i, 2).value
        If Not IsNumeric(newQty) Then newQty = 0

        found = False
        lastRowR = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
        For j = 2 To lastRowR
            If Trim(CStr(wsRes.Cells(j, 1).value)) = itemName Then
                wsRes.Cells(j, 2).value = newQty       ' 이후 값으로 덮어쓰기
                wsRes.Cells(j, KEEPCOL).value = 1       ' 유지 표시
                found = True
                Exit For
            End If
        Next j

        If Not found Then
            lastRowR = lastRowR + 1
            wsRes.Cells(lastRowR, 1).value = itemName
            wsRes.Cells(lastRowR, 2).value = newQty
            wsRes.Cells(lastRowR, FLAGCOL).value = 1    ' 신규 플래그
            wsRes.Cells(lastRowR, KEEPCOL).value = 1    ' 유지 표시
        End If
NextItem:
    Next i

    ' 4. 이후에 없는 품목(KEEP 비어있음) 행 삭제 → 전체합 = 이후 전체합
    lastRowR = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    For j = lastRowR To 2 Step -1
        If wsRes.Cells(j, 1).value <> "" And wsRes.Cells(j, KEEPCOL).value <> 1 Then
            wsRes.Rows(j).Delete
        End If
    Next j

    ' 5. 품목명(A열) 오름차순 정렬 ? 신규가 중간에 끼워짐
    lastRowR = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    With wsRes.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsRes.Range("A2:A" & lastRowR), Order:=xlAscending
        .SetRange wsRes.Range(wsRes.Cells(1, 1), wsRes.Cells(lastRowR, KEEPCOL))
        .header = xlYes
        .Apply
    End With

    ' 6. 정렬 후 신규 품목 파란색(#0070C0) 표시 + 보조열 제거
    For j = 2 To lastRowR
        If wsRes.Cells(j, FLAGCOL).value = 1 Then
            wsRes.Range(wsRes.Cells(j, 1), wsRes.Cells(j, lastCol)).Font.Color = RGB(0, 112, 192)
        End If
    Next j
    wsRes.Columns(KEEPCOL).Delete
    wsRes.Columns(FLAGCOL).Delete

    wsRes.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "완료되었습니다." & vbCrLf & _
           "· 같은 품목: 이후 값으로 교체" & vbCrLf & _
           "· 신규 품목: 오름차순 위치에 삽입(파란색)" & vbCrLf & _
           "· 이후에 없는 품목: 제외 → 전체합 = 이후 전체합", vbInformation
End Sub

