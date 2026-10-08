Attribute VB_Name = "modAirFormat"
'============================================================
' Air 시트에 '선박 느낌' 서식 입히기
'   · 흰 헤더 + 검정 굵은 글씨 + 연한 회색 격자 + 핵심열 굵게
'   · FAC F2=노랑 / F5=파랑
'   · NEED ETD (URGENT·날짜) 색은 그대로 유지 (건드리지 않음)
'   · 에어데이터 정리 실행 후, Air 시트에서 실행하세요
'============================================================
Sub 에어_서식_정리()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim facCol As Long, brCol As Long, prCol As Long, lastCol As Long, lastRow As Long, i As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    facCol = FindHdrA_(ws, "FAC"): If facCol = 0 Then facCol = 1
    lastRow = ws.Cells(ws.Rows.Count, facCol).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    brCol = FindHdrA_(ws, "B-R")
    prCol = FindHdrA_(ws, "Purchase request")

    Application.ScreenUpdating = False

    ' 본문 글꼴
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Font
        .Name = "맑은 고딕": .Size = 10
    End With

    ' 격자 테두리 (연한 회색)
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(208, 214, 224)
        .Weight = xlThin
    End With

    ' 머리글 (흰 바탕 · 검정 굵은 글씨 · 가운데)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Name = "맑은 고딕": .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(1).RowHeight = 30

    ' 핵심 열 굵게 (B-R / Purchase request)
    If brCol > 0 Then ws.Range(ws.Cells(2, brCol), ws.Cells(lastRow, brCol)).Font.Bold = True
    If prCol > 0 Then ws.Range(ws.Cells(2, prCol), ws.Cells(lastRow, prCol)).Font.Bold = True

    ' FAC 공장별 색상 (F2=노랑 / F5=파랑) ? NEED ETD 열은 손대지 않음
    For i = 2 To lastRow
        Select Case UCase$(Trim$(CStr(ws.Cells(i, facCol).value)))
            Case "F2": ws.Cells(i, facCol).Interior.Color = RGB(255, 255, 0)
            Case "F5": ws.Cells(i, facCol).Interior.Color = RGB(0, 176, 240)
        End Select
    Next i

    Application.ScreenUpdating = True
    MsgBox "Air 시트 서식을 선박 느낌으로 맞췄습니다." & vbCrLf & _
           "(URGENT·날짜 색은 그대로 유지)", vbInformation
End Sub

'--- [함수] 헤더명으로 열 찾기 (공백/특수문자/대소문자 무시) ---
Private Function FindHdrA_(ws As Worksheet, nameWanted As String) As Long
    Dim i As Long, h As String, t As String, s As String
    t = UCase(Replace(Replace(Replace(Replace(CStr(nameWanted), ".", ""), " ", ""), "_", ""), "-", ""))
    For i = 1 To 60
        s = CStr(ws.Cells(1, i).value)
        h = UCase(Replace(Replace(Replace(Replace(s, ".", ""), " ", ""), "_", ""), "-", ""))
        If Len(h) > 0 Then
            If h = t Or h Like t & "*" Or t Like h & "*" Then FindHdrA_ = i: Exit Function
        End If
    Next i
    FindHdrA_ = 0
End Function

