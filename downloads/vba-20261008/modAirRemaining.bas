Attribute VB_Name = "modAirRemaining"
Option Explicit

' ============================================================
'  Air up 생성
'   - Remain(E열) = Purchase request - (기존 Air Stock)
'   - Remain <= 0 (기존 재고로 충족) 행은 삭제
'   - 남은 행의 STOCK 자리 = 그 공장 Sto up 값 (없으면 0)
' ============================================================
Sub Air_up_생성()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsAir As Worksheet, wsOut As Worksheet

    Set wsAir = FindSheet_(wb, "Air")
    If wsAir Is Nothing Then Set wsAir = ActiveSheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo EH

    ' 기존 Air up 삭제 후 Air 복사
    Dim ex As Worksheet
    Set ex = FindSheet_(wb, "Air up")
    If Not ex Is Nothing Then ex.Delete

    wsAir.Copy After:=wsAir
    Set wsOut = wb.Sheets(wsAir.Index + 1)
    wsOut.Name = "Air up"

    ' 헤더 열 찾기 (이름 기반)
    Dim facCol As Long, itemCol As Long, prCol As Long, stoCol As Long
    facCol = FindHdr_(wsOut, "FAC")
    itemCol = FindHdr_(wsOut, "Item cd")
    prCol = FindHdr_(wsOut, "Purchase request")
    stoCol = FindHdr_(wsOut, "STOCK")
    If facCol = 0 Or itemCol = 0 Or prCol = 0 Or stoCol = 0 Then
        MsgBox "필요한 헤더(FAC / Item cd / Purchase request / STOCK)를 찾지 못했습니다." & vbCrLf & _
               "Air 시트의 헤더명을 확인해주세요.", vbExclamation
        GoTo CleanExit
    End If

    ' F1 / F2 / F5 Sto up 사전 구성
    Dim d1 As Object, d2 As Object, d5 As Object
    Set d1 = BuildStoUpDict_(FindSheet_(wb, "F1 Sto up"))
    Set d2 = BuildStoUpDict_(FindSheet_(wb, "F2 Sto up"))
    Set d5 = BuildStoUpDict_(FindSheet_(wb, "F5 Sto up"))

    ' 처리 (아래→위로: 충족 행 삭제 안전)
    Dim lastRow As Long, r As Long
    Dim fac As String, code As String
    Dim pr As Double, oldStock As Double, stoUp As Double, remain As Double
    Dim delCnt As Long, keepCnt As Long
    lastRow = wsOut.Cells(wsOut.Rows.Count, facCol).End(xlUp).Row

    For r = lastRow To 2 Step -1
        If Len(Trim$(CStr(wsOut.Cells(r, facCol).value))) = 0 Then GoTo NextR

        ' 1) 기존 Air Stock 으로 Remain 계산
        pr = NzNum_(wsOut.Cells(r, prCol).value)
        oldStock = NzNum_(wsOut.Cells(r, stoCol).value)   ' 덮어쓰기 전에 먼저 읽음
        remain = pr - oldStock

        If remain <= 0 Then
            wsOut.Rows(r).Delete            ' 기존 재고로 다 충족 → 삭제
            delCnt = delCnt + 1
        Else
            ' 2) E열 = Remain(부족분)
            wsOut.Cells(r, prCol).value = remain

            ' 3) STOCK 자리 = 그 공장 Sto up (없으면 0)
            fac = NormFac_(wsOut.Cells(r, facCol).value)
            code = NormCode_(wsOut.Cells(r, itemCol).value)
            stoUp = 0
            Select Case fac
                Case "F1": If d1.exists(code) Then stoUp = d1(code)
                Case "F2": If d2.exists(code) Then stoUp = d2(code)
                Case "F5": If d5.exists(code) Then stoUp = d5(code)
            End Select
            wsOut.Cells(r, stoCol).value = stoUp
            keepCnt = keepCnt + 1
        End If
NextR:
    Next r

    ' Purchase request 헤더 → Remain
    wsOut.Cells(1, prCol).value = "Remain"

    wsOut.Activate
    wsOut.Range("A1").Select
    MsgBox "'Air up' 생성 완료!" & vbCrLf & vbCrLf & _
           "· Remain(E열) = Purchase request ? 기존 Stock" & vbCrLf & _
           "· 남은 행 STOCK 자리 = Sto up 값(없으면 0)" & vbCrLf & _
           "· 기존 재고로 충족되어 삭제된 행 : " & delCnt & "건" & vbCrLf & _
           "· 남은(발주 필요) 행 : " & keepCnt & "건", vbInformation

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
EH:
    MsgBox "오류가 발생했습니다: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'--- [함수] Sto up 시트 -> {코드 : 재고수량} 사전 (코드열+2칸) ---
Private Function BuildStoUpDict_(ws As Worksheet) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    If ws Is Nothing Then Set BuildStoUpDict_ = d: Exit Function

    Dim codeCol As Long
    codeCol = FindCodeCol_(ws)
    If codeCol = 0 Then Set BuildStoUpDict_ = d: Exit Function

    Dim last As Long, r As Long, cC As String, key As String
    last = ws.Cells(ws.Rows.Count, codeCol).End(xlUp).Row
    For r = 1 To last
        cC = CStr(ws.Cells(r, codeCol).value)
        If cC Like "M0*" Or cC Like "S0*" Then
            key = NormCode_(cC)
            If Not d.exists(key) Then d.Add key, NzNum_(ws.Cells(r, codeCol + 2).value)
        End If
    Next r
    Set BuildStoUpDict_ = d
End Function

'--- [함수] 코드열(M0*/S0*) 위치 자동 감지 ---
Private Function FindCodeCol_(ws As Worksheet) As Long
    Dim r As Long, c As Long, v As String
    For r = 1 To 30
        For c = 1 To 15
            v = CStr(ws.Cells(r, c).value)
            If v Like "M0*" Or v Like "S0*" Then
                FindCodeCol_ = c: Exit Function
            End If
        Next c
    Next r
    FindCodeCol_ = 0
End Function

'--- [함수] 헤더명으로 열 찾기 (공백/특수문자/대소문자 무시) ---
Private Function FindHdr_(ws As Worksheet, nameWanted As String) As Long
    Dim i As Long, h As String, t As String
    t = NormHdr_(nameWanted)
    For i = 1 To 60
        h = NormHdr_(ws.Cells(1, i).value)
        If Len(h) > 0 Then
            If h = t Or h Like t & "*" Or t Like h & "*" Then
                FindHdr_ = i: Exit Function
            End If
        End If
    Next i
    FindHdr_ = 0
End Function

'--- [함수] 시트명으로 시트 찾기 (공백/대소문자 무시, 완전→포함) ---
Private Function FindSheet_(wb As Workbook, nameWanted As String) As Worksheet
    Dim sh As Worksheet, t As String
    t = NormHdr_(nameWanted)
    For Each sh In wb.Worksheets
        If NormHdr_(sh.Name) = t Then Set FindSheet_ = sh: Exit Function
    Next sh
    For Each sh In wb.Worksheets
        If InStr(1, NormHdr_(sh.Name), t, vbTextCompare) > 0 Then Set FindSheet_ = sh: Exit Function
    Next sh
    Set FindSheet_ = Nothing
End Function

'--- [함수] 헤더/시트명 정규화 (공백·.·_·- 제거, 대문자) ---
Private Function NormHdr_(v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(Replace(Replace(Replace(s, ".", ""), " ", ""), "_", ""), "-", "")
    NormHdr_ = UCase(s)
End Function

'--- [함수] 품목코드 정규화 (공백만 제거, 대문자 · 하이픈 유지) ---
Private Function NormCode_(v As Variant) As String
    NormCode_ = UCase(Trim(Replace(CStr(v), " ", "")))
End Function

'--- [함수] FAC 정규화 (F1/F2/F5) ---
Private Function NormFac_(v As Variant) As String
    NormFac_ = UCase(Replace(Trim(CStr(v)), " ", ""))
End Function

'--- [함수] 텍스트 숫자를 Double로 안전 변환 (콤마 제거) ---
Private Function NzNum_(v As Variant) As Double
    If IsError(v) Or CStr(v) = "" Then
        NzNum_ = 0
    Else
        NzNum_ = Val(Replace(CStr(v), ",", ""))
    End If
End Function

