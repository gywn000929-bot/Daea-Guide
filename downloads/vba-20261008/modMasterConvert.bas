Attribute VB_Name = "modMasterConvert"
Option Explicit

Sub 연호_대아_이름_변환()
    Dim ws As Worksheet, wsMaster As Worksheet
    Dim lastRow As Long, lastCol As Long, mRow As Long
    Dim i As Long
    Dim targetColStr As String, targetCol As Long
    
    Dim dictMaster As Object
    Dim vMaster As Variant
    Dim yName As String, dName As String
    Dim cleanY As String, cleanD As String, cleanTarget As String, targetVal As String
    Dim matchedArr As Variant
    
    ' 1. 시트 설정
    Set ws = ActiveSheet
    On Error Resume Next
    Set wsMaster = Sheets("마스터")
    On Error GoTo 0
    
    If wsMaster Is Nothing Then
        MsgBox "'마스터' 시트를 찾을 수 없습니다. 현재 통합 문서에 마스터 시트가 있는지 확인해 주세요.", vbCritical
        Exit Sub
    End If
    
    ' 2. 기준 열 입력
    targetColStr = InputBox("품목명이 적힌 열의 알파벳을 입력하세요." & vbCrLf & "(예: C, F, G 등)", "기준 열 입력")
    If Trim(targetColStr) = "" Then Exit Sub
    
    On Error Resume Next
    targetCol = ws.Columns(targetColStr).Column
    On Error GoTo 0
    
    If targetCol = 0 Then
        MsgBox "올바른 열 알파벳을 입력해 주세요.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 3. 마스터 시트 데이터 메모리에 저장 (핵심: 마스터 원본 그대로 저장)
    Set dictMaster = CreateObject("Scripting.Dictionary")
    
    mRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    If mRow >= 2 Then
        vMaster = wsMaster.Range("A1:B" & mRow).value
        For i = 2 To UBound(vMaster, 1)
            yName = Trim(CStr(vMaster(i, 1))) ' A열: 연호명 원본
            dName = Trim(CStr(vMaster(i, 2))) ' B열: 대아명 원본
            
            cleanY = GetPerfectCleanKey(yName)
            cleanD = GetPerfectCleanKey(dName)
            
            ' 입력값이 연호든 대아든, 무조건 [마스터 A열, 마스터 B열]을 뱉어내도록 배열로 저장
            If cleanY <> "" Then dictMaster(cleanY) = Array(yName, dName)
            If cleanD <> "" Then dictMaster(cleanD) = Array(yName, dName)
        Next i
    End If
    
    ' 4. 마지막 열 찾기
    On Error Resume Next
    lastCol = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookAt:=xlPart, _
                            LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, MatchCase:=False).Column
    On Error GoTo 0
    If lastCol = 0 Then lastCol = 1
    
    ' 5. 헤더 만들기
    Dim outYCol As Long, outDCol As Long
    outYCol = lastCol + 1
    outDCol = lastCol + 2
    
    With ws
        .Cells(1, outYCol).value = "마스터: 연호명"
        .Cells(1, outDCol).value = "마스터: 대아명"
        .Cells(1, outYCol).Interior.Color = RGB(220, 230, 241)
        .Cells(1, outDCol).Interior.Color = RGB(228, 223, 236)
        .Cells(1, outYCol).Font.Bold = True
        .Cells(1, outDCol).Font.Bold = True
        
        lastRow = .Cells(.Rows.Count, targetCol).End(xlUp).Row
        
        ' 6. 본격적인 매칭 로직
        For i = 2 To lastRow
            targetVal = CStr(.Cells(i, targetCol).value)
            
            If Trim(targetVal) <> "" Then
                cleanTarget = GetPerfectCleanKey(targetVal)
                
                ' 마스터에 데이터가 존재하면
                If dictMaster.exists(cleanTarget) Then
                    matchedArr = dictMaster(cleanTarget)
                    .Cells(i, outYCol).value = matchedArr(0) ' 무조건 마스터 A열(연호) 출력
                    .Cells(i, outDCol).value = matchedArr(1) ' 무조건 마스터 B열(대아) 출력
                Else
                    .Cells(i, outYCol).value = "미등록"
                    .Cells(i, outDCol).value = "미등록"
                End If
            End If
        Next i
        
        .Columns(outYCol).AutoFit
        .Columns(outDCol).AutoFit
    End With
    
    Application.ScreenUpdating = True
    MsgBox "마스터 기준 매칭 완료! 데이터 맨 우측을 확인해 보세요.", vbInformation
End Sub

' --- 보조 함수: 세척 (어떤 공백/특수문자가 와도 완벽히 파괴) ---
Function GetPerfectCleanKey(ByVal v As Variant) As String
    Dim s As String
    s = UCase(Trim(CStr(v)))
    
    ' 모든 종류의 공백 완벽 제거
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(160), "")   ' 웹/ERP 논브레이킹 스페이스
    s = Replace(s, ChrW(12288), "") ' 전각 공백
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    
    ' 회사명 제거
    s = Replace(s, "(YEONHO)", ""): s = Replace(s, "YEONHO", "")
    
    ' 컬러 약어 매칭
    s = Replace(s, "RED", "RE"): s = Replace(s, "GRN", "GN"): s = Replace(s, "BLU", "BL"): s = Replace(s, "YEL", "YE")
    s = Replace(s, "BLK", "BK"): s = Replace(s, "VIO", "VI"): s = Replace(s, "ORG", "OR"): s = Replace(s, "BRN", "BR")
    s = Replace(s, "WHT", ""): s = Replace(s, "WH", "")
    
    ' 특수문자 제거
    s = Replace(s, "-", ""): s = Replace(s, "(", ""): s = Replace(s, ")", ""): s = Replace(s, "[", ""): s = Replace(s, "]", ""): s = Replace(s, "/", ""): s = Replace(s, ".", ""): s = Replace(s, "_", ""): s = Replace(s, ",", "")
    GetPerfectCleanKey = s
End Function



