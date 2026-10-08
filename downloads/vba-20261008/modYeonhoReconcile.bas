Attribute VB_Name = "modYeonhoReconcile"
Option Explicit

Sub 연호_입고점검()
    Dim wsY As Worksheet, wsD As Worksheet, wsM As Worksheet, wsR As Worksheet
    Dim vY As Variant, vD As Variant, vM As Variant
    Dim dictY As Object, dictD As Object, dictM As Object, dictColor As Object
    Dim i As Long, outRow As Long, lastRow As Long
    Dim fac As String, dName As String, yName As String, keyStr As String
    Dim qtyY As Double, qtyD As Double, diff As Double, k As Variant
    
    ' 1. 시트 설정
    On Error Resume Next
    Set wsY = Sheets("연호 ERP"): Set wsD = Sheets("대아 ERP"): Set wsM = Sheets("마스터")
    On Error GoTo 0
    
    If wsY Is Nothing Or wsD Is Nothing Or wsM Is Nothing Then
        MsgBox "시트 이름을 확인하세요: [연호 ERP, 대아 ERP, 마스터]", vbCritical
        Exit Sub
    End If
    
    ' 2. 데이터 로드 및 마스터 색상 저장
    ' 마스터 시트 (A:연호명, B:대아명)
    lastRow = wsM.Cells(wsM.Rows.Count, "B").End(-4162).Row
    Set dictM = CreateObject("Scripting.Dictionary")
    Set dictColor = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        dName = Trim(CStr(wsM.Cells(i, 2).value))
        yName = Trim(CStr(wsM.Cells(i, 1).value))
        If dName <> "" Then
            Dim cleanD As String: cleanD = GetCleanKey(dName)
            Dim cleanY As String: cleanY = GetCleanKey(yName)
            
            dictM(cleanD) = cleanY ' 대아명을 연호명으로 매칭
            
            ' [핵심] B열(대아항목)에 색깔이 칠해져 있다면 연호명을 키로 색상 저장
            If wsM.Cells(i, 2).Interior.ColorIndex <> -4142 Then
                dictColor(cleanY) = wsM.Cells(i, 2).Interior.Color
            End If
        End If
    Next i
    
    ' 연호 ERP 집계 (A:품명, B:수량, C:Fac)
    lastRow = wsY.Cells(wsY.Rows.Count, "A").End(-4162).Row
    vY = wsY.Range("A1:C" & lastRow).value
    Set dictY = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(vY, 1)
        fac = Trim(CStr(vY(i, 3)))
        cleanY = GetCleanKey(vY(i, 1))
        keyStr = fac & "|" & cleanY
        dictY(keyStr) = dictY(keyStr) + Val(vY(i, 2))
    Next i
    
    ' 대아 ERP 집계 (A:Fac, B:품명, C:수량)
    lastRow = wsD.Cells(wsD.Rows.Count, "B").End(-4162).Row
    vD = wsD.Range("A1:C" & lastRow).value
    Set dictD = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(vD, 1)
        fac = Trim(CStr(vD(i, 1)))
        cleanD = GetCleanKey(vD(i, 2))
        ' 마스터 번역
        If dictM.exists(cleanD) Then
            yName = dictM(cleanD)
        Else
            yName = cleanD
        End If
        keyStr = fac & "|" & yName
        dictD(keyStr) = dictD(keyStr) + Val(vD(i, 3))
    Next i
    
    ' 3. 결과 시트 생성
    On Error Resume Next
    Set wsR = Sheets("최종수량대조_결과")
    If wsR Is Nothing Then Set wsR = Sheets.Add(After:=Sheets(Sheets.Count)): wsR.Name = "최종수량대조_결과"
    wsR.Cells.Clear
    On Error GoTo 0
    
    wsR.Range("A1:F1").value = Array("Fac", "품명(연호기준)", "연호 수량", "대아 수량", "차이(연호-대아)", "상태")
    outRow = 2
    
    ' 4. 대조 분석 및 색상 적용
    Dim allKeys As Object: Set allKeys = CreateObject("Scripting.Dictionary")
    For Each k In dictY.keys: allKeys(k) = True: Next
    For Each k In dictD.keys: allKeys(k) = True: Next
    
    For Each k In allKeys.keys
        fac = Split(CStr(k), "|")(0)
        Dim partKey As String: partKey = Split(CStr(k), "|")(1)
        
        qtyY = 0: If dictY.exists(k) Then qtyY = dictY(k)
        qtyD = 0: If dictD.exists(k) Then qtyD = dictD(k)
        diff = qtyY - qtyD
        
        wsR.Cells(outRow, 1).value = fac
        wsR.Cells(outRow, 2).value = partKey
        wsR.Cells(outRow, 3).value = qtyY
        wsR.Cells(outRow, 4).value = qtyD
        
        ' 차이가 있으면 표시, 없으면 공란
        If Abs(diff) > 0.0001 Then
            wsR.Cells(outRow, 5).value = diff
            wsR.Cells(outRow, 6).value = "불일치"
            wsR.Range("A" & outRow & ":F" & outRow).Font.Color = RGB(200, 0, 0)
        Else
            wsR.Cells(outRow, 5).value = ""
            wsR.Cells(outRow, 6).value = ""
        End If
        
        ' [색상 동기화] 마스터에서 칠한 색상을 결과 시트 수량 칸에 그대로 적용
        If dictColor.exists(partKey) Then
            With wsR.Cells(outRow, 3).Resize(1, 3) ' 연호수량, 대아수량, 차이 칸 색칠
                .Interior.Color = dictColor(partKey)
                If dictColor(partKey) < 8000000 Then .Font.Color = RGB(255, 255, 255) ' 어두운 색이면 글자 흰색
            End With
        End If
        
        outRow = outRow + 1
    Next k
    
    wsR.Columns.AutoFit
    MsgBox "최종 수량 대조가 완료되었습니다! 주황색 관리 품목을 확인하세요.", vbInformation
End Sub

' 이름 세척 함수 (YEONHO, WHT, 공백 제거)
Function GetCleanKey(ByVal v As Variant) As String
    Dim s As String: s = UCase(Trim(CStr(v)))
    s = Replace(s, "(YEONHO)", "")
    s = Replace(s, "YEONHO", "")
    s = Replace(s, "WHT", "")
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    GetCleanKey = s
End Function


