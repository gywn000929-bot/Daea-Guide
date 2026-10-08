Attribute VB_Name = "modYeonhoStockByCode"
Option Explicit

Sub 연호_발주_재고포함()
    ' --- 모든 변수 선언 (Option Explicit 대응) ---
    Dim wsERP As Worksheet, wsShip As Worksheet, wsAir As Worksheet, wsMaster As Worksheet, wsFinal As Worksheet
    Dim wsStock As Worksheet                              ' [추가] '재고' 시트 (선택)
    Dim vERP As Variant, vShip As Variant, vAir As Variant, vMaster As Variant, vStock As Variant
    Dim dictMaster As Object, dictERP As Object, dictUrgent As Object, dictColor As Object
    Dim dictShipDemand As Object, dictAirDemand As Object, dictIncoming As Object
    Dim dictStatus As Object, dictPrettyName As Object, dictRawDaea As Object
    Dim dictBalance As Object ' 통합 BALANCE 주머니
    Dim dictStock As Object, dictCode As Object          ' [추가] 공장별 본사재고 / 거래처코드→품명 매핑
    Dim dictDemandKeys As Object, shipKeys As Object, airOnlyKeys As Object

    Dim i As Long, outRow As Long, lastRow As Long
    Dim fac As String, rawName As String, cleanKey As String, yName As String, keyStr As String
    Dim arrears As Double, incomingQty As Double, totalDemand As Double
    Dim shipNeed As Double, airNeed As Double, finalOrderQty As Double, k As Variant
    Dim cleanYName As String, cleanERPName As String, statusVal As String, qVal As Double
    Dim dIdx As Integer, combinedKeys As Variant, remainNeed As Double, appliedUrgent As Double, uStock As Double
    Dim scode As String                                  ' [추가] 재고 거래처 품목코드(정규화)

    ' 1. 시트 설정
    On Error Resume Next
    Set wsERP = Sheets("연호 ERP"): Set wsShip = Sheets("선박"): Set wsAir = Sheets("Air"): Set wsMaster = Sheets("마스터")
    Set wsStock = Sheets("재고")                          ' [추가] 있으면 사용, 없으면 Nothing (선택)
    On Error GoTo 0

    If wsMaster Is Nothing Or wsERP Is Nothing Or wsShip Is Nothing Or wsAir Is Nothing Then
        MsgBox "시트 이름을 확인해주세요. (연호 ERP, 선박, Air, 마스터 필수)", vbCritical
        Exit Sub
    End If

    ' Dictionary 객체 초기화
    Set dictMaster = CreateObject("Scripting.Dictionary")
    Set dictERP = CreateObject("Scripting.Dictionary")
    Set dictUrgent = CreateObject("Scripting.Dictionary")
    Set dictShipDemand = CreateObject("Scripting.Dictionary")
    Set dictAirDemand = CreateObject("Scripting.Dictionary")
    Set dictDemandKeys = CreateObject("Scripting.Dictionary")
    Set dictIncoming = CreateObject("Scripting.Dictionary")
    Set dictColor = CreateObject("Scripting.Dictionary")
    Set dictStatus = CreateObject("Scripting.Dictionary")
    Set dictPrettyName = CreateObject("Scripting.Dictionary")
    Set dictRawDaea = CreateObject("Scripting.Dictionary")
    Set dictBalance = CreateObject("Scripting.Dictionary")
    Set dictStock = CreateObject("Scripting.Dictionary")   ' [추가]
    Set dictCode = CreateObject("Scripting.Dictionary")    ' [추가]
    Set shipKeys = CreateObject("Scripting.Dictionary")
    Set airOnlyKeys = CreateObject("Scripting.Dictionary")

    ' 2. [마스터] 로드
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "B").End(-4162).Row
    If lastRow >= 2 Then
        vMaster = wsMaster.Range("A1:B" & lastRow).value
        For i = 2 To UBound(vMaster, 1)
            rawName = CStr(vMaster(i, 2))
            If Trim(rawName) <> "" Then
                cleanKey = GetPerfectCleanKey(rawName)
                dictMaster(cleanKey) = Trim(CStr(vMaster(i, 1)))
                If wsMaster.Cells(i, 2).Interior.ColorIndex <> -4142 Then
                    dictColor(GetPerfectCleanKey(dictMaster(cleanKey))) = wsMaster.Cells(i, 2).Interior.Color
                End If
            End If
        Next i
    End If

    ' 3. [연호 ERP] 재고 로드
    lastRow = wsERP.Cells(wsERP.Rows.Count, "A").End(-4162).Row
    If lastRow >= 2 Then
        vERP = wsERP.Range("A1:C" & lastRow).value
        For i = 2 To UBound(vERP, 1)
            cleanERPName = GetPerfectCleanKey(CStr(vERP(i, 1)))
            fac = "": If UBound(vERP, 2) >= 3 Then fac = GetFacCodeFinal(CStr(vERP(i, 3)))
            qVal = 0: If UBound(vERP, 2) >= 2 Then qVal = Val(Replace(CStr(vERP(i, 2)), ",", ""))
            If fac = "긴급" Then
                dictUrgent(cleanERPName) = dictUrgent(cleanERPName) + qVal
            ElseIf fac <> "" Then
                dictERP(fac & "|" & cleanERPName) = dictERP(fac & "|" & cleanERPName) + qVal
            End If
        Next i
    End If

    ' 4. [선박] 로드 (BALANCE: F열)
    lastRow = wsShip.Cells(wsShip.Rows.Count, "C").End(-4162).Row
    If lastRow >= 2 Then
        vShip = wsShip.Range("A1:F" & lastRow).value
        For i = 2 To UBound(vShip, 1)
            rawName = CStr(vShip(i, 3))
            If Trim(rawName) <> "" Then
                fac = GetFacCodeFinal(CStr(vShip(i, 1)))
                cleanKey = GetPerfectCleanKey(rawName)
                If dictMaster.exists(cleanKey) Then
                    yName = dictMaster(cleanKey): statusVal = "매칭완료"
                Else
                    yName = rawName: statusVal = "미등록"
                End If
                cleanYName = GetPerfectCleanKey(yName)
                keyStr = fac & "|" & cleanYName
                dictShipDemand(keyStr) = dictShipDemand(keyStr) + Val(Replace(CStr(vShip(i, 4)), ",", ""))
                dictIncoming(keyStr) = dictIncoming(keyStr) + Val(Replace(CStr(vShip(i, 5)), ",", ""))

                ' [추가] 거래처 품목코드(B열) → 품명(cleanYName) 매핑 (재고 매칭용)
                If Trim(CStr(vShip(i, 2))) <> "" Then dictCode(GetPerfectCleanKey(CStr(vShip(i, 2)))) = cleanYName

                ' BALANCE 통합 저장 (선박 시트 우선)
                If UBound(vShip, 2) >= 6 Then
                    qVal = Val(Replace(CStr(vShip(i, 6)), ",", ""))
                    If qVal <> 0 Then dictBalance(keyStr) = qVal
                End If

                shipKeys(keyStr) = True: dictDemandKeys(keyStr) = True
                dictStatus(keyStr) = statusVal
                dictPrettyName(keyStr) = Trim(Replace(Replace(yName, "(YEONHO)", ""), "YEONHO", ""))
                If Not dictRawDaea.exists(keyStr) Then dictRawDaea(keyStr) = rawName
            End If
        Next i
    End If

    ' 5. [Air] 로드 (BALANCE: E열)
    lastRow = wsAir.Cells(wsAir.Rows.Count, "C").End(-4162).Row
    If lastRow >= 2 Then
        vAir = wsAir.Range("A1:E" & lastRow).value
        For i = 2 To UBound(vAir, 1)
            rawName = CStr(vAir(i, 3))
            If Trim(rawName) <> "" Then
                fac = GetFacCodeFinal(CStr(vAir(i, 1)))
                cleanKey = GetPerfectCleanKey(rawName)
                If dictMaster.exists(cleanKey) Then
                    yName = dictMaster(cleanKey): statusVal = "매칭완료"
                Else
                    yName = rawName: statusVal = "미등록"
                End If
                cleanYName = GetPerfectCleanKey(yName)
                keyStr = fac & "|" & cleanYName
                dictAirDemand(keyStr) = dictAirDemand(keyStr) + Val(Replace(CStr(vAir(i, 4)), ",", ""))

                ' [추가] 거래처 품목코드(B열) → 품명 매핑 (선박에 없던 것 보완)
                If Trim(CStr(vAir(i, 2))) <> "" Then
                    If Not dictCode.exists(GetPerfectCleanKey(CStr(vAir(i, 2)))) Then _
                        dictCode(GetPerfectCleanKey(CStr(vAir(i, 2)))) = cleanYName
                End If

                ' BALANCE 통합 저장 (선박에 없었을 경우 보완)
                If UBound(vAir, 2) >= 5 Then
                    qVal = Val(Replace(CStr(vAir(i, 5)), ",", ""))
                    If qVal <> 0 Then dictBalance(keyStr) = qVal
                End If

                If Not shipKeys.exists(keyStr) Then airOnlyKeys(keyStr) = True
                dictDemandKeys(keyStr) = True
                If Not dictStatus.exists(keyStr) Then dictStatus(keyStr) = statusVal
                If Not dictPrettyName.exists(keyStr) Then dictPrettyName(keyStr) = Trim(Replace(Replace(yName, "(YEONHO)", ""), "YEONHO", ""))
                If Not dictRawDaea.exists(keyStr) Then dictRawDaea(keyStr) = rawName
            End If
        Next i
    End If

    ' 5.5 [재고] 로드 (선택 시트) ? A:공장, B:거래처 품목코드, C:품목명, D:재고수량(출고수량 합계)
    '     공장별 본사재고를 fac|품명 키로 담아, 최종발주 '본사대기(F열)' 를 공장별로 정확히 채운다.
    If Not wsStock Is Nothing Then
        lastRow = wsStock.Cells(wsStock.Rows.Count, "C").End(-4162).Row
        If lastRow >= 2 Then
            vStock = wsStock.Range("A1:D" & lastRow).value
            For i = 2 To UBound(vStock, 1)
                rawName = CStr(vStock(i, 3))
                If Trim(rawName) <> "" Then
                    fac = GetFacCodeFinal(CStr(vStock(i, 1)))
                    ' 매칭: 거래처 품목코드(B) 우선 → 없으면 마스터(품명) → 그래도 없으면 품명 그대로
                    scode = GetPerfectCleanKey(CStr(vStock(i, 2)))
                    If dictCode.exists(scode) Then
                        cleanYName = dictCode(scode)
                    ElseIf dictMaster.exists(GetPerfectCleanKey(rawName)) Then
                        cleanYName = GetPerfectCleanKey(dictMaster(GetPerfectCleanKey(rawName)))
                    Else
                        cleanYName = GetPerfectCleanKey(rawName)
                    End If
                    keyStr = fac & "|" & cleanYName
                    dictStock(keyStr) = dictStock(keyStr) + Val(Replace(CStr(vStock(i, 4)), ",", ""))
                End If
            Next i
        End If
    End If

    ' 6. 결과 시트 생성 (A~I열 구조)
    On Error Resume Next
    Set wsFinal = Sheets("최종발주"): If wsFinal Is Nothing Then Set wsFinal = Sheets.Add(After:=Sheets(Sheets.Count)): wsFinal.Name = "최종발주"
    wsFinal.Cells.Clear: On Error GoTo 0

    ' [전체 서식 1] 시트 전체 글꼴, 크기, 검정색 적용
    With wsFinal.Cells.Font
        .Name = "Calibri"
        .Size = 11
        .Color = RGB(0, 0, 0)
    End With

    wsFinal.Range("A1:I1").value = Array("FAC", "품명", "선박", "AIR", "필요합", "본사대기", "연호대기", "발주필요량", "발란스")
    outRow = 2

    ' 7. 최종 출력
    combinedKeys = Array(shipKeys, airOnlyKeys)
    For dIdx = 0 To 1
        For Each k In combinedKeys(dIdx).keys
            fac = Split(CStr(k), "|")(0): cleanYName = Split(CStr(k), "|")(1)
            shipNeed = dictShipDemand(k): airNeed = dictAirDemand(k)
            totalDemand = shipNeed + airNeed

            ' [수정] 본사대기(본사재고): '재고' 시트가 있으면 공장별 재고 사용, 없으면 기존(선박 E열)
            If Not wsStock Is Nothing Then
                incomingQty = dictStock(k)
            Else
                incomingQty = dictIncoming(k)
            End If
            arrears = dictERP(k)

            remainNeed = totalDemand - incomingQty - arrears: If remainNeed < 0 Then remainNeed = 0

            ' 긴급 차감 블록
            appliedUrgent = 0
            If remainNeed > 0 Then
                If dictUrgent.exists(cleanYName) Then
                    uStock = dictUrgent(cleanYName)
                    If uStock >= remainNeed Then
                        appliedUrgent = remainNeed
                        dictUrgent(cleanYName) = uStock - remainNeed
                    Else
                        appliedUrgent = uStock
                        dictUrgent(cleanYName) = 0
                    End If
                End If
            End If

            finalOrderQty = remainNeed - appliedUrgent: If finalOrderQty < 0 Then finalOrderQty = 0

            ' 값 입력
            wsFinal.Cells(outRow, 1).value = fac              ' A열: FAC
            wsFinal.Cells(outRow, 2).value = dictRawDaea(k)   ' B열: 품명(원형)
            wsFinal.Cells(outRow, 3).value = shipNeed         ' C열: 선박
            wsFinal.Cells(outRow, 4).value = airNeed          ' D열: AIR
            wsFinal.Cells(outRow, 5).value = totalDemand      ' E열: 필요합
            wsFinal.Cells(outRow, 6).value = incomingQty      ' F열: 본사대기(=본사재고, 공장별)
            wsFinal.Cells(outRow, 7).value = arrears          ' G열: 연호대기
            wsFinal.Cells(outRow, 8).value = finalOrderQty    ' H열: 발주필요량
            wsFinal.Cells(outRow, 9).value = dictBalance(k)   ' I열: 발란스

            ' [수정] 터미널(마스터에서 색상이 있던 품목) 서식: 배경색 주황색(#FFA500), 흰색 폰트 제거(검정유지)
            If dictColor.exists(cleanYName) Then
                With wsFinal.Cells(outRow, 8)
                    .Interior.Color = RGB(255, 165, 0)
                    .Font.Bold = True
                End With
            End If

            ' 본사대기 + 연호대기가 AIR 수량보다 작으면 발주필요량(H열)의 폰트를 빨간색(#FF0000) 및 볼드체로 덮어쓰기
            If (incomingQty + arrears) < airNeed Then
                With wsFinal.Cells(outRow, 8)
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = True
                End With
            End If

            outRow = outRow + 1
        Next k
    Next dIdx

    ' [전체 서식 2] C열~I열 1,000단위 콤마 적용
    If outRow > 2 Then
        wsFinal.Range("C2:I" & outRow - 1).NumberFormat = "#,##0"
    End If

    wsFinal.Columns.AutoFit
    MsgBox "서식 변경 및 조건부 글꼴 색상 적용이 완료되었습니다!" & _
           IIf(wsStock Is Nothing, "", vbCrLf & "('재고' 시트를 반영해 본사대기를 공장별로 채웠습니다.)"), vbInformation
End Sub

' --- 보조 함수: 세척 ---
Function GetPerfectCleanKey(ByVal v As Variant) As String
    Dim s As String: s = UCase(Trim(CStr(v)))
    s = Replace(s, "(YEONHO)", ""): s = Replace(s, "YEONHO", "")
    s = Replace(s, "RED", "RE"): s = Replace(s, "GRN", "GN"): s = Replace(s, "BLU", "BL"): s = Replace(s, "YEL", "YE")
    s = Replace(s, "WHT", ""): s = Replace(s, "WH", "")
    s = Replace(s, " ", ""): s = Replace(s, "-", ""): s = Replace(s, "(", ""): s = Replace(s, ")", ""): s = Replace(s, "[", ""): s = Replace(s, "]", ""): s = Replace(s, "/", ""): s = Replace(s, ".", "")
    GetPerfectCleanKey = s
End Function

' --- 보조 함수: 공장 추출 ---
Function GetFacCodeFinal(ByVal v As String) As String
    Dim s As String: s = UCase(Trim(v))
    If InStr(s, "긴급") > 0 Then
        GetFacCodeFinal = "긴급"
    ElseIf InStr(s, "F1") > 0 Then
        GetFacCodeFinal = "F1"
    ElseIf InStr(s, "F2") > 0 Then
        GetFacCodeFinal = "F2"
    ElseIf InStr(s, "F5") > 0 Then
        GetFacCodeFinal = "F5"
    Else
        GetFacCodeFinal = s
    End If
End Function

