Attribute VB_Name = "modYeonhoReceivingFill"
Sub 연호_거래명세서()

    Dim wsSource As Worksheet, wsMatch As Worksheet, wsTarget As Worksheet

    Dim sht As Worksheet

    Dim lastRowSource As Long, lastRowMatch As Long

    Dim i As Long, targetRow As Long

    Dim sourceSheetName As String



    ' ==========================================

    ' 1. 필수 시트 확인 및 자동 탐색 (ActiveWorkbook 기준)

    ' ==========================================

    sourceSheetName = ""



    ' ActiveWorkbook(현재 보고 있는 창)의 모든 시트를 검사

    For Each sht In ActiveWorkbook.Sheets

        ' 시트 이름에서 모든 공백을 제거하고 "거래명세서"가 포함되어 있는지 확인

        If InStr(1, Replace(sht.Name, " ", ""), "거래명세서") > 0 Then

            sourceSheetName = sht.Name

            Exit For

        End If

    Next sht



    ' 만약 여전히 못 찾는다면 디버깅을 위해 메시지 표시

    If sourceSheetName = "" Then

        Dim allNames As String

        For Each sht In ActiveWorkbook.Sheets: allNames = allNames & vbCrLf & "- " & sht.Name: Next sht

MsgBox "현재 파일에서 '거래명세서' 포함 시트를 찾을 수 없습니다." & vbCrLf & "현재 시트 목록:" & allNames, vbCritical, "시트 찾기 오류"

        Exit Sub

    End If



    Set wsSource = ActiveWorkbook.Sheets(sourceSheetName)



    ' 매칭 시트는 코드가 있는 파일(ThisWorkbook)에 있을 확률이 높으므로 둘 다 체크

    On Error Resume Next

    Set wsMatch = ThisWorkbook.Sheets("연호매칭") ' 매크로 파일에 있는 경우

    If wsMatch Is Nothing Then Set wsMatch = ActiveWorkbook.Sheets("연호매칭") ' 데이터 파일에 있는 경우

    On Error GoTo 0



    If wsMatch Is Nothing Then

        MsgBox "'연호매칭' 시트를 찾을 수 없습니다.", vbCritical, "오류"

        Exit Sub

    End If



    ' 결과 시트 설정 (현재 보고 있는 파일에 생성)

    On Error Resume Next

    Set wsTarget = ActiveWorkbook.Sheets("연호입고")

    On Error GoTo 0

    If wsTarget Is Nothing Then

        Set wsTarget = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))

        wsTarget.Name = "연호입고"

    End If



    ' --- 이하 기존 로직과 동일 (공장코드 IF문 오류 수정 반영) ---

    wsTarget.Cells.Clear

    wsTarget.Cells(1, 1).value = "비 고"

    wsTarget.Cells(1, 2).value = "ITEM CD"

    wsTarget.Cells(1, 3).value = "MAKER"

    wsTarget.Cells(1, 5).value = "ITEM NM"

    wsTarget.Cells(1, 6).value = "요청"

    wsTarget.Cells(1, 7).value = "입고수량"



    ' [연호매칭] 데이터 로드 (Dictionary 활용)

    Dim dictSum As Object, dictMatchCD As Object, dictMatchNM As Object

    Set dictSum = CreateObject("Scripting.Dictionary")

    Set dictMatchCD = CreateObject("Scripting.Dictionary")

    Set dictMatchNM = CreateObject("Scripting.Dictionary")



    lastRowMatch = wsMatch.Cells(wsMatch.Rows.Count, "B").End(xlUp).Row

    For i = 2 To lastRowMatch

        Dim itemNameDelta As String, vendorCode As String, clientCode As String

        itemNameDelta = Trim(wsMatch.Cells(i, 1).value)

        vendorCode = Trim(wsMatch.Cells(i, 2).value)

        clientCode = Trim(wsMatch.Cells(i, 3).value)

        If vendorCode <> "" And Not dictMatchCD.exists(vendorCode) Then

            dictMatchCD.Add vendorCode, clientCode

            dictMatchNM.Add vendorCode, itemNameDelta

        End If

    Next i



    ' 거래명세서 정제 & 수량 합산

    Dim checkVal As Variant, isExclude As Boolean, excludeList As Variant, vItem As Variant

    Dim cleanB As String, rawG As String, facCode As String, qty As Double, dictKey As String



    excludeList = Array("거 래 명 세 서", "공", "급", "자", "작 성 일 자", "출고인:", "특기사항")

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row



    For i = 1 To lastRowSource

        checkVal = wsSource.Cells(i, 1).value

        isExclude = False



        If Trim(CStr(checkVal)) = "" Or Trim(CStr(checkVal)) = "No." Or IsDate(checkVal) Then

            isExclude = True

        Else

            For Each vItem In excludeList

                If InStr(CStr(wsSource.Cells(i, 1).value & wsSource.Cells(i, 2).value), vItem) > 0 Then

                    isExclude = True

                    Exit For

                End If

            Next vItem

        End If



        If Not isExclude Then

            qty = Val(wsSource.Cells(i, 3).value)

            cleanB = Trim(Replace(Replace(CStr(wsSource.Cells(i, 2).value), ChrW(160), " "), ChrW(&H3000), " "))

            rawG = UCase(Trim(Replace(Replace(CStr(wsSource.Cells(i, 7).value), ChrW(160), " "), ChrW(&H3000), " ")))



            facCode = ""

            If rawG <> "" Then

                Dim partsG As Variant, p As Variant, temp As String

                partsG = Split(rawG, ",")

                For Each p In partsG

                    temp = Trim(CStr(p))

                    If InStr(temp, "(") > 0 Then temp = Trim(Left(temp, InStr(temp, "(") - 1))



                    ' 공장코드 변환 로직

                    If InStr(1, temp, "F1") > 0 Then

                        temp = "1010"

                    ElseIf InStr(1, temp, "F2") > 0 Then

                        temp = "1060"

                    ElseIf InStr(1, temp, "F5") > 0 Then

                        temp = "1150"

                    End If

                    facCode = facCode & IIf(facCode = "", "", ", ") & temp

                Next p

            End If



            If cleanB <> "" And qty > 0 Then

                dictKey = facCode & "|" & cleanB

                dictSum(dictKey) = dictSum(dictKey) + qty

            End If

        End If

    Next i



    ' 결과 출력

    targetRow = 2

    Dim key As Variant, arrKey() As String

    For Each key In dictSum.keys

        arrKey = Split(key, "|")

        wsTarget.Cells(targetRow, 1).value = arrKey(0) ' 비고

        If dictMatchCD.exists(arrKey(1)) Then

            wsTarget.Cells(targetRow, 2).value = dictMatchCD(arrKey(1))

            wsTarget.Cells(targetRow, 5).value = dictMatchNM(arrKey(1))

        Else

            wsTarget.Cells(targetRow, 2).value = "매칭불가"

            wsTarget.Cells(targetRow, 5).value = arrKey(1)

        End If

        wsTarget.Cells(targetRow, 3).value = "연호"

        wsTarget.Cells(targetRow, 4).value = "3840"

        wsTarget.Cells(targetRow, 6).value = dictSum(key)

        wsTarget.Cells(targetRow, 7).value = dictSum(key)

        targetRow = targetRow + 1

    Next key



    ' 정렬 및 스타일

    If targetRow > 2 Then

        With wsTarget.Sort

            .SortFields.Clear

            .SortFields.Add key:=wsTarget.Range("A2:A" & targetRow - 1), Order:=xlAscending

            .SortFields.Add key:=wsTarget.Range("E2:E" & targetRow - 1), Order:=xlAscending

            .SetRange wsTarget.Range("A1:G" & targetRow - 1): .header = xlYes: .Apply

        End With

        wsTarget.ListObjects.Add(xlSrcRange, wsTarget.Range("A1").CurrentRegion, , xlYes).TableStyle = "TableStyleMedium2"

        wsTarget.Columns.AutoFit

    End If



    MsgBox "작업이 완료되었습니다!", vbInformation

End Sub
