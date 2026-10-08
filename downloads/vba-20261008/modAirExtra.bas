Attribute VB_Name = "modAirExtra"
Option Explicit



'--- [메인] 에어 데이터 정리 (모든 조건 통합 최종본) ---

Sub 에어리스트_Add()

    Dim ws As Worksheet

    Dim lastRow As Long, i As Long, j As Long

    Dim facCol As Long, itemCol As Long, makerCol As Long

    Dim reqCol As Long, balCol As Long, stoCol As Long, b1Col As Long, brCol As Long

    Dim partCol As Long, remCol As Long, jaegoCol As Long

    Dim facList As Variant, fName As String, fNum As String

    Dim reqVal As String, parts() As String



    Set ws = ActiveSheet



    ' 속도 향상 및 화면 깜빡임 방지

    Application.ScreenUpdating = False

    Application.EnableEvents = False

    Application.Calculation = xlCalculationManual



    On Error GoTo EH



    ' 1. 이미지 컬럼 순서대로 타겟 시트 재배치 (REMARK 추가)

    Call RearrangeColumnsToTarget_Final(ws)



    ' 필터 해제

    If ws.AutoFilterMode Then ws.AutoFilterMode = False



    ' 2. 핵심 헤더 위치 찾기

    facCol = FindHeaderColFlexible(ws, Array("FAC"))        ' A열

    itemCol = FindHeaderColFlexible(ws, Array("ITEM CD"))   ' B열

    makerCol = FindHeaderColFlexible(ws, Array("MAKER"))    ' C열

    partCol = FindHeaderColFlexible(ws, Array("PART NAME")) ' D열

    reqCol = FindHeaderColFlexible(ws, Array("PURCHASE REQUEST"))

    balCol = FindHeaderColFlexible(ws, Array("BALANCE"))

    stoCol = FindHeaderColFlexible(ws, Array("STOCK"))

    b1Col = FindHeaderColFlexible(ws, Array("BALANCE 1"))

    brCol = FindHeaderColFlexible(ws, Array("B-R"))

    remCol = FindHeaderColFlexible(ws, Array("REMARK"))

    jaegoCol = FindHeaderColFlexible(ws, Array("재고"))



    If facCol = 0 Or itemCol = 0 Or reqCol = 0 Then

        MsgBox "필수 헤더를 찾을 수 없습니다. 원본 데이터의 헤더명을 확인해주세요.", vbExclamation

        GoTo SafeExit

    End If



    lastRow = ws.Cells(ws.Rows.Count, itemCol).End(xlUp).Row

    If lastRow < 2 Then GoTo SafeExit



    ' 3. 공장별 데이터 매칭 (다른 시트에서 VLOOKUP)

    facList = Array("F1", "F2", "F5")

    For j = LBound(facList) To UBound(facList)

        fName = facList(j)

        fNum = Right(fName, 1)

        Call ProcessAllLookups(ws, lastRow, facCol, itemCol, makerCol, fName, fNum, balCol, stoCol)

    Next j



    ' 4. 최종 수식 계산, 수요일 수량 업데이트 및 조건부 서식

    For i = 2 To lastRow

        ' [FAC 배경색] A열이 F2이면 노란색(#FFFF00) 칠하기

        If Trim(UCase(ws.Cells(i, facCol).value)) = "F2" Then

            ws.Cells(i, facCol).Interior.Color = RGB(255, 255, 0)

        Else

            ws.Cells(i, facCol).Interior.ColorIndex = xlNone

        End If



        ' [수량 업데이트] Purchase Request 값에 ">>"가 있으면 최종 수량만 추출

        reqVal = CStr(ws.Cells(i, reqCol).value)

        If InStr(reqVal, ">>") > 0 Then

            parts = Split(reqVal, ">>")

            If UBound(parts) >= 1 Then

                ws.Cells(i, reqCol).value = NzNum(Trim(parts(UBound(parts))))

            End If

        Else

            ws.Cells(i, reqCol).value = NzNum(reqVal)

        End If



        ' 계산식: BALANCE 1 = BALANCE + STOCK

        ws.Cells(i, b1Col).value = NzNum(ws.Cells(i, balCol).value) + NzNum(ws.Cells(i, stoCol).value)

        ' 계산식: B-R = BALANCE 1 - PURCHASE REQUEST

        ws.Cells(i, brCol).value = NzNum(ws.Cells(i, b1Col).value) - NzNum(ws.Cells(i, reqCol).value)



        ' 부족분(B-R < 0) 빨간색 글꼴 표시

        If NzNum(ws.Cells(i, brCol).value) < 0 Then

            ws.Cells(i, brCol).Font.Color = vbRed

        Else

            ws.Cells(i, brCol).Font.ColorIndex = xlAutomatic

        End If



        ' [ADD 강조] REMARK에 "ADD"가 있으면 Part Name 빨간색, Purchase Request 빨간색+볼드체

        If remCol > 0 Then

            If InStr(1, UCase(CStr(ws.Cells(i, remCol).value)), "ADD", vbTextCompare) > 0 Then

                If partCol > 0 Then ws.Cells(i, partCol).Font.Color = RGB(255, 0, 0) ' Part Name 붉은색

                If reqCol > 0 Then

                    ws.Cells(i, reqCol).Font.Color = RGB(255, 0, 0) ' Purchase Request 붉은색

                    ws.Cells(i, reqCol).Font.Bold = True            ' Purchase Request 볼드체

                End If

            Else

                If partCol > 0 Then ws.Cells(i, partCol).Font.Color = vbBlack

                If reqCol > 0 Then

                    ws.Cells(i, reqCol).Font.ColorIndex = xlAutomatic

                    ws.Cells(i, reqCol).Font.Bold = False

                End If

            End If

        End If

    Next i



    ' 5. 날짜 포맷 및 URGENT 서식

    Dim dCol As Long, etdCol As Long

    dCol = FindHeaderColFlexible(ws, Array("SHORT DEL"))

    etdCol = FindHeaderColFlexible(ws, Array("NEED ETD"))



    If dCol > 0 Then ws.Columns(dCol).NumberFormat = "dd-mmm-yy"

    If etdCol > 0 Then

        For i = 2 To lastRow

            If InStr(UCase(ws.Cells(i, etdCol).value), "URGENT") > 0 Then

                ws.Cells(i, etdCol).Interior.Color = vbYellow

                ws.Cells(i, etdCol).Font.Color = vbRed

                ws.Cells(i, etdCol).Font.Bold = True

            End If

        Next i

    End If



    ' 6. 데이터 정렬 (1순위: A열 오름차순, 2순위: C열 오름차순, 3순위: D열 오름차순)

    With ws.Sort

        .SortFields.Clear

        .SortFields.Add key:=ws.Columns("A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SortFields.Add key:=ws.Columns("C"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SortFields.Add key:=ws.Columns("D"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SetRange ws.Range("A1").CurrentRegion

        .header = xlYes

        .MatchCase = False

        .Orientation = xlTopToBottom

        .Apply

    End With



    ' 7. 전체 테두리 및 숫자(1,000단위 쉼표) 서식 적용

    ws.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous



    Dim numCols As Variant, colIdx As Variant

    numCols = Array(reqCol, balCol, stoCol, b1Col, brCol, jaegoCol)

    For Each colIdx In numCols

        If colIdx > 0 Then ws.Columns(colIdx).NumberFormat = "#,##0"

    Next colIdx



    ' 부족분 필터 자동 적용 및 열 너비 자동 맞춤

    ws.Range("A1").CurrentRegion.AutoFilter Field:=brCol, Criteria1:="<0"

    ws.Columns.AutoFit



    MsgBox "모든 데이터 정리, 서식 지정 및 정렬이 완료되었습니다!", vbInformation



SafeExit:

    Application.ScreenUpdating = True

    Application.EnableEvents = True

    Application.Calculation = xlCalculationAutomatic

    Exit Sub

EH:

    MsgBox "매크로 실행 중 오류가 발생했습니다." & vbCrLf & "오류 내용: " & Err.Description, vbCritical, "오류"

    Resume SafeExit

End Sub



'--- [서브] 컬럼 순서 재배치 ---

Sub RearrangeColumnsToTarget_Final(ws As Worksheet)

    Dim targetHeaders As Variant, i As Long, lastRow As Long, srcCol As Long

    Dim tempWs As Worksheet



    targetHeaders = Array("FAC", "Item cd", "Maker", "Part Name", "Purchase request", "BALANCE", "STOCK", "BALANCE 1", "B-R", "Pallet N", "재고", "Short Del.", "NEED ETD", "REMARK")



    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then lastRow = ws.UsedRange.Rows.Count



    Set tempWs = Worksheets.Add

    For i = LBound(targetHeaders) To UBound(targetHeaders)

        tempWs.Cells(1, i + 1).value = targetHeaders(i)

        srcCol = FindHeaderColFlexible(ws, Array(targetHeaders(i)))

        If srcCol > 0 Then

            ws.Range(ws.Cells(2, srcCol), ws.Cells(lastRow, srcCol)).Copy tempWs.Cells(2, i + 1)

        End If

    Next i



    ws.Cells.Clear

    tempWs.UsedRange.Copy ws.Range("A1")

    Application.DisplayAlerts = False

    tempWs.Delete

    Application.DisplayAlerts = True

End Sub



'--- [서브] 다른 시트 순회하며 데이터 매칭 ---

Sub ProcessAllLookups(ws As Worksheet, lastRow As Long, fCol As Long, iCol As Long, mCol As Long, fName As String, fNum As String, bCol As Long, sCol As Long)

    Dim sh As Object, sn As String

    For Each sh In ActiveWorkbook.Sheets

        sn = NormalizeHeader(sh.Name)

        If sh.Name <> ws.Name Then

            If IsBalanceSheet(sn, fName, fNum) Then

                Call ApplyValueLookupByFac(ws, lastRow, fCol, iCol, mCol, fName, sh, bCol, False)

            ElseIf IsJstBalanceSheet(sn, fName, fNum) Then

                Call ApplyValueLookupByFac(ws, lastRow, fCol, iCol, mCol, fName, sh, bCol, True)

            ElseIf IsStockSheet(sn, fName, fNum) Then

                Call ApplyValueLookupByFac(ws, lastRow, fCol, iCol, mCol, fName, sh, sCol, False)

            End If

        End If

    Next sh

End Sub



'--- [서브] 에러 없이 안전하게 VLOOKUP 수행 ---

Sub ApplyValueLookupByFac(ws As Worksheet, lastRow As Long, facCol As Long, itemCol As Long, makerCol As Long, fac As String, tWs As Worksheet, targetCol As Long, isJST As Boolean)

    Dim r As Long, c As Long, rr As Long, codeCol As Long

    Dim itemCd As String, makerVal As String, resultVal As Variant, lookupRng As Range



    codeCol = 0

    ' M0 또는 S0로 시작하는 타겟 시트의 코드 열 번호 찾기

    For r = 1 To 20

        For c = 1 To 10

            If CStr(tWs.Cells(r, c).value) Like "M0*" Or CStr(tWs.Cells(r, c).value) Like "S0*" Then

                codeCol = c

                Exit For

            End If

        Next c

        If codeCol > 0 Then Exit For

    Next r



    If codeCol = 0 Then Exit Sub



    Set lookupRng = tWs.Range(tWs.Cells(1, codeCol), tWs.Cells(tWs.Rows.Count, codeCol + 2))



    For rr = 2 To lastRow

        If NormalizeHeader(ws.Cells(rr, facCol).value) = fac Then

            itemCd = Trim(CStr(ws.Cells(rr, itemCol).value))

            makerVal = NormalizeHeader(ws.Cells(rr, makerCol).value)



            If itemCd <> "" Then

                If isJST And makerVal <> "JST" Then GoTo SkipRow



                resultVal = SafeVLookup(itemCd, lookupRng)

                If Not IsError(resultVal) And resultVal <> "" Then

                    ws.Cells(rr, targetCol).value = resultVal

                End If

            End If

        End If

SkipRow:

    Next rr

End Sub



'--- [함수] 안전한 VLOOKUP 반환 ---

Function SafeVLookup(findVal As String, rng As Range) As Variant

    On Error Resume Next

    SafeVLookup = Application.WorksheetFunction.VLOOKUP(findVal, rng, 3, False)

    If Err.Number <> 0 Then

        SafeVLookup = ""

        Err.Clear

    End If

    On Error GoTo 0

End Function



'--- [함수] 헤더명 유연하게 찾기 ---

Function FindHeaderColFlexible(ws As Worksheet, candidates As Variant) As Long

    Dim i As Long, j As Long, h As String, c As String

    For i = 1 To 50

        h = NormalizeHeader(ws.Cells(1, i).value)

        For j = LBound(candidates) To UBound(candidates)

            c = NormalizeHeader(candidates(j))

            If h = c Or h Like c & "*" Or c Like h & "*" Then

                FindHeaderColFlexible = i

                Exit Function

            End If

        Next j

    Next i

    FindHeaderColFlexible = 0

End Function



'--- [함수] 텍스트 공백/특수문자 제거 정규화 ---

Function NormalizeHeader(v As Variant) As String

    Dim s As String

    s = CStr(v)

    s = Replace(Replace(Replace(Replace(s, ".", ""), " ", ""), "_", ""), "-", "")

    NormalizeHeader = UCase(s)

End Function



'--- [함수] 시트 식별용 ---

Function IsBalanceSheet(sn As String, fName As String, fNum As String) As Boolean

    IsBalanceSheet = (sn = fName & "BAL" Or sn = fName & "BALANCE" Or sn = "BALANCE" & fNum)

End Function



Function IsJstBalanceSheet(sn As String, fName As String, fNum As String) As Boolean

    IsJstBalanceSheet = (sn = fName & "JSTBAL" Or sn = fName & "JST" Or sn = "JST" & fNum)

End Function



Function IsStockSheet(sn As String, fName As String, fNum As String) As Boolean

    IsStockSheet = (sn = fName & "STO" Or sn = fName & "STOCK" Or sn = "STOCK" & fNum)

End Function



'--- [함수] 텍스트 숫자를 Double로 안전하게 변환 ---

Function NzNum(v As Variant) As Double

    If IsError(v) Or v = "" Then

        NzNum = 0

    Else

        NzNum = Val(Replace(CStr(v), ",", ""))

    End If

End Function
