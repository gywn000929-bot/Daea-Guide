Attribute VB_Name = "modSeaArrange"
Option Explicit

Sub 선박리스트_만들기()

    Dim ws As Worksheet, tWs As Worksheet, vendorWs As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim sumCol As Long, balCol As Long, stoCol As Long, b1Col As Long, brCol As Long, vendorCol As Long, itemNmCol As Long
    Dim headerRow As Long, dataRow As Long          ' [추가] 머리글 행 자동 감지 (1행 or 2행 그룹 디자인 대응)
    Dim facList As Variant, fName As String, fNum As String
    Dim sh As Object
    Dim deleteZeroRows As Boolean

    ' 거래처 조회를 위한 변수
    Dim makerDict As Object, exDict As Object
    Dim vRow As Long, vLastRow As Long
    Dim mk As String, vn As String, f_it As String, f_vn As String
    Dim mkStr As String, itStr As String

    Set ws = ActiveSheet
    deleteZeroRows = True   ' 필요 없으면 False 로 바꾸기

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo EH

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' [수정] SUM 머리글을 1~3행에서 찾아 '머리글 행'을 자동 감지
    '        (v2 처럼 1행=그룹라벨, 2행=실제 머리글 인 경우도 지원)
    headerRow = 0: sumCol = 0
    For j = 1 To 3
        sumCol = FindHeaderCol(ws, "SUM", j)
        If sumCol > 0 Then headerRow = j: Exit For
    Next j
    If sumCol = 0 Then
        MsgBox "1~3행에서 'SUM' 제목을 찾을 수 없습니다.", vbExclamation
        GoTo SafeExit
    End If
    dataRow = headerRow + 1

    balCol = sumCol + 1
    stoCol = sumCol + 2
    b1Col = sumCol + 3
    brCol = sumCol + 4
    vendorCol = sumCol + 5 ' 거래처 컬럼 추가

    ws.Cells(headerRow, balCol).value = "BALANCE"
    ws.Cells(headerRow, stoCol).value = "STOCK"
    ws.Cells(headerRow, b1Col).value = "BALANCE 1"
    ws.Cells(headerRow, brCol).value = "B-R"
    ws.Cells(headerRow, vendorCol).value = "거래처" ' 헤더명 지정

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = dataRow To lastRow
        ws.Cells(i, sumCol).value = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i, 5), ws.Cells(i, sumCol - 1)))
    Next i

    If deleteZeroRows Then
        For i = lastRow To dataRow Step -1
            If NzNum(ws.Cells(i, sumCol).value) = 0 Then
                ws.Rows(i).Delete
            End If
        Next i
    End If

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < dataRow Then GoTo FinalCalc

    ws.Range(ws.Cells(dataRow, balCol), ws.Cells(lastRow, balCol)).ClearContents
    ws.Range(ws.Cells(dataRow, stoCol), ws.Cells(lastRow, stoCol)).ClearContents
    ws.Range(ws.Cells(dataRow, b1Col), ws.Cells(lastRow, b1Col)).ClearContents
    ws.Range(ws.Cells(dataRow, brCol), ws.Cells(lastRow, brCol)).ClearContents

    facList = Array("F1", "F2", "F5")

    For j = LBound(facList) To UBound(facList)
        fName = facList(j)
        fNum = Right(fName, 1)

        ' 1. 일반 BALANCE 시트 확인
        Set tWs = Nothing
        For Each sh In ActiveWorkbook.Sheets
            If sh.Name <> ws.Name Then
                If IsBalanceSheet(UCaseNoSpace(sh.Name), fName, fNum) Then
                    Set tWs = sh
                    Exit For
                End If
            End If
        Next sh
        If Not tWs Is Nothing Then
            ApplyValueLookup ws, dataRow, lastRow, fName, tWs, balCol, False
        End If

        ' 2. JST BALANCE 시트 확인
        Set tWs = Nothing
        For Each sh In ActiveWorkbook.Sheets
            If sh.Name <> ws.Name Then
                If IsJstBalanceSheet(UCaseNoSpace(sh.Name), fName, fNum) Then
                    Set tWs = sh
                    Exit For
                End If
            End If
        Next sh
        If Not tWs Is Nothing Then
            ApplyValueLookup ws, dataRow, lastRow, fName, tWs, balCol, True
        End If

        ' 3. 일반 STOCK 시트 확인
        Set tWs = Nothing
        For Each sh In ActiveWorkbook.Sheets
            If sh.Name <> ws.Name Then
                If IsStockSheet(UCaseNoSpace(sh.Name), fName, fNum) Then
                    Set tWs = sh
                    Exit For
                End If
            End If
        Next sh
        If Not tWs Is Nothing Then
            ApplyValueLookup ws, dataRow, lastRow, fName, tWs, stoCol, False
        End If

        ' 4. JST STOCK 시트 확인 (추가된 부분)
        Set tWs = Nothing
        For Each sh In ActiveWorkbook.Sheets
            If sh.Name <> ws.Name Then
                If IsJstStockSheet(UCaseNoSpace(sh.Name), fName, fNum) Then
                    Set tWs = sh
                    Exit For
                End If
            End If
        Next sh
        If Not tWs Is Nothing Then
            ApplyValueLookup ws, dataRow, lastRow, fName, tWs, stoCol, True
        End If
    Next j

FinalCalc:
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = dataRow To lastRow
        ws.Cells(i, b1Col).value = NzNum(ws.Cells(i, balCol).value) + NzNum(ws.Cells(i, stoCol).value)
        ws.Cells(i, brCol).value = NzNum(ws.Cells(i, b1Col).value) - NzNum(ws.Cells(i, sumCol).value)
    Next i

    ' =========================================================
    ' 거래처명 매핑 로직 (딕셔너리 활용)
    ' =========================================================
    On Error Resume Next
    Set vendorWs = ActiveWorkbook.Sheets("거래처코드")
    On Error GoTo EH

    If Not vendorWs Is Nothing Then
        Set makerDict = CreateObject("Scripting.Dictionary")
        Set exDict = CreateObject("Scripting.Dictionary")

        vLastRow = vendorWs.Cells(vendorWs.Rows.Count, "B").End(xlUp).Row

        For vRow = 2 To vLastRow
            mk = UCase(Trim(CStr(vendorWs.Cells(vRow, "B").value)))
            vn = Trim(CStr(vendorWs.Cells(vRow, "A").value))
            If mk <> "" And Not makerDict.exists(mk) Then makerDict.Add mk, vn

            f_it = UCase(Trim(CStr(vendorWs.Cells(vRow, "F").value)))
            f_vn = Trim(CStr(vendorWs.Cells(vRow, "E").value))
            If f_it <> "" And Not exDict.exists(f_it) Then exDict.Add f_it, f_vn
        Next vRow

        itemNmCol = FindHeaderCol(ws, "ITEM NM", headerRow)
        If itemNmCol = 0 Then itemNmCol = FindHeaderCol(ws, "ITEM CD", headerRow)
        If itemNmCol = 0 Then itemNmCol = 2

        For i = dataRow To lastRow
            mkStr = UCase(Trim(CStr(ws.Cells(i, 3).value)))
            itStr = UCase(Trim(CStr(ws.Cells(i, itemNmCol).value)))

            If mkStr = "DONG-A" Or mkStr = "DONG-A BESTECH" Or mkStr = "KET" Then
                If exDict.exists(itStr) Then
                    ws.Cells(i, vendorCol).value = exDict(itStr)
                Else
                    If makerDict.exists(mkStr) Then
                        ws.Cells(i, vendorCol).value = makerDict(mkStr)
                    End If
                End If
            Else
                If makerDict.exists(mkStr) Then
                    ws.Cells(i, vendorCol).value = makerDict(mkStr)
                End If
            End If
        Next i
    Else
        MsgBox "'거래처코드' 시트를 찾을 수 없어 거래처 데이터는 불러오지 못했습니다.", vbExclamation
    End If

    ' =========================================================
    ' 데이터 정렬 (FAC, MAKER, ITEM NM 순)
    ' =========================================================
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(dataRow, 1), ws.Cells(lastRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(dataRow, 3), ws.Cells(lastRow, 3)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(dataRow, itemNmCol), ws.Cells(lastRow, itemNmCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With ws.Sort
        .SetRange ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, vendorCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' =========================================================
    ' 보기 좋게 서식 (머리글 강조 · 0 숨김 · 열 그룹 · 틀고정)
    ' =========================================================
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    Dim etdStart As Long, etdEnd As Long, cC As Long
    etdStart = 5: etdEnd = sumCol - 1

    ' 열 너비
    ws.Columns(1).ColumnWidth = 6           ' FAC
    ws.Columns(2).ColumnWidth = 14          ' ITEM CD
    ws.Columns(3).ColumnWidth = 12          ' MAKER
    ws.Columns(itemNmCol).ColumnWidth = 34  ' ITEM NM
    For cC = etdStart To vendorCol
        If cC = vendorCol Then
            ws.Columns(cC).ColumnWidth = 20
        Else
            ws.Columns(cC).ColumnWidth = 12
        End If
    Next cC

    ' 데이터 글꼴
    With ws.Range(ws.Cells(dataRow, 1), ws.Cells(lastRow, vendorCol)).Font
        .Name = "맑은 고딕": .Size = 10
    End With

    ' 숫자열: 천단위 콤마 + 0 은 빈칸 + 음수 빨강, 가운데 정렬
    With ws.Range(ws.Cells(dataRow, etdStart), ws.Cells(lastRow, brCol))
        .HorizontalAlignment = xlCenter
        .NumberFormat = "#,##0;[Red]-#,##0;"
    End With

    ' 정보열 정렬 (FAC 가운데, 코드/메이커 왼쪽)
    ws.Range(ws.Cells(dataRow, 1), ws.Cells(lastRow, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(dataRow, 2), ws.Cells(lastRow, 3)).HorizontalAlignment = xlLeft

    ' 본문 바탕 흰색
    ws.Range(ws.Cells(dataRow, 1), ws.Cells(lastRow, vendorCol)).Interior.ColorIndex = xlNone

    ' 핵심 열 강조 (SUM · B-R) ? 굵게
    ws.Range(ws.Cells(dataRow, sumCol), ws.Cells(lastRow, sumCol)).Font.Bold = True
    ws.Range(ws.Cells(dataRow, brCol), ws.Cells(lastRow, brCol)).Font.Bold = True

    ' 머리글 (흰 바탕 · 검정 굵은 글씨 · 가운데)
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, vendorCol))
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Name = "맑은 고딕": .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(headerRow).RowHeight = 30

    ' 격자 테두리 (연한 회색)
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, vendorCol)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(208, 214, 224)
        .Weight = xlThin
    End With

    ' ETD 열: 월별 옅은 바탕색 (파랑↔초록 번갈아) ? 머리글+데이터 함께
    Dim ci As Long, moNow As String, moPrev As String, grpIdx As Long, tintClr As Long
    grpIdx = 0: moPrev = ""
    For ci = etdStart To etdEnd
        moNow = MonthTokenOf(CStr(ws.Cells(headerRow, ci).value))
        If moNow <> moPrev And moPrev <> "" Then grpIdx = grpIdx + 1
        If grpIdx Mod 2 = 0 Then
            tintClr = RGB(219, 233, 247)   ' 파랑 계열
        Else
            tintClr = RGB(228, 240, 220)   ' 초록 계열
        End If
        ws.Range(ws.Cells(headerRow, ci), ws.Cells(lastRow, ci)).Interior.Color = tintClr
        moPrev = moNow
    Next ci

    ' 틀 고정 (머리글 + ITEM NM 까지 고정)
    ws.Activate
    ws.Cells(dataRow, etdStart).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    ws.Cells(headerRow, 1).Select

    MsgBox "선박 데이터 (JST STO 포함) 정리가 완료되었습니다!", vbInformation

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

EH:
    MsgBox "오류 번호: " & Err.Number & vbCrLf & "오류 내용: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

' --- 하단 사용자 정의 함수들 ---
'   [수정] dataStart 인자 추가 (머리글 행 다음부터 데이터 처리)
Sub ApplyValueLookup(ws As Worksheet, dataStart As Long, lastRow As Long, fac As String, tWs As Worksheet, targetCol As Long, isJST As Boolean)
    Dim r As Long, c As Long, rr As Long
    Dim codeCol As Long, itemCd As String, makerVal As String
    Dim resultVal As Variant

    codeCol = 0
    For r = 1 To 30
        For c = 1 To 10
            If tWs.Cells(r, c).value Like "M0*" Or tWs.Cells(r, c).value Like "S0*" Then
                codeCol = c
                Exit For
            End If
        Next c
        If codeCol > 0 Then Exit For
    Next r

    If codeCol = 0 Then Exit Sub

    For rr = dataStart To lastRow
        If Trim(UCase(ws.Cells(rr, 1).value)) = fac Then
            If isJST Then
                makerVal = Trim(UCase(ws.Cells(rr, 3).value))
                If makerVal <> "JST" Then GoTo NextRow
                ' 값이 이미 있으면 덮어쓰지 않음
                If NzNum(ws.Cells(rr, targetCol).value) <> 0 Or Trim(CStr(ws.Cells(rr, targetCol).value)) <> "" Then GoTo NextRow
            End If

            itemCd = Trim(CStr(ws.Cells(rr, 2).value))
            If itemCd <> "" Then
                resultVal = SafeVLookup(itemCd, tWs.Range(tWs.Cells(1, codeCol), tWs.Cells(tWs.Rows.Count, codeCol + 2)))
                If Not IsError(resultVal) Then
                    If Trim(CStr(resultVal)) <> "" Then
                        ws.Cells(rr, targetCol).value = resultVal
                    End If
                End If
            End If
        End If
NextRow:
    Next rr
End Sub

Function SafeVLookup(findVal As String, rng As Range) As Variant
    On Error Resume Next
    SafeVLookup = Application.WorksheetFunction.VLOOKUP(findVal, rng, 3, False)
    If Err.Number <> 0 Then
        Err.Clear
        SafeVLookup = ""
    End If
    On Error GoTo 0
End Function

'   [수정] 검색할 행(hrow) 인자 추가 (기본 1행, v2 는 2행)
Function FindHeaderCol(ws As Worksheet, headerName As String, Optional ByVal hrow As Long = 1) As Long
    Dim i As Long
    For i = 1 To 50
        If UCaseNoSpace(ws.Cells(hrow, i).value) = UCaseNoSpace(headerName) Then
            FindHeaderCol = i
            Exit Function
        End If
    Next i
End Function

Function UCaseNoSpace(v As Variant) As String
    UCaseNoSpace = UCase(Replace(Trim(CStr(v)), " ", ""))
End Function

Function IsBalanceSheet(sn As String, fName As String, fNum As String) As Boolean
    IsBalanceSheet = (sn = fName & "BAL" Or sn = fName & "BALANCE" Or sn = "BALANCE" & fName Or sn = "BALANCE" & fNum)
End Function

Function IsJstBalanceSheet(sn As String, fName As String, fNum As String) As Boolean
    IsJstBalanceSheet = (sn = fName & "JSTBAL" Or sn = "JST" & fNum Or sn = fName & "JST" Or sn = fName & "JSTBAL")
End Function

Function IsStockSheet(sn As String, fName As String, fNum As String) As Boolean
    IsStockSheet = (sn = fName & "STO" Or sn = fName & "STOCK" Or sn = "STOCK" & fName)
End Function

' 추가된 JST STO 확인 함수
Function IsJstStockSheet(sn As String, fName As String, fNum As String) As Boolean
    IsJstStockSheet = (sn = fName & "JSTSTO" Or sn = fName & "JSTSTOCK" Or sn = "JSTSTO" & fNum Or sn = "JSTSTOCK" & fNum Or sn = "JST" & fNum & "STO")
End Function

Function NzNum(v As Variant) As Double
    Dim s As String

    If IsError(v) Then
        NzNum = 0
        Exit Function
    End If

    s = Trim(CStr(v))
    If s = "" Then
        NzNum = 0
        Exit Function
    End If

    s = Replace(s, ",", "")
    If IsNumeric(s) Then
        NzNum = CDbl(s)
    Else
        NzNum = 0
    End If
End Function

' 헤더 문자열(예: "ETD 28-AUG-26")에서 월 약어(JAN~DEC)를 추출
Function MonthTokenOf(ByVal s As String) As String
    Dim mons As Variant, i As Long, u As String
    u = UCase(s)
    mons = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    For i = LBound(mons) To UBound(mons)
        If InStr(u, mons(i)) > 0 Then
            MonthTokenOf = mons(i)
            Exit Function
        End If
    Next i
    MonthTokenOf = ""
End Function
