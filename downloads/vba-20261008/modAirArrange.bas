Attribute VB_Name = "modAirArrange"
Option Explicit

'--- [메인] 에어 데이터 정리 (이미지 순서 반영 버전) ---
Sub 에어리스트_만들기()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim facCol As Long, itemCol As Long, makerCol As Long
    Dim reqCol As Long, balCol As Long, stoCol As Long, b1Col As Long, brCol As Long
    Dim facList As Variant, fName As String, fNum As String
    
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo EH
    
    ' 1. 이미지의 컬럼 순서로 재배치 (A~M열)
    Call RearrangeColumnsToTarget_Final(ws)
    
    ' 2. 재배치된 시트에서 핵심 헤더 위치 재확인
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    facCol = FindHeaderColFlexible(ws, Array("FAC"))
    itemCol = FindHeaderColFlexible(ws, Array("ITEM CD"))
    makerCol = FindHeaderColFlexible(ws, Array("MAKER"))
    reqCol = FindHeaderColFlexible(ws, Array("PURCHASE REQUEST"))
    balCol = FindHeaderColFlexible(ws, Array("BALANCE"))
    stoCol = FindHeaderColFlexible(ws, Array("STOCK"))
    b1Col = FindHeaderColFlexible(ws, Array("BALANCE 1"))
    brCol = FindHeaderColFlexible(ws, Array("B-R"))
    
    If facCol = 0 Or itemCol = 0 Or reqCol = 0 Then
        MsgBox "필수 헤더를 찾을 수 없습니다. 원본의 헤더명을 확인해주세요.", vbExclamation
        GoTo SafeExit
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, itemCol).End(xlUp).Row
    If lastRow < 2 Then GoTo SafeExit
    
    ' 3. 공장별 데이터 매칭
    facList = Array("F1", "F2", "F5")
    For j = LBound(facList) To UBound(facList)
        fName = facList(j): fNum = Right(fName, 1)
        Call ProcessAllLookups(ws, lastRow, facCol, itemCol, makerCol, fName, fNum, balCol, stoCol)
    Next j
    
    ' 4. 최종 수식 계산 및 서식 (이미지 기준 H, I열)
    For i = 2 To lastRow
        ' BALANCE 1 (H열) = BALANCE + STOCK
        ws.Cells(i, b1Col).value = NzNum(ws.Cells(i, balCol).value) + NzNum(ws.Cells(i, stoCol).value)
        ' B-R (I열) = BALANCE 1 - PURCHASE REQUEST
        ws.Cells(i, brCol).value = NzNum(ws.Cells(i, b1Col).value) - NzNum(ws.Cells(i, reqCol).value)
        
        ' 부족분(B-R < 0)일 경우 빨간색 표시 (이미지처럼)
        If NzNum(ws.Cells(i, brCol).value) < 0 Then
            ws.Cells(i, brCol).Font.Color = vbRed
        End If
    Next i
    
    ' 5. 날짜 및 URGENT 강조 서식
    Dim dCol As Long, etdCol As Long
    dCol = FindHeaderColFlexible(ws, Array("SHORT DEL"))
    etdCol = FindHeaderColFlexible(ws, Array("NEED ETD"))
    
    If dCol > 0 Then ws.Columns(dCol).NumberFormat = "dd-mmm-yy" ' 17-Mar-26 형태
    If etdCol > 0 Then
        ' NEED ETD 열에 URGENT가 있으면 노란색 배경 (이미지 반영)
        For i = 2 To lastRow
            If InStr(UCase(ws.Cells(i, etdCol).value), "URGENT") > 0 Then
                ws.Cells(i, etdCol).Interior.Color = vbYellow
                ws.Cells(i, etdCol).Font.Color = vbRed
                ws.Cells(i, etdCol).Font.Bold = True
            End If
        Next i
    End If
    
    ' 부족분 필터 적용
    ws.Range("A1").CurrentRegion.AutoFilter Field:=brCol, Criteria1:="<0"
    ws.Columns.AutoFit
    
    MsgBox "이미지 순서대로 정리가 완료되었습니다!", vbInformation

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
EH:
    MsgBox "오류: " & Err.Description, vbCritical
    Resume SafeExit
End Sub

'--- [핵심 수정] 이미지의 정확한 컬럼 순서 정의 ---
Sub RearrangeColumnsToTarget_Final(ws As Worksheet)
    Dim targetHeaders As Variant, i As Long, lastRow As Long, srcCol As Long
    Dim tempWs As Worksheet
    
    ' 이미지에서 확인된 순서: A(FAC) ~ M(NEED ETD)
    targetHeaders = Array("FAC", "Item cd", "Maker", "Part Name", "Purchase request", "BALANCE", "STOCK", "BALANCE 1", "B-R", "Pallet N", "재고", "Short Del.", "NEED ETD")
    
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
    Application.DisplayAlerts = False: tempWs.Delete: Application.DisplayAlerts = True
End Sub

'--- [보조 함수들 - 기존과 동일하게 유지하여 컴파일 오류 방지] ---
Sub ProcessAllLookups(ws As Worksheet, lastRow As Long, fCol As Long, iCol As Long, mCol As Long, fName As String, fNum As String, bCol As Long, sCol As Long)
    Dim sh As Object, sn As String
    For Each sh In ActiveWorkbook.Sheets
        sn = NormalizeHeader(sh.Name)
        If sh.Name <> ws.Name Then
            If IsBalanceSheet(sn, fName, fNum) Then
                ApplyValueLookupByFac ws, lastRow, fCol, iCol, mCol, fName, sh, bCol, False
            ElseIf IsJstBalanceSheet(sn, fName, fNum) Then
                ApplyValueLookupByFac ws, lastRow, fCol, iCol, mCol, fName, sh, bCol, True
            ElseIf IsStockSheet(sn, fName, fNum) Then
                ApplyValueLookupByFac ws, lastRow, fCol, iCol, mCol, fName, sh, sCol, False
            End If
        End If
    Next sh
End Sub

Sub ApplyValueLookupByFac(ws As Worksheet, lastRow As Long, facCol As Long, itemCol As Long, makerCol As Long, fac As String, tWs As Worksheet, targetCol As Long, isJST As Boolean)
    Dim r As Long, c As Long, rr As Long, codeCol As Long
    Dim itemCd As String, makerVal As String, resultVal As Variant, lookupRng As Range
    codeCol = 0
    For r = 1 To 20
        For c = 1 To 10
            If CStr(tWs.Cells(r, c).value) Like "M0*" Or CStr(tWs.Cells(r, c).value) Like "S0*" Then
                codeCol = c: Exit For
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
                If isJST And makerVal <> "JST" Then GoTo NextRow
                resultVal = SafeVLookup(itemCd, lookupRng)
                If Not IsError(resultVal) And resultVal <> "" Then ws.Cells(rr, targetCol).value = resultVal
            End If
        End If
NextRow:
    Next rr
End Sub

Function SafeVLookup(findVal As String, rng As Range) As Variant
    On Error Resume Next
    SafeVLookup = Application.WorksheetFunction.VLOOKUP(findVal, rng, 3, False)
    If Err.Number <> 0 Then SafeVLookup = "": Err.Clear
    On Error GoTo 0
End Function

Function FindHeaderColFlexible(ws As Worksheet, candidates As Variant) As Long
    Dim i As Long, j As Long, h As String, c As String
    For i = 1 To 50
        h = NormalizeHeader(ws.Cells(1, i).value)
        For j = LBound(candidates) To UBound(candidates)
            c = NormalizeHeader(candidates(j))
            If h = c Or h Like c & "*" Or c Like h & "*" Then
                FindHeaderColFlexible = i: Exit Function
            End If
        Next j
    Next i
End Function

Function NormalizeHeader(v As Variant) As String
    Dim s As String: s = CStr(v)
    s = Replace(Replace(Replace(Replace(s, ".", ""), " ", ""), "_", ""), "-", "")
    NormalizeHeader = UCase(s)
End Function

Function IsBalanceSheet(sn As String, fName As String, fNum As String) As Boolean
    IsBalanceSheet = (sn = fName & "BAL" Or sn = fName & "BALANCE" Or sn = "BALANCE" & fNum)
End Function

Function IsJstBalanceSheet(sn As String, fName As String, fNum As String) As Boolean
    IsJstBalanceSheet = (sn = fName & "JSTBAL" Or sn = fName & "JST" Or sn = "JST" & fNum)
End Function

Function IsStockSheet(sn As String, fName As String, fNum As String) As Boolean
    IsStockSheet = (sn = fName & "STO" Or sn = fName & "STOCK" Or sn = "STOCK" & fNum)
End Function

Function NzNum(v As Variant) As Double
    If IsError(v) Or v = "" Then NzNum = 0 Else NzNum = Val(Replace(CStr(v), ",", ""))
End Function

