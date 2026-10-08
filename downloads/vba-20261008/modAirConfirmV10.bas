Attribute VB_Name = "modAirConfirmV10"
Option Explicit
Private Const SH_TUE As String = "TUE_CONFIRM"
Private Const SH_THU As String = "THU_CONFIRM"
Private Const SH_F12 As String = "THAI_F1F2"
Private Const SH_F5 As String = "THAI_F5"
Private Const SH_CHECK As String = "CHECK"
Private mTuesday As Boolean
Public Sub BuildAirConfirm()
    에어_컨펌_만들기
End Sub
Public Sub BuildTuesdayConfirm()
    에어_컨펌_화요일
End Sub
Public Sub BuildThursdayConfirm()
    에어_컨펌_금요일
End Sub
Public Sub BuildFridayConfirm()
    에어_컨펌_금요일
End Sub
Public Sub 에어_컨펌_화요일()
    BuildAirMode True
End Sub
Public Sub 에어_컨펌_금요일()
    BuildAirMode False
End Sub
Public Sub 에어_컨펌_만들기()
    Dim choice As VbMsgBoxResult
    choice = MsgBox("예: 화요일 KET 또는 URGENT / 아니오: 금요일 REMAIN 전체", vbYesNoCancel + vbQuestion)
    If choice = vbCancel Then Exit Sub
    BuildAirMode (choice = vbYes)
End Sub
Private Sub BuildAirMode(ByVal tuesday As Boolean)
    mTuesday = tuesday
    Dim wb As Workbook
    Dim wsHyoju As Worksheet, wsKyungjin As Worksheet, wsMinji As Worksheet, wsBase As Worksheet
    Dim wsOut As Worksheet, wsF12 As Worksheet, wsF5 As Worksheet
    Dim hyojuHdr As Object, kyungjinHdr As Object, minjiHdr As Object, baseHdr As Object
    Dim confirmExact As Object, confirmByItem As Object
    Dim lastBase As Long, r As Long
    Dim outRow As Long, f12Row As Long, f5Row As Long
    Dim calcMode As XlCalculation
    Dim qtyHeader As String, baseLabel As String
    Dim currentStep As String
    calcMode = Application.Calculation
    Dim oldEvents As Boolean, oldScreen As Boolean
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo CleanFail
    currentStep = "Open workbook sheets"
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 109, , "No active workbook. Open the AIR confirmation workbook first."
    Set wsHyoju = OptionalPersonSheet(wb, "HYOJU")
    Set wsKyungjin = OptionalPersonSheet(wb, "KYUNGJIN")
    Set wsMinji = OptionalPersonSheet(wb, "MINJI")
    If TypeName(ActiveSheet) <> "Worksheet" Then Err.Raise vbObjectError + 112, , "요청 목록을 선택하세요."
    If wsHyoju Is Nothing And wsKyungjin Is Nothing And wsMinji Is Nothing Then Err.Raise vbObjectError + 115, , "효주·경진·민지 중 하나 이상의 담당자 시트가 필요합니다."
    Set wsBase = ActiveSheet
    If wsBase Is wsHyoju Or wsBase Is wsKyungjin Or wsBase Is wsMinji Then Err.Raise vbObjectError + 113, , "담당자 시트가 아닌 요청 목록을 선택하세요."
    Select Case UCase$(wsBase.Name)
        Case "AIR_CONFIRM", "TUE_CONFIRM", "THU_CONFIRM", "THAI_F1F2", "THAI_F5", "CHECK"
            Err.Raise vbObjectError + 114, , "결과 시트가 아닌 원본 요청 목록을 선택하세요."
    End Select
    qtyHeader = IIf(tuesday, "Purchase request", "Remain")
    baseLabel = wsBase.Name
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    currentStep = "Read headers"
    Set hyojuHdr = OptionalPersonHeaders(wsHyoju)
    Set kyungjinHdr = OptionalPersonHeaders(wsKyungjin)
    Set minjiHdr = OptionalPersonHeaders(wsMinji)
    Set baseHdr = GetHeaderMap(wsBase)
    ValidateHeaders baseHdr, wsBase.Name, Array("FAC", "ITEMCD", "QTY")
    If tuesday Then ValidateHeaders baseHdr, wsBase.Name, Array("MAKER", "NEEDETD")
    currentStep = "Prepare output sheets"
    Set wsOut = EnsureSheet(wb, "AIR_CONFIRM")
    Set wsF12 = EnsureSheet(wb, SH_F12)
    Set wsF5 = EnsureSheet(wb, SH_F5)
    ClearOutput wsOut
    ClearOutput wsF12
    ClearOutput wsF5
    WriteHeaders wsOut, Array("FAC", "Item cd", "Maker", "Part Name", qtyHeader, "STOCK", "Confirm", "Short Del.", "NEED ETD", "REMARK", "STATUS")
    WriteHeaders wsF12, Array("FAC", "Item cd", "Maker", "Part Name", qtyHeader, "STOCK", "Confirm", "SHORT", "NEED ETD")
    WriteHeaders wsF5, Array("FAC", "Item cd", "Maker", "Part Name", qtyHeader, "STOCK", "Confirm", "SHORT", "NEED ETD")
    ApplyConfirmHeaderFormats wsBase, baseHdr, wsHyoju, hyojuHdr, wsOut
    ApplyThaiHeaderFormats wsBase, baseHdr, wsF12
    ApplyThaiHeaderFormats wsBase, baseHdr, wsF5
    Set confirmExact = CreateObject("Scripting.Dictionary")
    Set confirmByItem = CreateObject("Scripting.Dictionary")
    currentStep = "Index confirmation sheets"
    IndexConfirmationSource wsHyoju, hyojuHdr, "H", confirmExact, confirmByItem
    IndexConfirmationSource wsKyungjin, kyungjinHdr, "K", confirmExact, confirmByItem
    IndexConfirmationSource wsMinji, minjiHdr, "M", confirmExact, confirmByItem
    outRow = 2: f12Row = 2: f5Row = 2
    currentStep = "Build output from base list"
    ProcessBaseList wsBase, baseHdr, baseLabel, wsHyoju, hyojuHdr, wsKyungjin, kyungjinHdr, wsMinji, minjiHdr, confirmExact, confirmByItem, wsOut, wsF12, wsF5, outRow, f12Row, f5Row
    currentStep = "Apply formats"
    ApplyConfirmDataFormats wsBase, baseHdr, wsOut, outRow - 1
    ApplyThaiFormatsFromOutput wsOut, wsF12, f12Row - 1
    ApplyThaiFormatsFromOutput wsOut, wsF5, f5Row - 1
    currentStep = "Sort outputs"
    SortOutput wsOut, outRow - 1, 11, 1, 3, 4
    SortOutput wsF12, f12Row - 1, 9, 1, 3, 4
    SortOutput wsF5, f5Row - 1, 9, 1, 3, 4
    FormatOutput wsOut, outRow - 1, 11, False
    FormatOutput wsF12, f12Row - 1, 9, False
    FormatOutput wsF5, f5Row - 1, 9, False
    ColorStatus wsOut, outRow - 1
    Application.Calculation = calcMode
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
Dim missingNames As String
    If wsHyoju Is Nothing Then missingNames = missingNames & "효주 "
    If wsKyungjin Is Nothing Then missingNames = missingNames & "경진 "
    If wsMinji Is Nothing Then missingNames = missingNames & "민지 "
    MsgBox IIf(tuesday, "화요일 KET + URGENT", "금요일 REMAIN 전체") & " 컨펌 생성 완료" & vbCrLf & "AIR_CONFIRM: " & (outRow - 2) & "행 / THAI_F1F2: " & (f12Row - 2) & "행 / THAI_F5: " & (f5Row - 2) & "행" & vbCrLf & "없는 담당자 시트: " & IIf(missingNames = "", "없음", missingNames) & vbCrLf & "확인사항은 AIR_CONFIRM 비고 열에 표시했습니다.", vbInformation
    Exit Sub
CleanFail:
    Application.Calculation = calcMode
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
MsgBox "Macro stopped." & vbCrLf & "Step: " & currentStep & vbCrLf & "Error " & CStr(Err.Number) & ": " & Err.Description, vbExclamation
End Sub
Public Sub CopyThaiF1F2()
    CopyOutputTable SH_F12
End Sub
Public Sub CopyThaiF5()
    CopyOutputTable SH_F5
End Sub
Private Function FindInputSheet(ByVal wb As Workbook, ByVal roleName As String) As Worksheet
    Dim aliases As Variant, ws As Worksheet, aliasValue As Variant
    Select Case UCase$(roleName)
        Case "HYOJU"
            aliases = Array(KoreanHyoju(), "HYOJU", "HYOJU_CONFIRM")
        Case "KYUNGJIN"
            aliases = Array(KoreanKyungjin(), "KYUNGJIN", "KYUNGJIN_CONFIRM")
        Case "MINJI"
            aliases = Array(KoreanMinji(), "MINJI", "MINJI_CONFIRM")
        Case "TUESDAY"
            aliases = Array(KoreanTuesdayAir(), "TUESDAY_AIR", "AIR_LIST", "AIR")
        Case "THURSDAY"
            aliases = Array(KoreanWednesdayRemain(), "WEDNESDAY_REMAIN", "REMAIN_WED", "REMAIN")
        Case Else
            Err.Raise vbObjectError + 110, , "Unknown sheet role: " & roleName
    End Select
    For Each ws In wb.Worksheets
        For Each aliasValue In aliases
            If NormalizeSheetName(ws.Name) = NormalizeSheetName(CStr(aliasValue)) Then
                Set FindInputSheet = ws
                Exit Function
            End If
        Next aliasValue
    Next ws
    Err.Raise vbObjectError + 111, , "Input sheet not found for " & roleName & ". Expected one of: " & Join(aliases, ", ")
End Function
Private Function FindOrCreatePersonSheet(ByVal wb As Workbook, ByVal roleName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = FindInputSheet(wb, roleName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        If UCase$(roleName) = "MINJI" Then
            ws.Name = KoreanMinji()
        Else
            ws.Name = roleName
        End If
        WriteHeaders ws, Array("FAC", "Item cd", "Maker", "Part Name", "Remain", "STOCK", "Confirm", "Short Del.", "NEED ETD", "REMARK")
    End If
    Set FindOrCreatePersonSheet = ws
End Function
Private Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureSheet = ws
End Function
Private Function NormalizeSheetName(ByVal valueIn As String) As String
    Dim s As String
    s = UCase$(Trim$(valueIn))
    s = Replace(s, " ", "")
    s = Replace(s, "-", "_")
    NormalizeSheetName = s
End Function
Private Function KoreanHyoju() As String
    KoreanHyoju = ChrW(54952) & ChrW(51452)
End Function
Private Function KoreanKyungjin() As String
    KoreanKyungjin = ChrW(44221) & ChrW(51652)
End Function
Private Function KoreanMinji() As String
    KoreanMinji = ChrW(48124) & ChrW(51648)
End Function
Private Function KoreanTuesdayAir() As String
    KoreanTuesdayAir = ChrW(54868) & ChrW(50836) & ChrW(51068) & "_AIR"
End Function
Private Function KoreanWednesdayRemain() As String
    KoreanWednesdayRemain = ChrW(49688) & ChrW(50836) & ChrW(51068) & "_REMAIN"
End Function
Private Sub IndexConfirmationSource(ByVal ws As Worksheet, ByVal hdr As Object, ByVal sourceId As String, ByVal exactIndex As Object, ByVal itemIndex As Object)
    Dim r As Long, lastRow As Long, fac As String, itemCd As String, exactKey As String, token As String
    If ws Is Nothing Then Exit Sub
    lastRow = LastDataRow(ws, hdr("ITEMCD"))
    For r = hdr("HEADERROW") + 1 To lastRow
        itemCd = NormalizeCode(CellValue(ws, r, hdr, "ITEMCD"))
        If itemCd <> "" Then
            fac = NormalizeCode(CellValue(ws, r, hdr, "FAC"))
            exactKey = fac & "|" & itemCd
            token = sourceId & "|" & CStr(r)
            If exactIndex.exists(exactKey) Then
                exactIndex(exactKey) = "DUP"
            Else
                exactIndex.Add exactKey, token
            End If
            If itemIndex.exists(itemCd) Then
                itemIndex(itemCd) = "DUP"
            Else
                itemIndex.Add itemCd, token
            End If
        End If
    Next r
End Sub
Private Sub ProcessBaseList(ByVal wsBase As Worksheet, ByVal baseHdr As Object, ByVal baseLabel As String, ByVal wsHyoju As Worksheet, ByVal hyojuHdr As Object, ByVal wsKyungjin As Worksheet, ByVal kyungjinHdr As Object, ByVal wsMinji As Worksheet, ByVal minjiHdr As Object, ByVal exactIndex As Object, ByVal itemIndex As Object, ByVal wsOut As Worksheet, ByVal wsF12 As Worksheet, ByVal wsF5 As Worksheet, ByRef outRow As Long, ByRef f12Row As Long, ByRef f5Row As Long)
    Dim lastBase As Long, r As Long, ownerRow As Long
    Dim fac As String, itemCd As String, exactKey As String, token As String, sourceId As String
    Dim statusText As String, issueText As String, sourceName As String
    Dim qtyVal As Variant, stockVal As Variant, confirmVal As Variant
    Dim shortVal As Variant, needVal As Variant, makerVal As Variant, partVal As Variant, remarkVal As Variant
    Dim wsOwner As Worksheet, ownerHdr As Object, tokenParts As Variant, hasOwner As Boolean
    lastBase = LastDataRow(wsBase, baseHdr("ITEMCD"))
    For r = baseHdr("HEADERROW") + 1 To lastBase
        itemCd = NormalizeCode(CellValue(wsBase, r, baseHdr, "ITEMCD"))
        If itemCd <> "" Then
            fac = NormalizeCode(CellValue(wsBase, r, baseHdr, "FAC"))
            exactKey = fac & "|" & itemCd
            token = ""
            issueText = ""
            hasOwner = False
            Set wsOwner = Nothing
            Set ownerHdr = Nothing
            If exactIndex.exists(exactKey) Then
                If CStr(exactIndex(exactKey)) = "DUP" Then
                    issueText = "담당자 자료 FAC + Item cd 중복: 자동 연결하지 않았습니다."
                Else
                    token = CStr(exactIndex(exactKey))
                End If
            End If
            If token <> "" Then
                tokenParts = Split(token, "|")
                sourceId = CStr(tokenParts(0))
                ownerRow = CLng(tokenParts(1))
                Select Case sourceId
                    Case "H": Set wsOwner = wsHyoju: Set ownerHdr = hyojuHdr
                    Case "K": Set wsOwner = wsKyungjin: Set ownerHdr = kyungjinHdr
                    Case "M": Set wsOwner = wsMinji: Set ownerHdr = minjiHdr
                End Select
                hasOwner = Not wsOwner Is Nothing
            End If
            makerVal = CellValue(wsBase, r, baseHdr, "MAKER")
            partVal = CellValue(wsBase, r, baseHdr, "PARTNAME")
            qtyVal = CellValue(wsBase, r, baseHdr, "QTY")
            stockVal = CellValue(wsBase, r, baseHdr, "STOCK")
            shortVal = CellValue(wsBase, r, baseHdr, "SHORTDEL")
            needVal = CellValue(wsBase, r, baseHdr, "NEEDETD")
            remarkVal = CellValue(wsBase, r, baseHdr, "REMARK")
            confirmVal = Empty
            If hasOwner Then
                makerVal = FirstValue(makerVal, CellValue(wsOwner, ownerRow, ownerHdr, "MAKER"))
                partVal = FirstValue(partVal, CellValue(wsOwner, ownerRow, ownerHdr, "PARTNAME"))
                stockVal = FirstValue(CellValue(wsOwner, ownerRow, ownerHdr, "STOCK"), stockVal)
                confirmVal = CellValue(wsOwner, ownerRow, ownerHdr, "CONFIRM")
                shortVal = FirstValue(shortVal, CellValue(wsOwner, ownerRow, ownerHdr, "SHORTDEL"))
                needVal = FirstValue(needVal, CellValue(wsOwner, ownerRow, ownerHdr, "NEEDETD"))
                remarkVal = FirstValue(remarkVal, CellValue(wsOwner, ownerRow, ownerHdr, "REMARK"))
                sourceName = wsOwner.Name
            Else
                sourceName = baseLabel
            End If
            If mTuesday Then
                If UCase$(Trim$(CStr(makerVal))) <> "KET" And InStr(1, CStr(needVal), "URGENT", vbTextCompare) = 0 Then GoTo NextBaseRow
            End If
            If Not hasOwner And issueText = "" Then issueText = "담당자 자료에 FAC + Item cd 일치 항목 없음"
            If fac <> "F1" And fac <> "F2" And fac <> "F5" Then issueText = issueText & " / 공장 확인 필요: " & fac
            If issueText <> "" Then remarkVal = CStr(remarkVal) & " / " & issueText
            statusText = EvaluateStatus(qtyVal, stockVal, confirmVal, needVal)
            If issueText <> "" Then statusText = "REVIEW"
            WriteConfirmRow wsOut, outRow, fac, itemCd, makerVal, partVal, qtyVal, stockVal, confirmVal, shortVal, needVal, remarkVal, statusText
            If True Then
                If UCase$(Trim$(CStr(fac))) = "F5" Then
                    WriteThaiRow wsF5, f5Row, fac, itemCd, makerVal, partVal, qtyVal, stockVal, confirmVal, shortVal, needVal, outRow
                    f5Row = f5Row + 1
                ElseIf UCase$(Trim$(CStr(fac))) = "F1" Or UCase$(Trim$(CStr(fac))) = "F2" Then
                    WriteThaiRow wsF12, f12Row, fac, itemCd, makerVal, partVal, qtyVal, stockVal, confirmVal, shortVal, needVal, outRow
                    f12Row = f12Row + 1
                Else
                wsOut.Cells(outRow, 10).Value = CStr(remarkVal) & " / 확인 필요"
                End If

            End If
            outRow = outRow + 1
        End If
NextBaseRow:
    Next r
End Sub
Private Sub ApplyConfirmHeaderFormats(ByVal wsBase As Worksheet, ByVal baseHdr As Object, ByVal wsOwner As Worksheet, ByVal ownerHdr As Object, ByVal wsOut As Worksheet)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "FAC", wsOut.Cells(1, 1)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "ITEMCD", wsOut.Cells(1, 2)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "MAKER", wsOut.Cells(1, 3)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "PARTNAME", wsOut.Cells(1, 4)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "QTY", wsOut.Cells(1, 5)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "STOCK", wsOut.Cells(1, 6)
    If Not wsOwner Is Nothing Then CopyFormatByKey wsOwner, CLng(ownerHdr("HEADERROW")), ownerHdr, "CONFIRM", wsOut.Cells(1, 7)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "SHORTDEL", wsOut.Cells(1, 8)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "NEEDETD", wsOut.Cells(1, 9)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "REMARK", wsOut.Cells(1, 10)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "NEEDETD", wsOut.Cells(1, 11)
    wsOut.Rows(1).RowHeight = wsBase.Rows(CLng(baseHdr("HEADERROW"))).RowHeight
    CopyColumnWidthByKey wsBase, baseHdr, "FAC", wsOut, 1
    CopyColumnWidthByKey wsBase, baseHdr, "ITEMCD", wsOut, 2
    CopyColumnWidthByKey wsBase, baseHdr, "MAKER", wsOut, 3
    CopyColumnWidthByKey wsBase, baseHdr, "PARTNAME", wsOut, 4
    CopyColumnWidthByKey wsBase, baseHdr, "QTY", wsOut, 5
    CopyColumnWidthByKey wsBase, baseHdr, "STOCK", wsOut, 6
    CopyColumnWidthByKey wsOwner, ownerHdr, "CONFIRM", wsOut, 7
    CopyColumnWidthByKey wsBase, baseHdr, "SHORTDEL", wsOut, 8
    CopyColumnWidthByKey wsBase, baseHdr, "NEEDETD", wsOut, 9
    CopyColumnWidthByKey wsBase, baseHdr, "REMARK", wsOut, 10
    wsOut.Columns(11).ColumnWidth = 14
End Sub
Private Sub ApplyConfirmRowFormats(ByVal wsBase As Worksheet, ByVal baseRow As Long, ByVal baseHdr As Object, ByVal wsOwner As Worksheet, ByVal ownerRow As Long, ByVal ownerHdr As Object, ByVal hasOwner As Boolean, ByVal wsOut As Worksheet, ByVal outRow As Long)
    CopyFormatByKey wsBase, baseRow, baseHdr, "FAC", wsOut.Cells(outRow, 1)
    CopyFormatByKey wsBase, baseRow, baseHdr, "ITEMCD", wsOut.Cells(outRow, 2)
    CopyFormatByKey wsBase, baseRow, baseHdr, "MAKER", wsOut.Cells(outRow, 3)
    CopyFormatByKey wsBase, baseRow, baseHdr, "PARTNAME", wsOut.Cells(outRow, 4)
    CopyFormatByKey wsBase, baseRow, baseHdr, "QTY", wsOut.Cells(outRow, 5)
    If hasOwner Then
        If ownerHdr.exists("STOCK") And Trim$(CStr(CellValue(wsOwner, ownerRow, ownerHdr, "STOCK"))) <> "" Then
            CopyFormatByKey wsOwner, ownerRow, ownerHdr, "STOCK", wsOut.Cells(outRow, 6)
        Else
            CopyFormatByKey wsBase, baseRow, baseHdr, "STOCK", wsOut.Cells(outRow, 6)
        End If
    Else
        CopyFormatByKey wsBase, baseRow, baseHdr, "STOCK", wsOut.Cells(outRow, 6)
    End If
    If hasOwner Then
        CopyFormatByKey wsOwner, ownerRow, ownerHdr, "CONFIRM", wsOut.Cells(outRow, 7)
    Else
        CopyFormatByKey wsBase, baseRow, baseHdr, "NEEDETD", wsOut.Cells(outRow, 7)
    End If
    CopyFormatByKey wsBase, baseRow, baseHdr, "SHORTDEL", wsOut.Cells(outRow, 8)
    CopyFormatByKey wsBase, baseRow, baseHdr, "NEEDETD", wsOut.Cells(outRow, 9)
    CopyFormatByKey wsBase, baseRow, baseHdr, "REMARK", wsOut.Cells(outRow, 10)
    CopyFormatByKey wsBase, baseRow, baseHdr, "NEEDETD", wsOut.Cells(outRow, 11)
    wsOut.Rows(outRow).RowHeight = wsBase.Rows(baseRow).RowHeight
End Sub
Private Sub ApplyThaiHeaderFormats(ByVal wsBase As Worksheet, ByVal baseHdr As Object, ByVal wsOut As Worksheet)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "FAC", wsOut.Cells(1, 1)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "ITEMCD", wsOut.Cells(1, 2)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "MAKER", wsOut.Cells(1, 3)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "PARTNAME", wsOut.Cells(1, 4)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "QTY", wsOut.Cells(1, 5)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "STOCK", wsOut.Cells(1, 6)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "NEEDETD", wsOut.Cells(1, 7)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "SHORTDEL", wsOut.Cells(1, 8)
    CopyFormatByKey wsBase, CLng(baseHdr("HEADERROW")), baseHdr, "NEEDETD", wsOut.Cells(1, 9)
    wsOut.Rows(1).RowHeight = wsBase.Rows(CLng(baseHdr("HEADERROW"))).RowHeight
    CopyColumnWidthByKey wsBase, baseHdr, "FAC", wsOut, 1
    CopyColumnWidthByKey wsBase, baseHdr, "ITEMCD", wsOut, 2
    CopyColumnWidthByKey wsBase, baseHdr, "MAKER", wsOut, 3
    CopyColumnWidthByKey wsBase, baseHdr, "PARTNAME", wsOut, 4
    CopyColumnWidthByKey wsBase, baseHdr, "QTY", wsOut, 5
    CopyColumnWidthByKey wsBase, baseHdr, "STOCK", wsOut, 6
    wsOut.Columns(7).ColumnWidth = 18
    CopyColumnWidthByKey wsBase, baseHdr, "SHORTDEL", wsOut, 8
    CopyColumnWidthByKey wsBase, baseHdr, "NEEDETD", wsOut, 9
End Sub
Private Sub ApplyThaiRowFormats(ByVal wsBase As Worksheet, ByVal baseRow As Long, ByVal baseHdr As Object, ByVal wsOwner As Worksheet, ByVal ownerRow As Long, ByVal ownerHdr As Object, ByVal hasOwner As Boolean, ByVal wsOut As Worksheet, ByVal outRow As Long)
    CopyFormatByKey wsBase, baseRow, baseHdr, "FAC", wsOut.Cells(outRow, 1)
    CopyFormatByKey wsBase, baseRow, baseHdr, "ITEMCD", wsOut.Cells(outRow, 2)
    CopyFormatByKey wsBase, baseRow, baseHdr, "MAKER", wsOut.Cells(outRow, 3)
    CopyFormatByKey wsBase, baseRow, baseHdr, "PARTNAME", wsOut.Cells(outRow, 4)
    CopyFormatByKey wsBase, baseRow, baseHdr, "QTY", wsOut.Cells(outRow, 5)
    CopyFormatByKey wsBase, baseRow, baseHdr, "STOCK", wsOut.Cells(outRow, 6)
    If hasOwner Then
        CopyFormatByKey wsOwner, ownerRow, ownerHdr, "CONFIRM", wsOut.Cells(outRow, 7)
    Else
        CopyFormatByKey wsBase, baseRow, baseHdr, "NEEDETD", wsOut.Cells(outRow, 7)
    End If
    CopyFormatByKey wsBase, baseRow, baseHdr, "SHORTDEL", wsOut.Cells(outRow, 8)
    CopyFormatByKey wsBase, baseRow, baseHdr, "NEEDETD", wsOut.Cells(outRow, 9)
    wsOut.Rows(outRow).RowHeight = wsBase.Rows(baseRow).RowHeight
End Sub
Private Sub ApplyConfirmDataFormats(ByVal wsBase As Worksheet, ByVal baseHdr As Object, ByVal wsOut As Worksheet, ByVal lastOutRow As Long)
    Dim sourceFirstRow As Long, rowCount As Long, r As Long
    If lastOutRow < 2 Then Exit Sub
    sourceFirstRow = CLng(baseHdr("HEADERROW")) + 1
    rowCount = lastOutRow - 1
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "FAC", wsOut, 2, 1, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "ITEMCD", wsOut, 2, 2, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "MAKER", wsOut, 2, 3, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "PARTNAME", wsOut, 2, 4, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "QTY", wsOut, 2, 5, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "STOCK", wsOut, 2, 6, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "STOCK", wsOut, 2, 7, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "SHORTDEL", wsOut, 2, 8, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "NEEDETD", wsOut, 2, 9, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "REMARK", wsOut, 2, 10, rowCount
    CopyFormatColumnBlock wsBase, sourceFirstRow, baseHdr, "REMARK", wsOut, 2, 11, rowCount
    wsOut.Range(wsOut.Cells(2, 7), wsOut.Cells(lastOutRow, 7)).NumberFormat = "dd-mmm"
    For r = 2 To lastOutRow
        wsOut.Rows(r).RowHeight = wsBase.Rows(sourceFirstRow + r - 2).RowHeight
    Next r
End Sub
Private Sub CopyFormatColumnBlock(ByVal wsSource As Worksheet, ByVal sourceFirstRow As Long, ByVal hdr As Object, ByVal key As String, ByVal wsDestination As Worksheet, ByVal destinationFirstRow As Long, ByVal destinationColumn As Long, ByVal rowCount As Long)
    Dim sourceRange As Range, destinationRange As Range
    If rowCount <= 0 Or Not hdr.exists(key) Then Exit Sub
    Set sourceRange = wsSource.Range(wsSource.Cells(sourceFirstRow, CLng(hdr(key))), wsSource.Cells(sourceFirstRow + rowCount - 1, CLng(hdr(key))))
    Set destinationRange = wsDestination.Range(wsDestination.Cells(destinationFirstRow, destinationColumn), wsDestination.Cells(destinationFirstRow + rowCount - 1, destinationColumn))
    sourceRange.Copy
    destinationRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub
Private Sub ApplyThaiFormatsFromOutput(ByVal wsOut As Worksheet, ByVal wsThai As Worksheet, ByVal lastThaiRow As Long)
    Dim r As Long, sourceOutRow As Long
    If lastThaiRow < 2 Then Exit Sub
    For r = 2 To lastThaiRow
        sourceOutRow = CLng(NumericValue(wsThai.Cells(r, 10).value))
        If sourceOutRow >= 2 Then
            wsOut.Range(wsOut.Cells(sourceOutRow, 1), wsOut.Cells(sourceOutRow, 9)).Copy
            wsThai.Range(wsThai.Cells(r, 1), wsThai.Cells(r, 9)).PasteSpecial Paste:=xlPasteFormats
            wsThai.Rows(r).RowHeight = wsOut.Rows(sourceOutRow).RowHeight
        End If
    Next r
    wsThai.Range(wsThai.Cells(2, 10), wsThai.Cells(lastThaiRow, 10)).ClearContents
    Application.CutCopyMode = False
End Sub
Private Sub CopyFormatByKey(ByVal wsSource As Worksheet, ByVal sourceRow As Long, ByVal hdr As Object, ByVal key As String, ByVal destination As Range)
    If Not hdr.exists(key) Then Exit Sub
    wsSource.Cells(sourceRow, CLng(hdr(key))).Copy
    destination.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub
Private Sub CopyColumnWidthByKey(ByVal wsSource As Worksheet, ByVal hdr As Object, ByVal key As String, ByVal wsDestination As Worksheet, ByVal destinationColumn As Long)
    If Not hdr.exists(key) Then Exit Sub
    wsDestination.Columns(destinationColumn).ColumnWidth = wsSource.Columns(CLng(hdr(key))).ColumnWidth
End Sub
Private Sub CopyOutputTable(ByVal sheetName As String)
    Dim ws As Worksheet, lastRow As Long, lastCol As Long
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then
        MsgBox "No rows to copy in " & sheetName, vbExclamation
        Exit Sub
    End If
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy
    MsgBox sheetName & " copied. Paste it into Outlook.", vbInformation
End Sub
Private Function GetHeaderMap(ByVal ws As Worksheet) As Object
    Dim m As Object, r As Long, c As Long, lastCol As Long, key As String
    Set m = CreateObject("Scripting.Dictionary")
    For r = 1 To 30
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastCol
            key = HeaderKey(CStr(ws.Cells(r, c).value))
            If key <> "" Then
                If Not m.exists(key) Then m.Add key, c
            End If
        Next c
        If m.exists("ITEMCD") Then
            m.Add "HEADERROW", r
            Set GetHeaderMap = m
            Exit Function
        End If
        m.RemoveAll
    Next r
    Err.Raise vbObjectError + 101, , "Header row not found in " & ws.Name
End Function
Private Function HeaderKey(ByVal rawText As String) As String
    Dim s As String
    s = UCase$(Trim$(rawText))
    s = Replace(s, " ", "")
    s = Replace(s, ".", "")
    s = Replace(s, "_", "")
    s = Replace(s, "-", "")
    s = Replace(s, "/", "")
    s = Replace(s, "'", "")
    Select Case s
        Case "FAC", "FACTORY": HeaderKey = "FAC"
        Case "ITEMCD", "ITEMCODE", "ITEMNO": HeaderKey = "ITEMCD"
        Case "MAKER", "SUPPLIER": HeaderKey = "MAKER"
        Case "PARTNAME", "ITEMNM", "ITEMNAME": HeaderKey = "PARTNAME"
        Case "REMAIN", "PURCHASEREQUEST", "REQUESTQTY", "QTY": HeaderKey = "QTY"
        Case "STOCK": HeaderKey = "STOCK"
        Case "CONFIRM", "CONFIRMED": HeaderKey = "CONFIRM"
        Case "SHORTDEL", "SHORT", "SHORTDELIVERY": HeaderKey = "SHORTDEL"
        Case "NEEDETD", "REQUESTETD": HeaderKey = "NEEDETD"
        Case "REMARK", "NOTE": HeaderKey = "REMARK"
        Case Else: HeaderKey = ""
    End Select
End Function
Private Sub ValidateHeaders(ByVal m As Object, ByVal sheetName As String, ByVal required As Variant)
    Dim i As Long
    For i = LBound(required) To UBound(required)
        If Not m.exists(CStr(required(i))) Then
            Err.Raise vbObjectError + 102, , "Missing header in " & sheetName & ": " & CStr(required(i))
        End If
    Next i
End Sub
Private Function CellValue(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal m As Object, ByVal key As String) As Variant
    If m.exists(key) Then
        CellValue = ws.Cells(rowNum, CLng(m(key))).value
    Else
        CellValue = Empty
    End If
End Function
Private Function FirstValue(ByVal primaryValue As Variant, ByVal fallbackValue As Variant) As Variant
    If Trim$(CStr(primaryValue)) <> "" Then
        FirstValue = primaryValue
    Else
        FirstValue = fallbackValue
    End If
End Function
Private Function NormalizeCode(ByVal valueIn As Variant) As String
    NormalizeCode = UCase$(Replace(Trim$(CStr(valueIn)), " ", ""))
End Function
Private Function LastDataRow(ByVal ws As Worksheet, ByVal colNum As Long) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
End Function
Private Function NumericValue(ByVal valueIn As Variant) As Double
    Dim s As String
    If IsNumeric(valueIn) Then
        NumericValue = CDbl(valueIn)
    Else
        s = Replace(Replace(Trim$(CStr(valueIn)), ",", ""), " ", "")
        If s = "" Then
            NumericValue = 0
        ElseIf IsNumeric(s) Then
            NumericValue = CDbl(s)
        Else
            NumericValue = 0
        End If
    End If
End Function
Private Function EvaluateStatus(ByVal qtyVal As Variant, ByVal stockVal As Variant, ByVal confirmVal As Variant, ByVal needVal As Variant) As String
    Dim qty As Double, stockQty As Double, confirmText As String
    qty = NumericValue(qtyVal)
    stockQty = NumericValue(stockVal)
    confirmText = UCase$(Trim$(CStr(confirmVal)))
    If qty <= 0 Then EvaluateStatus = "REVIEW": Exit Function
    If stockQty >= qty Then EvaluateStatus = "OK": Exit Function
    If confirmText = "" Then EvaluateStatus = "CHECKING": Exit Function
    EvaluateStatus = "OK"
End Function
Private Function ParseConfirmQty(ByVal textIn As String, ByRef qtyOut As Double) As Boolean
    Dim re As Object, matches As Object, token As String
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\s*([0-9]+(?:\.[0-9]+)?)\s*([Kk]?)\b"
    re.Global = False
    If re.Test(textIn) Then
        Set matches = re.Execute(textIn)
        token = matches(0).SubMatches(0)
        qtyOut = CDbl(token)
        If UCase$(matches(0).SubMatches(1)) = "K" Then qtyOut = qtyOut * 1000
        ParseConfirmQty = True
    End If
End Function
Private Function ParseFlexibleDate(ByVal valueIn As Variant, ByRef dateOut As Date) As Boolean
    Dim s As String, re As Object, matches As Object
    Dim d As Long, m As Long, y As Long, monText As String
    On Error GoTo DateFail
    If IsDate(valueIn) Then
        dateOut = CDate(valueIn)
        ParseFlexibleDate = True
        Exit Function
    End If
    s = Trim$(CStr(valueIn))
    If s = "" Or UCase$(s) = "URGENT" Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "([0-9]{1,2})[-/ ]([A-Za-z]{3}|[0-9]{1,2})(?:[-/ ]([0-9]{2,4}))?"
    re.Global = False
    If Not re.Test(s) Then Exit Function
    Set matches = re.Execute(s)
    d = CLng(matches(0).SubMatches(0))
    monText = UCase$(matches(0).SubMatches(1))
    If IsNumeric(monText) Then m = CLng(monText) Else m = MonthNo(monText)
    If matches(0).SubMatches(2) = "" Then
        y = Year(Date)
    Else
        y = CLng(matches(0).SubMatches(2))
        If y < 100 Then y = 2000 + y
    End If
    dateOut = DateSerial(y, m, d)
    ParseFlexibleDate = True
    Exit Function
DateFail:
    ParseFlexibleDate = False
End Function
Private Function MonthNo(ByVal monText As String) As Long
    Select Case Left$(monText, 3)
        Case "JAN": MonthNo = 1
        Case "FEB": MonthNo = 2
        Case "MAR": MonthNo = 3
        Case "APR": MonthNo = 4
        Case "MAY": MonthNo = 5
        Case "JUN": MonthNo = 6
        Case "JUL": MonthNo = 7
        Case "AUG": MonthNo = 8
        Case "SEP": MonthNo = 9
        Case "OCT": MonthNo = 10
        Case "NOV": MonthNo = 11
        Case "DEC": MonthNo = 12
        Case Else: Err.Raise vbObjectError + 103, , "Unknown month: " & monText
    End Select
End Function
Private Sub ClearOutput(ByVal ws As Worksheet)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.UsedRange.Clear
End Sub
Private Sub WriteHeaders(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = headers(i)
    Next i
End Sub
Private Sub WriteConfirmRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal fac As Variant, ByVal itemCd As Variant, ByVal makerVal As Variant, ByVal partVal As Variant, ByVal qtyVal As Variant, ByVal stockVal As Variant, ByVal confirmVal As Variant, ByVal shortVal As Variant, ByVal needVal As Variant, ByVal remarkVal As Variant, ByVal statusText As String)
    Dim rowData(1 To 1, 1 To 11) As Variant
    rowData(1, 1) = fac
    rowData(1, 2) = itemCd
    rowData(1, 3) = makerVal
    rowData(1, 4) = partVal
    rowData(1, 5) = qtyVal
    rowData(1, 6) = NumericValue(stockVal)
    rowData(1, 7) = confirmVal
    rowData(1, 8) = shortVal
    rowData(1, 9) = needVal
    rowData(1, 10) = remarkVal
    rowData(1, 11) = statusText
    ws.Cells(rowNum, 7).NumberFormat = "@"
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 11)).value = rowData
End Sub
Private Sub WriteThaiRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal fac As Variant, ByVal itemCd As Variant, ByVal makerVal As Variant, ByVal partVal As Variant, ByVal qtyVal As Variant, ByVal stockVal As Variant, ByVal confirmVal As Variant, ByVal shortVal As Variant, ByVal needVal As Variant, ByVal sourceOutRow As Long)
    Dim rowData(1 To 1, 1 To 10) As Variant
    rowData(1, 1) = fac
    rowData(1, 2) = itemCd
    rowData(1, 3) = makerVal
    rowData(1, 4) = partVal
    rowData(1, 5) = qtyVal
    rowData(1, 6) = NumericValue(stockVal)
    rowData(1, 7) = confirmVal
    rowData(1, 8) = shortVal
    rowData(1, 9) = needVal
    rowData(1, 10) = sourceOutRow
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 10)).value = rowData
End Sub
Private Sub FormatOutput(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long, ByVal useStandardStyle As Boolean)
    Dim rng As Range
    If lastRow < 1 Then lastRow = 1
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    If useStandardStyle Then
        With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter
        End With
        With rng.Borders
            .LineStyle = xlContinuous
            .Color = RGB(166, 166, 166)
            .Weight = xlThin
        End With
        rng.VerticalAlignment = xlCenter
        rng.WrapText = True
    End If
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Columns.AutoFit
    If lastCol = 11 Then
        ws.Columns(4).ColumnWidth = 48
        If lastRow >= 2 Then ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, 6)).NumberFormat = "#,##0"
    Else
        ws.Columns(4).ColumnWidth = 48
        If lastCol >= 6 And lastRow >= 2 Then ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, 6)).NumberFormat = "#,##0"
    End If
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub
Private Sub ColorStatus(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long, s As String
    For r = 2 To lastRow
        s = UCase$(Trim$(CStr(ws.Cells(r, 11).value)))
        Select Case s
            Case "OK": ws.Cells(r, 11).Interior.Color = RGB(198, 239, 206)
            Case "CHECKING": ws.Cells(r, 11).Interior.Color = RGB(255, 235, 156)
            Case Else: ws.Cells(r, 11).Interior.Color = RGB(255, 235, 156)
        End Select
    Next r
End Sub
Private Sub SortOutput(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long, ByVal facCol As Long, ByVal makerCol As Long, ByVal partCol As Long)
    If lastRow < 3 Then Exit Sub
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(2, facCol), ws.Cells(lastRow, facCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(2, makerCol), ws.Cells(lastRow, makerCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(2, partCol), ws.Cells(lastRow, partCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Function OptionalPersonSheet(ByVal wb As Workbook, ByVal role As String) As Worksheet
    On Error Resume Next
    Set OptionalPersonSheet = FindInputSheet(wb, role)
    On Error GoTo 0
End Function
Private Function OptionalPersonHeaders(ByVal ws As Worksheet) As Object
    Dim hdr As Object
    If ws Is Nothing Then
        Set OptionalPersonHeaders = CreateObject("Scripting.Dictionary")
    Else
        Set hdr = GetHeaderMap(ws)
        ValidateHeaders hdr, ws.Name, Array("FAC", "ITEMCD")
        Set OptionalPersonHeaders = hdr
    End If
End Function
