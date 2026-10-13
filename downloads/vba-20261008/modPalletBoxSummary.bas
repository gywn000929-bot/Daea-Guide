Attribute VB_Name = "modPalletBoxSummary"
Sub 파렛트_박스정리()
    Dim wsSrc As Worksheet, wsOut As Worksheet, wbTarget As Workbook
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, outRow As Long
    Dim palletCols() As Long, palletNames() As String, nPallet As Long

    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldLinks As Boolean
    oldCalc = Application.Calculation: oldScreen = Application.ScreenUpdating: oldLinks = Application.AskToUpdateLinks
    On Error GoTo Failed
    If TypeName(ActiveSheet) <> "Worksheet" Then Err.Raise vbObjectError + 800, , "원본 파렛트 시트를 선택하세요."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsSrc = ActiveSheet              ' 원본(파렛트) 시트를 선택한 상태에서 실행
    Set wbTarget = wsSrc.Parent          ' 결과를 만들 실제 작업 파일 (XLAM 아님)

    ' ===== MOQ 외부 파일 경로 =====
    Dim moqPath As String
    moqPath = Environ("USERPROFILE") & "\OneDrive - Dae-A Electronics (Thailand) Company Limited\바탕 화면\한국단자 MOQ(20250522).xlsx"

    If Dir(moqPath) = "" Then
        MsgBox "MOQ 파일을 찾을 수 없습니다:" & vbCrLf & moqPath, vbExclamation
        GoTo CleanExit
    End If

    ' ===== MOQ 파일 열기 (이미 열려 있으면 그걸 사용) =====
    Dim wbMOQ As Workbook, wsMOQ As Worksheet
    Dim AlreadyOpen As Boolean
    AlreadyOpen = False
    On Error Resume Next
    Set wbMOQ = Workbooks("한국단자 MOQ(20250522).xlsx")
    On Error GoTo Failed
    If Not wbMOQ Is Nothing Then
        AlreadyOpen = True
    Else
        Application.AskToUpdateLinks = False
        Set wbMOQ = Workbooks.Open(Filename:=moqPath, ReadOnly:=True, UpdateLinks:=0)
    End If

    ' MOQ 시트 잡기 (대아_MOQ 우선, 없으면 첫 시트)
    On Error Resume Next
    Set wsMOQ = wbMOQ.Worksheets("대아_MOQ")
    On Error GoTo Failed
    If wsMOQ Is Nothing Then Set wsMOQ = wbMOQ.Worksheets(1)

    ' ===== MOQ 헤더에서 ITEM NM / BOX 열 자동 인식 =====
    Dim moqNameCol As Long, moqBoxCol As Long, moqLastCol As Long, c As Long
    moqLastCol = wsMOQ.Cells(1, wsMOQ.Columns.Count).End(xlToLeft).Column
    For c = 1 To moqLastCol
        Dim hh As String
        hh = UCase(Replace(Trim(CStr(wsMOQ.Cells(1, c).value)), " ", ""))
        If hh = "ITEMNM" Then moqNameCol = c
        If hh = "BOX" Then moqBoxCol = c
    Next c
    If moqNameCol = 0 Or moqBoxCol = 0 Then
        MsgBox "MOQ 파일에서 'ITEM NM' / 'BOX' 헤더를 못 찾음.", vbExclamation
        If Not AlreadyOpen Then wbMOQ.Close SaveChanges:=False
        GoTo CleanExit
    End If

    ' ===== MOQ를 Dictionary에 적재 =====
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim moqLast As Long, mArr As Variant
    moqLast = wsMOQ.Cells(wsMOQ.Rows.Count, moqNameCol).End(xlUp).Row
    If moqLast < 2 Then Err.Raise vbObjectError + 801, , "MOQ 상세 데이터가 없습니다."
    mArr = wsMOQ.Range(wsMOQ.Cells(2, 1), wsMOQ.Cells(moqLast, Application.Max(moqBoxCol, moqNameCol))).value
    For i = 1 To UBound(mArr, 1)
        Dim key As String
        key = UCase(Replace(Trim(CStr(mArr(i, moqNameCol))), " ", ""))
        If key <> "" And Not dict.exists(key) Then
            If IsNumeric(mArr(i, moqBoxCol)) Then dict(key) = CDbl(mArr(i, moqBoxCol))
        End If
    Next i

    ' MOQ 파일 닫기 (원래 안 열려 있었을 때만)
    If Not AlreadyOpen Then wbMOQ.Close SaveChanges:=False
    Set wsMOQ = Nothing: Set wbMOQ = Nothing

    ' ===== 원본 파렛트 열 파악 (1열=품번, 2열=출고수량, sum 전까지 파렛트) =====
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    ReDim palletCols(1 To Application.Max(1, lastCol))
    ReDim palletNames(1 To Application.Max(1, lastCol))
    nPallet = 0
    For j = 3 To lastCol
        Dim hdr As String
        hdr = Trim(CStr(wsSrc.Cells(1, j).value))
        If LCase$(hdr) = "sum" Then Exit For
        If hdr <> "" Then
            nPallet = nPallet + 1
            palletCols(nPallet) = j
            If hdr = "대기" Then palletNames(nPallet) = "대기" Else palletNames(nPallet) = hdr & "번"
        End If
    Next j

    ' ===== 원본 데이터 배열로 읽기 =====
    Dim sArr As Variant
    sArr = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).value

    ' 한 결과 시트 안에서 파렛트마다 A4 한 페이지
    Dim qty As Double, nm As String, pb As Double, seg As String
    Dim boxes As Double, remQty As Double, created As Long, sumQty As Double
    Dim sheetName As String, suffix As Long, existing As Object, exists As Boolean
    Dim topRow As Long, firstData As Long, nItems As Long, bodyHeight As Double
    Dim sourceLabel As String, resultBase As String, totalBoxes As Double, totalEA As Double
    sourceLabel = wsSrc.Name
    If LCase$(Left$(sourceLabel, 5)) = "sheet" Then
        sourceLabel = wbTarget.Name
        If InStrRev(sourceLabel, ".") > 0 Then sourceLabel = Left$(sourceLabel, InStrRev(sourceLabel, ".") - 1)
    End If
    resultBase = Left$(sourceLabel, 20) & "_파레트 정리"
    sheetName = resultBase
    Do
        exists = False
        For Each existing In wbTarget.Sheets
            If StrComp(existing.Name, sheetName, vbTextCompare) = 0 Then exists = True: Exit For
        Next existing
        If Not exists Then Exit Do
        suffix = suffix + 1: sheetName = Left$(resultBase, 27) & "_" & suffix
    Loop
    Set wsOut = wbTarget.Worksheets.Add(After:=wbTarget.Sheets(wbTarget.Sheets.Count))
    wsOut.Name = sheetName
    wsOut.Columns("A").ColumnWidth = 6.7
    wsOut.Columns("B").ColumnWidth = 26.7
    wsOut.Columns("C").ColumnWidth = 14.7
    wsOut.Columns("D").ColumnWidth = 13
    wsOut.Columns("E").ColumnWidth = 18.7
    wsOut.Columns("B").NumberFormat = "@"
    outRow = 1
    For j = 1 To nPallet
        nItems = 0
        For i = 2 To lastRow
            If IsNumeric(sArr(i, palletCols(j))) Then
                If CDbl(sArr(i, palletCols(j))) <> 0 And Trim$(CStr(sArr(i, 1))) <> "" Then nItems = nItems + 1
            End If
        Next i
        If nItems = 0 Then GoTo NextPallet
        topRow = outRow
        If created > 0 Then wsOut.HPageBreaks.Add Before:=wsOut.Cells(topRow, 1)
        wsOut.Range(wsOut.Cells(topRow, 1), wsOut.Cells(topRow, 5)).Merge
        wsOut.Cells(topRow, 1).Value2 = "PALLET " & Replace(palletNames(j), "번", "")
        wsOut.Range(wsOut.Cells(topRow + 1, 1), wsOut.Cells(topRow + 1, 5)).Merge
        wsOut.Cells(topRow + 1, 1).Value2 = sourceLabel & " / KET        " & Format(Date, "dd-mmm-yyyy") & "        Hyoju"
        wsOut.Range(wsOut.Cells(topRow + 2, 1), wsOut.Cells(topRow + 2, 5)).Value = Array("No.", "ITEM NM", "Pallet Q'ty", "Box Q'ty", "Box / EA")
        outRow = topRow + 3: firstData = outRow: sumQty = 0: totalBoxes = 0: totalEA = 0
        For i = 2 To lastRow
            nm = Trim$(CStr(sArr(i, 1)))
            If nm = "" Then GoTo NextItem
            If Not IsNumeric(sArr(i, palletCols(j))) Then GoTo NextItem
            qty = CDbl(sArr(i, palletCols(j)))
            If qty = 0 Then GoTo NextItem
            pb = 0: key = UCase$(Replace(nm, " ", ""))
            If dict.Exists(key) Then pb = dict(key)
            seg = PalletQtyText(qty) & "EA"
            boxes = 0: remQty = qty
            If pb > 0 And qty > 0 Then
                boxes = Fix(qty / pb): remQty = qty - boxes * pb: seg = ""
                If boxes > 0 Then seg = Format(boxes, "#,##0") & "B"
                If remQty > 0 Then
                    If seg <> "" Then seg = seg & " + "
                    seg = seg & PalletQtyText(remQty) & "EA"
                End If
            End If
            wsOut.Cells(outRow, 1).Value2 = outRow - firstData + 1
            wsOut.Cells(outRow, 2).Value2 = nm
            If pb > 0 Then wsOut.Cells(outRow, 4).Value2 = pb Else wsOut.Cells(outRow, 4).Value2 = "MOQ없음"
            wsOut.Cells(outRow, 3).Value2 = qty
            wsOut.Cells(outRow, 5).Value2 = seg
            outRow = outRow + 1: sumQty = sumQty + qty
            totalBoxes = totalBoxes + boxes: totalEA = totalEA + remQty
NextItem:
        Next i
        If nItems > 1 Then wsOut.Range(wsOut.Cells(firstData, 1), wsOut.Cells(outRow - 1, 5)).Sort Key1:=wsOut.Cells(firstData, 2), Order1:=xlAscending, Header:=xlNo
        Dim nr As Long
        For nr = firstData To outRow - 1
            wsOut.Cells(nr, 1).Value2 = nr - firstData + 1
        Next nr
        wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 2)).Merge
        wsOut.Cells(outRow, 1).Value2 = "Total (" & nItems & " items)"
        wsOut.Range(wsOut.Cells(outRow, 4), wsOut.Cells(outRow, 5)).Merge
        wsOut.Cells(outRow, 4).Value2 = Format(totalBoxes, "#,##0") & "B" & IIf(totalEA <> 0, " + " & PalletQtyText(totalEA) & "EA", "")
        wsOut.Cells(outRow, 3).Value2 = sumQty
        ' 각 블록 높이를 제한하여 수동 페이지 구분 사이에 한 장으로 맞춤
        bodyHeight = Application.Min(30, 600 / (nItems + 1))
        With wsOut.Range(wsOut.Cells(topRow, 1), wsOut.Cells(outRow, 5))
            .Font.Name = "Arial"
            .Font.Size = Application.Min(13, Application.Max(1, bodyHeight * 0.52))
            .Font.Color = vbBlack
            .Interior.Pattern = xlNone
            .VerticalAlignment = xlCenter
            .WrapText = False
            .RowHeight = bodyHeight
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        With wsOut.Range(wsOut.Cells(topRow, 1), wsOut.Cells(topRow, 5))
            .Font.Bold = True
            .Font.Size = 28
            .Interior.Color = RGB(46, 92, 138)
            .Font.Color = vbWhite
            .RowHeight = 44
            .HorizontalAlignment = xlCenter
        End With
        With wsOut.Range(wsOut.Cells(topRow + 1, 1), wsOut.Cells(topRow + 1, 5))
            .Font.Size = 13
            .HorizontalAlignment = xlCenter
            .RowHeight = 24
        End With
        With wsOut.Range(wsOut.Cells(topRow + 2, 1), wsOut.Cells(topRow + 2, 5))
            .Font.Size = 13
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
            .HorizontalAlignment = xlCenter
            .RowHeight = 24
        End With
        wsOut.Rows(outRow).Font.Bold = True
        wsOut.Range(wsOut.Cells(firstData, 1), wsOut.Cells(outRow, 5)).ShrinkToFit = True
        wsOut.Range(wsOut.Cells(firstData, 3), wsOut.Cells(outRow, 4)).NumberFormat = "#,##0"
        wsOut.Range(wsOut.Cells(firstData, 3), wsOut.Cells(outRow, 5)).HorizontalAlignment = xlCenter
        created = created + 1
        outRow = outRow + 1
NextPallet:
    Next j
    If created = 0 Then MsgBox "출력할 파렛트 수량이 없습니다.", vbInformation: GoTo CleanExit
    With wsOut.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = 90
        .BlackAndWhite = True
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .PrintArea = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(outRow - 1, 5)).Address
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.2)
        .CenterHorizontally = True
        .CenterFooter = "&P / &N"
    End With
    wsOut.Activate
    MsgBox "완료! 결과 시트 하나에 " & created & "개 파렛트를 정리했습니다. 파렛트마다 A4 한 장씩 인쇄됩니다.", vbInformation
CleanExit:
    On Error Resume Next
    If Not wbMOQ Is Nothing Then
        If Not AlreadyOpen Then wbMOQ.Close SaveChanges:=False
    End If
    Application.ScreenUpdating = oldScreen
    Application.Calculation = oldCalc
    Application.AskToUpdateLinks = oldLinks
    Exit Sub
Failed:
    MsgBox "파렛트 정리 오류: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
Private Function PalletQtyText(ByVal value As Double) As String
    If value = Fix(value) Then
        PalletQtyText = Format(value, "#,##0")
    Else
        PalletQtyText = Format(value, "#,##0.########")
    End If
End Function
