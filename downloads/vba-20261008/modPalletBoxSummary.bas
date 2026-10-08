Attribute VB_Name = "modPalletBoxSummary"
Sub 파렛트_박스정리()

    Dim wsSrc As Worksheet, wsOut As Worksheet, wbTarget As Workbook

    Dim lastRow As Long, lastCol As Long

    Dim i As Long, j As Long, outRow As Long

    Dim palletCols() As Long, palletNames() As String, nPallet As Long



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

    Dim alreadyOpen As Boolean

    alreadyOpen = False

    On Error Resume Next

    Set wbMOQ = Workbooks("한국단자 MOQ(20250522).xlsx")

    On Error GoTo 0

    If Not wbMOQ Is Nothing Then

        alreadyOpen = True

    Else

        Application.AskToUpdateLinks = False

        Set wbMOQ = Workbooks.Open(Filename:=moqPath, ReadOnly:=True, UpdateLinks:=0)

    End If



    ' MOQ 시트 잡기 (대아_MOQ 우선, 없으면 첫 시트)

    On Error Resume Next

    Set wsMOQ = wbMOQ.Worksheets("대아_MOQ")

    On Error GoTo 0

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

        If Not alreadyOpen Then wbMOQ.Close SaveChanges:=False

        GoTo CleanExit

    End If



    ' ===== MOQ를 Dictionary에 적재 =====

    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")

    Dim moqLast As Long, mArr As Variant

    moqLast = wsMOQ.Cells(wsMOQ.Rows.Count, moqNameCol).End(xlUp).Row

    mArr = wsMOQ.Range(wsMOQ.Cells(2, 1), wsMOQ.Cells(moqLast, moqBoxCol)).value

    For i = 1 To UBound(mArr, 1)

        Dim key As String

        key = UCase(Replace(Trim(CStr(mArr(i, moqNameCol))), " ", ""))

        If key <> "" And Not dict.exists(key) Then

            If IsNumeric(mArr(i, moqBoxCol)) Then dict(key) = CDbl(mArr(i, moqBoxCol))

        End If

    Next i



    ' MOQ 파일 닫기 (원래 안 열려 있었을 때만)

    If Not alreadyOpen Then wbMOQ.Close SaveChanges:=False

    Set wsMOQ = Nothing: Set wbMOQ = Nothing



    ' ===== 원본 파렛트 열 파악 (1열=품번, 2열=출고수량, sum 전까지 파렛트) =====

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    ReDim palletCols(1 To 60)

    ReDim palletNames(1 To 60)

    nPallet = 0

    For j = 3 To lastCol

        Dim hdr As String

        hdr = Trim(CStr(wsSrc.Cells(1, j).value))

        If hdr = "sum" Then Exit For

        If hdr <> "" Then

            nPallet = nPallet + 1

            palletCols(nPallet) = j

            If hdr = "대기" Then palletNames(nPallet) = "대기" Else palletNames(nPallet) = hdr & "번"

        End If

    Next j



    ' ===== 원본 데이터 배열로 읽기 =====

    Dim sArr As Variant

    sArr = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).value



    ' ===== 출력 시트 생성 (XLAM 아니라 작업 파일에 만들기) =====

    On Error Resume Next

    Application.DisplayAlerts = False

    wbTarget.Worksheets("파렛트정리").Delete

    Application.DisplayAlerts = True

    On Error GoTo 0

    Set wsOut = wbTarget.Worksheets.Add(After:=wsSrc)

    wsOut.Name = "파렛트정리"

    wsOut.Range("A1:D1").value = Array("품명", "박스당", "총 수량", "파렛트별 위치 / 수량")



    ' ===== 결과 계산 → 배열에 담기 =====

    Dim res() As Variant

    ReDim res(1 To lastRow, 1 To 4)

    outRow = 0

    For i = 2 To lastRow

        Dim nm As String, pb As Double, total As Double, detail As String, pbFound As Boolean

        nm = Trim(CStr(sArr(i, 1)))

        If nm = "" Then GoTo NextRow



        Dim lk As String

        lk = UCase(Replace(nm, " ", ""))

        If dict.exists(lk) Then

            pb = dict(lk): pbFound = True

        Else

            pb = 0: pbFound = False

        End If



        If IsNumeric(sArr(i, 2)) Then total = sArr(i, 2) Else total = 0   ' B열 = 총 수량



        detail = ""

        For j = 1 To nPallet

            Dim qty As Double

            If IsNumeric(sArr(i, palletCols(j))) Then qty = sArr(i, palletCols(j)) Else qty = 0

            If qty <> 0 Then

                Dim seg As String: seg = ""

                If pbFound And pb > 0 Then

                    Dim boxes As Long, remQty As Long

                    boxes = Int(qty / pb): remQty = qty - boxes * pb

                    If boxes > 0 Then seg = boxes & "박스"

                    If remQty > 0 Then

                        If seg <> "" Then seg = seg & " + "

                        seg = seg & "낱개" & remQty

                    End If

                    If seg = "" Then seg = qty & "개"

                Else

                    seg = qty & "개"

                End If

                If detail <> "" Then detail = detail & "   "

                detail = detail & "[" & palletNames(j) & "] " & seg

            End If

        Next j



        outRow = outRow + 1

        res(outRow, 1) = nm

        If pbFound Then res(outRow, 2) = pb Else res(outRow, 2) = "MOQ없음"

        res(outRow, 3) = total

        res(outRow, 4) = detail

NextRow:

    Next i



    If outRow > 0 Then wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(outRow + 1, 4)).value = res



    ' ===== 서식 =====

    With wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(outRow + 1, 4))

        .Font.Name = "Arial"

        .Borders.LineStyle = xlContinuous

        .Borders.Weight = xlThin

        .VerticalAlignment = xlCenter

    End With

    With wsOut.Range("A1:D1")

        .Interior.Color = RGB(0, 0, 0)

        .Font.Color = RGB(255, 255, 255)

        .Font.Bold = True

        .Font.Size = 12

        .HorizontalAlignment = xlCenter

        .Borders.Weight = xlMedium

    End With

    wsOut.Rows(1).RowHeight = 24



    If outRow > 0 Then

wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(outRow + 1, 4)).Sort Key1:=wsOut.Cells(2, 1), Order1:=xlAscending, header:=xlNo

        wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(outRow + 1, 1)).Font.Bold = True

        wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(outRow + 1, 1)).Font.Size = 12

        wsOut.Range(wsOut.Cells(2, 2), wsOut.Cells(outRow + 1, 3)).Font.Size = 11

        wsOut.Range(wsOut.Cells(2, 2), wsOut.Cells(outRow + 1, 3)).HorizontalAlignment = xlCenter

        wsOut.Range(wsOut.Cells(2, 2), wsOut.Cells(outRow + 1, 2)).NumberFormat = "#,##0"

        wsOut.Range(wsOut.Cells(2, 3), wsOut.Cells(outRow + 1, 3)).NumberFormat = "#,##0"

        wsOut.Range(wsOut.Cells(2, 4), wsOut.Cells(outRow + 1, 4)).Font.Size = 11

        Dim r As Long

        For r = 2 To outRow + 1

            If r Mod 2 = 0 Then wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 4)).Interior.Color = RGB(232, 232, 232)

            wsOut.Rows(r).RowHeight = 18

        Next r

    End If



    ' ===== 열 너비 / 틀 고정 =====

    wsOut.Columns("A").ColumnWidth = 20

    wsOut.Columns("B").ColumnWidth = 10

    wsOut.Columns("C").ColumnWidth = 12

    wsOut.Columns("D").ColumnWidth = 90

    wsOut.Activate

    wsOut.Range("A2").Select

    ActiveWindow.FreezePanes = True



    ' ===== 인쇄 설정 =====

    With wsOut.PageSetup

        .Orientation = xlLandscape

        .PaperSize = xlPaperA4

        .Zoom = False

        .FitToPagesWide = 1

        .FitToPagesTall = False

        .PrintTitleRows = "$1:$1"

        .LeftMargin = Application.InchesToPoints(0.3)

        .RightMargin = Application.InchesToPoints(0.3)

        .TopMargin = Application.InchesToPoints(0.4)

        .BottomMargin = Application.InchesToPoints(0.4)

    End With



    MsgBox "완료! " & outRow & "개 품목 정리됨", vbInformation



CleanExit:

    Application.ScreenUpdating = True

    Application.Calculation = xlCalculationAutomatic

End Sub
