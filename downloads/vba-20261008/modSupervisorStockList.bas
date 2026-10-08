Attribute VB_Name = "modSupervisorStockList"
Option Explicit

Public Sub 반장님_재고리스트_만들기()

    Dim src As Worksheet, dest As Worksheet
    Dim wb As Workbook
    Dim dict As Object

    Dim facCol As Long, nameCol As Long
    Dim qtyCol As Long, codeCol As Long
    Dim lastRow As Long, r As Long
    Dim outRow As Long, lastOutputRow As Long

    Dim fac As String, itemName As String, itemCode As String
    Dim qty As Double
    Dim key As String, k As Variant, data As Variant

    Dim kind As String, qtyTitle As String
    Dim resultName As String, baseName As String
    Dim suffix As Long
    Dim oldScreen As Boolean
    Dim outputCreated As Boolean

    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "입고·출고 원본 또는 집계 결과 시트를 선택해주세요.", _
               vbExclamation
        Exit Sub
    End If

    Set src = ActiveSheet
    Set wb = src.Parent

    '-----------------------------------
    ' 1. 원본 유형 자동 판별
    '-----------------------------------

    ' 출고 원본
    facCol = RL_FindColumn(src, Array("출고처코드"))
    qtyCol = RL_FindColumn(src, Array("출고수량"))

    If facCol > 0 And qtyCol > 0 Then

        kind = "출고"
        qtyTitle = "출고수량"

        nameCol = RL_FindColumn(src, Array("품목명", "품명"))
        codeCol = RL_FindColumn(src, Array("거래처 품목코드"))

    Else

        ' 입고 원본
        facCol = RL_FindColumn(src, Array("수주처"))
        qtyCol = RL_FindColumn(src, Array("입고수량"))

        If facCol > 0 And qtyCol > 0 Then

            kind = "입고"
            qtyTitle = "입고수량"

            nameCol = RL_FindColumn(src, Array("품목명", "품명"))
            codeCol = RL_FindColumn(src, Array("수주처 품목코드"))

        Else

            ' 입고관리·STOCK 집계 결과
            kind = "집계"
            qtyTitle = "수량"

            facCol = RL_FindColumn(src, _
                Array("FAC", "공장", "수주처"))

            nameCol = RL_FindColumn(src, _
                Array("품목명", "품명", "Part Name", _
                      "ITEM NM", "ITEM NAME"))

            qtyCol = RL_FindColumn(src, _
                Array("입고수량", "입고수량 합계", _
                      "출고수량", "출고수량 합계", _
                      "Shipped Q'ty", "Shipped Qty", _
                      "수량", "QTY"))

            codeCol = RL_FindColumn(src, _
                Array("ITEM CD", "수주처 품목코드", _
                      "거래처 품목코드"))

            If qtyCol > 0 Then

                If InStr(CStr(src.Cells(1, qtyCol).Value2), _
                         "입고") > 0 Then

                    qtyTitle = "입고수량"
                    kind = "입고"

                ElseIf InStr(CStr(src.Cells(1, qtyCol).Value2), _
                             "출고") > 0 Or _
                       InStr(UCase$(CStr( _
                             src.Cells(1, qtyCol).Value2)), _
                             "SHIPPED") > 0 Then

                    qtyTitle = "출고수량"
                    kind = "출고"

                End If

            End If

        End If

    End If

    If facCol = 0 Or nameCol = 0 Or qtyCol = 0 Then

        MsgBox "첫 행에서 공장·품명·수량 열을 찾지 못했습니다." & _
               vbCrLf & vbCrLf & _
               "출고 원본: 출고처코드 / 품목명 / 출고수량" & vbCrLf & _
               "입고 원본: 수주처 / 품목명 / 입고수량" & vbCrLf & _
               "집계 결과: FAC 또는 공장 / 품명 / 수량", _
               vbExclamation

        Exit Sub
    End If

    lastRow = Application.Max( _
        src.Cells(src.Rows.Count, facCol).End(xlUp).Row, _
        src.Cells(src.Rows.Count, nameCol).End(xlUp).Row, _
        src.Cells(src.Rows.Count, qtyCol).End(xlUp).Row)

    If codeCol > 0 Then
        lastRow = Application.Max(lastRow, _
            src.Cells(src.Rows.Count, codeCol).End(xlUp).Row)
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    '-----------------------------------
    ' 2. FAC + 품명별 수량 합산
    '-----------------------------------
    For r = 2 To lastRow

        If IsError(src.Cells(r, facCol).Value2) Or _
           IsError(src.Cells(r, nameCol).Value2) Then

            Err.Raise vbObjectError + 901, , _
                "공장 또는 품명에 오류가 있습니다." & _
                vbCrLf & "원본 행: " & r
        End If

        fac = RL_Text(src.Cells(r, facCol).Value2)
        itemName = RL_Text(src.Cells(r, nameCol).Value2)
        itemCode = ""

        If codeCol > 0 Then

            If IsError(src.Cells(r, codeCol).Value2) Then
                Err.Raise vbObjectError + 902, , _
                    "품목코드에 오류가 있습니다." & _
                    vbCrLf & "원본 행: " & r
            End If

            itemCode = RL_Text(src.Cells(r, codeCol).Value2)

        End If

        ' 기존 합계행 제외
        If fac = "" And itemCode = "" Then

            Select Case UCase$(Replace(itemName, " ", ""))
                Case "합계", "총합계", "계", "TOTAL", "GRANDTOTAL"
                    GoTo NextRow
            End Select

        End If

        ' 완전히 빈 행 제외
        If fac = "" And itemName = "" And itemCode = "" Then

            If Not IsError(src.Cells(r, qtyCol).Value2) Then
                If RL_Text(src.Cells(r, qtyCol).Value2) = "" Then
                    GoTo NextRow
                End If
            End If

        End If

        If itemName = "" Then
            Err.Raise vbObjectError + 903, , _
                "품명이 비어 있습니다." & vbCrLf & _
                "원본 행: " & r
        End If

        ' 공장코드를 FAC로 변환
        Select Case UCase$(fac)

            Case "1010", "F1"
                fac = "F1"

            Case "1060", "F2"
                fac = "F2"

            Case "1150", "F5"
                fac = "F5"

            Case Else
                Err.Raise vbObjectError + 904, , _
                    "공장을 구분할 수 없습니다." & vbCrLf & _
                    "원본 행: " & r & vbCrLf & _
                    "공장 값: [" & fac & "]"

        End Select

        If IsError(src.Cells(r, qtyCol).Value2) Then
            Err.Raise vbObjectError + 905, , _
                "수량에 오류가 있습니다." & vbCrLf & _
                "원본 행: " & r
        End If

        If Not IsNumeric(src.Cells(r, qtyCol).Value2) Then
            Err.Raise vbObjectError + 906, , _
                "수량이 숫자가 아닙니다." & vbCrLf & _
                "원본 행: " & r
        End If

        qty = CDbl(src.Cells(r, qtyCol).Value2)

        ' 같은 FAC와 품명이면 품목코드와 관계없이 합산
        key = Len(fac) & ":" & fac & itemName

        If dict.exists(key) Then

            data = dict(key)
            data(2) = data(2) + qty
            dict(key) = data

        Else

            dict.Add key, Array(fac, itemName, qty)

        End If

NextRow:
    Next r

    If dict.Count = 0 Then
        MsgBox "출력할 상세 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    '-----------------------------------
    ' 3. 새 결과 시트 생성
    '-----------------------------------
    baseName = "반장님_재고_" & Format(Now, "hhmmss")
    resultName = baseName

    Do While RL_SheetExists(wb, resultName)
        suffix = suffix + 1
        resultName = baseName & "_" & suffix
    Loop

    Application.ScreenUpdating = False

    Set dest = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    outputCreated = True
    dest.Name = resultName

    With dest

        ' FAC 열은 좁으므로 제목과 날짜는 B1/C1에 표시
        .Cells(1, 2).Value2 = "재고리스트"
        .Cells(1, 3).Value2 = "(" & Format(Date, "yy-mm-dd") & ")"

        .Range("A2:I2").value = _
            Array("FAC", "품목명△", qtyTitle, _
                  "팔레트", "수량", _
                  "팔레트", "수량", _
                  "팔레트", "수량")

        .Columns("A:B").NumberFormat = "@"

    End With

    outRow = 3

    For Each k In dict.keys

        data = dict(k)

        dest.Cells(outRow, 1).Value2 = data(0)
        dest.Cells(outRow, 2).Value2 = data(1)
        dest.Cells(outRow, 3).Value2 = data(2)

        ' D:I는 팔레트·수량 기록용 빈칸
        outRow = outRow + 1

    Next k

    lastOutputRow = outRow - 1

    '-----------------------------------
    ' 4. FAC → 품명 오름차순 정렬
    '-----------------------------------
    If dict.Count > 1 Then

        With dest.Sort

            .SortFields.Clear

            .SortFields.Add _
                key:=dest.Range("A3:A" & lastOutputRow), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal

            .SortFields.Add _
                key:=dest.Range("B3:B" & lastOutputRow), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal

            .SetRange dest.Range("A2:I" & lastOutputRow)
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply

        End With

    End If

    '-----------------------------------
    ' 5. 연호 재고리스트 형태의 흑백 서식
    '-----------------------------------
    With dest.Range("A1:I" & lastOutputRow)

        .Font.Name = "맑은 고딕"
        .Font.Color = vbBlack
        .Interior.Pattern = xlNone
        .VerticalAlignment = xlCenter
        .WrapText = False

    End With

    With dest.Range("B1:C1")

        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .ShrinkToFit = True

    End With

    With dest.Range("A2:I" & lastOutputRow)

        .Font.Size = 14

        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack
        End With

    End With

    With dest.Range("A2:I2")

        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .ShrinkToFit = True

    End With

    With dest

        ' 품명 공간을 넓게 확보
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 34
        .Columns("C").ColumnWidth = 13
        .Columns("D:I").ColumnWidth = 9

        .Range("A3:A" & lastOutputRow).HorizontalAlignment = xlCenter
        .Range("C3:I" & lastOutputRow).HorizontalAlignment = xlRight

        ' 품명은 한 줄 표시
        ' 긴 품명만 셀 너비에 맞춰 글자 축소
        With .Range("B3:B" & lastOutputRow)
            .WrapText = False
            .ShrinkToFit = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With

        ' 수량 표시
        .Range("C3:I" & lastOutputRow).NumberFormat = "#,##0.######"

        ' 팔레트 번호는 텍스트로 입력 가능
        .Range("D3:D" & lastOutputRow).NumberFormat = "@"
        .Range("F3:F" & lastOutputRow).NumberFormat = "@"
        .Range("H3:H" & lastOutputRow).NumberFormat = "@"

        ' 행 높이 고정 - AutoFit 사용하지 않음
        .Rows(1).RowHeight = 32
        .Rows(2).RowHeight = 30
        .Rows("3:" & lastOutputRow).RowHeight = 28

    End With

    '-----------------------------------
    ' 6. A4 세로·흑백 인쇄
    '-----------------------------------
    With dest.PageSetup

        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .BlackAndWhite = True

        .PrintArea = dest.Range("A1:I" & lastOutputRow).Address
        .PrintTitleRows = "$1:$2"

        .PrintGridlines = False
        .PrintHeadings = False
        .CenterHorizontally = True

        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)

        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False

    End With

    Application.ScreenUpdating = oldScreen
    dest.Activate

    MsgBox "반장님 재고리스트를 만들었습니다." & vbCrLf & _
           "원본 구분: " & kind & vbCrLf & _
           "FAC·품명별 집계: " & dict.Count & "건" & vbCrLf & _
           "품명 한 줄 표시 · 행 높이 고정" & vbCrLf & _
           "Ctrl+P에서 인쇄 모양을 확인해주세요.", _
           vbInformation

    Exit Sub

Failed:

    Application.ScreenUpdating = oldScreen

    If outputCreated Then

        MsgBox "출력 또는 인쇄 설정 중 오류가 발생했습니다." & _
               vbCrLf & "생성된 시트를 확인해주세요." & _
               vbCrLf & Err.Description, vbExclamation

    Else

        MsgBox Err.Description, vbExclamation

    End If

End Sub


' 문자열 양쪽 공백 정리
Private Function RL_Text(ByVal value As Variant) As String

    RL_Text = Trim$(Replace(CStr(value), ChrW(160), " "))

End Function


' 머리글 비교용 문자열
Private Function RL_HeaderKey(ByVal value As Variant) As String

    Dim text As String

    text = UCase$(RL_Text(value))
    text = Replace(text, " ", "")
    text = Replace(text, vbTab, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, "△", "")
    text = Replace(text, "▲", "")

    RL_HeaderKey = text

End Function


' 첫 행에서 해당 머리글 열 찾기
Private Function RL_FindColumn( _
    ByVal ws As Worksheet, _
    ByVal aliases As Variant) As Long

    Dim c As Long, lastCol As Long
    Dim alias As Variant
    Dim header As String
    Dim matched As Boolean

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol

        If Not IsError(ws.Cells(1, c).Value2) Then

            header = RL_HeaderKey(ws.Cells(1, c).Value2)
            matched = False

            For Each alias In aliases

                If header = RL_HeaderKey(alias) Then
                    matched = True
                    Exit For
                End If

            Next alias

            If matched Then

                If RL_FindColumn <> 0 Then
                    Err.Raise vbObjectError + 920, , _
                        "같은 역할의 머리글이 여러 개 있습니다: " & _
                        CStr(ws.Cells(1, c).Value2)
                End If

                RL_FindColumn = c

            End If

        End If

    Next c

End Function


' 같은 이름의 시트 존재 여부
Private Function RL_SheetExists( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Boolean

    Dim sh As Object

    For Each sh In wb.Sheets

        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            RL_SheetExists = True
            Exit Function
        End If

    Next sh

End Function
