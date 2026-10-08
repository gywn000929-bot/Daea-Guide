Attribute VB_Name = "modYeonhoInventoryPrint"
Sub 연호_반장님_재고리스트()

    Dim srcSheet As Worksheet, destSheet As Worksheet

    Dim lastRow As Long, i As Long, r As Long

    Dim dict As Object, k As Variant

    Dim itemKey As String, itemQty As Double

    Dim invoiceDate As String, rawDate As String



    ' 속도 향상, 화면 깜빡임 차단, 파일 과부하 방지 설정

    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    Application.EnableEvents = False



    ' 1. 소스 시트 지정 (거래명세서 19 시트 우선 탐색)

    On Error Resume Next

    Set srcSheet = Sheets("거래명세서 (19)")

    On Error GoTo 0

    If srcSheet Is Nothing Then Set srcSheet = ActiveSheet



    ' 2. 거래명세서 내부에서 작성일자 자동 추출 및 변환 (YYYY-MM-DD -> YY/MM/DD)

    invoiceDate = ""

    lastRow = srcSheet.Cells(srcSheet.Rows.Count, 1).End(xlUp).Row

    For i = 1 To IIf(lastRow > 30, 30, lastRow)

        If InStr(Replace(srcSheet.Cells(i, 1).value, " ", ""), "작성일자") > 0 Then

            If srcSheet.Cells(i + 1, 1).value <> "" Then

                rawDate = Trim(srcSheet.Cells(i + 1, 1).value)

                If Len(rawDate) >= 10 Then

                    invoiceDate = Mid(rawDate, 3, 2) & "/" & Mid(rawDate, 6, 2) & "/" & Mid(rawDate, 9, 2)

                End If

            End If

            Exit For

        End If

    Next i

    If invoiceDate = "" Then invoiceDate = Format(Date, "yy/mm/dd")



    ' 3. 메모리 내 딕셔너리를 사용하여 품명별 수량 초고속 부분합 계산

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To lastRow

If IsNumeric(srcSheet.Cells(i, 1).value) And srcSheet.Cells(i, 1).value <> "" And IsNumeric(srcSheet.Cells(i, 3).value) And srcSheet.Cells(i, 3).value <> "" Then



            ' 유령 공백 제거 및 텍스트 정제

            itemKey = Trim(Replace(srcSheet.Cells(i, 2).value, Chr(160), " "))

            itemQty = srcSheet.Cells(i, 3).value



            If itemKey <> "" Then

                If dict.exists(itemKey) Then

                    dict(itemKey) = dict(itemKey) + itemQty

                Else

                    dict.Add itemKey, itemQty

                End If

            End If

        End If

    Next i



    ' 4. "재고 리스트" 출력 시트 초기화 (용량 누적 방지를 위해 완전 삭제 후 재생성)

    On Error Resume Next

    Set destSheet = Sheets("재고 리스트")

    On Error GoTo 0

    If Not destSheet Is Nothing Then destSheet.Delete



    Set destSheet = Sheets.Add(After:=srcSheet)

    destSheet.Name = "재고 리스트"



    ' 5. [오리지널 양식] 상단 뼈대 구조 작성 (A열부터 H열까지)

    With destSheet

        ' 1행: 타이틀 포맷

        .Cells(1, 1).value = "연호"

        .Cells(1, 2).value = "(" & invoiceDate & ")"



        ' 2행: 헤더명 지정 (원본 파일과 100% 일치)

        .Cells(2, 1).value = "품목명△"

        .Cells(2, 2).value = "출고수량"

        .Cells(2, 3).value = "팔레트"

        .Cells(2, 4).value = "수량"

        .Cells(2, 5).value = "팔레트"

        .Cells(2, 6).value = "수량"

        .Cells(2, 7).value = "팔레트"

        .Cells(2, 8).value = "수량"

    End With



    ' 6. 정제된 데이터를 행 순서대로 일괄 기입

    r = 3

    For Each k In dict.keys

        destSheet.Cells(r, 1).value = k

        destSheet.Cells(r, 2).value = dict(k)

        r = r + 1

    Next k



    Dim dataLastRow As Long

    dataLastRow = r - 1



    ' 7. 품목명 오름차순(가나다 순) 자동 정렬

    If dataLastRow >= 3 Then

        With destSheet.Sort

            .SortFields.Clear

.SortFields.Add key:=destSheet.Range("A3:A" & dataLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

            .SetRange destSheet.Range("A2:H" & dataLastRow)

            .header = xlYes

            .Apply

        End With

    End If



    ' 8. [서식 극대화 지정] 글자 크기 대폭 확대, 행 높이 증대, 반복 서식 적용

    With destSheet

        ' 전체 데이터 표 영역 지정 (A2부터 H열 끝까지)

        Dim tableRange As Range

        Set tableRange = .Range("A2:H" & dataLastRow)



        ' 폰트 및 글자 크기 눈에 띄게 대폭 확대 (14pt로 전면 확대)

        tableRange.Font.Name = "맑은 고딕"

        tableRange.Font.Size = 14



        ' 타이틀행 글자 크기 조정 (16pt, Bold)

        .Range("A1:B1").Font.Name = "맑은 고딕"

        .Range("A1:B1").Font.Size = 16

        .Range("A1:B1").Font.Bold = True



        ' 행 높이 설정 (현장에서 아주 시원시원하게 보이도록 넓힘)

        .Rows(1).RowHeight = 32        ' 타이틀 행 높이

        .Rows(2).RowHeight = 30        ' 헤더 행 높이

        .Range("3:" & dataLastRow).RowHeight = 28 ' 데이터 행 높이



        ' 2행 헤더 스타일 (글자 굵게, 가운데 정렬)

        With .Range("A2:H2")

            .Font.Bold = True

            .HorizontalAlignment = xlCenter

        End With



        ' 각 열별 데이터 정렬 및 숫자 포맷 (2번째 열 출고수량 쉼표 처리)

        .Range("A3:A" & dataLastRow).HorizontalAlignment = xlLeft   ' 품목명 왼쪽 정렬

        .Range("B3:B" & dataLastRow).HorizontalAlignment = xlRight  ' 출고수량 오른쪽 정렬

        .Range("B3:B" & dataLastRow).NumberFormat = "#,##0"        ' 천 단위 콤마



        ' [반복 서식] 3~8번째(C~H열) 빈 칸 서식도 완전히 동일하게 우측 정렬 및 천 단위 포맷 바인딩

        .Range("C3:H" & dataLastRow).HorizontalAlignment = xlRight

        .Range("C3:H" & dataLastRow).NumberFormat = "#,##0"



        ' 전 영역에 오리지널 느낌의 깨끗하고 선명한 격자 테두리선(Thin) 부여

        With tableRange.Borders

            .LineStyle = xlContinuous

            .Weight = xlThin

            .ColorIndex = xlAutomatic

        End With



        ' 열 너비 균형 조정 (글자가 크기 때문에 넉넉하게 확장)

        .Columns("A").ColumnWidth = 28 ' 품목명 칸 넉넉하게

        .Columns("B").ColumnWidth = 15 ' 출고수량 칸



        ' C열부터 H열까지 반복되는 팔레트/수량 열 너비 인쇄 밸런스에 맞춰 고정

        Dim colIdx As Long

        For colIdx = 3 To 8

            .Columns(colIdx).ColumnWidth = 11

        Next colIdx



        ' 9. 인쇄 페이지 맞춤 및 [★매 페이지 제목행 반복★] 설정

        With .PageSetup

            ' 이 코드가 핵심입니다: 인쇄 시 1행과 2행(제목과 헤더)을 모든 페이지 상단에 반복 출력합니다.

            .PrintTitleRows = "$1:$2"



            .Orientation = xlPortrait          ' 세로 인쇄

            .CenterHorizontally = True         ' 좌우 가운데 정렬

            .PrintGridlines = False            ' 테두리선만 깔끔하게 출력

            .TopMargin = Application.InchesToPoints(0.4)

            .BottomMargin = Application.InchesToPoints(0.4)

            .LeftMargin = Application.InchesToPoints(0.4)

            .RightMargin = Application.InchesToPoints(0.4)

            .Zoom = False

            .FitToPagesWide = 1                ' 가로 폭 자동 맞춤 (우측 서식이 절대 안 잘림)

            .FitToPagesTall = False            ' 세로 길이는 데이터 양에 따라 여러 장으로 자유롭게 넘어가도록 설정

        End With

    End With



    ' 기능 정상 복구

    Application.ScreenUpdating = True

    Application.DisplayAlerts = True

    Application.EnableEvents = True



    MsgBox "모든 페이지 제목 반복 설정이 반영된 재고 리스트가 완성되었습니다!", vbInformation, "완료"

End Sub
