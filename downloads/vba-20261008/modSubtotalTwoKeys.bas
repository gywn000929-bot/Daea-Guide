Attribute VB_Name = "modSubtotalTwoKeys"
Sub 부분합()
    Dim rng As Range
    Dim arr As Variant
    Dim dict As Object
    Dim i As Long, j As Long
    Dim keyCol1 As Variant, keyCol2 As Variant, valCol As Variant, extraInput As Variant
    Dim outWs As Worksheet
    Dim k As Variant, r As Long
    Dim headerKey1 As String, headerKey2 As String, headerVal As String

    ' 1. 범위 확인 및 방어 코드
    If TypeName(Selection) <> "Range" Then
        MsgBox "범위를 먼저 선택하세요."
        Exit Sub
    End If
    
    Set rng = Selection
    arr = rng.value

    If Not IsArray(arr) Then
        MsgBox "2행 이상 데이터 범위를 선택하세요."
        Exit Sub
    End If

    If UBound(arr, 1) < 2 Then
        MsgBox "머리글을 포함하여 최소 2행 이상 선택해야 합니다."
        Exit Sub
    End If

    ' 2. 열 번호 입력 받기
    keyCol1 = Application.InputBox("첫 번째 기준 열 번호(예: 거래처코드)를 입력하세요.", "1. 첫 번째 기준", Type:=1)
    If keyCol1 = False Or keyCol1 <= 0 Then Exit Sub

    keyCol2 = Application.InputBox("두 번째 기준 열 번호(예: 품목명)를 입력하세요.", "2. 두 번째 기준", Type:=1)
    If keyCol2 = False Or keyCol2 <= 0 Then Exit Sub

    valCol = Application.InputBox("합산할 금액/수량 열 번호를 입력하세요.", "3. 합계 열 선택", Type:=1)
    If valCol = False Or valCol <= 0 Then Exit Sub
    
    If keyCol1 > UBound(arr, 2) Or keyCol2 > UBound(arr, 2) Or valCol > UBound(arr, 2) Then
        MsgBox "선택한 범위 내의 열 번호를 입력해주세요."
        Exit Sub
    End If

    ' 2-1. 추가로 출력할 열 입력 받기
    Dim extraCols() As Long
    Dim hasExtra As Boolean
    Dim extraStrArr() As String
    Dim eCount As Long
    
    extraInput = Application.InputBox("추가로 함께 출력할 열 번호를 입력하세요." & vbCrLf & _
                                      "(예: 3,4 와 같이 쉼표로 구분 가능)" & vbCrLf & _
                                      "없으면 빈칸으로 '확인' 클릭", "4. 추가 정보 열 선택", Type:=2)
    
    hasExtra = False: eCount = 0
    If extraInput <> "False" And Trim(extraInput) <> "" Then
        extraStrArr = Split(extraInput, ",")
        ReDim extraCols(0 To UBound(extraStrArr))
        For j = 0 To UBound(extraStrArr)
            If IsNumeric(Trim(extraStrArr(j))) Then
                Dim colIdx As Long: colIdx = CLng(Trim(extraStrArr(j)))
                If colIdx > 0 And colIdx <= UBound(arr, 2) Then
                    extraCols(eCount) = colIdx
                    eCount = eCount + 1
                End If
            End If
        Next j
        If eCount > 0 Then
            ReDim Preserve extraCols(0 To eCount - 1)
            hasExtra = True
        End If
    End If

    ' 3. 데이터 합산 (Dictionary)
    Set dict = CreateObject("Scripting.Dictionary")
    
    headerKey1 = arr(1, keyCol1)
    headerKey2 = arr(1, keyCol2)
    headerVal = arr(1, valCol)

    Dim combinedKey As String ' 두 기준을 합친 고유 키
    Dim currentVal As Double
    Dim itemData() As Variant ' [0]합계, [1~n]추가 데이터

    For i = 2 To UBound(arr, 1)
        ' 두 기준 열을 "|" 구분자로 합쳐서 유니크한 키 생성
        combinedKey = CStr(arr(i, keyCol1)) & "|" & CStr(arr(i, keyCol2))
        currentVal = Val(arr(i, valCol))
        
        If dict.exists(combinedKey) Then
            itemData = dict(combinedKey)
            itemData(0) = itemData(0) + currentVal
            dict(combinedKey) = itemData
        Else
            ReDim itemData(0 To eCount)
            itemData(0) = currentVal
            If hasExtra Then
                For j = 0 To eCount - 1
                    itemData(j + 1) = arr(i, extraCols(j))
                Next j
            End If
            dict.Add combinedKey, itemData
        End If
    Next i

    ' 4. 결과 출력
    Set outWs = Worksheets.Add(After:=ActiveSheet)
    On Error Resume Next
    outWs.Name = "중복기준_부분합_" & Format(Now, "hhmmss")
    On Error GoTo 0

    ' 머리글 출력 (기준1, 기준2, 합계순)
    outWs.Cells(1, 1).value = headerKey1
    outWs.Cells(1, 2).value = headerKey2
    outWs.Cells(1, 3).value = headerVal & " 합계"
    
    If hasExtra Then
        For j = 0 To eCount - 1
            outWs.Cells(1, 4 + j).value = arr(1, extraCols(j))
        Next j
    End If
    
    ' 데이터 출력
    r = 2
    Dim splitKey As Variant
    For Each k In dict.keys
        splitKey = Split(k, "|")
        outWs.Cells(r, 1).value = splitKey(0) ' 기준1
        outWs.Cells(r, 2).value = splitKey(1) ' 기준2
        
        itemData = dict(k)
        outWs.Cells(r, 3).value = itemData(0) ' 합계
        
        If hasExtra Then
            For j = 0 To eCount - 1
                outWs.Cells(r, 4 + j).value = itemData(j + 1)
            Next j
        End If
        r = r + 1
    Next k

    ' 5. 읽기 쉬운 표와 흑백 인쇄 서식
    Dim lastCol As Long: lastCol = 3 + eCount
    Dim body As Range, col As Long
    Set body = outWs.Range(outWs.Cells(1, 1), outWs.Cells(r - 1, lastCol))
    With body
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Color = vbBlack
        .VerticalAlignment = xlCenter
        .WrapText = False
        .RowHeight = 23
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(190, 190, 190)
    End With
    With outWs.Range(outWs.Cells(1, 1), outWs.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = RGB(230, 230, 230)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With
    outWs.Range(outWs.Cells(2, 3), outWs.Cells(r - 1, 3)).NumberFormat = "#,##0.########;-#,##0.########;0"
    outWs.Range(outWs.Cells(2, 3), outWs.Cells(r - 1, 3)).HorizontalAlignment = xlRight
    outWs.Range(outWs.Columns(1), outWs.Columns(lastCol)).AutoFit
    For col = 1 To lastCol
        With outWs.Columns(col)
            If .ColumnWidth < 14 Then .ColumnWidth = 14
            If .ColumnWidth > 42 Then .ColumnWidth = 42
        End With
    Next col
    outWs.Range(outWs.Cells(2, 1), outWs.Cells(r - 1, lastCol)).ShrinkToFit = True
    body.AutoFilter
    With outWs.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        If lastCol > 5 Then .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .BlackAndWhite = True
        .PrintTitleRows = "$1:$1"
        .PrintArea = body.Address
        .CenterHorizontally = True
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .CenterFooter = "&P / &N"
    End With
    MsgBox "부분합 완료! 새 시트에 서식과 A4 인쇄 설정을 적용했습니다.", vbInformation
End Sub
