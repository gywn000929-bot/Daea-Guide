Attribute VB_Name = "modPalletLocations"
Sub 파렛트_번호_정리()

    Dim ws As Worksheet

    Dim startCol As Long, endCol As Long, sumCol As Long

    Dim resCol1 As Long, resCol2 As Long

    Dim lastRow As Long

    Dim r As Long, c As Long

    Dim palletList As String, itemName As String



    Set ws = ActiveSheet



    ' 1. 키워드로 범위 찾기

    On Error Resume Next

    startCol = ws.Rows(1).Find("출고수량").Column + 1

    sumCol = ws.Rows(1).Find("Sum").Column

    endCol = sumCol - 1



    ' 결과 열 위치 (Sum 뒤 두 칸)

    resCol1 = sumCol + 1

    resCol2 = sumCol + 2

    On Error GoTo 0



    ' 에러 체크

    If startCol <= 1 Or sumCol <= 0 Then

        MsgBox "'출고수량' 또는 'Sum' 머리글을 찾을 수 없습니다.", vbExclamation

        Exit Sub

    End If



    ' 2. 마지막 행 찾기

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row



    ' 3. 메인 루프

    For r = 2 To lastRow

        palletList = ""

        itemName = ws.Cells(r, 1).value



        ' 파렛트 번호 수집 및 "번" 결합

        For c = startCol To endCol

            If Trim(ws.Cells(r, c).value) <> "" Then

                ' 제목(숫자) 뒤에 "번"을 붙여서 저장

                If palletList = "" Then

                    palletList = ws.Cells(1, c).value & "번"

                Else

                    palletList = palletList & ", " & ws.Cells(1, c).value & "번"

                End If

            End If

        Next c



        ' 4. 데이터 입력 및 왼쪽 정렬 설정

        With ws.Cells(r, resCol1)

            .value = itemName

            .HorizontalAlignment = xlHAlignLeft ' 왼쪽 정렬

        End With



        With ws.Cells(r, resCol2)

            .value = palletList

            .HorizontalAlignment = xlHAlignLeft ' 왼쪽 정렬

        End With

    Next r



    ' 결과 열 디자인 정리

    With ws.Range(ws.Cells(1, resCol1), ws.Cells(1, resCol2))

        .Font.Bold = True

        .HorizontalAlignment = xlHAlignLeft

    End With



    ws.Cells(1, resCol1).value = "확인 품목"

    ws.Cells(1, resCol2).value = "해당 파렛트"



    ws.Columns(resCol1).AutoFit

    ws.Columns(resCol2).AutoFit



    MsgBox "정리가 완료되었습니다. 모든 결과는 왼쪽으로 정렬되었습니다!", vbInformation

End Sub
