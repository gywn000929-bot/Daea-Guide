Attribute VB_Name = "modYeonhoStatement"
Sub 연호거래명세서()

    Dim wsSource As Worksheet, wsTarget As Worksheet

    Dim lastRow As Long, i As Long, targetRow As Long

    Dim checkVal As Variant

    Dim excludeList As Variant

    Dim item As Variant

    Dim isExclude As Boolean

    Dim tbl As ListObject



    ' 변환 관련 변수

    Dim rawG As String, partsG As Variant, p As Variant

    Dim finalG As String, cleanB As String

    Dim arrG() As String, n As Integer, j As Integer, k As Integer, temp As String



    ' 1. 소스 시트 설정

    Set wsSource = ActiveSheet



    ' 2. 새 시트 생성 및 이름 설정

    Set wsTarget = Sheets.Add(After:=Sheets(Sheets.Count))

    wsTarget.Name = "연호_정제_" & Format(Now, "hhmm")



    ' 3. 19행을 머릿말(Header)로 새 시트의 1행에 복사

    wsSource.Rows(19).Copy destination:=wsTarget.Rows(1)

    targetRow = 2



    ' 4. 제외할 문구 리스트

excludeList = Array("거 래 명 세 서", "공", "급", "자", "작 성 일 자", "출고인:", "특기사항", "본 거래명세서는", "사용하는 계산서이며")



    ' 데이터 끝 확인 (B열 기준)

    lastRow = wsSource.Cells(wsSource.Rows.Count, 2).End(xlUp).Row



    Application.ScreenUpdating = False



    ' 5. 20행부터 끝까지 데이터 검사 및 복사

    For i = 20 To lastRow

        checkVal = wsSource.Cells(i, 1).value ' A열(No.) 기준

        isExclude = False



        ' 빈 칸, 제목줄, 날짜, 제외문구 걸러내기

        If Trim(CStr(checkVal)) = "" Or Trim(CStr(checkVal)) = "No." Then

            isExclude = True

        ElseIf IsDate(checkVal) Then

            isExclude = True

        Else

            For Each item In excludeList

                If InStr(CStr(wsSource.Cells(i, 1).value & wsSource.Cells(i, 2).value), item) > 0 Then

                    isExclude = True

                    Exit For

                End If

            Next item

        End If



        ' 필터에 걸리지 않은 알맹이 데이터 처리 및 복사

        If Not isExclude Then

            ' 전체 행 복사

            wsSource.Rows(i).Copy destination:=wsTarget.Rows(targetRow)



            ' ----------------------------------------------------

            ' [B열] 품명 및 규격: 악성 유니코드 공백(ChrW) 완벽 제거

            ' ----------------------------------------------------

            cleanB = CStr(wsTarget.Cells(targetRow, 2).value)

            cleanB = Replace(cleanB, ChrW(160), " ")     ' 한국어 엑셀용 특수공백 치환

            cleanB = Replace(cleanB, ChrW(&H3000), " ")  ' 전각 공백 치환

            wsTarget.Cells(targetRow, 2).value = Trim(cleanB)



            ' ----------------------------------------------------

            ' [G열] 비고: 날짜, 공백 제거 및 오름차순 정렬

            ' ----------------------------------------------------

            rawG = CStr(wsTarget.Cells(targetRow, 7).value)

            rawG = Replace(rawG, ChrW(160), " ")

            rawG = Replace(rawG, ChrW(&H3000), " ")

            rawG = Trim(rawG)



            If rawG <> "" Then

                partsG = Split(rawG, ",")

                n = 0

                ReDim arrG(UBound(partsG))



                For Each p In partsG

                    temp = Trim(CStr(p))

                    temp = Replace(temp, ChrW(160), " ")

                    temp = Trim(temp)



                    ' 날짜 부분 (04/13) 제거

                    If InStr(temp, "(") > 0 Then

                        temp = Trim(Left(temp, InStr(temp, "(") - 1))

                    End If



                    If temp <> "" Then

                        arrG(n) = temp

                        n = n + 1

                    End If

                Next p



                ' 오름차순 정렬 (F1, F2 순서대로)

                If n > 1 Then

                    For j = 0 To n - 2

                        For k = j + 1 To n - 1

                            If arrG(j) > arrG(k) Then

                                temp = arrG(j)

                                arrG(j) = arrG(k)

                                arrG(k) = temp

                            End If

                        Next k

                    Next j

                End If



                ' 결과 다시 합치기

                finalG = ""

                For j = 0 To n - 1

                    finalG = finalG & IIf(finalG = "", "", ", ") & arrG(j)

                Next j

                wsTarget.Cells(targetRow, 7).value = finalG

            End If

            ' ----------------------------------------------------



            targetRow = targetRow + 1

        End If

    Next i



    ' 6. 데이터를 엑셀 '표' 스타일로 변환

    If targetRow > 2 Then

        Set tbl = wsTarget.ListObjects.Add(xlSrcRange, wsTarget.Range("A1").CurrentRegion, , xlYes)

        tbl.TableStyle = "TableStyleMedium2"

    End If



    ' 열 너비 자동 맞춤

    wsTarget.Columns.AutoFit

    Application.ScreenUpdating = True



    MsgBox "악성 들여쓰기 공백 및 G열 정리가 완벽하게 완료되었습니다!", vbInformation

End Sub
