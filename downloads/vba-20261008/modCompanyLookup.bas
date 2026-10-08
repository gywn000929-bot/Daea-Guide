Attribute VB_Name = "modCompanyLookup"
Option Explicit



Sub 업체명_조회()

    Dim ws As Worksheet, vendorWs As Worksheet

    Dim makerColStr As String, itemColStr As String

    Dim makerCol As Long, itemCol As Long, outCol As Long

    Dim lastRow As Long, vRow As Long, vLastRow As Long



    Dim makerDict As Object, exDict As Object

    Dim mk As String, vn As String, f_it As String, f_vn As String

    Dim mkStr As String, itStr As String, resultVal As String

    Dim i As Long, parenStart As Long



    ' 1. 시트 확인

    Set ws = ActiveSheet

    On Error Resume Next

    Set vendorWs = ActiveWorkbook.Sheets("거래처코드")

    On Error GoTo 0



    If vendorWs Is Nothing Then

        MsgBox "'거래처코드' 시트를 찾을 수 없습니다.", vbCritical

        Exit Sub

    End If



    ' 2. 기준 열 입력

    makerColStr = InputBox("MAKER 열 알파벳을 입력하세요. (예: C)", "MAKER 열 입력")

    If Trim(makerColStr) = "" Then Exit Sub



    itemColStr = InputBox("ITEM NM 열 알파벳을 입력하세요. (예: B 또는 F)", "ITEM NM 열 입력")

    If Trim(itemColStr) = "" Then Exit Sub



    On Error Resume Next

    makerCol = ws.Columns(makerColStr).Column

    itemCol = ws.Columns(itemColStr).Column

    On Error GoTo 0



    If makerCol = 0 Or itemCol = 0 Then

        MsgBox "올바른 열 알파벳을 입력해 주세요.", vbExclamation

        Exit Sub

    End If



    Application.ScreenUpdating = False



    ' 3. 거래처코드 시트 → 딕셔너리 로드

    Set makerDict = CreateObject("Scripting.Dictionary")

    Set exDict = CreateObject("Scripting.Dictionary")



    vLastRow = vendorWs.Cells(vendorWs.Rows.Count, "B").End(xlUp).Row

    For vRow = 2 To vLastRow

        ' 기본 MAKER 매핑 (B열 MAKER → A열 거래처명)

        mk = UCase(Trim(CStr(vendorWs.Cells(vRow, "B").value)))

        vn = Trim(CStr(vendorWs.Cells(vRow, "A").value))

        If mk <> "" And Not makerDict.exists(mk) Then makerDict.Add mk, vn



        ' ? 수정: 예외 ITEM 매핑 시 괄호 안 메이커 표기 제거

        f_it = UCase(Trim(CStr(vendorWs.Cells(vRow, "F").value)))



        ' 괄호 이전 품목코드만 추출 ("PTSR-0402 (DONG-A)" → "PTSR-0402")

        parenStart = InStr(f_it, "(")

        If parenStart > 0 Then

            f_it = Trim(Left(f_it, parenStart - 1))

        End If



        ' 특수공백 및 일반공백 제거

        f_it = Replace(Replace(f_it, ChrW(160), ""), ChrW(12288), "")

        f_it = Replace(f_it, " ", "")



        f_vn = Trim(CStr(vendorWs.Cells(vRow, "E").value))

        If f_it <> "" And Not exDict.exists(f_it) Then exDict.Add f_it, f_vn

    Next vRow



    ' 4. 출력 열 결정

    On Error Resume Next

outCol = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column + 1

    On Error GoTo 0

    If outCol = 1 Then outCol = 2



    ws.Cells(1, outCol).value = "거래처명(매칭)"

    ws.Cells(1, outCol).Interior.Color = RGB(255, 242, 204)

    ws.Cells(1, outCol).Font.Bold = True



    ' 5. 매칭 로직

lastRow = Application.WorksheetFunction.Max( ws.Cells(ws.Rows.Count, makerCol).End(xlUp).Row, ws.Cells(ws.Rows.Count, itemCol).End(xlUp).Row)



    For i = 2 To lastRow

        mkStr = UCase(Trim(CStr(ws.Cells(i, makerCol).value)))



        ' 작업시트 품목명도 괄호 제거 후 비교 (혹시 괄호가 포함된 경우 대비)

        itStr = UCase(Trim(CStr(ws.Cells(i, itemCol).value)))

        parenStart = InStr(itStr, "(")

        If parenStart > 0 Then itStr = Trim(Left(itStr, parenStart - 1))

        itStr = Replace(Replace(itStr, ChrW(160), ""), ChrW(12288), "")

        itStr = Replace(itStr, " ", "")



        resultVal = ""



        If mkStr <> "" Then

            If mkStr = "DONG-A" Or mkStr = "DONG-A BESTECH" Or mkStr = "KET" Then

                If exDict.exists(itStr) Then

                    resultVal = exDict(itStr)

                ElseIf makerDict.exists(mkStr) Then

                    resultVal = makerDict(mkStr)

                End If

            Else

                If makerDict.exists(mkStr) Then

                    resultVal = makerDict(mkStr)

                End If

            End If

        End If



        If resultVal = "" And (mkStr <> "" Or itStr <> "") Then

            resultVal = "미등록"

        End If



        ws.Cells(i, outCol).value = resultVal

    Next i



    ws.Columns(outCol).AutoFit

    Application.ScreenUpdating = True



    MsgBox "거래처 매칭이 완료되었습니다!" & vbCrLf & "데이터 맨 우측을 확인해 보세요.", vbInformation

End Sub
