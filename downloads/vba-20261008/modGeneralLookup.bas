Attribute VB_Name = "modGeneralLookup"
Option Explicit
Public Sub 범용_값조회()
    Dim keys As Range, vals As Range, targets As Range, dest As Range
    Dim d As Object, dup As Object, result() As Variant
    Dim i As Long, k As String, miss As Long
    On Error Resume Next
    Set keys = Application.InputBox("원본의 검색 기준 한 열을 선택하세요. 머리글 제외", "범용 값조회 1/4", Type:=8)
    On Error GoTo Failed
    If keys Is Nothing Then Exit Sub
    On Error Resume Next
    Set vals = Application.InputBox("가져올 값 한 열을 선택하세요. 기준 열과 행 수가 같아야 합니다.", "범용 값조회 2/4", Type:=8)
    On Error GoTo Failed
    If vals Is Nothing Then Exit Sub
    On Error Resume Next
    Set targets = Application.InputBox("조회할 값이 있는 한 열을 선택하세요. 머리글 제외", "범용 값조회 3/4", Type:=8)
    On Error GoTo Failed
    If targets Is Nothing Then Exit Sub
    On Error Resume Next
    Set dest = Application.InputBox("결과를 넣을 첫 셀 하나를 선택하세요.", "범용 값조회 4/4", Type:=8)
    On Error GoTo Failed
    If dest Is Nothing Then Exit Sub
    If keys.Areas.Count <> 1 Or vals.Areas.Count <> 1 Or targets.Areas.Count <> 1 Or dest.Areas.Count <> 1 Then Err.Raise 5, , "연속 범위를 선택하세요."
    If keys.Columns.Count <> 1 Or vals.Columns.Count <> 1 Or targets.Columns.Count <> 1 Or dest.Cells.CountLarge <> 1 Then Err.Raise 5, , "기준/값/조회는 한 열, 결과 시작은 한 셀을 선택하세요."
    If keys.Rows.Count <> vals.Rows.Count Then Err.Raise 5, , "원본 기준 열과 값 열의 행 수가 다릅니다."
    If dest.Row + targets.Rows.Count - 1 > dest.Worksheet.Rows.Count Then Err.Raise 5, , "결과가 시트의 마지막 행을 넘습니다."
    Set d = CreateObject("Scripting.Dictionary")
    Set dup = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    dup.CompareMode = vbBinaryCompare
    For i = 1 To keys.Rows.Count
        If Not IsError(keys.Cells(i, 1).Value2) Then
            k = CStr(keys.Cells(i, 1).Value2)
            If Len(k) > 0 Then
                If d.Exists(k) Then
                    dup(k) = True
                Else
                    d.Add k, vals.Cells(i, 1).Value2
                End If
            End If
        End If
    Next i
    ReDim result(1 To targets.Rows.Count, 1 To 1)
    For i = 1 To targets.Rows.Count
        result(i, 1) = CVErr(xlErrNA)
        If Not IsError(targets.Cells(i, 1).Value2) Then
            k = CStr(targets.Cells(i, 1).Value2)
            If d.Exists(k) And Not dup.Exists(k) Then result(i, 1) = d(k)
        End If
        If IsError(result(i, 1)) Then miss = miss + 1
    Next i
    If MsgBox("선택한 셀부터 " & targets.Rows.Count & "행을 결과 값으로 덮어씁니다. 진행할까요?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    dest.Resize(targets.Rows.Count, 1).Value2 = result
    MsgBox "완료. 미조회/중복 기준/원본 오류 등 확인할 결과: " & miss & "건 (#N/A 또는 원본 오류)", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
