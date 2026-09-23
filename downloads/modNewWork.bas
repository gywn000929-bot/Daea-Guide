Attribute VB_Name = "modNewWork"
Option Explicit

Public Sub 새업무_대상확인()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "작업할 Excel 파일을 먼저 열어주세요."
        Exit Sub
    End If
    If wb Is ThisWorkbook Then Exit Sub
    If TypeName(Application.ActiveSheet) <> "Worksheet" Then
        MsgBox "작업할 워크시트를 선택해주세요."
        Exit Sub
    End If
    Set ws = Application.ActiveSheet
    MsgBox "작업 파일: " & wb.Name & vbCrLf & _
           "작업 시트: " & ws.Name
End Sub
