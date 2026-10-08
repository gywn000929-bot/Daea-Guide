Attribute VB_Name = "modAirGrayscale"
Option Explicit

' Run this macro from Alt+F8 while the weekly AIR workbook and target sheet are active.
Public Sub ø°æÓ_π›¿Â¥‘_¿Œº‚()
    Const HEADER_SEARCH_ROWS As Long = 10
    Const DEFAULT_HEADER_ROW As Long = 1

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim headerCell As Range
    Dim headerRow As Long
    Dim makerCol As Long
    Dim lastRow As Long
    Dim lastTableCol As Long
    Dim printLastCol As Long
    Dim dataFirstRow As Long
    Dim tableRange As Range
    Dim bodyRange As Range
    Dim rowRange As Range
    Dim colRange As Range
    Dim lastCell As Range
    Dim c As Long
    Dim r As Long
    Dim h As String
    Dim currentMaker As String
    Dim previousMaker As String
    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim completed As Boolean
    Dim errText As String

    On Error GoTo CleanFail

    If Application.Workbooks.Count = 0 Then
        Err.Raise vbObjectError + 1000, , "No workbook is open."
    End If
    If Not TypeOf ActiveSheet Is Worksheet Then
        Err.Raise vbObjectError + 1001, , "Activate the AIR worksheet and run the macro again."
    End If

    Set wb = ActiveWorkbook
    Set ws = ActiveSheet

    If UCase$(wb.Name) = "PERSONAL.XLSB" Then
        Err.Raise vbObjectError + 1002, , "Activate the weekly AIR workbook, then run the macro again."
    End If
    If ws.ProtectContents Then
        Err.Raise vbObjectError + 1003, , "The active sheet is protected. Unprotect it and run the macro again."
    End If

    Set headerCell = ws.Range(ws.Cells(1, 1), ws.Cells(HEADER_SEARCH_ROWS, 50)).Find( _
        What:="Maker", After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    If headerCell Is Nothing Then
        Err.Raise vbObjectError + 1004, , "The 'Maker' header was not found in the first 10 rows."
    End If

    headerRow = headerCell.Row
    makerCol = headerCell.Column
    dataFirstRow = headerRow + 1

    lastTableCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastTableCol < makerCol Then
        Err.Raise vbObjectError + 1005, , "The header row is incomplete."
    End If

    Set lastCell = ws.Range(ws.Cells(dataFirstRow, 1), ws.Cells(ws.Rows.Count, lastTableCol)).Find( _
        What:="*", After:=ws.Cells(dataFirstRow, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

    If lastCell Is Nothing Then
        Err.Raise vbObjectError + 1006, , "No AIR data rows were found below the header."
    End If
    lastRow = lastCell.Row

    printLastCol = lastTableCol
    For c = 1 To lastTableCol
        If IsMatchingHelperHeader(NormalizeHeader(ws.Cells(headerRow, c).Value2)) Then
            If c > 1 Then printLastCol = c - 1
            Exit For
        End If
    Next c

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set tableRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastTableCol))
    Set bodyRange = ws.Range(ws.Cells(dataFirstRow, 1), ws.Cells(lastRow, lastTableCol))

    ' Base font and row geometry. Existing body fills are intentionally preserved.
    With tableRange.Font
        .Name = "Calibri"
        .Size = 11
        .Color = RGB(0, 0, 0)
        .Bold = False
        .Italic = False
    End With
    ws.Rows(headerRow).RowHeight = 30
    ws.Range(ws.Rows(dataFirstRow), ws.Rows(lastRow)).RowHeight = 22
    bodyRange.VerticalAlignment = xlCenter

    ' Thin gray table grid.
    ApplyThinGrayBorders tableRange

    ' Header style from the approved black-and-white workbook.
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastTableCol))
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(231, 231, 231)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ShrinkToFit = False
    End With

    ' Header-driven formatting makes the macro independent of weekly column count.
    For c = 1 To lastTableCol
        h = NormalizeHeader(ws.Cells(headerRow, c).Value2)
        Set colRange = ws.Range(ws.Cells(dataFirstRow, c), ws.Cells(lastRow, c))
        FormatDataColumn colRange, h
        ws.Columns(c).ColumnWidth = PreferredColumnWidth(h)
    Next c

    ' Maker values are bold and centered.
    With ws.Range(ws.Cells(dataFirstRow, makerCol), ws.Cells(lastRow, makerCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' Medium black top border whenever Maker changes, including the first data row.
    previousMaker = vbNullString
    For r = dataFirstRow To lastRow
        currentMaker = Trim$(CStr(ws.Cells(r, makerCol).Value2))
        If Len(currentMaker) > 0 Then
            If r = dataFirstRow Or StrComp(currentMaker, previousMaker, vbTextCompare) <> 0 Then
                Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastTableCol))
                With rowRange.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(0, 0, 0)
                End With
            End If
            previousMaker = currentMaker
        End If
    Next r

    ' Filter and freeze panes follow the printable table, not helper columns.
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, printLastCol)).AutoFilter

    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = headerRow
        .FreezePanes = True
        .DisplayGridlines = False
    End With

    ' Exact print settings verified from the approved workbook.
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, printLastCol)).Address
        .PrintTitleRows = "$" & headerRow & ":$" & headerRow
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .BlackAndWhite = True
        .Draft = False
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.45)
        .BottomMargin = Application.InchesToPoints(0.45)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        .CenterHorizontally = False
        .CenterVertically = False
        .PrintGridlines = False
        .PrintHeadings = False
    End With

    completed = True

CleanExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If completed And Environ$("AIR_PRINT_SILENT") <> "1" Then
        MsgBox "AIR black-and-white print formatting is complete." & vbCrLf & _
               "Sheet: " & ws.Name & vbCrLf & _
               "Rows: " & headerRow & "-" & lastRow & vbCrLf & _
               "Print area: " & ws.PageSetup.PrintArea, vbInformation, "AIR Print Format"
    End If
    Exit Sub

CleanFail:
    errText = Err.Description
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0
    If Environ$("AIR_PRINT_SILENT") = "1" Then
        Err.Raise vbObjectError + 1099, "Air_Print_Grayscale", errText
    Else
        MsgBox "Formatting stopped: " & errText, vbExclamation, "AIR Print Format"
    End If
End Sub


Private Sub ApplyThinGrayBorders(ByVal target As Range)
    Dim borderIndex As Variant
    For Each borderIndex In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        With target.Borders(borderIndex)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(166, 166, 166)
        End With
    Next borderIndex
End Sub


Private Sub FormatDataColumn(ByVal target As Range, ByVal headerText As String)
    target.WrapText = False
    target.ShrinkToFit = False

    Select Case headerText
        Case "FAC"
            target.HorizontalAlignment = xlCenter
        Case "ITEM CD", "ITEM CODE", "PART NAME", "REMARK"
            target.HorizontalAlignment = xlLeft
        Case "MAKER"
            target.HorizontalAlignment = xlCenter
        Case "PURCHASE REQUEST"
            target.HorizontalAlignment = xlRight
            target.NumberFormat = "#,##0"
        Case "BALANCE", "BALANCE 1", "B-R", "STOCK", KoreanStockHeader()
            target.HorizontalAlignment = xlCenter
            target.NumberFormat = "#,##0"
        Case "SHORT DEL.", "SHORT DEL", "SHORT DELIVERY"
            target.HorizontalAlignment = xlCenter
            target.NumberFormat = "dd-mmm-yy"
        Case "NEED ETD"
            target.HorizontalAlignment = xlCenter
            target.NumberFormat = "[$-409]d-mmm-yy;@"
            target.Font.Bold = True
        Case Else
            target.HorizontalAlignment = xlCenter
    End Select
End Sub


Private Function PreferredColumnWidth(ByVal headerText As String) As Double
    Select Case headerText
        Case "FAC": PreferredColumnWidth = 7
        Case "ITEM CD", "ITEM CODE": PreferredColumnWidth = 16
        Case "MAKER": PreferredColumnWidth = 15
        Case "PART NAME": PreferredColumnWidth = 43
        Case "PURCHASE REQUEST": PreferredColumnWidth = 16
        Case "BALANCE", "BALANCE 1", "STOCK": PreferredColumnWidth = 12
        Case "B-R", KoreanStockHeader(): PreferredColumnWidth = 10
        Case "PALLET N", "PALLET NO": PreferredColumnWidth = 12
        Case "SHORT DEL.", "SHORT DEL", "SHORT DELIVERY", "NEED ETD": PreferredColumnWidth = 13
        Case "REMARK": PreferredColumnWidth = 18
        Case Else
            If IsMatchingHelperHeader(headerText) Then
                PreferredColumnWidth = 15
            Else
                PreferredColumnWidth = 12
            End If
    End Select
End Function


Private Function NormalizeHeader(ByVal value As Variant) As String
    Dim s As String
    s = CStr(value)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ChrW(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeHeader = UCase$(Trim$(s))
End Function


Private Function IsMatchingHelperHeader(ByVal headerText As String) As Boolean
    Dim koreanMatching As String
    koreanMatching = ChrW(&HB9E4) & ChrW(&HCE6D)
    IsMatchingHelperHeader = (InStr(1, headerText, "MATCHING", vbTextCompare) > 0) Or _
                             (InStr(1, headerText, koreanMatching, vbTextCompare) > 0)
End Function


Private Function KoreanStockHeader() As String
    KoreanStockHeader = ChrW(&HC7AC) & ChrW(&HACE0)
End Function


