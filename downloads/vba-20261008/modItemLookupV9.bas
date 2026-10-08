Attribute VB_Name = "modItemLookupV9"
Option Explicit



Private Const CF_SEP As String = "|"

Private gCFAlias As Object



Public Sub VLOOKUP()

    Dim targetWb As Workbook, targetWs As Worksheet

    Dim headerRow As Long, itemCol As Long, facCol As Long

    Dim balanceCol As Long, stockCol As Long, inboundCol As Long

    Dim balanceData As Object, stockData As Object, inboundData As Object

    Dim lastRow As Long, r As Long

    Dim itemText As String, facText As String, lookupKey As String

    Dim filledBalance As Long, filledStock As Long, filledInbound As Long

    Dim missingBalance As Long, missingStock As Long, missingInbound As Long

    Dim sourceBalance As Long, sourceStock As Long, sourceInbound As Long

    Dim oldCalc As XlCalculation

    Dim balanceSources As String, stockSources As String, inboundSources As String

    Dim missingSources As String



    If ActiveWorkbook Is Nothing Or ActiveSheet Is Nothing Then Exit Sub

    Set targetWb = ActiveWorkbook

    Set targetWs = ActiveSheet



    If Not CF_FindTargetColumns(targetWs, headerRow, itemCol, facCol, balanceCol, stockCol, inboundCol) Then

        MsgBox "ITEM CD and a BALANCE/STOCK/INBOUND column were not found in the active sheet.", vbExclamation

        Exit Sub

    End If



    oldCalc = Application.Calculation

    Application.ScreenUpdating = False

    Application.EnableEvents = False

    Application.Calculation = xlCalculationManual

    On Error GoTo Failed



    Set balanceData = CreateObject("Scripting.Dictionary")

    Set stockData = CreateObject("Scripting.Dictionary")

    Set inboundData = CreateObject("Scripting.Dictionary")

    Set gCFAlias = CreateObject("Scripting.Dictionary")

    balanceData.CompareMode = vbTextCompare

    stockData.CompareMode = vbTextCompare

    inboundData.CompareMode = vbTextCompare

    gCFAlias.CompareMode = vbTextCompare



    sourceBalance = CF_LoadOpenBalanceSources(targetWb, targetWs, balanceData, balanceSources)

    sourceStock = CF_LoadOpenNamedSources(targetWb, targetWs, "STOCK", stockData, True, stockSources)

    sourceInbound = CF_LoadOpenNamedSources(targetWb, targetWs, "INBOUND", inboundData, False, inboundSources)



    If balanceCol > 0 And balanceSources = "" Then missingSources = missingSources & "BALANCE, "

    If stockCol > 0 And stockSources = "" Then missingSources = missingSources & "STOCK, "

    If inboundCol > 0 And inboundSources = "" Then missingSources = missingSources & "INBOUND, "

    If missingSources <> "" Then

        missingSources = Left$(missingSources, Len(missingSources) - 2)

MsgBox "Source not found in this Excel instance: " & missingSources & vbCrLf & vbCrLf & "Open the target and source files from the same Excel window, then try again.", vbExclamation

        GoTo CleanExit

    End If



    lastRow = targetWs.Cells(targetWs.Rows.Count, itemCol).End(xlUp).Row

    For r = headerRow + 1 To lastRow

        itemText = CF_NormalizeItem(targetWs.Cells(r, itemCol).Value2)

        If facCol > 0 Then facText = CF_NormalizeFac(targetWs.Cells(r, facCol).Value2) Else facText = ""

        If itemText <> "" Then

            lookupKey = CF_MakeKey(facText, itemText)

            If balanceCol > 0 Then CF9_WriteResult targetWs.Cells(r, balanceCol), balanceData, lookupKey, filledBalance, missingBalance, False

            If stockCol > 0 Then CF9_WriteResult targetWs.Cells(r, stockCol), stockData, lookupKey, filledStock, missingStock, True

            If inboundCol > 0 Then CF9_WriteResult targetWs.Cells(r, inboundCol), inboundData, lookupKey, filledInbound, missingInbound, True

        End If

    Next r



    If balanceCol > 0 Then targetWs.Range(targetWs.Cells(headerRow + 1, balanceCol), targetWs.Cells(lastRow, balanceCol)).NumberFormat = "#,##0;[Red]-#,##0;0"

    If stockCol > 0 Then targetWs.Range(targetWs.Cells(headerRow + 1, stockCol), targetWs.Cells(lastRow, stockCol)).NumberFormat = "#,##0;[Red]-#,##0;0"

    If inboundCol > 0 Then targetWs.Range(targetWs.Cells(headerRow + 1, inboundCol), targetWs.Cells(lastRow, inboundCol)).NumberFormat = "#,##0;[Red]-#,##0;0"



MsgBox "Auto fill completed." & vbCrLf & "Balance: " & filledBalance & " filled / " & missingBalance & " not found" & vbCrLf & "Stock: " & filledStock & " filled / " & missingStock & " not found" & vbCrLf & "Inbound: " & filledInbound & " filled / " & missingInbound & " not found" & vbCrLf & vbCrLf & "Source rows read: " & sourceBalance & " / " & sourceStock & " / " & sourceInbound & vbCrLf & vbCrLf & "BAL: " & CF_SourceDisplay(balanceSources) & vbCrLf & "STO: " & CF_SourceDisplay(stockSources) & vbCrLf & "IN: " & CF_SourceDisplay(inboundSources), vbInformation



CleanExit:

    Application.Calculation = oldCalc

    Application.EnableEvents = True

    Application.ScreenUpdating = True

    Exit Sub

Failed:

    MsgBox "Auto fill failed: " & Err.Description, vbExclamation

    Resume CleanExit

End Sub



Public Sub VLOOKUP_¹öÆ°()

    Dim ws As Worksheet, shp As Shape, anchor As Range

    If ActiveSheet Is Nothing Then Exit Sub

    Set ws = ActiveSheet

    CF_DeleteShape ws, "btnItemAutoFill"

    Set anchor = ws.Cells(1, ws.UsedRange.Column + ws.UsedRange.Columns.Count + 1)

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, 150, 32)

    shp.Name = "btnItemAutoFill"

    shp.TextFrame2.TextRange.text = "ITEM AUTO FILL"

    shp.OnAction = "VLOOKUP"

    MsgBox "One-click button installed on the active sheet.", vbInformation

End Sub



Private Function CF_LoadOpenBalanceSources(ByVal targetWb As Workbook, ByVal targetWs As Worksheet, ByVal targetData As Object, ByRef sourceNames As String) As Long

    Dim wb As Workbook, loadedRows As Long

    For Each wb In Application.Workbooks

        If CF_IsUsableSourceWorkbook(wb, targetWb) Then

            loadedRows = CF_LoadBestSource(wb, targetWs, "BALANCE", targetData, True)

            If loadedRows > 0 Then

                CF_LoadOpenBalanceSources = CF_LoadOpenBalanceSources + loadedRows

                CF_AddSourceName sourceNames, wb.Name

            End If

        End If

    Next wb

End Function



Private Function CF_LoadOpenNamedSources(ByVal targetWb As Workbook, ByVal targetWs As Worksheet, ByVal sourceKind As String, ByVal targetData As Object, ByVal saveAliases As Boolean, ByRef sourceNames As String) As Long

    Dim wb As Workbook, combinedWb As Workbook, item As Variant

    Dim specificWbs As Collection, loadedRows As Long, facText As String, dataType As String

    Set specificWbs = New Collection



    For Each wb In Application.Workbooks

        If CF_IsUsableSourceWorkbook(wb, targetWb) Then

            If CF_NameMatchesSource(wb.Name, sourceKind) Then

                facText = CF_FacFromText(wb.Name)

                If facText = "" Then

                    Set combinedWb = wb

                Else

                    specificWbs.Add wb

                End If

            End If

        End If

    Next wb



    If UCase$(sourceKind) = "STOCK" Then dataType = "STOCKFILE" Else dataType = "INBOUND"



    If Not combinedWb Is Nothing Then

        loadedRows = CF_LoadBestSource(combinedWb, targetWs, dataType, targetData, saveAliases)

        If loadedRows > 0 Then

            CF_LoadOpenNamedSources = loadedRows

            CF_AddSourceName sourceNames, combinedWb.Name

        End If

        Exit Function

    End If



    If specificWbs.Count > 0 Then

        For Each item In specificWbs

            Set wb = item

            loadedRows = CF_LoadBestSource(wb, targetWs, dataType, targetData, saveAliases)

            If loadedRows > 0 Then

                CF_LoadOpenNamedSources = CF_LoadOpenNamedSources + loadedRows

                CF_AddSourceName sourceNames, wb.Name

            End If

        Next item

        Exit Function

    End If



End Function



Private Function CF_IsUsableSourceWorkbook(ByVal wb As Workbook, ByVal targetWb As Workbook) As Boolean

    If wb Is targetWb Then Exit Function

    If wb Is ThisWorkbook Then Exit Function

    If wb.IsAddin Then Exit Function

    CF_IsUsableSourceWorkbook = True

End Function



Private Function CF_NameMatchesSource(ByVal workbookName As String, ByVal sourceKind As String) As Boolean

    Dim nameText As String

    nameText = UCase$(workbookName)

    Select Case UCase$(sourceKind)

        Case "STOCK"

            CF_NameMatchesSource = (InStr(1, nameText, "STO", vbTextCompare) > 0 Or InStr(1, nameText, CF_W(51116, 44256), vbTextCompare) > 0 Or InStr(1, nameText, CF_W(52636, 44256), vbTextCompare) > 0)

        Case "INBOUND"

            CF_NameMatchesSource = (InStr(1, nameText, "INBOUND", vbTextCompare) > 0 Or InStr(1, nameText, CF_W(51077, 44256), vbTextCompare) > 0)

    End Select

End Function



Private Sub CF_AddSourceName(ByRef listText As String, ByVal workbookName As String)

    If listText <> "" Then listText = listText & ", "

    listText = listText & workbookName

End Sub



Private Function CF_SourceDisplay(ByVal listText As String) As String

    If listText = "" Then CF_SourceDisplay = "(none)" Else CF_SourceDisplay = listText

End Function



Private Function CF_FindTargetColumns(ByVal ws As Worksheet, ByRef bestHeaderRow As Long, ByRef bestItemCol As Long, ByRef bestFacCol As Long, ByRef bestBalanceCol As Long, ByRef bestStockCol As Long, ByRef bestInboundCol As Long) As Boolean

    Dim rowNo As Long, colNo As Long, maxRow As Long, maxCol As Long

    Dim itemCol As Long, facCol As Long, balanceCol As Long, stockCol As Long, inboundCol As Long

    Dim headerText As String, score As Long, bestScore As Long



    maxRow = WorksheetFunction.Min(10, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)

    maxCol = WorksheetFunction.Min(150, ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)

    For rowNo = 1 To maxRow

        itemCol = 0: facCol = 0: balanceCol = 0: stockCol = 0: inboundCol = 0

        For colNo = 1 To maxCol

            headerText = CF_NormalizeHeader(ws.Cells(rowNo, colNo).Value2)

            If CF_IsTargetItemHeader(headerText) Then itemCol = colNo

            If CF_IsFacHeader(headerText) Then facCol = colNo

            If CF_IsTargetBalanceHeader(headerText) Then balanceCol = colNo

            If CF_IsTargetStockHeader(headerText) Then stockCol = colNo

            If CF_IsTargetInboundHeader(headerText) Then inboundCol = colNo

        Next colNo

        score = 0

        If itemCol > 0 Then score = score + 10

        If facCol > 0 Then score = score + 2

        If balanceCol > 0 Then score = score + 3

        If stockCol > 0 Then score = score + 3

        If inboundCol > 0 Then score = score + 3

        If score > bestScore Then

            bestScore = score

            bestHeaderRow = rowNo

            bestItemCol = itemCol

            bestFacCol = facCol

            bestBalanceCol = balanceCol

            bestStockCol = stockCol

            bestInboundCol = inboundCol

        End If

    Next rowNo

    CF_FindTargetColumns = (bestItemCol > 0 And (bestBalanceCol > 0 Or bestStockCol > 0 Or bestInboundCol > 0))

End Function



Private Function CF_LoadBestSource(ByVal wb As Workbook, ByVal targetWs As Worksheet, ByVal dataType As String, ByVal targetData As Object, ByVal saveAliases As Boolean) As Long

    Dim ws As Worksheet, bestWs As Worksheet

    Dim headerRow As Long, itemCol As Long, altItemCol As Long, facCol As Long, qtyCol As Long

    Dim bestHeaderRow As Long, bestItemCol As Long, bestAltItemCol As Long, bestFacCol As Long, bestQtyCol As Long

    Dim score As Double, bestScore As Double

    Dim bucket As Variant, buckets As Variant, inferredFac As String



    buckets = Array("F1", "F2", "F5", "*")

    For Each bucket In buckets

        Set bestWs = Nothing

        bestScore = 0

        bestHeaderRow = 0: bestItemCol = 0: bestAltItemCol = 0: bestFacCol = 0: bestQtyCol = 0

        For Each ws In wb.Worksheets

            If Not (ws Is targetWs) Then

                inferredFac = CF_FacFromText(ws.Name)

                If inferredFac = "" Then inferredFac = CF_FacFromText(wb.Name)

                If inferredFac = "" Then inferredFac = "*"

                If inferredFac = CStr(bucket) Then

                    score = 0

                    If CF_FindSourceColumns(ws, wb.Name & " " & ws.Name, dataType, headerRow, itemCol, altItemCol, facCol, qtyCol, score) Then

                        If score > bestScore Then

                            bestScore = score

                            Set bestWs = ws

                            bestHeaderRow = headerRow

                            bestItemCol = itemCol

                            bestAltItemCol = altItemCol

                            bestFacCol = facCol

                            bestQtyCol = qtyCol

                        End If

                    End If

                End If

            End If

        Next ws

        If Not bestWs Is Nothing Then

            CF_LoadBestSource = CF_LoadBestSource + CF_LoadSourceSheet(bestWs, bestHeaderRow, bestItemCol, bestAltItemCol, bestFacCol, bestQtyCol, CStr(bucket), targetData, saveAliases)

        End If

    Next bucket

End Function



Private Function CF_LoadSourceSheet(ByVal sourceWs As Worksheet, ByVal headerRow As Long, ByVal itemCol As Long, ByVal altItemCol As Long, ByVal facCol As Long, ByVal qtyCol As Long, ByVal forcedFac As String, ByVal targetData As Object, ByVal saveAliases As Boolean) As Long

    Dim lastRow As Long, r As Long, rowsLoaded As Long

    Dim facText As String, itemText As String

    Dim primaryValue As Variant, alternateValue As Variant, qtyValue As Variant



    If forcedFac = "*" Then forcedFac = ""

    lastRow = sourceWs.Cells(sourceWs.Rows.Count, itemCol).End(xlUp).Row

    For r = headerRow + 1 To lastRow

        primaryValue = sourceWs.Cells(r, itemCol).Value2

        If altItemCol > 0 Then alternateValue = sourceWs.Cells(r, altItemCol).Value2 Else alternateValue = ""

        If facCol > 0 Then facText = CF_NormalizeFac(sourceWs.Cells(r, facCol).Value2) Else facText = ""

        If facText = "" Then facText = forcedFac

        itemText = CF_ResolveItem(primaryValue, alternateValue, facText)

        qtyValue = sourceWs.Cells(r, qtyCol).Value2

        If itemText <> "" And IsNumeric(qtyValue) Then

            CF_AddQuantity targetData, facText, itemText, CDbl(qtyValue)

            If saveAliases Then

                CF_SaveAlias facText, CF_NormalizeItem(primaryValue), itemText

                CF_SaveAlias facText, CF_NormalizeItem(alternateValue), itemText

            End If

            rowsLoaded = rowsLoaded + 1

        End If

    Next r

    CF_LoadSourceSheet = rowsLoaded

End Function



Private Function CF_FindSourceColumns(ByVal ws As Worksheet, ByVal workbookName As String, ByVal dataType As String, ByRef bestHeaderRow As Long, ByRef bestItemCol As Long, ByRef bestAltItemCol As Long, ByRef bestFacCol As Long, ByRef bestQtyCol As Long, ByRef bestScore As Double) As Boolean

    Dim rowNo As Long, colNo As Long, maxRow As Long, maxCol As Long

    Dim itemCol As Long, altItemCol As Long, facCol As Long, qtyCol As Long

    Dim itemScore As Long, secondItemScore As Long, facScore As Long, qtyScore As Long

    Dim thisItemScore As Long, thisFacScore As Long, thisQtyScore As Long

    Dim headerText As String, score As Double



    maxRow = WorksheetFunction.Min(10, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)

    maxCol = WorksheetFunction.Min(100, ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)

    For rowNo = 1 To maxRow

        itemCol = 0: altItemCol = 0: facCol = 0: qtyCol = 0

        itemScore = 0: secondItemScore = 0: facScore = 0: qtyScore = 0

        For colNo = 1 To maxCol

            headerText = CF_NormalizeHeader(ws.Cells(rowNo, colNo).Value2)

            thisItemScore = CF_SourceItemScore(headerText)

            If thisItemScore > itemScore Then

                secondItemScore = itemScore

                altItemCol = itemCol

                itemScore = thisItemScore

                itemCol = colNo

            ElseIf thisItemScore > secondItemScore And colNo <> itemCol Then

                secondItemScore = thisItemScore

                altItemCol = colNo

            End If

            thisFacScore = CF_FacHeaderScore(headerText)

            If thisFacScore > facScore Then facScore = thisFacScore: facCol = colNo

            thisQtyScore = CF_SourceQtyScore(headerText, dataType, workbookName)

            If thisQtyScore > qtyScore Then qtyScore = thisQtyScore: qtyCol = colNo

        Next colNo

        If itemCol > 0 And qtyCol > 0 Then

            score = itemScore + facScore + qtyScore + WorksheetFunction.Min(ws.UsedRange.Rows.Count, 2000) / 100#

            If score > bestScore Then

                bestScore = score

                bestHeaderRow = rowNo

                bestItemCol = itemCol

                bestAltItemCol = altItemCol

                bestFacCol = facCol

                bestQtyCol = qtyCol

            End If

        End If

    Next rowNo

    CF_FindSourceColumns = (bestItemCol > 0 And bestQtyCol > 0)

End Function



Private Sub CF9_WriteResult(ByVal outputCell As Range, ByVal sourceData As Object, ByVal lookupKey As String, ByRef filledCount As Long, ByRef missingCount As Long, ByVal missingAsZero As Boolean)

    Dim allKey As String, itemText As String

    If sourceData.exists(lookupKey) Then

        outputCell.value = CDbl(sourceData(lookupKey))

        outputCell.Interior.Pattern = xlNone

        filledCount = filledCount + 1

        Exit Sub

    End If

    itemText = Mid$(lookupKey, InStr(1, lookupKey, CF_SEP, vbBinaryCompare) + 1)

    allKey = "*" & CF_SEP & itemText

    If Left$(lookupKey, 1) = "*" And sourceData.exists(allKey) Then

        outputCell.value = CDbl(sourceData(allKey))

        outputCell.Interior.Pattern = xlNone

        filledCount = filledCount + 1

    Else

        If missingAsZero Then

            outputCell.value = 0

            outputCell.Interior.Pattern = xlNone

        Else

            outputCell.value = CVErr(xlErrNA)

            outputCell.Interior.Color = RGB(255, 235, 156)

        End If

        missingCount = missingCount + 1

    End If

End Sub



Private Function CF_IsTargetItemHeader(ByVal textValue As String) As Boolean

    CF_IsTargetItemHeader = (textValue = "ITEMCD" Or textValue = "ITEMCODE" Or textValue = CF_HCustomerItem())

End Function



Private Function CF_IsTargetBalanceHeader(ByVal textValue As String) As Boolean

    CF_IsTargetBalanceHeader = (textValue = "BAL" Or textValue = "BALANCE" Or textValue = CF_W(48156, 46976, 49828) Or textValue = CF_HRemainQty())

End Function



Private Function CF_IsTargetStockHeader(ByVal textValue As String) As Boolean

    CF_IsTargetStockHeader = (textValue = "STO" Or textValue = "STOCK" Or textValue = "STK" Or textValue = CF_W(49828, 53441) Or textValue = CF_W(51116, 44256) Or textValue = CF_HStockQty())

End Function



Private Function CF_IsTargetInboundHeader(ByVal textValue As String) As Boolean

    CF_IsTargetInboundHeader = (textValue = "IN" Or textValue = "INBOUND" Or textValue = "INBOUNDQTY" Or textValue = CF_W(51077, 44256) Or textValue = CF_HInboundQty())

End Function



Private Function CF_IsFacHeader(ByVal textValue As String) As Boolean

    CF_IsFacHeader = (CF_FacHeaderScore(textValue) > 0)

End Function



Private Function CF_SourceItemScore(ByVal textValue As String) As Long

    If textValue = CF_HOrderCustomerItem() Then

        CF_SourceItemScore = 70

    ElseIf textValue = CF_HCustomerItem() Or textValue = "ITEMCD" Or textValue = "ITEMCODE" Then

        CF_SourceItemScore = 60

    ElseIf textValue = CF_HInternalItem() Then

        CF_SourceItemScore = 40

    End If

End Function



Private Function CF_FacHeaderScore(ByVal textValue As String) As Long

    If textValue = "FAC" Or textValue = "FACTORY" Or textValue = CF_W(44277, 51109) Then

        CF_FacHeaderScore = 40

    ElseIf textValue = CF_W(49688, 51452, 52376) Then

        CF_FacHeaderScore = 40

    ElseIf textValue = CF_W(52636, 44256, 52376, 53076, 46300) Then

        CF_FacHeaderScore = 35

    ElseIf textValue = CF_W(44144, 47000, 52376, 53076, 46300) Then

        CF_FacHeaderScore = 30

    ElseIf textValue = CF_W(48708, 44256) Then

        CF_FacHeaderScore = 15

    End If

End Function



Private Function CF_SourceQtyScore(ByVal textValue As String, ByVal dataType As String, ByVal workbookName As String) As Long

    Dim nameText As String

    nameText = UCase$(workbookName)

    Select Case UCase$(dataType)

        Case "BALANCE"

            If textValue = CF_HRemainQty() Then CF_SourceQtyScore = 90

            If textValue = CF_HRemainQtyTotal() Then CF_SourceQtyScore = 85

            If textValue = "BALANCE" Or textValue = "BAL" Then CF_SourceQtyScore = 60

        Case "STOCK"

            If textValue = CF_HStockQty() Then CF_SourceQtyScore = 90

            If textValue = "STOCK" Or textValue = "STO" Or textValue = "STK" Then CF_SourceQtyScore = 60

        Case "STOCKFILE"

            If InStr(1, nameText, "STOCK", vbTextCompare) > 0 Or InStr(1, nameText, "STO", vbTextCompare) > 0 Then

                If textValue = CF_HStockQty() Then CF_SourceQtyScore = 90

                If textValue = "STOCK" Or textValue = "STO" Or textValue = "STK" Then CF_SourceQtyScore = 85

                If textValue = CF_HShipQtyTotal() Then CF_SourceQtyScore = 85

                If textValue = CF_HShipQty() Then CF_SourceQtyScore = 75

            End If

        Case "INBOUND"

            If textValue = CF_HInboundQtyTotal() Then CF_SourceQtyScore = 95

            If textValue = CF_HInboundQty() Then CF_SourceQtyScore = 90

            If textValue = "INBOUND" Or textValue = "INBOUNDQTY" Then CF_SourceQtyScore = 70

    End Select

End Function



Private Sub CF_AddQuantity(ByVal targetData As Object, ByVal facText As String, ByVal itemText As String, ByVal quantity As Double)

    CF_AddToKey targetData, "*" & CF_SEP & itemText, quantity

    If facText <> "" Then CF_AddToKey targetData, facText & CF_SEP & itemText, quantity

End Sub



Private Sub CF_AddToKey(ByVal targetData As Object, ByVal lookupKey As String, ByVal quantity As Double)

    If targetData.exists(lookupKey) Then

        targetData(lookupKey) = CDbl(targetData(lookupKey)) + quantity

    Else

        targetData.Add lookupKey, quantity

    End If

End Sub



Private Function CF_ResolveItem(ByVal primaryValue As Variant, ByVal alternateValue As Variant, ByVal facText As String) As String

    Dim primaryText As String, alternateText As String, mappedText As String

    primaryText = CF_NormalizeItem(primaryValue)

    alternateText = CF_NormalizeItem(alternateValue)

    If CF_IsCustomerCode(primaryText) Then CF_ResolveItem = primaryText: Exit Function

    mappedText = CF_GetAlias(facText, primaryText)

    If mappedText <> "" Then CF_ResolveItem = mappedText: Exit Function

    If CF_IsCustomerCode(alternateText) Then CF_ResolveItem = alternateText: Exit Function

    mappedText = CF_GetAlias(facText, alternateText)

    If mappedText <> "" Then CF_ResolveItem = mappedText: Exit Function

    If primaryText <> "" Then CF_ResolveItem = primaryText Else CF_ResolveItem = alternateText

End Function



Private Function CF_IsCustomerCode(ByVal itemText As String) As Boolean

    CF_IsCustomerCode = (Len(itemText) >= 5 And Left$(itemText, 1) = "M" And InStr(1, itemText, "-", vbBinaryCompare) > 0)

End Function



Private Sub CF_SaveAlias(ByVal facText As String, ByVal sourceItem As String, ByVal customerItem As String)

    Dim exactKey As String, allKey As String

    If sourceItem = "" Or customerItem = "" Or gCFAlias Is Nothing Then Exit Sub

    allKey = "*" & CF_SEP & sourceItem

    If Not gCFAlias.exists(allKey) Then gCFAlias.Add allKey, customerItem

    If facText <> "" Then

        exactKey = facText & CF_SEP & sourceItem

        If Not gCFAlias.exists(exactKey) Then gCFAlias.Add exactKey, customerItem

    End If

End Sub



Private Function CF_GetAlias(ByVal facText As String, ByVal sourceItem As String) As String

    Dim lookupKey As String

    If sourceItem = "" Or gCFAlias Is Nothing Then Exit Function

    If facText <> "" Then

        lookupKey = facText & CF_SEP & sourceItem

        If gCFAlias.exists(lookupKey) Then CF_GetAlias = CStr(gCFAlias(lookupKey)): Exit Function

    End If

    lookupKey = "*" & CF_SEP & sourceItem

    If gCFAlias.exists(lookupKey) Then CF_GetAlias = CStr(gCFAlias(lookupKey))

End Function



Private Function CF_MakeKey(ByVal facText As String, ByVal itemText As String) As String

    If facText = "" Then CF_MakeKey = "*" & CF_SEP & itemText Else CF_MakeKey = facText & CF_SEP & itemText

End Function



Private Function CF_NormalizeItem(ByVal value As Variant) As String

    If IsError(value) Or IsEmpty(value) Then Exit Function

    CF_NormalizeItem = UCase$(Replace(Trim$(CStr(value)), " ", ""))

End Function



Private Function CF_NormalizeHeader(ByVal value As Variant) As String

    Dim textValue As String

    If IsError(value) Or IsEmpty(value) Then Exit Function

    textValue = UCase$(Trim$(CStr(value)))

    textValue = Replace(textValue, " ", "")

    textValue = Replace(textValue, "_", "")

    textValue = Replace(textValue, "-", "")

    textValue = Replace(textValue, ".", "")

    textValue = Replace(textValue, "'", "")

    textValue = Replace(textValue, ChrW(9651), "")

    textValue = Replace(textValue, ChrW(9661), "")

    CF_NormalizeHeader = textValue

End Function



Private Function CF_NormalizeFac(ByVal value As Variant) As String

    If IsError(value) Or IsEmpty(value) Then Exit Function

    CF_NormalizeFac = CF_FacFromText(CStr(value))

End Function



Private Function CF_FacFromText(ByVal value As String) As String

    Dim textValue As String

    textValue = UCase$(Trim$(value))

    Select Case textValue

        Case "1010", "F1": CF_FacFromText = "F1": Exit Function

        Case "1060", "F2": CF_FacFromText = "F2": Exit Function

        Case "1150", "F5": CF_FacFromText = "F5": Exit Function

    End Select

    If InStr(1, textValue, "F1", vbTextCompare) > 0 Then CF_FacFromText = "F1": Exit Function

    If InStr(1, textValue, "F2", vbTextCompare) > 0 Then CF_FacFromText = "F2": Exit Function

    If InStr(1, textValue, "F5", vbTextCompare) > 0 Then CF_FacFromText = "F5"

End Function



Private Sub CF_DeleteShape(ByVal ws As Worksheet, ByVal shapeName As String)

    Dim shp As Shape

    For Each shp In ws.Shapes

        If StrComp(shp.Name, shapeName, vbTextCompare) = 0 Then shp.Delete: Exit For

    Next shp

End Sub



Private Function CF_HCustomerItem() As String

    CF_HCustomerItem = CF_W(44144, 47000, 52376, 54408, 47785, 53076, 46300)

End Function



Private Function CF_HOrderCustomerItem() As String

    CF_HOrderCustomerItem = CF_W(49688, 51452, 52376, 54408, 47785, 53076, 46300)

End Function



Private Function CF_HInternalItem() As String

    CF_HInternalItem = CF_W(54408, 47785, 53076, 46300)

End Function



Private Function CF_HRemainQty() As String

    CF_HRemainQty = CF_W(48120, 45225, 49688, 47049)

End Function



Private Function CF_HRemainQtyTotal() As String

    CF_HRemainQtyTotal = CF_W(48120, 45225, 49688, 47049, 54633, 44228)

End Function



Private Function CF_HStockQty() As String

    CF_HStockQty = CF_W(51116, 44256, 49688, 47049)

End Function



Private Function CF_HInboundQty() As String

    CF_HInboundQty = CF_W(51077, 44256, 49688, 47049)

End Function



Private Function CF_HInboundQtyTotal() As String

    CF_HInboundQtyTotal = CF_W(51077, 44256, 49688, 47049, 54633, 44228)

End Function



Private Function CF_HShipQty() As String

    CF_HShipQty = CF_W(52636, 44256, 49688, 47049)

End Function



Private Function CF_HShipQtyTotal() As String

    CF_HShipQtyTotal = CF_W(52636, 44256, 49688, 47049, 54633, 44228)

End Function



Private Function CF_W(ParamArray codes() As Variant) As String

    Dim i As Long, result As String

    For i = LBound(codes) To UBound(codes)

        result = result & ChrW(CLng(codes(i)))

    Next i

    CF_W = result

End Function
