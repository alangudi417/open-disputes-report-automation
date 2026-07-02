Sub All_in_One()

    'Part 1:

    Dim sourceWB As Workbook, targetWB As Workbook, thisWB As Workbook
    Dim sourceWS As Worksheet, rawWS As Worksheet
    Dim openWS As Worksheet, tarWS As Worksheet
    Dim todayStr As String
    Dim filePathRaw As String, filePathOutput As String
    Dim copyRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim cell As Range
    Dim ws As Worksheet
    Dim wsList As Variant
    Dim wsName As Variant
    Dim openData As Range
    Dim filterRange As Range

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set thisWB = ThisWorkbook

    ' Clear contents of "Open", "New_Largest" and "Mult_Disputes"
    With thisWB
        If SheetExists("Open", thisWB) Then .Sheets("Open").Cells.ClearContents
        If SheetExists("New_Largest", thisWB) Then .Sheets("New_Largest").Cells.ClearContents
        If SheetExists("Mult_Disputes", thisWB) Then .Sheets("Mult_Disputes").Cells.ClearContents
    End With

    todayStr = Format(Date, "mm-dd-yy")

    ' Paths
    filePathRaw = "/sample_data/RAW " & todayStr & ".xlsx"
    filePathOutput = "/sample_data/Open Disputes File/Open Disputes File " & todayStr & ".xlsm"

    Set sourceWB = Workbooks.Open(filePathRaw)
    Set sourceWS = sourceWB.Sheets("Sheet1")

    lastRow = sourceWS.Cells(sourceWS.Rows.Count, "A").End(xlUp).Row
    lastCol = sourceWS.Cells(1, sourceWS.Columns.Count).End(xlToLeft).Column
    Set copyRange = sourceWS.Range(sourceWS.Cells(1, 1), sourceWS.Cells(lastRow, lastCol))

    Set targetWB = Workbooks.Add
    Set rawWS = targetWB.Sheets(1)
    rawWS.Name = "Raw"
    copyRange.Copy Destination:=rawWS.Range("A1")

    Set openWS = targetWB.Sheets.Add(After:=rawWS)
    openWS.Name = "Open"
    Set tarWS = targetWB.Sheets.Add(After:=openWS)
    tarWS.Name = "TAR"

    With rawWS
        Dim lastRowU As Long
        lastRowU = .Cells(.Rows.Count, "U").End(xlUp).Row
        If lastRowU > 2 Then
            .Rows(lastRowU).Delete
            .Rows(lastRowU - 1).Delete
        End If
    End With

    With rawWS
        .AutoFilterMode = False
        lastRow = .Cells(.Rows.Count, "M").End(xlUp).Row

        .Range("A1").AutoFilter Field:=13, Criteria1:=Array("TAR", "DCP", "DPI"), Operator:=xlFilterValues
        .UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=tarWS.Range("A1")
        .AutoFilterMode = False

        Dim dict As Object
        Dim i As Long, val As String
        Set dict = CreateObject("Scripting.Dictionary")

        For i = 2 To lastRow
            val = Trim(.Cells(i, "M").Value)
            If val <> "" And val <> "TAR" And val <> "DCP" And val <> "DPI" Then
                If Not dict.exists(val) Then dict.Add val, Nothing
            End If
        Next i

        If dict.Count > 0 Then
            .Range("A1").AutoFilter Field:=13, Criteria1:=dict.Keys, Operator:=xlFilterValues
            .UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=openWS.Range("A1")
        Else
            openWS.Cells.Clear
        End If

        .AutoFilterMode = False
    End With

    rawWS.Delete

    wsList = Array("Open", "TAR")

    For Each wsName In wsList
        Set ws = targetWB.Sheets(wsName)
        With ws
            .Rows(1).Font.Bold = True
            .Rows(1).WrapText = True

            For Each cell In .Rows(1).Cells
                If Trim(cell.Value) = "Sales Organization" Then
                    cell.Value = "Sales Org"
                    Exit For
                End If
            Next cell

            .Columns("U").NumberFormat = "$#,##0.00"

            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            With .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With

            .Range(.Cells(1, 1), .Cells(1, lastCol)).AutoFilter
            .Columns.AutoFit

            .Columns("A").ColumnWidth = 6.1
            .Columns("B").ColumnWidth = 12
            .Columns("C").ColumnWidth = 7.1
            .Columns("D").ColumnWidth = 11.1
            .Columns("E").ColumnWidth = 8.2
            .Columns("F").ColumnWidth = 5.2
            .Columns("K").ColumnWidth = 9.1
            .Columns("L").ColumnWidth = 50
            .Columns("M").ColumnWidth = 7
            .Columns("N").ColumnWidth = 10
            .Columns("O").ColumnWidth = 15
            .Columns("Q").ColumnWidth = 7
            .Columns("R").ColumnWidth = 15
            .Columns("S").ColumnWidth = 15
            .Columns("U").ColumnWidth = 15

            .Columns("G").Hidden = True
            .Columns("H").Hidden = True
            .Columns("I").Hidden = True
            .Columns("J").Hidden = True
            .Columns("P").Hidden = True
            .Columns("T").Hidden = True
        End With
    Next wsName

    For Each wsName In wsList
        Set ws = targetWB.Sheets(wsName)
        With ws
            lastRow = .Cells(.Rows.Count, "U").End(xlUp).Row
            If lastRow > 1 Then
                .Range("A1").CurrentRegion.Sort _
                    Key1:=.Range("U2"), Order1:=xlDescending, Header:=xlYes
            End If
        End With
    Next wsName

    targetWB.SaveAs Filename:=filePathOutput, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    targetWB.Close SaveChanges:=False
    sourceWB.Close SaveChanges:=False

    MsgBox "Your Open Disputes File has been created"

    Set targetWB = Workbooks.Open(filePathOutput)
    Set openWS = targetWB.Sheets("Open")

    lastRow = openWS.Cells(openWS.Rows.Count, "A").End(xlUp).Row
    lastCol = openWS.Cells(1, openWS.Columns.Count).End(xlToLeft).Column
    Set openData = openWS.Range(openWS.Cells(1, 1), openWS.Cells(lastRow, lastCol))

    With thisWB.Sheets("Open")
        .Cells.ClearContents
        openData.Copy Destination:=.Range("A1")
    End With

    targetWB.Close SaveChanges:=False

    With thisWB.Sheets("Open")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set filterRange = .Range("A1", .Cells(lastRow, lastCol))
        filterRange.AutoFilter

        filterRange.AutoFilter Field:=1, Criteria1:=Array("0", "1"), Operator:=xlFilterValues
        filterRange.AutoFilter Field:=21, Criteria1:=">5000"

        On Error Resume Next
        filterRange.SpecialCells(xlCellTypeVisible).Copy Destination:=thisWB.Sheets("New_Largest").Range("A1")
        On Error GoTo 0

        .AutoFilterMode = False
    End With

    With thisWB.Sheets("New_Largest")
        .Columns.AutoFit
        .Columns("G").Hidden = True
        .Columns("H").Hidden = True
        .Columns("I").Hidden = True
        .Columns("J").Hidden = True
        .Columns("P").Hidden = True
        .Columns("T").Hidden = True
    End With

    MsgBox "The New Largest list has been created"

    'Part 2

    Dim wsMain As Worksheet
    Dim todayRow As Long
    Dim today As String
    Dim lastRowC As Long, countColC As Long
    Dim count1235549 As Long
    Dim sumColU As Double
    Dim col As Long, rowStart As Long, rowEnd As Long

    Set wsMain = thisWB.Sheets("Main")
    Set wsOpen = thisWB.Sheets("Open")

    today = Format(Date, "dddd")

    If today = "Monday" Then
        wsMain.Range("B12:F16").ClearContents
        For i = 12 To 16
            wsMain.Cells(i, "D").Formula = "=" & wsMain.Cells(i, "E").Address(False, False) & "-" & wsMain.Cells(i, "C").Address(False, False)
        Next i
    End If

    Select Case today
        Case "Monday": todayRow = 12
        Case "Tuesday": todayRow = 13
        Case "Wednesday": todayRow = 14
        Case "Thursday": todayRow = 15
        Case "Friday": todayRow = 16
        Case Else
            MsgBox "Today is not a working day (Monday-Friday)."
            GoTo Cleanup
    End Select

    wsMain.Cells(todayRow, 2).Value = Format(Date, "mm/dd/yyyy")
    count1235549 = Application.WorksheetFunction.CountIf(wsOpen.Range("K:K"), "1235549")
    wsMain.Cells(todayRow, 3).Value = count1235549

    With wsOpen
        lastRowC = .Cells(.Rows.Count, "C").End(xlUp).Row
        countColC = Application.WorksheetFunction.CountA(.Range("C2:C" & lastRowC))
        sumColU = Application.WorksheetFunction.Sum(.Range("U:U"))
    End With

    wsMain.Cells(todayRow, 5).Value = countColC
    wsMain.Cells(todayRow, 6).Value = sumColU

    rowStart = 13
    rowEnd = 16

    For col = 4 To 6
        For i = rowStart To rowEnd
            If wsMain.Cells(i, col).Value > wsMain.Cells(i - 1, col).Value Then
                wsMain.Cells(i, col).Font.Color = vbRed
            End If
        Next i
    Next col

    MsgBox "The table has been updated"
    
    'Part 3

    Dim multiWS As Worksheet
    Dim custRange As Range, visibleRange As Range
    Dim custCell As Range
    Dim lastRowOpen As Long
    Dim key As Variant
    Dim foundMatch As Boolean

    ' Create / reset Mult_Disputes sheet
    Set multiWS = thisWB.Worksheets("Mult_Disputes")
    multiWS.Cells.Clear

    ' Reference Open sheet
    Set openWS = thisWB.Worksheets("Open")

    ' Find last row in column A
    lastRowOpen = openWS.Cells(openWS.Rows.Count, "A").End(xlUp).Row

    ' Apply filter on Dispute Days (Column A = 0)
    If openWS.AutoFilterMode Then openWS.AutoFilterMode = False

    openWS.Range("A1:K" & lastRowOpen).AutoFilter Field:=1, Criteria1:="0"

    ' Get visible cells in Customer column (K)
    On Error Resume Next
    Set custRange = openWS.Range("K2:K" & lastRowOpen).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not custRange Is Nothing Then

        Set dict = CreateObject("Scripting.Dictionary")

        ' Count occurrences of each customer
        For Each custCell In custRange
            If custCell.Value <> "" Then
                If dict.exists(custCell.Value) Then
                    dict(custCell.Value) = dict(custCell.Value) + 1
                Else
                    dict.Add custCell.Value, 1
                End If
            End If
        Next custCell

        foundMatch = False

        ' Loop through dictionary to find customers with >=15 records
        For Each key In dict.Keys
            If dict(key) >= 15 Then
            
                foundMatch = True
            
            ' Apply additional filter for that customer
                openWS.Range("A1:K" & lastRowOpen).AutoFilter Field:=11, Criteria1:=key
            
            ' Copy visible rows
                On Error Resume Next
                Set visibleRange = openWS.Range("A1:L" & lastRowOpen).SpecialCells(xlCellTypeVisible)
                On Error GoTo 0
            
                If Not visibleRange Is Nothing Then
                    visibleRange.Copy multiWS.Cells(multiWS.Rows.Count, "A").End(xlUp).Offset(1, 0)
                End If
            End If
        Next key

    ' If no matches found, do nothing
    End If

    ' Clear filters
    If openWS.AutoFilterMode Then openWS.AutoFilterMode = False

    ' Part 4

'Dim wsMain As Worksheet
    Dim multiLastRow As Long
'    Dim descRange As Range, cell As Range
    Dim dictCust As Object
    Dim arrData() As Variant
'    Dim i As Long, j As Long
    Dim tempName As Variant, tempCount As Long
    Dim outRow As Long

' Reference sheets
'Set wsMain = thisWB.Worksheets("Main")
'Set multiWS = thisWB.Worksheets("Mult_Disputes")

    ' Clear Main table range
    wsMain.Range("H19:I23").ClearContents

    ' Find last row in column L
    multiLastRow = multiWS.Cells(multiWS.Rows.Count, "L").End(xlUp).Row

    ' Set range
    Set descRange = multiWS.Range("L2:L" & multiLastRow)

    ' Create dictionary
    Set dictCust = CreateObject("Scripting.Dictionary")

    ' Count occurrences
    For Each cell In descRange
        If cell.Value <> "" Then
            If dictCust.exists(cell.Value) Then
                dictCust(cell.Value) = dictCust(cell.Value) + 1
            Else
                dictCust.Add cell.Value, 1
            End If
        End If
    Next cell

    ' Create dictionary
    Set dictCust = CreateObject("Scripting.Dictionary")

    ' Count occurrences
    For Each cell In descRange
        If cell.Value <> "" Then
            If dictCust.exists(cell.Value) Then
                dictCust(cell.Value) = dictCust(cell.Value) + 1
            Else
                dictCust.Add cell.Value, 1
            End If
        End If
    Next cell

    ' IMPORTANT: exit if nothing qualifies
    If dictCust.Count = 0 Then GoTo NoResults

    ' Move only customers with >=15 into array
    ReDim arrData(1 To 2, 1 To dictCust.Count)
    
    i = 1
    For Each key In dictCust.Keys
        If dictCust(key) >= 15 Then
            arrData(1, i) = key
            arrData(2, i) = dictCust(key)
            i = i + 1
        End If
    Next key

    ' Resize array to actual size
    If i = 1 Then GoTo NoResults
    ReDim Preserve arrData(1 To 2, 1 To i - 1)

    ' Sort array DESCENDING by count (column 2)

    For i = 1 To UBound(arrData, 2) - 1
        For j = i + 1 To UBound(arrData, 2)
            If arrData(2, i) < arrData(2, j) Then
                tempName = arrData(1, i)
                tempCount = arrData(2, i)

                arrData(1, i) = arrData(1, j)
                arrData(2, i) = arrData(2, j)

                arrData(1, j) = tempName
                arrData(2, j) = tempCount
            End If
        Next j
    Next i

    ' Output to Main (max 5 rows: H19:I23)
    outRow = 19
    For i = 1 To UBound(arrData, 2)
        If outRow > 23 Then Exit For
        wsMain.Cells(outRow, "H").Value = arrData(1, i)
        wsMain.Cells(outRow, "I").Value = arrData(2, i)
        outRow = outRow + 1
    Next i

    GoTo Done

NoResults:
MsgBox "No customers found with 15 or more records."

Done:
    MsgBox "The macro finished successfully"

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub