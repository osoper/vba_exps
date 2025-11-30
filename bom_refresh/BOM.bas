Attribute VB_Name = "BOM"
Sub BOM()
' asks the user which check type(s), fleet(s), and location(s) to consider
' then, records/organizes part usage based on individual checks that match these filters
' finally, displays analytics for each PPN based on its usage

    Dim wsParts As Worksheet, wsBOMdata As Worksheet, wsBOM As Worksheet, wsParam As Worksheet
    Dim lastRowParts As Long, lastRowBOMdata As Long, lastRowBOM As Long, lastRowLoc As Long, lastRowFlt As Long, lastRowChk As Long
    Dim lastColBOMdataLetter As String
    Dim checkNum As Integer, checkCol As Integer
    Dim foundCheck As Range, foundPart As Range, rng As Range

    Set wsParts = ThisWorkbook.Sheets("Transactions")
    Set wsBOMdata = ThisWorkbook.Sheets("Usage")
    Set wsBOM = ThisWorkbook.Sheets("BOM")
    Set wsParam = ThisWorkbook.Sheets("Parameters")
    
    ' show the BOM filter selection menu
    UserForm1.Show
    
    ' clear and format the usage sheet
    wsBOMdata.Cells.ClearContents
    wsBOMdata.Cells.NumberFormat = "0;-0;""-"""
    wsBOMdata.Columns(1).NumberFormat = "@"
    wsBOMdata.Cells(1, 1).Value = "PPN"
    wsBOMdata.Cells(1, 2).Value = "Description"
    wsBOMdata.Cells(1, 3).Value = "Part Type"
    
    ' clear the BOM sheet
    lastRowBOM = wsBOM.Cells(wsBOM.Rows.Count, 1).End(xlUp).Row
    wsBOM.Range("A2:L" & lastRowBOM + 1).ClearContents

    ' get last rows
    lastRowParts = wsParts.Cells(wsParts.Rows.Count, 1).End(xlUp).Row
    lastRowBOMdata = 1
    lastRowLoc = wsParam.Cells(wsParam.Rows.Count, 5).End(xlUp).Row
    lastRowFlt = wsParam.Cells(wsParam.Rows.Count, 6).End(xlUp).Row
    lastRowChk = wsParam.Cells(wsParam.Rows.Count, 7).End(xlUp).Row
    
    ' go through the transactions sheet
    checkNum = 0
    For r = 2 To lastRowParts
        For l = 17 To lastRowLoc
            For f = 17 To lastRowFlt
                For c = 17 To lastRowChk
                
                    ' consider a transaction only if it matches the BOM filters from the Parameters worksheet
                    If wsParts.Cells(r, 5).Value = wsParam.Cells(l, 5).Value And wsParts.Cells(r, 7).Value = wsParam.Cells(f, 6).Value And _
                        wsParts.Cells(r, 8).Value = wsParam.Cells(c, 7).Value Then
                        Set foundPart = wsBOMdata.Range("A2:A" & lastRowBOMdata + 1).Find(What:=wsParts.Cells(r, 1).Value, LookAt:=xlWhole, MatchCase:=False)
                        Set foundCheck = wsBOMdata.Range("1:1").Find(What:=wsParts.Cells(r, 10).Value, LookAt:=xlWhole, MatchCase:=False)
                        
                        ' if the check ID doesn't exist in the usage sheet, make a new column with the new check ID
                        If foundCheck Is Nothing Then
                            checkNum = checkNum + 1
                            wsBOMdata.Cells(1, checkNum + 3).Value = wsParts.Cells(r, 10).Value
                            checkCol = checkNum + 3
                        
                        ' otherwise, go to the existing column
                        Else
                            checkCol = foundCheck.Column
                        End If
                        
                        ' if the PPN doesn't exists in the usage sheet, make a new row with the new PPN
                        If foundPart Is Nothing Then
                            lastRowBOMdata = lastRowBOMdata + 1
                            wsBOMdata.Range("A" & lastRowBOMdata & ":C" & lastRowBOMdata).Value = wsParts.Range("A" & r & ":C" & r).Value
                            
                            ' record the quantity used
                            wsBOMdata.Cells(lastRowBOMdata, checkCol).Value = wsParts.Cells(r, 4).Value
                            
                        ' otherwise, go to the existing row
                        Else
                            
                            ' add the quantity used
                            wsBOMdata.Cells(foundPart.Row, checkCol).Value = wsBOMdata.Cells(foundPart.Row, checkCol).Value + wsParts.Cells(r, 4).Value
                        End If
                    End If
                Next c
            Next f
        Next l
    Next r
    
    ' check if any checks for the selected filters were found before continuing
    If IsEmpty(wsBOMdata.Cells(2, 1).Value) Then
        MsgBox "No checks found for the selected filters.", vbExclamation
        Exit Sub
    End If
    
    ' remove rows/PPNs in the usage sheet which have no use quantities > 0
    lastColBOMdataLetter = Split(Cells(1, checkNum + 3).Address(True, False), "$")(0)
    For g = lastRowBOMdata To 2 Step -1
        If Application.WorksheetFunction.CountIf(wsBOMdata.Range("D" & g & ":" & lastColBOMdataLetter & g), ">0") = 0 Then
            wsBOMdata.Rows(g).Delete Shift:=xlUp
        End If
    Next g
    
    ' replace blank and negative cells with a use quantity of 0
    lastRowBOMdata = wsBOMdata.Cells(wsBOMdata.Rows.Count, 1).End(xlUp).Row
    Set rng = wsBOMdata.Range(wsBOMdata.Cells(2, 4), wsBOMdata.Cells(lastRowBOMdata, checkNum + 3))
    ' Loop through each cell in the range
    For Each cell In rng
        If cell.Value = "" Or cell.Value < 0 Then
            cell.Value = 0
        End If
    Next cell
    
    ' format BOM worksheet
    wsBOM.Columns(1).NumberFormat = "@"
    wsBOM.Columns(5).NumberFormat = "0.0%"
    wsBOM.Range("H:I").NumberFormat = "0.0"
    wsBOM.Range("A2:C" & lastRowBOMdata).Value = wsBOMdata.Range("A2:C" & lastRowBOMdata).Value
    lastRowBOM = wsBOM.Cells(wsBOM.Rows.Count, 1).End(xlUp).Row
    wsBOM.Range("D2:D" & lastRowBOM).HorizontalAlignment = xlRight
    
    ' populate BOM worksheet with formulas, which display analytics for each PPN
    wsBOM.Range("D2:D" & lastRowBOM).Formula = "=COUNTIF(Usage!D2:" & lastColBOMdataLetter & "2, "">0"") & ""/"" & COUNT(Usage!D2:" & lastColBOMdataLetter & "2)"
    wsBOM.Range("E2:E" & lastRowBOM).Formula = "=COUNTIF(Usage!D2:" & lastColBOMdataLetter & "2, "">0"")/COUNT(Usage!D2:" & lastColBOMdataLetter & "2)"
    wsBOM.Range("F2:F" & lastRowBOM).Formula = "=MIN(Usage!D2:" & lastColBOMdataLetter & "2)"
    wsBOM.Range("G2:G" & lastRowBOM).Formula = "=MIN(FILTER(Usage!D2:" & lastColBOMdataLetter & "2, Usage!D2:" & lastColBOMdataLetter & "2 > 0))"
    wsBOM.Range("H2:H" & lastRowBOM).Formula = "=AVERAGE(Usage!D2:" & lastColBOMdataLetter & "2)"
    wsBOM.Range("I2:I" & lastRowBOM).Formula = "=SUM(Usage!D2:" & lastColBOMdataLetter & "2)/COUNTIF(Usage!D2:" & lastColBOMdataLetter & "2, "">0"")"
    wsBOM.Range("J2:J" & lastRowBOM).Formula = "=MAX(Usage!D2:" & lastColBOMdataLetter & "2)"
    wsBOM.Range("K2:K" & lastRowBOM).Formula = "=MAX(MODE.MULT(Usage!D2:" & lastColBOMdataLetter & "2))"
    wsBOM.Range("L2:L" & lastRowBOM).Formula = "=MAX(MODE.MULT(FILTER(Usage!D2:" & lastColBOMdataLetter & "2, Usage!D2:" & lastColBOMdataLetter & "2 > 0)))"
    
    MsgBox "BOM generated." & vbCrLf & checkNum & " check(s) found.", vbInformation
    
End Sub

