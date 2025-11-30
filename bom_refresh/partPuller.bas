Attribute VB_Name = "partPuller"
Sub partPuller()
' gets relevant part transaction data from SQL server and puts it in the transactions sheet

    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim strConn As String
    
    ' SQL connection credentials
    strConn_Eagle = "Driver={MySQL ODBC _._ ANSI Driver};Server=___;Database=___;PORT=___;UID=___;PWD=___;Option=3"
    
    Dim wsParts As Worksheet, wsParam As Worksheet
    Dim lastRow As Long
    Dim daysBefore As Double, daysAfter As Double
    
    Set wsParts = ThisWorkbook.Sheets("Transactions")
    Set wsParam = ThisWorkbook.Sheets("Parameters")
    
    ' get transaction date tolerance values from the Parameters sheet
    daysBefore = wsParam.Cells(12, 3).Value
    daysAfter = wsParam.Cells(13, 3).Value
    
    ' make sure they are positive integers
    If Not (daysBefore = Int(daysBefore) And daysBefore > 0 And _
        daysAfter = Int(daysAfter) And daysAfter > 0) Then
        MsgBox "Transaction Date Tolerance values must be positive integers.", vbExclamation
        Exit Sub
    End If
    
    ' clear the transactions worksheet
    lastRow = wsParts.Cells(wsParts.Rows.Count, 1).End(xlUp).Row
'    wsParts.Range("A2:J" & lastRow + 1).ClearContents
    wsParts.Columns(1).NumberFormat = "@"
    
    ' get transaction data, and use joins + left joins to only pull transactions that are relevant
    ' also add the necessary check + fleet information as additional columns
    conn.Open strConn_Eagle
    
    rs.Open "SELECT CASE WHEN D.main IS null THEN A.pn ELSE D.main END AS ppn, F.description as des, F.itemType as itp, A.qty, E.code AS loc, A.job, B.FLEET AS flt, C.CHECK AS chk, C.DATE_OUT as dto FROM tx A JOIN tmp_fleet B ON A.job=B.TAIL LEFT JOIN org E ON A.org_id=E.id JOIN tmp_Heavychk C ON C.TAIL=B.TAIL_NUM AND C.ORG=E.code AND A.tx_dt BETWEEN DATE_SUB(C.DATE_IN, INTERVAL " & daysBefore & " DAY) AND DATE_ADD(C.DATE_OUT, INTERVAL " & daysAfter & " DAY) LEFT JOIN itx_normal D ON A.pn=D.pn LEFT JOIN part F ON D.main=F.pn", conn
    wsParts.Cells(2, 1).CopyFromRecordset rs
    
    rs.Close
    conn.Close
    
    ' assign a check ID to each transaction, which is the tail # and check date
    lastRow = wsParts.Cells(wsParts.Rows.Count, 1).End(xlUp).Row
    wsParts.Range("J2:J" & lastRow).Formula = "=F2 & ""-"" & TEXT(I2, ""m/d/yyyy"")"
    
    MsgBox "Transactions refreshed.", vbInformation

End Sub


