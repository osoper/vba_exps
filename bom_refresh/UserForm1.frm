VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "BOM Filters"
   ClientHeight    =   6140
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   9840.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
' when OK button is clicked:

    Dim wsParam As Worksheet
    Dim r As Integer
    Dim lastRowLoc As Long, lastRowFlt As Long, lastRowChk As Long
    Dim itemsSlected1 As Boolean, itemSelected2 As Boolean, itemSelected3 As Boolean
    
    Set wsParam = ThisWorkbook.Sheets("Parameters")
        
        ' make sure at least 1 option from each BOM filter is selected
        For j = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(j) Then
                itemselected1 = True
            End If
        Next j
        For j = 0 To ListBox2.ListCount - 1
            If ListBox2.Selected(j) Then
                itemSelected2 = True
            End If
        Next j
        For j = 0 To ListBox3.ListCount - 1
            If ListBox3.Selected(j) Then
                itemSelected3 = True
            End If
        Next j
        
        If Not (itemselected1 And itemSelected2 And itemSelected3) Then
            MsgBox "Please select at least 1 filter from each list.", vbExclamation
        Else
        
            ' clear the previous BOM filters on the Parameters worksheet
            lastRowLoc = wsParam.Cells(wsParam.Rows.Count, 5).End(xlUp).Row
            lastRowFlt = wsParam.Cells(wsParam.Rows.Count, 6).End(xlUp).Row
            lastRowChk = wsParam.Cells(wsParam.Rows.Count, 7).End(xlUp).Row
            wsParam.Range("E17:E" & lastRowLoc).ClearContents
            wsParam.Range("F17:F" & lastRowFlt).ClearContents
            wsParam.Range("G17:G" & lastRowChk).ClearContents
            
            ' record the new BOM filters on the Parameters worksheet
            r = 16
            For i = 0 To (Me.ListBox1.ListCount - 1)
                If Me.ListBox1.Selected(i) Then
                r = r + 1
                    wsParam.Cells(r, 5).Value = Me.ListBox1.List(i)
                End If
            Next i
            r = 16
            For i = 0 To (Me.ListBox2.ListCount - 1)
                If Me.ListBox2.Selected(i) Then
                r = r + 1
                    wsParam.Cells(r, 6).Value = Me.ListBox2.List(i)
                End If
            Next i
            r = 16
            For i = 0 To (Me.ListBox3.ListCount - 1)
                If Me.ListBox3.Selected(i) Then
                r = r + 1
                    wsParam.Cells(r, 7).Value = Me.ListBox3.List(i)
                End If
            Next i
    
            Unload Me
        End If
    
End Sub

Private Sub UserForm_Initialize()
' form setup

    Dim wsParts As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim dict As Object
    
    ' formatting
    Label1.Font.Size = 12
    Label2.Font.Size = 12
    Label3.Font.Size = 12
    ListBox1.Font.Size = 12
    ListBox2.Font.Size = 12
    ListBox3.Font.Size = 12

    Set wsParts = ThisWorkbook.Sheets("Transactions")
    Set dict = CreateObject("Scripting.Dictionary")

    ' fill the list boxes with each check type, fleet, and location found in the transactions sheet
    lastRow = wsParts.Cells(wsParts.Rows.Count, 1).End(xlUp).Row
    For Each cell In wsParts.Range("E2:E" & lastRow)
        If Not dict.exists(cell.Value) And cell.Value <> "" Then
            dict.Add cell.Value, Nothing
        End If
    Next cell
    Me.ListBox1.List = dict.keys
    Set dict = CreateObject("Scripting.Dictionary")
    For Each cell In wsParts.Range("G2:G" & lastRow)
        If Not dict.exists(cell.Value) And cell.Value <> "" Then
            dict.Add cell.Value, Nothing
        End If
    Next cell
    Me.ListBox2.List = dict.keys
    Set dict = CreateObject("Scripting.Dictionary")
    For Each cell In wsParts.Range("H2:H" & lastRow)
        If Not dict.exists(cell.Value) And cell.Value <> "" Then
            dict.Add cell.Value, Nothing
        End If
    Next cell
    Me.ListBox3.List = dict.keys

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' when X (top right) is clicked:

    ' close the form and stop the BOM macro
    If CloseMode = vbFormControlMenu Then
        End
    End If
    
End Sub
