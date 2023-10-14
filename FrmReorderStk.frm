VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMReorderStk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Less Than Re-Order Qty"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12750
   ClipControls    =   0   'False
   Icon            =   "FrmReorderStk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12750
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   30
      TabIndex        =   5
      Top             =   8100
      Width           =   1200
   End
   Begin MSDataListLib.DataCombo CMBSUPPLIERexp 
      Height          =   315
      Left            =   3105
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   255
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "&Display"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10155
      TabIndex        =   3
      Top             =   8100
      Width           =   1200
   End
   Begin VB.TextBox TXTsample 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   290
      Left            =   1740
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   11490
      TabIndex        =   1
      Top             =   8100
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   8025
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   14155
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   400
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   3
      Appearance      =   0
      GridLineWidth   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FRMReorderStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CLOSEALL As Integer
Dim MFG_REC As New ADODB.Recordset

Private Sub CMDDISPLAY_Click()
    Dim rststock As ADODB.Recordset
    Dim i As Long
    
    'PHY_FLAG = True
    
    i = 0
    GRDSTOCK.Rows = 1
    GRDSTOCK.FixedRows = 0
    On Error GoTo Errhand
    Set rststock = New ADODB.Recordset
    'RSTSTKLESS.Open "SELECT DISTINCT RTRXFILE.ITEM_CODE, RTRXFILE.ITEM_NAME, RTRXFILE.UNIT, ITEMMAST.ITEM_CODE, ITEMMAST.REORDER_QTY, ITEMMAST.CLOSE_QTY, ITEMMAST.ITEM_NAME FROM RTRXFILE RIGHT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE  ITEMMAST.REORDER_QTY > ITEMMAST.CLOSE_QTY ORDER BY ITEMMAST.ITEM_NAME", db, adOpenForwardOnly
    rststock.Open "SELECT * FROM ITEMMAST WHERE CLOSE_QTY = 0 OR CLOSE_QTY < REORDER_QTY ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
    If rststock.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        MDIMAIN.vbalProgressBar1.Visible = True
        MDIMAIN.vbalProgressBar1.value = 0
        MDIMAIN.vbalProgressBar1.ShowText = True
        MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
        MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
    End If
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!REORDER_QTY), "", rststock!REORDER_QTY)
        GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!CLOSE_QTY), "", rststock!CLOSE_QTY)
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!Remarks), "", rststock!Remarks)
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!Category), "", rststock!Category)
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
        
        rststock.MoveNext
        MDIMAIN.vbalProgressBar1.value = MDIMAIN.vbalProgressBar1.value + 1
    Loop
    rststock.Close
    Set rststock = Nothing
    
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdExport_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Zero Stock Items"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, N As Long
    
    On Error GoTo Errhand
    Screen.MousePointer = vbHourglass
    'Create an Excel instalce.
    Set oApp = CreateObject("Excel.Application")
    Set oWB = oApp.Workbooks.Add
    Set oWS = oWB.Worksheets(1)
    

    
    
'    xlRange = oWS.Range("A1", "C1")
'    xlRange.Font.Bold = True
'    xlRange.ColumnWidth = 15
'    'xlRange.Value = {"First Name", "Last Name", "Last Service"}
'    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
'    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
'
'    xlRange = oWS.Range("C1", "C999")
'    xlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'    xlRange.ColumnWidth = 12
    
    'If Sum_flag = False Then
        oWS.Range("A1", "E1").Merge
        oWS.Range("A1", "E1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "E2").Merge
        oWS.Range("A2", "E2").HorizontalAlignment = xlCenter
    'End If
    oWS.Range("A:A").ColumnWidth = 6
    oWS.Range("B:B").ColumnWidth = 10
    oWS.Range("C:C").ColumnWidth = 12
    oWS.Range("D:D").ColumnWidth = 12
    oWS.Range("E:E").ColumnWidth = 12
    oWS.Range("F:F").ColumnWidth = 12
    oWS.Range("G:G").ColumnWidth = 12
    oWS.Range("H:H").ColumnWidth = 12
    
    
    
    oWS.Range("A1").Select                      '-- particular cell selection
    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 14            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    'oApp.Columns("A:A").EntireColumn.AutoFit     '-- autofitted column

    oWS.Range("A2").Select                      '-- particular cell selection
    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 11            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True

    oWS.Range("A" & 1).value = MDIMAIN.StatusBar.Panels(5).Text
    oWS.Range("A" & 2).value = "Stock Less Than Re-Order Qty"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).value = GRDSTOCK.TextMatrix(0, 0)
    oWS.Range("B" & 3).value = GRDSTOCK.TextMatrix(0, 1)
    oWS.Range("C" & 3).value = GRDSTOCK.TextMatrix(0, 2)
    oWS.Range("D" & 3).value = GRDSTOCK.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 3).value = GRDSTOCK.TextMatrix(0, 4)
    oWS.Range("F" & 3).value = GRDSTOCK.TextMatrix(0, 5)
    oWS.Range("G" & 3).value = GRDSTOCK.TextMatrix(0, 6)
    oWS.Range("H" & 3).value = GRDSTOCK.TextMatrix(0, 7)
    
    On Error GoTo Errhand
    
    i = 4
    For N = 1 To GRDSTOCK.Rows - 1
        oWS.Range("A" & i).value = GRDSTOCK.TextMatrix(N, 0)
        oWS.Range("B" & i).value = GRDSTOCK.TextMatrix(N, 1)
        oWS.Range("C" & i).value = GRDSTOCK.TextMatrix(N, 2)
        oWS.Range("D" & i).value = GRDSTOCK.TextMatrix(N, 3)
        oWS.Range("E" & i).value = GRDSTOCK.TextMatrix(N, 4)
        oWS.Range("F" & i).value = GRDSTOCK.TextMatrix(N, 5)
        oWS.Range("G" & i).value = GRDSTOCK.TextMatrix(N, 6)
        oWS.Range("H" & i).value = GRDSTOCK.TextMatrix(N, 7)
        On Error GoTo Errhand
        i = i + 1
    Next N
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.HorizontalAlignment = xlRight
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 13            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    
   
SKIP:
    oApp.Visible = True
    
    
'    Set oWB = Nothing
'    oApp.Quit
'    Set oApp = Nothing
'
    
    Screen.MousePointer = vbNormal
    Exit Sub
Errhand:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox Err.Description
End Sub

Private Sub Form_Load()

    Set CMBSUPPLIERexp.DataSource = Nothing
    'MFG_REC.Open "SELECT DISTINCT CATEGORY FROM ITEMMAST RIGHT JOIN RTRXFILE ON ITEMMAST.ITEM_CODE = RTRXFILE.ITEM_CODE WHERE RTRXFILE.BAL_QTY > 0 ORDER BY [ITEMMAST.MANUFACTURER]", db, adOpenForwardOnly
    MFG_REC.Open "SELECT DISTINCT CATEGORY FROM CATEGORY ORDER BY CATEGORY", db, adOpenForwardOnly
    Set CMBSUPPLIERexp.RowSource = MFG_REC
    CMBSUPPLIERexp.ListField = "CATEGORY"
    
    CLOSEALL = 1
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "MIN STOCK"
    GRDSTOCK.TextMatrix(0, 4) = "STOCK"
    GRDSTOCK.TextMatrix(0, 5) = "PACK"
    GRDSTOCK.TextMatrix(0, 6) = "CATEGORY"
    GRDSTOCK.TextMatrix(0, 7) = "BIN LOCATION"
    
    GRDSTOCK.ColWidth(0) = 900
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 3200
    GRDSTOCK.ColWidth(3) = 1100
    GRDSTOCK.ColWidth(4) = 1000
    GRDSTOCK.ColWidth(5) = 0
    GRDSTOCK.ColWidth(6) = 2600
    GRDSTOCK.ColWidth(7) = 2500
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 3
    GRDSTOCK.ColAlignment(4) = 3
    GRDSTOCK.ColAlignment(5) = 3
    GRDSTOCK.ColAlignment(6) = 1
    GRDSTOCK.ColAlignment(7) = 1
    
    Left = 500
    Top = 0
    Height = 10000
    Width = 12840
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        MFG_REC.Close
        'If REPFLAG = False Then RSTREP.Close
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 555
        MDIMAIN.PCTMENU.SetFocus
    End If
   Cancel = CLOSEALL
End Sub

Private Sub GRDSTOCK_Click()
    TXTsample.Visible = False
    CMBSUPPLIERexp.Visible = False
    GRDSTOCK.SetFocus
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 3
                    TXTsample.MaxLength = 7
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 200
                    TXTsample.Left = GRDSTOCK.CellLeft + 50
                    TXTsample.Width = GRDSTOCK.CellWidth - 25
                    TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
                Case 7
                    TXTsample.MaxLength = 15
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 200
                    TXTsample.Left = GRDSTOCK.CellLeft + 50
                    TXTsample.Width = GRDSTOCK.CellWidth - 25
                    TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
                Case 6
                    CMBSUPPLIERexp.Visible = True
                    CMBSUPPLIERexp.Top = GRDSTOCK.CellTop + 200
                    CMBSUPPLIERexp.Left = GRDSTOCK.CellLeft + 50
                    CMBSUPPLIERexp.Width = GRDSTOCK.CellWidth
                    CMBSUPPLIERexp.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    CMBSUPPLIERexp.SetFocus
            End Select
        Case 114
            sitem = UCase(InputBox("Item Na..?", "STOCK"))
            For i = 1 To GRDSTOCK.Rows - 1
                    If Mid(GRDSTOCK.TextMatrix(i, 2), 1, Len(sitem)) = sitem Then
                        GRDSTOCK.Row = i
                        GRDSTOCK.TopRow = i
                    Exit For
                End If
            Next i
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub GRDSTOCK_Scroll()
    TXTsample.Visible = False
    GRDSTOCK.SetFocus
End Sub


Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim M_STOCK As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 3   'Min Qty
                    'If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT REORDER_QTY, ITEM_CODE FROM ITEMMAST WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!REORDER_QTY = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 7   'Location
                    'If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT BIN_LOCATION, ITEM_CODE FROM ITEMMAST WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!BIN_LOCATION = Trim(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
            End Select
            
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 3
            Select Case KeyAscii
                Case Asc("'")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 6
            Select Case KeyAscii
                Case Asc("'")
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub CMBSUPPLIERexp_Click(Area As Integer)
    CMBSUPPLIERexp.SelStart = 0
    CMBSUPPLIERexp.SelLength = Len(CMBSUPPLIERexp.Text)
End Sub

Private Sub CMBSUPPLIERexp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTSUPPLIER As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(CMBSUPPLIERexp.Text) = "" Then Exit Sub
            If CMBSUPPLIERexp.MatchedWithList = False Then
                MsgBox "Select Category from the list", vbOKOnly, "Min Stock!!!"
                CMBSUPPLIERexp.SelStart = 0
                CMBSUPPLIERexp.SelLength = Len(CMBSUPPLIERexp.Text)
                CMBSUPPLIERexp.SetFocus
                Exit Sub
            End If
            Set RSTSUPPLIER = New ADODB.Recordset
            RSTSUPPLIER.Open "SELECT CATEGORY, ITEM_CODE FROM ITEMMAST WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                RSTSUPPLIER!Category = Trim(CMBSUPPLIERexp.Text)
                RSTSUPPLIER.Update
            End If
            RSTSUPPLIER.Close
            Set RSTSUPPLIER = Nothing
            
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(CMBSUPPLIERexp.Text)
            GRDSTOCK.Enabled = True
            CMBSUPPLIERexp.Visible = False
            GRDSTOCK.SetFocus
        Case vbKeyEscape
            CMBSUPPLIERexp.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub CMBSUPPLIERexp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

