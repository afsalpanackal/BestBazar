VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmZeroStk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zero Stock Items"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11580
   Icon            =   "Frmeasybill.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   11580
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
      TabIndex        =   3
      Top             =   7920
      Width           =   1200
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
      Left            =   10335
      TabIndex        =   1
      Top             =   7905
      Width           =   1200
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
      Left            =   9000
      TabIndex        =   0
      Top             =   7905
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   7860
      Left            =   0
      TabIndex        =   2
      Top             =   15
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13864
      _Version        =   393216
      Rows            =   1
      Cols            =   5
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
Attribute VB_Name = "FrmZeroStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MFG_REC As New ADODB.Recordset
Dim SCH_REC As New ADODB.Recordset
Dim MOLEFLAG As Boolean
Dim RSTMOLE As New ADODB.Recordset

Private Sub CmDDisplay_Click()
    Dim rststock, RSTRTRXFILE, RSTSUPPLIER As ADODB.Recordset
    Dim I As Long
    
    'PHY_FLAG = True
    
    
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.rows = 1
    I = 0
    
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHand
    Set rststock = New ADODB.Recordset
    'RSTSTKLESS.Open "SELECT DISTINCT RTRXFILE.ITEM_CODE, RTRXFILE.ITEM_NAME, RTRXFILE.UNIT, ITEMMAST.ITEM_CODE, ITEMMAST.REORDER_QTY, ITEMMAST.CLOSE_QTY, ITEMMAST.ITEM_NAME FROM RTRXFILE RIGHT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE  ITEMMAST.REORDER_QTY > ITEMMAST.CLOSE_QTY ORDER BY ITEMMAST.ITEM_NAME", db, adOpenForwardOnly
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT DISTINCT ITEM_CODE, ITEM_NAME FROM RTRXFILE ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
'    If rststock.RecordCount > 0 Then
'        Screen.MousePointer = vbHourglass
'        MDIMAIN.vbalProgressBar1.Visible = True
'        MDIMAIN.vbalProgressBar1.value = 0
'        MDIMAIN.vbalProgressBar1.ShowText = True
'        MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
'        MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
'    End If
'    Do Until rststock.EOF
        Screen.MousePointer = vbHourglass
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * FROM ITEMMAST WHERE CLOSE_QTY <= 0 ", db, adOpenStatic, adLockReadOnly
                
        MDIMAIN.vbalProgressBar1.Visible = True
        MDIMAIN.vbalProgressBar1.Value = 0
        MDIMAIN.vbalProgressBar1.ShowText = True
        MDIMAIN.vbalProgressBar1.text = "PLEASE WAIT..."
        If RSTRTRXFILE.RecordCount > 0 Then MDIMAIN.vbalProgressBar1.Max = RSTRTRXFILE.RecordCount
        Do Until RSTRTRXFILE.EOF
            I = I + 1
            GRDSTOCK.rows = GRDSTOCK.rows + 1
            GRDSTOCK.FixedRows = 1
            GRDSTOCK.TextMatrix(I, 0) = I
            GRDSTOCK.TextMatrix(I, 1) = RSTRTRXFILE!ITEM_CODE
            GRDSTOCK.TextMatrix(I, 2) = RSTRTRXFILE!ITEM_NAME
            Set RSTSUPPLIER = New ADODB.Recordset
            RSTSUPPLIER.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PI' AND ITEM_CODE = '" & RSTRTRXFILE!ITEM_CODE & "' ORDER BY  TRX_YEAR DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
            If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                Select Case RSTSUPPLIER!TRX_TYPE
                    Case "PI"
                        GRDSTOCK.TextMatrix(I, 3) = IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", Mid(RSTSUPPLIER!VCH_DESC, 15))
                End Select
            End If
            RSTSUPPLIER.Close
            Set RSTSUPPLIER = Nothing
            GRDSTOCK.TextMatrix(I, 4) = IIf(IsNull(RSTRTRXFILE!SCHEDULE), "", RSTRTRXFILE!SCHEDULE)
            
            RSTRTRXFILE.MoveNext
        Loop
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing
        
        MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
'        rststock.MoveNext
'    Loop
'    rststock.Close
'    Set rststock = Nothing
        
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdExport_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Zero Stock Items"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim I, n As Long
    
    On Error GoTo ErrHand
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

    oWS.Range("A" & 1).Value = MDIMAIN.StatusBar.Panels(5).text
    oWS.Range("A" & 2).Value = "ZERO STOCK REPORT"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).Value = GRDSTOCK.TextMatrix(0, 0)
    oWS.Range("B" & 3).Value = GRDSTOCK.TextMatrix(0, 1)
    oWS.Range("C" & 3).Value = GRDSTOCK.TextMatrix(0, 2)
    oWS.Range("D" & 3).Value = GRDSTOCK.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 3).Value = GRDSTOCK.TextMatrix(0, 4)
    
    On Error GoTo ErrHand
    
    I = 4
    For n = 1 To GRDSTOCK.rows - 1
        oWS.Range("A" & I).Value = GRDSTOCK.TextMatrix(n, 0)
        oWS.Range("B" & I).Value = GRDSTOCK.TextMatrix(n, 1)
        oWS.Range("C" & I).Value = GRDSTOCK.TextMatrix(n, 2)
        oWS.Range("D" & I).Value = GRDSTOCK.TextMatrix(n, 3)
        oWS.Range("E" & I).Value = GRDSTOCK.TextMatrix(n, 4)
        On Error GoTo ErrHand
        I = I + 1
    Next n
    oWS.Range("A" & I, "Z" & I).Select                      '-- particular cell selection
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
ErrHand:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    
    MOLEFLAG = True
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "SUPPLIER"
    GRDSTOCK.TextMatrix(0, 4) = "SCHEDULE"
    
    GRDSTOCK.ColWidth(0) = 900
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 4200
    GRDSTOCK.ColWidth(3) = 5000
    GRDSTOCK.ColWidth(4) = 1000
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 1
    GRDSTOCK.ColAlignment(4) = 1
    
    
    Left = 500
    Top = 0
    'Height = 10000
    'Width = 12840
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If REPFLAG = False Then RSTREP.Close
    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim I As Long
    If GRDSTOCK.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 114
            sitem = UCase(InputBox("Item Name..?", "ZERO STOCK"))
            For I = 1 To GRDSTOCK.rows - 1
                    If Mid(GRDSTOCK.TextMatrix(I, 2), 1, Len(sitem)) = sitem Then
                        GRDSTOCK.Row = I
                        GRDSTOCK.TopRow = I
                    Exit For
                End If
            Next I
            GRDSTOCK.SetFocus
    End Select
End Sub

