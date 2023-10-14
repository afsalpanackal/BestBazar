VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmBinLoc 
   BackColor       =   &H00E8DFEC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price List"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmBinLoc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   13680
   Begin VB.CommandButton CmdExport 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8700
      TabIndex        =   23
      Top             =   1020
      Width           =   1530
   End
   Begin VB.CheckBox CHKCATEGORY2 
      BackColor       =   &H00E8DFEC&
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   10590
      TabIndex        =   21
      Top             =   285
      Width           =   1635
   End
   Begin VB.TextBox TXTDEALER2 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   10590
      TabIndex        =   20
      Top             =   630
      Width           =   3075
   End
   Begin VB.CheckBox chkcategory 
      BackColor       =   &H00E8DFEC&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   12285
      TabIndex        =   19
      Top             =   300
      Width           =   1410
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   19320
      Top             =   1935
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Re- Load"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8700
      TabIndex        =   13
      Top             =   600
      Width           =   1530
   End
   Begin VB.TextBox TxtCode 
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
      Height          =   345
      Left            =   4290
      TabIndex        =   1
      Top             =   270
      Width           =   2040
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1230
      Left            =   6360
      TabIndex        =   4
      Top             =   645
      Width           =   2310
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Display All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   75
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptStock 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Stock Items Only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Top             =   645
         Width           =   1935
      End
   End
   Begin VB.TextBox tXTMEDICINE 
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
      Height          =   345
      Left            =   45
      TabIndex        =   0
      Top             =   270
      Width           =   4215
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
      Height          =   405
      Left            =   8700
      TabIndex        =   3
      Top             =   1455
      Width           =   1545
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1230
      Left            =   45
      TabIndex        =   2
      Top             =   630
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   2170
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
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
   Begin VB.Frame Frame1 
      Height          =   6645
      Left            =   45
      TabIndex        =   7
      Top             =   1800
      Width           =   13635
      Begin MSDataListLib.DataCombo CMBMFGR 
         Height          =   360
         Left            =   6120
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TXTsample 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   210
         TabIndex        =   9
         Top             =   870
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
         Height          =   6480
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   13560
         _ExtentX        =   23918
         _ExtentY        =   11430
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   8438015
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DTFROM 
      Height          =   420
      Left            =   7905
      TabIndex        =   15
      Top             =   90
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   0
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   255
      Format          =   119996417
      CurrentDate     =   40498
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   780
      Left            =   10590
      TabIndex        =   22
      Top             =   975
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   1376
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   0
      TabIndex        =   18
      Top             =   810
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   450
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Stock Entry Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Index           =   3
      Left            =   6480
      TabIndex        =   16
      Top             =   60
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 - EDIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4590
      TabIndex        =   12
      Top             =   -15
      Width           =   1920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   75
      TabIndex        =   11
      Top             =   30
      Width           =   4380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Part"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   60
      TabIndex        =   10
      Top             =   660
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "FrmBinLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPFLAG As Boolean 'REP
Dim MFG_REC As New ADODB.Recordset
Dim RSTREP As New ADODB.Recordset
Dim PHY_FLAG As Boolean 'REP
Dim PHY_REC As New ADODB.Recordset

Private Sub CHKCATEGORY_Click()
    CHKCATEGORY2.value = 0
End Sub

Private Sub CHKCATEGORY2_Click()
    chkcategory.value = 0
End Sub

Private Sub CMBMFGR_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 11  'pack
                    If CMBMFGR.Text = "" Then
                        MsgBox "Please select Company from the List", vbOKOnly, "Stock Correction"
                        Exit Sub
                    End If
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MANUFACTURER = CMBMFGR.Text
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = CMBMFGR.Text
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    CMBMFGR.Visible = False
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            CMBMFGR.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdLoad_Click()
    Call Fillgrid
End Sub

Private Sub CmdExport_Click()
    If frmLogin.rs!Level <> "0" Then MsgBox "Permission Denied", vbOKOnly, "Price Analysis"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, N As Long
    
    On Error GoTo eRRhAND
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
    oWS.Range("C:C").ColumnWidth = 20
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

'    Range("C2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("C:C").EntireColumn.AutoFit     '-- autofitted column
'
'
'    Range("D2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("D:D").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("E2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("E:E").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("F2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("F:F").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("G2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("G:G").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("H2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("H:H").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("I2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("I:I").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("J2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("J:J").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("K2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("K:K").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("L2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("L:L").EntireColumn.AutoFit     '-- autofitted column

'    oWB.ActiveSheet.Font.Name = "Arial"
'    oApp.ActiveSheet.Font.Name = "Arial"
'    oWB.Font.Size = "11"
'    oWB.Font.Bold = True
    oWS.Range("A" & 1).value = MDIMAIN.StatusBar.Panels(5).Text
    oWS.Range("A" & 2).value = "PRICE LIST"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).value = GRDSTOCK.TextMatrix(0, 0)
    oWS.Range("B" & 3).value = GRDSTOCK.TextMatrix(0, 1)
    oWS.Range("C" & 3).value = GRDSTOCK.TextMatrix(0, 2)
    oWS.Range("D" & 3).value = GRDSTOCK.TextMatrix(0, 5)
    oWS.Range("E" & 3).value = GRDSTOCK.TextMatrix(0, 6)
    
    On Error GoTo eRRhAND
    
    i = 4
    For N = 1 To GRDSTOCK.Rows - 1
        oWS.Range("A" & i).value = GRDSTOCK.TextMatrix(N, 0)
        oWS.Range("B" & i).value = GRDSTOCK.TextMatrix(N, 1)
        oWS.Range("C" & i).value = GRDSTOCK.TextMatrix(N, 2)
        oWS.Range("D" & i).value = GRDSTOCK.TextMatrix(N, 5)
        oWS.Range("E" & i).value = GRDSTOCK.TextMatrix(N, 6)
        On Error GoTo eRRhAND
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
eRRhAND:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox Err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            GRDSTOCK.SetFocus
            'DataList2.SetFocus
        Case vbKeyEscape
            tXTMEDICINE.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo eRRhAND
    Set CMBMFGR.DataSource = Nothing
    MFG_REC.Open "SELECT DISTINCT MANUFACTURER FROM ITEMMAST ORDER BY ITEMMAST.MANUFACTURER", db, adOpenForwardOnly
    Set CMBMFGR.RowSource = MFG_REC
    CMBMFGR.ListField = "MANUFACTURER"
    
    REPFLAG = True
    PHY_FLAG = True
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "" '"LOC"
    GRDSTOCK.TextMatrix(0, 4) = "QTY"
    GRDSTOCK.TextMatrix(0, 5) = "PRICE"
    GRDSTOCK.TextMatrix(0, 6) = "PRICE(AC)"
    GRDSTOCK.TextMatrix(0, 7) = "COST"
    GRDSTOCK.TextMatrix(0, 8) = "" '"SPECIFICATIONS"
        
    GRDSTOCK.ColWidth(0) = 400
    GRDSTOCK.ColWidth(1) = 1500
    GRDSTOCK.ColWidth(2) = 5000
    GRDSTOCK.ColWidth(3) = 0 '2000
    GRDSTOCK.ColWidth(4) = 1100
    GRDSTOCK.ColWidth(5) = 1100
    GRDSTOCK.ColWidth(6) = 1100
    GRDSTOCK.ColWidth(7) = 1100
    GRDSTOCK.ColWidth(8) = 0
    GRDSTOCK.ColWidth(9) = 0

    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 4
    GRDSTOCK.ColAlignment(5) = 4
    GRDSTOCK.ColAlignment(6) = 4
    GRDSTOCK.ColAlignment(7) = 4
    GRDSTOCK.ColAlignment(8) = 1
    GRDSTOCK.ColAlignment(9) = 1
    
    DTFROM.value = Format(Date, "DD/MM/YYYY")
    Call Fillgrid
    'Me.Height = 8415
    'Me.Width = 6465
    Me.Left = 0
    Me.Top = 0
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    If PHY_FLAG = False Then PHY_REC.Close
    MFG_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            If frmLogin.rs!Level = "0" Then
                Select Case GRDSTOCK.Col
                    Case 1, 3, 2, 5, 6, 8
                        TXTsample.Visible = True
                        TXTsample.Top = GRDSTOCK.CellTop + 100
                        TXTsample.Left = GRDSTOCK.CellLeft '+ 60
                        TXTsample.Width = GRDSTOCK.CellWidth
                        TXTsample.Height = GRDSTOCK.CellHeight
                        TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        TXTsample.SetFocus
                    Case 12
                        CMBMFGR.Visible = True
                        CMBMFGR.Top = GRDSTOCK.CellTop + 100
                        CMBMFGR.Left = GRDSTOCK.CellLeft '+ 60
                        CMBMFGR.Width = GRDSTOCK.CellWidth
                        'CmbPack.Height = GRDSTOCK.CellHeight
                        On Error Resume Next
                        CMBMFGR.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        CMBMFGR.SetFocus
                End Select
            End If
        Case 114
            sitem = UCase(InputBox("Item Name...?", "STOCK"))
            For i = 1 To GRDSTOCK.Rows - 1
                If UCase(Mid(GRDSTOCK.TextMatrix(i, 2), 1, Len(sitem))) = sitem Then
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
    CMBMFGR.Visible = False
End Sub

Private Sub OptAll_Click()
    tXTMEDICINE.SetFocus
End Sub

Private Sub OptStock_Click()
    tXTMEDICINE.SetFocus
End Sub

Private Sub tXTMEDICINE_Change()
    On Error GoTo eRRhAND
    Call Fillgrid
    If REPFLAG = True Then
        If CHKCATEGORY2.value = 0 Then
            If OptStock.value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If CHKCATEGORY2.value = 0 Then
            If OptStock.value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.Text)
    'Call Fillgrid
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            TxtCode.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub DataList2_Click()
    Exit Sub
    Dim rststock As ADODB.Recordset
    Dim i As Long

    On Error GoTo eRRhAND
    
    i = 0
    Screen.MousePointer = vbHourglass
    
    GRDSTOCK.Rows = 1
    Set rststock = New ADODB.Recordset
    'WHERE ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%'
    'rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & DataList2.BoundText & "'ORDER BY VCH_NO DESC", db, adOpenStatic, adLockReadOnly
    rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly
    If Not (rststock.EOF And rststock.BOF) Then
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = rststock!CLOSE_QTY
        GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 1), "0.00"))
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 1), "0.00"))
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!CRTN_PACK), "", rststock!CRTN_PACK)
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!PACK_TYPE), "", rststock!PACK_TYPE)
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!P_CRTN), "", Format(Round(rststock!P_CRTN, 1), "0.00"))
        GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!P_LWS), "", Format(Round(rststock!P_LWS, 1), "0.00"))
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!P_VAN), "", Format(rststock!P_VAN, "0.00"))
        GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!Category), "", rststock!Category)
        'GRDSTOCK.TextMatrix(i, 12) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
        GRDSTOCK.TextMatrix(i, 12) = IIf(IsNull(rststock!MANUFACTURER), "", rststock!MANUFACTURER)
        GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST, "0.00"))
        GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.00"))
        GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!SALES_TAX), "", Format(rststock!SALES_TAX, "0.00"))
        GRDSTOCK.TextMatrix(i, 17) = IIf(IsNull(rststock!LOOSE_PACK), "", rststock!LOOSE_PACK)
        If Val(GRDSTOCK.TextMatrix(i, 17)) = 0 Then GRDSTOCK.TextMatrix(i, 17) = 1
        If Val(GRDSTOCK.TextMatrix(i, 13)) <> 0 Then
            GRDSTOCK.TextMatrix(i, 16) = Round(((Val(GRDSTOCK.TextMatrix(i, 4)) - Val(GRDSTOCK.TextMatrix(i, 12))) * 100) / Val(GRDSTOCK.TextMatrix(i, 12)), 2)
        Else
            GRDSTOCK.TextMatrix(i, 16) = 0
        End If
        GRDSTOCK.TextMatrix(i, 18) = IIf(IsNull(rststock!CUST_DISC), "", Format(rststock!CUST_DISC, "0.00"))
        Select Case rststock!COM_FLAG
            Case "P"
                GRDSTOCK.TextMatrix(i, 19) = IIf(IsNull(rststock!COM_PER), "", Format(rststock!COM_PER, "0.00"))
                GRDSTOCK.TextMatrix(i, 20) = "%"
            Case "A"
                GRDSTOCK.TextMatrix(i, 19) = IIf(IsNull(rststock!COM_AMT), "", Format(rststock!COM_AMT, "0.00"))
                GRDSTOCK.TextMatrix(i, 20) = "Rs"
        End Select
        rststock.MoveNext
    End If
    rststock.Close
    Set rststock = Nothing
    Screen.MousePointer = vbNormal
    Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
    
End Sub

Private Function Fillgrid()
    Dim rststock As ADODB.Recordset
    Dim rstopstock As ADODB.Recordset
    Dim i As Long

    On Error GoTo eRRhAND
    
    i = 0
    Screen.MousePointer = vbHourglass
    
    GRDSTOCK.Rows = 1
    Set rststock = New ADODB.Recordset
    If CHKCATEGORY2.value = 0 And chkcategory.value = 0 Then
        If OptStock.value = True Then
            'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
    Else
        If CHKCATEGORY2.value = 1 Then
            If OptStock.value = True Then
                'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.value = True Then
                'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If

        End If
    End If
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
        GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!CLOSE_QTY), "", rststock!CLOSE_QTY)
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 1), "0.00"))
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 1), "0.00"))
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST, "0.00"))
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!ITEM_SPEC), "", rststock!ITEM_SPEC)
        GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!Category), "", rststock!Category)
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    Screen.MousePointer = vbNormal
    Exit Function

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = vbCtrlMask Then Call Clipboard_fn(KeyCode, Shift, TXTsample)
    Dim rststock, RSTITEMMAST As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 1  ' Item Code
                    If Trim(TXTsample.Text) = "" Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_CODE = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!ITEM_CODE = Trim(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from TRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!ITEM_CODE = Trim(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 2  ' Item Name
                    If Trim(TXTsample.Text) = "" Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_NAME = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!ITEM_NAME = Trim(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 3  'LOC
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!BIN_LOCATION = Trim(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 5  'PRICE
                    db.Execute "Update ITEMMAST set P_RETAIL = " & Val(TXTsample.Text) & " where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' "
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                    db.Execute "Update RTRXFILE set P_RETAIL = " & Val(TXTsample.Text) & " where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0"
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 6  'PRICE AC
                    db.Execute "Update ITEMMAST set P_WS = " & Val(TXTsample.Text) & " where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' "
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                    db.Execute "Update RTRXFILE set P_WS = " & Val(TXTsample.Text) & " where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0"
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 8  'SPEC
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_SPEC = Trim(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 4, 5, 6, 8
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 1, 11, 12, 2
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub TXTCODE_Change()
    On Error GoTo eRRhAND
    Call Fillgrid
    If REPFLAG = True Then
        If OptStock.value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If OptStock.value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.Text)
    'Call Fillgrid
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            tXTMEDICINE.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select

End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTDEALER2_Change()
    
    On Error GoTo eRRhAND
    If FLAGCHANGE2.Caption <> "1" Then
        If chkcategory.value = 1 Then
            If PHY_FLAG = True Then
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            End If
            If (PHY_REC.EOF And PHY_REC.BOF) Then
                LBLDEALER2.Caption = ""
            Else
                LBLDEALER2.Caption = PHY_REC!Category
            End If
            Set Me.DataList1.RowSource = PHY_REC
            DataList1.ListField = "CATEGORY"
            DataList1.BoundColumn = "CATEGORY"
        Else
            If PHY_FLAG = True Then
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            End If
            If (PHY_REC.EOF And PHY_REC.BOF) Then
                LBLDEALER2.Caption = ""
            Else
                LBLDEALER2.Caption = PHY_REC!MANUFACTURER
            End If
            Set Me.DataList1.RowSource = PHY_REC
            DataList1.ListField = "MANUFACTURER"
            DataList1.BoundColumn = "MANUFACTURER"

        End If
    End If
    Exit Sub
eRRhAND:
    MsgBox Err.Description
    
End Sub


Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.Text)
    'CHKCATEGORY2.value = 1
End Sub

Private Sub TXTDEALER2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList1.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList1.SetFocus
    End Select

End Sub

Private Sub TXTDEALER2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_Click()
        
    TXTDEALER2.Text = DataList1.Text
    LBLDEALER2.Caption = TXTDEALER2.Text
    Call Fillgrid
    tXTMEDICINE.SetFocus
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER2.Text) = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Category From List", vbOKOnly, "Category List..."
                DataList1.SetFocus
                Exit Sub
            End If
        Case vbKeyEscape
            TXTDEALER2.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_GotFocus()
    FLAGCHANGE2.Caption = 1
    TXTDEALER2.Text = LBLDEALER2.Caption
    DataList1.Text = TXTDEALER2.Text
    Call DataList1_Click
    'CHKCATEGORY2.value = 1
End Sub

Private Sub DataList1_LostFocus()
     FLAGCHANGE2.Caption = ""
End Sub

