VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmBinLoc1 
   BackColor       =   &H00E8DFEC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price List"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmpricelistnew.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   19125
   Begin VB.CommandButton Command1 
      Caption         =   "Print Price List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10605
      TabIndex        =   25
      Top             =   1800
      Width           =   1320
   End
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8700
      TabIndex        =   23
      Top             =   1695
      Width           =   1830
   End
   Begin VB.CheckBox CHKCATEGORY2 
      BackColor       =   &H00E8DFEC&
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Arial"
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
         Name            =   "Arial"
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
      Height          =   510
      Left            =   8700
      TabIndex        =   13
      Top             =   645
      Width           =   1830
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
      Height          =   1200
      Left            =   6360
      TabIndex        =   4
      Top             =   660
      Width           =   2310
      Begin VB.OptionButton OptPLU 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Items with P&LU codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   60
         TabIndex        =   24
         Top             =   525
         Width           =   2160
      End
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
         Top             =   180
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
         Top             =   810
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
      Height          =   480
      Left            =   8700
      TabIndex        =   3
      Top             =   1185
      Width           =   1830
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   45
      TabIndex        =   2
      Top             =   645
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   2858
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
      Height          =   6210
      Left            =   45
      TabIndex        =   7
      Top             =   2250
      Width           =   19080
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
         Height          =   6060
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   19020
         _ExtentX        =   33549
         _ExtentY        =   10689
         _Version        =   393216
         Rows            =   1
         Cols            =   17
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
            Size            =   8.25
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
      Top             =   150
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Format          =   112984065
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
      Top             =   120
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
Attribute VB_Name = "FrmBinLoc1"
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
    CHKCATEGORY2.Value = 0
End Sub

Private Sub CHKCATEGORY2_Click()
    chkcategory.Value = 0
End Sub

Private Sub CMBMFGR_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 11  'pack
                    If CMBMFGR.text = "" Then
                        MsgBox "Please select Company from the List", vbOKOnly, "Stock Correction"
                        Exit Sub
                    End If
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MANUFACTURER = CMBMFGR.text
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = CMBMFGR.text
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
ErrHand:
    MsgBox err.Description
End Sub

Private Sub CMBMFGR_LostFocus()
    CMBMFGR.Visible = False
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdExport_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Price Analysis"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, n As Long
    
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
        oWS.Range("A1", "O1").Merge
        oWS.Range("A1", "O1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "O2").Merge
        oWS.Range("A2", "O2").HorizontalAlignment = xlCenter
    'End If
'    oWS.Range("A:A").ColumnWidth = 6
'    oWS.Range("B:B").ColumnWidth = 10
'    oWS.Range("C:C").ColumnWidth = 12
'    oWS.Range("D:D").ColumnWidth = 12
'    oWS.Range("E:E").ColumnWidth = 12
'    oWS.Range("F:F").ColumnWidth = 12
'    oWS.Range("G:G").ColumnWidth = 12
'    oWS.Range("H:H").ColumnWidth = 12
'    oWS.Range("I:I").ColumnWidth = 12
'    oWS.Range("J:J").ColumnWidth = 12
'    oWS.Range("K:K").ColumnWidth = 12
'    oWS.Range("L:L").ColumnWidth = 12
'    oWS.Range("M:M").ColumnWidth = 12
'    oWS.Range("N:N").ColumnWidth = 12
'    oWS.Range("O:O").ColumnWidth = 12
'    oWS.Range("P:P").ColumnWidth = 12
'    oWS.Range("Q:Q").ColumnWidth = 12
'    oWS.Range("R:R").ColumnWidth = 12
'    oWS.Range("S:S").ColumnWidth = 12
'    oWS.Range("T:T").ColumnWidth = 12
'    oWS.Range("U:U").ColumnWidth = 12
'    oWS.Range("V:V").ColumnWidth = 12
    
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
    oWS.Range("A" & 1).Value = MDIMAIN.StatusBar.Panels(5).text
    oWS.Range("A" & 2).Value = "PRICE LIST"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).Value = GRDSTOCK.TextMatrix(0, 0)
    oWS.Range("B" & 3).Value = GRDSTOCK.TextMatrix(0, 1)
    oWS.Range("C" & 3).Value = GRDSTOCK.TextMatrix(0, 2)
    oWS.Range("D" & 3).Value = GRDSTOCK.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 3).Value = GRDSTOCK.TextMatrix(0, 4)
    oWS.Range("F" & 3).Value = GRDSTOCK.TextMatrix(0, 5)
    oWS.Range("G" & 3).Value = GRDSTOCK.TextMatrix(0, 6)
    oWS.Range("H" & 3).Value = GRDSTOCK.TextMatrix(0, 7)
    oWS.Range("I" & 3).Value = GRDSTOCK.TextMatrix(0, 8)
    oWS.Range("J" & 3).Value = GRDSTOCK.TextMatrix(0, 9)
    oWS.Range("K" & 3).Value = GRDSTOCK.TextMatrix(0, 10)
    oWS.Range("L" & 3).Value = GRDSTOCK.TextMatrix(0, 11)
    oWS.Range("M" & 3).Value = GRDSTOCK.TextMatrix(0, 12)
    oWS.Range("N" & 3).Value = GRDSTOCK.TextMatrix(0, 13)
    oWS.Range("O" & 3).Value = GRDSTOCK.TextMatrix(0, 14)
    On Error GoTo ErrHand
    
    i = 4
    For n = 1 To GRDSTOCK.rows - 1
        oWS.Range("A" & i).Value = GRDSTOCK.TextMatrix(n, 0)
        oWS.Range("B" & i).Value = GRDSTOCK.TextMatrix(n, 1)
        oWS.Range("C" & i).Value = GRDSTOCK.TextMatrix(n, 2)
        oWS.Range("D" & i).Value = GRDSTOCK.TextMatrix(n, 3)
        oWS.Range("E" & i).Value = GRDSTOCK.TextMatrix(n, 4)
        oWS.Range("F" & i).Value = GRDSTOCK.TextMatrix(n, 5)
        oWS.Range("G" & i).Value = GRDSTOCK.TextMatrix(n, 6)
        oWS.Range("H" & i).Value = GRDSTOCK.TextMatrix(n, 7)
        oWS.Range("I" & i).Value = GRDSTOCK.TextMatrix(n, 8)
        oWS.Range("J" & i).Value = GRDSTOCK.TextMatrix(n, 9)
        oWS.Range("K" & i).Value = GRDSTOCK.TextMatrix(n, 10)
        oWS.Range("L" & i).Value = GRDSTOCK.TextMatrix(n, 11)
        oWS.Range("M" & i).Value = GRDSTOCK.TextMatrix(n, 12)
        oWS.Range("N" & i).Value = GRDSTOCK.TextMatrix(n, 13)
        oWS.Range("O" & i).Value = GRDSTOCK.TextMatrix(n, 14)
        i = i + 1
    Next n
    'oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    'oApp.Selection.HorizontalAlignment = xlRight
    'oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    'oApp.Selection.Font.Size = 13            '-- enabled bold cell style
    'oApp.Selection.Font.Bold = True
    oWS.Columns("A:Z").EntireColumn.AutoFit
   
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
    MsgBox err.Description
End Sub

Private Sub CmdLoad_Click()
    Call Fillgrid
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    
    On Error GoTo ErrHand
    ReportNameVar = Rptpath & "RPTPRICELIST"
    
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.RecordSelectionFormula = "( {ITEMMAST.CLOSE_QTY}<> 0 )"
    Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
    Set CRXFormulaFields = Report.FormulaFields
    
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        ''If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'A R STEELS' & chr(13) & 'Alappuzha'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Price List'"
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "STOCK REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
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
    On Error GoTo ErrHand
    Set CMBMFGR.DataSource = Nothing
    MFG_REC.Open "SELECT DISTINCT MANUFACTURER FROM ITEMMAST ORDER BY ITEMMAST.MANUFACTURER", db, adOpenForwardOnly
    Set CMBMFGR.RowSource = MFG_REC
    CMBMFGR.ListField = "MANUFACTURER"
    
    REPFLAG = True
    PHY_FLAG = True
    GRDSTOCK.TextMatrix(0, 0) = "Sl"
    GRDSTOCK.TextMatrix(0, 1) = "Item Code"
    GRDSTOCK.TextMatrix(0, 2) = "Item Description"
    GRDSTOCK.TextMatrix(0, 3) = "Qty"
    GRDSTOCK.TextMatrix(0, 4) = "Company"
    GRDSTOCK.TextMatrix(0, 5) = "Cost"
    GRDSTOCK.TextMatrix(0, 6) = "Tax"
    GRDSTOCK.TextMatrix(0, 7) = "Net Cost"
    GRDSTOCK.TextMatrix(0, 8) = "MRP"
    GRDSTOCK.TextMatrix(0, 9) = "Retail"
    GRDSTOCK.TextMatrix(0, 10) = "Wholesale"
    GRDSTOCK.TextMatrix(0, 11) = "Spl Price"
    GRDSTOCK.TextMatrix(0, 12) = "Price 5"
    GRDSTOCK.TextMatrix(0, 13) = "Price 6"
    GRDSTOCK.TextMatrix(0, 14) = "Price 7"
    
        
    GRDSTOCK.ColWidth(0) = 400
    GRDSTOCK.ColWidth(1) = 1500
    GRDSTOCK.ColWidth(2) = 5000
    GRDSTOCK.ColWidth(3) = 1100
    GRDSTOCK.ColWidth(4) = 2500
    GRDSTOCK.ColWidth(5) = 1100
    GRDSTOCK.ColWidth(6) = 900
    GRDSTOCK.ColWidth(7) = 1100
    GRDSTOCK.ColWidth(8) = 1100
    GRDSTOCK.ColWidth(9) = 1100
    GRDSTOCK.ColWidth(10) = 1100
    GRDSTOCK.ColWidth(11) = 1100
    GRDSTOCK.ColWidth(12) = 1100
    GRDSTOCK.ColWidth(13) = 1100
    GRDSTOCK.ColWidth(14) = 1100
    GRDSTOCK.ColWidth(15) = 0
    GRDSTOCK.ColWidth(16) = 0

    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 1
    GRDSTOCK.ColAlignment(5) = 4
    GRDSTOCK.ColAlignment(6) = 4
    GRDSTOCK.ColAlignment(7) = 4
    GRDSTOCK.ColAlignment(8) = 4
    GRDSTOCK.ColAlignment(9) = 4
    GRDSTOCK.ColAlignment(10) = 4
    GRDSTOCK.ColAlignment(11) = 4
    GRDSTOCK.ColAlignment(12) = 4
    GRDSTOCK.ColAlignment(13) = 4
    GRDSTOCK.ColAlignment(14) = 4
    
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
'        If RSTCOMPANY!hide_ws = "Y" Then
'            chkws.value = 1
'        Else
'            chkws.value = 0
'        End If
'        If RSTCOMPANY!hide_van = "Y" Then
'            chkvp.value = 1
'        Else
'            chkvp.value = 0
'        End If
'        If RSTCOMPANY!hide_lwp = "Y" Then
'            chklwp.value = 1
'        Else
'            chklwp.value = 0
'        End If
'        If RSTCOMPANY!hide_category = "Y" Then
'            chkhidecat.value = 1
'        Else
'            chkhidecat.value = 0
'        End If
        If RSTCOMPANY!hide_company = "Y" Then
            GRDSTOCK.ColWidth(4) = 0
        End If
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    DTFROM.Value = Format(Date, "DD/MM/YYYY")
    'Call Fillgrid
    'Me.Height = 8415
    'Me.Width = 6465
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    If PHY_FLAG = False Then PHY_REC.Close
    MFG_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDSTOCK_Click()
    TXTsample.Visible = False
    CMBMFGR.Visible = False
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                Select Case GRDSTOCK.Col
                    Case 1, 2, 5, 6, 8, 9, 10, 11, 12, 13, 14
                        TXTsample.Visible = True
                        TXTsample.Top = GRDSTOCK.CellTop + 100
                        TXTsample.Left = GRDSTOCK.CellLeft '+ 60
                        TXTsample.Width = GRDSTOCK.CellWidth
                        TXTsample.Height = GRDSTOCK.CellHeight
                        TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        TXTsample.SetFocus
'                    Case 12
'                        CMBMFGR.Visible = True
'                        CMBMFGR.Top = GRDSTOCK.CellTop + 100
'                        CMBMFGR.Left = GRDSTOCK.CellLeft '+ 60
'                        CMBMFGR.Width = GRDSTOCK.CellWidth
'                        'CmbPack.Height = GRDSTOCK.CellHeight
'                        On Error Resume Next
'                        CMBMFGR.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
'                        CMBMFGR.SetFocus
                End Select
            End If
        Case 114
            sitem = UCase(InputBox("Item Name...?", "STOCK"))
            For i = 1 To GRDSTOCK.rows - 1
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
    On Error GoTo ErrHand
    Call Fillgrid
    If REPFLAG = True Then
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
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
ErrHand:
    MsgBox err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.text)
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

    On Error GoTo ErrHand
    
    i = 0
    Screen.MousePointer = vbHourglass
    
    GRDSTOCK.rows = 1
    Set rststock = New ADODB.Recordset
    'WHERE ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%'
    'rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & DataList2.BoundText & "'ORDER BY VCH_NO DESC", db, adOpenStatic, adLockReadOnly
    rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly
    If Not (rststock.EOF And rststock.BOF) Then
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = rststock!CLOSE_QTY
        GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 2), "0.00"))
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 2), "0.00"))
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!CRTN_PACK), "", rststock!CRTN_PACK)
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!PACK_TYPE), "", rststock!PACK_TYPE)
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!P_CRTN), "", Format(Round(rststock!P_CRTN, 2), "0.00"))
        GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!P_LWS), "", Format(Round(rststock!P_LWS, 2), "0.00"))
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

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
    
End Sub

Private Function Fillgrid()
    Dim rststock As ADODB.Recordset
    Dim rstopstock As ADODB.Recordset
    Dim i As Long

    On Error GoTo ErrHand
    
    i = 0
    Screen.MousePointer = vbHourglass
    GRDSTOCK.rows = 1
    Set rststock = New ADODB.Recordset
    If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
        If OptStock.Value = True Then
            'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        ElseIf OptPLU.Value = True Then
            rststock.Open "SELECT * FROM ITEMMAST WHERE LENGTH(PLU_CODE)>0 and ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
    Else
        If CHKCATEGORY2.Value = 1 Then
            If OptStock.Value = True Then
                'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            ElseIf OptPLU.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE LENGTH(PLU_CODE)>0 and ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            ElseIf OptPLU.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE LENGTH(PLU_CODE)>0 and ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                'rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.Text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtCode.text) & "%' OR ITEM_CODE Like '%" & Me.TxtCode.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If

        End If
    End If
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!CLOSE_QTY), "", rststock!CLOSE_QTY)
        GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!MANUFACTURER), "", rststock!MANUFACTURER)
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST, "0.00"))
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!SALES_TAX), "0.00", Format(rststock!SALES_TAX, "0.00"))
        
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.00"))
        GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 2), "0.00"))
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 2), "0.00"))
        GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!P_VAN), "", Format(Round(rststock!P_VAN, 2), "0.00"))
        GRDSTOCK.TextMatrix(i, 12) = IIf(IsNull(rststock!PRICE5), "", Format(Round(rststock!PRICE5, 2), "0.00"))
        GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!PRICE6), "", Format(Round(rststock!PRICE6, 2), "0.00"))
        GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!PRICE7), "", Format(Round(rststock!PRICE7, 2), "0.00"))
        GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!CESS_PER), "", Format(rststock!CESS_PER, "0.00"))
        GRDSTOCK.TextMatrix(i, 16) = IIf(IsNull(rststock!cess_amt), "", Format(rststock!cess_amt, "0.00"))
                
        GRDSTOCK.TextMatrix(i, 7) = Format(Round((Val(GRDSTOCK.TextMatrix(i, 5)) + (Val(GRDSTOCK.TextMatrix(i, 5)) * Val(GRDSTOCK.TextMatrix(i, 6)) / 100)) + (Val(GRDSTOCK.TextMatrix(i, 5)) * Val(GRDSTOCK.TextMatrix(i, 15)) / 100) + Val(GRDSTOCK.TextMatrix(i, 16)), 4), "0.00")
'        GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!Category), "", rststock!Category)
'        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!ITEM_SPEC), "", rststock!ITEM_SPEC)
'        GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!Category), "", rststock!Category)
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = vbCtrlMask Then Call Clipboard_fn(KeyCode, Shift, TXTsample)
    Dim rststock, RSTITEMMAST As ADODB.Recordset
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 1  ' Item Code
                    If Trim(TXTsample.text) = "" Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_CODE = Trim(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!ITEM_CODE = Trim(TXTsample.text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from TRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!ITEM_CODE = Trim(TXTsample.text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 2  ' Item Name
                    If Trim(TXTsample.text) = "" Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_NAME = Trim(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!ITEM_NAME = Trim(TXTsample.text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
'                Case 3  'LOC
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (rststock.EOF And rststock.BOF) Then
'                        rststock!BIN_LOCATION = Trim(TXTsample.Text)
'                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
'                        rststock.Update
'                    End If
'                    rststock.Close
'                    Set rststock = Nothing
'                    GRDSTOCK.Enabled = True
'                    TXTsample.Visible = False
'                    GRDSTOCK.SetFocus
                
                
                                    
                Case 6  'TAX
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!SALES_TAX = Val(TXTsample.text)
                        rststock!CHECK_FLAG = "V"
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Val(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7) = Format(Round((Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6)) / 100)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15)) / 100) + Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16)), 4), "0.00")
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                                    
                Case 5  'COST
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_COST = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7) = Format(Round((Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6)) / 100)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15)) / 100) + Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16)), 4), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 8  'MRP
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!MRP = Val(TXTsample.text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 9  'RT
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT P_RETAIL from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_RETAIL) AND P_RETAIL <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_RETAIL = Val(TXTsample.text)
                        rststock!P_CRTN = Round(Val(TXTsample.text) / IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, 1, rststock!LOOSE_PACK), 2)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!P_RETAIL = Val(TXTsample.text)
                        rststock!P_CRTN = Round(Val(TXTsample.text) / IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, 1, rststock!LOOSE_PACK), 2)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                        
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 10  'WS
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT P_WS from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_WS) AND P_WS <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_WS = Val(TXTsample.text)
                        rststock!P_LWS = Round(Val(TXTsample.text) / IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, 1, rststock!LOOSE_PACK), 2)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!P_WS = Val(TXTsample.text)
                        rststock!P_LWS = Round(Val(TXTsample.text) / IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, 1, rststock!LOOSE_PACK), 2)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                        
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 11  'VP
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT P_VAN from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_VAN) AND P_VAN <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_VAN = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!P_VAN = Val(TXTsample.text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                        
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 12  'PRICE5
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT PRICE5 from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(PRICE5) AND PRICE5 <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!PRICE5 = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!PRICE5 = Val(TXTsample.text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                        
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 13  'PRICE6
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT PRICE6 from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(PRICE6) AND PRICE6 <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!PRICE6 = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!PRICE6 = Val(TXTsample.text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                        
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                                    
                Case 14  'PRICE7
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT PRICE7 from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(PRICE7) AND PRICE7 <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!PRICE7 = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!PRICE7 = Val(TXTsample.text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                        
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
'                Case 10  'SPEC
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (rststock.EOF And rststock.BOF) Then
'                        rststock!ITEM_SPEC = Trim(TXTsample.Text)
'                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
'                        rststock.Update
'                    End If
'                    rststock.Close
'                    Set rststock = Nothing
'                    GRDSTOCK.Enabled = True
'                    TXTsample.Visible = False
'                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 1, 2, 4
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub TXTCODE_Change()
    On Error GoTo ErrHand
    Call Fillgrid
    If REPFLAG = True Then
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
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
    
    On Error GoTo ErrHand
    If flagchange2.Caption <> "1" Then
        If chkcategory.Value = 1 Then
            If PHY_FLAG = True Then
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
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
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
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
ErrHand:
    MsgBox err.Description
    
End Sub


Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.text)
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
        
    TXTDEALER2.text = DataList1.text
    LBLDEALER2.Caption = TXTDEALER2.text
    Call Fillgrid
    tXTMEDICINE.SetFocus
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER2.text) = "" Then Exit Sub
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
    flagchange2.Caption = 1
    TXTDEALER2.text = LBLDEALER2.Caption
    DataList1.text = TXTDEALER2.text
    Call DataList1_Click
    'CHKCATEGORY2.value = 1
End Sub

Private Sub DataList1_LostFocus()
     flagchange2.Caption = ""
End Sub

Private Sub TXTsample_LostFocus()
    TXTsample.Visible = False
End Sub
