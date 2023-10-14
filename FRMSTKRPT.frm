VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMSTOCK 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   ControlBox      =   0   'False
   Icon            =   "FRMSTKRPT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   10950
   Begin VB.CommandButton cmddisplay 
      Caption         =   "&Print Company wise stock report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6570
      TabIndex        =   20
      Top             =   8730
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   105
      TabIndex        =   10
      Top             =   -135
      Width           =   10815
      Begin VB.TextBox TXTMRP 
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
         Left            =   3810
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1725
         Width           =   900
      End
      Begin VB.TextBox TXTDEALER 
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
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   495
         Width           =   4575
      End
      Begin VB.TextBox txtmedsearch 
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
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   5
         Top             =   840
         Width           =   3060
      End
      Begin VB.TextBox TXTSTCKAMT 
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
         Left            =   1515
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1710
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.TextBox txtmedname 
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
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   4
         Top             =   450
         Width           =   3645
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   840
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1482
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         Height          =   930
         Left            =   5130
         TabIndex        =   21
         Top             =   1155
         Width           =   5550
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   345
            Left            =   1590
            TabIndex        =   22
            Top             =   525
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   609
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   51838977
            CurrentDate     =   41275
            MaxDate         =   42004
            MinDate         =   41275
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   345
            Left            =   3765
            TabIndex        =   23
            Top             =   540
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   609
            _Version        =   393216
            Format          =   51838977
            CurrentDate     =   41275
            MaxDate         =   42004
            MinDate         =   41275
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "for the Period"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   7
            Left            =   120
            TabIndex        =   27
            Top             =   540
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   150
            Width           =   1080
         End
         Begin MSForms.ComboBox cmbcompany 
            Height          =   360
            Left            =   1590
            TabIndex        =   25
            Top             =   135
            Width           =   3645
            VariousPropertyBits=   746604571
            ForeColor       =   255
            DisplayStyle    =   7
            Size            =   "6429;635"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   255
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   5
            Left            =   3315
            TabIndex        =   24
            Top             =   600
            Visible         =   0   'False
            Width           =   285
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MRP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   6
         Left            =   3360
         TabIndex        =   18
         Top             =   1770
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   195
         TabIndex        =   14
         Top             =   255
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name Contains"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   5250
         TabIndex        =   13
         Top             =   885
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1770
         Visible         =   0   'False
         Width           =   1425
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   5250
         TabIndex        =   11
         Top             =   525
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8535
      TabIndex        =   6
      Top             =   8730
      Width           =   1125
   End
   Begin VB.CommandButton CMDEXIT 
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
      Height          =   510
      Left            =   9720
      TabIndex        =   8
      Top             =   8715
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid grdsTOCK 
      Height          =   6180
      Left            =   105
      TabIndex        =   19
      Top             =   2460
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   10901
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      HighLight       =   0
      AllowUserResizing=   1
      Appearance      =   0
      GridLineWidth   =   2
   End
   Begin VB.Label LBLCAPTION 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   135
      TabIndex        =   17
      Top             =   2055
      Width           =   9600
   End
   Begin VB.Label lbldealer 
      BackColor       =   &H00FF80FF&
      Height          =   315
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label flagchange 
      BackColor       =   &H00FF80FF&
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LBLSTAOCKVALUE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   2145
      TabIndex        =   9
      Top             =   8685
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Stock Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   8745
      Width           =   2235
   End
End
Attribute VB_Name = "FRMSTOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim CLOSEALL As Integer

Private Sub cmbcompany_Change()
    TXTDEALER.Text = ""
    TXTSTCKAMT.Text = ""
    txtmedsearch.Text = ""
    'cmbcompany.ListIndex = -1
    txtmedname.Text = ""
    txtmedsearch.Text = ""
    TxtMRP.Text = ""
    Call fillstockgrid(5)
    LBLCAPTION.Caption = "LIST OF AVAILABLE STOCK ITEMS WITH COMPANY NAME  """ & cmbcompany.Text & """"
End Sub

Private Sub CmdDelete_Click()
''    Dim PHY As ADODB.Recordset
''    Dim rststock As ADODB.Recordset
''    '''If grdsTOCK.ApproxCount < 1 Then Exit Sub
''    If MsgBox("Are You Sure You want to Delete " & "*** " & grdsTOCK.Columns(1).Text & " ****", vbYesNo, "DELETING .......") = vbNo Then Exit Sub
''    Conn.Execute ("DELETE * FROM STOCK WHERE STOCK.STK_SL = '" & Val(grdsTOCK.Columns(0).Text) & "'")
''
''    Set PHY = New ADODB.Recordset
''    PHY.Open "SELECT [STK_SL],[STK_ITEM],[STK_BATCH],[STK_EXPDATE],[STK_UNIT],[STK_MRP],[STK_QTY],[STK_VALUE] from [STOCK] ORDER BY STK_ITEM ", Conn, adOpenStatic,adLockReadOnly
''    Set grdsTOCK.DataSource = PHY
''    grdsTOCK.Columns(0).Caption = "SL"
''    grdsTOCK.Columns(1).Caption = "PRODUCT"
''    grdsTOCK.Columns(2).Caption = "Serial No"
''    grdsTOCK.Columns(3).Caption = "EXP DATE"
''    grdsTOCK.Columns(4).Caption = "UNIT"
''    grdsTOCK.Columns(5).Caption = "RATE"
''    grdsTOCK.Columns(6).Caption = "QTY"
''    grdsTOCK.Columns(7).Caption = "VALUE"
''
''    grdsTOCK.Columns(0).Width = 0
''    grdsTOCK.Columns(2).Width = 1000
''    grdsTOCK.Columns(3).Width = 1400
''    grdsTOCK.Columns(4).Width = 1000
''    grdsTOCK.Columns(5).Width = 1000
''    grdsTOCK.Columns(6).Width = 800
''    grdsTOCK.Columns(7).Width = 1000
''
''     'grdSTOCK.Columns(2).Alignment = dbgCenter
''
''    grdsTOCK.RowHeight = 270
''    LBLSTAOCKVALUE.Caption = ""
''    Set rststock = New ADODB.Recordset
''    rststock.Open "SELECT [STK_VALUE] from [STOCK]", Conn, adOpenStatic,adLockReadOnly
''    Do Until rststock.EOF
''        LBLSTAOCKVALUE.Caption = Val(LBLSTAOCKVALUE.Caption) + rststock!STK_VALUE
''        rststock.MoveNext
''    Loop
''    rststock.Close
''    Set rststock = Nothing
''    'Dim n As Integer
''    'n = Val(grdSTOCK.Columns(0).Text)
''    '
 
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo eRRHAND
    
    Screen.MousePointer = vbNormal
    Exit Sub
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description

End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDDISPLAY_Click()
    Dim OPQTY, OPVAL, CLOQTY, CLOVAL, RCVDQTY, RCVDVAL, ISSQTY, ISSVAL, DAMQTY, DAMVAL, FREEQTY, FREEVAL, SAMPLEQTY, SAMPLEVAL As Double
    Dim rststock As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    Dim FROMDATE As Date
    Dim TODATE As Date
    
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTO.Value, "MM,DD,YYYY")
    
    If cmbcompany.ListIndex = -1 Then
        MsgBox "Select Company from the list", vbOKOnly, "Report"
        cmbcompany.SetFocus
        Exit Sub
    End If
    
    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    db2.Execute "delete * From STOCKREPORT"
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM STOCKREPORT", db2, adOpenStatic, adLockOptimistic, adCmdText
    Set RSTITEM = New ADODB.Recordset
    RSTITEM.Open "SELECT *  FROM ITEMMAST WHERE MANUFACTURER = '" & cmbcompany.Text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTITEM.EOF
        OPQTY = 0
        OPVAL = 0
        RCVDQTY = 0
        RCVDVAL = 0
        ISSQTY = 0
        ISSVAL = 0
        CLOQTY = 0
        CLOVAL = 0
        DAMQTY = 0
        DAMVAL = 0
        FREEQTY = 0
        FREEVAL = 0
        SAMPLEQTY = 0
        SAMPLEVAL = 0
        
        OPQTY = RSTITEM!OPEN_QTY
        OPVAL = RSTITEM!OPEN_VAL
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & RSTITEM!ITEM_CODE & "' AND [VCH_DATE] <# " & DTFROM & " #", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            OPQTY = OPQTY + RSTTRXFILE!QTY
            OPVAL = OPVAL + RSTTRXFILE!TRX_TOTAL
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & RSTITEM!ITEM_CODE & "' AND [VCH_DATE] >=# " & DTFROM & " # AND [VCH_DATE] <=# " & DTTO & " #", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            RCVDQTY = RCVDQTY + RSTTRXFILE!QTY
            RCVDVAL = RCVDVAL + RSTTRXFILE!TRX_TOTAL
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEM!ITEM_CODE & "' AND ((TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='PR' OR TRX_TYPE='DG' OR TRX_TYPE='GF') AND [VCH_DATE] >=# " & DTFROM & " # AND [VCH_DATE] <=# " & DTTO & " # ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DG"
                    DAMQTY = DAMQTY + RSTTRXFILE!QTY
                    DAMVAL = DAMVAL + IIf(IsNull(RSTTRXFILE!TRX_TOTAL), 0, RSTTRXFILE!TRX_TOTAL)
                Case "GF"
                    SAMPLEQTY = SAMPLEQTY + RSTTRXFILE!QTY
                    SAMPLEVAL = SAMPLEVAL + IIf(IsNull(RSTTRXFILE!TRX_TOTAL), 0, RSTTRXFILE!TRX_TOTAL)
                Case Else
                    ISSQTY = ISSQTY + RSTTRXFILE!QTY
                    ISSVAL = ISSVAL + IIf(IsNull(RSTTRXFILE!TRX_TOTAL), 0, RSTTRXFILE!TRX_TOTAL)
                    FREEQTY = FREEQTY + IIf(IsNull(RSTTRXFILE!FREE_QTY), 0, RSTTRXFILE!FREE_QTY)
                    FREEVAL = FREEVAL + RSTTRXFILE!SALES_PRICE * FREEQTY
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        CLOQTY = (OPQTY + RCVDQTY) - (ISSQTY + DAMQTY + FREEQTY + SAMPLEQTY)
        CLOVAL = (OPVAL + RCVDVAL) - (ISSVAL + DAMVAL + FREEVAL + SAMPLEVAL)
        
        rststock.AddNew
        rststock!ITEM_CODE = RSTITEM!ITEM_CODE
        rststock!ITEM_NAME = RSTITEM!ITEM_NAME
        rststock!UNIT = RSTITEM!UNIT
        rststock!ITEM_COST = RSTITEM!ITEM_COST
        rststock!MRP = RSTITEM!MRP
        rststock!OPEN_QTY = OPQTY
        rststock!OPEN_VAL = OPVAL
        rststock!RCPT_QTY = RCVDQTY
        rststock!RCPT_VAL = RCVDVAL
        rststock!ISSUE_QTY = ISSQTY
        rststock!ISSUE_VAL = ISSVAL
        rststock!CLOSE_QTY = CLOQTY
        rststock!CLOSE_VAL = CLOVAL
        rststock!DAM_QTY = DAMQTY
        rststock!DAM_VAL = DAMVAL
        rststock!FREE_QTY = FREEQTY
        rststock!FREE_VAL = FREEVAL
        rststock!SAMP_QTY = SAMPLEQTY
        rststock!SAMP_VAL = SAMPLEVAL
        
        rststock.Update
        
        RSTITEM.MoveNext
    Loop
    RSTITEM.Close
    Set RSTITEM = Nothing
    
    rststock.Close
    Set rststock = Nothing
    Screen.MousePointer = vbNormal
    MsgBox "Report Generated", vbOKOnly, "Sales Report"
    ReportNameVar = App.Path & "\RPTSTOCK.rpt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
    Set CRXFormulaFields = Report.FormulaFields
    
    For i = 1 To Report.Database.Tables.Count
        Report.Database.Tables(i).SetLogOnInfo "ConnectionName", "G:\dbase\YEAR13-14\MEDINV.MDB", "admin", "!@#$%^&*())(*&^%$#@!"
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & cmbcompany.Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "COMPANY WISE REPORT"
    Call GENERATEREPORT
    
    Exit Sub
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    'RSTCOMPANY.Open "SELECT DISTINCT [MANUFACTURER]FROM ITEMMAST RIGHT JOIN RTRXFILE ON ITEMMAST.ITEM_CODE = RTRXFILE.ITEM_CODE WHERE RTRXFILE.BAL_QTY > 0 ORDER BY [ITEMMAST.MANUFACTURER]", db, adOpenStatic,adLockReadOnly
    RSTCOMPANY.Open "SELECT DISTINCT [MANUFACTURER]FROM ITEMMAST  WHERE MANUFACTURER<>'' ORDER BY [MANUFACTURER]", db, adOpenStatic, adLockReadOnly
    Do Until RSTCOMPANY.EOF
        If Not IsNull(RSTCOMPANY!MANUFACTURER) Then cmbcompany.AddItem (RSTCOMPANY!MANUFACTURER)
        RSTCOMPANY.MoveNext
    Loop
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    GRDSTOCK.ColWidth(0) = 300
    GRDSTOCK.ColWidth(1) = 600
    GRDSTOCK.ColWidth(2) = 2900
    GRDSTOCK.ColWidth(3) = 1200
    GRDSTOCK.ColWidth(4) = 0
    GRDSTOCK.ColWidth(5) = 0
    GRDSTOCK.ColWidth(6) = 800
    GRDSTOCK.ColWidth(7) = 800
    GRDSTOCK.ColWidth(8) = 800
    GRDSTOCK.ColWidth(9) = 1200
    
    GRDSTOCK.ColAlignment(0) = 4
    'grdsTOCK.ColAlignment(1) = 4
    'grdsTOCK.ColAlignment(2) = 4
    'grdsTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 0
    GRDSTOCK.ColAlignment(5) = 3
    GRDSTOCK.ColAlignment(6) = 7
    GRDSTOCK.ColAlignment(7) = 7
    GRDSTOCK.ColAlignment(8) = 7
    GRDSTOCK.ColAlignment(9) = 7
    
    GRDSTOCK.TextArray(0) = "SL"
    GRDSTOCK.TextArray(1) = "ITEM CODE"
    GRDSTOCK.TextArray(2) = "ITEM NAME"
    GRDSTOCK.TextArray(3) = "Serial No"
    GRDSTOCK.TextArray(4) = ""
    GRDSTOCK.TextArray(5) = "" '"PACK"
    GRDSTOCK.TextArray(6) = "COST"
    GRDSTOCK.TextArray(7) = "MRP"
    GRDSTOCK.TextArray(8) = "QTY"
    GRDSTOCK.TextArray(9) = "VALUE"
    
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    CLOSEALL = 1
    ACT_FLAG = True

    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If CLOSEALL = 0 Then
        If ACT_FLAG = False Then ACT_REC.Close
        If MDIMAIN.PCTMENU.Visible = True Then
            MDIMAIN.PCTMENU.Enabled = True
            MDIMAIN.PCTMENU.SetFocus
        Else
            MDIMAIN.pctmenu2.Enabled = True
            MDIMAIN.pctmenu2.SetFocus
        End If
    End If
   Cancel = CLOSEALL
   
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            
    End Select

End Sub

Private Sub TXTDEALER_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo eRRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ACT_NAME
        End If
        Set Me.DataList2.RowSource = ACT_REC
        DataList2.ListField = "ACT_NAME"
        DataList2.BoundColumn = "ACT_CODE"
    End If
    Exit Sub
eRRHAND:
    MsgBox Err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.Text = DataList2.Text
    TXTSTCKAMT.Text = ""
    txtmedsearch.Text = ""
    cmbcompany.ListIndex = -1
    txtmedname.Text = ""
    txtmedsearch.Text = ""
    TxtMRP.Text = ""
    Call fillstockgrid(1)
    LBLCAPTION.Caption = "AVAILABLE STOCK OF ITEMS FROM  " & DataList2.Text
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "STOCK..."
                DataList2.SetFocus
                Exit Sub
            End If
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub txtmedname_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTDEALER.Text = ""
            TXTSTCKAMT.Text = ""
            txtmedsearch.Text = ""
            cmbcompany.ListIndex = -1
            'txtmedname.Text = ""
            txtmedsearch.Text = ""
            TxtMRP.Text = ""
            Call fillstockgrid(3)
            LBLCAPTION.Caption = "LIST OF AVAILABLE STOCK ITEMS STARTING WITH  """ & txtmedname.Text & """"
    End Select
End Sub

Private Sub txtmedsearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTDEALER.Text = ""
            TXTSTCKAMT.Text = ""
            cmbcompany.ListIndex = -1
            txtmedname.Text = ""
            'txtmedsearch.Text = ""
            TxtMRP.Text = ""
            Call fillstockgrid(4)
            LBLCAPTION.Caption = "LIST OF AVAILABLE STOCK ITEMS CONTAINING  """ & txtmedsearch.Text & """"
    End Select
End Sub

Private Sub TXTMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTDEALER.Text = ""
            TXTSTCKAMT.Text = ""
            txtmedsearch.Text = ""
            cmbcompany.ListIndex = -1
            txtmedname.Text = ""
            txtmedsearch.Text = ""
            'TXTMRP.Text = ""
            Call fillstockgrid(2)
            LBLCAPTION.Caption = "LIST OF AVAILABLE STOCK ITEMS WITH MRP  """ & TxtMRP.Text & """"
    End Select
End Sub

Private Sub TXTMRP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTSTCKAMT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim i As Double
    Dim n As Double
    i = 0
    n = 0
    Screen.MousePointer = vbHourglass
    GRDSTOCK.Rows = 1
    Select Case KeyCode
        Case vbKeyReturn
            txtmedname.Text = ""
            txtmedsearch.Text = ""
            cmbcompany.ListIndex = -1
            flagchange.Caption = "1"
            TXTDEALER.Text = ""
            flagchange.Caption = ""
            TxtMRP.Text = ""
            Screen.MousePointer = vbHourglass
    
            On Error GoTo eRRHAND

            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT DISTINCT * from [RTRXFILE] WHERE RTRXFILE.BAL_QTY > 0 ORDER BY ITEM_CODE ", db, adOpenStatic, adLockReadOnly
            Do Until rststock.EOF
                If IsNull(rststock!ITEM_CODE) Then GoTo SKIP
                If IsNull(rststock!ITEM_NAME) Then GoTo SKIP
                'If IsNull(rststock!REF_NO) Then GoTo SKIP
                'If IsNull(rststock!EXP_DATE) Then GoTo SKIP
                If IsNull(rststock!UNIT) Then GoTo SKIP
                If IsNull(rststock!SALES_PRICE) Then GoTo SKIP
                If IsNull(rststock!BAL_QTY) Then GoTo SKIP
                If IsNull(rststock!TRX_TOTAL) Then GoTo SKIP

                'If DateDiff("d", Date, rststock!EXP_DATE) < 31 Then GoTo SKIP
                GRDSTOCK.Rows = GRDSTOCK.Rows + 1
                GRDSTOCK.FixedRows = 1
                i = i + 1
                        
                GRDSTOCK.TextMatrix(i, 0) = i
                GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
                GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
                GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
                GRDSTOCK.TextMatrix(i, 4) = "" 'IIf(IsNull(rststock!EXP_DATE), "", rststock!EXP_DATE)
                GRDSTOCK.TextMatrix(i, 5) = Val(rststock!UNIT)
                GRDSTOCK.TextMatrix(i, 6) = Format(Val(rststock!ITEM_COST) * Val(rststock!UNIT), ".00")
                GRDSTOCK.TextMatrix(i, 7) = Format(Val(rststock!MRP), ".00")
                GRDSTOCK.TextMatrix(i, 8) = rststock!BAL_QTY
                GRDSTOCK.TextMatrix(i, 9) = Format(rststock!ITEM_COST * rststock!BAL_QTY, ".00")
                n = n + Val(GRDSTOCK.TextMatrix(i, 9))
                LBLSTAOCKVALUE.Caption = Format(n, ".00")
                
                If n > Val(TXTSTCKAMT) And TXTSTCKAMT.Text <> "000" Then Exit Do
SKIP:
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            TXTSTCKAMT.Text = Val(TXTSTCKAMT.Text)

        '    grdSTOCK.Columns(0).Visible = False
        LBLCAPTION.Caption = "LIST OF AVAILABLE STOCK FOR AN APPROX AMOUNT OF RS. " & Format(TXTSTCKAMT.Text, "0.00")
    End Select
        Screen.MousePointer = vbNormal
   Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub TXTSTCKAMT_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Function fillstockgrid(mstag As Integer)
    
    Dim rststock As ADODB.Recordset
    Dim PHY As ADODB.Recordset

    Dim i As Double
    Dim n As Double
    i = 0
    n = 0
    Screen.MousePointer = vbHourglass
    GRDSTOCK.Rows = 1

    On Error GoTo eRRHAND
    
    Set rststock = New ADODB.Recordset
    Select Case mstag
        Case 1
            rststock.Open "Select * From [RTRXFILE] WHERE M_USER_ID = '" & Me.DataList2.BoundText & "'AND RTRXFILE.BAL_QTY > 0 ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        Case 2
            rststock.Open "Select * From [RTRXFILE] WHERE MRP= " & Val(TxtMRP.Text) & " AND RTRXFILE.BAL_QTY > 0 ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        Case 3
            rststock.Open "Select * From [RTRXFILE] WHERE ITEM_NAME Like '" & Me.txtmedname.Text & "%'AND RTRXFILE.BAL_QTY > 0 ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        Case 4
            rststock.Open "Select * From [RTRXFILE] WHERE ITEM_NAME Like '%" & Me.txtmedsearch.Text & "%'AND RTRXFILE.BAL_QTY > 0 ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        Case 5
            rststock.Open "Select * From [RTRXFILE] WHERE MFGR = '" & cmbcompany.Text & "'AND RTRXFILE.BAL_QTY > 0 ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
'        Case 6
'            rststock.Open "SELECT RTRXFILE.ITEM_CODE,RTRXFILE.ITEM_NAME,RTRXFILE.REF_NO,RTRXFILE.EXP_DATE,RTRXFILE.UNIT,RTRXFILE.SALES_PRICE,RTRXFILE.BAL_QTY,RTRXFILE.TRX_TOTAL,RTRXFILE.ITEM_COST FROM RTRXFILE RIGHT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE ITEMMAST.MANUFACTURER Like '" & Me.cmbcompany.Text & "%'AND RTRXFILE.BAL_QTY > 0 ORDER BY [RTRXFILE.ITEM_NAME]", db, adOpenStatic, adLockReadOnly
    End Select
    
    Do Until rststock.EOF
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        i = i + 1
                
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
        GRDSTOCK.TextMatrix(i, 4) = "" 'IIf(IsNull(rststock!EXP_DATE), "", rststock!EXP_DATE)
        GRDSTOCK.TextMatrix(i, 5) = Val(rststock!UNIT)
        GRDSTOCK.TextMatrix(i, 6) = Format(Val(rststock!ITEM_COST) * Val(rststock!UNIT), ".00")
        GRDSTOCK.TextMatrix(i, 7) = Format(Val(rststock!MRP), ".00")
        GRDSTOCK.TextMatrix(i, 8) = rststock!BAL_QTY
        GRDSTOCK.TextMatrix(i, 9) = Format(rststock!ITEM_COST * rststock!BAL_QTY, ".00")
        n = n + Val(GRDSTOCK.TextMatrix(i, 9))
        LBLSTAOCKVALUE.Caption = Format(n, ".00")

        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
   
    
    Screen.MousePointer = vbNormal
    Exit Function
   
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
    
End Function

