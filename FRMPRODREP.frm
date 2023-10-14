VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FRMPRODREP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRODUCTION REPORT"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMPRODREP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   9435
   Begin VB.Frame FRMEBILL 
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H00000080&
      Height          =   4725
      Left            =   1890
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   6045
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4125
         Left            =   30
         TabIndex        =   5
         Top             =   540
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   7276
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
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
         Height          =   255
         Index           =   0
         Left            =   3630
         TabIndex        =   7
         Top             =   210
         Width           =   780
      End
      Begin VB.Label LBLBILLNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4125
         TabIndex        =   6
         Top             =   165
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      Caption         =   "Frame1"
      Height          =   8955
      Left            =   -120
      TabIndex        =   0
      Top             =   -270
      Width           =   9555
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0C0FF&
         Caption         =   "PRODUCTION REPORT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   840
         Left            =   150
         TabIndex        =   10
         Top             =   285
         Width           =   9375
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0C0FF&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   75
            TabIndex        =   11
            Top             =   420
            Value           =   -1  'True
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1860
            TabIndex        =   12
            Top             =   330
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   123076609
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   4035
            TabIndex        =   13
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   123076609
            CurrentDate     =   40498
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0FF&
            Height          =   585
            Left            =   5625
            TabIndex        =   18
            Top             =   150
            Width           =   3675
            Begin VB.OptionButton OPTPROD2 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Production-2"
               ForeColor       =   &H00FF0000&
               Height          =   270
               Left            =   1815
               TabIndex        =   20
               Top             =   190
               Width           =   1575
            End
            Begin VB.OptionButton OPTPROD1 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Production-1"
               ForeColor       =   &H00FF0000&
               Height          =   270
               Left            =   75
               TabIndex        =   19
               Top             =   190
               Value           =   -1  'True
               Width           =   1545
            End
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   270
            Index           =   4
            Left            =   1110
            TabIndex        =   17
            Top             =   405
            Width           =   555
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
            ForeColor       =   &H00800080&
            Height          =   270
            Index           =   5
            Left            =   3585
            TabIndex        =   16
            Top             =   405
            Width           =   285
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   6465
            TabIndex        =   15
            Top             =   1965
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   8685
            TabIndex        =   14
            Top             =   1905
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.CommandButton CMDREGISTER 
         Caption         =   "PRINT REGISTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3555
         TabIndex        =   9
         Top             =   8295
         Width           =   1515
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6600
         TabIndex        =   2
         Top             =   8280
         Width           =   1335
      End
      Begin VB.CommandButton CMDDISPLAY 
         Caption         =   "&DISPLAY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5130
         TabIndex        =   1
         Top             =   8280
         Width           =   1380
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6660
         Left            =   165
         TabIndex        =   3
         Top             =   1125
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   11748
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
         Height          =   420
         Left            =   2310
         TabIndex        =   8
         Tag             =   "5"
         Top             =   7815
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   741
         Picture         =   "FRMPRODREP.frx":030A
         ForeColor       =   0
         BarPicture      =   "FRMPRODREP.frx":0326
         Max             =   150
         Text            =   "PLEASE WAIT..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
   End
End
Attribute VB_Name = "FRMPRODREP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmDDisplay_Click()
    Dim rstTRANX, RSTTRXFILE As ADODB.Recordset
    Dim n, M As Long
    
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    n = 1
    M = 0
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    If OPTPROD1.Value = True Then
        If OPTPERIOD.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='MI' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='MI' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    Else
        If OPTPERIOD.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='RM' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='RM' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    End If
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If
    
    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!TRX_TYPE
        GRDTranx.TextMatrix(M, 2) = "Production"
        GRDTranx.TextMatrix(M, 3) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 4) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 5) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
        
        Set RSTTRXFILE = New ADODB.Recordset
        If OPTPROD1.Value = True Then
            RSTTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_TYPE='MI' AND VCH_NO = " & rstTRANX!VCH_NO & " ", db, adOpenStatic, adLockReadOnly, adCmdText
        Else
            RSTTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_TYPE='RM' AND VCH_NO = " & rstTRANX!VCH_NO & " ", db, adOpenStatic, adLockReadOnly, adCmdText
        End If
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            GRDTranx.TextMatrix(M, 6) = RSTTRXFILE!QTY
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing

        CMDDISPLAY.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
            
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        n = n + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
        
    flagchange.Caption = ""
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub CMDDISPLAY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            DTTo.SetFocus
    End Select
End Sub

Private Sub CMDREGISTER_Click()
    
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    ReportNameVar = Rptpath & "RPTRAWREG"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
    If OPTPROD1.Value = True Then
        Report.RecordSelectionFormula = "( {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND {TRXMAST.TRX_TYPE} ='MI')"
    Else
        Report.RecordSelectionFormula = "( {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND {TRXMAST.TRX_TYPE} ='RM')"
    End If
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub DTFROM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DTTo.SetFocus
    End Select
End Sub

Private Sub DTTO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            DTFROM.SetFocus
    End Select
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'If Month(Date) > 1 Then
        'CMBMONTH.ListIndex = Month(Date) - 2
    'Else
        'CMBMONTH.ListIndex = 11
    'End If
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "Type"
    GRDTranx.TextMatrix(0, 2) = "TYPE"
    GRDTranx.TextMatrix(0, 3) = "BILL NO"
    GRDTranx.TextMatrix(0, 4) = "DATE"
    GRDTranx.TextMatrix(0, 5) = "Formula"
    GRDTranx.TextMatrix(0, 6) = "Qty"
    
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 0
    GRDTranx.ColWidth(2) = 0
    GRDTranx.ColWidth(3) = 1500
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 3700
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 3
    GRDTranx.ColAlignment(4) = 3
    GRDTranx.ColAlignment(5) = 1
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "Qty"
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 3800
    GRDBILL.ColWidth(2) = 800
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 11130
    'Me.Height = 10125
    Me.Left = 1500
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRMEMAIN.Enabled = True
            FRMEBILL.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDBILL_LostFocus()
    If FRMEBILL.Visible = True Then
        FRMEBILL.Visible = False
        GRDTranx.SetFocus
    End If
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.rows = 1 Then Exit Sub
            LBLBILLNO.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
             
            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILE WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "  AND TRX_TYPE = '" & Trim(GRDTranx.TextMatrix(GRDTranx.Row, 1)) & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.rows = GRDBILL.rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ITEM_NAME
                GRDBILL.TextMatrix(i, 2) = RSTTRXFILE!QTY
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        
            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    End Select
End Sub


