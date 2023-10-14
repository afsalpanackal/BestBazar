VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMDAMREG 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAMAGED GOODS REGISTER"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMDAMREG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   10755
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   765
      TabIndex        =   7
      Top             =   1890
      Visible         =   0   'False
      Width           =   9375
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4005
         Left            =   60
         TabIndex        =   8
         Top             =   540
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   7064
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColor       =   0
         ForeColor       =   16777215
         BackColorFixed  =   255
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
      Begin VB.Label LBLBILLAMT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6930
         TabIndex        =   12
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "BILL AMT"
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
         Index           =   1
         Left            =   5940
         TabIndex        =   11
         Top             =   210
         Width           =   885
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "BILL NO."
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
         Left            =   3825
         TabIndex        =   10
         Top             =   180
         Width           =   840
      End
      Begin VB.Label LBLBILLNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4710
         TabIndex        =   9
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00FFFF80&
      Height          =   10560
      Left            =   0
      TabIndex        =   0
      Top             =   -105
      Width           =   10875
      Begin VB.CommandButton CMDPRINTREGISTER 
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
         Left            =   9060
         TabIndex        =   19
         Top             =   9090
         Width           =   1635
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "&EXIT"
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
         Left            =   6060
         TabIndex        =   5
         Top             =   9090
         Width           =   1200
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
         Height          =   435
         Left            =   4725
         TabIndex        =   4
         Top             =   9090
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1185
         Left            =   105
         TabIndex        =   1
         Top             =   9090
         Width           =   8985
         Begin VB.Label LBLTRXTOTAL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   2205
            TabIndex        =   3
            Top             =   15
            Width           =   2220
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL AMOUNT"
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
            Height          =   315
            Index           =   3
            Left            =   195
            TabIndex        =   2
            Top             =   90
            Width           =   1935
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   7725
         Left            =   165
         TabIndex        =   6
         Top             =   1065
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   13626
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00FFFF80&
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
         Height          =   855
         Left            =   150
         TabIndex        =   13
         Top             =   90
         Width           =   10560
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00FFFF80&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   75
            TabIndex        =   18
            Top             =   420
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1860
            TabIndex        =   14
            Top             =   330
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   94568449
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   4035
            TabIndex        =   15
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   94568449
            CurrentDate     =   40498
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
            TabIndex        =   17
            Top             =   405
            Width           =   285
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
            TabIndex        =   16
            Top             =   405
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "FRMDAMREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim CLOSEALL As Integer

Private Sub CMDDISPLAY_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    Dim FROMDATE As Date
    Dim TODATE As Date
    Dim i As Integer

    
    LBLTRXTOTAL.Caption = ""
    On Error GoTo eRRHAND
    
    FROMDATE = Format(DTFROM.value, "MM,DD,YYYY")
    TODATE = Format(DTTO.value, "MM,DD,YYYY")
    
    GRDTranx.Rows = 1
    'If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    i = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT DISTINCT VCH_NO From DAMAGED WHERE [VCH_DATE] <=# " & TODATE & " # AND [VCH_DATE] >=# " & FROMDATE & " # ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        i = i + 1
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(i, 0) = i
        GRDTranx.TextMatrix(i, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(i, 5) = rstTRANX!TRX_TYPE
        GRDTranx.TextMatrix(i, 2) = rstTRANX!VCH_DATE
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * From DAMAGED WHERE VCH_NO = " & rstTRANX!VCH_NO & "", db, adOpenStatic, adLockReadOnly
        Do Until RSTTRXFILE.EOF
            GRDTranx.TextMatrix(i, 3) = Format(Val(GRDTranx.TextMatrix(i, 4)) + RSTTRXFILE!TRX_TOTAL, "0.00")
            GRDTranx.TextMatrix(i, 4) = IIf(IsNull(RSTTRXFILE!VCH_DESC), "", Mid(RSTTRXFILE!VCH_DESC, 15))
            LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + RSTTRXFILE!TRX_TOTAL, "0.00")
            RSTTRXFILE.MoveNext
        Loop
        GRDTranx.Col = 4
        GRDTranx.Row = i
        GRDTranx.CellForeColor = vbRed
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        rstTRANX.MoveNext
    Loop
    
    GRDTranx.Visible = True
    If i > 22 Then GRDTranx.TopRow = i
    GRDTranx.SetFocus
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDPRINTREGISTER_Click()
    rptPRINT.ReportFileName = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTSALESREG"
    rptPRINT.Formulas(0) = "PERIOD = '" & DTFROM.value & " " & " TO " & " " & DTTO.value & "'"
    rptPRINT.Action = 1
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "DG NO"
    GRDTranx.TextMatrix(0, 5) = "TYPE"
    GRDTranx.TextMatrix(0, 2) = "DG DATE"
    GRDTranx.TextMatrix(0, 3) = "AMOUNT"
    GRDTranx.TextMatrix(0, 4) = "REMARKS"
    
    
    GRDTranx.ColWidth(0) = 700
    GRDTranx.ColWidth(1) = 800
    GRDTranx.ColWidth(5) = 0
    GRDTranx.ColWidth(2) = 1350
    GRDTranx.ColWidth(3) = 1200
    GRDTranx.ColWidth(4) = 2800
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 3
    GRDTranx.ColAlignment(5) = 3
    GRDTranx.ColAlignment(2) = 3
    GRDTranx.ColAlignment(3) = 6
    GRDTranx.ColAlignment(4) = 1
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "MRP"
    GRDBILL.TextMatrix(0, 3) = "Rate"
    GRDBILL.TextMatrix(0, 4) = "Qty"
    GRDBILL.TextMatrix(0, 5) = "Amount"
    GRDBILL.TextMatrix(0, 6) = "Serial No"
    GRDBILL.TextMatrix(0, 7) = ""
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 2500
    GRDBILL.ColWidth(2) = 900
    GRDBILL.ColWidth(3) = 900
    GRDBILL.ColWidth(4) = 900
    GRDBILL.ColWidth(5) = 1100
    GRDBILL.ColWidth(6) = 1100
    GRDBILL.ColWidth(7) = 0
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 6
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 6
    GRDBILL.ColAlignment(6) = 1
    
    CLOSEALL = 1
    ACT_FLAG = True
    Month (Date) - 2
    DTFROM.value = "01/" & Month(Date) & "/" & Year(Date)
    DTTO.value = Format(Date, "DD/MM/YYYY")
    Me.Width = 10845
    Me.Height = 11025
    Me.Left = 1500
    Me.Top = 0
    txtpassword = "YEAR " & Year(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If ACT_FLAG = False Then ACT_REC.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            
            FRMEMAIN.Enabled = True
            FRMEBILL.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.Rows = 1 Then Exit Sub
            LBLBILLNO.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            LBLBILLAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
             
            GRDBILL.Rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From DAMAGED WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = '" & Trim(GRDTranx.TextMatrix(GRDTranx.Row, 5)) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.Rows = GRDBILL.Rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ITEM_NAME
                GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!MRP, "0.00")
                GRDBILL.TextMatrix(i, 3) = Format(RSTTRXFILE!SALES_PRICE, "0.00")
                GRDBILL.TextMatrix(i, 4) = RSTTRXFILE!QTY
                GRDBILL.TextMatrix(i, 5) = Format(RSTTRXFILE!SALES_PRICE * RSTTRXFILE!QTY, "0.00")
                GRDBILL.TextMatrix(i, 6) = RSTTRXFILE!REF_NO
                GRDBILL.TextMatrix(i, 7) = ""
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            FRMEMAIN.Enabled = False
            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    End Select
End Sub
