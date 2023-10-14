VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaybookwo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACCOUNTS SUMMARY"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   ControlBox      =   0   'False
   Icon            =   "DaybookWO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   10065
   Begin VB.Frame FRAME 
      Height          =   3780
      Left            =   30
      TabIndex        =   2
      Top             =   -15
      Width           =   9990
      Begin VB.OptionButton OPTPERIOD 
         Caption         =   "PERIOD FROM"
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
         Height          =   210
         Left            =   75
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.CommandButton CmdDisplay 
         BackColor       =   &H00400000&
         Caption         =   "&DISPLAY"
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
         Left            =   7080
         MaskColor       =   &H80000007&
         TabIndex        =   0
         Top             =   3105
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00400000&
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
         Left            =   8550
         MaskColor       =   &H80000007&
         TabIndex        =   1
         Top             =   3105
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   1890
         TabIndex        =   4
         Top             =   195
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   48955393
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   4080
         TabIndex        =   5
         Top             =   195
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   48955393
         CurrentDate     =   40498
      End
      Begin VB.Label LBLEXPENSE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   8040
         TabIndex        =   28
         Top             =   1320
         Width           =   1845
      End
      Begin VB.Label LBLEXPENSES 
         BackStyle       =   0  'Transparent
         Caption         =   "Expenses"
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
         Left            =   7095
         TabIndex        =   27
         Top             =   1365
         Width           =   885
      End
      Begin VB.Line Line1 
         X1              =   4065
         X2              =   4065
         Y1              =   795
         Y2              =   2820
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Sale Amt"
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
         Height          =   495
         Index           =   6
         Left            =   4125
         TabIndex        =   26
         Top             =   2325
         Width           =   1005
      End
      Begin VB.Label lblcashsale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   5175
         TabIndex        =   25
         Top             =   2325
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Sale Amt"
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
         Height          =   465
         Index           =   5
         Left            =   7095
         TabIndex        =   24
         Top             =   2325
         Width           =   870
      End
      Begin VB.Label lblcrdtsale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   8025
         TabIndex        =   23
         Top             =   2325
         Width           =   1845
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sale"
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
         Left            =   4125
         TabIndex        =   22
         Top             =   855
         Width           =   1365
      End
      Begin VB.Label LBLBTRXTOTAL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   5175
         TabIndex        =   21
         Top             =   810
         Width           =   1845
      End
      Begin VB.Label LBLCOST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   8025
         TabIndex        =   20
         Top             =   810
         Width           =   1845
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "COST"
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
         Index           =   7
         Left            =   7125
         TabIndex        =   19
         Top             =   795
         Width           =   660
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "PROFIT"
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
         Index           =   8
         Left            =   7110
         TabIndex        =   18
         Top             =   1845
         Width           =   810
      End
      Begin VB.Label LBLPROFIT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   8025
         TabIndex        =   17
         Top             =   1830
         Width           =   1845
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
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
         Index           =   9
         Left            =   4125
         TabIndex        =   16
         Top             =   1365
         Width           =   1155
      End
      Begin VB.Label LBLDISCOUNT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   5175
         TabIndex        =   15
         Top             =   1320
         Width           =   1845
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Sale"
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
         Index           =   10
         Left            =   4125
         TabIndex        =   14
         Top             =   1845
         Width           =   1185
      End
      Begin VB.Label LBLNET 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   5175
         TabIndex        =   13
         Top             =   1830
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Purchase"
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
         Left            =   105
         TabIndex        =   12
         Top             =   870
         Width           =   1920
      End
      Begin VB.Label LBLPTRXTOTAL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   2130
         TabIndex        =   11
         Top             =   810
         Width           =   1845
      End
      Begin VB.Label LBLCRDTPUR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   2130
         TabIndex        =   10
         Top             =   1830
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Crdt. Purchase"
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
         Left            =   105
         TabIndex        =   9
         Top             =   1860
         Width           =   2055
      End
      Begin VB.Label LBLCASHPUR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   2130
         TabIndex        =   8
         Top             =   1320
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Purchase"
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
         Left            =   105
         TabIndex        =   7
         Top             =   1380
         Width           =   2055
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
         Left            =   3615
         TabIndex        =   6
         Top             =   255
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmDaybookwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLOSEALL As Integer

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Exit Sub
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Paid To"
    GRDBILL.TextMatrix(0, 2) = "Paid Amt"
    GRDBILL.TextMatrix(0, 3) = "Paid Date"
    GRDBILL.TextMatrix(0, 4) = "Invoice Dtd"
    GRDBILL.TextMatrix(0, 5) = "Invoice No"
    GRDBILL.TextMatrix(0, 6) = "Ref No"
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 2700
    GRDBILL.ColWidth(2) = 1200
    GRDBILL.ColWidth(3) = 1200
    GRDBILL.ColWidth(4) = 1200
    GRDBILL.ColWidth(5) = 1100
    GRDBILL.ColWidth(6) = 1500
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(1) = 1
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 1
    GRDBILL.ColAlignment(5) = 3
    GRDBILL.ColAlignment(6) = 1
    
    grdrcpt.TextMatrix(0, 0) = "SL"
    grdrcpt.TextMatrix(0, 1) = "Received From"
    grdrcpt.TextMatrix(0, 2) = "Receipt Amt"
    grdrcpt.TextMatrix(0, 3) = "Receipt Date"
    grdrcpt.TextMatrix(0, 4) = "Invoice Dtd"
    grdrcpt.TextMatrix(0, 5) = "Invoice No"
    grdrcpt.TextMatrix(0, 6) = "Ref No"
    
    grdrcpt.ColWidth(0) = 500
    grdrcpt.ColWidth(1) = 2700
    grdrcpt.ColWidth(2) = 1200
    grdrcpt.ColWidth(3) = 1200
    grdrcpt.ColWidth(4) = 1200
    grdrcpt.ColWidth(5) = 1100
    grdrcpt.ColWidth(6) = 1500
    
    grdrcpt.ColAlignment(0) = 3
    grdrcpt.ColAlignment(1) = 1
    grdrcpt.ColAlignment(2) = 6
    grdrcpt.ColAlignment(3) = 3
    grdrcpt.ColAlignment(4) = 1
    grdrcpt.ColAlignment(5) = 3
    grdrcpt.ColAlignment(6) = 1
    
    CLOSEALL = 1
    DTFROM.Value = Format(Date, "DD/MM/YYYY")
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 9585
    'Me.Height = 10185
    Me.Left = 1500
    Me.Top = 0
End Sub

Private Sub CMDDISPLAY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    
    LBLBTRXTOTAL.Caption = "0.00"
    LBLDISCOUNT.Caption = "0.00"
    LBLNET.Caption = "0.00"
    LBLCOST.Caption = "0.00"
    LBLPROFIT.Caption = "0.00"
    LBLEXPENSE.Caption = "0.00"
    lblcashsale.Caption = "0.00"
    lblcrdtsale.Caption = "0.00"
    
    LBLPTRXTOTAL.Caption = "0.00"
    LBLCASHPUR.Caption = "0.00"
    LBLCRDTPUR.Caption = "0.00"

    LBLCASHPAY.Caption = "0.00"
    lblcashrcv.Caption = "0.00"
    
    'vbalProgressBar1.Value = 0
    'vbalProgressBar1.ShowText = True
    
    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRANSMASTWO WHERE [VCH_DATE] <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND [VCH_DATE] >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND TRX_TYPE='PI'", db2, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        LBLPTRXTOTAL.Caption = Format(Val(LBLPTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        If (rstTRANX!POST_FLAG = "Y") Then
            LBLCASHPUR.Caption = Format(Val(LBLCASHPUR.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        Else
            LBLCRDTPUR.Caption = Format(Val(LBLCRDTPUR.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        End If
        'vbalProgressBar1.Max = rstTRANX.RecordCount
        'vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXEXPENSE WHERE [VCH_DATE] <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND [VCH_DATE] >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND TRX_TYPE='EX'", db2, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        LBLEXPENSE.Caption = Format(Val(LBLEXPENSE.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMASTWO WHERE [VCH_DATE] <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND [VCH_DATE] >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND (TRX_TYPE='SI' OR TRX_TYPE='RI')", db2, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        LBLBTRXTOTAL.Caption = Format(Val(LBLBTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption) + rstTRANX!DISCOUNT, "0.00")
        LBLNET.Caption = Format(Val(LBLBTRXTOTAL.Caption) - Val(LBLDISCOUNT.Caption), "0.00")
        LBLCOST.Caption = Format(Val(LBLCOST.Caption) + rstTRANX!PAY_AMOUNT, "0.00")
        If (rstTRANX!POST_FLAG = "Y") Then
            lblcashsale.Caption = Format(Val(lblcashsale.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        Else
            lblcrdtsale.Caption = Format(Val(lblcrdtsale.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        End If
        'vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing

    
    LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(LBLEXPENSE.Caption)), "0.00")
    
    'vbalProgressBar1.ShowText = False
    'vbalProgressBar1.Value = 0
    Screen.MousePointer = vbDefault
    Exit Sub
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        'If ACT_FLAG = False Then ACT_REC.Close
    
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

Private Sub DTFROM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DTTO.SetFocus
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
