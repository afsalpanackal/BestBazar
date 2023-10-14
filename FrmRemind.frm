VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frmreminder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYMENT REMINDER"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   21180
   Icon            =   "FrmRemind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   21180
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
      Left            =   6720
      TabIndex        =   11
      Top             =   7860
      Width           =   1080
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save Receipt Entries"
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
      Left            =   14835
      TabIndex        =   18
      Top             =   7860
      Width           =   1590
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
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
      Left            =   7860
      TabIndex        =   12
      Top             =   7860
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1140
      Left            =   5685
      TabIndex        =   8
      Top             =   -60
      Width           =   15495
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Height          =   945
         Left            =   12885
         TabIndex        =   34
         Top             =   135
         Width           =   2550
         Begin VB.OptionButton OptHide 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Hide Day Analysis"
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
            Left            =   45
            TabIndex        =   36
            Top             =   525
            Value           =   -1  'True
            Width           =   2085
         End
         Begin VB.OptionButton OptShow 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Show Day Analysis"
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
            Left            =   30
            TabIndex        =   35
            Top             =   210
            Width           =   2130
         End
      End
      Begin VB.TextBox TxtRef 
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
         Height          =   360
         Left            =   9975
         TabIndex        =   24
         Top             =   540
         Width           =   2760
      End
      Begin VB.TextBox TxtName 
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
         Height          =   360
         Left            =   1380
         TabIndex        =   15
         Top             =   120
         Width           =   2745
      End
      Begin VB.TextBox txtCode 
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
         Height          =   360
         Left            =   45
         TabIndex        =   14
         Top             =   120
         Width           =   1320
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Area"
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
         Left            =   4815
         TabIndex        =   13
         Top             =   135
         Width           =   870
      End
      Begin VB.OptionButton OptAllCategory 
         BackColor       =   &H00C0E0FF&
         Caption         =   "All"
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
         Left            =   4155
         TabIndex        =   10
         Top             =   135
         Value           =   -1  'True
         Width           =   690
      End
      Begin MSComCtl2.DTPicker DTRCPT 
         Height          =   390
         Left            =   10950
         TabIndex        =   16
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   688
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
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         CheckBox        =   -1  'True
         Format          =   116719617
         CurrentDate     =   40498
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Height          =   570
         Left            =   4125
         TabIndex        =   30
         Top             =   480
         Width           =   4545
         Begin VB.OptionButton Optallagnts 
            BackColor       =   &H00C0E0FF&
            Caption         =   "All"
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
            Left            =   30
            TabIndex        =   32
            Top             =   165
            Value           =   -1  'True
            Width           =   690
         End
         Begin VB.OptionButton OptAgent 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Agent"
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
            Left            =   690
            TabIndex        =   31
            Top             =   165
            Width           =   870
         End
         Begin MSDataListLib.DataCombo CMBDISTI 
            Height          =   330
            Left            =   1560
            TabIndex        =   33
            Top             =   165
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   16711680
            Text            =   ""
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
      End
      Begin VB.Label LblInvoice 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F7 for B2C Sales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Index           =   6
         Left            =   75
         TabIndex        =   29
         Top             =   705
         Width           =   3510
      End
      Begin VB.Label LblInvoice 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F2 to enter Receipt Amounts"
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
         Height          =   300
         Index           =   5
         Left            =   75
         TabIndex        =   28
         Top             =   480
         Width           =   3510
      End
      Begin VB.Label LblInvoice 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. No."
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
         Height          =   300
         Index           =   3
         Left            =   9165
         TabIndex        =   25
         Top             =   585
         Width           =   810
      End
      Begin VB.Label LblInvoice 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
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
         Height          =   300
         Index           =   1
         Left            =   9630
         TabIndex        =   17
         Top             =   180
         Width           =   1275
      End
      Begin MSForms.ComboBox cmbarea 
         Height          =   360
         Left            =   5685
         TabIndex        =   9
         Top             =   150
         Width           =   2955
         VariousPropertyBits=   746604571
         ForeColor       =   255
         DisplayStyle    =   7
         Size            =   "5212;635"
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1155
      Left            =   -30
      TabIndex        =   2
      Top             =   -75
      Width           =   5715
      Begin VB.OptionButton OptCrPeriod 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Over Credit Period"
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
         Left            =   3585
         TabIndex        =   5
         Top             =   195
         Width           =   2070
      End
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "All Customers"
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
         TabIndex        =   4
         Top             =   180
         Width           =   1605
      End
      Begin VB.OptionButton OptBAL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Oustanding Only"
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
         Left            =   1710
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1875
      End
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
      Left            =   10110
      TabIndex        =   1
      Top             =   7860
      Width           =   1080
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
      Left            =   8985
      TabIndex        =   0
      Top             =   7860
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Height          =   6840
      Left            =   0
      TabIndex        =   21
      Top             =   990
      Width           =   21195
      Begin VB.TextBox TXTsample 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   7470
         TabIndex        =   23
         Top             =   870
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6675
         Left            =   15
         TabIndex        =   22
         Top             =   120
         Width           =   21150
         _ExtentX        =   37306
         _ExtentY        =   11774
         _Version        =   393216
         Rows            =   1
         Cols            =   16
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label LblLastRcpt 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   4755
      TabIndex        =   27
      Top             =   7860
      Width           =   1920
   End
   Begin VB.Label LblInvoice 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Rcpt Amt"
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
      Left            =   3330
      TabIndex        =   26
      Top             =   7950
      Width           =   1410
   End
   Begin VB.Label LblInvoice 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Amount"
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
      Left            =   11220
      TabIndex        =   20
      Top             =   7950
      Width           =   1590
   End
   Begin VB.Label LblReceipt 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   12825
      TabIndex        =   19
      Top             =   7875
      Width           =   1965
   End
   Begin VB.Label LblInvoice 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Top             =   7935
      Width           =   1410
   End
   Begin VB.Label lblAMT 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   1425
      TabIndex        =   6
      Top             =   7845
      Width           =   1875
   End
End
Attribute VB_Name = "Frmreminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim M_EDIT1, M_EDIT2 As Boolean
Dim ACT_AGNT As New ADODB.Recordset
Dim AGNT_FLAG As Boolean

Private Function Fillgrid()
    Dim rstTRANX, rstCust As ADODB.Recordset
    Dim OpBal, AC_DB, AC_CR, Total_DB, Total_CR As Double
    Dim DueDays As String
    Dim CR_PERIOD, DUE_DATE, Last_Rcpt_Amt As Long
    Dim i As Long
    
    If optCategory.Value = True And cmbarea.Text = "" Then
        MsgBox "Please select the Place from the List", vbOKOnly, "Receipt Dues"
        On Error Resume Next
        cmbarea.SetFocus
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    If OptShow.Value = True Then
        GRDTranx.ColWidth(0) = 700
        GRDTranx.ColWidth(1) = 900
        GRDTranx.ColWidth(2) = 3600
        GRDTranx.ColWidth(3) = 1300
        GRDTranx.ColWidth(4) = 1300
        GRDTranx.ColWidth(5) = 1000
        GRDTranx.ColWidth(6) = 1000
        GRDTranx.ColWidth(7) = 1000
        GRDTranx.ColWidth(8) = 1000
        GRDTranx.ColWidth(9) = 1000
        GRDTranx.ColWidth(10) = 1000
        GRDTranx.ColWidth(11) = 1300
        GRDTranx.ColWidth(12) = 1100
        GRDTranx.ColWidth(13) = 1000
        GRDTranx.ColWidth(14) = 1200
        GRDTranx.ColWidth(15) = 2600
    Else
        GRDTranx.ColWidth(0) = 700
        GRDTranx.ColWidth(1) = 900
        GRDTranx.ColWidth(2) = 3600
        GRDTranx.ColWidth(3) = 1300
        GRDTranx.ColWidth(4) = 1300
        GRDTranx.ColWidth(5) = 1000
        GRDTranx.ColWidth(6) = 0
        GRDTranx.ColWidth(7) = 0
        GRDTranx.ColWidth(8) = 0
        GRDTranx.ColWidth(9) = 0
        GRDTranx.ColWidth(10) = 0
        GRDTranx.ColWidth(11) = 1300
        GRDTranx.ColWidth(12) = 1100
        GRDTranx.ColWidth(13) = 1000
        GRDTranx.ColWidth(14) = 1200
        GRDTranx.ColWidth(15) = 2600
    End If
    GRDTranx.rows = 1
    i = 1
    lblAMT.Caption = ""
    LblLastRcpt.Caption = ""
    LblReceipt.Caption = ""
    On Error GoTo eRRhAND
    
    Dim OP_Sale As Double
    Dim OP_Rcpt As Double
    Dim dtcrdays As Date
    Dim dtdrdays As Date
    
    db.Execute "Update CUSTMAST set YTD_CR = 0 "
    Set rstCust = New ADODB.Recordset
    If optCategory.Value = True Then
        If Optallagnts.Value = True Then
            rstCust.Open "SELECT * From CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND AREA Like '%" & cmbarea.Text & "%' AND ACT_CODE Like '%" & Trim(txtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
        Else
            rstCust.Open "SELECT * From CUSTMAST WHERE AGENT_CODE = '" & CMBDISTI.BoundText & "' AND ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND AREA Like '%" & cmbarea.Text & "%' AND ACT_CODE Like '%" & Trim(txtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
        End If
    Else
        If Optallagnts.Value = True Then
            rstCust.Open "SELECT * From CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND ACT_CODE Like '%" & Trim(txtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
        Else
            rstCust.Open "SELECT * From CUSTMAST WHERE AGENT_CODE = '" & CMBDISTI.BoundText & "' AND ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND ACT_CODE Like '%" & Trim(txtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
        End If
    End If
    Do Until rstCust.EOF
        OpBal = 0
        Total_DB = 0
        Total_CR = 0
        DUE_DATE = 0
        DueDays = ""
        Last_Rcpt_Amt = 0
        CR_PERIOD = IIf(IsNull(rstCust!PYMT_PERIOD), 0, rstCust!PYMT_PERIOD)
        OpBal = IIf(IsNull(rstCust!OPEN_DB), 0, rstCust!OPEN_DB)
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & rstCust!ACT_CODE & "' AND (TRX_TYPE = 'CB' OR TRX_TYPE = 'DB' OR TRX_TYPE = 'RT' OR TRX_TYPE = 'DR' OR TRX_TYPE = 'RW' OR TRX_TYPE = 'SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'RD') ORDER BY CR_NO ASC, INV_DATE DESC", db, adOpenForwardOnly
        Do Until rstTRANX.EOF
            AC_DB = 0
            AC_CR = 0
            AC_DB = IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
            Select Case rstTRANX!check_flag
                Case "Y"
                    AC_CR = IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
                Case "N"
                    AC_CR = 0 '""IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
            End Select
            Select Case rstTRANX!TRX_TYPE
                Case "DR"
                    If IsDate(rstTRANX!INV_DATE) Then
                        DUE_DATE = DateDiff("d", rstTRANX!INV_DATE, Date)
                        DueDays = DUE_DATE & " days"
                    End If
                Case "DB"
                    AC_DB = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
                Case "RT"
                    AC_CR = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
                    Last_Rcpt_Amt = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
                Case "CB", "SR", "EP", "VC", "ER", "PY", "RW"
                    AC_CR = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
            End Select
            
            Total_DB = Total_DB + AC_DB
            Total_CR = Total_CR + AC_CR
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
    
        If OptAll.Value = False And Round((OpBal + Total_DB) - Total_CR, 2) = 0 Then GoTo SKIP
        If OptCrPeriod.Value = True And DUE_DATE < CR_PERIOD Then GoTo SKIP
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(i, 0) = i
        GRDTranx.TextMatrix(i, 1) = IIf(IsNull(rstCust!ACT_CODE), "", rstCust!ACT_CODE)
        'GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstCust!act_name), "", rstCust!act_name) & IIf(IsNull(rstCust!Address) Or rstCust!Address = "", "", "," & Mid(rstCust!Address, 1, 30)) & " - " & IIf(IsNull(rstCust!TELNO) Or rstCust!TELNO = "", "", rstCust!TELNO & ", ") & IIf(IsNull(rstCust!FAXNO), "", rstCust!FAXNO)
        GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstCust!ACT_NAME), "", rstCust!ACT_NAME) '& " - " & IIf(IsNull(rstCust!TELNO) Or rstCust!TELNO = "", "", rstCust!TELNO & ", ") & IIf(IsNull(rstCust!FAXNO), "", rstCust!FAXNO)
        GRDTranx.TextMatrix(i, 3) = OpBal + Total_DB
        GRDTranx.TextMatrix(i, 4) = Total_CR
        GRDTranx.TextMatrix(i, 5) = Round((OpBal + Total_DB) - Total_CR, 2)
        rstCust!YTD_CR = Val(GRDTranx.TextMatrix(i, 5))
        rstCust.Update
            
        If OptHide.Value = True Then GoTo SKKIP_DAY
        '>7days
        OP_Sale = OpBal
        OP_Rcpt = 0
        dtcrdays = DateDiff("d", 7, Format(Date, "DD/MM/YYYY"))
        dtdrdays = DateAdd("d", 15, Format(Date, "DD/MM/YYYY"))
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(INV_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='DR' OR TRX_TYPE = 'RD') and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' and INV_DATE < '" & Format(dtdrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and TRX_TYPE ='DB'  and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' and INV_DATE < '" & Format(dtdrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE = 'RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        GRDTranx.TextMatrix(i, 6) = Format(Round(OP_Sale - OP_Rcpt, 2), "0.00")
        If Val(GRDTranx.TextMatrix(i, 6)) < 0 Then GRDTranx.TextMatrix(i, 6) = 0
        
        '>15days
        OP_Sale = OpBal
        OP_Rcpt = 0
        dtcrdays = DateDiff("d", 15, Format(Date, "DD/MM/YYYY"))
        dtdrdays = DateAdd("d", 30, Format(Date, "DD/MM/YYYY"))
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(INV_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='DR' OR TRX_TYPE = 'RD') and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' and INV_DATE < '" & Format(dtdrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and TRX_TYPE ='DB'  and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' and INV_DATE < '" & Format(dtdrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE = 'RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' ) ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        GRDTranx.TextMatrix(i, 7) = Format(Round(OP_Sale - OP_Rcpt, 2), "0.00")
        If Val(GRDTranx.TextMatrix(i, 7)) < 0 Then GRDTranx.TextMatrix(i, 7) = 0
        
        '>30days
        OP_Sale = OpBal
        OP_Rcpt = 0
        dtcrdays = DateDiff("d", 30, Format(Date, "DD/MM/YYYY"))
        dtdrdays = DateAdd("d", 45, Format(Date, "DD/MM/YYYY"))
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(INV_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='DR' OR TRX_TYPE = 'RD') and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' and INV_DATE < '" & Format(dtdrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and TRX_TYPE ='DB'  and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' and INV_DATE < '" & Format(dtdrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE = 'RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        GRDTranx.TextMatrix(i, 8) = Format(Round(OP_Sale - OP_Rcpt, 2), "0.00")
        If Val(GRDTranx.TextMatrix(i, 8)) < 0 Then GRDTranx.TextMatrix(i, 8) = 0
        
        '>45days
        OP_Sale = OpBal
        OP_Rcpt = 0
        dtcrdays = DateDiff("d", 45, Format(Date, "DD/MM/YYYY"))
        dtdrdays = DateAdd("d", 60, Format(Date, "DD/MM/YYYY"))
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(INV_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='DR' OR TRX_TYPE = 'RD') and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' and INV_DATE < '" & Format(dtdrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and TRX_TYPE ='DB'  and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' and INV_DATE < '" & Format(dtdrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE = 'RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        GRDTranx.TextMatrix(i, 9) = Format(Round(OP_Sale - OP_Rcpt, 2), "0.00")
        If Val(GRDTranx.TextMatrix(i, 9)) < 0 Then GRDTranx.TextMatrix(i, 9) = 0
        
        '>60days
        OP_Sale = OpBal
        OP_Rcpt = 0
        dtcrdays = DateDiff("d", 60, Format(Date, "DD/MM/YYYY"))
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(INV_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='DR' OR TRX_TYPE = 'RD') and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and TRX_TYPE ='DB'  and INV_DATE <= '" & Format(dtcrdays, "yyyy/mm/dd") & "' and INV_DATE < '" & Format(dtdrdays, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCust!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE = 'RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        GRDTranx.TextMatrix(i, 10) = Format(Round(OP_Sale - OP_Rcpt, 2), "0.00")
        If Val(GRDTranx.TextMatrix(i, 10)) < 0 Then GRDTranx.TextMatrix(i, 10) = 0
SKKIP_DAY:
        GRDTranx.TextMatrix(i, 11) = DueDays
        If Last_Rcpt_Amt > 0 Then GRDTranx.TextMatrix(i, 12) = Last_Rcpt_Amt
        GRDTranx.TextMatrix(i, 13) = ""
        GRDTranx.TextMatrix(i, 14) = IIf(IsNull(rstCust!OPEN_DB), "0.00", Format(rstCust!OPEN_DB, "0.00"))
        GRDTranx.TextMatrix(i, 15) = IIf(IsNull(rstCust!Area), "", rstCust!Area) & " - " & IIf(IsNull(rstCust!TELNO) Or rstCust!TELNO = "", "", rstCust!TELNO & ", ") & IIf(IsNull(rstCust!FAXNO), "", rstCust!FAXNO)
        lblAMT.Caption = Format(Val(lblAMT.Caption) + GRDTranx.TextMatrix(i, 5), "0.00")
        LblLastRcpt.Caption = Val(LblLastRcpt.Caption) + Val(GRDTranx.TextMatrix(i, 12))
        i = i + 1
SKIP:
        rstCust.MoveNext
    Loop
    rstCust.Close
    Set rstCust = Nothing
    
    LblLastRcpt.Caption = Format(Round(Val(LblLastRcpt.Caption), 2), "0.00")
    
    DTRCPT.Value = Null
    M_EDIT1 = False
    M_EDIT2 = False
    TxtRef.Text = ""
    On Error Resume Next
    GRDTranx.SetFocus
    CmdPrint.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Function
    
eRRhAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Function

Private Sub cmbarea_GotFocus()
    optCategory.Value = True
    CmdPrint.Enabled = False
End Sub

Private Sub CmDDisplay_Click()
    If M_EDIT1 = True And M_EDIT2 = True Then
        If MsgBox("Changes have been made. Do you want to save the changes?", vbYesNo, "Receipt Entries...") = vbNo Then
            Call Fillgrid
            Exit Sub
        Else
            Call CmdSave_Click
            Exit Sub
        End If
    End If
    Call Fillgrid
End Sub

Private Sub CmdExit_Click()
    If M_EDIT1 = True And M_EDIT2 = True Then
        If MsgBox("Changes have been made. Do you want to save the changes?", vbYesNo, "Receipt Entries...") = vbYes Then
            Call CmdSave_Click
            If Not (M_EDIT1 = True And M_EDIT2 = True) Then
                Unload Me
            End If
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub CmdExport_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "EXPORT"
    If GRDTranx.rows <= 1 Then Exit Sub
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Debtor's Statement") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, n As Long
    
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
        oWS.Range("A1", "H1").Merge
        oWS.Range("A1", "H1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "H2").Merge
        oWS.Range("A2", "H2").HorizontalAlignment = xlCenter
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
    oWS.Range("A" & 1).Value = MDIMAIN.StatusBar.Panels(5).Text
    oWS.Range("A" & 2).Value = "DEBTOR'S STATEMENT"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).Value = GRDTranx.TextMatrix(0, 0)
    oWS.Range("B" & 3).Value = GRDTranx.TextMatrix(0, 1)
    oWS.Range("C" & 3).Value = GRDTranx.TextMatrix(0, 2)
    oWS.Range("D" & 3).Value = GRDTranx.TextMatrix(0, 3)
    oWS.Range("E" & 3).Value = GRDTranx.TextMatrix(0, 4)
    oWS.Range("F" & 3).Value = GRDTranx.TextMatrix(0, 5)
    If OptShow.Value = True Then
        oWS.Range("G" & 3).Value = GRDTranx.TextMatrix(0, 6)
        oWS.Range("H" & 3).Value = GRDTranx.TextMatrix(0, 7)
        oWS.Range("I" & 3).Value = GRDTranx.TextMatrix(0, 8)
        oWS.Range("J" & 3).Value = GRDTranx.TextMatrix(0, 9)
        oWS.Range("K" & 3).Value = GRDTranx.TextMatrix(0, 10)
        oWS.Range("L" & 3).Value = GRDTranx.TextMatrix(0, 11)
        oWS.Range("M" & 3).Value = GRDTranx.TextMatrix(0, 12)
        oWS.Range("N" & 3).Value = GRDTranx.TextMatrix(0, 13)
        oWS.Range("O" & 3).Value = GRDTranx.TextMatrix(0, 14)
        oWS.Range("P" & 3).Value = GRDTranx.TextMatrix(0, 15)
    Else
        oWS.Range("G" & 3).Value = GRDTranx.TextMatrix(0, 11)
        oWS.Range("H" & 3).Value = GRDTranx.TextMatrix(0, 12)
        oWS.Range("I" & 3).Value = GRDTranx.TextMatrix(0, 13)
        oWS.Range("J" & 3).Value = GRDTranx.TextMatrix(0, 14)
        oWS.Range("K" & 3).Value = GRDTranx.TextMatrix(0, 15)
    End If
    On Error GoTo eRRhAND
    
    i = 4
    For n = 1 To GRDTranx.rows - 1
        oWS.Range("A" & i).Value = GRDTranx.TextMatrix(n, 0)
        oWS.Range("B" & i).Value = GRDTranx.TextMatrix(n, 1)
        oWS.Range("C" & i).Value = GRDTranx.TextMatrix(n, 2)
        oWS.Range("D" & i).Value = GRDTranx.TextMatrix(n, 3)
        oWS.Range("E" & i).Value = GRDTranx.TextMatrix(n, 4)
        oWS.Range("F" & i).Value = GRDTranx.TextMatrix(n, 5)
        If OptShow.Value = True Then
            oWS.Range("G" & i).Value = GRDTranx.TextMatrix(n, 6)
            oWS.Range("H" & i).Value = GRDTranx.TextMatrix(n, 7)
            oWS.Range("I" & i).Value = GRDTranx.TextMatrix(n, 8)
            oWS.Range("J" & i).Value = GRDTranx.TextMatrix(n, 9)
            oWS.Range("K" & i).Value = GRDTranx.TextMatrix(n, 10)
            oWS.Range("L" & i).Value = GRDTranx.TextMatrix(n, 11)
            oWS.Range("M" & i).Value = GRDTranx.TextMatrix(n, 12)
            oWS.Range("N" & i).Value = GRDTranx.TextMatrix(n, 13)
            oWS.Range("O" & i).Value = GRDTranx.TextMatrix(n, 14)
            oWS.Range("P" & i).Value = GRDTranx.TextMatrix(n, 15)
        Else
            oWS.Range("G" & i).Value = GRDTranx.TextMatrix(n, 11)
            oWS.Range("H" & i).Value = GRDTranx.TextMatrix(n, 12)
            oWS.Range("I" & i).Value = GRDTranx.TextMatrix(n, 13)
            oWS.Range("J" & i).Value = GRDTranx.TextMatrix(n, 14)
            oWS.Range("K" & i).Value = GRDTranx.TextMatrix(n, 15)
        End If
        On Error GoTo eRRhAND
        i = i + 1
    Next n
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    oWS.Columns("A:Z").EntireColumn.AutoFit
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
    MsgBox err.Description
End Sub

Private Sub CmdPrint_Click()
    Dim i As Long
    
    On Error GoTo eRRhAND
    ReportNameVar = Rptpath & "RptRecSt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If optCategory.Value = True Then
        If Optallagnts.Value = True Then
            Report.RecordSelectionFormula = "(({CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.AREA} startswith '" & cmbarea.Text & "' AND {CUSTMAST.ACT_CODE} startswith '" & Trim(txtCode.Text) & "' AND {CUSTMAST.ACT_NAME} startswith '" & Trim(TxtName.Text) & "' AND {CUSTMAST.YTD_CR} <>0))"
        Else
            Report.RecordSelectionFormula = "(({CUSTMAST.AGENT_CODE} = '" & CMBDISTI.BoundText & "' AND {CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.AREA} startswith '" & cmbarea.Text & "' AND {CUSTMAST.ACT_CODE} startswith '" & Trim(txtCode.Text) & "' AND {CUSTMAST.ACT_NAME} startswith '" & Trim(TxtName.Text) & "' AND {CUSTMAST.YTD_CR} <>0))"
        End If
    Else
        If Optallagnts.Value = True Then
            Report.RecordSelectionFormula = "(({CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.ACT_CODE} startswith '" & Trim(txtCode.Text) & "' AND {CUSTMAST.ACT_NAME} startswith '" & Trim(TxtName.Text) & "' AND {CUSTMAST.YTD_CR} <>0))"
        Else
            Report.RecordSelectionFormula = "(({CUSTMAST.AGENT_CODE} = '" & CMBDISTI.BoundText & "' AND {CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.ACT_CODE} startswith '" & Trim(txtCode.Text) & "' AND {CUSTMAST.ACT_NAME} startswith '" & Trim(TxtName.Text) & "' AND {CUSTMAST.YTD_CR} <>0))"
        End If
    End If
    'Report.RecordSelectionFormula = "(({CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.YTD_CR} <>0))"
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'CUSTOMER DETAILS'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdSave_Click()
    
    Dim RSTTRXFILE, rstBILL As ADODB.Recordset
    Dim Sl_no, Max_Rec, Rec_Nos As Long
    
    If IsNull(DTRCPT.Value) Then
        MsgBox "Please select Date of Receipt", vbOKOnly, "Receipt"
        DTRCPT.SetFocus
        Exit Sub
    End If
    
'    If Trim(TxtRef.Text) = "" Then
'        MsgBox "Please enter the Reference No.", vbOKOnly, "Receipt"
'        TxtRef.SetFocus
'        Exit Sub
'    End If
    
    If MsgBox("ARE YOU SURE YOU WANT TO SAVE ALL THE RECEIPT ENTRIES", vbYesNo, "RECEIPT.....") = vbNo Then Exit Sub
    On Error GoTo eRRhAND
    Dim i As Long
    Dim RECNO, INVNO As Long
    Dim BillNO As Long
    Dim TRXTYPE, INVTRXTYPE, INVTYPE As String
    Dim rstMaxRec As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    Rec_Nos = 0
    LblReceipt.Caption = ""
    db.BeginTrans
    For Sl_no = 1 To GRDTranx.rows - 1
        If Val(GRDTranx.TextMatrix(Sl_no, 13)) = 0 Then GoTo SKIP
        Rec_Nos = Rec_Nos + 1
        Set rstBILL = New ADODB.Recordset
        rstBILL.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
        If Not (rstBILL.EOF And rstBILL.BOF) Then
            Max_Rec = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        End If
        rstBILL.Close
        Set rstBILL = Nothing
        
        Dim MAXRCPTNO As Long
        MAXRCPTNO = 1
        Set rstBILL = New ADODB.Recordset
        rstBILL.Open "Select MAX(REC_NO) From DBTPYMT WHERE TRX_TYPE = 'RT' AND '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenForwardOnly
        If Not (rstBILL.EOF And rstBILL.BOF) Then
            MAXRCPTNO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        End If
        rstBILL.Close
        Set rstBILL = Nothing
    
        i = 0
        Set rstMaxRec = New ADODB.Recordset
        rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
            i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
        End If
        rstMaxRec.Close
        Set rstMaxRec = Nothing
        
        'db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & Val(creditbill.LBLBILLNO.Caption) & " AND INV_TYPE = 'RT'"
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTITEMMAST.AddNew
        RSTITEMMAST!REC_NO = i + 1
        RSTITEMMAST!INV_TYPE = "RT"
        RSTITEMMAST!INV_TRX_TYPE = "RT"
        RSTITEMMAST!INV_NO = Max_Rec
        RSTITEMMAST!TRX_TYPE = "CR"
        RSTITEMMAST!ACT_CODE = GRDTranx.TextMatrix(Sl_no, 1)
        RSTITEMMAST!ACT_NAME = GRDTranx.TextMatrix(Sl_no, 2)
        RSTITEMMAST!AMOUNT = Val(GRDTranx.TextMatrix(Sl_no, 13))
        RSTITEMMAST!VCH_DATE = Format(DTRCPT.Value, "DD/MM/YYYY")
        RSTITEMMAST!BILL_TRX_TYPE = "SI"
        RSTITEMMAST!CASH_MODE = "C"
        RSTITEMMAST!CHQ_NO = ""
        'RSTITEMMAST!CHQ_DATE = Null
        RSTITEMMAST!BANK = ""
        RSTITEMMAST!CHQ_STATUS = ""
        RSTITEMMAST!check_flag = "S"
        RECNO = RSTITEMMAST!REC_NO
        INVNO = RSTITEMMAST!INV_NO
        TRXTYPE = RSTITEMMAST!TRX_TYPE
        INVTRXTYPE = RSTITEMMAST!INV_TRX_TYPE
        INVTYPE = RSTITEMMAST!INV_TYPE
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "RT"
        RSTTRXFILE!INV_TRX_TYPE = "RI"
        RSTTRXFILE!CR_NO = Max_Rec
        RSTTRXFILE!REC_NO = MAXRCPTNO
        RSTTRXFILE!RCPT_DATE = Format(DTRCPT.Value, "DD/MM/YYYY")
        RSTTRXFILE!RCPT_AMT = Val(GRDTranx.TextMatrix(Sl_no, 13))
        RSTTRXFILE!ACT_CODE = GRDTranx.TextMatrix(Sl_no, 1)
        RSTTRXFILE!ACT_NAME = GRDTranx.TextMatrix(Sl_no, 2)
        RSTTRXFILE!INV_DATE = Format(DTRCPT.Value, "DD/MM/YYYY")
        RSTTRXFILE!REF_NO = Trim(TxtRef.Text)
        RSTTRXFILE!INV_AMT = Null
        'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
        RSTTRXFILE!INV_NO = 0
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!BANK_FLAG = "N"
        RSTTRXFILE!B_TRX_TYPE = Null
        'RSTTRXFILE!B_TRX_NO = Null
        RSTTRXFILE!B_BILL_TRX_TYPE = Null
        RSTTRXFILE!B_TRX_YEAR = Null
        RSTTRXFILE!BANK_CODE = Null
        RSTTRXFILE!C_TRX_TYPE = TRXTYPE
        RSTTRXFILE!C_REC_NO = RECNO
        RSTTRXFILE!C_INV_TRX_TYPE = INVTRXTYPE
        RSTTRXFILE!C_INV_TYPE = INVTYPE
        RSTTRXFILE!C_INV_NO = INVNO
        
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        LblReceipt.Caption = Val(LblReceipt.Caption) + Val(GRDTranx.TextMatrix(Sl_no, 13))
SKIP:
    Next Sl_no
    db.CommitTrans
    LblReceipt.Caption = Format(LblReceipt.Caption, "0.00")
    Screen.MousePointer = vbNormal
    MsgBox Rec_Nos & " Entries Saved", vbOKOnly, "Receipt Entry"
    M_EDIT1 = False
    M_EDIT2 = False
    CmDDisplay_Click
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub Form_Load()
    
    AGNT_FLAG = True
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "CODE"
    GRDTranx.TextMatrix(0, 2) = "NAME"
    GRDTranx.TextMatrix(0, 3) = "CREDIT"
    GRDTranx.TextMatrix(0, 4) = "DEBIT"
    GRDTranx.TextMatrix(0, 5) = "Clo. Amount"
    GRDTranx.TextMatrix(0, 6) = ">=7 days"
    GRDTranx.TextMatrix(0, 7) = ">=15 days"
    GRDTranx.TextMatrix(0, 8) = ">=30 days"
    GRDTranx.TextMatrix(0, 9) = ">=45 days"
    GRDTranx.TextMatrix(0, 10) = ">60 days"
    GRDTranx.TextMatrix(0, 11) = "Last Bill"
    GRDTranx.TextMatrix(0, 12) = "Last Rcpt"
    GRDTranx.TextMatrix(0, 13) = "Rcpt Amt"
    GRDTranx.TextMatrix(0, 14) = "Op. Bal"
    GRDTranx.TextMatrix(0, 15) = "Area - Phone"
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    GRDTranx.ColAlignment(7) = 4
    GRDTranx.ColAlignment(8) = 4
    GRDTranx.ColAlignment(9) = 4
    GRDTranx.ColAlignment(10) = 4
    GRDTranx.ColAlignment(11) = 4
    GRDTranx.ColAlignment(12) = 4
    GRDTranx.ColAlignment(13) = 4
    GRDTranx.ColAlignment(14) = 4
    GRDTranx.ColAlignment(15) = 1
    
    Dim RSTCOMPANY As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "Select DISTINCT AREA From CUSTMAST ORDER BY AREA", db, adOpenStatic, adLockReadOnly
    Do Until RSTCOMPANY.EOF
        If Not IsNull(RSTCOMPANY!Area) Then cmbarea.AddItem (RSTCOMPANY!Area)
        RSTCOMPANY.MoveNext
    Loop
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If AGNT_FLAG = True Then
        ACT_AGNT.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='911')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    Else
        ACT_AGNT.Close
        ACT_AGNT.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='911')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    End If
    
    Set Me.CMBDISTI.RowSource = ACT_AGNT
    CMBDISTI.ListField = "ACT_NAME"
    CMBDISTI.BoundColumn = "ACT_CODE"
    Screen.MousePointer = vbNormal
    
    
    DTRCPT.Value = Format(Date, "DD/MM/YYYY")
    DTRCPT.Value = Null
    CmdPrint.Enabled = False
    Call Fillgrid
    Me.Left = 0
    Me.Top = 0
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AGNT_FLAG = False Then ACT_AGNT.Close
End Sub

Private Sub GRDTranx_Click()
    TXTsample.Visible = False
    GRDTranx.SetFocus
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDTranx.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            For Each F In Forms
                If F.Name = "FRMRECEIPT" Then
                    MsgBox "Please close the Receipt Window", vbOKOnly, "Receipt Entry"
                    GRDTranx.SetFocus
                    Exit Sub
                End If
                If F.Name = "FRMDRCR" Then
                    MsgBox "Please close the Receipt Window", vbOKOnly, "Receipt Entry"
                    GRDTranx.SetFocus
                    Exit Sub
                End If
            Next F
            FRMRcptReg.TXTDEALER.Text = GRDTranx.TextMatrix(GRDTranx.Row, 2)
            FRMRcptReg.DataList2.BoundText = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            FRMRcptReg.Show
            FRMRcptReg.SetFocus
            
        Case 118
            For Each F In Forms
                If F.Name = "FRMGSTR2" Then
                    MsgBox "Sales Window Already Opened", vbOKOnly, "Sales"
                    GRDTranx.SetFocus
                    Exit Sub
                End If
            Next F
'            FRMRcptReg.TXTDEALER.Text = GRDTranx.TextMatrix(GRDTranx.Row, 2)
'            FRMRcptReg.DataList2.BoundText = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            FRMGSTR2.Show
            FRMGSTR2.SetFocus
            FRMGSTR2.TXTDEALER.Text = GRDTranx.TextMatrix(GRDTranx.Row, 2)
            FRMGSTR2.DataList2.BoundText = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            FRMGSTR2.DataList2_Click
            FRMGSTR2.DataList2.BoundText = GRDTranx.TextMatrix(GRDTranx.Row, 1)
        Case 113
            If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                Select Case GRDTranx.Col
                    Case 13, 14, 15
                        TXTsample.Visible = True
                        TXTsample.Top = GRDTranx.CellTop + 90
                        TXTsample.Left = GRDTranx.CellLeft
                        TXTsample.Width = GRDTranx.CellWidth
                        TXTsample.Height = GRDTranx.CellHeight
                        TXTsample.Text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                        TXTsample.SetFocus
                End Select
            End If
    End Select
End Sub

Private Sub GRDTranx_Scroll()
    TXTsample.Visible = False
    GRDTranx.SetFocus
End Sub

Private Sub OptAgent_Click()
    CmdPrint.Enabled = False
End Sub

Private Sub OptAll_Click()
    CmdPrint.Enabled = False
End Sub

Private Sub Optallagnts_Click()
    CmdPrint.Enabled = False
End Sub

Private Sub OptAllCategory_Click()
    CmdPrint.Enabled = False
End Sub

Private Sub OptBAL_Click()
    CmdPrint.Enabled = False
End Sub

Private Sub optCategory_Click()
    CmdPrint.Enabled = False
End Sub

Private Sub OptCrPeriod_Click()
    CmdPrint.Enabled = False
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmDDisplay_Click
            txtCode.SetFocus
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

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmDDisplay_Click
            TxtName.SetFocus
    End Select
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTsample_Change()
    M_EDIT1 = True
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Sl_no As Long
    Dim rststock As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 13  ' Rcpt
                    If Val(TXTsample.Text) > Val(GRDTranx.TextMatrix(GRDTranx.Row, 5)) Then
                        MsgBox "Receipt Amount could not be greater than Balance Amount"
                        Exit Sub
                    End If
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Val(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    LblReceipt.Caption = ""
                    For Sl_no = 1 To GRDTranx.rows - 1
                        LblReceipt.Caption = Val(LblReceipt.Caption) + Val(GRDTranx.TextMatrix(Sl_no, 13))
                    Next Sl_no
                    LblReceipt.Caption = Format(LblReceipt.Caption, "0.00")
                    GRDTranx.SetFocus
                    M_EDIT2 = True
                Case 14  ' OP BAL
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Val(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from CUSTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!OPEN_DB = Format(Val(TXTsample.Text), "0.00")
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    GRDTranx.SetFocus
                Case 15  ' Area
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from CUSTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!Area = Trim(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    GRDTranx.SetFocus
                    
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDTranx.Col
        Case 13, 14
             Select Case KeyAscii
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 15
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub CMBDISTI_GotFocus()
    OptAgent.Value = True
    CmdPrint.Enabled = False
End Sub

