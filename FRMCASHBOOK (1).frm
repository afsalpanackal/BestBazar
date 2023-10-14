VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FRMCASHBOOK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CASH BOOK"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18660
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   18660
   Begin VB.Frame FRMEBILL 
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H00000080&
      Height          =   4725
      Left            =   60
      TabIndex        =   4
      Top             =   1950
      Visible         =   0   'False
      Width           =   10845
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4005
         Left            =   30
         TabIndex        =   5
         Top             =   540
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   7064
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         Appearance      =   0
         GridLineWidth   =   2
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
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "NET AMT"
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
         Index           =   6
         Left            =   8565
         TabIndex        =   13
         Top             =   210
         Width           =   825
      End
      Begin VB.Label LBLNETAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9390
         TabIndex        =   12
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "DISC"
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
         Index           =   2
         Left            =   7320
         TabIndex        =   11
         Top             =   210
         Width           =   495
      End
      Begin VB.Label LBLDISC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7785
         TabIndex        =   10
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LBLBILLAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6150
         TabIndex        =   9
         Top             =   180
         Width           =   1080
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
         Left            =   5190
         TabIndex        =   8
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
         Left            =   3300
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
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      Caption         =   "Frame1"
      Height          =   9885
      Left            =   -120
      TabIndex        =   0
      Top             =   -285
      Width           =   18720
      Begin VB.CommandButton cmdwoprint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   135
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   27
         Top             =   9450
         Width           =   420
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0C0FF&
         Caption         =   "CASH BOOK"
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
         Height          =   1890
         Left            =   150
         TabIndex        =   16
         Top             =   285
         Width           =   18510
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0C0FF&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   75
            TabIndex        =   19
            Top             =   420
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton OPTCUSTOMER 
            BackColor       =   &H00C0C0FF&
            Caption         =   "TYPE"
            Height          =   210
            Left            =   90
            TabIndex        =   18
            Top             =   870
            Visible         =   0   'False
            Width           =   1320
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
            Left            =   1845
            TabIndex        =   17
            Top             =   825
            Visible         =   0   'False
            Width           =   3720
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1860
            TabIndex        =   20
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
            TabIndex        =   21
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   123076609
            CurrentDate     =   40498
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1845
            TabIndex        =   22
            Top             =   1170
            Visible         =   0   'False
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   1138
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
         Begin MSForms.CheckBox ChkDelete 
            Height          =   270
            Left            =   17130
            TabIndex        =   36
            Top             =   150
            Width           =   1350
            BackColor       =   12632319
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2381;476"
            Value           =   "0"
            Caption         =   "Force Delete"
            FontName        =   "Tahoma"
            FontHeight      =   135
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label LBLEXPENSES 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Paid Cash"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   5
            Left            =   7620
            TabIndex        =   35
            Top             =   1155
            Width           =   1590
         End
         Begin VB.Label lblpaidcash 
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
            ForeColor       =   &H0000C000&
            Height          =   405
            Left            =   7575
            TabIndex        =   34
            Top             =   1395
            Width           =   1635
         End
         Begin VB.Label lblcloscash 
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
            ForeColor       =   &H00800080&
            Height          =   405
            Left            =   10965
            TabIndex        =   33
            Top             =   1395
            Width           =   1755
         End
         Begin VB.Label LBLEXPENSES 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Closing Cash"
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
            Height          =   315
            Index           =   4
            Left            =   11010
            TabIndex        =   32
            Top             =   1170
            Width           =   1695
         End
         Begin VB.Label LBLEXPENSES 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rcvd cash"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   3
            Left            =   9285
            TabIndex        =   31
            Top             =   1155
            Width           =   1620
         End
         Begin VB.Label lblrcvdcash 
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
            ForeColor       =   &H0000C000&
            Height          =   405
            Left            =   9255
            TabIndex        =   30
            Top             =   1395
            Width           =   1665
         End
         Begin VB.Label lblopcash 
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
            ForeColor       =   &H00800080&
            Height          =   405
            Left            =   5970
            TabIndex        =   29
            Top             =   1395
            Width           =   1560
         End
         Begin VB.Label LBLEXPENSES 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Cash"
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
            Height          =   315
            Index           =   2
            Left            =   5850
            TabIndex        =   28
            Top             =   1140
            Width           =   1665
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
            TabIndex        =   26
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
            TabIndex        =   25
            Top             =   405
            Width           =   285
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   6465
            TabIndex        =   24
            Top             =   1965
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   8685
            TabIndex        =   23
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
         Height          =   480
         Left            =   6930
         TabIndex        =   15
         Top             =   8310
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
         Height          =   480
         Left            =   9945
         TabIndex        =   2
         Top             =   8310
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
         Height          =   480
         Left            =   8490
         TabIndex        =   1
         Top             =   8310
         Width           =   1380
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   5595
         Left            =   165
         TabIndex        =   3
         Top             =   2205
         Width           =   18465
         _ExtentX        =   32570
         _ExtentY        =   9869
         _Version        =   393216
         Rows            =   1
         Cols            =   14
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
         Height          =   330
         Left            =   5430
         TabIndex        =   14
         Tag             =   "5"
         Top             =   7860
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   582
         Picture         =   "FRMCASHBOOK.frx":0000
         ForeColor       =   0
         BarPicture      =   "FRMCASHBOOK.frx":001C
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
Attribute VB_Name = "FRMCASHBOOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CmDDisplay_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim FROMDATE As Date
    Dim TODATE As Date
    
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    ChkDelete.Value = False
    
    n = 1
    M = 0
    On Error GoTo ERRHAND
    Dim OPVAL, CLOVAL, RCVDVAL, ISSVAL As Double
    CLOVAL = 0
    
    OPVAL = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "select OPEN_DB from ACTMAST  WHERE ACT_CODE = '111001' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OPVAL = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "select SUM(OPEN_DB) from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3) ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OPVAL = OPVAL + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    FROMDATE = Format(DTFROM.Value, "yyyy/mm/dd")
    TODATE = Format(DTTo.Value, "yyyy/mm/dd")
    
    RCVDVAL = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT SUM(AMOUNT)  FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='S'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RCVDVAL = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    ISSVAL = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "select SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='P'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        ISSVAL = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
    
    lblopcash.Caption = Round(OPVAL + (RCVDVAL - ISSVAL), 2)

    lblpaidcash.Caption = 0
    lblrcvdcash.Caption = 0

    Set RSTTRXFILE = New ADODB.Recordset
    'RSTTRXFILE.Open "SELECT *  FROM CASHATRXFILE WHERE (INV_TYPE = 'WD') and VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    RSTTRXFILE.Open "SELECT *  FROM CASHATRXFILE WHERE AMOUNT <> 0 AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly, adCmdText
    If RSTTRXFILE.RecordCount > 0 Then
        vbalProgressBar1.Max = RSTTRXFILE.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If

    Do Until RSTTRXFILE.EOF
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = RSTTRXFILE!ACT_NAME
        GRDTranx.TextMatrix(M, 2) = RSTTRXFILE!INV_NO
        Select Case RSTTRXFILE!check_flag
            Case "S"
                GRDTranx.TextMatrix(M, 3) = RSTTRXFILE!AMOUNT
                lblrcvdcash.Caption = Val(lblrcvdcash.Caption) + RSTTRXFILE!AMOUNT
            Case "P"
                GRDTranx.TextMatrix(M, 4) = RSTTRXFILE!AMOUNT
                lblpaidcash.Caption = Val(lblpaidcash.Caption) + RSTTRXFILE!AMOUNT
        End Select
        GRDTranx.TextMatrix(M, 5) = RSTTRXFILE!INV_TYPE
        GRDTranx.TextMatrix(M, 6) = RSTTRXFILE!INV_TRX_TYPE
        Select Case GRDTranx.TextMatrix(M, 5)
            Case "RT"
                Select Case GRDTranx.TextMatrix(M, 6)
                    Case "WO"
                        GRDTranx.TextMatrix(M, 7) = "Petty Cash Sales"
                    Case "RI"
                        GRDTranx.TextMatrix(M, 7) = "8B Cash Sales"
                    Case "GI"
                        GRDTranx.TextMatrix(M, 7) = "B2B Sales"
                    Case "HI"
                        GRDTranx.TextMatrix(M, 7) = "Sales"
                    Case "SI"
                        GRDTranx.TextMatrix(M, 7) = "8 Cash Sales"
                    Case "SV"
                        GRDTranx.TextMatrix(M, 7) = "Service Bills"
                    Case "RT"
                        GRDTranx.TextMatrix(M, 7) = "Cash Receipt"
                End Select
            Case "ES"
                GRDTranx.TextMatrix(M, 7) = "Expense to Staff"
            Case "EX"
                GRDTranx.TextMatrix(M, 7) = "Office Expense"
            Case "FA"
                GRDTranx.TextMatrix(M, 7) = "Fixed Assets"
            Case "PY"
                GRDTranx.TextMatrix(M, 7) = "Cash Payment"
            Case "CN"
                GRDTranx.TextMatrix(M, 7) = "Credit Note"
            Case "DN"
                GRDTranx.TextMatrix(M, 7) = "Debit Note"
            Case "DP"
                GRDTranx.TextMatrix(M, 7) = "Bank Deposit"
            Case "WD"
                GRDTranx.TextMatrix(M, 7) = "Bank Withdrawal"
            Case "SR"
                GRDTranx.TextMatrix(M, 7) = "Sales Return"
            Case "RW"
                GRDTranx.TextMatrix(M, 7) = "Sales Return(W)"
        End Select
        GRDTranx.TextMatrix(M, 8) = RSTTRXFILE!VCH_DATE
        GRDTranx.TextMatrix(M, 9) = RSTTRXFILE!TRX_TYPE
        GRDTranx.TextMatrix(M, 10) = RSTTRXFILE!REC_NO
        GRDTranx.TextMatrix(M, 11) = RSTTRXFILE!INV_TRX_TYPE
        GRDTranx.TextMatrix(M, 12) = RSTTRXFILE!INV_TYPE
        GRDTranx.TextMatrix(M, 13) = RSTTRXFILE!INV_NO
        RSTTRXFILE.MoveNext
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    lblcloscash.Caption = Round(Val(lblopcash.Caption) + (Val(lblrcvdcash.Caption) - Val(lblpaidcash.Caption)), 2)
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
    
    On Error GoTo ERRHAND
    Dim i As Integer
    Dim FROMDATE As Date
    Dim TODATE As Date
    
    ChkDelete.Value = False
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTo.Value, "MM,DD,YYYY")
     ReportNameVar = Rptpath & "RptCashBook"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    '<=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "'
    
    'Report.RecordSelectionFormula = "(({CASHATRXFILE.VCH_DATE}<=# " & Format(DTTO.Value, "MM,DD,YYYY") & " #) AND ({CASHATRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #) AND (({CASHATRXFILE.INV_TYPE}='SI' OR {CASHATRXFILE.INV_TYPE}='WO') and {CASHATRXFILE.TRX_TYPE} ='DR') OR ({CASHATRXFILE.INV_TYPE}='RT' and {CASHATRXFILE.TRX_TYPE} ='DR') OR ({CASHATRXFILE.INV_TYPE}='PI' and {CASHATRXFILE.TRX_TYPE} ='DR') OR ({CASHATRXFILE.INV_TYPE}='PY' and {CASHATRXFILE.TRX_TYPE} ='DR') OR ({CASHATRXFILE.INV_TYPE}='EX' and {CASHATRXFILE.TRX_TYPE} ='DR'))"
    Report.RecordSelectionFormula = "({CASHATRXFILE.VCH_DATE}<=# " & Format(DTTo.Value, "MM,DD,YYYY") & " #) AND ({CASHATRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
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
        If CRXFormulaField.Name = "{@opcash}" Then CRXFormulaField.text = "'" & lblopcash.Caption & "' "
        If CRXFormulaField.Name = "{@clocash}" Then CRXFormulaField.text = "'" & lblcloscash.Caption & "' "
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "DAY BOOK"
    
    GENERATEREPORT
    'GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub DTFROM_Change()
    ChkDelete.Value = False
End Sub

Private Sub DTFROM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DTTo.SetFocus
    End Select
End Sub

Private Sub DTTO_Change()
    ChkDelete.Value = False
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
    GRDTranx.TextMatrix(0, 1) = "Account Head"
    GRDTranx.TextMatrix(0, 2) = "Bill No"
    GRDTranx.TextMatrix(0, 3) = "Credit"
    GRDTranx.TextMatrix(0, 4) = "Debit"
    GRDTranx.TextMatrix(0, 7) = "Description"
    GRDTranx.TextMatrix(0, 8) = "Date"
    
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 4000
    GRDTranx.ColWidth(2) = 1500
    GRDTranx.ColWidth(3) = 1500
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 0
    GRDTranx.ColWidth(6) = 0
    GRDTranx.ColWidth(7) = 3000
    GRDTranx.ColWidth(8) = 1300
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 3
    GRDTranx.ColAlignment(4) = 3
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Code"
    GRDBILL.TextMatrix(0, 2) = "Description"
    GRDBILL.TextMatrix(0, 3) = "Amount"
    GRDBILL.TextMatrix(0, 4) = "Remarks"
    

    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 0
    GRDBILL.ColWidth(2) = 3800
    GRDBILL.ColWidth(3) = 1200
    GRDBILL.ColWidth(4) = 4000
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(1) = 3
    GRDBILL.ColAlignment(2) = 1
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 11130
    'Me.Height = 10125
    Me.Left = 1500
    Me.Top = 0
    ACT_FLAG = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close

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
    
    If Not (Trim(GRDTranx.TextMatrix(GRDTranx.Row, 5)) = "FA" Or Trim(GRDTranx.TextMatrix(GRDTranx.Row, 5)) = "EX" Or Trim(GRDTranx.TextMatrix(GRDTranx.Row, 5)) = "ES") Then Exit Sub
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.rows = 1 Then Exit Sub
            LBLBILLNO.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 2)
            If Val(GRDTranx.TextMatrix(GRDTranx.Row, 3)) = 0 Then
                LBLBILLAMT.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 4), "0.00")
            Else
                LBLBILLAMT.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 3), "0.00")
            End If
            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            Select Case Trim(GRDTranx.TextMatrix(GRDTranx.Row, 5))
                Case "ES"
                    RSTTRXFILE.Open "Select * From TRXFILE_EXP WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "  AND TRX_TYPE = 'EX'", db, adOpenStatic, adLockReadOnly
                Case "EX"
                    RSTTRXFILE.Open "Select * From TRXEXPENSE WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "  AND TRX_TYPE = 'EX'", db, adOpenStatic, adLockReadOnly
                Case "FA"
                    RSTTRXFILE.Open "Select * From TRXFXDASSETS WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "  AND TRX_TYPE = 'FA'", db, adOpenStatic, adLockReadOnly
                Case Else
                    RSTTRXFILE.Open "Select * From TRXFILE_EXP WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "  AND TRX_TYPE = 'EX'", db, adOpenStatic, adLockReadOnly
            End Select
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.rows = GRDBILL.rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                Select Case Trim(GRDTranx.TextMatrix(GRDTranx.Row, 5))
                    Case "EX"
                        GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ACT_CODE
                        GRDBILL.TextMatrix(i, 2) = RSTTRXFILE!ACT_NAME
                        GRDBILL.TextMatrix(i, 3) = RSTTRXFILE!VCH_AMOUNT
                        GRDBILL.TextMatrix(i, 4) = RSTTRXFILE!REMARKS
                    Case "ES"
                        GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!EXP_CODE
                        GRDBILL.TextMatrix(i, 2) = RSTTRXFILE!EXP_NAME
                        GRDBILL.TextMatrix(i, 3) = RSTTRXFILE!TRX_TOTAL
                        GRDBILL.TextMatrix(i, 4) = RSTTRXFILE!REMARKS
                    Case "FA"
                        GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ACT_CODE
                        GRDBILL.TextMatrix(i, 2) = RSTTRXFILE!ACT_NAME
                        GRDBILL.TextMatrix(i, 3) = RSTTRXFILE!VCH_AMOUNT
                        GRDBILL.TextMatrix(i, 4) = RSTTRXFILE!REMARKS
                End Select
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing

            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    End Select
End Sub

Private Sub GRDTranx_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRHAND
    Select Case KeyAscii
        Case vbKeyD, Asc("d")
            CMDDISPLAY.Tag = KeyAscii
        Case vbKeyE, Asc("e")
            CMDEXIT.Tag = KeyAscii
        Case vbKeyL, Asc("l")
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
            If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                If ChkDelete.Value = False Then Exit Sub
                If MsgBox("Are you sure you want to delete this entry", vbYesNo + vbDefaultButton2, "DELETE !!!") = vbNo Then
                    ChkDelete.Value = False
                    Exit Sub
                End If
                db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 13) & " AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 10) & " AND TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "'"
                Call CmDDisplay_Click
            End If
        Case Else
            CMDEXIT.Tag = ""
            CMDDISPLAY.Tag = ""
    End Select
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

'Private Sub TMPDELETE_Click()
'    If GRDTranx.Rows = 1 Then Exit Sub
'    If MsgBox("Are You Sure You want to Delete PRINT_BILL NO." & "*** " & GRDTranx.TextMatrix(GRDTranx.Row, 2) & " ****", vbYesNo, "DELETING BILL....") = vbNo Then Exit Sub
'    DB.Execute ("DELETE from SALESREG WHERE VCH_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 2) & " AND (TRX_TYPE='SI' OR TRX_TYPE='SI')")
'    Call fillSTOCKREG
'
'End Sub
'
'Private Function fillSTOCKREG()
'    Dim rstTRANX As ADODB.Recordset
'    Dim i As lONG
'
'    LBLTRXTOTAL.Caption = "0.00"
'    LBLDISCOUNT.Caption = "0.00"
'    LBLNET.Caption = "0.00"
'    LBLCOST.Caption = "0.00"
'    LBLPROFIT.Caption = "0.00"
'
'   On Error GoTo eRRHAND
'
'
'    Screen.MousePointer = vbHourglass
'
'    GRDTranx.Rows = 1
'    i = 0
'    GRDTranx.Visible = False
'    vbalProgressBar1.Value = 0
'    vbalProgressBar1.ShowText = True
'    vbalProgressBar1.Text = "PLEASE WAIT..."
'
'    Set rstTRANX = New ADODB.Recordset
'    rstTRANX.Open "SELECT * From SALESREG", DB, adOpenStatic,adLockReadOnly
'    Do Until rstTRANX.EOF
'        i = i + 1
'        GRDTranx.Rows = GRDTranx.Rows + 1
'        GRDTranx.FixedRows = 1
'        GRDTranx.TextMatrix(i, 0) = i
'        GRDTranx.TextMatrix(i, 2) = rstTRANX!VCH_NO
'        GRDTranx.TextMatrix(i, 3) = rstTRANX!VCH_DATE
'        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!VCH_AMOUNT, "0.00")
'        GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!DISCOUNT, "0.00")
'        GRDTranx.TextMatrix(i, 6) = Format(Val(GRDTranx.TextMatrix(i, 4)) - Val(GRDTranx.TextMatrix(i, 4)) * Val(GRDTranx.TextMatrix(i, 5)) / 100)
'        GRDTranx.TextMatrix(i, 7) = Format(rstTRANX!PAYAMOUNT, "0.00")
'        GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
'
'        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
'        LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption) + rstTRANX!DISCOUNT, "0.00")
'        LBLNET.Caption = Format(Val(LBLTRXTOTAL.Caption) - Val(LBLDISCOUNT.Caption), "0.00")
'        LBLCOST.Caption = Format(Val(LBLCOST.Caption) + rstTRANX!PAYAMOUNT, "0.00")
'        LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
'
'        vbalProgressBar1.Max = rstTRANX.RecordCount
'        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
'    Loop
'
'    rstTRANX.Close
'    Set rstTRANX = Nothing
'
'    vbalProgressBar1.ShowText = False
'    vbalProgressBar1.Value = 0
'    GRDTranx.Visible = True
'    Screen.MousePointer = vbDefault
'    Exit Function
'
'eRRHAND:
'    Screen.MousePointer = vbDefault
'    MsgBox Err.Description
'End Function

Private Sub OPTCUSTOMER_Click()
    ChkDelete.Value = False
    TXTDEALER.SetFocus
End Sub

Private Sub OPTCUSTOMER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub

Private Sub OPTPERIOD_Click()
    ChkDelete.Value = False
End Sub

Private Sub TXTDEALER_GotFocus()
    OPTCUSTOMER.Value = True
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select

End Sub

Private Sub TXTDEALER_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
    ChkDelete.Value = False
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    GRDTranx.rows = 1
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub


