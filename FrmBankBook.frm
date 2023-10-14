VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMBankBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BANK BOOK"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17625
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   17625
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      Caption         =   "PRESS ESC TO CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5520
      Left            =   75
      TabIndex        =   9
      Top             =   2250
      Visible         =   0   'False
      Width           =   9735
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4080
         Left            =   105
         TabIndex        =   10
         Top             =   1335
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   7197
         _Version        =   393216
         Rows            =   1
         Cols            =   7
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LBLINVDATE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   135
         TabIndex        =   21
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV DATE"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   4
         Left            =   405
         TabIndex        =   20
         Top             =   735
         Width           =   885
      End
      Begin VB.Label LBLSUPPLIER 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1140
         TabIndex        =   19
         Top             =   315
         Width           =   4410
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer\"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   18
         Top             =   345
         Width           =   885
      End
      Begin VB.Label LBLBILLAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2535
         TabIndex        =   14
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV AMT"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   2685
         TabIndex        =   13
         Top             =   735
         Width           =   810
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV NO"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   1665
         TabIndex        =   12
         Top             =   735
         Width           =   675
      End
      Begin VB.Label LBLBILLNO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1515
         TabIndex        =   11
         Top             =   975
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFC0&
      Height          =   8925
      Left            =   -75
      TabIndex        =   0
      Top             =   -90
      Width           =   17715
      Begin VB.ComboBox CmbMode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmBankBook.frx":0000
         Left            =   12885
         List            =   "FrmBankBook.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2865
         Visible         =   0   'False
         Width           =   2610
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
         Left            =   11130
         TabIndex        =   42
         Top             =   3300
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
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
         Left            =   6855
         TabIndex        =   5
         Top             =   345
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         Height          =   1260
         Left            =   10860
         TabIndex        =   38
         Top             =   570
         Width           =   2715
         Begin VB.OptionButton OptAll 
            BackColor       =   &H00C0C0FF&
            Caption         =   "All"
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   90
            TabIndex        =   40
            Top             =   735
            Value           =   -1  'True
            Width           =   2505
         End
         Begin VB.OptionButton OptPend 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Cheque Pending"
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   90
            TabIndex        =   39
            Top             =   330
            Width           =   2505
         End
      End
      Begin MSMask.MaskEdBox TXTEXPIRY 
         Height          =   390
         Left            =   11235
         TabIndex        =   37
         Top             =   4335
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.ComboBox CmbPend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmBankBook.frx":0032
         Left            =   12345
         List            =   "FrmBankBook.frx":003C
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   4320
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "TOTAL"
         ForeColor       =   &H000000FF&
         Height          =   870
         Left            =   120
         TabIndex        =   23
         Top             =   8010
         Width           =   9810
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Op. Balance"
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
            Height          =   300
            Index           =   8
            Left            =   1035
            TabIndex        =   32
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label lblOPBal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   990
            TabIndex        =   31
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLPAIDAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   5325
            TabIndex        =   29
            Top             =   435
            Width           =   1875
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Debit Amt"
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
            Height          =   240
            Index           =   3
            Left            =   5340
            TabIndex        =   28
            Top             =   150
            Width           =   1875
         End
         Begin VB.Label LBLINVAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   3135
            TabIndex        =   27
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Amt"
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
            Height          =   300
            Index           =   6
            Left            =   3180
            TabIndex        =   26
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label LBLBALAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   7395
            TabIndex        =   25
            Top             =   435
            Width           =   1830
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bal Amt"
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
            Height          =   315
            Index           =   7
            Left            =   7395
            TabIndex        =   24
            Top             =   150
            Width           =   1815
         End
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
         Left            =   9345
         TabIndex        =   7
         Top             =   345
         Width           =   1170
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
         Left            =   8115
         TabIndex        =   6
         Top             =   345
         Width           =   1185
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0FFC0&
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
         Height          =   1740
         Left            =   105
         TabIndex        =   15
         Top             =   120
         Width           =   6465
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
            Height          =   360
            Left            =   1230
            TabIndex        =   1
            Top             =   210
            Width           =   3735
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1230
            TabIndex        =   2
            Top             =   585
            Width           =   3735
            _ExtentX        =   6588
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
         Begin VB.PictureBox rptPRINT 
            Height          =   480
            Left            =   9990
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   30
            Top             =   -45
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1230
            TabIndex        =   3
            Top             =   1275
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   94109697
            CurrentDate     =   40498
            MinDate         =   -614701
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   3150
            TabIndex        =   4
            Top             =   1275
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   688
            _Version        =   393216
            Format          =   94109697
            CurrentDate     =   40498
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Period"
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
            Left            =   135
            TabIndex        =   35
            Top             =   1320
            Width           =   915
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
            Index           =   10
            Left            =   2805
            TabIndex        =   34
            Top             =   1350
            Width           =   285
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
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
            Index           =   5
            Left            =   135
            TabIndex        =   22
            Top             =   345
            Width           =   1485
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   16
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   17
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6150
         Left            =   105
         TabIndex        =   8
         Top             =   1860
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   10848
         _Version        =   393216
         Rows            =   1
         Cols            =   22
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
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
      Begin MSForms.CheckBox ChkDelete 
         Height          =   270
         Left            =   16305
         TabIndex        =   41
         Top             =   255
         Width           =   1350
         BackColor       =   12648384
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F6 to make Bank Entries"
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
         Height          =   300
         Index           =   8
         Left            =   6690
         TabIndex        =   33
         Top             =   1515
         Width           =   3120
      End
   End
End
Attribute VB_Name = "FRMBankBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CmDDisplay_Click()
    Call Fillgrid
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim Op_Bal, OP_DR, OP_CR As Double

    Dim RSTTRXFILE, rstTRANX As ADODB.Recordset
    Dim i As Long
    
    If DataList2.BoundText = "" Then
        MsgBox "please Select Bank from the List", vbOKOnly, "Statement"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Op_Bal = 0
    OP_DR = 0
    OP_CR = 0
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select OPEN_DB from BANKCODE  WHERE BANK_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        Op_Bal = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "Select * from BANK_TRX WHERE BANK_CODE = '" & DataList2.BoundText & "' and TRX_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rstTRANX.EOF
        Select Case rstTRANX!TRX_TYPE
            Case "DR"
                OP_DR = OP_DR + IIf(IsNull(rstTRANX!TRX_AMOUNT), 0, rstTRANX!TRX_AMOUNT)
            Case "CR"
                OP_CR = OP_CR + IIf(IsNull(rstTRANX!TRX_AMOUNT), 0, rstTRANX!TRX_AMOUNT)
        End Select
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    Op_Bal = Op_Bal + OP_CR - OP_DR
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM BANKCODE  WHERE BANK_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!OPEN_CR = Op_Bal
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim BAL_AMOUNT As Double
    BAL_AMOUNT = 0
    Set RSTTRXFILE = New ADODB.Recordset
    '< '" & Format(DTFROM.value, "yyyy/mm/dd") & "'
    RSTTRXFILE.Open "SELECT * From BANK_TRX WHERE BANK_CODE = '" & DataList2.BoundText & "' AND TRX_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND TRX_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY BNK_SL_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        'RSTTRXFILE!BAL_AMT = Op_Bal
        Select Case RSTTRXFILE!TRX_TYPE
            Case "CR"
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!TRX_AMOUNT), 0, RSTTRXFILE!TRX_AMOUNT) '- IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
            Case Else
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!TRX_AMOUNT), 0, RSTTRXFILE!TRX_AMOUNT)
                'RSTTRXFILE!BAL_AMT = Op_Bal

        End Select
        RSTTRXFILE!BAL_AMT = BAL_AMOUNT
        Op_Bal = 0
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
'    Dim DR_AMOUNT As Double
'    Dim CR_AMOUNT As Double
'    Dim CR_FLAG As Boolean
'    CR_FLAG = False
'    DR_AMOUNT = 0
'    CR_AMOUNT = 0
'    Set RSTTRXFILE = New ADODB.Recordset
'    If Optall.value = True Then
'        RSTTRXFILE.Open "SELECT * From BANK_TRX WHERE BANK_CODE = '" & DataList2.BoundText & "' AND TRX_DATE <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND TRX_DATE >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # ORDER BY BNK_SL_NO", db, adOpenForwardOnly
'    Else
'        RSTTRXFILE.Open "SELECT * From BANK_TRX WHERE BANK_FLAG ='Y' AND CHECK_FLAG = 'N' AND BANK_CODE = '" & DataList2.BoundText & "' AND TRX_DATE <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND TRX_DATE >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # ORDER BY BNK_SL_NO", db, adOpenForwardOnly
'    End If
'    Do Until RSTTRXFILE.EOF
'        CR_FLAG = True
'        Select Case RSTTRXFILE!TRX_TYPE
'            Case "DR"
'                DR_AMOUNT = DR_AMOUNT + IIf(IsNull(RSTTRXFILE!TRX_AMOUNT), 0, RSTTRXFILE!TRX_AMOUNT)
'            Case "CR"
'                CR_AMOUNT = CR_AMOUNT + IIf(IsNull(RSTTRXFILE!TRX_AMOUNT), 0, RSTTRXFILE!TRX_AMOUNT)
'        End Select
'        RSTTRXFILE.MoveNext
'    Loop
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    
    
    Screen.MousePointer = vbHourglass
    Sleep (300)
    ReportNameVar = Rptpath & "RptBANKREPORT1"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.RecordSelectionFormula = "({CUSTMAST.CR_FLAG}='Y')"
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
    Report.RecordSelectionFormula = "({BANK_TRX.BANK_CODE}='" & DataList2.BoundText & "') and ({BANK_TRX.TRX_TYPE} = 'DR' OR {BANK_TRX.TRX_TYPE} = 'CR') AND ({BANK_TRX.TRX_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {BANK_TRX.TRX_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Statement of ' & '" & UCase(DataList2.text) & "' & ' for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Activate()
    Call Fillgrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF6
            If DataList2.BoundText = "" Then Exit Sub
            'If GRDTranx.Rows <= 1 Then Exit Sub
            Enabled = False
            FRMbankentry.LBLSUPPLIER.Caption = DataList2.text
            FRMbankentry.lblactcode.Caption = DataList2.BoundText
            'FRMRECEIPTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMRECEIPTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMRECEIPTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMRECEIPTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMRECEIPTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMbankentry.Show
    End Select
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "DATE"
    GRDTranx.TextMatrix(0, 2) = "TRX TYPE"
    GRDTranx.TextMatrix(0, 3) = "Credit"
    GRDTranx.TextMatrix(0, 4) = "Debit"
    GRDTranx.TextMatrix(0, 5) = "Balance"
    GRDTranx.TextMatrix(0, 6) = "Ref."
    GRDTranx.TextMatrix(0, 7) = "Mode"
    GRDTranx.TextMatrix(0, 8) = "Pymnt Pending?"
    GRDTranx.TextMatrix(0, 9) = "Chq Date"
    GRDTranx.TextMatrix(0, 10) = "Chq No."
    GRDTranx.TextMatrix(0, 16) = "Remarks"
    
    GRDTranx.ColWidth(0) = 850
    GRDTranx.ColWidth(1) = 1400
    GRDTranx.ColWidth(2) = 1100
    GRDTranx.ColWidth(3) = 1400
    GRDTranx.ColWidth(4) = 1400
    GRDTranx.ColWidth(5) = 1400
    GRDTranx.ColWidth(6) = 1900
    GRDTranx.ColWidth(7) = 1200
    GRDTranx.ColWidth(8) = 1400
    GRDTranx.ColWidth(9) = 1400
    GRDTranx.ColWidth(10) = 1800
    GRDTranx.ColWidth(11) = 0
    GRDTranx.ColWidth(12) = 0
    GRDTranx.ColWidth(13) = 0
    GRDTranx.ColWidth(14) = 0
    GRDTranx.ColWidth(15) = 0
    GRDTranx.ColWidth(16) = 2500
    GRDTranx.ColWidth(17) = 0
    GRDTranx.ColWidth(18) = 0
    GRDTranx.ColWidth(19) = 0
    GRDTranx.ColWidth(20) = 0
    GRDTranx.ColWidth(21) = 0
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 1
    GRDTranx.ColAlignment(7) = 4
    GRDTranx.ColAlignment(8) = 4
    GRDTranx.ColAlignment(9) = 4
    GRDTranx.ColAlignment(10) = 1
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "Rate"
    GRDBILL.TextMatrix(0, 3) = "Disc %"
    GRDBILL.TextMatrix(0, 4) = "Tax %"
    GRDBILL.TextMatrix(0, 5) = "Qty"
    GRDBILL.TextMatrix(0, 6) = "Amount"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 2800
    GRDBILL.ColWidth(2) = 800
    GRDBILL.ColWidth(3) = 800
    GRDBILL.ColWidth(4) = 800
    GRDBILL.ColWidth(5) = 900
    GRDBILL.ColWidth(6) = 1100
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 3
    GRDBILL.ColAlignment(6) = 6
    
    DTFROM.Value = "01/04/" & Year(MDIMAIN.DTFROM.Value)
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 1500
    Top = 0
    TXTDEALER.text = " "
    TXTDEALER.text = ""
    'MDIMAIN.MNUPYMNT.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close

    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.MNUPYMNT.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
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
        Frmeperiod.Enabled = True
        FRMEBILL.Visible = False
        GRDTranx.SetFocus
    End If
End Sub

Private Sub GRDTranx_Click()
    TXTsample.Visible = False
    TXTEXPIRY.Visible = False
    CmbPend.Visible = False
    CmbMode.Visible = False
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim i As Long
    Dim E_TABLE As String
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            
            Select Case GRDTranx.Col
                Case 7
                    If GRDTranx.TextMatrix(GRDTranx.Row, 7) = "Cash" Then Exit Sub
                    CmbMode.Visible = True
                    CmbMode.Top = GRDTranx.CellTop + 1900
                    CmbMode.Left = GRDTranx.CellLeft + 150
                    CmbMode.Width = GRDTranx.CellWidth
                    On Error Resume Next
                    CmbMode.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                    err.Clear
                    On Error GoTo ERRHAND
                    CmbMode.SetFocus
                 Case 8
                    If GRDTranx.TextMatrix(GRDTranx.Row, 7) = "Cash" Then Exit Sub
                    CmbPend.Visible = True
                    CmbPend.Top = GRDTranx.CellTop + 1900
                    CmbPend.Left = GRDTranx.CellLeft + 150
                    CmbPend.Width = GRDTranx.CellWidth
                    'CmbPend.Height = GRDTranx.CellHeight
                    If Trim(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)) = "Passed" Then
                        CmbPend.ListIndex = 0
                    Else
                        CmbPend.ListIndex = 1
                   End If
                    CmbPend.SetFocus
                Case 1
                    If GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Cheque Returned" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Receipt" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Payment" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Debit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Credit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Off Expense" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Staff Expense" Then Exit Sub
                    TXTEXPIRY.Visible = True
                    TXTEXPIRY.Top = GRDTranx.CellTop + 1850
                    TXTEXPIRY.Left = GRDTranx.CellLeft + 100
                    TXTEXPIRY.Width = GRDTranx.CellWidth
                    TXTEXPIRY.Height = GRDTranx.CellHeight
                    If Not (IsDate(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))) Then
                        TXTEXPIRY.text = "  /  /    "
                    Else
                        TXTEXPIRY.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                    End If
                    
                    TXTEXPIRY.SetFocus
                Case 3, 4
                    If Val(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)) = 0 Then Exit Sub
                    If GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Cheque Returned" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Receipt" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Payment" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Debit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Credit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Off Expense" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Staff Expense" Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDTranx.CellTop + 1850
                    TXTsample.Left = GRDTranx.CellLeft + 100
                    TXTsample.Width = GRDTranx.CellWidth
                    TXTsample.Height = GRDTranx.CellHeight
                    TXTsample.text = Format(Val(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)), "0.00")
                    TXTsample.SetFocus
                Case 6
                    If Val(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)) = 0 Then Exit Sub
                    If GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Cheque Returned" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Receipt" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Payment" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Debit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Credit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Off Expense" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Staff Expense" Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDTranx.CellTop + 1850
                    TXTsample.Left = GRDTranx.CellLeft + 100
                    TXTsample.Width = GRDTranx.CellWidth
                    TXTsample.Height = GRDTranx.CellHeight
                    TXTsample.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                    TXTsample.SetFocus
                    
            End Select
            
            Exit Sub
            If GRDTranx.rows = 1 Then Exit Sub
            If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Receipt" Then Exit Sub
            LBLSUPPLIER.Caption = " " & DataList2.text
            LBLINVDATE.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 2)
            LBLBILLNO.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 3)
            LBLBILLAMT.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'LBLPAID.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'LBLBAL.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 6)

            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILE WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "'", db, adOpenForwardOnly
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.rows = GRDBILL.rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ITEM_NAME
                GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!SALES_PRICE, "0.00")
                GRDBILL.TextMatrix(i, 3) = Val(RSTTRXFILE!LINE_DISC)
                GRDBILL.TextMatrix(i, 4) = Val(RSTTRXFILE!SALES_TAX)
                GRDBILL.TextMatrix(i, 5) = RSTTRXFILE!QTY
                GRDBILL.TextMatrix(i, 6) = Format(RSTTRXFILE!TRX_TOTAL, "0.00")
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing

            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
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
            If ChkDelete.Value = False Then
                If GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Cheque Returned" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Receipt" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Payment" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Debit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Credit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Off Expense" Or GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Staff Expense" Then Exit Sub
            End If
            
            If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                If MsgBox("Are you sure you want to delete this entry", vbYesNo + vbDefaultButton2, "DELETE !!!") = vbYes Then
                    If GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Contra" Then
                        db.BeginTrans
                        db.Execute "delete FROM BANK_TRX WHERE TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' "
                        db.CommitTrans
                    Else
                        db.BeginTrans
                        'db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "'"
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 17) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 18) = "") Then
                            db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 18) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 19) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 20) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 21) & " "
                        End If
                        db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' "
                        db.CommitTrans
                    End If
                    Call Fillgrid
                Else
                    GRDTranx.SetFocus
                End If
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

Private Sub GRDTranx_Scroll()
    TXTsample.Visible = False
    TXTEXPIRY.Visible = False
    CmbPend.Visible = False
    CmbMode.Visible = False
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
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
            ACT_REC.Open "select BANK_CODE, BANK_NAME from BANKCODE  WHERE BANK_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select BANK_CODE, BANK_NAME from BANKCODE  WHERE BANK_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!BANK_NAME
        End If
        Set DataList2.RowSource = ACT_REC
        DataList2.ListField = "BANK_NAME"
        DataList2.BoundColumn = "BANK_CODE"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    GRDTranx.rows = 1
    Call Fillgrid
    'LBL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
           
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

Private Function Fillgrid()
    Dim rstTRANX As ADODB.Recordset
    Dim i As Long
    
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.rows = 1
    LBLINVAMT.Caption = ""
    LBLPAIDAMT.Caption = ""
    LBLBALAMT.Caption = ""
    lblOPBal.Caption = ""
    i = 1
    
    Dim Op_Bal, OP_DR, OP_CR As Double
    Op_Bal = 0
    OP_DR = 0
    OP_CR = 0
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select OPEN_DB from BANKCODE  WHERE BANK_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        Op_Bal = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "Select * from BANK_TRX WHERE BANK_CODE = '" & DataList2.BoundText & "' and TRX_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rstTRANX.EOF
        Select Case rstTRANX!TRX_TYPE
            Case "DR"
                OP_DR = OP_DR + IIf(IsNull(rstTRANX!TRX_AMOUNT), 0, rstTRANX!TRX_AMOUNT)
            Case "CR"
                OP_CR = OP_CR + IIf(IsNull(rstTRANX!TRX_AMOUNT), 0, rstTRANX!TRX_AMOUNT)
        End Select
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    lblOPBal.Caption = Format(Op_Bal + OP_CR - OP_DR, "0.00")
    
    Set rstTRANX = New ADODB.Recordset
    If Optall.Value = True Then
        rstTRANX.Open "SELECT * From BANK_TRX WHERE BANK_CODE = '" & DataList2.BoundText & "' AND TRX_DATE <='" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND TRX_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY TRX_DATE, BNK_SL_NO", db, adOpenForwardOnly
    Else
        rstTRANX.Open "SELECT * From BANK_TRX WHERE BANK_FLAG ='Y' AND CHECK_FLAG = 'N' AND BANK_CODE = '" & DataList2.BoundText & "' AND TRX_DATE <='" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND TRX_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY TRX_DATE,  BNK_SL_NO", db, adOpenForwardOnly
    End If
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0
        GRDTranx.TextMatrix(i, 0) = i
        GRDTranx.TextMatrix(i, 1) = Format(rstTRANX!TRX_DATE, "DD/MM/YYYY")
        Select Case rstTRANX!TRX_TYPE
            Case "DR"
                Select Case rstTRANX!BILL_TRX_TYPE
                    Case "PY", "ML"
                        GRDTranx.TextMatrix(i, 2) = "Payment"
                    Case "WD"
                        GRDTranx.TextMatrix(i, 2) = "Withdrawal"
                    Case "BC"
                        GRDTranx.TextMatrix(i, 2) = "Bank Charges"
                    Case "DN"
                        GRDTranx.TextMatrix(i, 2) = "Credit Note"
                    Case "CT"
                        GRDTranx.TextMatrix(i, 2) = "Contra"
                    Case "EX"
                        GRDTranx.TextMatrix(i, 2) = "Off Expense"
                    Case "DB"
                        GRDTranx.TextMatrix(i, 2) = "Credit Note"
                    Case "ES"
                        GRDTranx.TextMatrix(i, 2) = "Staff Expense"
                    Case "FP"
                        GRDTranx.TextMatrix(i, 2) = "Payment"
                    Case "RC"
                        GRDTranx.TextMatrix(i, 2) = "Cheque Returned"
                    Case Else
                        GRDTranx.TextMatrix(i, 2) = "Others"
                End Select
                GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!TRX_AMOUNT, "0.00")
'                GRDTranx.TextMatrix(i, 0) = "Sale"
'                GRDTranx.CellForeColor = vbRed
            Case "CR"
                Select Case rstTRANX!BILL_TRX_TYPE
                    Case "RT", "ML"
                        GRDTranx.TextMatrix(i, 2) = "Receipt"
                    Case "DP"
                        GRDTranx.TextMatrix(i, 2) = "Deposit"
                    Case "IN"
                        GRDTranx.TextMatrix(i, 2) = "Bank Interest"
                    Case "CN"
                        GRDTranx.TextMatrix(i, 2) = "Debit Note"
                    Case "CB"
                        GRDTranx.TextMatrix(i, 2) = "Debit Note"
                    Case "CT"
                        GRDTranx.TextMatrix(i, 2) = "Contra"
                    Case "FR"
                        GRDTranx.TextMatrix(i, 2) = "Receipt"
                    Case "RC"
                        GRDTranx.TextMatrix(i, 2) = "Cheque Returned"
                    Case Else
                        GRDTranx.TextMatrix(i, 2) = "Others"
                End Select
                GRDTranx.TextMatrix(i, 3) = Format(rstTRANX!TRX_AMOUNT, "0.00")
        End Select
        If i = 1 Then
            GRDTranx.TextMatrix(i, 5) = Format(Val(lblOPBal.Caption) - Val(GRDTranx.TextMatrix(i, 4)) + Val(GRDTranx.TextMatrix(i, 3)), "0.00")
        Else
            GRDTranx.TextMatrix(i, 5) = Format(Val(GRDTranx.TextMatrix(i - 1, 5)) - Val(GRDTranx.TextMatrix(i, 4)) + Val(GRDTranx.TextMatrix(i, 3)), "0.00")
        End If
        GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        Select Case rstTRANX!BANK_FLAG
            Case "Y"
                'GRDTranx.TextMatrix(i, 7) = "Cheque"
                Select Case rstTRANX!BANK_MODE
                    Case "C"
                        GRDTranx.TextMatrix(i, 7) = "Cheque"
                    Case "U"
                        GRDTranx.TextMatrix(i, 7) = "UPI Transfer"
                    Case "N"
                        GRDTranx.TextMatrix(i, 7) = "NEFT/RTGS"
                    Case Else
                        GRDTranx.TextMatrix(i, 7) = "Cheque"
                End Select
                Select Case rstTRANX!check_flag
                    Case "Y"
                        GRDTranx.TextMatrix(i, 8) = "Passed"
                    Case Else
                        GRDTranx.TextMatrix(i, 8) = "Pending"
                End Select
                GRDTranx.TextMatrix(i, 9) = IIf(IsNull(rstTRANX!CHQ_DATE), "", Format(rstTRANX!CHQ_DATE, "DD/MM/YYYY"))
                GRDTranx.TextMatrix(i, 10) = IIf(IsNull(rstTRANX!CHQ_NO), "", rstTRANX!CHQ_NO)
                
                GRDTranx.TextMatrix(i, 17) = ""
                GRDTranx.TextMatrix(i, 18) = ""
                GRDTranx.TextMatrix(i, 19) = ""
                GRDTranx.TextMatrix(i, 20) = ""
                GRDTranx.TextMatrix(i, 21) = ""
            Case Else
                GRDTranx.TextMatrix(i, 7) = "Cash"
                GRDTranx.TextMatrix(i, 8) = ""
                GRDTranx.TextMatrix(i, 9) = ""
                GRDTranx.TextMatrix(i, 10) = ""
                
                GRDTranx.TextMatrix(i, 17) = IIf(IsNull(rstTRANX!C_TRX_TYPE), "", rstTRANX!C_TRX_TYPE)
                GRDTranx.TextMatrix(i, 18) = IIf(IsNull(rstTRANX!C_REC_NO), "", rstTRANX!C_REC_NO)
                GRDTranx.TextMatrix(i, 19) = IIf(IsNull(rstTRANX!C_INV_TRX_TYPE), "", rstTRANX!C_INV_TRX_TYPE)
                GRDTranx.TextMatrix(i, 20) = IIf(IsNull(rstTRANX!C_INV_TYPE), "", rstTRANX!C_INV_TYPE)
                GRDTranx.TextMatrix(i, 21) = IIf(IsNull(rstTRANX!C_INV_NO), "", rstTRANX!C_INV_NO)
        End Select
        GRDTranx.TextMatrix(i, 11) = IIf(IsNull(rstTRANX!TRX_NO), "", rstTRANX!TRX_NO)
        GRDTranx.TextMatrix(i, 12) = IIf(IsNull(rstTRANX!TRX_TYPE), "", rstTRANX!TRX_TYPE)
        GRDTranx.TextMatrix(i, 13) = IIf(IsNull(rstTRANX!BILL_TRX_TYPE), "", rstTRANX!BILL_TRX_TYPE)
        GRDTranx.TextMatrix(i, 14) = IIf(IsNull(rstTRANX!TRX_YEAR), "", rstTRANX!TRX_YEAR)
        GRDTranx.TextMatrix(i, 15) = IIf(IsNull(rstTRANX!BANK_CODE), "", rstTRANX!BANK_CODE)
        GRDTranx.TextMatrix(i, 16) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        
        'GRDTranx.Row = i
        'GRDTranx.Col = 0
        LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + Val(GRDTranx.TextMatrix(i, 3)), "0.00")
        LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + Val(GRDTranx.TextMatrix(i, 4)), "0.00")
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    
    LBLBALAMT.Caption = Format(Val(lblOPBal.Caption) + (Val(LBLINVAMT.Caption) - Val(LBLPAIDAMT.Caption)), "0.00")
    ChkDelete.Value = False
    flagchange.Caption = ""
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Function

Private Sub CmbPend_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPend.ListIndex = -1 Then Exit Sub
            If CmbPend.text = "Passed" And DateValue(GRDTranx.TextMatrix(GRDTranx.Row, 9)) > DateValue(Date) Then
                MsgBox "From Date could not be higher than Today.. Please make proper changes", vbOKOnly, "CHEQUE POSTING..."
            End If
                    
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (rststock.EOF And rststock.BOF) Then
                If CmbPend.ListIndex = 0 Then
                    rststock!check_flag = "Y"
                Else
                    rststock!check_flag = "N"
                End If
                rststock.Update
                GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = CmbPend.text
            End If
            rststock.Close
            Set rststock = Nothing
            GRDTranx.Enabled = True
            CmbPend.Visible = False
            GRDTranx.SetFocus
        Case vbKeyEscape
            CmbPend.Visible = False
            GRDTranx.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmbPend_LostFocus()
    CmbPend.Visible = False
End Sub


Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 9
                    If Not (IsDate(TXTEXPIRY.text)) Then
                        TXTEXPIRY.SetFocus
                        Exit Sub
                    End If
                    If GRDTranx.TextMatrix(GRDTranx.Row, 8) = "Passed" And DateValue(TXTEXPIRY.text) > DateValue(Date) Then
                        MsgBox "From Date could not be higher than Today", vbOKOnly, "CHEQUE POSTING..."
                        TXTEXPIRY.SetFocus
                        Exit Sub
                    End If
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!CHQ_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                        rststock.Update
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDTranx.Enabled = True
                    TXTEXPIRY.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                Case 1
                    If Not (IsDate(TXTEXPIRY.text)) Then
                        TXTEXPIRY.SetFocus
                        Exit Sub
                    End If
                    If GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Contra" Then
                        db.BeginTrans
                        'db.Execute "delete FROM BANK_TRX WHERE TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' "
                        'Format(rstTRANX!TRX_DATE, "DD/MM/YYYY")
                        db.Execute "UPDATE BANK_TRX SET TRX_DATE = '" & Format(TXTEXPIRY.text, "yyyy/mm/dd") & "' WHERE TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' "
                        db.CommitTrans
                    Else
                        db.BeginTrans
                        'db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "'"
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 17) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 18) = "") Then
                            'db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 18) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 19) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 20) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 21) & " "
                            db.Execute "UPDATE CASHATRXFILE SET VCH_DATE = '" & Format(TXTEXPIRY.text, "yyyy/mm/dd") & "' WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 18) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 19) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 20) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 21) & " "
                        End If
                        'db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' "
                        db.Execute "UPDATE BANK_TRX SET TRX_DATE = '" & Format(TXTEXPIRY.text, "yyyy/mm/dd") & "' WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' "
                        db.CommitTrans
                    End If
                    
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                    GRDTranx.Enabled = True
                    TXTEXPIRY.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
            End Select
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            GRDTranx.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTEXPIRY_LostFocus()
    TXTEXPIRY.Visible = False
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 3, 4 ' AMOUNT
                    If Val(TXTsample.text) = 0 Then Exit Sub
                    If GRDTranx.TextMatrix(GRDTranx.Row, 2) = "Contra" Then
                        db.BeginTrans
                        'db.Execute "delete FROM BANK_TRX WHERE TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' "
                        'Format(rstTRANX!TRX_DATE, "DD/MM/YYYY")
                        db.Execute "UPDATE BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' "
                        db.CommitTrans
                    Else
                        db.BeginTrans
                        'db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "'"
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 17) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 18) = "") Then
                            'db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 18) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 19) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 20) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 21) & " "
                            db.Execute "UPDATE CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 18) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 19) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 20) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 21) & " "
                        End If
                        'db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' "
                        db.Execute "UPDATE BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' "
                        db.CommitTrans
                    End If
                    
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                                    
                Case 6 ' REF
                    db.Execute "UPDATE BANK_TRX SET REF_NO = '" & Trim(TXTsample.text) & "' WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' "
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                    
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDTranx.SetFocus
    End Select
        Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDTranx.Col
        Case 3, 4
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub CmbMode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If CmbMode.ListIndex = -1 Then Exit Sub
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 11)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (rststock.EOF And rststock.BOF) Then
                If CmbMode.ListIndex = 0 Then
                    rststock!BANK_MODE = "C"
                ElseIf CmbMode.ListIndex = 1 Then
                    rststock!BANK_MODE = "U"
                ElseIf CmbMode.ListIndex = 2 Then
                    rststock!BANK_MODE = "N"
                Else
                    rststock!check_flag = "C"
                End If
                rststock.Update
                GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = CmbMode.text
            End If
            rststock.Close
            Set rststock = Nothing
            GRDTranx.Enabled = True
            CmbMode.Visible = False
            GRDTranx.SetFocus
        Case vbKeyEscape
            CmbMode.Visible = False
            GRDTranx.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmbMode_LostFocus()
    CmbMode.Visible = False
End Sub
