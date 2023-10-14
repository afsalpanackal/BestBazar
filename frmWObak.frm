VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmWO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESTIMATE...."
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12930
   Icon            =   "frmWO.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   12930
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   13155
      TabIndex        =   64
      Top             =   3090
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1482
      _Version        =   393216
   End
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   2205
      TabIndex        =   50
      Top             =   3630
      Visible         =   0   'False
      Width           =   6030
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   2835
         Left            =   75
         TabIndex        =   51
         Top             =   120
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   5001
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   2205
      TabIndex        =   46
      Top             =   3645
      Visible         =   0   'False
      Width           =   5835
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2535
         Left            =   90
         TabIndex        =   49
         Top             =   360
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   "BATCH WISE LIST FOR THE ITEM "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   9
         Left            =   90
         TabIndex        =   48
         Top             =   105
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   "MEDICINE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Index           =   0
         Left            =   3135
         TabIndex        =   47
         Top             =   105
         Visible         =   0   'False
         Width           =   2610
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   -135
      TabIndex        =   16
      Top             =   -15
      Width           =   13005
      Begin VB.TextBox txtexpirydate 
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
         Height          =   285
         Left            =   13965
         MaxLength       =   15
         TabIndex        =   54
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H0080FF80&
         Height          =   1800
         Left            =   210
         TabIndex        =   17
         Top             =   -60
         Width           =   7575
         Begin VB.TextBox TxtBillName 
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
            Left            =   3840
            MaxLength       =   35
            TabIndex        =   0
            Top             =   150
            Width           =   3645
         End
         Begin VB.TextBox txtBillNo 
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
            Left            =   1305
            TabIndex        =   14
            Top             =   210
            Visible         =   0   'False
            Width           =   885
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   345
            Left            =   1305
            TabIndex        =   55
            Top             =   615
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSForms.TextBox TxtBillAddress 
            Height          =   1230
            Left            =   3840
            TabIndex        =   97
            Top             =   510
            Width           =   3645
            VariousPropertyBits=   746604571
            MaxLength       =   100
            BorderStyle     =   1
            Size            =   "6429;2170"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            Index           =   3
            Left            =   2835
            TabIndex        =   92
            Top             =   255
            Width           =   975
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "INV DATE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   8
            Left            =   135
            TabIndex        =   56
            Top             =   675
            Width           =   885
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE NO."
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
            Left            =   105
            TabIndex        =   21
            Top             =   255
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Date"
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
            Index           =   1
            Left            =   135
            TabIndex        =   20
            Top             =   1050
            Width           =   1170
         End
         Begin VB.Label LBLDATE 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1320
            TabIndex        =   19
            Top             =   1065
            Width           =   1410
         End
         Begin VB.Label LBLBILLNO 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1305
            TabIndex        =   18
            Top             =   210
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080FF80&
         Height          =   6210
         Left            =   210
         TabIndex        =   22
         Top             =   1650
         Width           =   12795
         Begin VB.TextBox TXTAMOUNT 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   855
            TabIndex        =   84
            Top             =   5820
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox TXTTOTALDISC 
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
            Left            =   9840
            TabIndex        =   83
            Top             =   5805
            Width           =   930
         End
         Begin VB.OptionButton OPTDISCPERCENT 
            BackColor       =   &H0080FF80&
            Caption         =   "Discount %"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6690
            TabIndex        =   82
            Top             =   5805
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton OptDiscAmt 
            BackColor       =   &H0080FF80&
            Caption         =   "Discount Amt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   8190
            TabIndex        =   81
            Top             =   5805
            Width           =   1605
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0080FF80&
            Height          =   6060
            Left            =   10935
            TabIndex        =   23
            Top             =   120
            Width           =   1815
            Begin VB.Label LBLDISCAMT 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00800080&
               Height          =   435
               Left            =   195
               TabIndex        =   86
               Top             =   2535
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "DISC AMOUNT"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   375
               Index           =   4
               Left            =   165
               TabIndex        =   85
               Top             =   2280
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Tax On Free"
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
               Height          =   375
               Index           =   31
               Left            =   165
               TabIndex        =   80
               Top             =   1545
               Width           =   1515
            End
            Begin VB.Label LBLFOT 
               Alignment       =   2  'Center
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
               ForeColor       =   &H000000FF&
               Height          =   465
               Left            =   195
               TabIndex        =   79
               Top             =   1755
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "SELLING PRICE"
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
               Height          =   375
               Index           =   28
               Left            =   180
               TabIndex        =   70
               Top             =   4365
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ITEM COST"
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
               Height          =   375
               Index           =   27
               Left            =   210
               TabIndex        =   69
               Top             =   3795
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label LBLSELPRICE 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   180
               TabIndex        =   67
               Top             =   4605
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label LBLITEMCOST 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   180
               TabIndex        =   66
               Top             =   4035
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "RETAIL BILL"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   690
               Index           =   26
               Left            =   150
               TabIndex        =   62
               Top             =   3390
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label LBLPROFIT 
               Alignment       =   2  'Center
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
               ForeColor       =   &H000000FF&
               Height          =   495
               Left            =   180
               TabIndex        =   61
               Top             =   3555
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label LBLTOTALCOST 
               Alignment       =   2  'Center
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
               ForeColor       =   &H000000FF&
               Height          =   465
               Left            =   180
               TabIndex        =   60
               Top             =   3255
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL COST"
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
               Height          =   375
               Index           =   25
               Left            =   150
               TabIndex        =   59
               Top             =   2970
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "NET AMOUNT"
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
               Height          =   375
               Index           =   23
               Left            =   150
               TabIndex        =   58
               Top             =   810
               Width           =   1515
            End
            Begin VB.Label lblnetamount 
               Alignment       =   2  'Center
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
               ForeColor       =   &H000000FF&
               Height          =   480
               Left            =   180
               TabIndex        =   57
               Top             =   1020
               Width           =   1440
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "GROSS AMOUNT"
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
               Height          =   375
               Index           =   6
               Left            =   150
               TabIndex        =   25
               Top             =   120
               Width           =   1755
            End
            Begin VB.Label LBLTOTAL 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label2"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   480
               Left            =   195
               TabIndex        =   24
               Top             =   330
               Width           =   1440
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   5520
            Left            =   90
            TabIndex        =   15
            Top             =   225
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9737
            _Version        =   393216
            Rows            =   1
            Cols            =   25
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            GridLineWidth   =   2
         End
      End
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   465
         Left            =   11100
         TabIndex        =   45
         Top             =   5130
         Visible         =   0   'False
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   820
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0080FF80&
         Height          =   1545
         Left            =   210
         TabIndex        =   26
         Top             =   7770
         Width           =   12795
         Begin VB.TextBox TxtRetailmode 
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
            Height          =   300
            Left            =   4995
            MaxLength       =   6
            TabIndex        =   96
            Top             =   780
            Width           =   795
         End
         Begin VB.TextBox txtretaildummy 
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
            Height          =   300
            Left            =   4065
            MaxLength       =   6
            TabIndex        =   95
            Top             =   765
            Width           =   795
         End
         Begin VB.TextBox txtcommi 
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
            Height          =   300
            Left            =   10620
            MaxLength       =   6
            TabIndex        =   93
            Top             =   450
            Width           =   825
         End
         Begin VB.TextBox TXTSALETYPE 
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
            Height          =   300
            Left            =   3045
            MaxLength       =   6
            TabIndex        =   91
            Top             =   780
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtretail 
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
            Height          =   300
            Left            =   7440
            MaxLength       =   6
            TabIndex        =   90
            Top             =   450
            Width           =   945
         End
         Begin VB.TextBox TXTRETAILNOTAX 
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
            Height          =   300
            Left            =   6540
            MaxLength       =   6
            TabIndex        =   89
            Top             =   450
            Width           =   885
         End
         Begin VB.OptionButton optnet 
            BackColor       =   &H0000FFFF&
            Caption         =   "Net"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   9540
            TabIndex        =   88
            Top             =   780
            Width           =   1005
         End
         Begin VB.TextBox txtmrpbt 
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
            Height          =   300
            Left            =   2115
            MaxLength       =   6
            TabIndex        =   87
            Top             =   795
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H0000FFFF&
            Caption         =   "VAT %"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   8475
            TabIndex        =   77
            Top             =   780
            Width           =   1065
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H0000FFFF&
            Caption         =   "Tax on MRP %"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6630
            TabIndex        =   76
            Top             =   780
            Width           =   1875
         End
         Begin VB.TextBox TXTPTR 
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
            Height          =   300
            Left            =   11115
            MaxLength       =   6
            TabIndex        =   74
            Top             =   1095
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox TXTFREE 
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
            Height          =   300
            Left            =   4725
            MaxLength       =   7
            TabIndex        =   72
            Top             =   450
            Width           =   540
         End
         Begin VB.CommandButton cmdstockadjst 
            Caption         =   "Adjust Stock"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   555
            TabIndex        =   71
            Top             =   1035
            Width           =   1380
         End
         Begin VB.CommandButton cmdwoprint 
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
            Left            =   90
            TabIndex        =   68
            Top             =   1035
            Width           =   420
         End
         Begin VB.CommandButton CMDHIDE 
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
            Left            =   10335
            TabIndex        =   63
            Top             =   1035
            Width           =   420
         End
         Begin VB.TextBox TxtMRP 
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
            Height          =   300
            Left            =   5280
            MaxLength       =   6
            TabIndex        =   52
            Top             =   450
            Width           =   705
         End
         Begin VB.CommandButton CMDMODIFY 
            Caption         =   "&Modify Line"
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
            Left            =   2895
            TabIndex        =   8
            Top             =   1035
            Width           =   1125
         End
         Begin VB.TextBox TXTSLNO 
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
            Height          =   300
            Left            =   60
            TabIndex        =   1
            Top             =   450
            Width           =   480
         End
         Begin VB.TextBox TXTPRODUCT 
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
            Height          =   300
            Left            =   555
            TabIndex        =   2
            Top             =   450
            Width           =   3555
         End
         Begin VB.TextBox TXTQTY 
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
            Height          =   300
            Left            =   4125
            MaxLength       =   7
            TabIndex        =   3
            Top             =   450
            Width           =   585
         End
         Begin VB.TextBox TXTTAX 
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
            Height          =   300
            Left            =   6000
            MaxLength       =   4
            TabIndex        =   4
            Top             =   450
            Width           =   525
         End
         Begin VB.TextBox TXTDISC 
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
            Height          =   300
            Left            =   9960
            MaxLength       =   4
            TabIndex        =   6
            Top             =   450
            Width           =   645
         End
         Begin VB.CommandButton CMDPRINT 
            Caption         =   "&PRINT"
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
            Left            =   6495
            TabIndex        =   11
            Top             =   1035
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
            Height          =   405
            Left            =   8895
            TabIndex        =   13
            Top             =   1035
            Width           =   1125
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
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
            Height          =   405
            Left            =   5295
            TabIndex        =   10
            Top             =   1035
            Width           =   1125
         End
         Begin VB.CommandButton CmdDelete 
            Caption         =   "&Delete Line"
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
            Left            =   4095
            TabIndex        =   9
            Top             =   1035
            Width           =   1125
         End
         Begin VB.TextBox TXTITEMCODE 
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
            Height          =   300
            Left            =   2070
            TabIndex        =   31
            Top             =   1245
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox txtBatch 
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
            Height          =   300
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   5
            Top             =   450
            Width           =   1545
         End
         Begin VB.TextBox TXTVCHNO 
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
            Height          =   300
            Left            =   4035
            TabIndex        =   30
            Top             =   1260
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox TXTLINENO 
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
            Height          =   300
            Left            =   5895
            TabIndex        =   29
            Top             =   1260
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox TXTTRXTYPE 
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
            Height          =   300
            Left            =   7725
            TabIndex        =   28
            Top             =   1290
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox TXTUNIT 
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
            Height          =   300
            Left            =   9600
            TabIndex        =   27
            Top             =   1335
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
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
            Height          =   405
            Left            =   7695
            TabIndex        =   12
            Top             =   1035
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Commi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   5
            Left            =   10620
            TabIndex        =   94
            Top             =   225
            Width           =   825
         End
         Begin VB.Label lblP_Rate 
            Caption         =   "0"
            Height          =   390
            Left            =   10935
            TabIndex        =   78
            Top             =   960
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Lblprice 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "C Price"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   30
            Left            =   6540
            TabIndex        =   75
            Top             =   225
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Free"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   29
            Left            =   4725
            TabIndex        =   73
            Top             =   225
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   24
            Left            =   5280
            TabIndex        =   53
            Top             =   225
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "SL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   8
            Left            =   60
            TabIndex        =   44
            Top             =   225
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Product Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   9
            Left            =   555
            TabIndex        =   43
            Top             =   225
            Width           =   3555
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Qty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   10
            Left            =   4125
            TabIndex        =   42
            Top             =   225
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Net Rate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   11
            Left            =   7440
            TabIndex        =   41
            Top             =   225
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Tax%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   12
            Left            =   6000
            TabIndex        =   40
            Top             =   225
            Width           =   525
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Disc%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   13
            Left            =   9960
            TabIndex        =   39
            Top             =   225
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Sub Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   14
            Left            =   11460
            TabIndex        =   38
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "ITEM CODE."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   15
            Left            =   930
            TabIndex        =   37
            Top             =   1260
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Batch No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   7
            Left            =   8400
            TabIndex        =   36
            Top             =   225
            Width           =   1545
         End
         Begin VB.Label LBLSUBTOTAL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   11460
            TabIndex        =   7
            Top             =   450
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "VCH NO."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   17
            Left            =   2895
            TabIndex        =   35
            Top             =   1275
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "TRX TYPE."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   19
            Left            =   6585
            TabIndex        =   34
            Top             =   1305
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "UNIT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   20
            Left            =   8460
            TabIndex        =   33
            Top             =   1350
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "LINE NO."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   18
            Left            =   4755
            TabIndex        =   32
            Top             =   1275
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   65
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "frmWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim TMPREC As New ADODB.Recordset
Dim TMPFLAG As Boolean
Dim ACT_REC As New ADODB.Recordset
Dim ACT_AGNT As New ADODB.Recordset

Dim ACT_FLAG As Boolean
Dim AGNT_FLAG As Boolean
Dim PHY_BATCH As New ADODB.Recordset
Dim BATCH_FLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean

Dim CLOSEALL As Integer
Dim M_STOCK As Integer
Dim cr_days As Boolean
Dim M_EDIT As Boolean
Dim EDIT_INV As Boolean

Private Sub cmdstockadjst_Click()
    FrmStkAdj.Show
    FrmStkAdj.SetFocus
End Sub

Private Sub CMDADD_Click()
    Dim rststock As ADODB.Recordset
    'Dim RSTMINQTY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Format(Val(TxtMRP.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTRETAILNOTAX.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Format(Val(txtretail.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(TXTDISC.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Val(TXTTAX.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = ""
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Trim(TXTTRXTYPE.Text)
    
    If OPTVAT.Value = True And Val(TXTTAX.Text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "V"
    If OPTTaxMRP.Value = True And Val(TXTTAX.Text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "M"
    If Val(TXTTAX.Text) <= 0 Or optnet.Value = True Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "N"
  
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = IIf(IsNull(RSTTRXFILE!MANUFACTURER), "", Trim(RSTTRXFILE!MANUFACTURER))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = "B"
    grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = Val(TXTFREE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 21) = Format(Val(txtretail.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 22) = Format(Val(TXTRETAILNOTAX.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 23) = Trim(TXTSALETYPE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = Val(txtcommi.Text)
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            '!ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
            !FREE_QTY = !FREE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            !CLOSE_QTY = !CLOSE_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & " AND BAL_QTY > 0", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
            
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    optnet.Value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Text = ""
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    TxtRetailmode.Text = ""
    txtretail.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtBatch.Text = ""
    LBLSUBTOTAL.Caption = ""
    lblP_Rate.Caption = "0"
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    M_EDIT = False
    EDIT_INV = True
    Call COSTCALCULATION
    TXTSLNO.SetFocus
    grdsales.TopRow = grdsales.Rows - 1
Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            txtcommi.Enabled = True
            txtcommi.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
            !FREE_QTY = !FREE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
      
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
    
    For i = Val(TXTSLNO.Text) - 1 To grdsales.Rows - 2
        grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = i
        grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(i + 1, 6)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = grdsales.TextMatrix(i + 1, 7)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = grdsales.TextMatrix(i + 1, 8)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = grdsales.TextMatrix(i + 1, 9)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = grdsales.TextMatrix(i + 1, 10)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = grdsales.TextMatrix(i + 1, 11)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = grdsales.TextMatrix(i + 1, 12)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = grdsales.TextMatrix(i + 1, 13)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = grdsales.TextMatrix(i + 1, 14)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = grdsales.TextMatrix(i + 1, 15)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = grdsales.TextMatrix(i + 1, 16)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = grdsales.TextMatrix(i + 1, 17)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = grdsales.TextMatrix(i + 1, 18)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = grdsales.TextMatrix(i + 1, 19)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = grdsales.TextMatrix(i + 1, 20)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 21) = grdsales.TextMatrix(i + 1, 21)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 22) = grdsales.TextMatrix(i + 1, 22)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 23) = grdsales.TextMatrix(i + 1, 23)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = grdsales.TextMatrix(i + 1, 24)
    Next i
    grdsales.Rows = grdsales.Rows - 1
    
    Call COSTCALCULATION
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    optnet.Value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    
    txtretail.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    TxtRetailmode.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtBatch.Text = ""
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    EDIT_INV = True
    If grdsales.Rows = 1 Then
'        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    
End Sub

Private Sub CmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            
            txtretail.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.Value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            TxtRetailmode.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDHIDE_Click()
    If LBLPROFIT.Visible = True Then
        LBLPROFIT.Visible = False
        LBLTOTALCOST.Visible = False
        Label1(25).Visible = False
        Label1(26).Visible = False
        Label1(27).Visible = False
        Label1(28).Visible = False
        LBLITEMCOST.Visible = False
        LBLSELPRICE.Visible = False
    Else
        LBLPROFIT.Visible = True
        LBLTOTALCOST.Visible = True
        Label1(25).Visible = True
        Label1(26).Visible = True
        Label1(27).Visible = True
        Label1(28).Visible = True
        LBLITEMCOST.Visible = True
        LBLSELPRICE.Visible = True
    End If
End Sub

Private Sub CMDMODIFY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    On Error GoTo eRRHAND
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            '!ISSUE_QTY = !ISSUE_QTY - ((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))))
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
            !FREE_QTY = !FREE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = True
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            
            txtretail.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.Value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            TxtRetailmode.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub cmdPrint_Click()
    
    If grdsales.Rows = 1 Then Exit Sub
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bil..."
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bil..."
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    End If
    
    Generateprint
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Integer
    Dim CN As Integer
    Dim DN As Integer
    Dim B As Integer
    Dim num As Currency
    
    'On Error GoTo eRRHAND
    DN = 0
    CN = 0
    B = 0
    
    db.Execute "delete * From TRXFILE WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "SI"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!CATEGORY = "MEDICINE"
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 13)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 5))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "Issued as     " & "ESTIMATE"
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 10)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!COM_AMT = Trim(grdsales.TextMatrix(i, 24))
        
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "131001"
        RSTTRXFILE!Area = ""
        
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Call ReportGeneratION
    
    On Error GoTo CLOSEFILE
    Open App.Path & "\repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open App.Path & "\repo.bat" For Output As #1 '//Creating Batch file
    End If
    On Error GoTo eRRHAND
    
    Print #1, "TYPE " & App.Path & "\Report.txt > PRN"
    Print #1, "EXIT"
    Close #1
    
    '//HERE write the proper path where your command.com file exist
    'Shell "C:\WINDOW\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
    Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
    'If MsgBox("DO YOU WANT ONE MORE PRINT...", vbYesNo, "BILL..") = vbYes Then Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide

    ''cmdRefresh.SetFocus

'    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption), 2)) - Val(Round(Val(LBLTOTAL.Caption), 0)), 2)
'
    
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    txtretail.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    txtcommi.Enabled = False
    
    ''rptPRINT.Action = 1
    Exit Function
eRRHAND:
    MsgBox Err.Description
End Function

Private Sub CMDPRINT_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            
            txtretail.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.Value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            TxtRetailmode.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub cmdRefresh_Click()
    
   ' If grdsales.Rows = 1 Then GoTo SKIP
    
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bil..."
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bil..."
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    End If

    Call AppendSale
'    Me.Enabled = False
'    FRMDEBIT.Show
    
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            
            txtretail.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.Value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            TxtRetailmode.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub Form_Activate()
    If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If TxtMRP.Enabled = True Then TxtMRP.SetFocus
    If txtretail.Enabled = True Then txtretail.SetFocus
    If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
    If TXTTAX.Enabled = True Then TXTTAX.SetFocus
    If txtBatch.Enabled = True Then txtBatch.SetFocus
    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
    If txtcommi.Enabled = True Then txtcommi.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo eRRHAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(VCH_NO)) From TRXMAST WHERE TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    ACT_FLAG = True
    AGNT_FLAG = True
    lblP_Rate.Caption = "0"
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2000
    grdsales.ColWidth(3) = 500
    grdsales.ColWidth(4) = 500
    grdsales.ColWidth(5) = 700
    grdsales.ColWidth(6) = 600
    grdsales.ColWidth(7) = 600
    grdsales.ColWidth(8) = 500
    grdsales.ColWidth(9) = 600
    grdsales.ColWidth(10) = 800
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 1000
    grdsales.ColWidth(21) = 600
    grdsales.ColWidth(22) = 0
    grdsales.ColWidth(24) = 600
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "UNIT"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "PTR"
    grdsales.TextArray(7) = "RATE"
    grdsales.TextArray(8) = "DISC %"
    grdsales.TextArray(9) = "TAX %"
    grdsales.TextArray(10) = "Serial No"
    grdsales.TextArray(11) = ""
    grdsales.TextArray(12) = "SUB TOTAL"
    grdsales.TextArray(13) = "ITEM CODE"
    grdsales.TextArray(14) = "Vch No"
    grdsales.TextArray(15) = "Line No"
    grdsales.TextArray(16) = "Trx Type"
    grdsales.TextArray(17) = "Tax Mode"
    grdsales.TextArray(18) = "MFGR"
    grdsales.TextArray(19) = "CN/DN"
    grdsales.TextArray(20) = "Free"
    grdsales.TextArray(21) = "PTR"
    grdsales.TextArray(22) = "PTRWOTAX"
    grdsales.TextArray(24) = "Comm"
    'grdsales.ColWidth(12) = 0
    'grdsales.ColWidth(13) = 0
    'grdsales.ColWidth(14) = 0
   'grdsales.ColWidth(15) = 0
    'grdsales.ColWidth(16) = 0
    
    LBLTOTAL.Caption = 0
    
    PHYFLAG = True
    TMPFLAG = True
    BATCH_FLAG = True
    ITEM_FLAG = True
    cr_days = False
    
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TxtMRP.Enabled = False
    
    txtretail.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    txtcommi.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    TXTSLNO.Text = 1
    TXTSLNO.Enabled = False
    CLOSEALL = 1
    M_EDIT = False
    EDIT_INV = False
'    Me.Width = 11700
'    Me.Height = 10185
    Me.Left = 0
    Me.Top = 0
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If TMPFLAG = False Then TMPREC.Close
        If BATCH_FLAG = False Then PHY_BATCH.Close
        If ITEM_FLAG = False Then PHY_ITEM.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If AGNT_FLAG = False Then ACT_AGNT.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDPOPUP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTtax As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            'TXTQTY.Text = GRDPOPUP.Columns(1)
            TxtMRP.Text = GRDPOPUP.Columns(3)
            TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(13)), "", GRDPOPUP.Columns(13))
            If GRDPOPUP.Columns(14) = "A" Then
                txtretaildummy.Text = IIf(IsNull(GRDPOPUP.Columns(16)), "P", GRDPOPUP.Columns(16))
                TxtRetailmode.Text = "A"
            Else
                txtretaildummy.Text = IIf(IsNull(GRDPOPUP.Columns(15)), "P", GRDPOPUP.Columns(15))
                TxtRetailmode.Text = "P"
            End If
            
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & GRDPOPUP.Columns(6) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            With RSTtax
                If Not (.EOF And .BOF) Then
                    Select Case GRDPOPUP.Columns(12)
                        Case "M"
                            OPTTaxMRP.Value = True
                            TXTTAX.Text = GRDPOPUP.Columns(5)
                            TXTSALETYPE.Text = "2"
                        Case "V"
                            If (!CATEGORY = "MEDICINE" And !Remarks = "1") Then
                                OPTTaxMRP.Value = True
                                TXTSALETYPE.Text = "1"
                            Else
                                OPTVAT.Value = True
                                TXTSALETYPE.Text = "2"
                            End If
                            TXTTAX.Text = GRDPOPUP.Columns(5)
                        Case Else
                            TXTSALETYPE.Text = "2"
                            optnet.Value = True
                            TXTTAX.Text = "0"
                    End Select
                Else
                    optnet.Value = True
                    TXTTAX.Text = "0"
                End If
            End With
            RSTtax.Close
            Set RSTtax = Nothing
            
            txtBatch.Text = GRDPOPUP.Columns(0)
            
            TXTVCHNO.Text = GRDPOPUP.Columns(8)
            TXTLINENO.Text = GRDPOPUP.Columns(9)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(10)
            TXTUNIT.Text = GRDPOPUP.Columns(11)
                        
            Set GRDPOPUP.DataSource = Nothing
            
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            
            Set GRDPOPUP.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
        
    End Select
End Sub

Private Sub GRDPOPUPITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim NONSTOCKFLAG As Boolean
    Dim MINUSFLAG As Boolean
    Dim i As Integer
    
    On Error GoTo eRRHAND
    Select Case KeyCode
        Case vbKeyReturn
            NONSTOCKFLAG = False
            MINUSFLAG = False
            M_STOCK = Val(GRDPOPUPITEM.Columns(2))
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            i = 0
            If M_STOCK <= 0 Then
''''''                Set RSTNONSTOCK = New ADODB.Recordset
''''''                RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
''''''                i = RSTNONSTOCK.RecordCount
''''''                RSTNONSTOCK.Close
''''''                Set RSTNONSTOCK = Nothing
''''''                If i = 0 Then
''''''                    If (MsgBox("NO STOCK AVAILABLE..Do you want to add to Stockless", vbYesNo, "SALES") = vbYes) Then
''''''                        Set RSTNONSTOCK = New ADODB.Recordset
''''''                        RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
''''''                        If (RSTNONSTOCK.EOF And RSTNONSTOCK.BOF) Then
''''''                            RSTNONSTOCK.AddNew
''''''                            RSTNONSTOCK!ITEM_NAME = TXTPRODUCT.Text
''''''                            RSTNONSTOCK!ITEM_CODE = TXTITEMCODE.Text
''''''                            RSTNONSTOCK!Date = Date & " " & Time
''''''                            RSTNONSTOCK.Update
''''''                        End If
''''''                        RSTNONSTOCK.Close
''''''                        Set RSTNONSTOCK = Nothing
''''''                    End If
''''''                    Exit Sub
''''''                End If
                
                MsgBox "AVAILABLE STOCK IS  " & M_STOCK & "", vbOKOnly, "BILL.."
                Exit Sub
                If (MsgBox("AVAILABLE STOCK IS  " & M_STOCK & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                    Exit Sub
                Else
                    MINUSFLAG = True
                End If
                NONSTOCKFLAG = True
            End If
            For i = 1 To grdsales.Rows - 1
                If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists.... Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then
                        Set GRDPOPUPITEM.DataSource = Nothing
                        FRMEITEM.Visible = False
                        FRMEMAIN.Enabled = True
                        TXTPRODUCT.Enabled = True
                        TXTQTY.Enabled = False
                        TXTPRODUCT.SetFocus
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next i
            Set GRDPOPUPITEM.DataSource = Nothing
            If ITEM_FLAG = True Then
                If NONSTOCKFLAG = True Then
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY [VCH_DATE]", db, adOpenStatic, adLockReadOnly
                Else
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db, adOpenStatic, adLockReadOnly
                End If
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                If NONSTOCKFLAG = True Then
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY [VCH_DATE]", db, adOpenStatic, adLockReadOnly
                Else
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db, adOpenStatic, adLockReadOnly
                End If
                ITEM_FLAG = False
            End If
            Set GRDPOPUPITEM.DataSource = PHY_ITEM
            If PHY_ITEM.RecordCount = 0 Then
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            End If
            If PHY_ITEM.RecordCount = 1 Or MINUSFLAG = True Then
                'TXTQTY.Text = GRDPOPUPITEM.Columns(2)
                TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", GRDPOPUPITEM.Columns(13))
                TxtMRP.Text = GRDPOPUPITEM.Columns(11)
                If GRDPOPUPITEM.Columns(14) = "A" Then
                    txtretaildummy.Text = IIf(IsNull(GRDPOPUPITEM.Columns(16)), "P", GRDPOPUPITEM.Columns(16))
                    TxtRetailmode.Text = "A"
                Else
                    txtretaildummy.Text = IIf(IsNull(GRDPOPUPITEM.Columns(15)), "P", GRDPOPUPITEM.Columns(15))
                    TxtRetailmode.Text = "P"
                End If
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                With RSTtax
                    If Not (.EOF And .BOF) Then
                        Select Case PHY_ITEM!CHECK_FLAG
                            Case "M"
                                OPTTaxMRP.Value = True
                                TXTTAX.Text = GRDPOPUPITEM.Columns(4)
                                TXTSALETYPE.Text = "2"
                            Case "V"
                                If (!CATEGORY = "MEDICINE" And !Remarks = "1") Then
                                'If !CATEGORY = "MEDICINE" Then
                                    OPTTaxMRP.Value = True
                                    TXTSALETYPE.Text = "1"
                                Else
                                    OPTVAT.Value = True
                                    TXTSALETYPE.Text = "2"
                                End If
                                TXTTAX.Text = GRDPOPUPITEM.Columns(4)
                            Case Else
                                TXTSALETYPE.Text = "2"
                                optnet.Value = True
                                TXTTAX.Text = "0"
                        End Select
                    Else
                        optnet.Value = True
                        TXTTAX.Text = "0"
                    End If
                End With
                RSTtax.Close
                Set RSTtax = Nothing
            
                'TXTTAX.Text = 0 'GRDPOPUPITEM.Columns(4)
                txtBatch.Text = GRDPOPUPITEM.Columns(6)
                
                TXTVCHNO.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(8), "")
                TXTLINENO.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(9), "")
                TXTTRXTYPE.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(10), "")
                TXTUNIT.Text = GRDPOPUPITEM.Columns(5)
                            
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            ElseIf PHY_ITEM.RecordCount > 1 And MINUSFLAG = False Then
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEGRDTMP.Visible = False
                Call FILL_BATCHGRID
            End If
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
            
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub optnet_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OPTTaxMRP_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OPTVAT_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtBatch.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyEscape
            txtBatch.Enabled = False
            txtretail.Enabled = True
            txtretail.SetFocus
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtBillAddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            FRMEHEAD.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TxtBillName.Enabled = True
            TxtBillName.SetFocus
    End Select
End Sub

Private Sub TxtBillAddress_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
    cr_days = False
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TRXMAST As ADODB.Recordset
    Dim TRXSUB As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    
    Dim i As Integer
    Dim N As Integer
    Dim M As Integer

    On Error GoTo eRRHAND

    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtBillNo.Text) = 0 Then Exit Sub
            grdsales.Rows = 1
            i = 0
            Set TRXSUB = New ADODB.Recordset
            TRXSUB.Open "Select * From TRXSUB WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until TRXSUB.EOF
                Set TRXFILE = New ADODB.Recordset
                TRXFILE.Open "Select * From TRXFILE WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(TRXSUB!LINE_NO) & "", db, adOpenStatic, adLockReadOnly
                If Not (TRXFILE.EOF And TRXFILE.BOF) Then
                    i = i + 1
                    LBLDATE.Caption = Format(TRXFILE!VCH_DATE, "DD/MM/YYYY")
                    grdsales.Rows = grdsales.Rows + 1
                    grdsales.FixedRows = 1
                    grdsales.TextMatrix(i, 0) = i
                    grdsales.TextMatrix(i, 1) = TRXFILE!ITEM_CODE
                    grdsales.TextMatrix(i, 2) = TRXFILE!ITEM_NAME
                    grdsales.TextMatrix(i, 3) = TRXFILE!QTY
                    Set TRXMAST = New ADODB.Recordset
                    TRXMAST.Open "SELECT UNIT FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(TRXSUB!R_TRX_TYPE) & "' AND RTRXFILE.VCH_NO = " & Val(TRXSUB!R_VCH_NO) & " AND RTRXFILE.LINE_NO = " & Val(TRXSUB!R_LINE_NO) & "", db, adOpenStatic, adLockReadOnly
                    If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                        grdsales.TextMatrix(i, 4) = Val(TRXMAST!UNIT)
                    End If
                    TRXMAST.Close
                    Set TRXMAST = Nothing
                    
                    Set TRXMAST = New ADODB.Recordset
                    TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMAST WHERE ITEMMAST.ITEM_CODE = '" & Trim(TRXFILE!ITEM_CODE) & "'", db, adOpenStatic, adLockReadOnly
                    If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                        grdsales.TextMatrix(i, 18) = IIf(IsNull(TRXMAST!MANUFACTURER), "", Trim(TRXMAST!MANUFACTURER))
                    End If
                    TRXMAST.Close
                    Set TRXMAST = Nothing
                    
                    grdsales.TextMatrix(i, 5) = Format(TRXFILE!MRP, ".000")
                    grdsales.TextMatrix(i, 6) = Format(TRXFILE!PTR, ".000")
                    grdsales.TextMatrix(i, 7) = Format(TRXFILE!SALES_PRICE, ".000")
                    grdsales.TextMatrix(i, 8) = 0 'DISC
                    grdsales.TextMatrix(i, 9) = Val(TRXFILE!SALES_TAX)
            
                    grdsales.TextMatrix(i, 10) = TRXFILE!REF_NO
                    grdsales.TextMatrix(i, 11) = ""
                    grdsales.TextMatrix(i, 12) = Format(Val(TRXFILE!TRX_TOTAL), ".000")
                    
                    grdsales.TextMatrix(i, 13) = TRXFILE!ITEM_CODE
                    grdsales.TextMatrix(i, 14) = Val(TRXSUB!R_VCH_NO)
                    grdsales.TextMatrix(i, 15) = Val(TRXSUB!R_LINE_NO)
                    grdsales.TextMatrix(i, 16) = Trim(TRXSUB!R_TRX_TYPE)
                    grdsales.TextMatrix(i, 17) = IIf(IsNull(TRXFILE!CHECK_FLAG), "", Trim(TRXFILE!CHECK_FLAG))
                    TXTINVDATE.Text = IIf(IsNull(TRXFILE!VCH_DATE), Date, TRXFILE!VCH_DATE)
                    grdsales.TextMatrix(i, 19) = "B"
                    grdsales.TextMatrix(i, 20) = TRXFILE!FREE_QTY
                    grdsales.TextMatrix(i, 21) = IIf(IsNull(TRXFILE!P_RETAIL), "0.00", Format(TRXFILE!P_RETAIL, ".000"))
                    grdsales.TextMatrix(i, 22) = IIf(IsNull(TRXFILE!P_RETAILWOTAX), "0.00", Format(TRXFILE!P_RETAILWOTAX, ".000"))
                    grdsales.TextMatrix(i, 23) = IIf(IsNull(TRXFILE!SALE_1_FLAG), "2", TRXFILE!SALE_1_FLAG)
                    grdsales.TextMatrix(i, 24) = IIf(IsNull(TRXFILE!COM_AMT), "2", TRXFILE!COM_AMT)
                    cr_days = True
                    'txtBillNo.Text = ""
                    'LBLBILLNO.Caption = ""
                End If
                TRXFILE.Close
                Set TRXFILE = Nothing
                TRXSUB.MoveNext
            Loop
            TRXSUB.Close
            Set TRXSUB = Nothing
            
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "Select * From TRXMAST WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                If TRXMAST!SLSM_CODE = "A" Then
                    TXTTOTALDISC.Text = IIf(IsNull(TRXMAST!DISCOUNT), "", Round((TRXMAST!DISCOUNT * TRXMAST!VCH_AMOUNT) / 100, 2))
                    OptDiscAmt.Value = True
                ElseIf TRXMAST!SLSM_CODE = "P" Then
                    TXTTOTALDISC.Text = IIf(IsNull(TRXMAST!DISCOUNT), "", TRXMAST!DISCOUNT)
                    OPTDISCPERCENT.Value = True
                End If
                TxtBillName.Text = IIf(IsNull(TRXMAST!BILL_NAME), "", TRXMAST!BILL_NAME)
                TxtBillAddress.Text = IIf(IsNull(TRXMAST!BILL_ADDRESS), "", TRXMAST!BILL_ADDRESS)
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            LBLBILLNO.Caption = Val(txtBillNo.Text)
            
            
            LBLTOTAL.Caption = ""
            lblnetamount.Caption = ""
            LBLFOT.Caption = ""
            For i = 1 To grdsales.Rows - 1
                grdsales.TextMatrix(i, 0) = i
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
            Next i
            LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
            TXTAMOUNT.Text = ""
            If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
                TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
            ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
                TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
            End If
            LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
            lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
            
            Call COSTCALCULATION
            
            
            TXTSLNO.Text = grdsales.Rows
            txtBillNo.Visible = False
            TXTSLNO.Enabled = True
            
            If grdsales.Rows > 1 Then
                TXTSLNO.SetFocus
            Else
                TXTSLNO.Enabled = False
                TxtBillName.Text = ""
                TxtBillName.SetFocus
            End If
    
    End Select

    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTBILLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBillNo_LostFocus()
    Dim TRXMAST As ADODB.Recordset
    Dim i As Integer

    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(Val(VCH_NO)) From TRXMAST WHERE TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        If Val(txtBillNo.Text) > i Then
            MsgBox "The last bill No. is " & i, vbCritical, "BILL..."
            txtBillNo.Visible = True
            txtBillNo.SetFocus
            Exit Sub
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
      
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MIN(Val(VCH_NO)) From TRXFILE WHERE TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
        If Val(txtBillNo.Text) < i Then
            MsgBox "This Year Starting Bill No. is " & i, vbCritical, "BILL..."
            txtBillNo.Visible = True
            txtBillNo.SetFocus
            Exit Sub
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    txtBillNo.Visible = False
    Call TXTBILLNO_KeyDown(13, 0)
    Exit Sub
End Sub

Private Sub TXTDISC_GotFocus()
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.Text)
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcommi.Enabled = True
            TXTDISC.Enabled = False
            txtcommi.SetFocus
        Case vbKeyEscape
            txtBatch.Enabled = True
            TXTDISC.Enabled = False
            txtBatch.SetFocus
    End Select
End Sub

Private Sub TXTDISC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDISC_LostFocus()

    ''TXTDISC.Text = Format(TXTDISC.Text, ".000")
    TXTDISC.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(txtretail.Text) * Val(TXTDISC.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(txtretail.Text), 3)) - Val(TXTDISC.Tag), ".000")

End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.Text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.Text = "  /  /    " Then
                TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
                TxtBillName.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
                TxtBillName.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                TxtBillName.SetFocus
            End If
        Case vbKeyEscape
            If EDIT_INV = True Then Exit Sub
            txtBillNo.Visible = True
            txtBillNo.SetFocus
    End Select
End Sub

Private Sub TXTINVDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTMRP_GotFocus()
    TxtMRP.SelStart = 0
    TxtMRP.SelLength = Len(TxtMRP.Text)
End Sub

Private Sub TXTMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtMRP.Text) = 0 Then Exit Sub
            TxtMRP.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
        Case vbKeyEscape
            TxtMRP.Enabled = False
            TXTFREE.Enabled = True
            TXTFREE.SetFocus
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

Private Sub TxtMRP_LostFocus()
    TxtMRP.Text = Format(TxtMRP.Text, ".000")
End Sub

Private Sub TXTPRODUCT_GotFocus()
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim RSTZEROSTOCK As ADODB.Recordset
    Dim RSTBALQTY As ADODB.Recordset
    
    On Error GoTo eRRHAND
    Select Case KeyCode
        Case 106
            If TXTQTY.Tag <> "" Then
                TXTPRODUCT.Text = Trim(TXTQTY.Tag)
                TXTPRODUCT.SelStart = 0
                TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
            End If
        Case vbKeyReturn
            M_STOCK = 0
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            CmdDelete.Enabled = False
            TXTQTY.Text = ""
            txtretail.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.Value = True
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            txtBatch.Text = ""
            LBLSUBTOTAL.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
           
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                Next i
                Set grdtmp.DataSource = Nothing
                If TMPFLAG = True Then
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                Else
                    TMPREC.Close
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                End If
                
                Set grdtmp.DataSource = TMPREC
                If TMPREC.RecordCount = 1 Then
                    'TXTQTY.Text = grdtmp.Columns(2)
                    TxtMRP.Text = grdtmp.Columns(3)
                    TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(13)), "", grdtmp.Columns(13))
                    If grdtmp.Columns(14) = "A" Then
                        txtretaildummy.Text = IIf(IsNull(grdtmp.Columns(16)), "P", grdtmp.Columns(16))
                        TxtRetailmode.Text = "A"
                    Else
                        txtretaildummy.Text = IIf(IsNull(grdtmp.Columns(15)), "P", grdtmp.Columns(15))
                        TxtRetailmode.Text = "P"
                    End If
                    Set RSTtax = New ADODB.Recordset
                    RSTtax.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdtmp.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                    With RSTtax
                        If Not (.EOF And .BOF) Then
                            Select Case TMPREC!CHECK_FLAG
                                Case "M"
                                    OPTTaxMRP.Value = True
                                    TXTTAX.Text = grdtmp.Columns(5)
                                    TXTSALETYPE.Text = "2"
                                Case "V"
                                    If (!CATEGORY = "MEDICINE" And !Remarks = "1") Then
                                    'If !CATEGORY = "MEDICINE" Then
                                        TXTSALETYPE.Text = "1"
                                        OPTTaxMRP.Value = True
                                    Else
                                        OPTVAT.Value = True
                                        TXTSALETYPE.Text = "2"
                                    End If
                                    TXTTAX.Text = grdtmp.Columns(5)
                                Case Else
                                    TXTSALETYPE.Text = "2"
                                    optnet.Value = True
                                    TXTTAX.Text = "0"
                            End Select
                        Else
                            optnet.Value = True
                            TXTTAX.Text = "0"
                        End If
                    End With
                    RSTtax.Close
                    Set RSTtax = Nothing
                    
                    txtBatch.Text = grdtmp.Columns(7)
                    
                    TXTVCHNO.Text = grdtmp.Columns(9)
                    TXTLINENO.Text = grdtmp.Columns(10)
                    TXTTRXTYPE.Text = grdtmp.Columns(11)
                    TXTUNIT.Text = grdtmp.Columns(6)
                                        
                    TXTPRODUCT.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                    Exit Sub
                ElseIf TMPREC.RecordCount = 0 Then
                    Set RSTBALQTY = New ADODB.Recordset
                    RSTBALQTY.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                    With RSTBALQTY
                        If Not (.EOF And .BOF) Then
                            M_STOCK = !CLOSE_QTY
                        End If
                    End With
                    RSTBALQTY.Close
                    Set RSTBALQTY = Nothing
            
                    TXTQTY.Text = 0
                    i = 0
                    MsgBox "AVAILABLE STOCK IS  " & M_STOCK & "", vbOKOnly, "BILL.."
                    'MsgBox "NO STOCK...", vbOKOnly, "BILL.."
                    TXTPRODUCT.Enabled = True
                    TXTQTY.Enabled = False
                    TXTPRODUCT.SelStart = 0
                    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
                    TXTPRODUCT.SetFocus
                    Exit Sub
                ElseIf TMPREC.RecordCount > 1 Then
                    Call FILL_BATCHGRID
                    Exit Sub
                End If
JUMPNONSTOCK:
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                
                TXTTAX.Enabled = False
                txtBatch.Enabled = False
                TXTDISC.Enabled = False
                TXTQTY.SetFocus
                Exit Sub
            ElseIf PHY.RecordCount > 1 Then
                'FRMSUB.grdsub.Columns(0).Visible = True
                'FRMSUB.grdsub.Columns(1).Caption = "ITEM NAME"
                'FRMSUB.grdsub.Columns(1).Width = 3200
                'FRMSUB.grdsub.Columns(2).Caption = "QTY"
                'FRMSUB.grdsub.Columns(2).Width = 1300
                Call FILL_ITEMGRID
            End If

            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CmdDelete.Enabled = False
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTQTY_GotFocus()
    Dim RSTITEMCOST As ADODB.Recordset
    
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
    TXTQTY.Tag = Trim(TXTPRODUCT.Text)
    On Error GoTo eRRHAND
    
    Set RSTITEMCOST = New ADODB.Recordset
    RSTITEMCOST.Open "SELECT ITEM_COST, SALES_PRICE FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMCOST.EOF Or RSTITEMCOST.BOF) Then
        LBLITEMCOST.Caption = RSTITEMCOST!ITEM_COST
        LBLSELPRICE.Caption = RSTITEMCOST!SALES_PRICE
    End If
    RSTITEMCOST.Close
    Set RSTITEMCOST = Nothing
    
    
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            
            'If Val(TXTQTY.Text) = 0 Then Exit Sub
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT BAL_QTY  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                i = RSTTRXFILE!BAL_QTY
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            Set RSTTRXFILE = Nothing
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If i > 0 Then
                If Val(TXTQTY.Text) > i Then
                    MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                    TXTQTY.SelStart = 0
                    TXTQTY.SelLength = Len(TXTQTY.Text)
                    Exit Sub
                End If
            End If
SKIP:
            TXTQTY.Enabled = False
            TXTFREE.Enabled = True
            TXTFREE.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    TXTQTY.Text = Format(TXTQTY.Text, ".000")
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    If Val(TXTRETAILNOTAX.Text) = 0 Then
        TXTDISC.Tag = Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(txtretail.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    End If
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.Text) = 0 Then
                TXTSLNO.Text = ""
                TXTPRODUCT.Text = ""
                TXTQTY.Text = ""
                TXTFREE.Text = ""
                optnet.Value = True
                TxtMRP.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTUNIT.Text = ""
                LBLSUBTOTAL.Caption = ""
                txtBatch.Text = ""
                TXTSLNO.Text = grdsales.Rows
                CmdDelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.Text) >= grdsales.Rows Then
                TXTSLNO.Text = grdsales.Rows
                CmdDelete.Enabled = False
                CMDMODIFY.Enabled = False
            End If
            If Val(TXTSLNO.Text) < grdsales.Rows Then
                lblP_Rate.Caption = "1"
                TXTSLNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 0)
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 3)
                TXTFREE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 20)
                TxtMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), ".000")
                
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 15)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 16)
                TXTUNIT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                'TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 22)
                txtretaildummy.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 24)
                TxtRetailmode.Text = "A"
                txtcommi.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 24)
                LBLSUBTOTAL.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 12)
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                
                Select Case Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
                    Case "M"
                        OPTTaxMRP.Value = True
                    Case "V"
                        OPTVAT.Value = True
                    Case Else
                        optnet.Value = True
                End Select
                TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                txtretail.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                TXTSALETYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 23)
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                
                TXTTAX.Enabled = False
                TXTFREE.Enabled = False
                txtretail.Enabled = False
                TXTRETAILNOTAX.Enabled = False
                txtBatch.Enabled = False
                TXTDISC.Enabled = False
                TxtMRP.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True

                Exit Sub
            End If
SKIP:
            lblP_Rate.Caption = "0"
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                optnet.Value = True
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTUNIT.Text = ""
                TXTQTY.Text = ""
                
                txtretail.Text = ""
                TXTRETAILNOTAX.Text = ""
                TXTSALETYPE.Text = ""
                TXTFREE.Text = ""
                TxtMRP.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                txtBatch.Text = ""
                lblP_Rate.Caption = "0"
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.Rows > 1 Then
                TXTSLNO.Enabled = False
                CMDPRINT.Enabled = True
                cmdRefresh.Enabled = True
                CMDPRINT.SetFocus
            Else
                TXTSLNO.Enabled = False
                FRMEHEAD.Enabled = True
                TxtBillName.Enabled = True
                TxtBillName.SetFocus
            End If
    End Select
End Sub

Private Sub TXTSLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case vbKeyTab
            KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTTAX_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
End Sub

Private Sub TXTTAX_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRETAILNOTAX.Enabled = True
            TXTTAX.Enabled = False
            TXTRETAILNOTAX.SetFocus
        Case vbKeyEscape
            TxtMRP.Enabled = True
            TXTTAX.Enabled = False
            TxtMRP.SetFocus
    End Select
End Sub

Private Sub TXTTAX_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTTAX_LostFocus()
'    TXTDISC.Tag = 0
'    TXTTAX.Tag = 0
'    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
'    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
'    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    txtmrpbt.Text = 100 * Val(TxtMRP.Text) / (100 + Val(TXTTAX.Text))
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Format(Month(DateIn), "00") & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Function FILL_BATCHGRID()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "BATCH NO."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = ""
    GRDPOPUP.Columns(3).Caption = "MRP"
    GRDPOPUP.Columns(4).Caption = "RATE"
    GRDPOPUP.Columns(5).Caption = "TAX"
    GRDPOPUP.Columns(8).Caption = "VCH No"
    GRDPOPUP.Columns(9).Caption = "Line No"
    GRDPOPUP.Columns(10).Caption = "Trx Type"
    
    GRDPOPUP.Columns(0).Width = 1400
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 0
    GRDPOPUP.Columns(3).Width = 1000
    GRDPOPUP.Columns(4).Width = 1000
    GRDPOPUP.Columns(5).Width = 900
    
    GRDPOPUP.Columns(8).Visible = False
    GRDPOPUP.Columns(9).Visible = False
    GRDPOPUP.Columns(10).Visible = False
    
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(7).Text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
End Function

Function FILL_ITEMGRID()
    FRMEMAIN.Enabled = False
    FRMEITEM.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEGRDTMP.Visible = False
    
    
    If ITEM_FLAG = True Then
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    End If

    Set GRDPOPUPITEM.DataSource = PHY_ITEM
    GRDPOPUPITEM.RowHeight = 250
    GRDPOPUPITEM.Columns(0).Visible = False
    GRDPOPUPITEM.Columns(1).Caption = "ITEM NAME"
    GRDPOPUPITEM.Columns(1).Width = 3800
    GRDPOPUPITEM.Columns(2).Caption = "QTY"
    GRDPOPUPITEM.Columns(2).Width = 1300
    GRDPOPUPITEM.SetFocus
End Function

Private Function STOCKADJUST()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    M_STOCK = 0
    On Error GoTo eRRHAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rststock.EOF
        M_STOCK = M_STOCK + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTITEMMAST
        If Not (.EOF And .BOF) Then
'            !OPEN_QTY = M_STOCK
'            !OPEN_VAL = 0
'            !RCPT_QTY = 0
'            !RCPT_VAL = 0
'            !ISSUE_QTY = 0
'            !ISSUE_VAL = 0
            !CLOSE_QTY = M_STOCK
'            !CLOSE_VAL = 0
'            !DAM_QTY = 0
'            !DAM_VAL = 0
            RSTITEMMAST.Update
        End If
    End With
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Exit Function
    
eRRHAND:
    MsgBox Err.Description
End Function


Private Function COSTCALCULATION()
    Dim RSTCOST As ADODB.Recordset
    Dim COST As Double
    Dim N As Integer
    'Dim RSTITEMMAST As ADODB.Recordset
    
     LBLTOTALCOST.Caption = ""
     LBLPROFIT.Caption = ""
        COST = 0
    On Error GoTo eRRHAND
    For N = 1 To grdsales.Rows - 1
        Set RSTCOST = New ADODB.Recordset
        RSTCOST.Open "SELECT [ITEM_COST] FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(N, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(N, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(N, 15)) & "", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTCOST.EOF
            COST = COST + (RSTCOST!ITEM_COST) * Val(grdsales.TextMatrix(N, 3))
            RSTCOST.MoveNext
        Loop
        RSTCOST.Close
        Set RSTCOST = Nothing
    Next N
    
    LBLTOTALCOST.Caption = Round(COST, 2)
    LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - COST, 2)

    Exit Function
    
eRRHAND:
    MsgBox Err.Description
End Function

Private Function AppendSale()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTP_RATE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim i As Double
    Dim TRXVALUE As Double
    
    Dim DAY_DATE As String
    Dim MONTH_DATE As String
    Dim YEAR_DATE As String
    Dim E_DATE As Date
    i = 0
    'On Error GoTo eRRHAND
    
    db.Execute "delete * From TRXMAST WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete * From TRXSUB WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete * From TRXFILE WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    'db2.Execute "delete * From P_Rate WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    
    Set rstMaxRec = New ADODB.Recordset
    rstMaxRec.Open "Select MAX(Val(REC_NO)) From ATRXFILE ", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
        i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
    End If
    rstMaxRec.Close
    Set rstMaxRec = Nothing
    
    E_DATE = Format(TXTINVDATE.Text, "MM/DD/YYYY")
    If Day(E_DATE) <= 12 Then
        DAY_DATE = Format(Month(E_DATE), "00")
        MONTH_DATE = Format(Day(E_DATE), "00")
        YEAR_DATE = Format(Year(E_DATE), "0000")
        E_DATE = DAY_DATE & "/" & MONTH_DATE & "/" & YEAR_DATE
    End If
    E_DATE = Format(E_DATE, "MM/DD/YYYY")
    
    TRXVALUE = 0
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFILE WHERE VCH_DATE = # " & E_DATE & " # ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        TRXVALUE = TRXVALUE + RSTTRXFILE!TRX_TOTAL
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From SALESLEDGER WHERE TRX_TYPE='SI' AND INV_NO = " & Val(txtBillNo.Text) & "", db2, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!INV_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!BAL_AMOUNT = RSTTRXFILE!INV_AMOUNT - RSTTRXFILE!RCPT_AMOUNT
        RSTTRXFILE!ACT_CODE = "131001"
        RSTTRXFILE!ACT_NAME = "ESTIMATE"
        RSTTRXFILE.Update
    Else
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "SI"
        RSTTRXFILE!INV_NO = Val(txtBillNo.Text)
        RSTTRXFILE!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!INV_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!RCPT_AMOUNT = 0
        RSTTRXFILE!BAL_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!ACT_CODE = "131001"
        RSTTRXFILE!ACT_NAME = "ESTIMATE"
        RSTTRXFILE!CHECK_FLAG = "N"
        
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From ATRXFILE  WHERE TRX_TYPE = 'SI' AND VCH_DATE = # " & E_DATE & " # ", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!VCH_AMOUNT = TRXVALUE + Val(LBLTOTAL.Caption)
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
        RSTTRXFILE!VCH_AMOUNT = TRXVALUE + Val(LBLTOTAL.Caption)
        RSTTRXFILE.Update
    Else
        RSTTRXFILE.AddNew
        i = i + 1
        RSTTRXFILE!REC_NO = i
        RSTTRXFILE!TRX_TYPE = "SI"
        RSTTRXFILE!VCH_NO = 0
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = "501001"
        RSTTRXFILE!ACT_NAME = "Sales"
        RSTTRXFILE!VCH_DESC = "Second Sales"
        RSTTRXFILE!VCH_AMOUNT = TRXVALUE + Val(LBLTOTAL.Caption)
        RSTTRXFILE!CD_FLAG = 2
        RSTTRXFILE!POST_FLAG = "Y"
        RSTTRXFILE!CREATE_DATE = Date
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE.Update
        
        RSTTRXFILE.AddNew
        i = i + 1
        RSTTRXFILE!REC_NO = i
        RSTTRXFILE!TRX_TYPE = "SI"
        RSTTRXFILE!VCH_NO = 0
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = "111001"
        RSTTRXFILE!ACT_NAME = "Sales"
        RSTTRXFILE!VCH_DESC = "Cash on Hand"
        RSTTRXFILE!VCH_AMOUNT = TRXVALUE + Val(LBLTOTAL.Caption)
        RSTTRXFILE!CD_FLAG = 1
        RSTTRXFILE!POST_FLAG = "Y"
        RSTTRXFILE!CREATE_DATE = Date
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXMAST WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "SI"
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
    Else
        RSTTRXFILE!VCH_AMOUNT = RSTTRXFILE!VCH_AMOUNT + Val(LBLTOTAL.Caption)
    End If
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = "131001"
    RSTTRXFILE!ACT_NAME = "ESTIMATE"
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = 0
    RSTTRXFILE!REF_NO = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "A"
        RSTTRXFILE!DISCOUNT = Round((Val(TXTTOTALDISC.Text) * 100) / RSTTRXFILE!VCH_AMOUNT, 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "P"
        RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
    End If
    RSTTRXFILE!PAY_AMOUNT = 0
    RSTTRXFILE!CHECK_FLAG = "I"
    RSTTRXFILE!POST_FLAG = "Y"
    RSTTRXFILE!CFORM_NO = Time
    RSTTRXFILE!Remarks = "ESTIMATE"
    RSTTRXFILE!DISC_PERS = 0
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!cr_days = 0
    RSTTRXFILE!BILL_NAME = Trim(TxtBillName.Text)
    RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.Text)
    txtcommi.Tag = ""
    RSTTRXFILE!AGENT_CODE = ""
    RSTTRXFILE!AGENT_NAME = ""
    
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXSUB ", db, adOpenStatic, adLockOptimistic, adCmdText
    
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTTRXTYPE.Text)
    
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "SI"
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(i, 14) = "", 0, grdsales.TextMatrix(i, 14))
        RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(i, 15) = "", 0, grdsales.TextMatrix(i, 15))
        RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(i, 16) = "", "MI", grdsales.TextMatrix(i, 16))
        RSTTRXFILE!QTY = grdsales.TextMatrix(i, 3)
        RSTTRXFILE.Update
    Next i
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "SI"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!CATEGORY = "MEDICINE"
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 13)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 5))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
        RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(i, 24))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "Issued as     " & "ESTIMATE"
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 10)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "131001"
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!Area = ""
        
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
''''    For i = 1 To grdsales.Rows - 1
''''        If Val(grdsales.TextMatrix(i, 6)) <> 0 Then
''''            Set RSTP_RATE = New ADODB.Recordset
''''            RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & grdsales.TextMatrix(i, 13) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
''''            If (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
''''                RSTP_RATE.AddNew
''''            End If
''''            RSTP_RATE!ENTRY_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
''''            RSTP_RATE!ITEM_CODE = grdsales.TextMatrix(i, 13)
''''            RSTP_RATE!ITEM_NAME = grdsales.TextMatrix(i, 2)
''''            RSTP_RATE!PTR = Val(grdsales.TextMatrix(i, 6))
''''            RSTP_RATE!Rate = Val(grdsales.TextMatrix(i, 7))
''''            RSTP_RATE!SALES_TAX = grdsales.TextMatrix(i, 9)
''''            RSTP_RATE!UNIT = grdsales.TextMatrix(i, 4)
''''            RSTP_RATE!CUST_CODE = DataList2.BoundText
''''            RSTP_RATE.Update
''''            RSTP_RATE.Close
''''            Set RSTP_RATE = Nothing
''''        End If
''''    Next i
    
    For i = 1 To grdsales.Rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE!CHECK_FLAG = "Y"
            RSTTRXFILE.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    Next i
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(VCH_NO)) From TRXMAST WHERE TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    LBLPROFIT.Caption = ""
    LBLDATE.Caption = Date
    LBLTOTAL.Caption = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    TXTAMOUNT.Text = ""
    LBLDISCAMT.Caption = ""
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    M_EDIT = False
    cmdRefresh.Enabled = False
    CMDEXIT.Enabled = True
    CMDPRINT.Enabled = False
    CMDEXIT.Enabled = True
    TXTSLNO.Enabled = False
    FRMEHEAD.Enabled = True
    TxtBillName.Enabled = True
    TxtBillName.SetFocus
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Tag = ""
    TxtBillAddress.Text = ""
    TxtBillName.Text = ""
    cr_days = False
    EDIT_INV = False
    Exit Function
eRRHAND:
    MsgBox Err.Description
End Function

Private Sub TxtFree_GotFocus()
    TXTFREE.SelStart = 0
    TXTFREE.SelLength = Len(TXTFREE.Text)
    TXTFREE.Tag = Trim(TXTPRODUCT.Text)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTFREE.Text) = 0 Then GoTo SKIP
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT BAL_QTY  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                i = RSTTRXFILE!BAL_QTY
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            If i > 0 Then
                If Val(TXTFREE.Text) + Val(TXTQTY.Text) > i Then
                    MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                    TXTFREE.SelStart = 0
                    TXTFREE.SelLength = Len(TXTFREE.Text)
                    Exit Sub
                End If
            End If
SKIP:
            TXTFREE.Enabled = False
            TxtMRP.Enabled = True
            TxtMRP.SetFocus
         Case vbKeyEscape
            TXTFREE.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub TxtFree_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtFree_LostFocus()
'    TXTFREE.Text = Format(TXTFREE.Text, ".000")
'    TXTDISC.Tag = 0
'    TXTTAX.Tag = 0
'    If Val(TXTRATE.Text) = 0 Then
'        TXTDISC.Tag = Val(TXTDISC.Text) / 100
'        TXTTAX.Tag = Val(TXTTAX.Text) / 100
'        LBLSUBTOTAL.Caption = Format((Val(TXTFREE.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
'    Else
'        TXTDISC.Tag = Val(TXTFREE.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
'        TXTTAX.Tag = Val(TXTFREE.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
'        LBLSUBTOTAL.Caption = Format((Val(TXTFREE.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
'    End If
End Sub

Private Sub OPTDISCPERCENT_Click()
    TXTTOTALDISC.SetFocus
End Sub

Private Sub OptDiscAmt_Click()
    TXTTOTALDISC.SetFocus
End Sub

Private Sub TXTTOTALDISC_GotFocus()
    TXTTOTALDISC.SelStart = 0
    TXTTOTALDISC.SelLength = Len(TXTTOTALDISC.Text)
End Sub

Private Sub TXTTOTALDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
        End Select
End Sub

Private Sub TXTTOTALDISC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTTOTALDISC_LostFocus()
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        lblnetamount.Caption = Val(lblnetamount.Caption) + Val(grdsales.TextMatrix(i, 12))
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = 0
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
    LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - Val(LBLTOTALCOST.Caption), 2)
    
End Sub

Private Function ReportGeneratION()
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim num As Currency
    Dim SN As Integer
    Dim i As Integer
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    End If
    'On Error GoTo eRRHAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold


    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) & _
              Chr(27) & Chr(72)
        Print #1, Chr(13)
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & AlignLeft(RSTCOMPANY!ADDRESS, 50)
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(7) & "Phone: " & RSTCOMPANY!TEL_NO & ", " & RSTCOMPANY!FAX_NO
        Print #1,
        '''Print #1, Space(7) & "TIN No. " & RSTCOMPANY!KGST
        
        Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & AlignLeft("ESTIMATE", 25)
        'Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & AlignLeft("INVOICE FORM 8H", 16)
    
        'If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, LBLDATE.Caption)
        Print #1, Space(7) & "Bill No. " & Trim(LBLBILLNO.Caption) & Chr(27) & Chr(72) & Space(2) & AlignRight("Date:" & LBLDATE.Caption, 57) '& Space(2) & LBLTIME.Caption
        'LBLDATE.Caption = Date
    
       ' Print #1, Chr(27) & Chr(72) & Space(7) & "Salesman: CS"
    
        Print #1, Space(7) & RepeatString("-", 73)
        Print #1, Space(7) & AlignLeft("SL NO.", 6) & Space(2) & _
                AlignLeft("Description", 25) & Space(1) & _
                AlignLeft("Rate", 7) & Space(3) & _
                AlignLeft("Qty", 4) & Space(4) & _
                AlignLeft("Disc%", 6) & Space(8) & _
                AlignLeft("Amount", 6) & _
                Chr(27) & Chr(72)  '//Bold Ends
    
        Print #1, Space(7) & RepeatString("-", 73)
    
        For i = 1 To grdsales.Rows - 1
            Print #1, Space(7) & AlignLeft(Str(i), 5) & Space(2) & _
                AlignLeft(grdsales.TextMatrix(i, 2), 26) & _
                AlignLeft(grdsales.TextMatrix(i, 7), 2) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 10) & _
                AlignRight(Format(Round(grdsales.TextMatrix(i, 8), 2), "0.00"), 10) & _
                AlignRight(Format(grdsales.TextMatrix(i, 12), "0.00"), 16) & _
                Chr(27) & Chr(72)  '//Bold Ends
        Next i
    
        Print #1, Space(7) & AlignRight("-------------", 73)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("BILL AMOUNT ", 57) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 10)
            Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("DISC AMOUNT ", 57) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 10)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("BILL AMOUNT ", 57) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 10)
        End If
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 57) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Round off ", 57) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 10)
        Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 9)
        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
        num = CCur(Round(LBLTOTAL.Caption, 0))
        Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Rupees " & Words_1_all(num), 40)
        Print #1, Chr(27) & Chr(67) & Chr(0)
        Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** THANK YOU ****", 40)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing

    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)

    Close #1 '//Closing the file
    Exit Function

eRRHAND:
    MsgBox Err.Description
End Function

Private Sub TXTRETAIL_GotFocus()
    txtretail.SelStart = 0
    txtretail.SelLength = Len(txtretail.Text)
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtretail.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = True
            TXTRETAILNOTAX.SetFocus
    End Select
End Sub

Private Sub TXTRETAIL_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRETAILNOTAX_LostFocus()
    txtretail.Text = Format(Val(txtretail.Text), "0.000")
    ''If lblP_Rate.Caption = "0" Then
    If OPTTaxMRP.Value = True Then
        txtretail.Text = Val(TXTRETAILNOTAX.Text) + Val(txtmrpbt.Text) * Val(TXTTAX.Text) / 100
    End If
    If OPTVAT.Value = True Then
        txtretail.Text = Val(TXTRETAILNOTAX.Text) + Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100
    End If
    If optnet.Value = True Then
        txtretail.Text = TXTRETAILNOTAX.Text
    End If
    TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.000")
    If TxtRetailmode.Text = "A" Then
        txtcommi.Text = Format(Round(Val(txtretaildummy.Text), 2), "0.00")
    Else
        txtcommi.Text = Format(Round(Val(TXTRETAILNOTAX.Text) * Val(txtretaildummy.Text) / 100, 2), "0.00")
    End If
    '''End If
End Sub

Private Sub TXTRETAILNOTAX_GotFocus()
    TXTRETAILNOTAX.SelStart = 0
    TXTRETAILNOTAX.SelLength = Len(TXTRETAILNOTAX.Text)
End Sub

Private Sub TXTRETAILNOTAX_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTRETAILNOTAX.Text) = 0 Then Exit Sub
            TXTRETAILNOTAX.Enabled = False
            txtretail.Enabled = True
            txtretail.SetFocus
        Case vbKeyEscape
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
    End Select
End Sub

Private Sub TXTRETAILNOTAX_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub TXTRETAIL_LostFocus()
    TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.000")
    txtretail.Text = Format(Val(txtretail.Text), "0.000")
    'TXTDISC.Tag = 0
    'TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
    'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag), ".000")
End Sub

Private Sub TxtBillName_GotFocus()
    TxtBillName.SelStart = 0
    TxtBillName.SelLength = Len(TxtBillName.Text)
End Sub

Private Sub TxtBillName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            TxtBillAddress.SetFocus
        Case vbKeyEscape
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
    End Select

End Sub

Private Sub TxtBillName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtBillAddress_GotFocus()
    TxtBillAddress.SelStart = 0
    TxtBillAddress.SelLength = Len(TxtBillAddress.Text)
End Sub

Private Sub txtcommi_GotFocus()
    txtcommi.SelStart = 0
    txtcommi.SelLength = Len(txtcommi.Text)
End Sub

Private Sub txtcommi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.Enabled = True
            txtcommi.Enabled = False
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTDISC.Enabled = True
            txtcommi.Enabled = False
            TXTDISC.SetFocus
    End Select
End Sub

Private Sub txtcommi_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcommi_LostFocus()
    txtcommi.Text = Format(txtcommi.Text, ".000")
End Sub

