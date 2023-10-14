VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMDUPLI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DUPLICATE BILL"
   ClientHeight    =   11160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "FRMDUPLI.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   11160
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   3585
      Top             =   10110
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   11190
      TabIndex        =   63
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
      Left            =   1785
      TabIndex        =   56
      Top             =   3255
      Visible         =   0   'False
      Width           =   6030
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   2835
         Left            =   75
         TabIndex        =   57
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
      Left            =   1860
      TabIndex        =   52
      Top             =   3270
      Visible         =   0   'False
      Width           =   5835
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2535
         Left            =   90
         TabIndex        =   55
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
      Begin VB.Label lblhead 
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
         TabIndex        =   54
         Top             =   105
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.Label lblhead 
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
         TabIndex        =   53
         Top             =   105
         Visible         =   0   'False
         Width           =   2610
      End
   End
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   9780
      Left            =   -180
      TabIndex        =   20
      Top             =   -15
      Width           =   11310
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
         TabIndex        =   60
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00FFC0C0&
         Height          =   1335
         Left            =   210
         TabIndex        =   21
         Top             =   -45
         Width           =   11085
         Begin VB.Frame FRMEMODE 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Bill Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   1155
            Left            =   8580
            TabIndex        =   86
            Top             =   150
            Width           =   2475
            Begin VB.OptionButton optothers 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Others"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   75
               TabIndex        =   89
               Top             =   840
               Width           =   1830
            End
            Begin VB.OptionButton optoushadi 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Oushadhi"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   75
               TabIndex        =   88
               Top             =   525
               Width           =   1830
            End
            Begin VB.OptionButton OPTAUTOMATIC 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Automatic"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   75
               TabIndex        =   87
               Top             =   225
               Value           =   -1  'True
               Width           =   1485
            End
         End
         Begin VB.TextBox TXTPATIENT 
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
            Left            =   1230
            MaxLength       =   35
            TabIndex        =   2
            Top             =   645
            Width           =   2850
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
            Height          =   345
            Left            =   1290
            TabIndex        =   0
            Top             =   210
            Width           =   915
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   345
            Left            =   3840
            TabIndex        =   1
            Top             =   225
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor"
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
            Left            =   4215
            TabIndex        =   70
            Top             =   705
            Width           =   765
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "PATIENT"
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
            Left            =   120
            TabIndex        =   69
            Top             =   705
            Width           =   990
         End
         Begin MSForms.ComboBox TXTDOCTOR 
            Height          =   360
            Left            =   4995
            TabIndex        =   3
            Top             =   645
            Width           =   3465
            VariousPropertyBits=   746604571
            ForeColor       =   255
            MaxLength       =   35
            DisplayStyle    =   3
            Size            =   "6112;635"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            DropButtonStyle =   0
            BorderColor     =   255
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE DATE"
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
            Index           =   8
            Left            =   2490
            TabIndex        =   61
            Top             =   255
            Width           =   1335
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
            TabIndex        =   25
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DATE"
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
            Left            =   5505
            TabIndex        =   24
            Top             =   255
            Width           =   645
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
            Left            =   6120
            TabIndex        =   23
            Top             =   225
            Width           =   1215
         End
         Begin VB.Label LBLTIME 
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
            Left            =   7335
            TabIndex        =   22
            Top             =   225
            Width           =   1110
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   6465
         Left            =   210
         TabIndex        =   26
         Top             =   1200
         Width           =   11085
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Height          =   6255
            Left            =   9180
            TabIndex        =   71
            Top             =   165
            Width           =   1815
            Begin VB.Label lblhead 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "DUMMY BILL"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   1110
               Index           =   1
               Left            =   45
               TabIndex        =   98
               Top             =   2670
               Width           =   1650
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
               Index           =   6
               Left            =   135
               TabIndex        =   85
               Top             =   870
               Width           =   1515
            End
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
               Left            =   165
               TabIndex        =   84
               Top             =   1125
               Width           =   1440
            End
            Begin VB.Label lbltotal 
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
               ForeColor       =   &H00000000&
               Height          =   435
               Left            =   165
               TabIndex        =   83
               Top             =   390
               Width           =   1440
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
               ForeColor       =   &H00004000&
               Height          =   375
               Index           =   22
               Left            =   135
               TabIndex        =   82
               Top             =   1575
               Width           =   1515
            End
            Begin VB.Label lblnetamount 
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
               ForeColor       =   &H00000080&
               Height          =   435
               Left            =   165
               TabIndex        =   81
               Top             =   1815
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "BILL AMOUNT"
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
               Height          =   375
               Index           =   21
               Left            =   135
               TabIndex        =   80
               Top             =   150
               Width           =   1515
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
               ForeColor       =   &H000000FF&
               Height          =   375
               Index           =   25
               Left            =   135
               TabIndex        =   79
               Top             =   2490
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label LBLTOTALCOST 
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
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   165
               TabIndex        =   78
               Top             =   2715
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label LBLPROFIT 
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
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   165
               TabIndex        =   77
               Top             =   3405
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   26
               Left            =   135
               TabIndex        =   76
               Top             =   3180
               Visible         =   0   'False
               Width           =   1515
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
               TabIndex        =   75
               Top             =   4980
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
               TabIndex        =   74
               Top             =   5775
               Visible         =   0   'False
               Width           =   1425
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
               TabIndex        =   73
               Top             =   4740
               Visible         =   0   'False
               Width           =   1425
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
               TabIndex        =   72
               Top             =   5520
               Visible         =   0   'False
               Width           =   1395
            End
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
            Left            =   6240
            TabIndex        =   28
            Top             =   6075
            Width           =   600
         End
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
            Left            =   7755
            TabIndex        =   27
            Top             =   6060
            Width           =   1080
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   5730
            Left            =   90
            TabIndex        =   19
            Top             =   270
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   10107
            _Version        =   393216
            Rows            =   1
            Cols            =   20
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
         Begin VB.Label Label1 
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
            Height          =   300
            Index           =   4
            Left            =   5325
            TabIndex        =   30
            Top             =   6105
            Width           =   870
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Index           =   5
            Left            =   6900
            TabIndex        =   29
            Top             =   6105
            Width           =   780
         End
      End
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   465
         Left            =   11100
         TabIndex        =   51
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
         BackColor       =   &H00FFC0C0&
         Height          =   2130
         Left            =   195
         TabIndex        =   31
         Top             =   7605
         Width           =   11100
         Begin VB.CommandButton cmditemcreate 
            Caption         =   "&Create Item"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   8895
            TabIndex        =   91
            Top             =   1350
            Width           =   1125
         End
         Begin VB.TextBox TXTCATEGORY 
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
            Left            =   8910
            TabIndex        =   90
            Top             =   1170
            Visible         =   0   'False
            Width           =   2010
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
            Height          =   465
            Left            =   90
            TabIndex        =   66
            Top             =   825
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
            Height          =   465
            Left            =   10335
            TabIndex        =   62
            Top             =   810
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
            Left            =   6090
            MaxLength       =   6
            TabIndex        =   58
            Top             =   450
            Width           =   630
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
            Height          =   480
            Left            =   2895
            TabIndex        =   13
            Top             =   780
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
            Left            =   30
            TabIndex        =   4
            Top             =   450
            Width           =   570
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
            Left            =   615
            TabIndex        =   5
            Top             =   450
            Width           =   3885
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
            Left            =   5310
            MaxLength       =   7
            TabIndex        =   6
            Top             =   450
            Width           =   765
         End
         Begin VB.TextBox TXTRATE 
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
            Left            =   8325
            MaxLength       =   6
            TabIndex        =   7
            Top             =   1860
            Visible         =   0   'False
            Width           =   630
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
            Left            =   6735
            MaxLength       =   4
            TabIndex        =   8
            Top             =   450
            Width           =   600
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
            Left            =   9450
            MaxLength       =   4
            TabIndex        =   11
            Top             =   450
            Width           =   660
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
            Height          =   465
            Left            =   6495
            TabIndex        =   16
            Top             =   810
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
            Height          =   465
            Left            =   8895
            TabIndex        =   18
            Top             =   825
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
            Height          =   465
            Left            =   5295
            TabIndex        =   15
            Top             =   795
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
            Height          =   465
            Left            =   4095
            TabIndex        =   14
            Top             =   795
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
            TabIndex        =   36
            Top             =   1950
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
            Left            =   8490
            MaxLength       =   15
            TabIndex        =   10
            Top             =   450
            Width           =   930
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
            TabIndex        =   35
            Top             =   1965
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
            TabIndex        =   34
            Top             =   1965
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
            TabIndex        =   33
            Top             =   1995
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
            Left            =   4530
            TabIndex        =   32
            Top             =   450
            Width           =   765
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
            Height          =   465
            Left            =   7695
            TabIndex        =   17
            Top             =   810
            Width           =   1125
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   300
            Left            =   7365
            TabIndex        =   9
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Others Bill #"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   300
            Index           =   30
            Left            =   4425
            TabIndex        =   97
            Top             =   1395
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Oushadhi Bill #"
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
            Index           =   29
            Left            =   1935
            TabIndex        =   96
            Top             =   1395
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "SLIP #"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   300
            Index           =   23
            Left            =   165
            TabIndex        =   95
            Top             =   1395
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label LBLOTHERS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00008080&
            Height          =   345
            Left            =   5640
            TabIndex        =   94
            Top             =   1335
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label LBLOUSHADI 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H000000C0&
            Height          =   345
            Left            =   3405
            TabIndex        =   93
            Top             =   1335
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label LBLSLIP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00008000&
            Height          =   345
            Left            =   825
            TabIndex        =   92
            Top             =   1335
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label LBLDNORCN 
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
            Left            =   705
            TabIndex        =   67
            Top             =   750
            Visible         =   0   'False
            Width           =   510
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
            Left            =   6090
            TabIndex        =   59
            Top             =   225
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "SL No"
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
            Left            =   30
            TabIndex        =   50
            Top             =   225
            Width           =   570
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
            Left            =   615
            TabIndex        =   49
            Top             =   225
            Width           =   3885
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
            Left            =   5310
            TabIndex        =   48
            Top             =   225
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Rate"
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
            Left            =   8325
            TabIndex        =   47
            Top             =   1920
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Tax %"
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
            Left            =   6735
            TabIndex        =   46
            Top             =   225
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Disc %"
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
            Left            =   9450
            TabIndex        =   45
            Top             =   240
            Width           =   660
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
            Left            =   10125
            TabIndex        =   44
            Top             =   240
            Width           =   930
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
            TabIndex        =   43
            Top             =   1965
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Exp Date"
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
            Index           =   16
            Left            =   7365
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Batch"
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
            Left            =   8490
            TabIndex        =   41
            Top             =   240
            Width           =   930
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
            Left            =   10125
            TabIndex        =   12
            Top             =   450
            Width           =   930
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
            TabIndex        =   40
            Top             =   1980
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
            TabIndex        =   39
            Top             =   2010
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Pack"
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
            Index           =   20
            Left            =   4530
            TabIndex        =   38
            Top             =   225
            Width           =   765
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
            TabIndex        =   37
            Top             =   1980
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   15
      TabIndex        =   68
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11385
      TabIndex        =   65
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11595
      TabIndex        =   64
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMDUPLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim TMPREC As New ADODB.Recordset
Dim TMPFLAG As Boolean
Dim ACT_REC As New ADODB.Recordset

Dim ACT_FLAG As Boolean
Dim PHY_BATCH As New ADODB.Recordset
Dim BATCH_FLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean

Dim CLOSEALL As Integer
Dim M_STOCK As Integer
Dim M_EDIT As Boolean

Private Sub CMDSALERETURN_Click()

End Sub

Private Sub cmditemcreate_Click()
    MDIMAIN.Enabled = False
    frmitemmaster.Show
End Sub
Private Sub cmditemcreate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub cmdstockadjst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
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
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TxtMRP.Text), ".000") 'Format(Val(TXTRATE.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Val(TXTDISC.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(TXTTAX.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = IIf(TXTEXPIRY.Text = "  /  ", "", Trim(TXTEXPIRY.Text))
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTTRXTYPE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = ""
    grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = IIf(Trim(TXTCATEGORY.Text) = "", "Others", Trim(TXTCATEGORY.Text))
  
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = IIf(IsNull(RSTTRXFILE!MANUFACTURER), "", Trim(RSTTRXFILE!MANUFACTURER))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Select Case LBLDNORCN.Caption
        Case "DN"
            grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = "DN"
        Case "CN"
            grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = "CN"
        Case Else
            grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = "B"
    End Select

    
    lbltotal.Caption = ""
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 18)
            Case "CN"
                lbltotal.Caption = Round(Val(lbltotal.Caption) - Val(grdsales.TextMatrix(i, 11)), 2)
            Case Else
                lbltotal.Caption = Round(Val(lbltotal.Caption) + Val(grdsales.TextMatrix(i, 11)), 2)
        End Select
    Next i
    lbltotal.Tag = Val(lbltotal.Caption)
    TXTAMOUNT.Text = Round((Val(lbltotal.Caption) * Val(TXTTOTALDISC.Text) / 100), 2)
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(lbltotal.Caption) - Val(TXTAMOUNT.Text), 2)
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTCATEGORY.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Text = ""
    TxtMRP.Text = ""
    TXTRATE.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtBatch.Text = ""
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    cmddelete.Enabled = False
    cmdexit.Enabled = False
    TXTSLNO.Enabled = True
    M_EDIT = False
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
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            If Val(!ISSUE_VAL) > 0 Then !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            If Val(!CLOSE_VAL) > 0 Then !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    For i = Val(TXTSLNO.Text) - 1 To grdsales.Rows - 2
        grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = i
        grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(i + 1, 6)
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
    Next i
    grdsales.Rows = grdsales.Rows - 1
    lbltotal.Caption = ""
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 18)
            Case "CN"
                lbltotal.Caption = Val(lbltotal.Caption) - Val(grdsales.TextMatrix(i, 11))
            Case Else
                lbltotal.Caption = Val(lbltotal.Caption) + Val(grdsales.TextMatrix(i, 11))
        End Select
    Next i
    lbltotal.Tag = Val(lbltotal.Caption)
    TXTAMOUNT.Text = Round((Val(lbltotal.Caption) * Val(TXTTOTALDISC.Text) / 100), 2)
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(lbltotal.Caption) - Val(TXTAMOUNT.Text), 2)
    
    Call COSTCALCULATION
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTCATEGORY.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    TxtMRP.Text = ""
    TXTTAX.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTDISC.Text = ""
    txtBatch.Text = ""
    LBLSUBTOTAL.Caption = ""
    LBLDNORCN.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    cmddelete.Enabled = False
    CMDMODIFY.Enabled = False
    cmdexit.Enabled = False
    M_EDIT = False
    If grdsales.Rows = 1 Then
        cmdexit.Enabled = True
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
            TXTRATE.Text = ""
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTCATEGORY.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTUNIT.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            cmddelete.Enabled = False
    End Select
End Sub

Private Sub cmdexit_Click()
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
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            If Val(!ISSUE_VAL) > 0 Then !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            If Val(!CLOSE_VAL) > 0 Then !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    CMDMODIFY.Enabled = False
    cmddelete.Enabled = False
    cmdexit.Enabled = False
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
            TXTRATE.Text = ""
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTCATEGORY.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTUNIT.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            cmddelete.Enabled = False
    End Select
End Sub

Private Sub cmdprint_Click()
    If Val(txtBillNo.Text) = 0 Then
        MsgBox "Please enter a valid Bill number", vbOKOnly, "Duplicate Sale"
        txtBillNo.Enabled = True
        txtBillNo.SetFocus
        Exit Sub
    End If
    Me.Enabled = False
    frmDUPbilltype.Show
End Sub

Private Sub CMDPRINT_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTCATEGORY.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTUNIT.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            cmddelete.Enabled = False
    End Select
End Sub

Private Sub cmdRefresh_Click()
    
   ' If grdsales.Rows = 1 Then GoTo SKIP
    
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Dummy Bill..."
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Dummy Bill..."
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    End If
    Call AppendSale
    
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTCATEGORY.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            cmddelete.Enabled = False
    End Select
End Sub

Private Sub Form_Activate()
     Dim rstDOCTORS As ADODB.Recordset
'    If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
'    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
'    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
'    If TxtMRP.Enabled = True Then TxtMRP.SetFocus
'    If TXTTAX.Enabled = True Then TXTTAX.SetFocus
'    If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
'    'If TXTEXPEnabled = True Then TXTEXPDATE.SetFocus
'    If txtBatch.Enabled = True Then txtBatch.SetFocus
'    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
'    If cmdadd.Enabled = True Then cmdadd.SetFocus
    
    On Error GoTo eRRHAND
    
    Set rstDOCTORS = New ADODB.Recordset
    rstDOCTORS.Open "Select DISTINCT [ACT_NAME] From DOCTORLIST ORDER BY [ACT_NAME]", db2, adOpenForwardOnly
    Do Until rstDOCTORS.EOF
        If Not IsNull(rstDOCTORS!ACT_NAME) Then TXTDOCTOR.AddItem (rstDOCTORS!ACT_NAME)
        rstDOCTORS.MoveNext
    Loop
    rstDOCTORS.Close
    Set rstDOCTORS = Nothing
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    ACT_FLAG = True
    lblcredit.Caption = "1"
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2000
    grdsales.ColWidth(3) = 500
    grdsales.ColWidth(4) = 500
    grdsales.ColWidth(5) = 1000
    grdsales.ColWidth(6) = 1000
    grdsales.ColWidth(7) = 800
    grdsales.ColWidth(8) = 600
    grdsales.ColWidth(9) = 800
    grdsales.ColWidth(10) = 1100
    grdsales.ColWidth(11) = 1000
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "UNIT"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "RATE"
    grdsales.TextArray(7) = "DISC %"
    grdsales.TextArray(8) = "TAX %"
    grdsales.TextArray(9) = "BATCH"
    grdsales.TextArray(10) = "EXPIRY"
    grdsales.TextArray(11) = "SUB TOTAL"
    grdsales.TextArray(12) = "ITEM CODE"
    grdsales.TextArray(13) = "Vch No"
    grdsales.TextArray(14) = "Line No"
    grdsales.TextArray(15) = "Trx Type"
    grdsales.TextArray(16) = "Flag"
    grdsales.TextArray(17) = "MFGR"
    grdsales.TextArray(18) = "CN/DN"
    'grdsales.ColWidth(12) = 0
    'grdsales.ColWidth(13) = 0
    'grdsales.ColWidth(14) = 0
   'grdsales.ColWidth(15) = 0
    'grdsales.ColWidth(16) = 0
    
    lbltotal.Caption = 0
    
    PHYFLAG = True
    TMPFLAG = True
    BATCH_FLAG = True
    ITEM_FLAG = True
    
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTUNIT.Enabled = False
    TxtMRP.Enabled = False
    
    TXTTAX.Enabled = False
    TXTEXPIRY.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    cmddelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    TXTSLNO.Text = 1
    
    CLOSEALL = 1
    M_EDIT = False
    Me.Width = 11100
    Me.Height = 10000
    Me.Left = 0
    Me.Top = 0
    'If TXTPATIENT.Enabled = True Then TXTPATIENT.SetFocus
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
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDPOPUP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTP_RATE As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            'TXTQTY.Text = GRDPOPUP.Columns(1)
            TxtMRP.Text = GRDPOPUP.Columns(3)
            TXTRATE.Text = GRDPOPUP.Columns(3)
            ''''TXTRATE.Text = GRDPOPUP.Columns(4)
            If IsNull(GRDPOPUP.Columns(12)) Or GRDPOPUP.Columns(12) <> "V" Then
                TXTTAX.Text = "0"
            ElseIf GRDPOPUP.Columns(12) = "V" Then
                TXTTAX.Text = GRDPOPUP.Columns(5)
            End If
            'TXTTAX.Text = 0  'GRDPOPUP.Columns(5)
            TXTEXPIRY.Text = IIf(GRDPOPUP.Columns(2) = "", "  /  ", Format(GRDPOPUP.Columns(2), "mm/yy"))
            txtBatch.Text = GRDPOPUP.Columns(0)
            
            TXTVCHNO.Text = GRDPOPUP.Columns(8)
            TXTLINENO.Text = GRDPOPUP.Columns(9)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(10)
            TXTUNIT.Text = GRDPOPUP.Columns(11)
            
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT CATEGORY FROM ITEMMAST WHERE ITEM_CODE = '" & GRDPOPUP.Columns(6) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                TXTCATEGORY.Text = IIf(IsNull(RSTITEMMAST!CATEGORY), "OTHERS", RSTITEMMAST!CATEGORY)
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
''''            Set RSTP_RATE = New ADODB.Recordset
''''            RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & GRDPOPUP.Columns(6) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
''''            If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
''''                TXTRATE.Text = RSTP_RATE!SALES_PRICE
''''            End If
''''            RSTP_RATE.Close
''''            Set RSTP_RATE = Nothing
            
            Set GRDPOPUP.DataSource = Nothing
            
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTVCHNO.Text = ""
            TXTCATEGORY.Text = ""
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
    Dim RSTP_RATE As ADODB.Recordset
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
                Set RSTNONSTOCK = New ADODB.Recordset
                RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                i = RSTNONSTOCK.RecordCount
                RSTNONSTOCK.Close
                Set RSTNONSTOCK = Nothing
''''                If i = 0 Then
''''                    If (MsgBox("NO STOCK AVAILABLE..Do you want to add to Stockless", vbYesNo, "SALES") = vbYes) Then
''''                        Set RSTNONSTOCK = New ADODB.Recordset
''''                        RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
''''                        If (RSTNONSTOCK.EOF And RSTNONSTOCK.BOF) Then
''''                            RSTNONSTOCK.AddNew
''''                            RSTNONSTOCK!ITEM_NAME = TXTPRODUCT.Text
''''                            RSTNONSTOCK!ITEM_CODE = TXTITEMCODE.Text
''''                            RSTNONSTOCK!Date = Date & " " & Time
''''                            RSTNONSTOCK.Update
''''                        End If
''''                        RSTNONSTOCK.Close
''''                        Set RSTNONSTOCK = Nothing
''''                    End If
''''                    Exit Sub
''''                End If
                
                MINUSFLAG = True
                NONSTOCKFLAG = True
            End If
            For i = 1 To grdsales.Rows - 1
                If Trim(grdsales.TextMatrix(i, 12)) = Trim(TXTITEMCODE.Text) Then
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
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                Else
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                End If
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                If NONSTOCKFLAG = True Then
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                Else
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
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
                ''''''''TXTRATE.Text = GRDPOPUPITEM.Columns(3)
                TXTRATE.Text = GRDPOPUPITEM.Columns(11)
                TxtMRP.Text = GRDPOPUPITEM.Columns(11)
                If IsNull(PHY_ITEM!CHECK_FLAG) Or PHY_ITEM!CHECK_FLAG <> "V" Then
                    TXTTAX.Text = "0"
                ElseIf PHY_ITEM!CHECK_FLAG = "V" Then
                    TXTTAX.Text = GRDPOPUPITEM.Columns(4)
                End If
                'TXTTAX.Text = 0 'GRDPOPUPITEM.Columns(4)
                TXTEXPIRY.Text = IIf(GRDPOPUPITEM.Columns(7) = "", "  /  ", Format(GRDPOPUPITEM.Columns(7), "MM/YY"))
                txtBatch.Text = GRDPOPUPITEM.Columns(6)
                
                TXTVCHNO.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(8), "")
                TXTLINENO.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(9), "")
                TXTTRXTYPE.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(10), "")
                TXTUNIT.Text = GRDPOPUPITEM.Columns(5)
                
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT CATEGORY FROM ITEMMAST WHERE ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                    TXTCATEGORY.Text = IIf(IsNull(RSTITEMMAST!CATEGORY), "OTHERS", RSTITEMMAST!CATEGORY)
                End If
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                
'                Set RSTP_RATE = New ADODB.Recordset
'                RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & GRDPOPUPITEM.Columns(0) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
'                If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
'                    TXTRATE.Text = RSTP_RATE!SALES_PRICE
'                End If
'                RSTP_RATE.Close
'                Set RSTP_RATE = Nothing
            
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
            TXTCATEGORY.Text = ""
            TXTCATEGORY.Text = ""
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

Private Sub Label2_Click()

End Sub

Private Sub OPTAUTOMATIC_Click()
    If TXTPATIENT.Enabled = True Then TXTPATIENT.SetFocus
    If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If TxtMRP.Enabled = True Then TxtMRP.SetFocus
    If TXTTAX.Enabled = True Then TXTTAX.SetFocus
    If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
    If TXTEXPEnabled = True Then TXTEXPDATE.SetFocus
    If txtBatch.Enabled = True Then txtBatch.SetFocus
    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
End Sub

Private Sub OPTAUTOMATIC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TXTPATIENT.Enabled = True Then TXTPATIENT.SetFocus
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
            If TXTEXPEnabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
        End Select
End Sub

Private Sub optothers_Click()
    If TXTPATIENT.Enabled = True Then TXTPATIENT.SetFocus
    If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If TxtMRP.Enabled = True Then TxtMRP.SetFocus
    If TXTTAX.Enabled = True Then TXTTAX.SetFocus
    If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
    If TXTEXPEnabled = True Then TXTEXPDATE.SetFocus
    If txtBatch.Enabled = True Then txtBatch.SetFocus
    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
End Sub

Private Sub optothers_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TXTPATIENT.Enabled = True Then TXTPATIENT.SetFocus
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
            If TXTEXPEnabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
        End Select
End Sub

Private Sub optoushadi_Click()
    If TXTPATIENT.Enabled = True Then TXTPATIENT.SetFocus
    If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If TxtMRP.Enabled = True Then TxtMRP.SetFocus
    If TXTTAX.Enabled = True Then TXTTAX.SetFocus
    If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
    If TXTEXPEnabled = True Then TXTEXPDATE.SetFocus
    If txtBatch.Enabled = True Then txtBatch.SetFocus
    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
End Sub

Private Sub optoushadi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TXTPATIENT.Enabled = True Then TXTPATIENT.SetFocus
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
            If TXTEXPEnabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
        End Select
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = True
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTEXPIRY.SetFocus
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

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtBillNo.Text) = 0 Then Exit Sub
            TXTINVDATE.Enabled = True
            txtBillNo.Enabled = False
            TXTINVDATE.SetFocus
    End Select
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

Private Sub TXTDISC_GotFocus()
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.Text)
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            cmdadd.Enabled = True
            TXTDISC.Enabled = False
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
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
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TxtMRP.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TxtMRP.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TxtMRP.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
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
                DataList2.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                MsgBox "Enter Proper Invoice Date", vbOKOnly, "Dummy Bill..."
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
                MsgBox "Enter Proper Invoice Date", vbOKOnly, "Dummy Bill..."
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                TXTPATIENT.SetFocus
            End If
        Case vbKeyEscape
            txtBillNo.Enabled = True
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
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TxtMRP.Enabled = False
            TXTTAX.Enabled = True
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTTAX.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TxtMRP.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTQTY.SetFocus
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
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TxtMRP.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TxtMRP.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TxtMRP.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
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
    Dim RSTP_RATE As ADODB.Recordset
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
            cmddelete.Enabled = False
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
            LBLSUBTOTAL.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
           
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME],[CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME],[CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 12)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                Next i
                Set grdtmp.DataSource = Nothing
                If TMPFLAG = True Then
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                Else
                    TMPREC.Close
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                End If
                
                Set grdtmp.DataSource = TMPREC
                If TMPREC.RecordCount = 1 Then
                    'TXTQTY.Text = grdtmp.Columns(2)
                    TxtMRP.Text = grdtmp.Columns(3)
                    '''TXTRATE.Text = grdtmp.Columns(4)
                    TXTRATE.Text = grdtmp.Columns(3)
                    If IsNull(TMPREC!CHECK_FLAG) Or TMPREC!CHECK_FLAG <> "V" Then
                        TXTTAX.Text = "0"
                    ElseIf TMPREC!CHECK_FLAG = "V" Then
                        TXTTAX.Text = grdtmp.Columns(5)
                    End If
                    'IIf (IsNull(TMPREC!CHECK_FLAG) Or TMPREC!CHECK_FLAG <> "V"), TXTTAX.Text = "", TXTTAX.Text = grdtmp.Columns(5)
                    TXTEXPIRY.Text = IIf(grdtmp.Columns(8) = "", "  /  ", Format(grdtmp.Columns(8), "MM/YY"))
                    txtBatch.Text = grdtmp.Columns(7)
                    
                    TXTVCHNO.Text = grdtmp.Columns(9)
                    TXTLINENO.Text = grdtmp.Columns(10)
                    TXTTRXTYPE.Text = grdtmp.Columns(11)
                    TXTUNIT.Text = grdtmp.Columns(6)
                    
                    Set RSTITEMMAST = New ADODB.Recordset
                    RSTITEMMAST.Open "SELECT CATEGORY FROM ITEMMAST WHERE ITEM_CODE = '" & grdtmp.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                        TXTCATEGORY.Text = IIf(IsNull(RSTITEMMAST!CATEGORY), "OTHERS", RSTITEMMAST!CATEGORY)
                    End If
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
                    
'                    Set RSTP_RATE = New ADODB.Recordset
'                    RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & grdtmp.Columns(0) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
'                    If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
'                        TXTRATE.Text = RSTP_RATE!SALES_PRICE
'                    End If
'                    RSTP_RATE.Close
'                    Set RSTP_RATE = Nothing
                    
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
                    Set RSTNONSTOCK = New ADODB.Recordset
                    RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                    i = RSTNONSTOCK.RecordCount
                    RSTNONSTOCK.Close
                    Set RSTNONSTOCK = New ADODB.Recordset
'''                    If i = 0 Then
'''                        If (MsgBox("NO STOCK AVAILABLE..Do you want to add to Stockless", vbYesNo, "SALES") = vbYes) Then
'''                            Set RSTNONSTOCK = New ADODB.Recordset
'''                            RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
'''                            If (RSTNONSTOCK.EOF And RSTNONSTOCK.BOF) Then
'''                                RSTNONSTOCK.AddNew
'''                                RSTNONSTOCK!ITEM_NAME = TXTPRODUCT.Text
'''                                RSTNONSTOCK!ITEM_CODE = TXTITEMCODE.Text
'''                                RSTNONSTOCK!Date = Date & " " & Time
'''                                RSTNONSTOCK.Update
'''                            End If
'''                            RSTNONSTOCK.Close
'''                            Set RSTNONSTOCK = Nothing
'''                        End If
'''                    End If
                
                    Set RSTZEROSTOCK = New ADODB.Recordset
                    RSTZEROSTOCK.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY [VCH_NO]", db, adOpenStatic, adLockReadOnly
                    If Not (RSTZEROSTOCK.EOF And RSTZEROSTOCK.BOF) Then
                        TxtMRP.Text = RSTZEROSTOCK!MRP
                        If IsNull(RSTZEROSTOCK!CHECK_FLAG) Or RSTZEROSTOCK!CHECK_FLAG <> "V" Then
                            TXTTAX.Text = "0"
                        ElseIf RSTZEROSTOCK!CHECK_FLAG = "V" Then
                            TXTTAX.Text = RSTZEROSTOCK!SALES_TAX
                        End If
                        TXTEXPIRY.Text = IIf(IsNull(RSTZEROSTOCK!EXP_DATE), "  /  ", Format(RSTZEROSTOCK!EXP_DATE, "MM/YY"))
                        txtBatch.Text = IIf(IsNull(RSTZEROSTOCK!REF_NO), "", RSTZEROSTOCK!REF_NO)
                        
                        TXTVCHNO.Text = ""
                        TXTCATEGORY.Text = ""
                        TXTLINENO.Text = ""
                        TXTTRXTYPE.Text = ""
                        TXTUNIT.Text = RSTZEROSTOCK!UNIT
                        
'                            Set RSTP_RATE = New ADODB.Recordset
'                            RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & RSTZEROSTOCK!ITEM_CODE & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
'                            If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
'                                TXTRATE.Text = RSTP_RATE!SALES_PRICE
'                            End If
'                            RSTP_RATE.Close
'                            Set RSTP_RATE = Nothing
                    End If
                    RSTZEROSTOCK.Close
                    Set RSTZEROSTOCK = Nothing
                    
                    GoTo JUMPNONSTOCK
                    
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
                TXTEXPIRY.Enabled = False
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
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            cmddelete.Enabled = False
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTSLNO.SetFocus
            cmddelete.Enabled = False
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
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
            
            'TXTEXPIRY.Text = Date
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT EXP_DATE  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!EXP_DATE)) Then
                    TXTEXPIRY.Text = "  /  "
                    txtexpirydate.Text = ""
                Else
                    TXTEXPIRY.Text = Format(RSTTRXFILE!EXP_DATE, "MM/YY")
                    txtexpirydate.Text = RSTTRXFILE!EXP_DATE
                End If
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            If Val(TXTQTY.Text) = 0 Then Exit Sub
''''            If i > 0 Then
''''                If Val(TXTQTY.Text) > i Then
''''                    MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
''''                    TXTQTY.SelStart = 0
''''                    TXTQTY.SelLength = Len(TXTQTY.Text)
''''                    Exit Sub
''''                End If
''''            End If
            If TXTEXPIRY.Text = "  /  " Then GoTo SKIP
            If txtexpirydate.Text = "" Then GoTo SKIP
            
            If DateDiff("d", Date, txtexpirydate.Text) < 0 Then
                MsgBox "Item Expired....", vbOKOnly, "BILL.."
                TXTQTY.SelStart = 0
                TXTQTY.SelLength = Len(TXTQTY.Text)
                Exit Sub
            End If
            
            If DateDiff("d", Date, txtexpirydate.Text) < 90 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, txtexpirydate.Text)) & "Days", vbOKOnly, "BILL.."
                TXTQTY.SelStart = 0
                TXTQTY.SelLength = Len(TXTQTY.Text)
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TxtMRP.Enabled = True
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TxtMRP.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTQTY.Enabled = False
            TXTUNIT.Enabled = True
            TXTUNIT.SetFocus
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
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TxtMRP.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TxtMRP.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TxtMRP.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
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
                TXTRATE.Text = ""
                TxtMRP.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTCATEGORY.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTUNIT.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTEXPIRY.Text = "  /  "
                txtBatch.Text = ""
                TXTSLNO.Text = grdsales.Rows
                cmddelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.Text) >= grdsales.Rows Then
                TXTSLNO.Text = grdsales.Rows
                cmddelete.Enabled = False
                CMDMODIFY.Enabled = False
            End If
            If Val(TXTSLNO.Text) < grdsales.Rows Then
                TXTSLNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 0)
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 3)
                TxtMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                ''''TXTRATE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                TXTRATE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 11), ".000")
                
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 12)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
                TXTCATEGORY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 19)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 15)
                TXTUNIT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                LBLSUBTOTAL.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 11)
                If grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = "" Then
                    TXTEXPIRY.Text = "  /  "
                Else
                    TXTEXPIRY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                End If
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                
                TXTTAX.Enabled = False
                TXTEXPIRY.Enabled = False
                txtBatch.Enabled = False
                TXTDISC.Enabled = False
                TxtMRP.Enabled = False
                Select Case grdsales.TextMatrix(Val(TXTSLNO.Text), 18)
                    Case "CN", "DN"
                        cmddelete.Enabled = True
                        cmddelete.SetFocus
                        
                    Case Else
                        CMDMODIFY.Enabled = True
                        CMDMODIFY.SetFocus
                        cmddelete.Enabled = True
                End Select
                
                LBLDNORCN.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 18)
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If cmddelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTCATEGORY.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTUNIT.Text = ""
                TXTQTY.Text = ""
                TXTRATE.Text = ""
                TxtMRP.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTEXPIRY.Text = "  /  "
                txtBatch.Text = ""
                cmdadd.Enabled = False
                cmddelete.Enabled = False
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
                TXTDOCTOR.Enabled = True
                TXTDOCTOR.SetFocus
            End If
            LBLDNORCN.Caption = ""
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
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = True
            TXTEXPIRY.SetFocus
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TxtMRP.Enabled = True
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
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TxtMRP.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TxtMRP.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TxtMRP.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
End Sub

Private Sub TXTTOTALDISC_GotFocus()
    TXTTOTALDISC.SelStart = 0
    TXTTOTALDISC.SelLength = Len(TXTTOTALDISC.Text)
End Sub

Private Sub TXTTOTALDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            If TXTPATIENT.Enabled = True Then TXTPATIENT.SetFocus
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
            If TXTEXPEnabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
        
            If Len(Trim(TXTEXPIRY.Text)) = 1 Then GoTo SKIP
            If Len(Trim(TXTEXPIRY.Text)) < 5 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
SKIP:
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
             If Len(Trim(TXTEXPIRY.Text)) = 1 Then GoTo Nextstep
            If Len(Trim(TXTEXPIRY.Text)) < 5 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
Nextstep:
            TXTEXPIRY.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    'TXTEXPIRY.Text = Format(TXTEXPIRY.Text, "MM/YY")
    'TXTEXPIRY.Visible = False
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
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "BATCH NO."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "EXP DATE"
    GRDPOPUP.Columns(3).Caption = "MRP"
    GRDPOPUP.Columns(4).Caption = "RATE"
    GRDPOPUP.Columns(5).Caption = "TAX"
    GRDPOPUP.Columns(6).Caption = "VCH No"
    GRDPOPUP.Columns(7).Caption = "Line No"
    GRDPOPUP.Columns(8).Caption = "Trx Type"
    
    GRDPOPUP.Columns(0).Width = 1400
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 1400
    GRDPOPUP.Columns(3).Width = 1000
    GRDPOPUP.Columns(4).Width = 1000
    GRDPOPUP.Columns(5).Width = 900
    
    GRDPOPUP.Columns(6).Visible = False
    GRDPOPUP.Columns(7).Visible = False
    GRDPOPUP.Columns(8).Visible = False
    
    GRDPOPUP.SetFocus
    lblhead(0).Caption = GRDPOPUP.Columns(6).Text
    lblhead(9).Visible = True
    lblhead(0).Visible = True
End Function

Function FILL_ITEMGRID()
    FRMEMAIN.Enabled = False
    FRMEITEM.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEGRDTMP.Visible = False
    
    
    If ITEM_FLAG = True Then
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
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
            !OPEN_QTY = M_STOCK
            !OPEN_VAL = 0
            !RCPT_QTY = 0
            !RCPT_VAL = 0
            !ISSUE_QTY = 0
            !ISSUE_VAL = 0
            !CLOSE_QTY = M_STOCK
            !CLOSE_VAL = 0
            !DAM_QTY = 0
            !DAM_VAL = 0
            RSTITEMMAST.Update
        End If
    End With
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Exit Function
    
eRRHAND:
    MsgBox Err.Description
End Function

Private Sub TXTTOTALDISC_LostFocus()
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 18)
            Case "CN"
                lblnetamount.Caption = Val(lblnetamount.Caption) - Val(grdsales.TextMatrix(i, 11))
            Case Else
                lblnetamount.Caption = Val(lblnetamount.Caption) + Val(grdsales.TextMatrix(i, 11))
        End Select
    Next i
    lbltotal.Tag = Val(lbltotal.Caption)
    TXTAMOUNT.Text = Round((Val(lbltotal.Caption) * Val(TXTTOTALDISC.Text) / 100), 2)
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(lbltotal.Caption) - Val(TXTAMOUNT.Text), 2)
    LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - Val(LBLTOTALCOST.Caption), 2)
    
End Sub

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
        RSTCOST.Open "SELECT [ITEM_COST] FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(N, 15)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(N, 13)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(N, 14)) & "", db, adOpenStatic, adLockReadOnly, adCmdText
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

Public Function AppendSale()
    
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    TXTPATIENT.Text = ""
    TXTDOCTOR.Text = ""
    LBLDNORCN.Caption = ""
    lblnetamount.Caption = ""
    LBLPROFIT.Caption = ""
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    lbltotal.Caption = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    TXTAMOUNT.Text = ""
    LBLDISCAMT.Caption = ""
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    M_EDIT = False
    cmdRefresh.Enabled = False
    cmdexit.Enabled = True
    CMDPRINT.Enabled = False
    cmdexit.Enabled = True
    TXTSLNO.Enabled = True
    FRMEHEAD.Enabled = True
    TXTPATIENT.SetFocus
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Tag = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblcredit.Caption = "1"
    OPTAUTOMATIC.Value = True

End Function

Private Sub TXTPATIENT_GotFocus()
    TXTPATIENT.SelStart = 0
    TXTPATIENT.SelLength = Len(TXTPATIENT.Text)
End Sub

Private Sub TXTPATIENT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(TXTPATIENT.Text) = "" Then TXTPATIENT.Text = "CASH"
            TXTDOCTOR.Enabled = True
            TXTDOCTOR.SetFocus
        Case vbKeyEscape
            If grdsales.Rows = 1 Then
                txtBillNo.Enabled = True
                txtBillNo.SetFocus
                Exit Sub
            ElseIf grdsales.Rows > 1 Then
            'If grdsales.Rows > 1 Then
                CMDPRINT.Enabled = True
                cmdRefresh.Enabled = True
                CMDPRINT.SetFocus
            End If
    End Select
End Sub

Private Sub TXTPATIENT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), vbKey0 To vbKey9
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub TXTDOCTOR_GotFocus()
    TXTDOCTOR.SelStart = 0
    TXTDOCTOR.SelLength = Len(TXTDOCTOR.Text)
End Sub

Private Sub TXTDOCTOR_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error GoTo eRRHAND
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(TXTDOCTOR.Text) = "" Then Exit Sub
            TXTSLNO.Text = grdsales.Rows
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTPATIENT.Enabled = True
            TXTPATIENT.SetFocus
    End Select
    Exit Sub

eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTDOCTOR_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Function Addnewbatch()
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * from [RTRXFILE]", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTRTRXFILE.AddNew
    RSTRTRXFILE!TRX_TYPE = "OP"
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(Val(VCH_NO)) From RTRXFILE WHERE TRX_TYPE = 'OP'", db, adOpenForwardOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        RSTRTRXFILE!VCH_NO = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    RSTRTRXFILE!VCH_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!LINE_NO = Val(TXTSLNO.Text)
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT CATEGORY FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTRTRXFILE!CATEGORY = IIf(IsNull(RSTITEMMAST!CATEGORY), "MEDICINE", RSTITEMMAST!CATEGORY)
    Else
        RSTRTRXFILE!CATEGORY = "MEDICINE"
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    RSTRTRXFILE!ITEM_CODE = Trim(TXTITEMCODE.Text)
    RSTRTRXFILE!ITEM_NAME = Trim(TXTPRODUCT.Text)
    RSTRTRXFILE!QTY = 0
    RSTRTRXFILE!ITEM_COST = 0
    RSTRTRXFILE!MRP = Val(TxtMRP.Text)
    RSTRTRXFILE!PTR = Val(TXTRATE.Text)
    RSTRTRXFILE!SALES_PRICE = Val(TXTRATE.Text) + Val(TXTTAX.Text) / 100
    RSTRTRXFILE!SALES_TAX = Val(TXTTAX.Text)
    RSTRTRXFILE!UNIT = Val(TXTUNIT.Text)
    RSTRTRXFILE!VCH_DESC = "Opening Balance"
    RSTRTRXFILE!REF_NO = Trim(txtBatch.Text)
    RSTRTRXFILE!ISSUE_QTY = 0
    RSTRTRXFILE!CST = 0
    RSTRTRXFILE!BAL_QTY = 0
    RSTRTRXFILE!TRX_TOTAL = 0
    RSTRTRXFILE!LINE_DISC = 0
    RSTRTRXFILE!SCHEME = 0
    RSTRTRXFILE!EXP_DATE = Null 'IIf(TXTEXPIRY.Text = "/", Null, Format(TXTEXPIRY.Text, "dd/mm/yyyy"))
    RSTRTRXFILE!FREE_QTY = 0
    RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!C_USER_ID = "SM"
    RSTRTRXFILE!M_USER_ID = ""
    RSTRTRXFILE!CHECK_FLAG = ""
    RSTRTRXFILE!PINV = "OP. Bal"
    RSTRTRXFILE!M_USER_ID = ""
    RSTRTRXFILE.Update
    
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
End Function

Private Sub TXTUNIT_GotFocus()
    TXTUNIT.SelStart = 0
    TXTUNIT.SelLength = Len(TXTUNIT.Text)
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.Text) = 0 Then Exit Sub
            
            TXTUNIT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTCATEGORY.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTUNIT.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTUNIT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Function PRINTSLIP()
    
    Dim i As Integer
    Dim Total_Amt As Double
    Dim num As Currency

    'Rs.Open strSQL, cnn
    
    '//NOTE : Report file name should never contain blank space.
    On Error GoTo CLOSEFILE
    Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    End If
    On Error GoTo eRRHAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold


    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    'If (Trim(txtremarks.Text) <> "") Then Print #1, Chr(27) & Chr(67) & Chr(0) & Space(48) & AlignRight(Trim(txtremarks.Text), 30)
    'Print #1, Chr(27) & Chr(71) & Chr(10) & _
        Space(7) & Chr(14) & Chr(15) & AlignLeft("ESTIMATE", 30) & _
        Chr(27) & Chr(72)
        
        
    Print #1, Chr(27) & Chr(71) & Chr(10) & _
      Space(7) & Chr(14) & Chr(15) & AlignLeft("ESTIMATE", 30) & _
      Chr(27) & Chr(72)
    Print #1, Chr(13)
    Print #1, Space(7) & "No. " & Trim(LBLSLIP.Caption) & Chr(27) & Chr(72) & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 57) '& Space(2) & LBLTIME.Caption
    Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Patient: " & Trim(TXTPATIENT.Text), 38) & Space(4) & AlignRight("Doctor: " & TXTDOCTOR.Text, 29)

    Print #1, Space(7) & RepeatString("-", 73)
    Print #1, Space(7) & AlignLeft("Description", 35) & Space(5) & _
            AlignLeft("MFR", 8) & Space(2) & _
            AlignLeft("Rate", 8) & Space(1) & _
            AlignLeft("Qty", 4) & Space(1) & _
            AlignLeft("Amount", 6) & _
            Chr(27) & Chr(72)  '//Bold Ends

    Print #1, Space(7) & RepeatString("-", 73)
    
    Total_Amt = 0
    For i = 1 To grdsales.Rows - 1
        Print #1, Space(7) & AlignLeft(grdsales.TextMatrix(i, 2), 40) & _
            AlignLeft(grdsales.TextMatrix(i, 17), 9) & Space(0) & _
            AlignRight(Format(Round(grdsales.TextMatrix(i, 6), 2), "0.00"), 6) & _
            AlignRight(grdsales.TextMatrix(i, 3), 6) & Space(0) & _
            AlignRight(Format(grdsales.TextMatrix(i, 11), "0.00"), 10) & _
            Chr(27) & Chr(72)  '//Bold Ends
            Total_Amt = Total_Amt + Val(grdsales.TextMatrix(i, 11))
    Next i

    Print #1, Space(7) & AlignRight("-------------", 73)
    
    If Val(LBLDISCAMT.Caption) > 0 Then
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("BILL AMOUNT ", 57) & AlignRight((Format(lbltotalwodiscount.Caption, "####.00")), 10)
        Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("DISC AMOUNT ", 57) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 10)
    End If

    Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Round(Total_Amt - Val(LBLDISCAMT.Caption), 0), "####.00")), 9)
    num = CCur(Round(Total_Amt - Val(LBLDISCAMT.Caption), 0))
    
    Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Rupees " & Words_1_all(num), 40)
    Print #1, Chr(27) & Chr(72) & Space(7) & "E&OE, Dispensed medicines cannot be taken back or exchanged"

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
        
    cmdexit.Enabled = False
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTUNIT.Enabled = False
    
    TXTTAX.Enabled = False
    TXTEXPIRY.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    
    Call DOSPRINTING
    Exit Function

eRRHAND:
    MsgBox Err.Description
End Function

Public Function PRINTBILL()
    
    Dim i As Integer
    Dim Total_Amt As Double
    Dim num As Currency

    'Rs.Open strSQL, cnn

    '//NOTE : Report file name should never contain blank space.
    On Error GoTo CLOSEFILE
    Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    End If
    On Error GoTo eRRHAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold


    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    If optoushadi.Value = True Then
        
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "Tin No. 32511150406" & Space(38) & "Phone: 2262611"
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(43) & "Form 8B"
        Print #1, Chr(27) & Chr(71) & Chr(10) & _
          Space(7) & Chr(14) & Chr(15) & Space(15) & "OUSHADHI" & _
          Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "The Pharmaceutical Corporation(IM), Kerala Ltd. H.O: Trissur - 14"
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "(A Government of Kerala Undertaking)"
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "D.S Depot: Alappuzha - 11"
        Print #1, Chr(13)
        ''Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Patient: " & Trim(TXTPATIENT.Text), 38) & Space(4) & AlignRight("Doctor: " & TXTDOCTOR.Text, 29)
        Print #1, Space(7) & "Invoice No. " & Trim(txtBillNo.Text) & Chr(27) & Chr(72) & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 57) '& Space(2) & LBLTIME.Caption
        
        
    ElseIf optothers.Value = True Then
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "Tin No. 32511150406" & Space(38) & "Phone: 2262611"
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(43) & "FORM 8B"
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(42) & "CASH MEMO"
        Print #1, Chr(27) & Chr(71) & Chr(10) & _
          Space(7) & Chr(14) & Chr(15) & Space(9) & "Oushadhis' Pharmacy" & _
          Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(40) & "Alappuzha - 11"
        Print #1, Chr(13)
        ''Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Patient: " & Trim(TXTPATIENT.Text), 38) & Space(4) & AlignRight("Doctor: " & TXTDOCTOR.Text, 29)
        Print #1, Space(7) & "Invoice No. P" & Trim(txtBillNo.Text) & Chr(27) & Chr(72) & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 57) '& Space(2) & LBLTIME.Caption
        Print #1, Chr(13)
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "OP No."
    End If
    
    'If Weekday(Date) = 1 Then TXTINVDATE.TEXT = DateAdd("d", 1, TXTINVDATE.TEXT)
    'TXTINVDATE.TEXT = Date
    Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Patient: " & Trim(TXTPATIENT.Text), 38) & Space(4) & AlignRight("Doctor: " & TXTDOCTOR.Text, 29)

    Print #1, Space(7) & RepeatString("-", 73)
    Print #1, Space(7) & AlignLeft("Code", 4) & Space(1); AlignLeft("Name of Commodity", 35) & Space(5) & _
            AlignLeft("Unit", 6) & Space(0) & _
            AlignLeft("Rate", 6) & Space(2) & _
            AlignLeft("Qty", 6) & Space(1) & _
            AlignLeft("Amount", 10) & _
            Chr(27) & Chr(72)  '//Bold Ends

    Print #1, Space(7) & RepeatString("-", 73)
    
    Total_Amt = 0
    For i = 1 To grdsales.Rows - 1
        Print #1, Space(12) & AlignLeft(grdsales.TextMatrix(i, 2), 40) & _
            AlignRight(grdsales.TextMatrix(i, 4), 4) & Space(1) & _
            AlignRight(Format(Round(grdsales.TextMatrix(i, 6), 2), "0.00"), 8) & _
            AlignRight(grdsales.TextMatrix(i, 3), 4) & Space(1) & _
            AlignRight(Format(grdsales.TextMatrix(i, 11), "0.00"), 10) & _
            Chr(27) & Chr(72)  '//Bold Ends
            Total_Amt = Total_Amt + Val(grdsales.TextMatrix(i, 11))
    Next i
    
    Print #1, Space(7) & AlignRight("-------------", 73)
    If Val(LBLDISCAMT.Caption) > 0 Then
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("BILL AMOUNT ", 57) & AlignRight((Format(lbltotalwodiscount.Caption, "####.00")), 10)
        Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("DISC AMOUNT ", 57) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 10)
    End If

    Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Round(Total_Amt - Val(LBLDISCAMT.Caption), 0), "####.00")), 9)
    num = CCur(Round(Total_Amt - Val(LBLDISCAMT.Caption), 0))
    
    Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Rupees " & Words_1_all(num), 40)
    'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
    Print #1, Chr(27) & Chr(67) & Chr(0)
    'Print #1, Chr(27) & Chr(72) & Space(8) & AlignRight("Pharmacist:" & RSTCOMPANY!EMAIL_ADD, 48)
    Print #1, Space(7) & RepeatString("-", 73)
    Print #1, Chr(27) & Chr(72) & Space(7) & "E & O E. Medicine once dispensed cannot be taken back or exchanged"
    Print #1, Chr(27) & Chr(72) & Space(7) & "Sign:"
        
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
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Close #1 '//Closing the file
        
    cmdexit.Enabled = False
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTUNIT.Enabled = False
    
    TXTTAX.Enabled = False
    TXTEXPIRY.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    Call DOSPRINTING
    Exit Function

eRRHAND:
    MsgBox Err.Description

End Function


' Return words for this value between 1 and 999.
Private Function Words_1_999(ByVal num As Integer) As String
Dim hundreds As Integer
Dim remainder As Integer
Dim result As String

    hundreds = num \ 100
    remainder = num - hundreds * 100

    If hundreds > 0 Then
        result = Words_1_19(hundreds) & " hundred "
    End If

    If remainder > 0 Then
        result = result & Words_1_99(remainder)
    End If

    Words_1_999 = Trim$(result)
End Function
' Return a word for this value between 1 and 99.
Private Function Words_1_99(ByVal num As Integer) As String
Dim result As String
Dim tens As Integer

    tens = num \ 10

    If tens <= 1 Then
        ' 1 <= num <= 19
        result = result & " " & Words_1_19(num)
    Else
        ' 20 <= num
        ' Get the tens digit word.
        Select Case tens
            Case 2
                result = "twenty"
            Case 3
                result = "thirty"
            Case 4
                result = "forty"
            Case 5
                result = "fifty"
            Case 6
                result = "sixty"
            Case 7
                result = "seventy"
            Case 8
                result = "eighty"
            Case 9
                result = "ninety"
        End Select

        ' Add the ones digit number.
        result = result & " " & Words_1_19(num - tens * 10)
    End If

    Words_1_99 = Trim$(result)
End Function
' Return a word for this value between 1 and 19.
Private Function Words_1_19(ByVal num As Integer) As String
    Select Case num
        Case 1
            Words_1_19 = "one"
        Case 2
            Words_1_19 = "two"
        Case 3
            Words_1_19 = "three"
        Case 4
            Words_1_19 = "four"
        Case 5
            Words_1_19 = "five"
        Case 6
            Words_1_19 = "six"
        Case 7
            Words_1_19 = "seven"
        Case 8
            Words_1_19 = "eight"
        Case 9
            Words_1_19 = "nine"
        Case 10
            Words_1_19 = "ten"
        Case 11
            Words_1_19 = "eleven"
        Case 12
            Words_1_19 = "twelve"
        Case 13
            Words_1_19 = "thirteen"
        Case 14
            Words_1_19 = "fourteen"
        Case 15
            Words_1_19 = "fifteen"
        Case 16
            Words_1_19 = "sixteen"
        Case 17
            Words_1_19 = "seventeen"
        Case 18
            Words_1_19 = "eightteen"
        Case 19
            Words_1_19 = "nineteen"
    End Select
End Function
' Return a string of words to represent the
' integer part of this value.
Private Function Words_1_all(ByVal num As Currency) As String
Dim power_value(1 To 5) As Currency
Dim power_name(1 To 5) As String
Dim digits As Integer
Dim result As String
Dim i As Integer

    ' Initialize the power names and values.
    power_name(1) = "trillion": power_value(1) = 1000000000000#
    power_name(2) = "billion":  power_value(2) = 1000000000
    power_name(3) = "million":  power_value(3) = 1000000
    power_name(4) = "thousand": power_value(4) = 1000
    power_name(5) = "":         power_value(5) = 1

    For i = 1 To 5
        ' See if we have digits in this range.
        If num >= power_value(i) Then
            ' Get the digits.
            digits = Int(num / power_value(i))

            ' Add the digits to the result.
            If Len(result) > 0 Then result = result & ", "
            result = result & _
                Words_1_999(digits) & _
                " " & power_name(i)

            ' Get the number without these digits.
            num = num - digits * power_value(i)
        End If
    Next i

    Words_1_all = Trim$(result)
End Function
' Return a string of words to represent this
' currency value in dollars and cents.
Private Function Words_Money(ByVal num As Currency) As String
Dim dollars As Currency
Dim cents As Integer
Dim dollars_result As String
Dim cents_result As String

    ' Dollars.
    dollars = Int(num)
    dollars_result = Words_1_all(dollars)
    If Len(dollars_result) = 0 Then dollars_result = "zero"

    If dollars_result = "one" Then
        dollars_result = dollars_result & " rupee"
    Else
        dollars_result = dollars_result & " rupees"
    End If

    ' Cents.
    cents = CInt((num - dollars) * 100#)
    cents_result = Words_1_all(cents)
    If Len(cents_result) = 0 Then cents_result = "zero"

    If cents_result = "one" Then
        cents_result = cents_result & " paise"
    Else
        cents_result = cents_result & " paise"
    End If

    ' Combine the results.
    Words_Money = dollars_result & _
        " and " & cents_result
End Function

Public Function AutoPRINTBILL()
    
    Dim i, Oushadi, others, N As Integer
    Dim Total_Amt, DISCOUNT As Double
    Dim num As Currency

    'Rs.Open strSQL, cnn

    '//NOTE : Report file name should never contain blank space.
    On Error GoTo CLOSEFILE
    Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    End If
    On Error GoTo eRRHAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold

    Oushadi = 0
    others = 0
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    For i = 1 To grdsales.Rows - 1
        ''''If grdsales.TextMatrix(i, 19) = "Oushadhi" Or grdsales.TextMatrix(i, 19) = "oushadhi" Or grdsales.TextMatrix(i, 19) = "OUSHADHI" Then
        If UCase(grdsales.TextMatrix(i, 19)) = "OUSHADHI" Then
            Oushadi = Oushadi + 1
        Else
            others = others + 1
        End If
    Next i
    
    For N = 1 To 2
        If Oushadi > 0 Then
            
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "Tin No. 32511150406" & Space(38) & "Phone: 2262611"
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(43) & "Form 8B"
            Print #1, Chr(27) & Chr(71) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & Space(15) & "OUSHADHI" & _
              Chr(27) & Chr(72)
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "The Pharmaceutical Corporation(IM), Kerala Ltd. H.O: Trissur - 14"
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "(A Government of Kerala Undertaking)"
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "D.S Depot: Alappuzha - 11"
            Print #1, Chr(13)
            ''Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Patient: " & Trim(TXTPATIENT.Text), 38) & Space(4) & AlignRight("Doctor: " & TXTDOCTOR.Text, 29)
            Print #1, Space(7) & "Invoice No. " & Trim(txtBillNo.Text) & Chr(27) & Chr(72) & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 57) '& Space(2) & LBLTIME.Caption
            Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Patient: " & Trim(TXTPATIENT.Text), 38) & Space(4) & AlignRight("Doctor: " & TXTDOCTOR.Text, 29)

            Print #1, Space(7) & RepeatString("-", 73)
            Print #1, Space(7) & AlignLeft("Code", 4) & Space(1); AlignLeft("Name of Commodity", 35) & Space(5) & _
                    AlignLeft("Unit", 6) & Space(0) & _
                    AlignLeft("Rate", 6) & Space(2) & _
                    AlignLeft("Qty", 6) & Space(1) & _
                    AlignLeft("Amount", 10) & _
                    Chr(27) & Chr(72)  '//Bold Ends
        
            Print #1, Space(7) & RepeatString("-", 73)
            
            Total_Amt = 0
            For i = 1 To grdsales.Rows - 1
                If UCase(grdsales.TextMatrix(i, 19)) = "OUSHADHI" Then
                    Print #1, Space(12) & AlignLeft(grdsales.TextMatrix(i, 2), 40) & _
                        AlignRight(grdsales.TextMatrix(i, 4), 4) & Space(1) & _
                        AlignRight(Format(Round(grdsales.TextMatrix(i, 6), 2), "0.00"), 8) & _
                        AlignRight(grdsales.TextMatrix(i, 3), 4) & Space(1) & _
                        AlignRight(Format(grdsales.TextMatrix(i, 11), "0.00"), 10) & _
                        Chr(27) & Chr(72)  '//Bold Ends
                        Total_Amt = Total_Amt + Val(grdsales.TextMatrix(i, 11))
                End If
            Next i
            
            Print #1, Space(7) & AlignRight("-------------", 73)
            If Val(LBLDISCAMT.Caption) > 0 Then
                'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("BILL AMOUNT ", 57) & AlignRight((Format(lbltotalwodiscount.Caption, "####.00")), 10)
                Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("DISC AMOUNT ", 57) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 10)
            End If
        
            Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Round(Total_Amt - Val(LBLDISCAMT.Caption), 0), "####.00")), 9)
            num = CCur(Round(Total_Amt - Val(LBLDISCAMT.Caption), 0))
            
            Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Rupees " & Words_1_all(num), 40)
            'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
            Print #1, Chr(27) & Chr(67) & Chr(0)
            'Print #1, Chr(27) & Chr(72) & Space(8) & AlignRight("Pharmacist:" & RSTCOMPANY!EMAIL_ADD, 48)
            Print #1, Space(7) & RepeatString("-", 73)
            Print #1, Chr(27) & Chr(72) & Space(7) & "E & O E. Medicine once dispensed cannot be taken back or exchanged"
            Print #1, Chr(27) & Chr(72) & Space(7) & "Sign:"

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
            Oushadi = 0
        ElseIf others > 0 Then

            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "Tin No. 32511150406" & Space(38) & "Phone: 2262611"
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(43) & "FORM 8B"
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(42) & "CASH MEMO"
            Print #1, Chr(27) & Chr(71) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & Space(9) & "Oushadhis' Pharmacy" & _
              Chr(27) & Chr(72)
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(40) & "Alappuzha - 11"
            Print #1, Chr(13)
            ''Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Patient: " & Trim(TXTPATIENT.Text), 38) & Space(4) & AlignRight("Doctor: " & TXTDOCTOR.Text, 29)
            Print #1, Space(7) & "Invoice No. P" & Trim(txtBillNo.Text) & Chr(27) & Chr(72) & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 57) '& Space(2) & LBLTIME.Caption
            Print #1, Chr(13)
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & "OP No."
            Print #1, Space(7) & RepeatString("-", 73)
            Print #1, Space(7) & AlignLeft("Code", 4) & Space(1); AlignLeft("Name of Commodity", 35) & Space(5) & _
                    AlignLeft("Unit", 6) & Space(0) & _
                    AlignLeft("Rate", 6) & Space(2) & _
                    AlignLeft("Qty", 6) & Space(1) & _
                    AlignLeft("Amount", 10) & _
                    Chr(27) & Chr(72)  '//Bold Ends
        
            Print #1, Space(7) & RepeatString("-", 73)
            
            Total_Amt = 0
            For i = 1 To grdsales.Rows - 1
                If UCase(grdsales.TextMatrix(i, 19)) <> "OUSHADHI" Then
                    Print #1, Space(12) & AlignLeft(grdsales.TextMatrix(i, 2), 40) & _
                        AlignRight(grdsales.TextMatrix(i, 4), 4) & Space(1) & _
                        AlignRight(Format(Round(grdsales.TextMatrix(i, 6), 2), "0.00"), 8) & _
                        AlignRight(grdsales.TextMatrix(i, 3), 4) & Space(1) & _
                        AlignRight(Format(grdsales.TextMatrix(i, 11), "0.00"), 10) & _
                        Chr(27) & Chr(72)  '//Bold Ends
                        Total_Amt = Total_Amt + Val(grdsales.TextMatrix(i, 11))
                End If
            Next i
            
            Print #1, Space(7) & AlignRight("-------------", 73)
            If Val(LBLDISCAMT.Caption) > 0 Then
                'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("BILL AMOUNT ", 57) & AlignRight((Format(lbltotalwodiscount.Caption, "####.00")), 10)
                Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("DISC AMOUNT ", 57) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 10)
            End If
        
            Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Round(Total_Amt - Val(LBLDISCAMT.Caption), 0), "####.00")), 9)
            num = CCur(Round(Total_Amt - Val(LBLDISCAMT.Caption), 0))
            
            Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Rupees " & Words_1_all(num), 40)
            'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
            Print #1, Chr(27) & Chr(67) & Chr(0)
            'Print #1, Chr(27) & Chr(72) & Space(8) & AlignRight("Pharmacist:" & RSTCOMPANY!EMAIL_ADD, 48)
            Print #1, Space(7) & RepeatString("-", 73)
            Print #1, Chr(27) & Chr(72) & Space(7) & "E & O E. Medicine once dispensed cannot be taken back or exchanged"
            Print #1, Chr(27) & Chr(72) & Space(7) & "Sign:"
                
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
            others = 0
        End If
    Next N
    
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Close #1 '//Closing the file
    cmdexit.Enabled = False
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTUNIT.Enabled = False
    
    TXTTAX.Enabled = False
    TXTEXPIRY.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    Call DOSPRINTING
    Exit Function

eRRHAND:
    MsgBox Err.Description

End Function


Private Function DOSPRINTING()
''''    On Error GoTo CLOSEFILE
''''    Open App.Path & "\repo.bat" For Output As #1 '//Creating Batch file
''''CLOSEFILE:
''''    If Err.Number = 55 Then
''''        Close #1
''''        Open App.Path & "\repo.bat" For Output As #1 '//Creating Batch file
''''    End If
''''    On Error GoTo eRRHAND
''''
''''    Print #1, "TYPE " & App.Path & "\Report.txt > PRN"
''''    Print #1, "EXIT"
''''    Close #1
''''
    '//HERE write the proper path where your command.com file exist
    'Shell "C:\WINDOW\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
    Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
End Function

Private Sub Timer_Timer()
    'Timer.Interval = 30
    If lblhead(1).Visible = True Then
        lblhead(1).Visible = False
    Else
        lblhead(1).Visible = True
    End If
End Sub
