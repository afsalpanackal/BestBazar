VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMSALE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WHOLE SALE INVOICE !!!!!"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   Icon            =   "Frmmain.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   12270
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   12630
      TabIndex        =   76
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
      TabIndex        =   55
      Top             =   3645
      Visible         =   0   'False
      Width           =   6030
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   2835
         Left            =   75
         TabIndex        =   56
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
      TabIndex        =   51
      Top             =   3645
      Visible         =   0   'False
      Width           =   5835
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2535
         Left            =   90
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   105
         Visible         =   0   'False
         Width           =   2610
      End
   End
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   -135
      TabIndex        =   19
      Top             =   -15
      Width           =   12375
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
         TabIndex        =   59
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00FFFFC0&
         Height          =   1725
         Left            =   210
         TabIndex        =   20
         Top             =   45
         Width           =   12165
         Begin VB.TextBox txtcrdays 
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
            Left            =   9900
            TabIndex        =   108
            Top             =   210
            Width           =   885
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
            Left            =   1545
            TabIndex        =   0
            Top             =   615
            Width           =   3735
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
            Left            =   1545
            TabIndex        =   17
            Top             =   195
            Visible         =   0   'False
            Width           =   885
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   345
            Left            =   3840
            TabIndex        =   60
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
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1545
            TabIndex        =   1
            Top             =   960
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
         Begin VB.Label LBLDLNO2 
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   210
            TabIndex        =   110
            Top             =   1260
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Days"
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
            Index           =   32
            Left            =   8670
            TabIndex        =   109
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            Index           =   22
            Left            =   5325
            TabIndex        =   69
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "TIN"
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
            Index           =   21
            Left            =   8325
            TabIndex        =   68
            Top             =   1365
            Width           =   360
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DL No."
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
            Left            =   5370
            TabIndex        =   67
            Top             =   1365
            Width           =   645
         End
         Begin VB.Label lbltin 
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   8670
            TabIndex        =   66
            Top             =   1335
            Width           =   2145
         End
         Begin VB.Label lbldlno 
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6150
            TabIndex        =   65
            Top             =   1335
            Width           =   2130
         End
         Begin VB.Label lbladdress 
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
            ForeColor       =   &H80000008&
            Height          =   660
            Left            =   6150
            TabIndex        =   64
            Top             =   630
            Width           =   4635
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CUSTOMER"
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
            Left            =   90
            TabIndex        =   62
            Top             =   675
            Width           =   1230
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
            Left            =   1545
            TabIndex        =   21
            Top             =   210
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Height          =   6465
         Left            =   210
         TabIndex        =   26
         Top             =   1695
         Width           =   12180
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
            TabIndex        =   100
            Top             =   6030
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
            Left            =   8145
            TabIndex        =   99
            Top             =   6060
            Width           =   930
         End
         Begin VB.OptionButton OPTDISCPERCENT 
            BackColor       =   &H00FFFFC0&
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
            Left            =   4995
            TabIndex        =   98
            Top             =   6060
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton OptDiscAmt 
            BackColor       =   &H00FFFFC0&
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
            Left            =   6495
            TabIndex        =   97
            Top             =   6060
            Width           =   1605
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFC0&
            Height          =   6255
            Left            =   10290
            TabIndex        =   27
            Top             =   135
            Width           =   1815
            Begin VB.CommandButton CMDSALERETURN 
               Appearance      =   0  'Flat
               BackColor       =   &H000080FF&
               Caption         =   "Add Reurned Items"
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
               Height          =   525
               Left            =   60
               MaskColor       =   &H008080FF&
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   5085
               Width           =   1695
            End
            Begin VB.CommandButton CMDDELIVERY 
               BackColor       =   &H00FF8080&
               Caption         =   "Add Delivered Items"
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
               Height          =   525
               Left            =   45
               Style           =   1  'Graphical
               TabIndex        =   75
               Top             =   5670
               Width           =   1695
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "WHOLESALE BILL"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   690
               Index           =   26
               Left            =   105
               TabIndex        =   111
               Top             =   3390
               Width           =   1605
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
               Left            =   195
               TabIndex        =   102
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
               TabIndex        =   101
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
               TabIndex        =   96
               Top             =   1530
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
               TabIndex        =   95
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
               TabIndex        =   83
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
               TabIndex        =   82
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
               TabIndex        =   80
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
               TabIndex        =   79
               Top             =   4035
               Visible         =   0   'False
               Width           =   1425
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
               TabIndex        =   73
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
               TabIndex        =   72
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
               TabIndex        =   71
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
               TabIndex        =   70
               Top             =   795
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
               TabIndex        =   63
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
               Left            =   165
               TabIndex        =   29
               Top             =   90
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
               TabIndex        =   28
               Top             =   315
               Width           =   1440
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   5730
            Left            =   90
            TabIndex        =   18
            Top             =   270
            Width           =   10140
            _ExtentX        =   17886
            _ExtentY        =   10107
            _Version        =   393216
            Rows            =   1
            Cols            =   24
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
         TabIndex        =   50
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
         BackColor       =   &H00FFFFC0&
         Height          =   1545
         Left            =   210
         TabIndex        =   30
         Top             =   8070
         Width           =   12180
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
            Left            =   4275
            MaxLength       =   6
            TabIndex        =   112
            Top             =   765
            Visible         =   0   'False
            Width           =   795
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
            Left            =   3135
            MaxLength       =   6
            TabIndex        =   107
            Top             =   795
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
            Left            =   10500
            MaxLength       =   6
            TabIndex        =   105
            Top             =   450
            Width           =   645
         End
         Begin VB.OptionButton optnet 
            BackColor       =   &H00FFFFC0&
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
            Height          =   225
            Left            =   9540
            TabIndex        =   104
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
            TabIndex        =   103
            Top             =   795
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00FFFFC0&
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
            Height          =   225
            Left            =   8475
            TabIndex        =   93
            Top             =   780
            Width           =   1005
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00FFFFC0&
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
            Height          =   240
            Left            =   6630
            TabIndex        =   92
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
            Left            =   6450
            MaxLength       =   6
            TabIndex        =   90
            Top             =   450
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
            Left            =   4665
            MaxLength       =   7
            TabIndex        =   88
            Top             =   450
            Width           =   555
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
            Height          =   465
            Left            =   555
            TabIndex        =   87
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
            Height          =   465
            Left            =   90
            TabIndex        =   81
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
            Height          =   465
            Left            =   10335
            TabIndex        =   74
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
            Left            =   5235
            MaxLength       =   6
            TabIndex        =   57
            Top             =   450
            Width           =   660
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
            TabIndex        =   11
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
            TabIndex        =   2
            Top             =   450
            Width           =   390
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
            Left            =   465
            TabIndex        =   3
            Top             =   450
            Width           =   3570
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
            Left            =   4050
            MaxLength       =   7
            TabIndex        =   4
            Top             =   450
            Width           =   600
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
            Left            =   7110
            MaxLength       =   6
            TabIndex        =   5
            Top             =   450
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
            Left            =   5910
            MaxLength       =   4
            TabIndex        =   6
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
            Left            =   9840
            MaxLength       =   4
            TabIndex        =   9
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
            Height          =   465
            Left            =   6495
            TabIndex        =   14
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
            Height          =   465
            Left            =   8895
            TabIndex        =   16
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
            Height          =   465
            Left            =   5295
            TabIndex        =   13
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
            Height          =   465
            Left            =   4095
            TabIndex        =   12
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
            TabIndex        =   35
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
            Left            =   8880
            MaxLength       =   15
            TabIndex        =   8
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            Height          =   465
            Left            =   7695
            TabIndex        =   15
            Top             =   1035
            Width           =   1125
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   300
            Left            =   7770
            TabIndex        =   7
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
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "PTR"
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
            Left            =   10500
            TabIndex        =   106
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblP_Rate 
            Caption         =   "0"
            Height          =   390
            Left            =   10935
            TabIndex        =   94
            Top             =   960
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "PTS"
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
            Left            =   6450
            TabIndex        =   91
            Top             =   225
            Width           =   630
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
            Left            =   4665
            TabIndex        =   89
            Top             =   225
            Width           =   555
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
            TabIndex        =   85
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
            Left            =   5235
            TabIndex        =   58
            Top             =   225
            Width           =   660
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
            TabIndex        =   49
            Top             =   225
            Width           =   390
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
            Left            =   465
            TabIndex        =   48
            Top             =   225
            Width           =   3570
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
            Left            =   4050
            TabIndex        =   47
            Top             =   225
            Width           =   600
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
            Left            =   7110
            TabIndex        =   46
            Top             =   225
            Width           =   630
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
            Left            =   5910
            TabIndex        =   45
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
            Left            =   9840
            TabIndex        =   44
            Top             =   240
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
            Left            =   11160
            TabIndex        =   43
            Top             =   240
            Width           =   975
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
            TabIndex        =   42
            Top             =   1260
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
            Left            =   7770
            TabIndex        =   41
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
            Left            =   8880
            TabIndex        =   40
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
            Left            =   11160
            TabIndex        =   10
            Top             =   450
            Width           =   975
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
            Top             =   1275
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   -15
      TabIndex        =   86
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11355
      TabIndex        =   78
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   77
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMSALE"
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
Dim cr_days As Boolean
Dim M_EDIT As Boolean

Private Sub CMDSALERETURN_Click()
    Dim RSTMFGR As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
    'If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM SALERETURN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y'", db2, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        grdsales.Rows = grdsales.Rows + 1
        grdsales.FixedRows = 1
        grdsales.TextMatrix(grdsales.Rows - 1, 0) = grdsales.Rows
        grdsales.TextMatrix(grdsales.Rows - 1, 1) = RSTTRXFILE!ITEM_CODE
        grdsales.TextMatrix(grdsales.Rows - 1, 2) = RSTTRXFILE!ITEM_NAME
        grdsales.TextMatrix(grdsales.Rows - 1, 3) = RSTTRXFILE!QTY
        grdsales.TextMatrix(grdsales.Rows - 1, 4) = RSTTRXFILE!UNIT
        grdsales.TextMatrix(grdsales.Rows - 1, 5) = Format(RSTTRXFILE!MRP, ".000")
        grdsales.TextMatrix(grdsales.Rows - 1, 6) = Format(RSTTRXFILE!PTR, ".000")
        grdsales.TextMatrix(grdsales.Rows - 1, 7) = Format(RSTTRXFILE!SALES_PRICE, ".000")
        grdsales.TextMatrix(grdsales.Rows - 1, 8) = Format(RSTTRXFILE!LINE_DISC, ".00")
        grdsales.TextMatrix(grdsales.Rows - 1, 9) = Format(RSTTRXFILE!SALES_TAX, ".00")
        grdsales.TextMatrix(grdsales.Rows - 1, 10) = IIf(IsNull(RSTTRXFILE!REF_NO), "", RSTTRXFILE!REF_NO)
        grdsales.TextMatrix(grdsales.Rows - 1, 11) = IIf(IsNull(RSTTRXFILE!EXP_DATE), "", Format(RSTTRXFILE!EXP_DATE, "MM/YY"))
        grdsales.TextMatrix(grdsales.Rows - 1, 12) = Format(RSTTRXFILE!TRX_TOTAL, ".000")
        
        grdsales.TextMatrix(grdsales.Rows - 1, 13) = RSTTRXFILE!ITEM_CODE
        grdsales.TextMatrix(grdsales.Rows - 1, 14) = RSTTRXFILE!VCH_NO
        grdsales.TextMatrix(grdsales.Rows - 1, 15) = RSTTRXFILE!LINE_NO
        grdsales.TextMatrix(grdsales.Rows - 1, 16) = RSTTRXFILE!TRX_TYPE
        grdsales.TextMatrix(grdsales.Rows - 1, 17) = "N"
        Set RSTMFGR = New ADODB.Recordset
        RSTMFGR.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(grdsales.Rows - 1, 1)) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTMFGR.EOF And RSTMFGR.BOF) Then
            grdsales.TextMatrix(grdsales.Rows - 1, 18) = Trim(RSTMFGR!MANUFACTURER)
        End If
        RSTMFGR.Close
        Set RSTMFGR = Nothing
        grdsales.TextMatrix(grdsales.Rows - 1, 19) = "CN"
        grdsales.TextMatrix(grdsales.Rows - 1, 20) = "0"
        RSTTRXFILE!CHECK_FLAG = "Y"
        RSTTRXFILE!BILL_NO = Val(txtBillNo.Text)
        RSTTRXFILE!BILL_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
        CMDSALERETURN.Enabled = False
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))), 2)
            Case Else
                LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))), 2)
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
        End Select
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    OPTNET.Value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    
    TXTQTY.Text = ""
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    TXTRATE.Text = ""
    TXTRETAIL.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTPTR.Text = ""
    TxtFree.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtBatch.Text = ""
    TXTEXPIRY.Text = "  /  "
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    M_EDIT = False
    FRMEHEAD.Enabled = False
    Call COSTCALCULATION
    TXTSLNO.SetFocus
    If grdsales.Rows > 1 Then grdsales.TopRow = grdsales.Rows - 1
Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub cmdstockadjst_Click()
    FrmStkAdj.Show
    FrmStkAdj.SetFocus
End Sub

Private Sub DataList2_Click()
    Dim rstCustomer As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    On Error GoTo eRRHAND
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from [ACTMAST]  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = Trim(rstCustomer!ADDRESS)
        lbldlno.Caption = IIf(IsNull(rstCustomer!DL_NO), "", Trim(rstCustomer!DL_NO))
        LBLDLNO2.Caption = IIf(IsNull(rstCustomer!Remarks), "", Trim(rstCustomer!Remarks))
        lbltin.Caption = IIf(IsNull(Trim(rstCustomer!KGST)), "", rstCustomer!KGST)
    Else
        lbladdress.Caption = ""
        lbldlno.Caption = ""
        LBLDLNO2.Caption = ""
        lbltin.Caption = ""
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y' AND TRX_TYPE = 'SI'", db2, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        CMDDELIVERY.Enabled = True
    Else
        CMDDELIVERY.Enabled = False
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM SALERETURN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y'", db2, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        CMDSALERETURN.Enabled = True
    Else
        CMDSALERETURN.Enabled = False
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
        
    If cr_days = False Then
        Set rstCustomer = New ADODB.Recordset
        rstCustomer.Open "Select PYMT_PERIOD, ACT_CODE From ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
        If Not (rstCustomer.EOF And rstCustomer.BOF) Then
            txtcrdays.Text = IIf(IsNull(rstCustomer!PYMT_PERIOD), "", rstCustomer!PYMT_PERIOD)
        End If
        rstCustomer.Close
        Set rstCustomer = Nothing
    End If
    
    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text
    Exit Sub
    
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
                DataList2.SetFocus
                Exit Sub
            End If
            FRMEHEAD.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
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
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTPTR.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Format(Val(TXTRATE.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(TXTDISC.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Val(TXTTAX.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = IIf(TXTEXPIRY.Text = "  /  ", "", Trim(TXTEXPIRY.Text))
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Trim(TXTTRXTYPE.Text)
    
    If OPTVAT.Value = True And Val(TXTTAX.Text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "V"
    If OPTTaxMRP.Value = True And Val(TXTTAX.Text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "M"
    If Val(TXTTAX.Text) <= 0 Or OPTNET.Value = True Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "N"
  
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = Trim(RSTTRXFILE!MANUFACTURER)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Select Case LBLDNORCN.Caption
        Case "DN"
            grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = "DN"
        Case "CN"
            grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = "CN"
        Case Else
            grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = "B"
    End Select
    grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = Val(TxtFree.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 21) = Format(Val(TXTRETAIL.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 22) = Format(Val(TXTRETAILNOTAX.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 23) = Trim(TXTSALETYPE.Text)
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
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
        End Select
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
    OPTNET.Value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Text = ""
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    TXTRATE.Text = ""
    TXTRETAIL.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTPTR.Text = ""
    TxtFree.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtBatch.Text = ""
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lblP_Rate.Caption = "0"
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
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
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
        End Select
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
    Next i
    grdsales.Rows = grdsales.Rows - 1
    
    Call COSTCALCULATION
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    OPTNET.Value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    TXTRETAIL.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTPTR.Text = ""
    TxtFree.Text = ""
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    TXTTAX.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTDISC.Text = ""
    txtBatch.Text = ""
    LBLSUBTOTAL.Caption = ""
    LBLDNORCN.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
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
            TXTRATE.Text = ""
            TXTRETAIL.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTPTR.Text = ""
            TxtFree.Text = ""
            OPTNET.Value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
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
            TXTRATE.Enabled = False
            TXTRETAIL.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            TxtFree.Enabled = False
            TXTPTR.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub CMDDELIVERY_Click()
    Dim RSTMFGR As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
    'If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y' AND TRX_TYPE = 'SI'", db2, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        grdsales.Rows = grdsales.Rows + 1
        grdsales.FixedRows = 1
        grdsales.TextMatrix(grdsales.Rows - 1, 0) = grdsales.Rows
        grdsales.TextMatrix(grdsales.Rows - 1, 1) = RSTTRXFILE!ITEM_CODE
        grdsales.TextMatrix(grdsales.Rows - 1, 2) = RSTTRXFILE!ITEM_NAME
        grdsales.TextMatrix(grdsales.Rows - 1, 3) = RSTTRXFILE!QTY
        grdsales.TextMatrix(grdsales.Rows - 1, 4) = RSTTRXFILE!UNIT
        grdsales.TextMatrix(grdsales.Rows - 1, 5) = Format(RSTTRXFILE!MRP, ".000")
        grdsales.TextMatrix(grdsales.Rows - 1, 6) = Format(RSTTRXFILE!PTR, ".000")
        grdsales.TextMatrix(grdsales.Rows - 1, 7) = Format(RSTTRXFILE!SALES_PRICE, ".000")
        grdsales.TextMatrix(grdsales.Rows - 1, 8) = Format(RSTTRXFILE!LINE_DISC, ".00")
        grdsales.TextMatrix(grdsales.Rows - 1, 9) = Format(RSTTRXFILE!SALES_TAX, ".00")
        grdsales.TextMatrix(grdsales.Rows - 1, 10) = IIf(IsNull(RSTTRXFILE!REF_NO), "", RSTTRXFILE!REF_NO)
        grdsales.TextMatrix(grdsales.Rows - 1, 11) = IIf(IsNull(RSTTRXFILE!EXP_DATE), "", Format(RSTTRXFILE!EXP_DATE, "MM/YY"))
        grdsales.TextMatrix(grdsales.Rows - 1, 12) = Format(RSTTRXFILE!TRX_TOTAL, ".000")
    
        grdsales.TextMatrix(grdsales.Rows - 1, 13) = RSTTRXFILE!ITEM_CODE
        grdsales.TextMatrix(grdsales.Rows - 1, 14) = RSTTRXFILE!R_VCH_NO
        grdsales.TextMatrix(grdsales.Rows - 1, 15) = RSTTRXFILE!R_LINE_NO
        grdsales.TextMatrix(grdsales.Rows - 1, 16) = RSTTRXFILE!R_TRX_TYPE
        grdsales.TextMatrix(grdsales.Rows - 1, 17) = IIf(IsNull(RSTTRXFILE!FLAG), "N", RSTTRXFILE!FLAG)
        Set RSTMFGR = New ADODB.Recordset
        RSTMFGR.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(grdsales.Rows - 1, 1)) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTMFGR.EOF And RSTMFGR.BOF) Then
            grdsales.TextMatrix(grdsales.Rows - 1, 18) = Trim(RSTMFGR!MANUFACTURER)
        End If
        RSTMFGR.Close
        Set RSTMFGR = Nothing
        grdsales.TextMatrix(grdsales.Rows - 1, 19) = "DN"
        grdsales.TextMatrix(grdsales.Rows - 1, 20) = IIf(IsNull(RSTTRXFILE!FREE_QTY), 0, RSTTRXFILE!FREE_QTY)
        grdsales.TextMatrix(grdsales.Rows - 1, 21) = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
        grdsales.TextMatrix(grdsales.Rows - 1, 22) = IIf(IsNull(RSTTRXFILE!P_RETAILWOTAX), 0, RSTTRXFILE!P_RETAILWOTAX)
        grdsales.TextMatrix(grdsales.Rows - 1, 23) = IIf(IsNull(RSTTRXFILE!SALE_1_FLAG), "2", RSTTRXFILE!SALE_1_FLAG)
        RSTTRXFILE!CHECK_FLAG = "Y"
        RSTTRXFILE!BILL_NO = Val(txtBillNo.Text)
        RSTTRXFILE!BILL_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
        CMDDELIVERY.Enabled = False
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))), 2)
            Case Else
                LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))), 2)
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
        End Select
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    OPTNET.Value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    
    TXTQTY.Text = ""
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    TXTRATE.Text = ""
    TXTRETAIL.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTPTR.Text = ""
    TxtFree.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtBatch.Text = ""
    TXTEXPIRY.Text = "  /  "
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    M_EDIT = False
    FRMEHEAD.Enabled = False
    Call COSTCALCULATION
    TXTSLNO.SetFocus
    If grdsales.Rows > 1 Then grdsales.TopRow = grdsales.Rows - 1
Exit Sub
eRRHAND:
    MsgBox Err.Description
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
            TXTRATE.Text = ""
            TXTRETAIL.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTPTR.Text = ""
            TxtFree.Text = ""
            OPTNET.Value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            
            TXTVCHNO.Text = ""
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
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            TxtFree.Enabled = False
            TXTPTR.Enabled = False
            TXTRETAIL.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub cmdprint_Click()
    
    If grdsales.Rows = 1 Then Exit Sub
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
        DataList2.SetFocus
        Exit Sub
    End If
    
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
    
    Me.Enabled = False
    FRMDEBIT.Show
    
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Integer
    Dim CN As Integer
    Dim DN As Integer
    Dim B As Integer
    Dim num As Currency
    
    On Error GoTo eRRHAND
    If CMDDELIVERY.Enabled = True Then
        If (MsgBox("Delivered Items Available... Do you want to add these Items too...", vbYesNo, "SALES") = vbYes) Then CMDDELIVERY_Click
    End If
    
'    If CMDSALERETURN.Enabled = True Then
'        If (MsgBox("Returned Items Available... Do you want to add these Items too...", vbYesNo, "SALES") = vbYes) Then CMDSALERETURN_Click
'    End If
    
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
        RSTTRXFILE!VCH_DESC = "Issued to     " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 10)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
        Select Case grdsales.TextMatrix(i, 19)
            Case "DN"
                RSTTRXFILE!CST = 1
            Case "CN"
                RSTTRXFILE!CST = 2
            Case Else
                RSTTRXFILE!CST = 0
        End Select
        
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        If grdsales.TextMatrix(i, 11) = "" Then
            RSTTRXFILE!EXP_DATE = Null
        Else
            RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(i, 11))) & "/" & Trim(grdsales.TextMatrix(i, 11))
        End If
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT AREA FROM ACTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!Area = RSTITEMMAST!Area
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Call ReportGeneratION
    
    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)) - Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 2)), 2)
    num = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0))
    LBLFOT.Tag = "(Rupees " & Words_1_all(num) & " Only)"
    ReportNameVar = App.Path & "\rptbill.RPT"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "( {TRXFILE.TRX_TYPE}='SI' AND {TRXFILE.VCH_NO}= " & Val(txtBillNo.Text) & " )"
    Set CRXFormulaFields = Report.FormulaFields
    
    For i = 1 To Report.Database.Tables.Count
        Report.Database.Tables(i).SetLogOnInfo "ConnectionName", "M:\dbase\YEAR12-13\MDINV.MDB", "admin", "!@#$%^&*())(*&^%$#@!"
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & DataList2.Text & "'"
        If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & lbladdress.Caption & "'"
        If CRXFormulaField.Name = "{@DLNO2}" Then CRXFormulaField.Text = "'" & LBLDLNO2.Caption & "'"
        If CRXFormulaField.Name = "{@DLNO}" Then CRXFormulaField.Text = "'" & lbldlno.Caption & "'"
        If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
        If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
        If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
        If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
        If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
        If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'TAX INVOICE FORM 8H/8B/8'"
        If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & lbltin.Caption & "'"
        If lblcredit.Caption = "0" Then
            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'CASH'"
        Else
            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'" & txtcrdays.Text & "'" & "' Days'"
        End If
'        If lbltin.Caption = "" Then
'            If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "''"
'            If CRXFormulaField.Name = "{@ZVAT}" Then CRXFormulaField.Text = "'THE KERALA VALUE ADDED TAX RULES - 2005'"
'            If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'FORM 8B'"
'            If CRXFormulaField.Name = "{@ZRULE}" Then CRXFormulaField.Text = "'[See rule 58(10)]'"
'            If CRXFormulaField.Name = "{@ZRETAIL}" Then CRXFormulaField.Text = "'RETAIL INVOICE'"
'        Else
'
'            If CRXFormulaField.Name = "{@ZVAT}" Then CRXFormulaField.Text = "''"
'            If CRXFormulaField.Name = "{@ZRULE}" Then CRXFormulaField.Text = "''"
'            If CRXFormulaField.Name = "{@ZRETAIL}" Then CRXFormulaField.Text = "''"
'        End If
    Next
    frmreport.Caption = "BILL"
    Call GENERATEREPORT
    
    ''cmdRefresh.SetFocus

'    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption), 2)) - Val(Round(Val(LBLTOTAL.Caption), 0)), 2)
'
    
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTTAX.Enabled = False
    TXTEXPIRY.Enabled = False
    TxtFree.Enabled = False
    TXTPTR.Enabled = False
    TXTRETAIL.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    
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
            TXTRATE.Text = ""
            TXTRETAIL.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTPTR.Text = ""
            TxtFree.Text = ""
            OPTNET.Value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
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
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            TxtFree.Enabled = False
            TXTPTR.Enabled = False
            TXTRETAIL.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
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
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
        DataList2.SetFocus
        Exit Sub
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
            TXTRETAIL.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTPTR.Text = ""
            TxtFree.Text = ""
            OPTNET.Value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
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
            TXTRATE.Enabled = False
            TXTRETAIL.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            TxtFree.Enabled = False
            TXTPTR.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub Form_Activate()
    If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If TxtMRP.Enabled = True Then TxtMRP.SetFocus
    If TXTRATE.Enabled = True Then TXTRATE.SetFocus
    If TXTPTR.Enabled = True Then TXTPTR.SetFocus
    If TXTRETAIL.Enabled = True Then TXTRETAIL.SetFocus
    If TXTTAX.Enabled = True Then TXTTAX.SetFocus
    If TXTEXPIRY.Enabled = True Then TXTEXPIRY.SetFocus
    'If TXTEXPEnabled = True Then TXTEXPDATE.SetFocus
    If txtBatch.Enabled = True Then txtBatch.SetFocus
    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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
    lblcredit.Caption = "1"
    txtcrdays.Text = ""
    lblP_Rate.Caption = "0"
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
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
    grdsales.ColWidth(11) = 1100
    grdsales.ColWidth(12) = 1000
    grdsales.ColWidth(21) = 600
    grdsales.ColWidth(22) = 0
    
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
    grdsales.TextArray(11) = "EXPIRY"
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
    TXTRATE.Enabled = False
    TXTRETAIL.Enabled = False
    TXTTAX.Enabled = False
    TXTEXPIRY.Enabled = False
    TxtFree.Enabled = False
    TXTPTR.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    TXTSLNO.Text = 1
    
    TXTSLNO.Enabled = False
    CLOSEALL = 1
    M_EDIT = False
'    Me.Width = 11700
'    Me.Height = 10185
    Me.Left = 500
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
            'FRMSALE.TXTQTY.Text = GRDPOPUP.Columns(1)
            FRMSALE.TxtMRP.Text = GRDPOPUP.Columns(3)
            FRMSALE.TXTPTR.Text = IIf(IsNull(GRDPOPUP.Columns(4)), "", GRDPOPUP.Columns(4))
            FRMSALE.TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(13)), "", GRDPOPUP.Columns(13))
            'FRMSALE.TXTRATE.Text = GRDPOPUP.Columns(4)
            
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & GRDPOPUP.Columns(6) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            With RSTtax
                If Not (.EOF And .BOF) Then
                    Select Case GRDPOPUP.Columns(12)
                        Case "M"
                            OPTTaxMRP.Value = True
                            FRMSALE.TXTTAX.Text = GRDPOPUP.Columns(5)
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
                            FRMSALE.TXTTAX.Text = GRDPOPUP.Columns(5)
                        Case Else
                            TXTSALETYPE.Text = "2"
                            OPTNET.Value = True
                            FRMSALE.TXTTAX.Text = "0"
                    End Select
                Else
                    OPTNET.Value = True
                    FRMSALE.TXTTAX.Text = "0"
                End If
            End With
            RSTtax.Close
            Set RSTtax = Nothing
            
            FRMSALE.TXTEXPIRY.Text = IIf(GRDPOPUP.Columns(2) = "", "  /  ", Format(GRDPOPUP.Columns(2), "mm/yy"))
            FRMSALE.txtBatch.Text = GRDPOPUP.Columns(0)
            
            FRMSALE.TXTVCHNO.Text = GRDPOPUP.Columns(8)
            FRMSALE.TXTLINENO.Text = GRDPOPUP.Columns(9)
            FRMSALE.TXTTRXTYPE.Text = GRDPOPUP.Columns(10)
            FRMSALE.TXTUNIT.Text = GRDPOPUP.Columns(11)
            
'''''            Set RSTP_RATE = New ADODB.Recordset
'''''            RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & GRDPOPUP.Columns(6) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
'''''            If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
'''''                FRMSALE.TXTPTR.Text = IIf(IsNull(RSTP_RATE!PTR), "", RSTP_RATE!PTR)
'''''                'FRMSALE.TXTPTR.Text = RSTP_RATE!PTR
'''''                lblP_Rate.Caption = "1"
'''''            Else
'''''                FRMSALE.TXTPTR.Text = ""
'''''                lblP_Rate.Caption = "0"
'''''            End If
'''''            RSTP_RATE.Close
'''''            Set RSTP_RATE = Nothing
            
            Set GRDPOPUP.DataSource = Nothing
            
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            FRMSALE.TXTPRODUCT.Enabled = False
            FRMSALE.TXTQTY.Enabled = True
            FRMSALE.TXTQTY.SetFocus
        Case vbKeyEscape
            FRMSALE.TXTQTY.Text = ""
            FRMSALE.TXTVCHNO.Text = ""
            FRMSALE.TXTLINENO.Text = ""
            FRMSALE.TXTTRXTYPE.Text = ""
            FRMSALE.TXTUNIT.Text = ""
            
            Set GRDPOPUP.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            FRMSALE.TXTPRODUCT.Enabled = True
            FRMSALE.TXTQTY.Enabled = False
            FRMSALE.TXTPRODUCT.SetFocus
        
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
            FRMSALE.TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            FRMSALE.TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            i = 0
            If M_STOCK <= 0 Then
''''''                Set RSTNONSTOCK = New ADODB.Recordset
''''''                RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & FRMSALE.TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
''''''                i = RSTNONSTOCK.RecordCount
''''''                RSTNONSTOCK.Close
''''''                Set RSTNONSTOCK = Nothing
''''''                If i = 0 Then
''''''                    If (MsgBox("NO STOCK AVAILABLE..Do you want to add to Stockless", vbYesNo, "SALES") = vbYes) Then
''''''                        Set RSTNONSTOCK = New ADODB.Recordset
''''''                        RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & FRMSALE.TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
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
            For i = 1 To FRMSALE.grdsales.Rows - 1
                If Trim(FRMSALE.grdsales.TextMatrix(i, 13)) = Trim(FRMSALE.TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists.... Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then
                        Set GRDPOPUPITEM.DataSource = Nothing
                        FRMEITEM.Visible = False
                        FRMEMAIN.Enabled = True
                        FRMSALE.TXTPRODUCT.Enabled = True
                        FRMSALE.TXTQTY.Enabled = False
                        FRMSALE.TXTPRODUCT.SetFocus
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next i
            Set GRDPOPUPITEM.DataSource = Nothing
            If ITEM_FLAG = True Then
                If NONSTOCKFLAG = True Then
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                Else
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                End If
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                If NONSTOCKFLAG = True Then
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG, P_RETAIL From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                Else
                    PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP, CHECK_FLAG, P_RETAIL From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                End If
                ITEM_FLAG = False
            End If
            Set GRDPOPUPITEM.DataSource = PHY_ITEM
            If PHY_ITEM.RecordCount = 0 Then
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                FRMSALE.TXTPRODUCT.Enabled = False
                FRMSALE.TXTQTY.Enabled = True
                FRMSALE.TXTQTY.SetFocus
                Exit Sub
            End If
            If PHY_ITEM.RecordCount = 1 Or MINUSFLAG = True Then
                'FRMSALE.TXTQTY.Text = GRDPOPUPITEM.Columns(2)
                FRMSALE.TXTPTR.Text = IIf(IsNull(GRDPOPUPITEM.Columns(3)), "", GRDPOPUPITEM.Columns(3))
                FRMSALE.TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", GRDPOPUPITEM.Columns(13))
                'FRMSALE.TXTRATE.Text = GRDPOPUPITEM.Columns(3)
                FRMSALE.TxtMRP.Text = GRDPOPUPITEM.Columns(11)
                
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                With RSTtax
                    If Not (.EOF And .BOF) Then
                        Select Case PHY_ITEM!CHECK_FLAG
                            Case "M"
                                OPTTaxMRP.Value = True
                                FRMSALE.TXTTAX.Text = GRDPOPUPITEM.Columns(4)
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
                                FRMSALE.TXTTAX.Text = GRDPOPUPITEM.Columns(4)
                            Case Else
                                TXTSALETYPE.Text = "2"
                                OPTNET.Value = True
                                FRMSALE.TXTTAX.Text = "0"
                        End Select
                    Else
                        OPTNET.Value = True
                        FRMSALE.TXTTAX.Text = "0"
                    End If
                End With
                RSTtax.Close
                Set RSTtax = Nothing
            
                'FRMSALE.TXTTAX.Text = 0 'GRDPOPUPITEM.Columns(4)
                FRMSALE.TXTEXPIRY.Text = IIf(GRDPOPUPITEM.Columns(7) = "", "  /  ", Format(GRDPOPUPITEM.Columns(7), "MM/YY"))
                FRMSALE.txtBatch.Text = GRDPOPUPITEM.Columns(6)
                
                FRMSALE.TXTVCHNO.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(8), "")
                FRMSALE.TXTLINENO.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(9), "")
                FRMSALE.TXTTRXTYPE.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(10), "")
                FRMSALE.TXTUNIT.Text = GRDPOPUPITEM.Columns(5)
                
''''''                Set RSTP_RATE = New ADODB.Recordset
''''''                RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & GRDPOPUPITEM.Columns(0) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
''''''                If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
''''''                    FRMSALE.TXTPTR.Text = IIf(IsNull(RSTP_RATE!PTR), "", RSTP_RATE!PTR)
''''''                    ''FRMSALE.TXTPTR.Text = IIF(ISNULL(RSTP_RATE!PTR),"",
''''''                    lblP_Rate.Caption = "1"
''''''                Else
''''''                    FRMSALE.TXTPTR.Text = ""
''''''                    lblP_Rate.Caption = "0"
''''''                End If
''''''                RSTP_RATE.Close
''''''                Set RSTP_RATE = Nothing
            
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                FRMSALE.TXTPRODUCT.Enabled = False
                FRMSALE.TXTQTY.Enabled = True
                FRMSALE.TXTQTY.SetFocus
                Exit Sub
            ElseIf PHY_ITEM.RecordCount > 1 And MINUSFLAG = False Then
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEGRDTMP.Visible = False
                Call FILL_BATCHGRID
            End If
        Case vbKeyEscape
            FRMSALE.TXTQTY.Text = ""
            FRMSALE.TXTVCHNO.Text = ""
            FRMSALE.TXTLINENO.Text = ""
            FRMSALE.TXTTRXTYPE.Text = ""
            FRMSALE.TXTUNIT.Text = ""
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            FRMSALE.TXTPRODUCT.Enabled = True
            FRMSALE.TXTQTY.Enabled = False
            FRMSALE.TXTPRODUCT.SetFocus
            
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub optnet_Click()
    TXTPTR_LostFocus
End Sub

Private Sub OPTTaxMRP_Click()
    TXTPTR_LostFocus
End Sub

Private Sub OPTVAT_Click()
    TXTPTR_LostFocus
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
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
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
                    LBLTIME.Caption = Time
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
                        grdsales.TextMatrix(i, 18) = Trim(TRXMAST!MANUFACTURER)
                    End If
                    TRXMAST.Close
                    Set TRXMAST = Nothing
                    
                    grdsales.TextMatrix(i, 5) = Format(TRXFILE!MRP, ".000")
                    grdsales.TextMatrix(i, 6) = Format(TRXFILE!PTR, ".000")
                    grdsales.TextMatrix(i, 7) = Format(TRXFILE!SALES_PRICE, ".000")
                    grdsales.TextMatrix(i, 8) = 0 'DISC
                    grdsales.TextMatrix(i, 9) = Val(TRXFILE!SALES_TAX)
            
                    grdsales.TextMatrix(i, 10) = TRXFILE!REF_NO
                    grdsales.TextMatrix(i, 11) = Format(TRXFILE!EXP_DATE, "MM/YY")
                    grdsales.TextMatrix(i, 12) = Format(Val(TRXFILE!TRX_TOTAL), ".000")
                    
                    grdsales.TextMatrix(i, 13) = TRXFILE!ITEM_CODE
                    grdsales.TextMatrix(i, 14) = Val(TRXSUB!R_VCH_NO)
                    grdsales.TextMatrix(i, 15) = Val(TRXSUB!R_LINE_NO)
                    grdsales.TextMatrix(i, 16) = Trim(TRXSUB!R_TRX_TYPE)
                    grdsales.TextMatrix(i, 17) = IIf(IsNull(TRXFILE!CHECK_FLAG), "", Trim(TRXFILE!CHECK_FLAG))
                    TXTDEALER.Text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                    'DataList2.Text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                    TXTINVDATE.Text = IIf(IsNull(TRXFILE!VCH_DATE), Date, TRXFILE!VCH_DATE)
                    Select Case TRXFILE!CST
                        Case 0
                            grdsales.TextMatrix(i, 19) = "B"
                        Case 1
                            grdsales.TextMatrix(i, 19) = "DN"
                        Case 2
                            grdsales.TextMatrix(i, 19) = "CN"
                    End Select
                    grdsales.TextMatrix(i, 20) = TRXFILE!FREE_QTY
                    grdsales.TextMatrix(i, 21) = IIf(IsNull(TRXFILE!P_RETAIL), "0.00", Format(TRXFILE!P_RETAIL, ".000"))
                    grdsales.TextMatrix(i, 22) = IIf(IsNull(TRXFILE!P_RETAILWOTAX), "0.00", Format(TRXFILE!P_RETAILWOTAX, ".000"))
                    grdsales.TextMatrix(i, 23) = IIf(IsNull(TRXFILE!SALE_1_FLAG), "2", TRXFILE!SALE_1_FLAG)
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
                If TRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                LBLTIME.Caption = IIf(IsNull(TRXMAST!CFORM_NO), Time, TRXMAST!CFORM_NO)
                txtcrdays.Text = IIf(IsNull(TRXMAST!cr_days), "", TRXMAST!cr_days)
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            LBLBILLNO.Caption = Val(txtBillNo.Text)
            
            LBLTOTAL.Caption = ""
            lblnetamount.Caption = ""
            LBLFOT.Caption = ""
            For i = 1 To grdsales.Rows - 1
                grdsales.TextMatrix(i, 0) = i
                Select Case grdsales.TextMatrix(i, 19)
                    Case "CN"
                        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                    Case Else
                        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                End Select
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
                TXTDEALER.Text = ""
                TXTDEALER.SetFocus
            End If
    
    End Select
    
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click

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

Private Sub txtcrdays_GotFocus()
    txtcrdays.SelStart = 0
    txtcrdays.SelLength = Len(txtcrdays.Text)
End Sub

Private Sub txtcrdays_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TxtFree.Enabled = True Then TxtFree.SetFocus
            If TXTPTR.Enabled = True Then TXTPTR.SetFocus
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

Private Sub txtcrdays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo eRRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='131')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='131')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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

Private Sub TXTDEALER_LostFocus()
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y' AND TRX_TYPE = 'SI'", db2, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        CMDDELIVERY.Enabled = True
    Else
        CMDDELIVERY.Enabled = False
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM SALERETURN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y'", db2, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        CMDSALERETURN.Enabled = True
    Else
        CMDSALERETURN.Enabled = False
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
End Sub

Private Sub TXTDISC_GotFocus()
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.Text)
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRETAIL.Enabled = True
            TXTDISC.Enabled = False
            TXTRETAIL.SetFocus
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
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag), ".000")

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
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                DataList2.SetFocus
            End If
        Case vbKeyEscape
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

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
    CMDDELIVERY.Enabled = False
    CMDSALERETURN.Enabled = False
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            lbladdress.Caption = ""
            lbldlno.Caption = ""
            LBLDLNO2.Caption = ""
            lbltin.Caption = ""
            DataList2.SetFocus
        Case vbKeyEscape
            txtBillNo.Visible = True
            txtBillNo.SetFocus
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
            TxtFree.Enabled = True
            TxtFree.SetFocus
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
            TXTRATE.Text = ""
            TXTRETAIL.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTPTR.Text = ""
            TxtFree.Text = ""
            OPTNET.Value = True
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
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG, P_RETAIL From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                Else
                    TMPREC.Close
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG, P_RETAIL From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                End If
                
                Set grdtmp.DataSource = TMPREC
                If TMPREC.RecordCount = 1 Then
                    'FRMSALE.TXTQTY.Text = grdtmp.Columns(2)
                    FRMSALE.TxtMRP.Text = grdtmp.Columns(3)
                    FRMSALE.TXTPTR.Text = IIf(IsNull(grdtmp.Columns(4)), "", grdtmp.Columns(4))
                    FRMSALE.TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(13)), "", grdtmp.Columns(13))
                    'FRMSALE.TXTRATE.Text = grdtmp.Columns(4)
                    
                    Set RSTtax = New ADODB.Recordset
                    RSTtax.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdtmp.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                    With RSTtax
                        If Not (.EOF And .BOF) Then
                            Select Case TMPREC!CHECK_FLAG
                                Case "M"
                                    OPTTaxMRP.Value = True
                                    FRMSALE.TXTTAX.Text = grdtmp.Columns(5)
                                    TXTSALETYPE.Text = "2"
                                Case "V"
                                    If (!CATEGORY = "MEDICINE" And !Remarks = "1") Then
                                        TXTSALETYPE.Text = "1"
                                    'If !CATEGORY = "MEDICINE" Then
                                        OPTTaxMRP.Value = True
                                    Else
                                        TXTSALETYPE.Text = "2"
                                        OPTVAT.Value = True
                                    End If
                                    FRMSALE.TXTTAX.Text = grdtmp.Columns(5)
                                Case Else
                                    TXTSALETYPE.Text = "2"
                                    OPTNET.Value = True
                                    FRMSALE.TXTTAX.Text = "0"
                            End Select
                        Else
                            OPTNET.Value = True
                            FRMSALE.TXTTAX.Text = "0"
                        End If
                    End With
                    RSTtax.Close
                    Set RSTtax = Nothing
                    
                    FRMSALE.TXTEXPIRY.Text = IIf(grdtmp.Columns(8) = "", "  /  ", Format(grdtmp.Columns(8), "MM/YY"))
                    FRMSALE.txtBatch.Text = grdtmp.Columns(7)
                    
                    FRMSALE.TXTVCHNO.Text = grdtmp.Columns(9)
                    FRMSALE.TXTLINENO.Text = grdtmp.Columns(10)
                    FRMSALE.TXTTRXTYPE.Text = grdtmp.Columns(11)
                    FRMSALE.TXTUNIT.Text = grdtmp.Columns(6)
                    
''''''                    Set RSTP_RATE = New ADODB.Recordset
''''''                    RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & grdtmp.Columns(0) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
''''''                    If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
''''''                        FRMSALE.TXTPTR.Text = IIf(IsNull(RSTP_RATE!PTR), "", RSTP_RATE!PTR)
''''''                        'FRMSALE.TXTRATE.Text = RSTP_RATE!PTR
''''''                        lblP_Rate.Caption = "1"
''''''                    Else
''''''                        FRMSALE.TXTPTR.Text = ""
''''''                        lblP_Rate.Caption = "0"
''''''                    End If
''''''                    RSTP_RATE.Close
''''''                    Set RSTP_RATE = Nothing
                    
                    FRMSALE.TXTPRODUCT.Enabled = False
                    FRMSALE.TXTQTY.Enabled = True
                    FRMSALE.TXTQTY.SetFocus
                    Exit Sub
                ElseIf TMPREC.RecordCount = 0 Then
                    Set RSTBALQTY = New ADODB.Recordset
                    RSTBALQTY.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                    With RSTBALQTY
                        If Not (.EOF And .BOF) Then
                            M_STOCK = !CLOSE_QTY
                        End If
                    End With
                    RSTBALQTY.Close
                    Set RSTBALQTY = Nothing
            
                    FRMSALE.TXTQTY.Text = 0
                    i = 0
''''''                    Set RSTNONSTOCK = New ADODB.Recordset
''''''                    RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & FRMSALE.TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
''''''                    i = RSTNONSTOCK.RecordCount
''''''                    RSTNONSTOCK.Close
''''''                    Set RSTNONSTOCK = New ADODB.Recordset
''''''                    If i = 0 Then
''''''                        If (MsgBox("NO STOCK AVAILABLE..Do you want to add to Stockless", vbYesNo, "SALES") = vbYes) Then
''''''                            Set RSTNONSTOCK = New ADODB.Recordset
''''''                            RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & FRMSALE.TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
''''''                            If (RSTNONSTOCK.EOF And RSTNONSTOCK.BOF) Then
''''''                                RSTNONSTOCK.AddNew
''''''                                RSTNONSTOCK!ITEM_NAME = TXTPRODUCT.Text
''''''                                RSTNONSTOCK!ITEM_CODE = TXTITEMCODE.Text
''''''                                RSTNONSTOCK!Date = Date & " " & Time
''''''                                RSTNONSTOCK.Update
''''''                            End If
''''''                            RSTNONSTOCK.Close
''''''                            Set RSTNONSTOCK = Nothing
''''''                        End If
''''''                    End If
                    
''''''                    If (MsgBox("AVAILABLE STOCK IS  " & M_STOCK & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbYes) Then
''''''                        Set RSTZEROSTOCK = New ADODB.Recordset
''''''                        RSTZEROSTOCK.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' ORDER BY [VCH_NO]", db, adOpenStatic, adLockReadOnly
''''''                        If Not (RSTZEROSTOCK.EOF And RSTZEROSTOCK.BOF) Then
''''''                            FRMSALE.TxtMRP.Text = RSTZEROSTOCK!MRP
''''''                            Select Case RSTZEROSTOCK!CHECK_FLAG
''''''                                Case "M"
''''''                                    OPTTaxMRP.Value = True
''''''                                    OPTVAT.Value = False
''''''                                    FRMSALE.TXTTAX.Text = RSTZEROSTOCK!SALES_TAX
''''''                                Case "V"
''''''                                    OPTVAT.Value = True
''''''                                    OPTTaxMRP.Value = False
''''''                                    FRMSALE.TXTTAX.Text = RSTZEROSTOCK!SALES_TAX
''''''                                Case Else
''''''                                    OPTTaxMRP.Value = False
''''''                                    OPTVAT.Value = False
''''''                                    FRMSALE.TXTTAX.Text = "0"
''''''                            End Select
''''''                            FRMSALE.TXTEXPIRY.Text = IIf(IsNull(RSTZEROSTOCK!EXP_DATE), "  /  ", Format(RSTZEROSTOCK!EXP_DATE, "MM/YY"))
''''''                            FRMSALE.txtBatch.Text = RSTZEROSTOCK!REF_NO
''''''
''''''                            FRMSALE.TXTVCHNO.Text = ""
''''''                            FRMSALE.TXTLINENO.Text = ""
''''''                            FRMSALE.TXTTRXTYPE.Text = ""
''''''                            FRMSALE.TXTUNIT.Text = RSTZEROSTOCK!UNIT
''''''
''''''                            Set RSTP_RATE = New ADODB.Recordset
''''''                            RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & RSTZEROSTOCK!ITEM_CODE & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
''''''                            If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
''''''                                FRMSALE.TXTPTR.Text = IIf(IsNull(RSTP_RATE!PTR), "", RSTP_RATE!PTR)
''''''                                'FRMSALE.TXTRATE.Text = RSTP_RATE!PTR
''''''                                lblP_Rate.Caption = "1"
''''''                            Else
''''''                                lblP_Rate.Caption = "0"
''''''                            End If
''''''                            RSTP_RATE.Close
''''''                            Set RSTP_RATE = Nothing
''''''                        End If
''''''                        RSTZEROSTOCK.Close
''''''                        Set RSTZEROSTOCK = Nothing
''''''
''''''                        GoTo JUMPNONSTOCK
''''''                    End If
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
                TXTRATE.Enabled = False
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
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CmdDelete.Enabled = False
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
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
            If i > 0 Then
                If Val(TXTQTY.Text) > i Then
                    MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                    TXTQTY.SelStart = 0
                    TXTQTY.SelLength = Len(TXTQTY.Text)
                    Exit Sub
                End If
            End If
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
            TXTQTY.Enabled = False
            TxtFree.Enabled = True
            TxtFree.SetFocus
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
    If Val(TXTRATE.Text) = 0 Then
        TXTDISC.Tag = Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    End If
    If Val(TXTPTR.Text) = 0 Then TXTPTR.Text = Format(Val(LBLSELPRICE.Caption), "0.00")
End Sub

Private Sub TXTRATE_GotFocus()
    TXTRATE.SelStart = 0
    TXTRATE.SelLength = Len(TXTRATE.Text)
End Sub

Private Sub TXTRATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTRATE.Text) = 0 Then Exit Sub
            TXTRATE.Enabled = False
            TXTEXPIRY.Enabled = True
            TXTEXPIRY.SetFocus
        Case vbKeyEscape
            TXTRATE.Enabled = False
            TXTPTR.Enabled = True
            TXTPTR.SetFocus
    End Select
End Sub

Private Sub TXTRATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub TXTRATE_LostFocus()
    TXTRATE.Text = Format(TXTRATE.Text, ".000")
    TXTDISC.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag), ".000")
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
                TXTPTR.Text = ""
                TxtFree.Text = ""
                OPTNET.Value = True
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
                TXTEXPIRY.Text = "  /  "
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
                TxtFree.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 20)
                TxtMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTPTR.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                TXTRATE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), ".000")
                
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 15)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 16)
                TXTUNIT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 22)
                TXTSALETYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 23)
                LBLSUBTOTAL.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 12)
                If grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = "" Then
                    TXTEXPIRY.Text = "  /  "
                Else
                    TXTEXPIRY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 11)
                End If
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                
                Select Case Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
                    Case "M"
                        OPTTaxMRP.Value = True
                    Case "V"
                        OPTVAT.Value = True
                    Case Else
                        OPTNET.Value = True
                End Select
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                TXTRATE.Enabled = False
                TXTTAX.Enabled = False
                TXTEXPIRY.Enabled = False
                TxtFree.Enabled = False
                TXTPTR.Enabled = False
                TXTRETAIL.Enabled = False
                txtBatch.Enabled = False
                TXTDISC.Enabled = False
                TxtMRP.Enabled = False
                Select Case grdsales.TextMatrix(Val(TXTSLNO.Text), 19)
                    Case "CN", "DN"
                        CmdDelete.Enabled = True
                        CmdDelete.SetFocus
                        
                    Case Else
                        CMDMODIFY.Enabled = True
                        CMDMODIFY.SetFocus
                        CmdDelete.Enabled = True
                End Select
                
                LBLDNORCN.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 19)
                Exit Sub
            End If
SKIP:
            lblP_Rate.Caption = "0"
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                OPTNET.Value = True
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTUNIT.Text = ""
                TXTQTY.Text = ""
                TXTRATE.Text = ""
                TXTRETAIL.Text = ""
                TXTRETAILNOTAX.Text = ""
                TXTSALETYPE.Text = ""
                TXTPTR.Text = ""
                TxtFree.Text = ""
                TxtMRP.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTEXPIRY.Text = "  /  "
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
                DataList2.Enabled = True
                DataList2.SetFocus
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
            TXTPTR.Enabled = True
            TXTTAX.Enabled = False
            TXTPTR.SetFocus
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
            TXTPTR.Enabled = True
            TXTPTR.SetFocus
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
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT, CHECK_FLAG, P_RETAIL From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT, CHECK_FLAG, P_RETAIL From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "BATCH NO."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "EXP DATE"
    GRDPOPUP.Columns(3).Caption = "MRP"
    GRDPOPUP.Columns(4).Caption = "RATE"
    GRDPOPUP.Columns(5).Caption = "TAX"
    GRDPOPUP.Columns(8).Caption = "VCH No"
    GRDPOPUP.Columns(9).Caption = "Line No"
    GRDPOPUP.Columns(10).Caption = "Trx Type"
    
    GRDPOPUP.Columns(0).Width = 1400
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 1400
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
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & FRMSALE.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & FRMSALE.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
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
    On Error GoTo eRRHAND
    
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
        RSTTRXFILE!act_code = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.Text
        RSTTRXFILE.Update
    Else
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "SI"
        RSTTRXFILE!INV_NO = Val(txtBillNo.Text)
        RSTTRXFILE!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!INV_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!RCPT_AMOUNT = 0
        RSTTRXFILE!BAL_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!act_code = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.Text
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
        RSTTRXFILE!act_code = "501001"
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
        RSTTRXFILE!act_code = "111001"
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
    RSTTRXFILE!act_code = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = "Sales"
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
    If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
    RSTTRXFILE!CFORM_NO = LBLTIME.Caption
    RSTTRXFILE!Remarks = DataList2.Text
    RSTTRXFILE!DISC_PERS = 0
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!cr_days = Val(txtcrdays.Text)
    
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
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "Issued to     " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 10)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
        Select Case grdsales.TextMatrix(i, 19)
            Case "DN"
                RSTTRXFILE!CST = 1
            Case "CN"
                RSTTRXFILE!CST = 2
            Case Else
                RSTTRXFILE!CST = 0
        End Select
        
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        If grdsales.TextMatrix(i, 11) = "" Then
            RSTTRXFILE!EXP_DATE = Null
        Else
            RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(i, 11))) & "/" & Trim(grdsales.TextMatrix(i, 11))
        End If
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT AREA FROM ACTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!Area = RSTITEMMAST!Area
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
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

    i = 0
    Set rstMaxRec = New ADODB.Recordset
    rstMaxRec.Open "Select MAX(Val(CR_NO)) From CRDTPYMT", db2, adOpenStatic, adLockReadOnly
    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
        i = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
    End If
    rstMaxRec.Close
    Set rstMaxRec = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CRDTPYMT WHERE INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'DR'", db2, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!TRX_TYPE = "DR"
        RSTITEMMAST!CR_NO = i
        RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
        RSTITEMMAST!RCPT_AMT = 0
    End If
    RSTITEMMAST!INV_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
    RSTITEMMAST!INV_AMT = Val(LBLTOTAL.Caption)
    RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption) - RSTITEMMAST!RCPT_AMT
    If lblcredit.Caption = "0" Then RSTITEMMAST!CHECK_FLAG = "Y" Else RSTITEMMAST!CHECK_FLAG = "N"
    RSTITEMMAST!PINV = ""
    RSTITEMMAST!act_code = DataList2.BoundText
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(VCH_NO)) From TRXMAST WHERE TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    lbladdress.Caption = ""
    LBLDNORCN.Caption = ""
    lbldlno.Caption = ""
    LBLDLNO2.Caption = ""
    lbltin.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    LBLPROFIT.Caption = ""
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
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
    TXTDEALER.Enabled = True
    TXTDEALER.SetFocus
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Tag = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblcredit.Caption = "0"
    txtcrdays.Text = ""
    cr_days = False
    Exit Function
eRRHAND:
    MsgBox Err.Description
End Function

Private Sub TxtFree_GotFocus()
    TxtFree.SelStart = 0
    TxtFree.SelLength = Len(TxtFree.Text)
    TxtFree.Tag = Trim(TXTPRODUCT.Text)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TxtFree.Text) = 0 Then GoTo SKIP
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
                If Val(TxtFree.Text) + Val(TXTQTY.Text) > i Then
                    MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                    TxtFree.SelStart = 0
                    TxtFree.SelLength = Len(TxtFree.Text)
                    Exit Sub
                End If
            End If
SKIP:
            TxtFree.Enabled = False
            TxtMRP.Enabled = True
            TxtMRP.SetFocus
         Case vbKeyEscape
            TxtFree.Enabled = False
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

Private Sub TXTPTR_GotFocus()
    TXTPTR.SelStart = 0
    TXTPTR.SelLength = Len(TXTPTR.Text)
End Sub

Private Sub TXTPTR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            TXTPTR.Enabled = False
            TXTEXPIRY.Enabled = True
            TXTEXPIRY.SetFocus
        Case vbKeyEscape
            TXTPTR.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
    End Select
End Sub

Private Sub TXTPTR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPTR_LostFocus()
    TXTPTR.Text = Format(TXTPTR.Text, ".000")
    ''If lblP_Rate.Caption = "0" Then
        If OPTTaxMRP.Value = True Then
            TXTRETAIL.Text = Val(TXTRETAILNOTAX.Text) + Val(txtmrpbt.Text) * Val(TXTTAX.Text) / 100
            TXTRATE.Text = Val(TXTPTR.Text) + Val(txtmrpbt.Text) * Val(TXTTAX.Text) / 100
        End If
        If OPTVAT.Value = True Then
            TXTRETAIL.Text = Val(TXTRETAILNOTAX.Text) + Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100
            TXTRATE.Text = Val(TXTPTR.Text) + Val(TXTPTR.Text) * Val(TXTTAX.Text) / 100
        End If
        If OPTNET.Value = True Then
            TXTRATE.Text = TXTPTR.Text
            TXTRETAIL.Text = TXTRETAILNOTAX.Text
        End If
    '''End If
End Sub

Private Sub OPTDISCPERCENT_Click()
    TXTTOTALDISC.SetFocus
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
            If TxtFree.Enabled = True Then TxtFree.SetFocus
            If TXTPTR.Enabled = True Then TXTPTR.SetFocus
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

Private Sub TXTTOTALDISC_LostFocus()
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                lblnetamount.Caption = Val(lblnetamount.Caption) - Val(grdsales.TextMatrix(i, 12))
            Case Else
                lblnetamount.Caption = Val(lblnetamount.Caption) + Val(grdsales.TextMatrix(i, 12))
        End Select
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

Private Sub ReportGeneratION()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    Dim rstSUBfile As ADODB.Recordset
    Dim SN As Integer
    Dim TRXTOTAL As Double
    
    SN = 0
    TRXTOTAL = 0
   ' On Error GoTo errHand
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

    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenForwardOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!ADDRESS & ", " & RSTCOMPANY!HO_NAME, 140)
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(48) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(48) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(48) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "SALES SUMMARY FOR THE PERIOD"
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    'Set RSTTRXFILE = New ADODB.Recordset
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(13) & RepeatString("-", 59)
    Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SN", 3) & Space(2) & _
            AlignLeft("INV DATE", 8) & Space(10) & _
            AlignLeft("INV AMT", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Print #1, Space(12) & RepeatString("-", 59)
    SN = 0
'    RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_NO", Conn, adOpenForwardOnly
'    Do Until RSTTRXFILE.EOF
'        SN = SN + 1
'        Print #1, Chr(27) & Chr(71) & Space(5) & Chr(14) & Chr(15) & AlignRight(Str(SN), 4) & ". " & Space(1) & _
'            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
'            AlignRight(Format(Round(RSTTRXFILE!VCH_AMOUNT, 0), "0.00"), 16)
'        'Print #1, Chr(13)
'        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
'        RSTTRXFILE.MoveNext
'    Loop
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    Print #1,
    
    Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(13) & AlignRight("TOTAL AMOUNT", 12) & AlignRight((Format(TRXTOTAL, "####.00")), 11)
    Print #1, Space(56) & RepeatString("-", 16)
    'Print #1, Chr(27) & Chr(67) & Chr(0)
    'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** WISH YOU A SPEEDY RECOVERY ****", 40)


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
    
    'Print #1, Chr(27) & Chr(80)
    Close #1 '//Closing the file
    'MsgBox "Report file generated at " & App.Path & "\Report.txt" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub TXTRETAIL_GotFocus()
    TXTRETAIL.SelStart = 0
    TXTRETAIL.SelLength = Len(TXTRETAIL.Text)
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.Enabled = True
            TXTRETAIL.Enabled = False
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTRETAIL.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
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

Private Sub TXTRETAIL_LostFocus()
    TXTRETAIL.Text = Format(TXTRETAIL.Text, ".000")
End Sub



