VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDaybook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACCOUNTS SUMMARY"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19095
   ControlBox      =   0   'False
   Icon            =   "Daybook.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   19095
   Begin VB.Frame FRAME 
      Height          =   8985
      Left            =   30
      TabIndex        =   2
      Top             =   -60
      Width           =   19065
      Begin VB.CommandButton CmdDenom 
         BackColor       =   &H00400000&
         Caption         =   "Denom"
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
         Left            =   6495
         MaskColor       =   &H80000007&
         TabIndex        =   103
         Top             =   2730
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cash Flow"
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
         Height          =   600
         Left            =   45
         TabIndex        =   92
         Top             =   3060
         Width           =   8115
         Begin VB.Label LBLEXPENSES 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Paid"
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
            Height          =   255
            Index           =   5
            Left            =   1920
            TabIndex        =   100
            Top             =   240
            Width           =   645
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
            Height          =   345
            Left            =   2475
            TabIndex        =   99
            Top             =   195
            Width           =   1545
         End
         Begin VB.Label LBLEXPENSES 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Clo."
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
            Height          =   255
            Index           =   4
            Left            =   6120
            TabIndex        =   98
            Top             =   240
            Width           =   345
         End
         Begin VB.Label LBLEXPENSES 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rcvd"
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
            Height          =   255
            Index           =   3
            Left            =   4020
            TabIndex        =   97
            Top             =   240
            Width           =   555
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
            Height          =   345
            Left            =   4560
            TabIndex        =   96
            Top             =   195
            Width           =   1545
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
            Height          =   345
            Left            =   510
            TabIndex        =   95
            Top             =   195
            Width           =   1440
         End
         Begin VB.Label LBLEXPENSES 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OP."
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
            Height          =   255
            Index           =   2
            Left            =   105
            TabIndex        =   94
            Top             =   240
            Width           =   285
         End
         Begin VB.Label LBLWCLOCASH 
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
            Height          =   345
            Left            =   6465
            TabIndex        =   93
            Top             =   195
            Width           =   1605
         End
      End
      Begin VB.CommandButton CmdoffExp 
         Caption         =   "Couter wise Office Expense"
         Height          =   495
         Left            =   4245
         TabIndex        =   91
         Top             =   8175
         Width           =   1335
      End
      Begin VB.CommandButton CmdCunterSales 
         Caption         =   "Couter wise Sales Report"
         Height          =   495
         Left            =   2985
         TabIndex        =   90
         Top             =   8175
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton Cmdtax 
         BackColor       =   &H00400000&
         Caption         =   "&Tax Details"
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
         Left            =   10995
         MaskColor       =   &H80000007&
         TabIndex        =   87
         Top             =   8175
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00400000&
         Caption         =   "Print Item Wise Report"
         Height          =   495
         Left            =   1620
         MaskColor       =   &H80000007&
         TabIndex        =   75
         Top             =   8175
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdPrnCoolie 
         BackColor       =   &H00400000&
         Caption         =   "Print Coolie"
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
         Left            =   9810
         MaskColor       =   &H80000007&
         TabIndex        =   74
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   1170
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tax Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2790
         Left            =   9390
         TabIndex        =   65
         Top             =   105
         Width           =   1635
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Profit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Index           =   8
            Left            =   45
            TabIndex        =   89
            Top             =   2190
            Width           =   690
         End
         Begin VB.Label LBLPROFIT 
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
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   30
            TabIndex        =   88
            Top             =   2415
            Width           =   1560
         End
         Begin VB.Label LBLdifftax 
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
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   30
            TabIndex        =   73
            Top             =   1380
            Width           =   1560
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Diff Tax in Sales"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Index           =   15
            Left            =   45
            TabIndex        =   72
            Top             =   1155
            Width           =   1695
         End
         Begin VB.Label lblpurtax 
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
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   30
            TabIndex        =   71
            Top             =   855
            Width           =   1560
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Tax"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Index           =   14
            Left            =   45
            TabIndex        =   70
            Top             =   630
            Width           =   1365
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   4
            Left            =   45
            TabIndex        =   69
            Top             =   135
            Width           =   1365
         End
         Begin VB.Label lbltaxamt 
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
            ForeColor       =   &H000000C0&
            Height          =   320
            Left            =   30
            TabIndex        =   68
            Top             =   345
            Width           =   1560
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Flood Cess"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Index           =   11
            Left            =   45
            TabIndex        =   67
            Top             =   1680
            Width           =   1305
         End
         Begin VB.Label lblcess 
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
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   30
            TabIndex        =   66
            Top             =   1920
            Width           =   1560
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00400000&
         Caption         =   "Print Cash / Credit Report"
         Height          =   495
         Left            =   75
         MaskColor       =   &H80000007&
         TabIndex        =   48
         Top             =   8175
         UseMaskColor    =   -1  'True
         Width           =   1470
      End
      Begin VB.CommandButton CmdDayBook 
         BackColor       =   &H00400000&
         Caption         =   "Print Day &Book"
         Height          =   495
         Left            =   5625
         MaskColor       =   &H80000007&
         TabIndex        =   47
         Top             =   8175
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00400000&
         Caption         =   "&Print Cash Book"
         Height          =   495
         Left            =   7005
         MaskColor       =   &H80000007&
         TabIndex        =   33
         Top             =   8175
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
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
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.CommandButton CmdDisplay 
         BackColor       =   &H00400000&
         Caption         =   "&DISPLAY"
         Height          =   495
         Left            =   8385
         MaskColor       =   &H80000007&
         TabIndex        =   0
         Top             =   8175
         UseMaskColor    =   -1  'True
         Width           =   1215
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
         Height          =   495
         Left            =   9645
         MaskColor       =   &H80000007&
         TabIndex        =   1
         Top             =   8175
         UseMaskColor    =   -1  'True
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   1680
         TabIndex        =   4
         Top             =   120
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
         Format          =   144113665
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   3570
         TabIndex        =   5
         Top             =   120
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
         Format          =   144113665
         CurrentDate     =   40498
      End
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   2130
         Left            =   60
         TabIndex        =   23
         Top             =   3660
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   3757
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
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdrcpt 
         Height          =   1935
         Left            =   45
         TabIndex        =   24
         Top             =   6225
         Width           =   10950
         _ExtentX        =   19315
         _ExtentY        =   3413
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
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   3465
         Left            =   11010
         TabIndex        =   41
         Top             =   4695
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   6112
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdlend 
         Height          =   1950
         Left            =   11040
         TabIndex        =   49
         Top             =   375
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   3440
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grddeposit 
         Height          =   1770
         Left            =   11025
         TabIndex        =   55
         Top             =   2550
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   3122
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblPurchret 
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
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   1230
         TabIndex        =   102
         Top             =   2715
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purch Ret."
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
         Left            =   45
         TabIndex        =   101
         Top             =   2685
         Width           =   1020
      End
      Begin VB.Label LBLWSALE 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1230
         TabIndex        =   86
         Top             =   1380
         Width           =   1665
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Charge Paid"
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
         Height          =   270
         Index           =   19
         Left            =   5940
         TabIndex        =   85
         Top             =   2445
         Width           =   2115
      End
      Begin VB.Label lblServicepaid 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   7905
         TabIndex        =   84
         Top             =   2430
         Width           =   1470
      End
      Begin VB.Label LblPettyPur 
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1230
         TabIndex        =   83
         Top             =   705
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Purch"
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
         Left            =   60
         TabIndex        =   82
         Top             =   690
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Assets Purchase"
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
         Index           =   18
         Left            =   5940
         TabIndex        =   81
         Top             =   1785
         Width           =   1635
      End
      Begin VB.Label lblassets 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   7590
         TabIndex        =   80
         Top             =   1770
         Width           =   1785
      End
      Begin VB.Label lblexptax 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   7905
         TabIndex        =   79
         Top             =   2100
         Width           =   1470
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense(Tax Credit)"
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
         Height          =   270
         Index           =   17
         Left            =   5940
         TabIndex        =   78
         Top             =   2115
         Width           =   2115
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange"
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
         Index           =   16
         Left            =   2955
         TabIndex        =   77
         Top             =   2640
         Width           =   1410
      End
      Begin VB.Label lblxchange 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   4110
         TabIndex        =   76
         Top             =   2640
         Width           =   1740
      End
      Begin VB.Label LBLPETTY 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   4110
         TabIndex        =   64
         Top             =   750
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label LBLSERVICE 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   4110
         TabIndex        =   63
         Top             =   1695
         Width           =   1740
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Bills"
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
         Index           =   13
         Left            =   2955
         TabIndex        =   62
         Top             =   1710
         Width           =   1185
      End
      Begin VB.Line Line2 
         X1              =   5895
         X2              =   5895
         Y1              =   135
         Y2              =   2685
      End
      Begin VB.Label LBLB2C 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   4110
         TabIndex        =   61
         Top             =   1380
         Width           =   1740
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "B2C Sales"
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
         Index           =   12
         Left            =   2955
         TabIndex        =   60
         Top             =   1380
         Width           =   1185
      End
      Begin VB.Label LBLB2B 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   4110
         TabIndex        =   59
         Top             =   1065
         Width           =   1740
      End
      Begin VB.Label LBLB2BSALES 
         BackStyle       =   0  'Transparent
         Caption         =   "B2B Sales"
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
         Index           =   11
         Left            =   2955
         TabIndex        =   58
         Top             =   1065
         Width           =   1155
      End
      Begin VB.Label LBLPETTYSALE 
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Sales"
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
         Index           =   4
         Left            =   2955
         TabIndex        =   57
         Top             =   750
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label LBLEXPENSES 
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Details"
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
         Index           =   10
         Left            =   11040
         TabIndex        =   56
         Top             =   2310
         Width           =   2250
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Frieght /Handle"
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
         Height          =   270
         Index           =   2
         Left            =   8205
         TabIndex        =   54
         Top             =   2955
         Width           =   1560
      End
      Begin VB.Label lblhandle 
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   8160
         TabIndex        =   53
         Top             =   3195
         Width           =   1620
      End
      Begin VB.Label LBLEXPENSES 
         BackStyle       =   0  'Transparent
         Caption         =   "Money Lend Details"
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
         Index           =   9
         Left            =   11055
         TabIndex        =   52
         Top             =   105
         Width           =   2250
      End
      Begin VB.Label lblsaleret 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   4110
         TabIndex        =   51
         Top             =   2325
         Width           =   1740
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Return"
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
         Left            =   2955
         TabIndex        =   50
         Top             =   2340
         Width           =   1410
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
         Left            =   12210
         TabIndex        =   46
         Top             =   4380
         Width           =   1905
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
         Left            =   13725
         TabIndex        =   45
         Top             =   4350
         Width           =   1965
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
         Index           =   0
         Left            =   15270
         TabIndex        =   44
         Top             =   4395
         Width           =   1875
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
         Left            =   16725
         TabIndex        =   43
         Top             =   4350
         Width           =   1875
      End
      Begin VB.Label LBLEXPENSES 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Details"
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
         Left            =   11025
         TabIndex        =   42
         Top             =   4455
         Width           =   1410
      End
      Begin VB.Label LBLEXPENSES 
         BackStyle       =   0  'Transparent
         Caption         =   "Office Income"
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
         Height          =   285
         Index           =   7
         Left            =   5940
         TabIndex        =   40
         Top             =   1470
         Width           =   1515
      End
      Begin VB.Label LBLINCOME 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   7590
         TabIndex        =   39
         Top             =   1455
         Width           =   1785
      End
      Begin VB.Label LBLEXPENSES 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff Expenses"
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
         Height          =   435
         Index           =   6
         Left            =   5940
         TabIndex        =   38
         Top             =   810
         Width           =   1440
      End
      Begin VB.Label LblStaff 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   7590
         TabIndex        =   37
         Top             =   795
         Width           =   1785
      End
      Begin VB.Label LBLEXPENSES 
         BackStyle       =   0  'Transparent
         Caption         =   "Commission"
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
         TabIndex        =   36
         Top             =   135
         Width           =   1095
      End
      Begin VB.Label lblCommi 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   7590
         TabIndex        =   35
         Top             =   135
         Width           =   1785
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
         Height          =   345
         Left            =   7935
         TabIndex        =   34
         Top             =   3585
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label LBLEXPENSE 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   7590
         TabIndex        =   32
         Top             =   1125
         Width           =   1785
      End
      Begin VB.Label LBLEXPENSES 
         BackStyle       =   0  'Transparent
         Caption         =   "Office Expenses"
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
         Height          =   285
         Index           =   0
         Left            =   5940
         TabIndex        =   31
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Line Line1 
         X1              =   2910
         X2              =   2910
         Y1              =   645
         Y2              =   3090
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Sale"
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
         Index           =   6
         Left            =   60
         TabIndex        =   30
         Top             =   2355
         Width           =   1000
      End
      Begin VB.Label lblcashsale 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1230
         TabIndex        =   29
         Top             =   2370
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Sale"
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
         Height          =   285
         Index           =   5
         Left            =   2955
         TabIndex        =   28
         Top             =   2010
         Width           =   1515
      End
      Begin VB.Label lblcrdtsale 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   4110
         TabIndex        =   27
         Top             =   2010
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Received"
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
         Index           =   1
         Left            =   5265
         TabIndex        =   26
         Top             =   5865
         Width           =   2040
      End
      Begin VB.Label lblcashrcv 
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
         Height          =   390
         Left            =   7350
         TabIndex        =   25
         Top             =   5820
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
         Height          =   300
         Index           =   3
         Left            =   60
         TabIndex        =   22
         Top             =   1365
         Width           =   1000
      End
      Begin VB.Label LBLBTRXTOTAL 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   135
         TabIndex        =   21
         Top             =   3555
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label LBLCOST 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   7590
         TabIndex        =   20
         Top             =   465
         Width           =   1785
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
         Left            =   5940
         TabIndex        =   19
         Top             =   480
         Width           =   660
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
         Height          =   300
         Index           =   9
         Left            =   60
         TabIndex        =   18
         Top             =   1695
         Width           =   1000
      End
      Begin VB.Label LBLDISCOUNT 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1230
         TabIndex        =   17
         Top             =   1710
         Width           =   1665
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
         Height          =   300
         Index           =   10
         Left            =   60
         TabIndex        =   16
         Top             =   2025
         Width           =   1000
      End
      Begin VB.Label LBLNET 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1230
         TabIndex        =   15
         Top             =   2040
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase"
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
         Left            =   60
         TabIndex        =   14
         Top             =   1035
         Width           =   1000
      End
      Begin VB.Label LBLPTRXTOTAL 
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1230
         TabIndex        =   13
         Top             =   1050
         Width           =   1665
      End
      Begin VB.Label LBLCASHPAY 
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
         Height          =   390
         Left            =   1755
         TabIndex        =   12
         Top             =   5820
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Paid"
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
         Left            =   120
         TabIndex        =   11
         Top             =   5880
         Width           =   1590
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
         Left            =   5535
         TabIndex        =   10
         Top             =   3660
         Visible         =   0   'False
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
         Left            =   3510
         TabIndex        =   9
         Top             =   3480
         Visible         =   0   'False
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
         Left            =   3960
         TabIndex        =   8
         Top             =   3630
         Visible         =   0   'False
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
         Left            =   3510
         TabIndex        =   7
         Top             =   3570
         Visible         =   0   'False
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
         Left            =   3285
         TabIndex        =   6
         Top             =   180
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmDaybook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCunterSales_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTCOUNTERSALES"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            If Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            End If
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    If frmLogin.rs!Level = "5" Then
        Report.RecordSelectionFormula = "({TRXMAST.SYS_NAME}= '" & system_name & "' AND ({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "COUNTER WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CmdDayBook_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim i As Long
    Dim FROMDATE As Date
    Dim TODATE As Date
    
        
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    
    db.Execute "Delete from day_book"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM day_book", db, adOpenStatic, adLockOptimistic, adCmdText
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTO.Value, "MM,DD,YYYY")
        
        
    Dim OPVAL, CLOVAL, RCVDVAL, ISSVAL As Double
    
    OPVAL = 0
    CLOVAL = 0
    RCVDVAL = 0
    ISSVAL = 0
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select OPEN_DB from ACTMAST  WHERE ACT_CODE = '111001' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        OPVAL = IIf(IsNull(rstTRANX!OPEN_DB), 0, rstTRANX!OPEN_DB)
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(OPEN_DB) from ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3) ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        OPVAL = OPVAL + IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='S'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        RCVDVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='P'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        ISSVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    OPVAL = Round(OPVAL + (RCVDVAL - ISSVAL), 2)
    
'    RSTTRXFILE.AddNew
'    RSTTRXFILE!TRX_TYPE = "OC"
'    RSTTRXFILE!VCH_NO = 0
'    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
'    RSTTRXFILE!CR_AMT = OPVAL
'    RSTTRXFILE!DR_AMT = 0
'    RSTTRXFILE!VCH_DESC = "Opening Balance:...................."
'    RSTTRXFILE!REMARKS = ""
'    RSTTRXFILE!VCH_DATE = Format(DTFROM.Value, "dd/mm/yyyy")
'    RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
'    RSTTRXFILE.Update
    
    '===============
    
    
    FROMDATE = DTFROM.Value 'Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = DTTO.Value 'Format(DTTO.Value, "MM,DD,YYYY")
    i = 1
    Do Until FROMDATE > TODATE
        
        RCVDVAL = 0
        Set rstTRANX = New ADODB.Recordset
        'RSTTRXFILE.Open "SELECT SUM(NET_AMOUNT) FROM TRXMAST WHERE (ACT_CODE = '130000' Or ACT_CODE = '130001') AND VCH_DATE = '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI'", db, adOpenStatic, adLockReadOnly, adCmdText
        rstTRANX.Open "SELECT SUM(NET_AMOUNT) FROM TRXMAST WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_TYPE='HI'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            RCVDVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "HI"
        RSTTRXFILE!VCH_NO = 0
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!CR_AMT = RCVDVAL
        RSTTRXFILE!DR_AMT = 0
        RSTTRXFILE!VCH_DESC = "BY: B2C SALES:...................."
        RSTTRXFILE!REMARKS = ""
        RSTTRXFILE!VCH_DATE = Format(FROMDATE, "dd/mm/yyyy")
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!DAY_CHANGE = i
        RSTTRXFILE.Update
        
        RCVDVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(NET_AMOUNT) FROM TRXMAST WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_TYPE='GI'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            RCVDVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "GI"
        RSTTRXFILE!VCH_NO = 0
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!CR_AMT = RCVDVAL
        RSTTRXFILE!DR_AMT = 0
        RSTTRXFILE!VCH_DESC = "BY: B2B SALES:...................."
        RSTTRXFILE!REMARKS = ""
        RSTTRXFILE!VCH_DATE = Format(FROMDATE, "dd/mm/yyyy")
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!DAY_CHANGE = i
        RSTTRXFILE.Update
        
        
        RCVDVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(NET_AMOUNT) FROM TRXMAST WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_TYPE='SV'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            RCVDVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "SV"
        RSTTRXFILE!VCH_NO = 0
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!CR_AMT = RCVDVAL
        RSTTRXFILE!DR_AMT = 0
        RSTTRXFILE!VCH_DESC = "BY: Service Bills:...................."
        RSTTRXFILE!REMARKS = ""
        RSTTRXFILE!VCH_DATE = Format(FROMDATE, "dd/mm/yyyy")
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!DAY_CHANGE = i
        RSTTRXFILE.Update
        
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * FROM TRXMAST WHERE Not(ACT_CODE = '130000' Or ACT_CODE = '130001') AND VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' or TRX_TYPE='GI' or TRX_TYPE='SV') ORDER BY TRX_TYPE", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until rstTRANX.EOF
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = 0
            RSTTRXFILE!DR_AMT = rstTRANX!NET_AMOUNT
            Select Case rstTRANX!TRX_TYPE
                Case "SV"
                    RSTTRXFILE!VCH_DESC = "TO: Service Bill" & IIf(IsNull(rstTRANX!ACT_NAME), "", "(" & rstTRANX!ACT_NAME & ")")
                Case "HI"
                    RSTTRXFILE!VCH_DESC = "TO:  B2C Sales" & IIf(IsNull(rstTRANX!ACT_NAME), "", " (" & rstTRANX!ACT_NAME & ")")
                Case "GI"
                    RSTTRXFILE!VCH_DESC = "TO:  B2B Sales" & IIf(IsNull(rstTRANX!ACT_NAME), "", " (" & rstTRANX!ACT_NAME & ")")
                Case "WO"
                    RSTTRXFILE!VCH_DESC = "TO: Petty Sales" & IIf(IsNull(rstTRANX!ACT_NAME), "", " (" & rstTRANX!ACT_NAME & ")")
            End Select
            RSTTRXFILE!REMARKS = "Credit Sale"
            RSTTRXFILE!VCH_DATE = rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
            
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI' OR TRX_TYPE='LP') ORDER BY TRX_TYPE", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = 0
            RSTTRXFILE!DR_AMT = rstTRANX!NET_AMOUNT
            Select Case rstTRANX!TRX_TYPE
                Case "PI"
                    RSTTRXFILE!VCH_DESC = "TO: Purchase: " & rstTRANX!ACT_NAME
                Case "LP"
                    RSTTRXFILE!VCH_DESC = "TO: Local Purchase: " & rstTRANX!ACT_NAME
                Case "PW"
                    RSTTRXFILE!VCH_DESC = "TO: Petty Purchase: " & rstTRANX!ACT_NAME
            End Select
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
            
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = rstTRANX!NET_AMOUNT
            RSTTRXFILE!DR_AMT = 0
            Select Case rstTRANX!TRX_TYPE
                Case "PI"
                    RSTTRXFILE!VCH_DESC = "BY: Purchase: " & rstTRANX!ACT_NAME
                Case "LP"
                    RSTTRXFILE!VCH_DESC = "BY: Local Purchase: " & rstTRANX!ACT_NAME
                Case "PW"
                    RSTTRXFILE!VCH_DESC = "BY: Petty Purchase: " & rstTRANX!ACT_NAME
            End Select
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
            
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        'assets purchase
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select * from ASTRXMAST  WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='AP' OR TRX_TYPE='EP') ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until rstTRANX.EOF
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = 0
            RSTTRXFILE!DR_AMT = rstTRANX!NET_AMOUNT
            Select Case rstTRANX!TRX_TYPE
                Case "AP"
                    RSTTRXFILE!VCH_DESC = "TO: Asset Purchase: " & rstTRANX!ACT_NAME
                Case "EP"
                    RSTTRXFILE!VCH_DESC = "TO: Expense (Input Tax): " & rstTRANX!ACT_NAME
            End Select
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
            
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = rstTRANX!NET_AMOUNT
            RSTTRXFILE!DR_AMT = 0
            Select Case rstTRANX!TRX_TYPE
                Case "AP"
                    RSTTRXFILE!VCH_DESC = "BY: Asset Purchase: " & rstTRANX!ACT_NAME
                Case "EP"
                    RSTTRXFILE!VCH_DESC = "BY: Expense (Input Tax): " & rstTRANX!ACT_NAME
            End Select
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        
        ISSVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(VCH_AMOUNT) FROM TRXEXPMAST WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_TYPE='EX'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            ISSVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If ISSVAL > 0 Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "EX" 'rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = 0 'rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM) 'rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = 0
            RSTTRXFILE!DR_AMT = ISSVAL
            RSTTRXFILE!VCH_DESC = "TO: Office Expense"
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = Format(FROMDATE, "dd/mm/yyyy") 'rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy") 'rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
        End If
        
        RCVDVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(VCH_AMOUNT) FROM TRXINCMAST WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_TYPE='IN'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            RCVDVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If RCVDVAL > 0 Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "IN" 'rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = 0 'rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM) 'rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = RCVDVAL
            RSTTRXFILE!DR_AMT = 0
            RSTTRXFILE!VCH_DESC = "BY: Office Income"
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = Format(FROMDATE, "dd/mm/yyyy") 'rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy") 'rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
        End If
        
        ISSVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(VCH_AMOUNT) FROM TRXEXP_MAST WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_TYPE='EX'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            ISSVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If ISSVAL > 0 Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "ES" 'rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = 0 'rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM) 'rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = 0
            RSTTRXFILE!DR_AMT = ISSVAL
            RSTTRXFILE!VCH_DESC = "TO: Staff Expense"
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = Format(FROMDATE, "dd/mm/yyyy") 'rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy") 'rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
        End If
            
        ISSVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMOUNT) FROM STAFFPYMT WHERE INV_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_TYPE='EX'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            ISSVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If ISSVAL > 0 Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "SE" 'rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = 0 'rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM) 'rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = 0
            RSTTRXFILE!DR_AMT = ISSVAL
            RSTTRXFILE!VCH_DESC = "TO: Staff Expense"
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = Format(FROMDATE, "dd/mm/yyyy") 'rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy") 'rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
        End If
        
        ISSVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(VCH_AMOUNT) from RETURNMAST  WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            ISSVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        If ISSVAL > 0 Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "SR" 'rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = 0 'rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM) 'rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = 0
            RSTTRXFILE!DR_AMT = ISSVAL
            RSTTRXFILE!VCH_DESC = "TO: Sales Return: " '& rstTRANX!ACT_NAME
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = Format(FROMDATE, "dd/mm/yyyy") 'rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy") 'rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
        End If
        
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select * from RETURNMAST  WHERE POST_FLAG = 'N' AND VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until rstTRANX.EOF
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = rstTRANX!VCH_NO
            RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
            RSTTRXFILE!CR_AMT = rstTRANX!VCH_AMOUNT
            RSTTRXFILE!DR_AMT = 0
            RSTTRXFILE!VCH_DESC = "BY: Sales Return: " & rstTRANX!ACT_NAME
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = rstTRANX!VCH_DATE
            RSTTRXFILE!CREATE_DATE = rstTRANX!CREATE_DATE
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        'RSTTRXFILE!POST_FLAG
        
    ''    '========
    ''    Set rstTRANX = New ADODB.Recordset
    ''    rstTRANX.Open "SELECT * From RTRXFILE WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE='SV') ", db, adOpenStatic, adLockReadOnly
    ''    Do Until rstTRANX.EOF
    ''        RSTTRXFILE.AddNew
    ''        RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
    ''        RSTTRXFILE!VCH_NO = rstTRANX!VCH_NO
    ''        RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
    ''        RSTTRXFILE!CR_AMT = 0
    ''        RSTTRXFILE!DR_AMT = rstTRANX!TRX_TOTAL
    ''        RSTTRXFILE!VCH_DESC = "TO: Sales Return (Exchange): " & rstTRANX!ACT_NAME
    ''        RSTTRXFILE!Remarks = ""
    ''        RSTTRXFILE!VCH_DATE = rstTRANX!VCH_DATE
    ''        RSTTRXFILE!CREATE_DATE = rstTRANX!CREATE_DATE
    ''        RSTTRXFILE.Update
    ''        rstTRANX.MoveNext
    ''    Loop
    ''    rstTRANX.Close
    ''    Set rstTRANX = Nothing
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From CRDTPYMT WHERE RCPT_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "'  AND TRX_TYPE='PY' order by CR_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = rstTRANX!CR_NO
            RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
            RSTTRXFILE!DR_AMT = 0
            RSTTRXFILE!CR_AMT = rstTRANX!RCPT_AMOUNT
            RSTTRXFILE!VCH_DESC = "Payment(Sundry Credtors): " & rstTRANX!ACT_NAME
            RSTTRXFILE!REMARKS = ""
            RSTTRXFILE!VCH_DATE = rstTRANX!RCPT_DATE
            RSTTRXFILE!CREATE_DATE = IIf(IsDate(rstTRANX!ENTRY_DATE), rstTRANX!ENTRY_DATE, rstTRANX!RCPT_DATE)
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From DBTPYMT WHERE INV_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "'  AND (TRX_TYPE = 'RT' OR TRX_TYPE = 'RL' OR TRX_TYPE = 'PL' OR TRX_TYPE = 'FR' OR TRX_TYPE = 'FP') order by CR_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = rstTRANX!CR_NO
            RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
            Select Case rstTRANX!TRX_TYPE
                Case "RT"
                    RSTTRXFILE!VCH_DESC = "Receipt(Sundry Debtors): " & rstTRANX!ACT_NAME
                    RSTTRXFILE!DR_AMT = 0
                    RSTTRXFILE!CR_AMT = rstTRANX!RCPT_AMT
                Case "RL"
                    RSTTRXFILE!VCH_DESC = "Receipt(Money Lender): " & rstTRANX!ACT_NAME
                    RSTTRXFILE!DR_AMT = 0
                    RSTTRXFILE!CR_AMT = rstTRANX!RCPT_AMT
                Case "PL"
                    RSTTRXFILE!VCH_DESC = "Payment(Money Lender): " & rstTRANX!ACT_NAME
                    RSTTRXFILE!CR_AMT = 0
                    RSTTRXFILE!DR_AMT = rstTRANX!RCPT_AMT
                Case "FR"
                    RSTTRXFILE!VCH_DESC = "Receipt(Savings /Deposit): " & rstTRANX!ACT_NAME
                    RSTTRXFILE!DR_AMT = 0
                    RSTTRXFILE!CR_AMT = rstTRANX!RCPT_AMT
                Case "FP"
                    RSTTRXFILE!VCH_DESC = "Payment(Savings /Deposit): " & rstTRANX!ACT_NAME
                    RSTTRXFILE!CR_AMT = 0
                    RSTTRXFILE!DR_AMT = rstTRANX!RCPT_AMT
            End Select
            RSTTRXFILE!REMARKS = IIf(IsNull(rstTRANX!INV_NO) Or rstTRANX!INV_NO = 0, "", rstTRANX!INV_NO) & IIf(IsNull(rstTRANX!REF_NO) Or rstTRANX!REF_NO = "", "", ", " & rstTRANX!REF_NO)
            RSTTRXFILE!VCH_DATE = rstTRANX!INV_DATE
            RSTTRXFILE!CREATE_DATE = IIf(IsDate(rstTRANX!ENTRY_DATE), rstTRANX!ENTRY_DATE, rstTRANX!INV_DATE)
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From BANK_TRX WHERE TRX_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' ORDER BY BNK_SL_NO, TRX_DATE", db, adOpenForwardOnly
        Do Until rstTRANX.EOF
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTTRXFILE!VCH_NO = rstTRANX!BNK_SL_NO
            RSTTRXFILE!TRX_YEAR = rstTRANX!TRX_YEAR
            Select Case rstTRANX!TRX_TYPE
                Case "DR"
                    Select Case rstTRANX!BILL_TRX_TYPE
                        Case "PY"
                            RSTTRXFILE!VCH_DESC = "Payment: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case "WD"
                            RSTTRXFILE!VCH_DESC = "Withdrawal: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case "BC"
                            RSTTRXFILE!VCH_DESC = "Bank Charges: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case "DN"
                            RSTTRXFILE!VCH_DESC = "Credit Note: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case "CT"
                            RSTTRXFILE!VCH_DESC = "Contra: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case "DB"
                            RSTTRXFILE!VCH_DESC = "Credit Note: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case "EX"
                            RSTTRXFILE!VCH_DESC = "Office Expense: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case "ES"
                            RSTTRXFILE!VCH_DESC = "Staff Expense: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case "FP"
                            RSTTRXFILE!VCH_DESC = "Payment(Savings /Deposit): " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case "RD"
                            RSTTRXFILE!VCH_DESC = "Cheque  Returned (Debtor): " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                        Case Else
                            RSTTRXFILE!VCH_DESC = "Others: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!DR_AMT = 0
                            RSTTRXFILE!CR_AMT = rstTRANX!TRX_AMOUNT
                    End Select
                Case "CR"
                    Select Case rstTRANX!BILL_TRX_TYPE
                        Case "RT"
                            RSTTRXFILE!VCH_DESC = "Receipt: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!CR_AMT = 0
                            RSTTRXFILE!DR_AMT = rstTRANX!TRX_AMOUNT
                        Case "DP"
                            RSTTRXFILE!VCH_DESC = "Deposit: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!CR_AMT = 0
                            RSTTRXFILE!DR_AMT = rstTRANX!TRX_AMOUNT
                        Case "IN"
                            RSTTRXFILE!VCH_DESC = "Bank Interest: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!CR_AMT = 0
                            RSTTRXFILE!DR_AMT = rstTRANX!TRX_AMOUNT
                        Case "CN"
                            RSTTRXFILE!VCH_DESC = "Debit Note: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!CR_AMT = 0
                            RSTTRXFILE!DR_AMT = rstTRANX!TRX_AMOUNT
                        Case "CB"
                            RSTTRXFILE!VCH_DESC = "Debit Note: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!CR_AMT = 0
                            RSTTRXFILE!DR_AMT = rstTRANX!TRX_AMOUNT
                        Case "CT"
                            RSTTRXFILE!VCH_DESC = "Contra: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!CR_AMT = 0
                            RSTTRXFILE!DR_AMT = rstTRANX!TRX_AMOUNT
                        Case "FR"
                            RSTTRXFILE!VCH_DESC = "Receipt(Savings /Deposit): " & rstTRANX!ACT_NAME
                            RSTTRXFILE!CR_AMT = 0
                            RSTTRXFILE!DR_AMT = rstTRANX!TRX_AMOUNT
                        Case "RC"
                            RSTTRXFILE!VCH_DESC = "Cheque Returned(Credtor)" & rstTRANX!ACT_NAME
                            RSTTRXFILE!CR_AMT = 0
                            RSTTRXFILE!DR_AMT = rstTRANX!TRX_AMOUNT
                        Case Else
                            RSTTRXFILE!VCH_DESC = "Others: " & rstTRANX!ACT_NAME
                            RSTTRXFILE!CR_AMT = 0
                            RSTTRXFILE!DR_AMT = rstTRANX!TRX_AMOUNT
                            
                    End Select
            End Select
            RSTTRXFILE!REMARKS = IIf(IsNull(rstTRANX!BANK_NAME), "", rstTRANX!BANK_NAME) & IIf(IsNull(rstTRANX!REF_NO) Or rstTRANX!REF_NO = "", "", ", " & rstTRANX!REF_NO)
            RSTTRXFILE!VCH_DATE = rstTRANX!TRX_DATE
            RSTTRXFILE!CREATE_DATE = IIf(IsDate(rstTRANX!ENTRY_DATE), rstTRANX!ENTRY_DATE, rstTRANX!TRX_DATE)
            RSTTRXFILE!DAY_CHANGE = i
            RSTTRXFILE.Update
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        CLOVAL = 0
        RCVDVAL = 0
        ISSVAL = 0
        
'        Set rstTRANX = New ADODB.Recordset
'        rstTRANX.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND CHECK_FLAG='S'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
'            RCVDVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
'        End If
'        rstTRANX.Close
'        Set rstTRANX = Nothing
'
'        Set rstTRANX = New ADODB.Recordset
'        rstTRANX.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' AND CHECK_FLAG='P'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
'            ISSVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
'        End If
'        rstTRANX.Close
'        Set rstTRANX = Nothing
'
'        CLOVAL = Round(OPVAL + (RCVDVAL - ISSVAL), 2)
'
'        RSTTRXFILE.AddNew
'        RSTTRXFILE!TRX_TYPE = "CC"
'        RSTTRXFILE!VCH_NO = 0
'        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
'        RSTTRXFILE!CR_AMT = 0
'        RSTTRXFILE!DR_AMT = CLOVAL
'        RSTTRXFILE!VCH_DESC = "Closing Balance:...................."
'        RSTTRXFILE!REMARKS = ""
'        RSTTRXFILE!VCH_DATE = Format(FROMDATE, "dd/mm/yyyy")
'        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
'        RSTTRXFILE!DAY_CHANGE = i
'        RSTTRXFILE.Update
        
        FROMDATE = DateAdd("d", FROMDATE, 1)
        i = i + 1
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
        
    ReportNameVar = Rptpath & "RptDayBook"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@OPBAL}" Then CRXFormulaField.text = " " & OPVAL & " "
    Next
    frmreport.Caption = "DAY BOOK"
    Call GENERATEREPORT
    
    Screen.MousePointer = vbDefault
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub CmdDenom_Click()
    FrmDenom.Show
    FrmDenom.SetFocus
    FrmDenom.TxtCAmount.text = Val(LBLWCLOCASH.Caption)
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdoffExp_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTCOUNTEREXP"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            If Report.Database.Tables(i).Name = "TRXEXPMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            End If
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    If frmLogin.rs!Level = "5" Then
        Report.RecordSelectionFormula = "({TRXEXPMAST.SYS_NAME}= '" & system_name & "' AND {TRXEXPMAST.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXEXPMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "({TRXEXPMAST.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXEXPMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "COUNTER WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CmdPrint_Click()
    On Error GoTo ERRHAND
    Dim i As Integer
    Dim FROMDATE As Date
    Dim TODATE As Date
    
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTO.Value, "MM,DD,YYYY")
    ReportNameVar = Rptpath & "RptCashBook"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
    '<=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND VCH_DATE >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #
    
    'Report.RecordSelectionFormula = "(({CASHATRXFILE.VCH_DATE}<=# " & Format(DTTO.Value, "MM,DD,YYYY") & " #) AND ({CASHATRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #) AND (({CASHATRXFILE.INV_TYPE}='SI' OR {CASHATRXFILE.INV_TYPE}='WO') and {CASHATRXFILE.TRX_TYPE} ='DR') OR ({CASHATRXFILE.INV_TYPE}='RT' and {CASHATRXFILE.TRX_TYPE} ='DR') OR ({CASHATRXFILE.INV_TYPE}='PI' and {CASHATRXFILE.TRX_TYPE} ='DR') OR ({CASHATRXFILE.INV_TYPE}='PY' and {CASHATRXFILE.TRX_TYPE} ='DR') OR ({CASHATRXFILE.INV_TYPE}='EX' and {CASHATRXFILE.TRX_TYPE} ='DR'))"
    db.Execute "UPDATE CASHATRXFILE SET AMOUNT = 0 WHERE ISNULL(AMOUNT)"
    Report.RecordSelectionFormula = "(NOT ISNULL({CASHATRXFILE.AMOUNT}) AND {CASHATRXFILE.AMOUNT} <> 0 AND {CASHATRXFILE.VCH_DATE}<=# " & Format(DTTO.Value, "MM,DD,YYYY") & " #) AND ({CASHATRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    'Report.RecordSelectionFormula = "({CASHATRXFILE.VCH_DATE}<=# " & Format(DTTO.value, "MM,DD,YYYY") & " #) AND ({CASHATRXFILE.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@opcash}" Then CRXFormulaField.text = "'" & lblopcash.Caption & "' "
        If CRXFormulaField.Name = "{@clocash}" Then CRXFormulaField.text = "'" & lblcloscash.Caption & "' "
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "CASH BOOK"
    
    GENERATEREPORT
    'GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub CmdPrnCoolie_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTCOOLIE"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            End If
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "COOLIE REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdTax_Click()
    FrmTaxdetails.Show
    FrmTaxdetails.SetFocus
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
'    Dim rstdbt As ADODB.Recordset
'    Dim rstdbt2 As ADODB.Recordset
'    Set rstTRANX = New ADODB.Recordset
'    rstTRANX.Open "SELECT DISTINCT ACT_CODE From TRXMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO') ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly
'    Do Until rstTRANX.EOF
'        Set rstdbt = New ADODB.Recordset
'        rstdbt.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "' and TRX_TYPE = 'DR'", db, adOpenStatic, adLockOptimistic, adCmdText
'        Do Until rstdbt.EOF
'
'            Set rstdbt2 = New ADODB.Recordset
'            rstdbt2.Open "select SUM(RCPT_AMOUNT) from trnxrcpt WHERE ACT_CODE = '" & rstdbt!ACT_CODE & "' AND INV_NO  = " & rstdbt!INV_NO & " AND INV_TRX_TYPE = '" & rstdbt!INV_TRX_TYPE & "' AND INV_TRX_YEAR = '" & rstdbt!TRX_YEAR & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
'                rstdbt!RCVD_AMOUNT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
'                rstdbt.Update
'                'db.Execute "Update DBTPYMT set RCVD_AMOUNT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0)) where ACT_CODE = '" & rstdbt!ACT_CODE & "' AND TRX_TYPE = 'DR' AND INV_TRX_TYPE  = '" & rstdbt!TRX_TYPE & "' AND INV_NO = '" & rstdbt!VCH_NO & "' AND TRX_YEAR = '" & rstdbt!TRX_YEAR & "'"
'                'lblsaleret.Caption = Format(IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0)), "0.00")
'            End If
'            rstdbt2.Close
'            Set rstdbt2 = Nothing
'
'            rstdbt.MoveNext
'        Loop
'        rstdbt.Close
'        Set rstdbt = Nothing
'
'        rstTRANX.MoveNext
'    Loop
'    rstTRANX.Close
'    Set rstTRANX = Nothing
    
    On Error GoTo ERRHAND
    Dim rstdbt As ADODB.Recordset
    Dim rstdbt2 As ADODB.Recordset
    Set rstdbt = New ADODB.Recordset
    rstdbt.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstdbt.EOF
            
        Set rstdbt2 = New ADODB.Recordset
        rstdbt2.Open "select SUM(RCPT_AMOUNT) from trnxrcpt WHERE ACT_CODE = '" & rstdbt!ACT_CODE & "' AND INV_NO  = " & rstdbt!VCH_NO & " AND INV_TRX_TYPE = '" & rstdbt!TRX_TYPE & "' AND INV_TRX_YEAR = '" & rstdbt!TRX_YEAR & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstdbt2.EOF And rstdbt2.BOF) Then
            rstdbt!RCVD_AMOUNT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
            rstdbt.Update
            'db.Execute "Update DBTPYMT set RCVD_AMOUNT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0)) where ACT_CODE = '" & rstdbt!ACT_CODE & "' AND TRX_TYPE = 'DR' AND INV_TRX_TYPE  = '" & rstdbt!TRX_TYPE & "' AND INV_NO = '" & rstdbt!VCH_NO & "' AND TRX_YEAR = '" & rstdbt!TRX_YEAR & "'"
            'lblsaleret.Caption = Format(IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0)), "0.00")
        End If
        rstdbt2.Close
        Set rstdbt2 = Nothing
            
        rstdbt.MoveNext
    Loop
    rstdbt.Close
    Set rstdbt = Nothing
    
    ReportNameVar = Rptpath & "RPTCRSALEREP"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If LBLPETTYSALE(4).Visible = True Then
        Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='RI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    End If
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" Then
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "CASH / CREDIT REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbHourglass
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Command3_Click()
    
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    ReportNameVar = Rptpath & "RPTSALEITEMS"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            End If
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    If LBLPETTYSALE(4).Visible = True Then
        Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    If frmLogin.rs!Level <> "0" Then
        Frame1.Visible = False
        LBLCOST.Visible = False
        LBLTOTAL(7).Visible = False
    End If
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        LBLB2BSALES(11).Visible = False
        LBLB2B.Visible = False
        LBLTOTAL(13).Caption = "Sales (M)"
        LBLTOTAL(12).Caption = "Sales"
        
        LBLTOTAL(4).Visible = False
        LBLTOTAL(15).Visible = False
        LBLTOTAL(11).Visible = False
        
        lbltaxamt.Visible = False
        LBLdifftax.Visible = False
        lblcess.Visible = False
        
        Cmdtax.Visible = False
    End If
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
    
    GRDBILL.ColAlignment(0) = 4
    GRDBILL.ColAlignment(1) = 1
    GRDBILL.ColAlignment(2) = 1
    GRDBILL.ColAlignment(3) = 4
    GRDBILL.ColAlignment(4) = 4
    GRDBILL.ColAlignment(5) = 4
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
    
    grdrcpt.ColAlignment(0) = 4
    grdrcpt.ColAlignment(1) = 1
    grdrcpt.ColAlignment(2) = 1
    grdrcpt.ColAlignment(3) = 4
    grdrcpt.ColAlignment(4) = 4
    grdrcpt.ColAlignment(5) = 4
    grdrcpt.ColAlignment(6) = 1
    
    GRDTranx.FixedRows = 0
    GRDTranx.rows = 1
    
    DTFROM.Value = Format(Date, "DD/MM/YYYY")
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 9585
    'Me.Height = 10185
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub CmDDisplay_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim i As Long
    Dim FROMDATE As Date
    Dim TODATE As Date
    LBLBTRXTOTAL.Caption = ""
    LBLDISCOUNT.Caption = ""
    LBLNET.Caption = ""
    LBLCOST.Caption = ""
    LBLPROFIT.Caption = "0.00"
    LBLEXPENSE.Caption = "0.00"
    LBLINCOME.Caption = "0.00"
    LblStaff.Caption = "0.00"
    lblCommi.Caption = "0.00"
    lblcashsale.Caption = ""
    lblcrdtsale.Caption = ""
    lblhandle.Caption = ""
    lblServicepaid.Caption = ""
    LblPurchret.Caption = "0.00"
    
    LBLPTRXTOTAL.Caption = "0.00"
    LblPettyPur.Caption = "0.00"
    LBLCASHPUR.Caption = "0.00"
    LBLCRDTPUR.Caption = "0.00"

    LBLCASHPAY.Caption = "0.00"
    lblcashrcv.Caption = "0.00"
    
    lblopcash.Caption = "0.00"
    lblcloscash.Caption = "0.00"
    
    LBLSERVICE.Caption = "0.00"
    LBLPETTY.Caption = "0.00"
    LBLB2C.Caption = "0.00"
    LBLB2B.Caption = "0.00"
    
    lblassets.Caption = "0.00"
    lblexptax.Caption = "0.00"
    'vbalProgressBar1.Value = 0
    'vbalProgressBar1.ShowText = True
        
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI' OR TRX_TYPE='LP')", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        LBLPTRXTOTAL.Caption = Format(Val(LBLPTRXTOTAL.Caption) + rstTRANX!NET_AMOUNT, "0.00")
        If (rstTRANX!POST_FLAG = "Y") Then
            LBLCASHPUR.Caption = Format(Val(LBLCASHPUR.Caption) + rstTRANX!NET_AMOUNT, "0.00")
        Else
            LBLCRDTPUR.Caption = Format(Val(LBLCRDTPUR.Caption) + rstTRANX!NET_AMOUNT, "0.00")
        End If
        'vbalProgressBar1.Max = rstTRANX.RecordCount
        'vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PW')", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        LblPettyPur.Caption = Format(Val(LblPettyPur.Caption) + rstTRANX!NET_AMOUNT, "0.00")
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    '===========
    'assets purchase
    lblassets.Caption = ""
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(NET_AMOUNT) from ASTRXMAST  WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='AP' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblassets.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    '============
    
    '===========
    'EXPENSE (INPUT TAX)
    lblexptax.Caption = ""
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(NET_AMOUNT) from ASTRXMAST  WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='EP' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblexptax.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    '============
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(VCH_AMOUNT) From TRXEXPMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='EX'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        LBLEXPENSE.Caption = Format(Val(LBLEXPENSE.Caption) + IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(VCH_AMOUNT) From TRXINCMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='IN'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        LBLINCOME.Caption = Format(Val(LBLINCOME.Caption) + IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(VCH_AMOUNT) From TRXEXP_MAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='EX'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        LblStaff.Caption = Format(Val(LblStaff.Caption) + IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
        
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(RCPT_AMOUNT) from STAFFPYMT WHERE INV_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='EX'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        LblStaff.Caption = Format(Val(LblStaff.Caption) + IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(COMM_AMT) From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SI' OR TRX_TYPE='RI' OR TRX_TYPE='WO' OR TRX_TYPE='VI')", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblCommi.Caption = Format(Val(lblCommi.Caption) + IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SI' OR TRX_TYPE='RI' OR TRX_TYPE='WO' OR TRX_TYPE='VI')", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        Select Case rstTRANX!TRX_TYPE
            Case "SV"
                LBLSERVICE.Caption = Val(LBLPETTY.Caption) + rstTRANX!NET_AMOUNT
            Case "HI"
                LBLB2C.Caption = Val(LBLB2C.Caption) + rstTRANX!NET_AMOUNT
            Case "GI"
                LBLB2B.Caption = Val(LBLB2B.Caption) + rstTRANX!NET_AMOUNT
            Case "WO"
                LBLPETTY.Caption = Val(LBLPETTY.Caption) + rstTRANX!NET_AMOUNT
        End Select
        LBLBTRXTOTAL.Caption = Val(LBLBTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT
        LBLDISCOUNT.Caption = Val(LBLDISCOUNT.Caption) + rstTRANX!DISCOUNT
        'LBLNET.Caption = Format(Val(LBLBTRXTOTAL.Caption) - Val(LBLDISCOUNT.Caption), "0.00")
        LBLNET.Caption = Val(LBLNET.Caption) + rstTRANX!NET_AMOUNT
        LBLCOST.Caption = Val(LBLCOST.Caption) + rstTRANX!PAY_AMOUNT
        lblhandle.Caption = Val(lblhandle.Caption) + IIf(IsNull(rstTRANX!HANDLE), 0, rstTRANX!HANDLE) + IIf(IsNull(rstTRANX!FRIEGHT), 0, rstTRANX!FRIEGHT)
        
        'If (rstTRANX!POST_FLAG = "Y") Then
        '    lblcashsale.Caption = Format(Val(lblcashsale.Caption) + rstTRANX!NET_AMOUNT, "0.00")
        'Else
            'lblcrdtsale.Caption = Format(Val(lblcrdtsale.Caption) + rstTRANX!NET_AMOUNT, "0.00")
        'End If
        'If (rstTRANX!POST_FLAG = "Y") Then
'            lblcashsale.Caption = Format(Val(lblcashsale.Caption) + rstTRANX!NET_AMOUNT, "0.00")
'        Else
'            lblcrdtsale.Caption = Format(Val(lblcrdtsale.Caption) + rstTRANX!NET_AMOUNT, "0.00")
'        End If
        If (rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001") Then
            lblcashsale.Caption = Val(lblcashsale.Caption) + rstTRANX!NET_AMOUNT
        Else
            lblcrdtsale.Caption = Val(lblcrdtsale.Caption) + rstTRANX!NET_AMOUNT
        End If
        'vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    LBLBTRXTOTAL.Caption = Format(Val(LBLBTRXTOTAL.Caption), "0.00")
    LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption), "0.00")
    LBLNET.Caption = Format(Val(LBLNET.Caption), "0.00")
    LBLCOST.Caption = Format(Val(LBLCOST.Caption), "0.00")
    lblhandle.Caption = Format(Val(lblhandle.Caption), "0.00")
    lblcashsale.Caption = Format(Val(lblcashsale.Caption), "0.00")
    lblcrdtsale.Caption = Format(Val(lblcrdtsale.Caption), "0.00")
        
    LBLSERVICE.Caption = Format(Val(LBLSERVICE.Caption), "0.00")
    LBLPETTY.Caption = Format(Val(LBLPETTY.Caption), "0.00")
    LBLB2C.Caption = Format(Val(LBLB2C.Caption), "0.00")
    LBLB2B.Caption = Format(Val(LBLB2B.Caption), "0.00")
    
    '===========
    lblsaleret.Caption = ""
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(VCH_AMOUNT) from RETURNMAST  WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblsaleret.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    '===========
    LblPurchret.Caption = ""
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select SUM(VCH_AMOUNT) from PURCAHSERETURN  WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='PR' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        LblPurchret.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    '========
    
    lblxchange.Caption = ""
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(TRX_TOTAL) From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE='SV') ", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblxchange.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    lblServicepaid.Caption = ""
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(TRX_TOTAL) From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CATEGORY = 'SERVICE CHARGE'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblServicepaid.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    GRDBILL.rows = 1
    i = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From CRDTPYMT WHERE RCPT_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND RCPT_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'  AND TRX_TYPE='PY' order by CR_NO", db, adOpenStatic, adLockReadOnly
    'rstTRANX.Open "SELECT * From CRDTPYMT WHERE RCPT_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND RCPT_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        i = i + 1
        GRDBILL.rows = GRDBILL.rows + 1
        GRDBILL.FixedRows = 1
        GRDBILL.TextMatrix(i, 0) = i
        GRDBILL.TextMatrix(i, 1) = rstTRANX!ACT_NAME
        GRDBILL.TextMatrix(i, 2) = Format(rstTRANX!RCPT_AMOUNT, "0.00")
        GRDBILL.TextMatrix(i, 3) = Format(rstTRANX!RCPT_DATE, "DD/MM/YYYY")
        'GRDBILL.TextMatrix(i, 4) = rstTRANX!INV_NO & " dtd " & Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        GRDBILL.TextMatrix(i, 4) = IIf(IsNull(rstTRANX!INV_DATE), "", Format(rstTRANX!INV_DATE, "DD/MM/YYYY"))
        GRDBILL.TextMatrix(i, 5) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        GRDBILL.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        'LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + rstTRANX!INV_AMT, "0.00")
        LBLCASHPAY.Caption = Format(Val(LBLCASHPAY.Caption) + rstTRANX!RCPT_AMOUNT, "0.00")
        'lblbalamt.Caption = Format(Val(LBLINVAMT.Caption) - Val(LBLPAIDAMT.Caption), "0.00")
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    grdrcpt.rows = 1
    i = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From DBTPYMT WHERE INV_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'  AND TRX_TYPE='RT' order by CR_NO", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        i = i + 1
        grdrcpt.rows = grdrcpt.rows + 1
        grdrcpt.FixedRows = 1
        grdrcpt.TextMatrix(i, 0) = i
        grdrcpt.TextMatrix(i, 1) = rstTRANX!ACT_NAME
        grdrcpt.TextMatrix(i, 2) = Format(rstTRANX!RCPT_AMT, "0.00")
        grdrcpt.TextMatrix(i, 3) = Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        'GRDBILL.TextMatrix(i, 4) = rstTRANX!INV_NO & " dtd " & Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        grdrcpt.TextMatrix(i, 4) = IIf(IsNull(rstTRANX!INV_DATE), "", Format(rstTRANX!INV_DATE, "DD/MM/YYYY"))
        grdrcpt.TextMatrix(i, 5) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        grdrcpt.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        lblcashrcv.Caption = Format(Val(lblcashrcv.Caption) + rstTRANX!RCPT_AMT, "0.00")
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    
'    Set rstTRANX = New ADODB.Recordset
'    rstTRANX.Open "SELECT * From TRNXRCPT WHERE RCPT_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND RCPT_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "'  AND TRX_TYPE='RT' order by RCPT_NO", db, adOpenStatic, adLockReadOnly
'    Do Until rstTRANX.EOF
'        i = i + 1
'        grdrcpt.Rows = grdrcpt.Rows + 1
'        grdrcpt.FixedRows = 1
'        grdrcpt.TextMatrix(i, 0) = i
'        grdrcpt.TextMatrix(i, 1) = rstTRANX!ACT_NAME
'        grdrcpt.TextMatrix(i, 2) = Format(rstTRANX!RCPT_AMOUNT, "0.00")
'        grdrcpt.TextMatrix(i, 3) = Format(rstTRANX!RCPT_DATE, "DD/MM/YYYY")
'        'grdrcpt.TextMatrix(i, 4) = rstTRANX!INV_NO & " dtd " & Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
'        grdrcpt.TextMatrix(i, 4) = IIf(IsNull(rstTRANX!INV_DATE), "", Format(rstTRANX!INV_DATE, "DD/MM/YYYY"))
'        grdrcpt.TextMatrix(i, 5) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
'        grdrcpt.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
'        'LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + rstTRANX!INV_AMT, "0.00")
'        lblcashrcv.Caption = Format(Val(lblcashrcv.Caption) + rstTRANX!RCPT_AMOUNT, "0.00")
'        'lblbalamt.Caption = Format(Val(LBLINVAMT.Caption) - Val(LBLPAIDAMT.Caption), "0.00")
'        rstTRANX.MoveNext
'    Loop
'    rstTRANX.Close
'    Set rstTRANX = Nothing
'
    'LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(LBLEXPENSE.Caption) + Val(LblStaff.Caption) + Val(lblCommi.Caption)), "0.00")
    
    
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
    RSTTRXFILE.Open "select SUM(OPEN_DB) from ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3) ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OPVAL = OPVAL + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTO.Value, "MM,DD,YYYY")
    
    RCVDVAL = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='S'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RCVDVAL = RCVDVAL + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        'RSTTRXFILE.MoveNext
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    ISSVAL = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='P'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        ISSVAL = ISSVAL + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        'RSTTRXFILE.MoveNext
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    lblopcash.Caption = Round(OPVAL + (RCVDVAL - ISSVAL), 2)

    lblpaidcash.Caption = 0
    lblrcvdcash.Caption = 0

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM CASHATRXFILE WHERE VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        Select Case RSTTRXFILE!check_flag
            Case "S"
                lblrcvdcash.Caption = Val(lblrcvdcash.Caption) + RSTTRXFILE!AMOUNT
            Case "P"
                lblpaidcash.Caption = Val(lblpaidcash.Caption) + RSTTRXFILE!AMOUNT
        End Select
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    lblcloscash.Caption = Round(Val(lblopcash.Caption) + (Val(lblrcvdcash.Caption) - Val(lblpaidcash.Caption)), 2)
    If LBLPETTY.Visible = True Then
        LBLWCLOCASH.Caption = Format(lblcloscash.Caption, "0.00")
    Else
        LBLWCLOCASH.Caption = Format(Round(Val(lblcloscash.Caption) - Val(LBLPETTY.Caption), 2), "0.00")
    End If
    
    GRDTranx.FixedRows = 0
    GRDTranx.rows = 1
    GRDTranx.TextMatrix(0, 0) = "Sl"
    GRDTranx.TextMatrix(0, 1) = "Date"
    GRDTranx.TextMatrix(0, 2) = "Type"
    GRDTranx.TextMatrix(0, 3) = "Credit"
    GRDTranx.TextMatrix(0, 4) = "Debit"
    GRDTranx.TextMatrix(0, 5) = "TRX No"
    GRDTranx.TextMatrix(0, 6) = "Bank"
    GRDTranx.TextMatrix(0, 7) = "PARTY"
    GRDTranx.TextMatrix(0, 8) = "REMARKS"
    
    GRDTranx.ColWidth(0) = 700
    GRDTranx.ColWidth(1) = 1100
    GRDTranx.ColWidth(2) = 1000
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1100
    GRDTranx.ColWidth(5) = 1100
    GRDTranx.ColWidth(6) = 1600
    GRDTranx.ColWidth(7) = 1600
    GRDTranx.ColWidth(8) = 2500
    LBLINVAMT.Caption = "0.00"
    LBLPAIDAMT.Caption = "0.00"
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 1
    'GRDTranx.ColAlignment(3) = 1
    'GRDTranx.ColAlignment(4) = 1
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 1
    GRDTranx.ColAlignment(7) = 1
    GRDTranx.ColAlignment(8) = 1
    
    grdlend.FixedRows = 0
    grdlend.rows = 1
    grdlend.TextMatrix(0, 0) = "SL"
    grdlend.TextMatrix(0, 1) = "DESCRIPTION"
    grdlend.TextMatrix(0, 2) = "DATE"
    grdlend.TextMatrix(0, 3) = "TRX NO"
    grdlend.TextMatrix(0, 4) = "RCVD AMT"
    grdlend.TextMatrix(0, 5) = "PAID AMT"
    grdlend.TextMatrix(0, 6) = "REF NO"
    grdlend.TextMatrix(0, 7) = "CR NO"
    
    grdlend.ColWidth(0) = 600
    grdlend.ColWidth(1) = 2000
    grdlend.ColWidth(2) = 1100
    grdlend.ColWidth(3) = 1200
    grdlend.ColWidth(4) = 1200
    grdlend.ColWidth(5) = 1200
    grdlend.ColWidth(6) = 1600
    grdlend.ColWidth(7) = 0
        
    grdlend.ColAlignment(0) = 4
    grdlend.ColAlignment(1) = 1
    grdlend.ColAlignment(2) = 4
    grdlend.ColAlignment(3) = 4
    grdlend.ColAlignment(4) = 4
    grdlend.ColAlignment(5) = 4
    grdlend.ColAlignment(6) = 4
    
    i = 1
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From BANK_TRX WHERE TRX_DATE <='" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND TRX_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY BNK_SL_NO, TRX_DATE", db, adOpenForwardOnly
    Do Until rstTRANX.EOF
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0
        GRDTranx.TextMatrix(i, 0) = i
        GRDTranx.TextMatrix(i, 1) = Format(rstTRANX!TRX_DATE, "DD/MM/YYYY")
        Select Case rstTRANX!TRX_TYPE
            Case "DR"
                Select Case rstTRANX!BILL_TRX_TYPE
                    Case "PY"
                        GRDTranx.TextMatrix(i, 2) = "Payment"
                    Case "WD"
                        GRDTranx.TextMatrix(i, 2) = "Withdrawal"
                    Case "BC"
                        GRDTranx.TextMatrix(i, 2) = "Bank Charges"
                    Case "DN"
                        GRDTranx.TextMatrix(i, 2) = "Credit Note"
                    Case "CT"
                        GRDTranx.TextMatrix(i, 2) = "Contra"
                    Case "DB"
                        GRDTranx.TextMatrix(i, 2) = "Credit Note"
                    Case "EX"
                        GRDTranx.TextMatrix(i, 2) = "Off Expense"
                    Case "ES"
                        GRDTranx.TextMatrix(i, 2) = "Staff Expense"
                    Case "FP"
                        GRDTranx.TextMatrix(i, 2) = "Payment(Deposit)"
                    Case "RD"
                        GRDTranx.TextMatrix(i, 2) = "Cheque  Returned (Debtor)"
                    Case Else
                        GRDTranx.TextMatrix(i, 2) = "Others"
                End Select
                GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!TRX_AMOUNT, "0.00")
'                GRDTranx.TextMatrix(i, 0) = "Sale"
'                GRDTranx.CellForeColor = vbRed
            Case "CR"
                Select Case rstTRANX!BILL_TRX_TYPE
                    Case "RT"
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
                        GRDTranx.TextMatrix(i, 2) = "Cheque Returned(Credtor)"
                    Case Else
                        GRDTranx.TextMatrix(i, 2) = "Others"
                End Select
                GRDTranx.TextMatrix(i, 3) = Format(rstTRANX!TRX_AMOUNT, "0.00")
        End Select
        GRDTranx.TextMatrix(i, 5) = IIf(IsNull(rstTRANX!TRX_NO), "", rstTRANX!TRX_NO)
        GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!BANK_NAME), "", rstTRANX!BANK_NAME)
        GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        
        LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + Val(GRDTranx.TextMatrix(i, 3)), "0.00")
        LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + Val(GRDTranx.TextMatrix(i, 4)), "0.00")
        
        'GRDTranx.Row = i
        'GRDTranx.Col = 0
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    i = 1
    LBLINVAMT.Tag = ""
    LBLPAIDAMT.Tag = ""
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From DBTPYMT WHERE (TRX_TYPE = 'RL' OR TRX_TYPE = 'PL') AND RCPT_DATE <='" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND RCPT_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE DESC", db, adOpenForwardOnly
    Do Until rstTRANX.EOF
        grdlend.rows = grdlend.rows + 1
        grdlend.FixedRows = 1
        grdlend.Row = i
        grdlend.Col = 0

        grdlend.TextMatrix(i, 0) = i
        grdlend.TextMatrix(i, 2) = Format(rstTRANX!RCPT_DATE, "DD/MM/YYYY")
        grdlend.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        Select Case rstTRANX!TRX_TYPE
            Case "RL"
                grdlend.TextMatrix(i, 1) = "Receipt" & IIf(IsNull(rstTRANX!ACT_NAME), "", " (" & rstTRANX!ACT_NAME) & ")"
                grdlend.TextMatrix(i, 4) = Format(rstTRANX!RCPT_AMT, "0.00")
                grdlend.CellForeColor = vbRed
            Case "PL"
                grdlend.TextMatrix(i, 1) = "Payment" & IIf(IsNull(rstTRANX!ACT_NAME), "", " (" & rstTRANX!ACT_NAME) & ")"
                grdlend.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                grdlend.CellForeColor = vbBlue
        End Select
        grdlend.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        grdlend.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!CR_NO), "", rstTRANX!CR_NO)
        grdlend.Row = i
        grdlend.Col = 0
        LBLINVAMT.Tag = Val(LBLINVAMT.Tag) + Val(grdlend.TextMatrix(i, 4))
        LBLPAIDAMT.Tag = Val(LBLPAIDAMT.Tag) + Val(grdlend.TextMatrix(i, 5))
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    If i > 1 Then
       grdlend.rows = grdlend.rows + 1
       grdlend.TextMatrix(i, 1) = "TOTAL"
       grdlend.TextMatrix(i, 4) = Format(LBLINVAMT.Tag, "0.00")
       grdlend.TextMatrix(i, 5) = Format(LBLPAIDAMT.Tag, "0.00")
    End If
     
    grddeposit.FixedRows = 0
    grddeposit.rows = 1
    grddeposit.TextMatrix(0, 0) = "SL"
    grddeposit.TextMatrix(0, 1) = "DESCRIPTION"
    grddeposit.TextMatrix(0, 2) = "DATE"
    grddeposit.TextMatrix(0, 3) = "TRX NO"
    grddeposit.TextMatrix(0, 4) = "RCVD AMT"
    grddeposit.TextMatrix(0, 5) = "PAID AMT"
    grddeposit.TextMatrix(0, 6) = "REF NO"
    grddeposit.TextMatrix(0, 7) = "CR NO"
    
    grddeposit.ColWidth(0) = 600
    grddeposit.ColWidth(1) = 2000
    grddeposit.ColWidth(2) = 1100
    grddeposit.ColWidth(3) = 1200
    grddeposit.ColWidth(4) = 1200
    grddeposit.ColWidth(5) = 1200
    grddeposit.ColWidth(6) = 1600
    grddeposit.ColWidth(7) = 0
        
    grddeposit.ColAlignment(0) = 4
    grddeposit.ColAlignment(1) = 1
    grddeposit.ColAlignment(2) = 4
    grddeposit.ColAlignment(3) = 4
    grddeposit.ColAlignment(4) = 4
    grddeposit.ColAlignment(5) = 4
    grddeposit.ColAlignment(6) = 4
    
    i = 1
    LBLINVAMT.Tag = ""
    LBLPAIDAMT.Tag = ""
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From DBTPYMT WHERE (TRX_TYPE = 'FR' OR TRX_TYPE = 'FP') AND INV_DATE <='" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE DESC", db, adOpenForwardOnly
    Do Until rstTRANX.EOF
        grddeposit.rows = grddeposit.rows + 1
        grddeposit.FixedRows = 1
        grddeposit.Row = i
        grddeposit.Col = 0

        grddeposit.TextMatrix(i, 0) = i
        grddeposit.TextMatrix(i, 2) = Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        grddeposit.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        Select Case rstTRANX!TRX_TYPE
            Case "FR"
                grddeposit.TextMatrix(i, 1) = "Receipt" & IIf(IsNull(rstTRANX!ACT_NAME), "", " (" & rstTRANX!ACT_NAME) & ")"
                grddeposit.TextMatrix(i, 4) = Format(rstTRANX!RCPT_AMT, "0.00")
                grddeposit.CellForeColor = vbRed
            Case "FP"
                grddeposit.TextMatrix(i, 1) = "Payment" & IIf(IsNull(rstTRANX!ACT_NAME), "", " (" & rstTRANX!ACT_NAME) & ")"
                grddeposit.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                grddeposit.CellForeColor = vbBlue
        End Select
        grddeposit.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        grddeposit.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!CR_NO), "", rstTRANX!CR_NO)
        grddeposit.Row = i
        grddeposit.Col = 0
        LBLINVAMT.Tag = Val(LBLINVAMT.Tag) + Val(grddeposit.TextMatrix(i, 4))
        LBLPAIDAMT.Tag = Val(LBLPAIDAMT.Tag) + Val(grddeposit.TextMatrix(i, 5))
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    If i > 1 Then
       grddeposit.rows = grddeposit.rows + 1
       grddeposit.TextMatrix(i, 1) = "TOTAL"
       grddeposit.TextMatrix(i, 4) = Format(LBLINVAMT.Tag, "0.00")
       grddeposit.TextMatrix(i, 5) = Format(LBLPAIDAMT.Tag, "0.00")
    End If
    
    
    Dim KFC As Double
    Dim CESSPER As Double
    Dim CESSAMT As Double
    Dim TAX_AMT As Double
    Dim PUR_TAX As Double
    Dim PUR_CESS As Double
    Dim PUR_ADDCESS As Double
    
    CESSPER = 0
    CESSAMT = 0
    KFC = 0
    TAX_AMT = 0
    PUR_TAX = 0
    PUR_CESS = 0
    PUR_ADDCESS = 0
    
    If MDIMAIN.lblgst.Caption = "C" Then
        'TAX_AMT = Val(LBLB2B.Caption) * 1 / 100
        TAX_AMT = Val(LBLB2C.Caption) * 1 / 100
        TAX_AMT = TAX_AMT + Val(LBLSERVICE.Caption) * 1 / 100
    ElseIf MDIMAIN.lblgst.Caption = "N" Then
        TAX_AMT = 0
    Else
        Dim RSTtax As ADODB.Recordset
        Dim rstTRANX2 As ADODB.Recordset
        Set rstTRANX2 = New ADODB.Recordset
        rstTRANX2.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SV') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX2.EOF
            
            'B2C
            Set rstTRANX = New ADODB.Recordset
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI'  ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            If Not (rstTRANX.EOF And rstTRANX.BOF) Then
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & rstTRANX2!SALES_TAX & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        Select Case RSTTRXFILE!SLSM_CODE
                            Case "P"
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                            Case Else
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                        End Select
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                        TAX_AMT = TAX_AMT + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
    
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                        'If Not (IsNull(RSTtax!PUR_TAX)) Then
                            PUR_TAX = PUR_TAX + (IIf(IsNull(RSTtax!ITEM_COST), 0, RSTtax!ITEM_COST) * IIf(IsNull(RSTtax!PUR_TAX) Or RSTtax!PUR_TAX = 0, RSTtax!SALES_TAX, RSTtax!PUR_TAX) / 100) * RSTtax!QTY
                        'End If
                        If Not (IsNull(RSTtax!CESS_PER)) Then
                            PUR_CESS = PUR_CESS + (IIf(IsNull(RSTtax!ITEM_COST), 0, RSTtax!ITEM_COST) * RSTtax!CESS_PER / 100) * RSTtax!QTY
                        End If
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    'GRDPranx.TextMatrix(M, N) = Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.Tag) * Val(RSTtax!QTY)
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                'GRDPranx.TextMatrix(M, N) = Format(Round(Val(GRDPranx.TextMatrix(M, N)), 3), "0.00")
                'TOTAL_AMT = TOTAL_AMT + Val(GRDPranx.TextMatrix(M, N)) + TAX_AMT
                
                'GRDPranx.TextMatrix(M, N) = Format(Round(CESSPER, 3), "0.00")
                'GRDPranx.TextMatrix(M, N + 1) = Format(Round(CESSAMT, 3), "0.00")
                'GRDPranx.TextMatrix(M, N + 2) = Format(Round(KFC, 3), "0.00")
                'GRDPranx.TextMatrix(M, 4) = TOTAL_AMT + KFC + Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.TextMatrix(M, N + 1))
            End If
            rstTRANX.Close
            Set rstTRANX = Nothing
            
            'B2B
            Set rstTRANX = New ADODB.Recordset
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            If Not (rstTRANX.EOF And rstTRANX.BOF) Then
                
                
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & rstTRANX2!SALES_TAX & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        Select Case RSTTRXFILE!SLSM_CODE
                            Case "P"
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                            Case Else
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                        End Select
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                        TAX_AMT = TAX_AMT + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
    
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                        'If Not (IsNull(RSTtax!PUR_TAX)) Then
                            PUR_TAX = PUR_TAX + (IIf(IsNull(RSTtax!ITEM_COST), 0, RSTtax!ITEM_COST) * IIf(IsNull(RSTtax!PUR_TAX) Or RSTtax!PUR_TAX = 0, RSTtax!SALES_TAX, RSTtax!PUR_TAX) / 100) * RSTtax!QTY
                        'End If
                        If Not (IsNull(RSTtax!CESS_PER)) Then
                            PUR_CESS = PUR_CESS + (IIf(IsNull(RSTtax!ITEM_COST), 0, RSTtax!ITEM_COST) * RSTtax!CESS_PER / 100) * RSTtax!QTY
                        End If
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    'GRDPranx.TextMatrix(M, N) = Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.Tag) * Val(RSTtax!QTY)
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
            End If
            rstTRANX.Close
            Set rstTRANX = Nothing
            
            'SERVICE BILL
            Set rstTRANX = New ADODB.Recordset
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV'  ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            If Not (rstTRANX.EOF And rstTRANX.BOF) Then
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & rstTRANX2!SALES_TAX & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        Select Case RSTTRXFILE!SLSM_CODE
                            Case "P"
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                            Case Else
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                        End Select
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                        TAX_AMT = TAX_AMT + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
    
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                        'If Not (IsNull(RSTtax!PUR_TAX)) Then
                            PUR_TAX = PUR_TAX + (IIf(IsNull(RSTtax!ITEM_COST), 0, RSTtax!ITEM_COST) * IIf(IsNull(RSTtax!PUR_TAX) Or RSTtax!PUR_TAX = 0, RSTtax!SALES_TAX, RSTtax!PUR_TAX) / 100) * RSTtax!QTY
                        'End If
                        If Not (IsNull(RSTtax!CESS_PER)) Then
                            PUR_CESS = PUR_CESS + (IIf(IsNull(RSTtax!ITEM_COST), 0, RSTtax!ITEM_COST) * RSTtax!CESS_PER / 100) * RSTtax!QTY
                        End If
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    'GRDPranx.TextMatrix(M, N) = Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.Tag) * Val(RSTtax!QTY)
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
            End If
            rstTRANX.Close
            Set rstTRANX = Nothing
            
            rstTRANX2.MoveNext
        Loop
        rstTRANX2.Close
        Set rstTRANX2 = Nothing
    End If
    lbltaxamt.Caption = Format(TAX_AMT + CESSAMT + CESSPER, "0.00")
    lblpurtax.Caption = Format(PUR_TAX + CESSAMT + PUR_CESS, "0.00")
    LBLdifftax.Caption = Format(Val(lbltaxamt.Caption) - Val(lblpurtax.Caption), "0.00")
    lblcess.Caption = Format(KFC, "0.00")
    'LBLPROFIT.Caption = Val(LBLPROFIT.Caption) - ((Val(lbltaxamt.Caption) - Val(lblpurtax.Caption)) + Val(lblcess.Caption))
    If MDIMAIN.lblgst.Caption = "C" Then
        LBLPROFIT.Caption = Format(Round(Val(LBLNET.Caption) - (Val(lblCommi.Caption) + Val(LBLCOST.Caption)), 2), "0.00")
    ElseIf MDIMAIN.lblgst.Caption = "N" Then
        LBLPROFIT.Caption = Format(Round(Val(LBLNET.Caption) - (Val(lblCommi.Caption) + Val(LBLCOST.Caption)), 2), "0.00")
    Else
        LBLPROFIT.Caption = (Val(LBLBTRXTOTAL.Caption) - (Val(LBLDISCOUNT.Caption) + Val(LBLCOST.Caption) + Val(lblpurtax.Caption))) - (Val(LBLdifftax.Caption) + Val(lblcess.Caption))
    End If
    'LBLPROFIT.Caption = Format(Round(Val(LBLPROFIT.Caption) + Val(lblxchange.Caption), 2), "0.00")
    LBLPROFIT.Caption = Format(Round(Val(LBLPROFIT.Caption), 2), "0.00")
    
    'vbalProgressBar1.ShowText = False
    'vbalProgressBar1.Value = 0
    If LBLPETTY.Visible = True Then
        LBLWSALE.Caption = Format(Val(LBLBTRXTOTAL.Caption), "0.00")
    Else
        LBLWSALE.Caption = Format(Val(LBLBTRXTOTAL.Caption) - Val(LBLPETTY.Caption), "0.00")
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If ACT_FLAG = False Then ACT_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
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
            CmdDisplay.SetFocus
        Case vbKeyEscape
            DTFROM.SetFocus
    End Select
End Sub

Private Sub lblcloscash_DblClick()
    If LBLPETTYSALE(4).Visible = True Then
        LBLPETTYSALE(4).Visible = False
        LBLPETTY.Visible = False
        Label1(7).Visible = False
        LblPettyPur.Visible = False
        CmdCunterSales.Visible = False
        LBLWSALE.Caption = Format(Val(LBLBTRXTOTAL.Caption) - Val(LBLPETTY.Caption), "0.00")
        LBLWCLOCASH.Caption = Format(Round(Val(lblcloscash.Caption) - Val(LBLPETTY.Caption), 2), "0.00")
    Else
        LBLPETTYSALE(4).Visible = True
        LBLPETTY.Visible = True
        Label1(7).Visible = True
        LblPettyPur.Visible = True
        CmdCunterSales.Visible = True
        LBLWSALE.Caption = Format(Val(LBLBTRXTOTAL.Caption), "0.00")
        LBLWCLOCASH.Caption = Format(lblcloscash.Caption, "0.00")
    End If
End Sub

Private Sub LBLWCLOCASH_DblClick()
    On Error GoTo ERRHAND
    Dim rstTRANX As ADODB.Recordset
    If LBLPETTYSALE(4).Visible = True Then
        LBLPETTYSALE(4).Visible = False
        LBLPETTY.Visible = False
        Label1(7).Visible = False
        LblPettyPur.Visible = False
        CmdCunterSales.Visible = False
        LBLWSALE.Caption = Format(Val(LBLBTRXTOTAL.Caption) - Val(LBLPETTY.Caption), "0.00")
        LBLWCLOCASH.Caption = Format(Round(Val(lblcloscash.Caption) - Val(LBLPETTY.Caption), 2), "0.00")
        lblsaleret.Caption = ""
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select SUM(VCH_AMOUNT) from RETURNMAST  WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            lblsaleret.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        LblPurchret.Caption = ""
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select SUM(VCH_AMOUNT) from PURCAHSERETURN  WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='PR' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            LblPurchret.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
    Else
        LBLPETTYSALE(4).Visible = True
        LBLPETTY.Visible = True
        Label1(7).Visible = True
        LblPettyPur.Visible = True
        CmdCunterSales.Visible = True
        LBLWSALE.Caption = Format(Val(LBLBTRXTOTAL.Caption), "0.00")
        LBLWCLOCASH.Caption = Format(lblcloscash.Caption, "0.00")
        lblsaleret.Caption = ""
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select SUM(VCH_AMOUNT) from RETURNMAST  WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR' OR TRX_TYPE='RW')", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            lblsaleret.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        LblPurchret.Caption = ""
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select SUM(VCH_AMOUNT) from PURCAHSERETURN  WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PR' OR TRX_TYPE='WP') ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            LblPurchret.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
    End If
    Exit Sub
ERRHAND:
End Sub
