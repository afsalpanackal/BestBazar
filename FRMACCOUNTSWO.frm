VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FRMBILLPRINT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES REPORT"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMACCOUNTSWO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   18660
   Begin VB.CommandButton CmdCunterSales 
      Caption         =   "Couter wise Sales Report"
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
      Left            =   6570
      TabIndex        =   116
      Top             =   8205
      Width           =   1275
   End
   Begin VB.Frame FRMEBILL 
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H00000080&
      Height          =   4725
      Left            =   60
      TabIndex        =   8
      Top             =   1950
      Visible         =   0   'False
      Width           =   10845
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4140
         Left            =   30
         TabIndex        =   9
         Top             =   540
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   7303
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      Caption         =   "Frame1"
      Height          =   10320
      Left            =   -120
      TabIndex        =   0
      Top             =   -285
      Width           =   18720
      Begin VB.CheckBox ChkProfit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Profit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9300
         TabIndex        =   119
         Top             =   8970
         Width           =   1245
      End
      Begin VB.CommandButton CmdCounterReg 
         Caption         =   "Couter Register"
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
         Left            =   5400
         TabIndex        =   117
         Top             =   8040
         Width           =   1275
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export to Excel"
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
         Left            =   5400
         TabIndex        =   108
         Top             =   8505
         Width           =   1260
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   17535
         TabIndex        =   107
         Top             =   7545
         Width           =   1095
      End
      Begin VB.CommandButton CmdZeroBills 
         Caption         =   "Cancelled Bills"
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
         Left            =   17550
         TabIndex        =   106
         Top             =   8010
         Width           =   1080
      End
      Begin VB.CheckBox chkunbill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Petty Bills"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   14265
         TabIndex        =   104
         Top             =   8280
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.CommandButton Cmddaywise 
         Caption         =   "Day Wise"
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
         Left            =   14265
         TabIndex        =   95
         Top             =   8475
         Width           =   1036
      End
      Begin VB.CommandButton Cmdyear 
         Caption         =   "Year Wise"
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
         Left            =   16455
         TabIndex        =   97
         Top             =   8475
         Width           =   1036
      End
      Begin VB.CommandButton CmdMonthWise 
         Caption         =   "Month Wise"
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
         Left            =   15360
         TabIndex        =   96
         Top             =   8475
         Width           =   1036
      End
      Begin VB.CommandButton CmdUserrep 
         Caption         =   "Print User wise Report"
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
         Left            =   11985
         TabIndex        =   88
         Top             =   8490
         Width           =   1275
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Print Category wise Report"
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
         Left            =   10695
         TabIndex        =   87
         Top             =   8490
         Width           =   1260
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Print Itemwise Details"
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
         Left            =   13305
         TabIndex        =   86
         Top             =   7560
         Width           =   1200
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Customer wise Sale Analysis"
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
         Left            =   10695
         TabIndex        =   84
         Top             =   8040
         Width           =   1260
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print Area Wise Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10695
         TabIndex        =   83
         Top             =   7575
         Width           =   1260
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Print Cash / Credit Report"
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
         Left            =   12000
         TabIndex        =   74
         Top             =   8040
         Width           =   1260
      End
      Begin VB.CommandButton Cmdday 
         Caption         =   "Damage Analysis Report"
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
         Left            =   9360
         TabIndex        =   69
         Top             =   8040
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Bill Wise Report"
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
         Left            =   7995
         TabIndex        =   68
         Top             =   8040
         Width           =   1335
      End
      Begin VB.CommandButton CMDEfile 
         Caption         =   "Print Report for e-Filing"
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
         Left            =   9330
         TabIndex        =   53
         Top             =   8490
         Width           =   1335
      End
      Begin VB.CommandButton CmdReport 
         Caption         =   "Print Report"
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
         Left            =   6705
         TabIndex        =   52
         Top             =   8040
         Width           =   1275
      End
      Begin VB.CommandButton CmdMonthly 
         Caption         =   "Monthly wise Chart"
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
         Left            =   15585
         TabIndex        =   46
         Top             =   7545
         Width           =   960
      End
      Begin VB.CommandButton CmdDaily 
         Caption         =   "Daily wise Chart"
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
         Left            =   14520
         TabIndex        =   45
         Top             =   7545
         Width           =   1035
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0C0FF&
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
         Height          =   1965
         Left            =   150
         TabIndex        =   30
         Top             =   195
         Width           =   18510
         Begin VB.ListBox LstCategory 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1560
            Left            =   16095
            Style           =   1  'Checkbox
            TabIndex        =   118
            Top             =   360
            Width           =   2370
         End
         Begin VB.TextBox TXTDEALER5 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   12315
            TabIndex        =   111
            Top             =   825
            Width           =   1530
         End
         Begin VB.TextBox TXTREFNO 
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
            Left            =   9090
            TabIndex        =   105
            Top             =   1305
            Width           =   1185
         End
         Begin VB.CheckBox chkverify 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Verify Each Accounts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   13905
            TabIndex        =   93
            Top             =   150
            Width           =   1815
         End
         Begin VB.TextBox txtPhone 
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
            Left            =   6855
            TabIndex        =   90
            Top             =   930
            Width           =   2880
         End
         Begin VB.CommandButton cmdwoprint 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   30
            MaskColor       =   &H00C0C0FF&
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   1530
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox TXTDEALER4 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   13875
            TabIndex        =   78
            Top             =   825
            Width           =   2190
         End
         Begin VB.TextBox TXTDEALER2 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   10290
            TabIndex        =   56
            Top             =   825
            Width           =   1995
         End
         Begin VB.TextBox TxtAgent 
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
            Left            =   5895
            TabIndex        =   48
            Top             =   1305
            Width           =   2460
         End
         Begin VB.TextBox txtCustomercode 
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
            Left            =   6855
            TabIndex        =   47
            Top             =   165
            Width           =   3405
         End
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0C0FF&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   75
            TabIndex        =   33
            Top             =   420
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton OPTCUSTOMER 
            BackColor       =   &H00C0C0FF&
            Caption         =   "CUSTOMER"
            Height          =   210
            Left            =   90
            TabIndex        =   32
            Top             =   930
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
            Left            =   1470
            TabIndex        =   31
            Top             =   900
            Width           =   3720
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1680
            TabIndex        =   34
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   51511297
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   3630
            TabIndex        =   35
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   51511297
            CurrentDate     =   40498
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1470
            TabIndex        =   36
            Top             =   1245
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
         Begin MSDataListLib.DataList DataList1 
            Height          =   780
            Left            =   10290
            TabIndex        =   57
            Top             =   1140
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   1376
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0FF&
            Height          =   615
            Left            =   10290
            TabIndex        =   41
            Top             =   15
            Width           =   4950
            Begin VB.OptionButton Optservice 
               BackColor       =   &H00C0C0FF&
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
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   1470
               TabIndex        =   85
               Top             =   135
               Width           =   1530
            End
            Begin VB.OptionButton Optall 
               BackColor       =   &H00C0C0FF&
               Caption         =   "All"
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
               Height          =   240
               Left            =   3060
               TabIndex        =   77
               Top             =   135
               Width           =   645
            End
            Begin VB.OptionButton OPTGST 
               BackColor       =   &H00C0C0FF&
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
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   60
               TabIndex        =   76
               Top             =   135
               Width           =   1350
            End
            Begin VB.OptionButton Optpetty 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Petty"
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
               Height          =   240
               Left            =   1470
               TabIndex        =   75
               Top             =   360
               Width           =   945
            End
            Begin VB.OptionButton OptRT 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Stock Transfer"
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
               Height          =   240
               Left            =   3075
               TabIndex        =   44
               Top             =   360
               Width           =   1740
            End
            Begin VB.OptionButton OptWS 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Old 8"
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
               Height          =   240
               Left            =   7230
               TabIndex        =   43
               Top             =   345
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.OptionButton OptVan 
               BackColor       =   &H00C0C0FF&
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
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   75
               TabIndex        =   42
               Top             =   360
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin MSDataListLib.DataList DataList4 
            Height          =   780
            Left            =   13875
            TabIndex        =   79
            Top             =   1140
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   1376
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
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
            Height          =   240
            Left            =   5205
            TabIndex        =   102
            Tag             =   "5"
            Top             =   1665
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   423
            Picture         =   "FRMACCOUNTSWO.frx":030A
            ForeColor       =   0
            BarPicture      =   "FRMACCOUNTSWO.frx":0326
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
         Begin MSDataListLib.DataList DataList5 
            Height          =   780
            Left            =   12315
            TabIndex        =   112
            Top             =   1140
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   1376
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
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
         Begin VB.Label lbldealer5 
            Height          =   315
            Left            =   0
            TabIndex        =   114
            Top             =   1470
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent"
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
            Index           =   23
            Left            =   12345
            TabIndex        =   113
            Top             =   600
            Width           =   1350
         End
         Begin MSForms.ComboBox txtCustomerName 
            Height          =   345
            Left            =   6855
            TabIndex        =   92
            Top             =   540
            Width           =   3405
            VariousPropertyBits=   746604571
            ForeColor       =   255
            MaxLength       =   30
            DisplayStyle    =   3
            Size            =   "6006;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            DropButtonStyle =   0
            BorderColor     =   255
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Phone"
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
            Index           =   20
            Left            =   5235
            TabIndex        =   91
            Top             =   990
            Width           =   1635
         End
         Begin VB.Label lbldealer4 
            Height          =   315
            Left            =   -540
            TabIndex        =   82
            Top             =   1605
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange4 
            Height          =   315
            Left            =   0
            TabIndex        =   81
            Top             =   1065
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Area"
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
            Index           =   19
            Left            =   13905
            TabIndex        =   80
            Top             =   600
            Width           =   1665
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Category Name"
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
            Index           =   18
            Left            =   16065
            TabIndex        =   65
            Top             =   120
            Width           =   1350
         End
         Begin VB.Label LBLDEALER2 
            Height          =   315
            Left            =   0
            TabIndex        =   60
            Top             =   -345
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label FLAGCHANGE2 
            Height          =   315
            Left            =   0
            TabIndex        =   59
            Top             =   -360
            Width           =   495
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name"
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
            Index           =   15
            Left            =   10305
            TabIndex        =   58
            Top             =   600
            Width           =   1710
         End
         Begin VB.Label lblbillnos 
            Height          =   735
            Left            =   18585
            TabIndex        =   54
            Top             =   1110
            Width           =   1695
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Name"
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
            Index           =   14
            Left            =   5235
            TabIndex        =   51
            Top             =   1350
            Width           =   1155
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
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
            Index           =   13
            Left            =   5235
            TabIndex        =   50
            Top             =   585
            Width           =   1635
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Code"
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
            Index           =   11
            Left            =   5235
            TabIndex        =   49
            Top             =   210
            Width           =   1710
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
            TabIndex        =   40
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
            Left            =   3255
            TabIndex        =   39
            Top             =   405
            Width           =   285
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   6465
            TabIndex        =   38
            Top             =   1965
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   8685
            TabIndex        =   37
            Top             =   1905
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdview 
         Caption         =   "Print Agent wise Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7995
         TabIndex        =   27
         Top             =   7560
         Width           =   1320
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
         Height          =   420
         Left            =   7980
         TabIndex        =   26
         Top             =   8505
         Width           =   1350
      End
      Begin VB.CommandButton CmdCompany 
         Caption         =   "Print Company Wise Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9360
         TabIndex        =   3
         Top             =   7560
         Width           =   1320
      End
      Begin VB.CommandButton CMDPRINTREGISTER 
         Caption         =   "Print Item Wise Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6720
         TabIndex        =   4
         Top             =   7560
         Width           =   1260
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
         Height          =   435
         Left            =   16560
         TabIndex        =   2
         Top             =   7545
         Width           =   945
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
         Height          =   420
         Left            =   12000
         TabIndex        =   1
         Top             =   7575
         Width           =   1260
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   5325
         Left            =   165
         TabIndex        =   55
         Top             =   2160
         Width           =   18465
         _ExtentX        =   32570
         _ExtentY        =   9393
         _Version        =   393216
         Rows            =   1
         Cols            =   23
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton CmdPrintBills 
         Caption         =   "Print Bills"
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
         Left            =   13305
         TabIndex        =   94
         Top             =   8475
         Width           =   930
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0FF&
         Height          =   390
         Left            =   13305
         TabIndex        =   70
         Top             =   7890
         Width           =   4200
         Begin VB.OptionButton OptCredit 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Credit"
            ForeColor       =   &H00004000&
            Height          =   210
            Left            =   1350
            TabIndex        =   73
            Top             =   135
            Width           =   1110
         End
         Begin VB.OptionButton OptCash 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Cash"
            ForeColor       =   &H00004000&
            Height          =   210
            Left            =   2760
            TabIndex        =   72
            Top             =   135
            Width           =   1185
         End
         Begin VB.OptionButton OptBoth 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Both"
            ForeColor       =   &H00004000&
            Height          =   210
            Left            =   120
            TabIndex        =   71
            Top             =   135
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2460
         Left            =   120
         TabIndex        =   5
         Top             =   7470
         Width           =   8430
         Begin VB.Label label 
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
            Height          =   315
            Index           =   23
            Left            =   45
            TabIndex        =   110
            Top             =   1725
            Width           =   1245
         End
         Begin VB.Label LBLxchange 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   1395
            TabIndex        =   109
            Top             =   1680
            Width           =   1320
         End
         Begin VB.Label lblcess 
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
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   3765
            TabIndex        =   101
            Top             =   1710
            Visible         =   0   'False
            Width           =   1230
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
            Index           =   22
            Left            =   2745
            TabIndex        =   100
            Top             =   1725
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lbltaxamt 
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
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   5040
            TabIndex        =   99
            Top             =   195
            Width           =   1530
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot. Tax Collected"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Index           =   21
            Left            =   5040
            TabIndex        =   98
            Top             =   0
            Width           =   1665
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Handling"
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
            Index           =   17
            Left            =   2790
            TabIndex        =   64
            Top             =   1350
            Width           =   1200
         End
         Begin VB.Label lblhandle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   3675
            TabIndex        =   63
            Top             =   1320
            Width           =   1320
         End
         Begin VB.Label lblfrieght 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   1395
            TabIndex        =   62
            Top             =   1290
            Width           =   1320
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Frieght"
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
            Index           =   16
            Left            =   45
            TabIndex        =   61
            Top             =   1335
            Width           =   1245
         End
         Begin VB.Label LBLTOTAL 
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
            Height          =   315
            Index           =   12
            Left            =   45
            TabIndex        =   29
            Top             =   945
            Width           =   1245
         End
         Begin VB.Label lblcommi 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   1395
            TabIndex        =   28
            Top             =   900
            Width           =   1320
         End
         Begin VB.Label LBLNET 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   3675
            TabIndex        =   25
            Top             =   930
            Width           =   1320
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
            Height          =   495
            Index           =   10
            Left            =   2790
            TabIndex        =   24
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label LBLDISCOUNT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   1395
            TabIndex        =   23
            Top             =   465
            Width           =   1320
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "DISCOUNT"
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
            Left            =   45
            TabIndex        =   22
            Top             =   510
            Width           =   1155
         End
         Begin VB.Label LBLPROFIT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   3675
            TabIndex        =   21
            Top             =   495
            Width           =   1320
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
            Height          =   315
            Index           =   8
            Left            =   2775
            TabIndex        =   20
            Top             =   510
            Width           =   810
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
            Left            =   2775
            TabIndex        =   19
            Top             =   60
            Width           =   660
         End
         Begin VB.Label LBLCOST 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   3675
            TabIndex        =   18
            Top             =   45
            Width           =   1320
         End
         Begin VB.Label LBLTRXTOTAL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   1400
            TabIndex        =   7
            Top             =   45
            Width           =   1320
         End
         Begin VB.Label LBLTOTAL 
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   3
            Left            =   45
            TabIndex        =   6
            Top             =   105
            Width           =   1365
         End
      End
      Begin MSForms.TabStrip TabStrip1 
         Height          =   30
         Left            =   15105
         TabIndex        =   103
         Top             =   8790
         Width           =   30
         ListIndex       =   0
         Size            =   "53;53"
         Items           =   "Tab1;Tab2;"
         TipStrings      =   ";;"
         Names           =   "Tab1;Tab2;"
         NewVersion      =   -1  'True
         TabsAllocated   =   2
         Tags            =   ";;"
         TabData         =   2
         Accelerator     =   ";;"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
         TabState        =   "3;3"
      End
   End
   Begin VB.Label flagchange5 
      Height          =   315
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LBLDEALER3 
      Height          =   315
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label flagchange3 
      Height          =   315
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "FRMBILLPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY_REC As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG, PHY_FLAG As Boolean
Dim CAT_REC As New ADODB.Recordset
Dim CAT_FLAG As Boolean
Dim AGNT_REC As New ADODB.Recordset
Dim AGNT_FLAG As Boolean
Dim AREA_REC As New ADODB.Recordset
Dim AREA_FLAG As Boolean

Private Sub CmdCompany_Click()
    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Long
    
    On Error GoTo ERRHAND
'    If DataList1.BoundText = "" Then
'        MsgBox "Please select the Company from the list", vbOKOnly, "Company wise Report"
'        Exit Sub
'    End If
    db.Execute "Update itemmast set manufacturer ='' where isnull(manufacturer)"
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select the Customer from the list", vbOKOnly, "Company wise Report"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    If ChkProfit.Value = 1 Then
        ReportNameVar = Rptpath & "RPTCompRepPr"
    Else
        ReportNameVar = Rptpath & "RPTCompRep"
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If DataList1.BoundText = "" Then
        If OPTCUSTOMER.Value = False Then
            Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    Else
        If OPTCUSTOMER.Value = False Then
            Report.RecordSelectionFormula = "({itemmast.manufacturer}='" & DataList1.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {itemmast.manufacturer}='" & DataList1.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    End If
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdCounterReg_Click()
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    FRMCounterReg.Show
    FRMCounterReg.SetFocus
End Sub

Private Sub CmdCunterSales_Click()
    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
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
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            End If
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    If frmLogin.rs!Level = "5" Then
        Report.RecordSelectionFormula = "({TRXMAST.SYS_NAME}= '" & system_name & "' AND ({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "COUNTER WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CmdDaily_Click()
    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    
    ReportNameVar = Rptpath & "RptDaywise"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='WO') )"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
    Next
    frmreport.Caption = "DAILY SALES ANALYSIS"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Cmdday_Click()
   
    Dim i As Long
    
    On Error GoTo ERRHAND
'    If DataList1.BoundText = "" Then
'        MsgBox "Please select the Company from the list", vbOKOnly, "Company wise Report"
'        Exit Sub
'    End If
    
    db.Execute "Update trxfile set m_user_id ='130000' where isnull(m_user_id) or m_user_id =''"
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select the Customer from the list", vbOKOnly, "Damage Report"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    ReportNameVar = Rptpath & "RPTDamRept"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OPTCUSTOMER.Value = False Then
        Report.RecordSelectionFormula = "({TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "({CUSTMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    End If
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXFILE" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
    Next
    frmreport.Caption = "DAMAGE REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Cmddaywise_Click()
    Call Report_Generate("D")
End Sub

Private Sub CmDDisplay_Click()
    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
    db.Execute "delete From SALESREG"
    
    GRDTranx.TextMatrix(0, 12) = "AREA"
    GRDTranx.ColWidth(12) = 1200
    LBLTRXTOTAL.Caption = "0.00"
    LBLDISCOUNT.Caption = "0.00"
    LBLNET.Caption = "0.00"
    LBLCOST.Caption = "0.00"
    LBLPROFIT.Caption = "0.00"
    lblcommi.Caption = "0.00"
    lblFrieght.Caption = ""
    lblhandle.Caption = ""
    lbltaxamt.Caption = ""
    lblcess.Caption = ""
    'GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    
    On Error GoTo ERRHAND
    'BILL_NAME LIKE '%" & txtCustomerName.Text & "%' AND
    Screen.MousePointer = vbHourglass
    
    If chkverify.Value = 0 Then GoTo SKIP_VERIFY
    
    n = 1
    M = 0
    Dim rstdbt As ADODB.Recordset
    Dim rstdbt2 As ADODB.Recordset
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT DISTINCT ACT_CODE From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO') ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        Set rstdbt = New ADODB.Recordset
        rstdbt.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "' and TRX_TYPE = 'DR'", db, adOpenStatic, adLockOptimistic, adCmdText
        rstdbt.Properties("Update Criteria").Value = adCriteriaKey
        Do Until rstdbt.EOF
                
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(RCPT_AMOUNT) from trnxrcpt WHERE ACT_CODE = '" & rstdbt!ACT_CODE & "' AND INV_NO  = " & rstdbt!INV_NO & " AND INV_TRX_TYPE = '" & rstdbt!INV_TRX_TYPE & "' AND INV_TRX_YEAR = '" & rstdbt!TRX_YEAR & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
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
        
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
SKIP_VERIFY:
    
'    Dim searchstring, searchstring2 As String
'    Dim selcat As Boolean
'    searchstring2 = ""
'    searchstring = ""
'    selcat = False
'    For n = 0 To LstCategory.ListCount - 1
'        If LstCategory.Selected(n) = True Then
'            searchstring = searchstring & " CATEGORY LIKE '%" & LstCategory.List(n) & "%'" & " OR "
'            selcat = True
'        End If
'    Next n
'    If Len(searchstring) > 4 Then
'        searchstring = Left(searchstring, Len(searchstring) - 4)
'        searchstring = "(" & searchstring & ")"
'    End If
    
    n = 1
    M = 0
    Dim TOT_AMT As Double
    Dim DISC_AMT, RETAMT, FRIEGHT, HANDLE As Double
    Dim TRXMAST As ADODB.Recordset
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or  TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO' OR TRX_TYPE='VI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    rstTRANX.Properties("Update Criteria").Value = adCriteriaKey
    Do Until rstTRANX.EOF
        TOT_AMT = 0
        DISC_AMT = 0
        RETAMT = 0
        FRIEGHT = 0
        HANDLE = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT SUM(TRX_TOTAL) FROM TRXFILE WHERE VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            TOT_AMT = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        If rstTRANX!SLSM_CODE = "A" Then
            DISC_AMT = IIf(IsNull(rstTRANX!DISCOUNT), "", rstTRANX!DISCOUNT)
        ElseIf rstTRANX!SLSM_CODE = "P" Then
            If IsNull(rstTRANX!VCH_AMOUNT) Or rstTRANX!VCH_AMOUNT = 0 Then
                DISC_AMT = 0
            Else
                'DISC_AMT = IIf(IsNull(rstTRANX!DISCOUNT), 0, Round((rstTRANX!DISCOUNT * 100 / rstTRANX!VCH_AMOUNT), 2))
                DISC_AMT = IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            End If
        Else
            DISC_AMT = IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
        End If
        RETAMT = IIf(IsNull(rstTRANX!ADD_AMOUNT), 0, rstTRANX!ADD_AMOUNT)
        FRIEGHT = IIf(IsNull(rstTRANX!FRIEGHT), 0, rstTRANX!FRIEGHT)
        HANDLE = IIf(IsNull(rstTRANX!HANDLE), 0, rstTRANX!HANDLE)
        'rstTRANX!VCH_AMOUNT = Round(TOT_AMT, 0)
        'rstTRANX!NET_AMOUNT = Round((TOT_AMT + HANDLE + FRIEGHT) - (RETAMT + DISC_AMT), 0)
        rstTRANX!VCH_AMOUNT = Format(TOT_AMT, "0")
        rstTRANX!NET_AMOUNT = Format((TOT_AMT + HANDLE + FRIEGHT) - (RETAMT + DISC_AMT), "0")
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    db.Execute "Update TRXMAST Set PHONE='' where isnull(PHONE)"
    db.Execute "Update TRXMAST Set BILL_NAME='Cash' where isnull(BILL_NAME)"
    db.Execute "Update TRXMAST Set ACT_CODE='130001' where isnull(ACT_CODE)"
    db.Execute "Update TRXMAST Set ACT_NAME='Cash' where isnull(ACT_NAME)"
    
    Set rstTRANX = New ADODB.Recordset
    If frmLogin.rs!Level = "5" Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE C_USER_ID = '" & frmLogin.rs!USER_ID & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or  TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO' OR TRX_TYPE='VI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        If OPTPERIOD.Value = True Then
            If OPTGST.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptVan.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
                'rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
                'rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' ", db, adOpenStatic, adLockReadOnly
            ElseIf Optservice.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptPetty.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='WO')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptRT.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='TF')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptWs.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            Else
                If OptPetty.Visible = False Then
                    rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='VI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
                Else
                    rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO' OR TRX_TYPE='VI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
                End If
            End If
        Else
            If OPTGST.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptVan.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf Optservice.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptPetty.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='WO'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptRT.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='TF'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptWs.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SI'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            Else
                If OptPetty.Visible = False Then
                    rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or  TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='VI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
                Else
                    rstTRANX.Open "SELECT * From TRXMAST WHERE PHONE Like '%" & Trim(TxtPhone.Text) & "%' AND (BILL_NAME Like '%" & Trim(txtCustomerName.Text) & "%' OR ACT_NAME Like '%" & Trim(txtCustomerName.Text) & "%') AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or  TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO' OR TRX_TYPE='VI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
                End If
            End If
        End If
    End If
    lblbillnos = ""
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
        rstTRANX.MoveLast
        lblbillnos.Caption = rstTRANX!VCH_NO
        rstTRANX.MoveFirst
        lblbillnos.Caption = "From : " & rstTRANX!VCH_NO & " to " & lblbillnos.Caption
        
    Else
        vbalProgressBar1.Max = 100
    End If
    
    Dim Tax_collected As Double
    Dim TAXABLE_AMT As Double
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTSALEREG.Properties("Update Criteria").Value = adCriteriaKey
    Do Until rstTRANX.EOF
        If GRDTranx.rows >= 14250 Then
            GRDTranx.rows = 2
            'GRDTranx.Clear
            M = 1
        End If
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = n
        GRDTranx.TextMatrix(M, 1) = rstTRANX!TRX_TYPE
        Select Case rstTRANX!TRX_TYPE
            Case "GI"
                GRDTranx.TextMatrix(M, 2) = "GI"
            Case "SV"
                GRDTranx.TextMatrix(M, 2) = "SV"
            Case "RI"
                GRDTranx.TextMatrix(M, 2) = "RT"
            Case "TF"
                GRDTranx.TextMatrix(M, 2) = "ST"
            Case "WO"
                GRDTranx.TextMatrix(M, 2) = "PT"
            Case "SI"
                GRDTranx.TextMatrix(M, 2) = "WS"
            Case "VI"
                GRDTranx.TextMatrix(M, 2) = "VN"
            Case "HI"
                GRDTranx.TextMatrix(M, 2) = "GR"
        End Select
        GRDTranx.TextMatrix(M, 3) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 4) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 5) = Format(Round(rstTRANX!VCH_AMOUNT, 2), "0.00")
'        If rstTRANX!SLSM_CODE = "A" Then
'
'        ElseIf rstTRANX!SLSM_CODE = "P" Then
'            GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(Round((rstTRANX!DISCOUNT * 100 / rstTRANX!VCH_AMOUNT), 2), "0.00"))
'        End If
        GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(rstTRANX!DISCOUNT, "0.00"))
        GRDTranx.TextMatrix(M, 7) = Format(Round(rstTRANX!NET_AMOUNT, 2), "0.00") 'Format(Round(Val(GRDTranx.TextMatrix(M, 5)) - Val(GRDTranx.TextMatrix(M, 6)), 2), "0.00")
        GRDTranx.TextMatrix(M, 22) = Format(Round(rstTRANX!ADD_AMOUNT, 2), "0.00")
        
        
        CMDEXIT.Tag = IIf(IsNull(rstTRANX!DISCOUNT), "0", Format(rstTRANX!DISCOUNT, "0.00"))
        'GRDTranx.TextMatrix(M, 7) = Format(Round(Val(GRDTranx.TextMatrix(M, 5)), 2), "0.00")
        If frmLogin.rs!Level <> "0" Then
            GRDTranx.TextMatrix(M, 8) = "xxx"
            GRDTranx.TextMatrix(M, 9) = "xxx"
        Else
            GRDTranx.TextMatrix(M, 8) = IIf(IsNull(rstTRANX!COMM_AMT), "0", Format(rstTRANX!COMM_AMT, "0.00"))
            GRDTranx.TextMatrix(M, 9) = IIf(IsNull(rstTRANX!PAY_AMOUNT), "0", Format(rstTRANX!PAY_AMOUNT, "0.00"))
        End If
        GRDTranx.TextMatrix(M, 10) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        GRDTranx.TextMatrix(M, 11) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS), "", ", " & rstTRANX!BILL_ADDRESS)
        
        CMDDISPLAY.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
        Tax_collected = 0
        TAXABLE_AMT = 0
        'If rstTRANX!TRX_TYPE <> "SI" Then GoTo SKIP
        Set RSTTRXFILE = New ADODB.Recordset
        'RSTTRXFILE.Open "Select DISTINCT SALES_TAX From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND CATEGORY LIKE '%" & DataList3.BoundText & "%'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If selcat = True Then
'            searchstring2 = "Select DISTINCT SALES_TAX From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND " & searchstring
'            RSTTRXFILE.Open searchstring2, db, adOpenStatic, adLockReadOnly, adCmdText
'        Else
'            RSTTRXFILE.Open "Select DISTINCT SALES_TAX From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'        End If
        RSTTRXFILE.Open "Select DISTINCT SALES_TAX From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            EXSALEAMT = 0
            TAXSALEAMT = 0
            TaxAmt = 0
            MRPVALUE = 0
            DISCAMT = 0
            TAXRATE = RSTTRXFILE!SALES_TAX
            Set RSTtax = New ADODB.Recordset
            'RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & RSTTRXFILE!SALES_TAX & " AND CATEGORY LIKE '%" & DataList3.BoundText & "%'", db, adOpenStatic, adLockReadOnly, adCmdText
            RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & RSTTRXFILE!SALES_TAX & " ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                'PUR_TAX = PUR_TAX + (IIf(IsNull(RSTtax!ITEM_COST), 0, RSTtax!ITEM_COST) * IIf(IsNull(RSTtax!PUR_TAX) Or RSTtax!PUR_TAX = 0, RSTtax!SALES_TAX, RSTtax!PUR_TAX) / 100) * RSTtax!QTY
                If MDIMAIN.lblgst.Caption = "R" Then
                    'Tax_collected = Tax_collected + (IIf(IsNull(RSTtax!ITEM_COST), 0, RSTtax!ITEM_COST) * IIf(IsNull(RSTtax!PUR_TAX) Or RSTtax!PUR_TAX = 0, RSTtax!SALES_TAX, RSTtax!PUR_TAX) / 100) * RSTtax!QTY
                    
                    Select Case rstTRANX!SLSM_CODE
                            Case "P"
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                            Case Else
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                        End Select
                        'KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!KFC_TAX), 0, RSTtax!KFC_TAX / 100)) * RSTtax!QTY
                        Tax_collected = Tax_collected + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                End If
                If RSTTRXFILE!SALES_TAX > 0 Then 'And RSTtax!CHECK_FLAG = "V" Then
                    TAXSALEAMT = TAXSALEAMT + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                    'TAXAMT = TAXAMT + Round((RSTtax!PTR * RSTtax!SALES_TAX / 100) * RSTtax!QTY, 2)
                    TaxAmt = TaxAmt + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                Else
                    If RSTtax!SALE_1_FLAG = "1" Then
                        TaxAmt = TaxAmt + Round((RSTtax!SALES_PRICE - RSTtax!PTR) * RSTtax!QTY, 2)
                        MRPVALUE = Round(MRPVALUE + (100 * RSTtax!MRP / 105) * RSTtax!QTY, 2)
                    End If
                    EXSALEAMT = EXSALEAMT + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                    TAXSALEAMT = TAXSALEAMT + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                End If
                DISCAMT = Round(DISCAMT + IIf(IsNull(RSTtax!LINE_DISC), 0, RSTtax!TRX_TOTAL * RSTtax!LINE_DISC / 100), 2)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            RSTSALEREG.AddNew
            TAXSALEAMT = TAXSALEAMT - TaxAmt
            TAXABLE_AMT = TAXABLE_AMT + TAXSALEAMT
            RSTSALEREG!VCH_NO = rstTRANX!VCH_NO 'N
            RSTSALEREG!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTSALEREG!VCH_DATE = rstTRANX!VCH_DATE
            RSTSALEREG!DISCOUNT = DISCAMT
            RSTSALEREG!VCH_AMOUNT = Val(GRDTranx.TextMatrix(M, 7))
            RSTSALEREG!PAYAMOUNT = Val(GRDTranx.TextMatrix(M, 9))
            RSTSALEREG!ACT_NAME = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            RSTSALEREG!ACT_CODE = IIf(IsNull(rstTRANX!ACT_CODE), "", rstTRANX!ACT_CODE)
            RSTSALEREG!TIN_NO = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
'            Dim RSTACTCODE As ADODB.Recordset
'            Set RSTACTCODE = New ADODB.Recordset
'            RSTACTCODE.Open "SELECT KGST FROM CUSTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'            If Not (RSTACTCODE.EOF And RSTACTCODE.BOF) Then
'                RSTSALEREG!TIN_NO = RSTACTCODE!KGST
'            End If
'            RSTACTCODE.Close
'            Set RSTACTCODE = Nothing
            RSTSALEREG!EXMPSALES_AMT = EXSALEAMT
            RSTSALEREG!TAXSALES_AMT = TAXSALEAMT
            RSTSALEREG!TAXAMOUNT = TaxAmt
            RSTSALEREG!TAXRATE = TAXRATE
            CMDDISPLAY.Tag = Val(CMDDISPLAY.Tag) + EXSALEAMT
            FRMEMAIN.Tag = Val(FRMEMAIN.Tag) + TAXSALEAMT
            FRMEBILL.Tag = Val(FRMEBILL.Tag) + TaxAmt
            RSTSALEREG.Update
            
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        'GRDTranx.TextMatrix(M, 12) = Format(Val(CMDDISPLAY.Tag), "0.00")
        If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Or rstTRANX!TRX_TYPE = "WO" Then
            GRDTranx.TextMatrix(M, 20) = Format(Round(rstTRANX!NET_AMOUNT, 2), "0.00")
        Else
            'GRDTranx.TextMatrix(M, 20) = Format(Round(IIf(IsNull(rstTRANX!GROSS_AMT) Or rstTRANX!GROSS_AMT = 0, rstTRANX!NET_AMOUNT, rstTRANX!GROSS_AMT), 2), "0.00")
            GRDTranx.TextMatrix(M, 20) = Format(Val(TAXABLE_AMT), "0.00") 'Format(Round(Val(GRDTranx.TextMatrix(M, 7)) - Tax_collected, 2), "0.00")
        End If
        GRDTranx.TextMatrix(M, 12) = IIf(IsNull(rstTRANX!Area), "", rstTRANX!Area)
        GRDTranx.TextMatrix(M, 13) = Format(Val(FRMEMAIN.Tag), "0.00")
        GRDTranx.TextMatrix(M, 14) = Format(Val(FRMEBILL.Tag), "0.00")
        GRDTranx.TextMatrix(M, 15) = rstTRANX!TRX_YEAR
        GRDTranx.TextMatrix(M, 21) = Format(Tax_collected, "0.00")
        
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + Val(GRDTranx.TextMatrix(M, 5)), "0.00")
        LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption) + Val(GRDTranx.TextMatrix(M, 6)), "0.00")
        LBLNET.Caption = Format(Val(LBLNET.Caption) + Val(GRDTranx.TextMatrix(M, 7)), "0.00")
        
        lblFrieght.Caption = Format(Val(lblFrieght.Caption) + IIf(IsNull(rstTRANX!FRIEGHT), 0, rstTRANX!FRIEGHT), "0.00")
        lblhandle.Caption = Format(Val(lblhandle.Caption) + IIf(IsNull(rstTRANX!HANDLE), 0, rstTRANX!HANDLE), "0.00")
        If frmLogin.rs!Level <> "0" Then
            lblcommi.Caption = "xxx"
            LBLCOST.Caption = "xxx"
            LBLPROFIT.Caption = "xxx"
            GRDTranx.TextMatrix(M, 16) = "xxx"
            GRDTranx.TextMatrix(M, 17) = "xxx"
        Else
            lblcommi.Caption = Format(Val(lblcommi.Caption) + Val(GRDTranx.TextMatrix(M, 8)), "0.00")
            LBLCOST.Caption = Format(Val(LBLCOST.Caption) + Val(GRDTranx.TextMatrix(M, 9)), "0.00")
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Or rstTRANX!TRX_TYPE = "WO" Then
                GRDTranx.TextMatrix(M, 16) = Format(Round((Val(GRDTranx.TextMatrix(M, 20)) + Val(GRDTranx.TextMatrix(M, 22))) - Val(GRDTranx.TextMatrix(M, 9)), 2), "0.00")
            Else
                GRDTranx.TextMatrix(M, 16) = Format(Round(((Val(GRDTranx.TextMatrix(M, 20)) + Val(GRDTranx.TextMatrix(M, 22))) - (Val(GRDTranx.TextMatrix(M, 6)) + Val(GRDTranx.TextMatrix(M, 8)))) - Val(GRDTranx.TextMatrix(M, 9)), 2), "0.00")
            End If
            If Val(GRDTranx.TextMatrix(M, 9)) = 0 Then
                GRDTranx.TextMatrix(M, 17) = "0.00"
            'ElseIf Val(GRDTranx.TextMatrix(M, 5)) = 0 Then
                'GRDTranx.TextMatrix(M, 17) = "100.00"
            Else
                GRDTranx.TextMatrix(M, 17) = Format(Round(((((Val(GRDTranx.TextMatrix(M, 20)) + Val(GRDTranx.TextMatrix(M, 22))) - (Val(GRDTranx.TextMatrix(M, 6)) + Val(GRDTranx.TextMatrix(M, 8)))) * 100) / Val(GRDTranx.TextMatrix(M, 9))) - 100, 2), "0.00")
            End If
            'LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
        End If
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND TRX_TYPE = 'DR' AND INV_TRX_TYPE  = '" & rstTRANX!TRX_TYPE & "' AND INV_NO = '" & rstTRANX!VCH_NO & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "'", db, adOpenForwardOnly
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            GRDTranx.TextMatrix(M, 18) = IIf(IsNull(RSTTRXFILE!RCVD_AMOUNT), "", Format(RSTTRXFILE!RCVD_AMOUNT, "0.00"))
            GRDTranx.TextMatrix(M, 19) = Format(Round(Val(GRDTranx.TextMatrix(M, 7)) - Val(GRDTranx.TextMatrix(M, 18)), 2), "0.00")
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        If frmLogin.rs!Level = "0" Then
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
                lbltaxamt.Caption = ""
            Else
                lbltaxamt.Caption = Val(lbltaxamt.Caption) + Val(GRDTranx.TextMatrix(M, 21))
                LBLPROFIT.Caption = (Val(LBLNET.Caption) + Val(LBLxchange.Caption)) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption) + Val(lbltaxamt.Caption))
                
                'LBLPROFIT.Caption = Round(Val(LBLPROFIT.Caption) - Val(lblxchange.Caption), 2)
                'LBLPROFIT.Caption = Val(LBLPROFIT.Caption)
            End If
        End If
SKIP:
        GRDTranx.Refresh
        n = n + 1
        rstTRANX.MoveNext
    Loop
    
'    If frmLogin.rs!Level = "0" And MDIMAIN.lblgst.Caption = "R" Then
        LBLPROFIT.Caption = Format(Round(Val(LBLPROFIT.Caption), 2), "0.00")
'    End If
    
    If GRDTranx.rows > 14 Then GRDTranx.TopRow = GRDTranx.rows - 1
    rstTRANX.Close
    Set rstTRANX = Nothing
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    LBLxchange.Caption = ""
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(TRX_TOTAL) From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE='SV') ", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        LBLxchange.Caption = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    
    LBLPROFIT.Caption = Format(LBLPROFIT.Caption, "0.00")
'    If frmLogin.rs!Level = "0" Then
'        LBLPROFIT.Caption = "" 'Format(Val(LBLNET.Caption) - (Val(lblcost.Caption) + Val(lblCommi.Caption)), "0.00")
'    End If
        
    flagchange.Caption = ""
    flagchange2.Caption = ""
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    'GRDTranx.Visible = True
    
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

Private Sub CmdExport_Click()
    If GRDTranx.rows <= 1 Then Exit Sub
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "REPORT"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, n As Long
    
    On Error GoTo ERRHAND
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
        oWS.Range("A1", "J1").Merge
        oWS.Range("A1", "J1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "J2").Merge
        oWS.Range("A2", "J2").HorizontalAlignment = xlCenter
    'End If
    oWS.Range("A:A").ColumnWidth = 6
    oWS.Range("B:B").ColumnWidth = 10
    oWS.Range("C:C").ColumnWidth = 12
    oWS.Range("D:D").ColumnWidth = 12
    oWS.Range("E:E").ColumnWidth = 12
    oWS.Range("F:F").ColumnWidth = 12
    oWS.Range("G:G").ColumnWidth = 12
    oWS.Range("H:H").ColumnWidth = 12
    oWS.Range("I:I").ColumnWidth = 12
    oWS.Range("J:J").ColumnWidth = 12
    oWS.Range("K:K").ColumnWidth = 12
    oWS.Range("L:L").ColumnWidth = 12
    oWS.Range("M:M").ColumnWidth = 12
    oWS.Range("N:N").ColumnWidth = 12
    oWS.Range("O:O").ColumnWidth = 12
    oWS.Range("P:P").ColumnWidth = 12
'    oWS.Range("Q:Q").ColumnWidth = 12
'    oWS.Range("R:R").ColumnWidth = 12
'    oWS.Range("S:S").ColumnWidth = 12
'    oWS.Range("T:T").ColumnWidth = 12
'    oWS.Range("U:U").ColumnWidth = 12
'    oWS.Range("V:V").ColumnWidth = 12
    
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
    oWS.Range("A" & 2).Value = "SALES REPORT"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).Value = GRDTranx.TextMatrix(0, 0)
    oWS.Range("B" & 3).Value = GRDTranx.TextMatrix(0, 1)
    oWS.Range("C" & 3).Value = GRDTranx.TextMatrix(0, 2)
    oWS.Range("D" & 3).Value = GRDTranx.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 3).Value = GRDTranx.TextMatrix(0, 4)
    oWS.Range("F" & 3).Value = GRDTranx.TextMatrix(0, 5)
    oWS.Range("G" & 3).Value = GRDTranx.TextMatrix(0, 6)
    oWS.Range("H" & 3).Value = GRDTranx.TextMatrix(0, 7)
    oWS.Range("I" & 3).Value = GRDTranx.TextMatrix(0, 8)
    oWS.Range("J" & 3).Value = GRDTranx.TextMatrix(0, 9)
    oWS.Range("K" & 3).Value = GRDTranx.TextMatrix(0, 10)
    oWS.Range("L" & 3).Value = GRDTranx.TextMatrix(0, 11)
    oWS.Range("M" & 3).Value = GRDTranx.TextMatrix(0, 12)
    oWS.Range("0" & 3).Value = GRDTranx.TextMatrix(0, 13)
    oWS.Range("O" & 3).Value = GRDTranx.TextMatrix(0, 14)
    oWS.Range("P" & 3).Value = GRDTranx.TextMatrix(0, 15)
'    oWS.Range("Q" & 3).value = GRDTranx.TextMatrix(0, 16)
'    oWS.Range("R" & 3).value = GRDTranx.TextMatrix(0, 17)
'    oWS.Range("S" & 3).value = GRDTranx.TextMatrix(0, 18)
'    oWS.Range("T" & 3).value = GRDTranx.TextMatrix(0, 19)
    
    On Error GoTo ERRHAND
    
    i = 4
    For n = 1 To GRDTranx.rows - 1
        oWS.Range("A" & i).Value = GRDTranx.TextMatrix(n, 0)
        oWS.Range("B" & i).Value = GRDTranx.TextMatrix(n, 1)
        oWS.Range("C" & i).Value = GRDTranx.TextMatrix(n, 2)
        oWS.Range("D" & i).Value = GRDTranx.TextMatrix(n, 3)
        oWS.Range("E" & i).Value = GRDTranx.TextMatrix(n, 4)
        oWS.Range("F" & i).Value = GRDTranx.TextMatrix(n, 5)
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
'        oWS.Range("Q" & i).value = GRDTranx.TextMatrix(N, 16)
'        oWS.Range("R" & i).value = GRDTranx.TextMatrix(N, 17)
'        oWS.Range("S" & i).value = GRDTranx.TextMatrix(N, 18)
'        oWS.Range("T" & i).value = GRDTranx.TextMatrix(N, 19)
        On Error GoTo ERRHAND
        i = i + 1
    Next n
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.HorizontalAlignment = xlRight
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 13            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    
   
SKIP:
    oApp.Visible = True
    
'    If Sum_flag = True Then
'        'oWS.Columns("C:C").Select
'        oWS.Columns("C:C").NumberFormat = "0"
'        oWS.Columns("A:Z").EntireColumn.AutoFit
'    End If
'
'    Set oWB = Nothing
'    oApp.Quit
'    Set oApp = Nothing
'
    
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox err.Description
End Sub

Private Sub CmdMonthWise_Click()
    Call Report_Generate("M")
End Sub

Private Sub CmdPrintBills_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select Customer from the list", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    
    On Error GoTo ERRHAND
    
        Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompAddress4, CompAddress5, CompTin, CompCST, BIL_PRE, BILL_SUF, DL, ML, DL1, DL2, INV_TERMS, INV_MSG, BANK_DET, PAN_NO, OS_FLAG As String
    Dim QtnTerms, QtnTerms1, QtnTerms2, QtnTerms3, QtnTerms4, T2COPIES As String
    
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Val(GRDTranx.TextMatrix(GRDTranx.Row, 15)) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        CompName = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
        CompAddress1 = IIf(IsNull(RSTCOMPANY!Address), "", RSTCOMPANY!Address)
        CompAddress2 = IIf(IsNull(RSTCOMPANY!HO_NAME), "", RSTCOMPANY!HO_NAME)
        CompAddress5 = IIf(IsNull(RSTCOMPANY!TEL_NO) Or RSTCOMPANY!TEL_NO = "", "", "Ph: " & RSTCOMPANY!TEL_NO)
        CompAddress3 = IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", "Ph: " & RSTCOMPANY!FAX_NO)
        CompAddress4 = IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!EMAIL_ADD)
        CompTin = IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", "", "GSTIN No. " & RSTCOMPANY!CST)
        CompCST = IIf(IsNull(RSTCOMPANY!DL_NO) Or RSTCOMPANY!DL_NO = "", "", "CST No. " & RSTCOMPANY!DL_NO)
        DL = IIf(IsNull(RSTCOMPANY!DL_NO) Or RSTCOMPANY!DL_NO = "", "", "DL No. " & RSTCOMPANY!DL_NO)
        ML = IIf(IsNull(RSTCOMPANY!ML_NO) Or RSTCOMPANY!DL_NO = "", "", "ML No. " & RSTCOMPANY!ML_NO)
        If Trim(GRDTranx.TextMatrix(GRDTranx.Row, 1)) = "GI" Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        ElseIf Trim(GRDTranx.TextMatrix(GRDTranx.Row, 1)) = "GI" Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
        Else
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        End If
        'If Trim(TxtVehicle.text) = "" Then TxtVehicle.text = IIf(IsNull(RSTCOMPANY!VEHICLE), "", RSTCOMPANY!VEHICLE)
        INV_TERMS = IIf(IsNull(RSTCOMPANY!INV_TERMS) Or RSTCOMPANY!INV_TERMS = "", "", RSTCOMPANY!INV_TERMS)
        INV_MSG = IIf(IsNull(RSTCOMPANY!INV_MSGS) Or RSTCOMPANY!INV_MSGS = "", "", RSTCOMPANY!INV_MSGS)
        BANK_DET = IIf(IsNull(RSTCOMPANY!bank_details) Or RSTCOMPANY!bank_details = "", "", RSTCOMPANY!bank_details)
        PAN_NO = IIf(IsNull(RSTCOMPANY!PAN_NO) Or RSTCOMPANY!PAN_NO = "", "", RSTCOMPANY!PAN_NO)
'        T2COPIES = IIf(IsNull(RSTCOMPANY!T2_COPIES) Or RSTCOMPANY!T2_COPIES = "", "N", RSTCOMPANY!T2_COPIES)
'        If thermalprn = True Then
'            OS_FLAG = IIf(IsNull(RSTCOMPANY!OSPTY_FLAG) Or RSTCOMPANY!OSPTY_FLAG = "", "", RSTCOMPANY!OSPTY_FLAG)
'        Else
'            OS_FLAG = IIf(IsNull(RSTCOMPANY!OSB2C_FLAG) Or RSTCOMPANY!OSB2C_FLAG = "", "", RSTCOMPANY!OSB2C_FLAG)
'        End If
        If RSTCOMPANY!TERMS_FLAG = "Y" Then
            QtnTerms = "TERMS & CONDITIONS:"
            QtnTerms1 = IIf(IsNull(RSTCOMPANY!Terms1) Or RSTCOMPANY!Terms1 = "", "", RSTCOMPANY!Terms1)
            QtnTerms2 = IIf(IsNull(RSTCOMPANY!Terms2) Or RSTCOMPANY!Terms2 = "", "", RSTCOMPANY!Terms2)
            QtnTerms3 = IIf(IsNull(RSTCOMPANY!Terms3) Or RSTCOMPANY!Terms3 = "", "", RSTCOMPANY!Terms3)
            QtnTerms4 = IIf(IsNull(RSTCOMPANY!Terms4) Or RSTCOMPANY!Terms4 = "", "", RSTCOMPANY!Terms4)
        Else
            QtnTerms = ""
            QtnTerms1 = ""
            QtnTerms2 = ""
            QtnTerms3 = ""
            QtnTerms4 = ""
        End If
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing


    ReportNameVar = Rptpath & "RPTPRINTBILLS"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 3)) & " and ({TRXFILE.TRX_TYPE}= '" & Trim(GRDTranx.TextMatrix(GRDTranx.Row, 1)) & "' or {TRXFILE.TRX_YEAR}= '" & Val(GRDTranx.TextMatrix(GRDTranx.Row, 15)) & "'))"
        
    Set CRXFormulaFields = Report.FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.Text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
        If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.Text = "'" & CompName & "'"
        If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.Text = "'" & CompAddress1 & "'"
        If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.Text = "'" & CompAddress2 & "'"
        If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.Text = "'" & CompAddress3 & "'"
        If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.Text = "'" & CompAddress4 & "'"
        If CRXFormulaField.Name = "{@Comp_Address5}" Then CRXFormulaField.Text = "'" & CompAddress5 & "'"
        If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.Text = "'" & CompTin & "'"
        If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.Text = "'" & CompCST & "'"
        If CRXFormulaField.Name = "{@DL}" Then CRXFormulaField.Text = "'" & DL & "'"
        If CRXFormulaField.Name = "{@ML}" Then CRXFormulaField.Text = "'" & ML & "'"
        If CRXFormulaField.Name = "{@HSNSUM_FLAG}" Then
'            If Val(lblnetamount.Caption) >= Val(MDIMAIN.LBLHSNSUM.Caption) Or Trim(lblIGST.Caption) = "Y" Then
'                CRXFormulaField.text = "'Y'"
'            Else
                CRXFormulaField.Text = "'N'"
'            End If
        End If
        'If CRXFormulaField.Name = "{@salesman}" Then CRXFormulaField.text = "'" & CMBDISTI.text & "'"
        If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.Text = "'" & INV_TERMS & "'"
        If CRXFormulaField.Name = "{@inv_msg}" Then CRXFormulaField.Text = "'" & INV_MSG & "'"
        If CRXFormulaField.Name = "{@Terms}" Then CRXFormulaField.Text = "'" & QtnTerms & "'"
        If CRXFormulaField.Name = "{@Terms1}" Then CRXFormulaField.Text = "'" & QtnTerms1 & "'"
        If CRXFormulaField.Name = "{@Terms2}" Then CRXFormulaField.Text = "'" & QtnTerms2 & "'"
        If CRXFormulaField.Name = "{@Terms3}" Then CRXFormulaField.Text = "'" & QtnTerms3 & "'"
        If CRXFormulaField.Name = "{@Terms4}" Then CRXFormulaField.Text = "'" & QtnTerms4 & "'"

        'If CRXFormulaField.Name = "{@TaxSplit}" Then CRXFormulaField.text = "'" & Taxsplit & "'"
        If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.Text = "'" & BANK_DET & "'"
        If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.Text = "'" & PAN_NO & "'"
        If CRXFormulaField.Name = "{@DL1}" Then CRXFormulaField.Text = "'" & DL1 & "'"
        If CRXFormulaField.Name = "{@DL2}" Then CRXFormulaField.Text = "'" & DL2 & "'"
'        If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.text = "'" & TxtBillName.text & "'"
'        If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.text = "'" & Trim(TXTDEALER.text) & "'"
'        If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.text = "'" & lbladdress.Caption & "'"
'        If TxtPhone.text = "" Then
'            If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
'        Else
'            If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
'        End If
'        If lblIGST.Caption = "Y" Then
'            If CRXFormulaField.Name = "{@IGSTFLAG}" Then CRXFormulaField.text = "'Y'"
'        Else
'            If CRXFormulaField.Name = "{@IGSTFLAG}" Then CRXFormulaField.text = "'N'"
'        End If
'        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
'        If chkTerms.Value = 1 And Trim(Terms1.text) <> "" Then
'            If CRXFormulaField.Name = "{@condition}" Then CRXFormulaField.text = "'" & Trim(Terms1.text) & "'"
'        End If
'        If CRXFormulaField.Name = "{@Area}" Then CRXFormulaField.text = "'" & Trim(TXTAREA.text) & "'"
'        If CRXFormulaField.Name = "{@area2}" Then CRXFormulaField.text = "'" & Trim(TXTAREA.text) & "'"
'        If Trim(TXTTIN.text) <> "" Then
'            If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.text = "'GSTIN: ' & '" & Trim(TXTTIN.text) & "'"
'        Else
'            If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.text = "'UID: ' & '" & Trim(TxtUID.text) & "'"
'        End If
'
'        If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.text = "'" & TxtPhone.text & "'"
'        If CRXFormulaField.Name = "{@Pin}" Then CRXFormulaField.text = "'" & txtPin.text & "'"
'        If CRXFormulaField.Name = "{@VCH_NO}" Then
'            Me.Tag = BIL_PRE & Format(Trim(txtBillNo.text), bill_for) & BILL_SUF
'            CRXFormulaField.text = "'" & Me.Tag & "' "
'        End If
'        If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.text = "'" & Trim(TxtVehicle.text) & "'"
'        If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.text = "'" & Trim(TxtOrder.text) & "'"
'        If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.text = " " & Val(LBLDISCAMT.Caption) & " "
'        If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.text = "'" & Trim(TxtCode.text) & "'"
'        If OptDiscAmt.Value = True Then
'            If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'N'"
'        Else
'            If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'Y'"
'        End If
'        If CRXFormulaField.Name = "{@RcptAmt}" Then CRXFormulaField.text = " " & Rcptamt & " "
'        If CRXFormulaField.Name = "{@Frieght}" Then CRXFormulaField.text = "'" & Trim(lblFrieght.text) & "'"
'        If CRXFormulaField.Name = "{@FC}" Then CRXFormulaField.text = " " & Val(TxtFrieght.text) & " "
'        If CRXFormulaField.Name = "{@HANDLE}" Then CRXFormulaField.text = " '" & Trim(lblhandle.text) & "' "
'        If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.text = " " & Val(Txthandle.text) & " "
'        If CRXFormulaField.Name = "{@DISCPER}" Then CRXFormulaField.text = " " & Val(TXTTOTALDISC.text) & " "
'
'        If Val(LBLRETAMT.Caption) = 0 Then
'            If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'N' "
'        Else
'            If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'Y' "
'        End If
'        If CRXFormulaField.Name = "{@EXCHANGE}" Then CRXFormulaField.text = " " & Val(LBLRETAMT.Caption) & " "
'        If lblcredit.Caption = "0" Then
'            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Cash'"
'        Else
'            If Val(txtcrdays.text) > 0 Then
'                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'" & txtcrdays.text & "'" & "' Days Credit'"
'            Else
'                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Credit'"
'            End If
'        End If
    Next
    
    'ACT_CODE = '" & DataList2.BoundText & "' AND
    If OPTPERIOD.Value = True Then
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='GI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='HI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='SV' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='WO' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='TF' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='SI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Value = False Then
                Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    Else
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='GI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='HI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='SV' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='WO' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='TF' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='SI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Value = False Then
                Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    End If
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
        Report.Database.SetDataSource oRs, 3, 1
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM ITEMMAST ")
        Report.Database.SetDataSource oRs, 3, 2
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM ITEMMAST ")
        Report.Database.SetDataSource oRs, 3, 3
        Set oRs = Nothing
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "PRINT BILLS"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdSearch_Click()
    If Trim(TXTREFNO.Text) = "" Then Exit Sub
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim n, M As Long
    
    
    GRDTranx.TextMatrix(0, 12) = "BILL DETAILS"
    GRDTranx.ColWidth(12) = 2400
    LBLTRXTOTAL.Caption = "0.00"
    LBLDISCOUNT.Caption = "0.00"
    LBLNET.Caption = "0.00"
    LBLCOST.Caption = "0.00"
    LBLPROFIT.Caption = "0.00"
    lblcommi.Caption = "0.00"
    lblFrieght.Caption = ""
    lblhandle.Caption = ""
    lbltaxamt.Caption = ""
    lblcess.Caption = ""
    
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    n = 1
    M = 0
    On Error GoTo ERRHAND
    'BILL_NAME LIKE '%" & txtCustomerName.Text & "%' AND
    Screen.MousePointer = vbHourglass
    
    db.Execute "Update TRXMAST Set REF_NO='' where isnull(REF_NO)"
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMAST WHERE REF_NO Like '%" & Trim(TXTREFNO.Text) & "%' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO' OR TRX_TYPE='VI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        
    lblbillnos = ""
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
        rstTRANX.MoveLast
        lblbillnos.Caption = rstTRANX!VCH_NO
        rstTRANX.MoveFirst
        lblbillnos.Caption = "From : " & rstTRANX!VCH_NO & " to " & lblbillnos.Caption
        
    Else
        vbalProgressBar1.Max = 100
    End If
    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!TRX_TYPE
        Select Case rstTRANX!TRX_TYPE
            Case "GI"
                GRDTranx.TextMatrix(M, 2) = "GI"
            Case "SV"
                GRDTranx.TextMatrix(M, 2) = "SV"
            Case "RI"
                GRDTranx.TextMatrix(M, 2) = "RT"
            Case "TF"
                GRDTranx.TextMatrix(M, 2) = "ST"
            Case "WO"
                GRDTranx.TextMatrix(M, 2) = "PT"
            Case "SI"
                GRDTranx.TextMatrix(M, 2) = "WS"
            Case "VI"
                GRDTranx.TextMatrix(M, 2) = "VN"
            Case "HI"
                GRDTranx.TextMatrix(M, 2) = "GR"
        End Select
        GRDTranx.TextMatrix(M, 3) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 4) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 5) = Format(Round(rstTRANX!VCH_AMOUNT, 2), "0.00")
'        If rstTRANX!SLSM_CODE = "A" Then
'
'        ElseIf rstTRANX!SLSM_CODE = "P" Then
'            GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(Round((rstTRANX!DISCOUNT * 100 / rstTRANX!VCH_AMOUNT), 2), "0.00"))
'        End If
        GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(rstTRANX!DISCOUNT, "0.00"))
        GRDTranx.TextMatrix(M, 7) = Format(Round(rstTRANX!NET_AMOUNT, 2), "0.00") 'Format(Round(Val(GRDTranx.TextMatrix(M, 5)) - Val(GRDTranx.TextMatrix(M, 6)), 2), "0.00")
        
        CMDEXIT.Tag = IIf(IsNull(rstTRANX!DISCOUNT), "0", Format(rstTRANX!DISCOUNT, "0.00"))
        'GRDTranx.TextMatrix(M, 7) = Format(Round(Val(GRDTranx.TextMatrix(M, 5)), 2), "0.00")
        If frmLogin.rs!Level <> "0" Then
            GRDTranx.TextMatrix(M, 8) = "xxx"
            GRDTranx.TextMatrix(M, 9) = "xxx"
        Else
            GRDTranx.TextMatrix(M, 8) = IIf(IsNull(rstTRANX!COMM_AMT), "0", Format(rstTRANX!COMM_AMT, "0.00"))
            GRDTranx.TextMatrix(M, 9) = IIf(IsNull(rstTRANX!PAY_AMOUNT), "0", Format(rstTRANX!PAY_AMOUNT, "0.00"))
        End If
        GRDTranx.TextMatrix(M, 10) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        GRDTranx.TextMatrix(M, 11) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS), "", ", " & rstTRANX!BILL_ADDRESS)

        CMDDISPLAY.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
        
        
        'GRDTranx.TextMatrix(M, 12) = Format(Val(CMDDISPLAY.Tag), "0.00")
        GRDTranx.TextMatrix(M, 12) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        GRDTranx.TextMatrix(M, 13) = Format(Val(FRMEMAIN.Tag), "0.00")
        GRDTranx.TextMatrix(M, 14) = Format(Val(FRMEBILL.Tag), "0.00")
        GRDTranx.TextMatrix(M, 15) = rstTRANX!TRX_YEAR
        
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + Val(GRDTranx.TextMatrix(M, 5)), "0.00")
        LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption) + Val(GRDTranx.TextMatrix(M, 6)), "0.00")
        LBLNET.Caption = Format(Val(LBLNET.Caption) + Val(GRDTranx.TextMatrix(M, 7)), "0.00")
        
        lblFrieght.Caption = Format(Val(lblFrieght.Caption) + IIf(IsNull(rstTRANX!FRIEGHT), 0, rstTRANX!FRIEGHT), "0.00")
        lblhandle.Caption = Format(Val(lblhandle.Caption) + IIf(IsNull(rstTRANX!HANDLE), 0, rstTRANX!HANDLE), "0.00")
        If frmLogin.rs!Level <> "0" Then
            lblcommi.Caption = "xxx"
            LBLCOST.Caption = "xxx"
            LBLPROFIT.Caption = "xxx"
            GRDTranx.TextMatrix(M, 16) = "xxx"
            GRDTranx.TextMatrix(M, 17) = "xxx"
        Else
            lblcommi.Caption = Format(Val(lblcommi.Caption) + Val(GRDTranx.TextMatrix(M, 8)), "0.00")
            LBLCOST.Caption = Format(Val(LBLCOST.Caption) + Val(GRDTranx.TextMatrix(M, 9)), "0.00")
            GRDTranx.TextMatrix(M, 16) = Format(Round((Val(GRDTranx.TextMatrix(M, 5)) - (Val(GRDTranx.TextMatrix(M, 6)) + Val(GRDTranx.TextMatrix(M, 8)))) - Val(GRDTranx.TextMatrix(M, 9)), 2), "0.00")
            If Val(GRDTranx.TextMatrix(M, 9)) = 0 Then
                GRDTranx.TextMatrix(M, 17) = "0.00"
            ElseIf Val(GRDTranx.TextMatrix(M, 5)) = 0 Then
                GRDTranx.TextMatrix(M, 17) = "100.00"
            Else
                GRDTranx.TextMatrix(M, 17) = Format(Round((((Val(GRDTranx.TextMatrix(M, 5)) - (Val(GRDTranx.TextMatrix(M, 6)) + Val(GRDTranx.TextMatrix(M, 8)))) * 100) / Val(GRDTranx.TextMatrix(M, 9))) - 100, 2), "0.00")
            End If
            'LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
        End If
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND TRX_TYPE = 'DR' AND INV_TRX_TYPE  = '" & rstTRANX!TRX_TYPE & "' AND INV_NO = '" & rstTRANX!VCH_NO & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "'", db, adOpenForwardOnly
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            GRDTranx.TextMatrix(M, 18) = IIf(IsNull(RSTTRXFILE!RCVD_AMOUNT), "", Format(RSTTRXFILE!RCVD_AMOUNT, "0.00"))
            GRDTranx.TextMatrix(M, 19) = Format(Round(Val(GRDTranx.TextMatrix(M, 7)) - Val(GRDTranx.TextMatrix(M, 18)), 2), "0.00")
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        
        n = n + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If frmLogin.rs!Level = "0" Then
        LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
    End If
        
    flagchange.Caption = ""
    flagchange2.Caption = ""
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

Private Sub CmdUserrep_Click()

    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select Customer from the list", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTUSERREPORT"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'ACT_CODE = '" & DataList2.BoundText & "' AND
    
    If frmLogin.rs!Level = "5" Then
        Report.RecordSelectionFormula = "({TRXMAST.C_USER_ID}= '" & frmLogin.rs!USER_ID & "' AND ({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        If OPTPERIOD.Value = True Then
            Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    End If
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "DAY WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Cmdyear_Click()
    Call Report_Generate("Y")
End Sub

Private Sub CmdZeroBills_Click()
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim n, M As Long
    
    
    GRDTranx.TextMatrix(0, 12) = "BILL DETAILS"
    GRDTranx.ColWidth(12) = 2400
    LBLTRXTOTAL.Caption = "0.00"
    LBLDISCOUNT.Caption = "0.00"
    LBLNET.Caption = "0.00"
    LBLCOST.Caption = "0.00"
    LBLPROFIT.Caption = "0.00"
    lblcommi.Caption = "0.00"
    lblFrieght.Caption = ""
    lblhandle.Caption = ""
    lbltaxamt.Caption = ""
    lblcess.Caption = ""
    
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    n = 1
    M = 0
    On Error GoTo ERRHAND
    'BILL_NAME LIKE '%" & txtCustomerName.Text & "%' AND
    Screen.MousePointer = vbHourglass
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMAST WHERE NET_AMOUNT <= 0 AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='VI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        
    lblbillnos = ""
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
        rstTRANX.MoveLast
        lblbillnos.Caption = rstTRANX!VCH_NO
        rstTRANX.MoveFirst
        lblbillnos.Caption = "From : " & rstTRANX!VCH_NO & " to " & lblbillnos.Caption
        
    Else
        vbalProgressBar1.Max = 100
    End If
    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!TRX_TYPE
        Select Case rstTRANX!TRX_TYPE
            Case "GI"
                GRDTranx.TextMatrix(M, 2) = "GI"
            Case "SV"
                GRDTranx.TextMatrix(M, 2) = "SV"
            Case "TF"
                GRDTranx.TextMatrix(M, 2) = "ST"
            Case "RI"
                GRDTranx.TextMatrix(M, 2) = "RT"
            Case "TF"
                GRDTranx.TextMatrix(M, 2) = "ST"
            Case "WO"
                GRDTranx.TextMatrix(M, 2) = "PT"
            Case "SI"
                GRDTranx.TextMatrix(M, 2) = "WS"
            Case "VI"
                GRDTranx.TextMatrix(M, 2) = "VN"
            Case "HI"
                GRDTranx.TextMatrix(M, 2) = "GR"
        End Select
        GRDTranx.TextMatrix(M, 3) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 4) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 5) = Format(Round(rstTRANX!VCH_AMOUNT, 2), "0.00")
'        If rstTRANX!SLSM_CODE = "A" Then
'
'        ElseIf rstTRANX!SLSM_CODE = "P" Then
'            GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(Round((rstTRANX!DISCOUNT * 100 / rstTRANX!VCH_AMOUNT), 2), "0.00"))
'        End If
        GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(rstTRANX!DISCOUNT, "0.00"))
        GRDTranx.TextMatrix(M, 7) = Format(Round(rstTRANX!NET_AMOUNT, 2), "0.00") 'Format(Round(Val(GRDTranx.TextMatrix(M, 5)) - Val(GRDTranx.TextMatrix(M, 6)), 2), "0.00")
        
        CMDEXIT.Tag = IIf(IsNull(rstTRANX!DISCOUNT), "0", Format(rstTRANX!DISCOUNT, "0.00"))
        'GRDTranx.TextMatrix(M, 7) = Format(Round(Val(GRDTranx.TextMatrix(M, 5)), 2), "0.00")
        If frmLogin.rs!Level <> "0" Then
            GRDTranx.TextMatrix(M, 8) = "xxx"
            GRDTranx.TextMatrix(M, 9) = "xxx"
        Else
            GRDTranx.TextMatrix(M, 8) = IIf(IsNull(rstTRANX!COMM_AMT), "0", Format(rstTRANX!COMM_AMT, "0.00"))
            GRDTranx.TextMatrix(M, 9) = IIf(IsNull(rstTRANX!PAY_AMOUNT), "0", Format(rstTRANX!PAY_AMOUNT, "0.00"))
        End If
        GRDTranx.TextMatrix(M, 10) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        GRDTranx.TextMatrix(M, 11) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS), "", ", " & rstTRANX!BILL_ADDRESS)

        CMDDISPLAY.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
        
        
        'GRDTranx.TextMatrix(M, 12) = Format(Val(CMDDISPLAY.Tag), "0.00")
        GRDTranx.TextMatrix(M, 12) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        GRDTranx.TextMatrix(M, 13) = Format(Val(FRMEMAIN.Tag), "0.00")
        GRDTranx.TextMatrix(M, 14) = Format(Val(FRMEBILL.Tag), "0.00")
        GRDTranx.TextMatrix(M, 15) = rstTRANX!TRX_YEAR
        
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + Val(GRDTranx.TextMatrix(M, 5)), "0.00")
        LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption) + Val(GRDTranx.TextMatrix(M, 6)), "0.00")
        LBLNET.Caption = Format(Val(LBLNET.Caption) + Val(GRDTranx.TextMatrix(M, 7)), "0.00")
        
        lblFrieght.Caption = Format(Val(lblFrieght.Caption) + IIf(IsNull(rstTRANX!FRIEGHT), 0, rstTRANX!FRIEGHT), "0.00")
        lblhandle.Caption = Format(Val(lblhandle.Caption) + IIf(IsNull(rstTRANX!HANDLE), 0, rstTRANX!HANDLE), "0.00")
        If frmLogin.rs!Level <> "0" Then
            lblcommi.Caption = "xxx"
            LBLCOST.Caption = "xxx"
            LBLPROFIT.Caption = "xxx"
            GRDTranx.TextMatrix(M, 16) = "xxx"
            GRDTranx.TextMatrix(M, 17) = "xxx"
        Else
            lblcommi.Caption = Format(Val(lblcommi.Caption) + Val(GRDTranx.TextMatrix(M, 8)), "0.00")
            LBLCOST.Caption = Format(Val(LBLCOST.Caption) + Val(GRDTranx.TextMatrix(M, 9)), "0.00")
            GRDTranx.TextMatrix(M, 16) = Format(Round((Val(GRDTranx.TextMatrix(M, 5)) - (Val(GRDTranx.TextMatrix(M, 6)) + Val(GRDTranx.TextMatrix(M, 8)))) - Val(GRDTranx.TextMatrix(M, 9)), 2), "0.00")
            If Val(GRDTranx.TextMatrix(M, 9)) = 0 Then
                GRDTranx.TextMatrix(M, 17) = "0.00"
            ElseIf Val(GRDTranx.TextMatrix(M, 5)) = 0 Then
                GRDTranx.TextMatrix(M, 17) = "100.00"
            Else
                GRDTranx.TextMatrix(M, 17) = Format(Round((((Val(GRDTranx.TextMatrix(M, 5)) - (Val(GRDTranx.TextMatrix(M, 6)) + Val(GRDTranx.TextMatrix(M, 8)))) * 100) / Val(GRDTranx.TextMatrix(M, 9))) - 100, 2), "0.00")
            End If
            'LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
        End If
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND TRX_TYPE = 'DR' AND INV_TRX_TYPE  = '" & rstTRANX!TRX_TYPE & "' AND INV_NO = '" & rstTRANX!VCH_NO & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "'", db, adOpenForwardOnly
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            GRDTranx.TextMatrix(M, 18) = IIf(IsNull(RSTTRXFILE!RCVD_AMOUNT), "", Format(RSTTRXFILE!RCVD_AMOUNT, "0.00"))
            GRDTranx.TextMatrix(M, 19) = Format(Round(Val(GRDTranx.TextMatrix(M, 7)) - Val(GRDTranx.TextMatrix(M, 18)), 2), "0.00")
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        
        n = n + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If frmLogin.rs!Level = "0" Then
        LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
    End If
        
    flagchange.Caption = ""
    flagchange2.Caption = ""
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

Private Sub Command1_Click()
    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select Customer from the list", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTSALESDAY"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'ACT_CODE = '" & DataList2.BoundText & "' AND
    If OPTPERIOD.Value = True Then
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='GI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='HI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='SV' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='WO' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='TF' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='SI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Value = False Then
                Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    Else
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='GI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='HI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='SV' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='WO' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='TF' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE}='SI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Value = False Then
                Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='WO' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    End If
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "DAY WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CMDEfile_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    ReportNameVar = Rptpath & "RPTSALESREG3"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
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
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@BillNos}" Then CRXFormulaField.Text = "'" & lblbillnos.Caption & "' "
        If CRXFormulaField.Name = "{@Amount}" Then CRXFormulaField.Text = "'" & LBLNET.Caption & "' "
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
    Next
    frmreport.Caption = "SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
   Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
    db.Execute "delete From SALESREG"
    
    LBLTRXTOTAL.Caption = "0.00"
    LBLDISCOUNT.Caption = "0.00"
    LBLNET.Caption = "0.00"
    LBLCOST.Caption = "0.00"
    LBLPROFIT.Caption = "0.00"
    lblcommi.Caption = "0.00"
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    n = 1
    M = 0
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    If OPTPERIOD.Value = True Then
        If OPTGST.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptVan.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Optservice.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptPetty.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='WO')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptRT.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='TF')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptWs.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            If OptPetty.Value = False Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SI' OR TRX_TYPE='VI' OR TRX_TYPE='RI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            Else
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='HI' OR TRX_TYPE='WO' OR TRX_TYPE='GI' OR TRX_TYPE='SI' OR TRX_TYPE='VI' OR TRX_TYPE='RI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            End If
        End If
    Else
        If OPTGST.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptVan.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Optservice.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptPetty.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='WO'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptRT.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='TF'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptWs.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SI'  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            If OptPetty.Value = False Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SI' OR TRX_TYPE='VI' OR TRX_TYPE='RI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            Else
                rstTRANX.Open "SELECT * From TRXMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='HI' OR TRX_TYPE='WO' OR TRX_TYPE='GI' OR TRX_TYPE='SI' OR TRX_TYPE='VI' OR TRX_TYPE='RI') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            End If
        End If
    End If
        
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTSALEREG.Properties("Update Criteria").Value = adCriteriaKey
    Do Until rstTRANX.EOF
        M = M + 1
        
        CMDDISPLAY.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
        
        'If rstTRANX!TRX_TYPE <> "SI" Then GoTo SKIP
        
        EXSALEAMT = 0
        TAXSALEAMT = 0
        TaxAmt = 0
        MRPVALUE = 0
        DISCAMT = 0
        'TAXRATE = RSTTRXFILE!SALES_TAX
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select * From TRXFILE WHERE VCH_NO = " & rstTRANX!VCH_NO & " ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            If RSTtax!SALES_TAX > 0 And RSTtax!check_flag = "V" Then
                TAXSALEAMT = TAXSALEAMT + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                TaxAmt = TaxAmt + Round((RSTtax!PTR * RSTtax!SALES_TAX / 100) * RSTtax!QTY, 2)
                
            Else
                If RSTtax!SALE_1_FLAG = "1" Then
                    TaxAmt = TaxAmt + Round((RSTtax!SALES_PRICE - RSTtax!PTR) * RSTtax!QTY, 2)
                    MRPVALUE = Round(MRPVALUE + (100 * RSTtax!MRP / 105) * RSTtax!QTY, 2)
                End If
                EXSALEAMT = EXSALEAMT + RSTtax!TRX_TOTAL
            End If
            DISCAMT = Round(DISCAMT + IIf(IsNull(RSTtax!LINE_DISC), 0, RSTtax!TRX_TOTAL * RSTtax!LINE_DISC / 100), 2)
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
        RSTSALEREG.AddNew
        TAXSALEAMT = TAXSALEAMT - TaxAmt
        RSTSALEREG!VCH_NO = rstTRANX!VCH_NO 'N
        RSTSALEREG!TRX_TYPE = rstTRANX!TRX_TYPE
        RSTSALEREG!VCH_DATE = rstTRANX!VCH_DATE
        RSTSALEREG!DISCOUNT = DISCAMT
        RSTSALEREG!VCH_AMOUNT = rstTRANX!VCH_AMOUNT - IIf(IsNull(rstTRANX!DISCOUNT), "", rstTRANX!DISCOUNT)
        'RSTSALEREG!PAYAMOUNT = Val(GRDTranx.TextMatrix(M, 9))
        RSTSALEREG!ACT_NAME = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
        RSTSALEREG!ACT_CODE = IIf(IsNull(rstTRANX!ACT_CODE), "", rstTRANX!ACT_CODE)
        
        Dim RSTACTCODE As ADODB.Recordset
        Set RSTACTCODE = New ADODB.Recordset
        RSTACTCODE.Open "SELECT KGST FROM CUSTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTACTCODE.EOF And RSTACTCODE.BOF) Then
            RSTSALEREG!TIN_NO = RSTACTCODE!KGST
        End If
        RSTACTCODE.Close
        Set RSTACTCODE = Nothing
        RSTSALEREG!EXMPSALES_AMT = EXSALEAMT
        RSTSALEREG!TAXSALES_AMT = TAXSALEAMT
        RSTSALEREG!TAXAMOUNT = TaxAmt
        RSTSALEREG!TAXRATE = TAXRATE
        CMDDISPLAY.Tag = Val(CMDDISPLAY.Tag) + EXSALEAMT
        FRMEMAIN.Tag = Val(FRMEMAIN.Tag) + TAXSALEAMT
        FRMEBILL.Tag = Val(FRMEBILL.Tag) + TaxAmt
        RSTSALEREG.Update
            
        
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        n = n + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    LBLNET.Caption = Format(Val(LBLTRXTOTAL.Caption) - Val(LBLDISCOUNT.Caption), "0.00")
    If frmLogin.rs!Level = "0" Then
        LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
    End If
        
    flagchange.Caption = ""
    flagchange2.Caption = ""
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

Private Sub CmdMonthly_Click()

    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RptMonthwise"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.RecordSelectionFormula = "( {TRXMAST.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='VI') )"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
    Next
    
    frmreport.Caption = "DAILY SALES ANALYSIS"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CMDREGISTER_Click()
    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Integer
    Screen.MousePointer = vbHourglass

    ReportNameVar = Rptpath & "RPTLEDREP"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If DataList2.BoundText = "" Then
        Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    End If
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
'    Dim i As lONG
'    Screen.MousePointer = vbHourglass
'
'    ReportNameVar = Rptpath & "RPTSALESREG2"
'    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
'    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
'    Set CRXFormulaFields = Report.FormulaFields
'    For i = 1 To Report.Database.Tables.COUNT
'        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
'    Next i
'    For Each CRXFormulaField In CRXFormulaFields
'        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
'        If CRXFormulaField.Name = "{@BillNos}" Then CRXFormulaField.Text = "'" & lblbillnos.Caption & "' "
'        If CRXFormulaField.Name = "{@Amount}" Then CRXFormulaField.Text = "'" & LBLNET.Caption & "' "
'        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
'    Next
'    frmreport.Caption = "SALES REGISTER"
'    Call GENERATEREPORT
'    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdReport_Click()
    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select Customer from the list", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    
    ReportNameVar = Rptpath & "RPTSALESREP1"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OPTPERIOD.Value = True Then
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE} = 'GI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE} = 'SV' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE} = 'TF' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE} = 'WO' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE} = 'SI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE} = 'HI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Visible = False Then
                Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE} = 'SV' OR {TRXMAST.TRX_TYPE} = 'HI' OR {TRXMAST.TRX_TYPE} = 'GI' OR {TRXMAST.TRX_TYPE} = 'VI' OR {TRXMAST.TRX_TYPE} = 'RI' OR {TRXMAST.TRX_TYPE} = 'SI')AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE} = 'SV' OR {TRXMAST.TRX_TYPE} = 'HI' OR {TRXMAST.TRX_TYPE} = 'GI' OR {TRXMAST.TRX_TYPE} = 'WO' OR {TRXMAST.TRX_TYPE} = 'VI' OR {TRXMAST.TRX_TYPE} = 'RI' OR {TRXMAST.TRX_TYPE} = 'SI')AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    Else
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE} = 'GI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE} = 'SV' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE} = 'TF' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE} = 'WO' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE} = 'SI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.TRX_TYPE} = 'HI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Visible = False Then
                Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXMAST.TRX_TYPE} = 'SV' OR {TRXMAST.TRX_TYPE} = 'HI' OR {TRXMAST.TRX_TYPE} = 'GI' OR {TRXMAST.TRX_TYPE} = 'VI' OR {TRXMAST.TRX_TYPE} = 'RI' OR {TRXMAST.TRX_TYPE} = 'SI')AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXMAST.TRX_TYPE} = 'SV' OR {TRXMAST.TRX_TYPE} = 'HI' OR {TRXMAST.TRX_TYPE} = 'GI' OR {TRXMAST.TRX_TYPE} = 'WO' OR {TRXMAST.TRX_TYPE} = 'VI' OR {TRXMAST.TRX_TYPE} = 'RI' OR {TRXMAST.TRX_TYPE} = 'SI')AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    End If
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdview_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select Customer from the list", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTSALESREPORTAG"
    
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'ACT_CODE = '" & DataList2.BoundText & "' AND
    If OPTPERIOD.Value = True Then
        If DataList5.BoundText = "" Then
            Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.AGENT_CODE} = '" & DataList5.BoundText & "' AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    Else
        If DataList5.BoundText = "" Then
            Report.RecordSelectionFormula = "({TRXFILE.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.AGENT_CODE} = '" & DataList5.BoundText & "' AND {TRXFILE.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    End If
    '
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "DAY WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
    
'    Dim i As Integer
'    Screen.MousePointer = vbHourglass
'
'    ReportNameVar = Rptpath & "RPTSALESREG1"
'    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
'    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
'    Set CRXFormulaFields = Report.FormulaFields
'    For i = 1 To Report.Database.Tables.COUNT
'        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
'        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
'            Set oRs = New ADODB.Recordset
'            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
'            Report.Database.Tables(i).SetDataSource oRs, 3
'            Set oRs = Nothing
'        End If
'    Next i
'    Report.DiscardSavedData
'    Report.VerifyOnEveryPrint = True
'    For Each CRXFormulaField In CRXFormulaFields
'        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
'        If CRXFormulaField.Name = "{@BillNos}" Then CRXFormulaField.Text = "'" & lblbillnos.Caption & "' "
'        If CRXFormulaField.Name = "{@Amount}" Then CRXFormulaField.Text = "'" & LBLNET.Caption & "' "
'        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
'    Next
'    frmreport.Caption = "SALES REGISTER"
'    Call GENERATEREPORT
'    Screen.MousePointer = vbNormal
'
'    Exit Sub
'
'    Dim TRXFILE As ADODB.Recordset
'    Dim RSTTRXFILE As ADODB.Recordset
'    Dim RSTSALEREG As ADODB.Recordset
'    Dim rstTRANX As ADODB.Recordset
'    Dim FROMDATE As Date
'    Dim TODATE As Date
'    Dim SLIPAMT As Double
'    Dim M As Long
'
'    db.Execute "delete From SLIP_REG"
'
'    FROMDATE = DTFROM.value 'Format(DTFROM.Value, "MM,DD,YYYY")
'    TODATE = DTTO.value 'Format(DTTO.Value, "MM,DD,YYYY")
'
'    vbalProgressBar1.value = 0
'    vbalProgressBar1.ShowText = True
'
'    On Error GoTo eRRHAND
'    Screen.MousePointer = vbHourglass
'    Set rstTRANX = New ADODB.Recordset
'    If OptGST.value = True Then
'        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' ", db, adOpenStatic, adLockReadOnly
'    ElseIf OptRT.value = True Then
'        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='RI' ", db, adOpenStatic, adLockReadOnly
'    ElseIf OptService.value = True Then
'        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV' ", db, adOpenStatic, adLockReadOnly
'    ElseIf Optpetty.value = True Then
'        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='WO' ", db, adOpenStatic, adLockReadOnly
'    ElseIf OptWS.value = True Then
'        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SI' ", db, adOpenStatic, adLockReadOnly
'    ElseIf OptVan.value = True Then
'        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' ", db, adOpenStatic, adLockReadOnly
'    Else
'        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' ", db, adOpenStatic, adLockReadOnly
'    End If
'
'    If rstTRANX.RecordCount > 0 Then
'        vbalProgressBar1.Max = rstTRANX.RecordCount
'    Else
'        vbalProgressBar1.Max = 100
'    End If
'    rstTRANX.Close
'    Set rstTRANX = Nothing
'
'    Set RSTSALEREG = New ADODB.Recordset
'    RSTSALEREG.Open "SELECT * From SLIP_REG", db, adOpenStatic, adLockOptimistic, adCmdText
'    M = 0
'    Do Until FROMDATE > TODATE
'        SLIPAMT = 0
'        M = M + 1
'
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_DATE = '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SI' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
'        Do Until RSTTRXFILE.EOF
'            CmdDisplay.Tag = ""
'            If RSTTRXFILE!SLSM_CODE = "A" Then
'                CmdDisplay.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(RSTTRXFILE!DISCOUNT, "0.00"))
'            ElseIf RSTTRXFILE!SLSM_CODE = "P" Then
'                CmdDisplay.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(Round((RSTTRXFILE!DISCOUNT * RSTTRXFILE!VCH_AMOUNT) / 100, 2), "0.00"))
'            End If
'            cmdview.Tag = ""
'            'cmdview.Tag = IIf(IsNull(RSTTRXFILE!ADD_AMOUNT), "", RSTTRXFILE!ADD_AMOUNT)
'            SLIPAMT = SLIPAMT + Round(RSTTRXFILE!VCH_AMOUNT - Val(CmdDisplay.Tag), 0) '+ Val(cmdview.Tag))
'            RSTTRXFILE.MoveNext
'            vbalProgressBar1.value = vbalProgressBar1.value + 1
'        Loop
'        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'            RSTSALEREG.AddNew
'            RSTTRXFILE.MoveLast
'            RSTSALEREG!VCH_END_NO = RSTTRXFILE!VCH_NO
'            RSTTRXFILE.MoveFirst
'            RSTSALEREG!VCH_START_NO = RSTTRXFILE!VCH_NO
'            RSTSALEREG!VCH_DATE = RSTTRXFILE!VCH_DATE
'            RSTSALEREG!rec_no = M
'            RSTSALEREG!TRX_TYPE = "S"
'            RSTSALEREG!VCH_AMOUNT = SLIPAMT
'            RSTSALEREG.Update
'        End If
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
'
'        FROMDATE = DateAdd("d", FROMDATE, 1)
'    Loop
'    RSTSALEREG.Close
'    Set RSTSALEREG = Nothing
'
'    'CHECKFLAG = 1
'    Sleep (300)
'
'    ReportNameVar = Rptpath & "RptSalreg"
'    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
'    'selectionformla = "( {TRXFILE.FREE_QTY}>0 and {TRXFILE.VCH_DATE}<=# " & TODATE & " # and {TRXFILE.VCH_DATE}>=# " & FROMDATE & " # and {TRXFILE.MFGR}='" & DataList3.BoundText & "')"
'    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
'    Set CRXFormulaFields = Report.FormulaFields
'
'    'Dim i As lONG
'    For i = 1 To Report.Database.Tables.COUNT
'        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
'    Next i
'    Report.DiscardSavedData
'    Report.VerifyOnEveryPrint = True
'    For Each CRXFormulaField In CRXFormulaFields
'        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
'        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
'    Next
'    frmreport.Caption = "SALES REGISTER"
'
'    vbalProgressBar1.ShowText = False
'    vbalProgressBar1.value = 0
'    GRDTranx.Visible = True
'
'
'    GENERATEREPORT
'    'GRDTranx.SetFocus
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'eRRHAND:
'    Screen.MousePointer = vbDefault
'    MsgBox Err.Description

End Sub

Private Sub cmdwoprint_Click()
    FRMBillTransfer.Show
    FRMBillTransfer.SetFocus
End Sub

Private Sub Command2_Click()

    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Long
    Screen.MousePointer = vbHourglass
        
    On Error GoTo ERRHAND
    Dim rstdbt As ADODB.Recordset
    Dim rstdbt2 As ADODB.Recordset
    Set rstdbt = New ADODB.Recordset
    rstdbt.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    rstdbt.Properties("Update Criteria").Value = adCriteriaKey
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
        
    ReportNameVar = Rptpath & "RPTCRSALE"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OPTGST.Value = True Then
        If OptBoth = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='GI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf optCash = True Then
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'Y' AND {TRXMAST.TRX_TYPE}='GI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'N' AND {TRXMAST.TRX_TYPE}='GI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    ElseIf OptVan.Value = True Then
        If OptBoth = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='HI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf optCash = True Then
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'Y' AND {TRXMAST.TRX_TYPE}='HI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'N' AND {TRXMAST.TRX_TYPE}='HI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    ElseIf Optservice.Value = True Then
        If OptBoth = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='SV' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf optCash = True Then
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'Y' AND {TRXMAST.TRX_TYPE}='SV' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'N' AND {TRXMAST.TRX_TYPE}='SV' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    ElseIf OptPetty.Value = True Then
        If OptBoth = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='WO' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf optCash = True Then
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'Y' AND {TRXMAST.TRX_TYPE}='WO' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'N' AND {TRXMAST.TRX_TYPE}='WO' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    ElseIf OptRT.Value = True Then
        If OptBoth = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='TF' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf optCash = True Then
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'Y' AND {TRXMAST.TRX_TYPE}='TF' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'N' AND {TRXMAST.TRX_TYPE}='TF' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    ElseIf OptWs.Value = True Then
        If OptBoth = True Then
            Report.RecordSelectionFormula = "({TRXMAST.TRX_TYPE}='SI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf optCash = True Then
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'Y' AND {TRXMAST.TRX_TYPE}='SI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.POST_FLAG} = 'N' AND {TRXMAST.TRX_TYPE}='SI' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    Else
        If OptBoth = True Then
            Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf optCash = True Then
            Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.POST_FLAG} = 'Y' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='VI' OR {TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.POST_FLAG} = 'N' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    End If
    
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "CASH / CREDIT REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Command3_Click()

    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Long
    
    On Error GoTo ERRHAND
'    If DataList4.BoundText = "" Then
'        MsgBox "Please select the Area from the list", vbOKOnly, "Area wise Report"
'        Exit Sub
'    End If
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select the Customer from the list", vbOKOnly, "Area wise Report"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    ReportNameVar = Rptpath & "RPTAreaRep"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    If DataList4.BoundText = "" Then
        If OPTCUSTOMER.Value = False Then
            Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    Else
        If OPTCUSTOMER.Value = False Then
            Report.RecordSelectionFormula = "({custmast.AREA}='" & DataList4.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND {custmast.AREA}='" & DataList4.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    End If
    Set CRXFormulaFields = Report.FormulaFields
    
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Command4_Click()

    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Long
    Screen.MousePointer = vbHourglass
                        
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select the customer from the list", , "Sales Register"
        Exit Sub
    End If
    
    ReportNameVar = Rptpath & "RPTSALEITEM2"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
    'db.Execute "Update trxmast set REF_BILL="" where isnull(REF_BILL)"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            End If
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    If OPTCUSTOMER.Value = True Then
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR (ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1)) AND {TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='GI' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='HI' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='SV' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='WO' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='TF' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='SI'AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Visible = False Then
                Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    Else
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.TRX_TYPE}='GI' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.TRX_TYPE}='HI' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.TRX_TYPE}='SV' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.TRX_TYPE}='WO' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.TRX_TYPE}='TF' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "((ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.TRX_TYPE}='SI'AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Visible = False Then
                Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND (ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND (ISNULL({TRXMAST.REF_BILL}) OR {TRXMAST.REF_BILL} <>1) AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "CUSTOMER WISE SALES ANALYSIS"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Command5_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
                        
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select the customer from the list", , "Sales Register"
        Exit Sub
    End If
    
    ReportNameVar = Rptpath & "RPTSALEITEM"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            End If
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    If OPTCUSTOMER.Value = True Then
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='GI' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='HI' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='SV' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='WO' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='TF' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND {TRXFILE.TRX_TYPE}='SI'AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Visible = False Then
                Report.RecordSelectionFormula = "({TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "({TRXFILE.M_USER_ID} = '" & DataList2.BoundText & "' AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    Else
        If OPTGST.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.TRX_TYPE}='GI' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptVan.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.TRX_TYPE}='HI' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf Optservice.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.TRX_TYPE}='SV' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptPetty.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.TRX_TYPE}='WO' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptRT.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.TRX_TYPE}='TF' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        ElseIf OptWs.Value = True Then
            Report.RecordSelectionFormula = "({TRXFILE.TRX_TYPE}='SI'AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            If OptPetty.Visible = False Then
                Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            Else
                Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
            End If
        End If
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Command6_Click()

    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And OptPetty.Visible = False Then Exit Sub
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select Customer from the list", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    
    Dim searchstring As String
    Dim selcat As Boolean
    searchstring = ""
    selcat = False
    For i = 0 To LstCategory.ListCount - 1
        If LstCategory.Selected(i) = True Then
            searchstring = searchstring & "{ITEMMAST.CATEGORY} = " & "'" & LstCategory.List(i) & "'" & " OR "
            selcat = True
        End If
    Next i
    If Len(searchstring) > 4 Then
        searchstring = Left(searchstring, Len(searchstring) - 4)
    End If
    On Error GoTo ERRHAND
    If selcat = False Then
        ReportNameVar = Rptpath & "RPTSALESREPORT1"
    Else
        ReportNameVar = Rptpath & "RPTSALESREPORT2"
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'ACT_CODE = '" & DataList2.BoundText & "' AND
    If selcat = False Then
        If OPTPERIOD.Value = True Then
            Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "({TRXFILE.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    Else
'        searchstring = "(" & searchstring & ")" & " AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #"
'        Report.RecordSelectionFormula = searchstring
        If OPTPERIOD.Value = True Then
            searchstring = "(" & searchstring & ")" & " AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #"
            Report.RecordSelectionFormula = searchstring
        Else
            searchstring = "(" & searchstring & ")" & " AND {TRXMAST.ACT_CODE} = '" & DataList2.BoundText & "' AND ({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #"
            Report.RecordSelectionFormula = searchstring
        End If
    End If
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "DAY WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
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

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CMDPRINTREGISTER_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
   
    ReportNameVar = Rptpath & "RPTSALESREPORT"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    Dim rpt_Sql As String
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            If Report.Database.Tables(i).Name = "TRXFILE" Or Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            End If
            Report.Database.Tables(i).SetDataSource oRs, 3
            'Report.Database.Tables.Item(i).SetDataSource
            Set oRs = Nothing
        End If
    Next i
    If OPTGST.Value = True Then
        Report.RecordSelectionFormula = "(((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND {TRXFILE.TRX_TYPE}='GI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptVan.Value = True Then
        Report.RecordSelectionFormula = "(((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND {TRXFILE.TRX_TYPE}='HI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf Optservice.Value = True Then
        Report.RecordSelectionFormula = "(((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND {TRXFILE.TRX_TYPE}='SV')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptPetty.Value = True Then
        Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='WO')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptRT.Value = True Then
        Report.RecordSelectionFormula = "({TRXFILE.TRX_TYPE}='TF' AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptWs.Value = True Then
        Report.RecordSelectionFormula = "(((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        If OptPetty.Visible = False Then
            Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        Else
            Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='VI' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI' OR {TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    End If
    
    
'    For i = 1 To Report.Database.Tables.COUNT
'        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
'        Set oRs = New ADODB.Recordset
'        Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
'        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then  Report.Database.Tables(i).SetDataSource oRs, 3
'        Set oRs = Nothing
'    Next i
    
    
    
    Set CRXFormulaFields = Report.FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    'If Month(Date) > 1 Then
        'CMBMONTH.ListIndex = Month(Date) - 2
    'Else
        'CMBMONTH.ListIndex = 11
    'End If
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        'OPTGST.Visible = False
        OPTGST.Caption = "Sales (CS)"
        OptVan.Caption = "Sales"
    End If
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "Type"
    GRDTranx.TextMatrix(0, 2) = "TYPE"
    GRDTranx.TextMatrix(0, 3) = "BILL NO"
    GRDTranx.TextMatrix(0, 4) = "BILL DATE"
    GRDTranx.TextMatrix(0, 5) = "BILL AMT"
    GRDTranx.TextMatrix(0, 6) = "DISC AMT"
    GRDTranx.TextMatrix(0, 7) = "NET AMT"
    GRDTranx.TextMatrix(0, 8) = "COMMI"
    GRDTranx.TextMatrix(0, 9) = "COST"
    GRDTranx.TextMatrix(0, 10) = "CUSTOMER"
    GRDTranx.TextMatrix(0, 11) = "Bill Address"
    GRDTranx.TextMatrix(0, 12) = "AREA"
    GRDTranx.TextMatrix(0, 13) = "TAX SALES"
    GRDTranx.TextMatrix(0, 14) = "TAX AMT"
    GRDTranx.TextMatrix(0, 15) = "YEAR"
    GRDTranx.TextMatrix(0, 16) = "Profit"
    GRDTranx.TextMatrix(0, 17) = "Profit%"
    GRDTranx.TextMatrix(0, 18) = "Rcvd Amt"
    GRDTranx.TextMatrix(0, 19) = "Bal Amt"
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        GRDTranx.TextMatrix(0, 20) = ""
        GRDTranx.TextMatrix(0, 21) = ""
    Else
        GRDTranx.TextMatrix(0, 20) = "Gross Amt"
        GRDTranx.TextMatrix(0, 21) = "Tax Amt"
    End If
    GRDTranx.TextMatrix(0, 22) = "Exchange"
    GRDTranx.ColWidth(0) = 700
    GRDTranx.ColWidth(1) = 0
    GRDTranx.ColWidth(2) = 500
    GRDTranx.ColWidth(3) = 1000
    GRDTranx.ColWidth(4) = 1100
    GRDTranx.ColWidth(5) = 1200
    GRDTranx.ColWidth(6) = 1000
    GRDTranx.ColWidth(7) = 1000
    If frmLogin.rs!Level <> "0" Then
        OptPetty.Visible = False
        GRDTranx.ColWidth(8) = 0
        GRDTranx.ColWidth(9) = 0
    Else
        OptPetty.Visible = True
        CmdCunterSales.Visible = True
        CmdCounterReg.Visible = True
        GRDTranx.ColWidth(8) = 1000
        GRDTranx.ColWidth(9) = 1000
    End If
    GRDTranx.ColWidth(10) = 2000
    GRDTranx.ColWidth(11) = 2000
    GRDTranx.ColWidth(12) = 1200
    GRDTranx.ColWidth(13) = 0
    GRDTranx.ColWidth(14) = 0
    GRDTranx.ColWidth(15) = 0
    GRDTranx.ColWidth(16) = 1100
    GRDTranx.ColWidth(17) = 700
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        GRDTranx.ColWidth(20) = 0
        GRDTranx.ColWidth(21) = 0
    Else
        GRDTranx.ColWidth(20) = 1200
        GRDTranx.ColAlignment(20) = 3
        GRDTranx.ColWidth(21) = 1200
        GRDTranx.ColAlignment(21) = 3
    End If
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 3
    GRDTranx.ColAlignment(4) = 3
    GRDTranx.ColAlignment(5) = 3
    GRDTranx.ColAlignment(6) = 6
    GRDTranx.ColAlignment(7) = 6
    GRDTranx.ColAlignment(8) = 6
    GRDTranx.ColAlignment(9) = 6
    GRDTranx.ColAlignment(10) = 1
    GRDTranx.ColAlignment(11) = 1
    GRDTranx.ColAlignment(12) = 1
    GRDTranx.ColAlignment(13) = 6
    GRDTranx.ColAlignment(14) = 6
    GRDTranx.ColAlignment(18) = 6
    GRDTranx.ColAlignment(19) = 6
    GRDTranx.ColWidth(22) = 1200
    GRDTranx.ColAlignment(22) = 6
        
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "Rate"
    GRDBILL.TextMatrix(0, 3) = "Disc %"
    GRDBILL.TextMatrix(0, 4) = "Tax %"
    GRDBILL.TextMatrix(0, 5) = "Qty"
    GRDBILL.TextMatrix(0, 6) = "Amount"
    GRDBILL.TextMatrix(0, 7) = "Batch"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 4800
    GRDBILL.ColWidth(2) = 800
    GRDBILL.ColWidth(3) = 800
    GRDBILL.ColWidth(4) = 800
    GRDBILL.ColWidth(5) = 900
    GRDBILL.ColWidth(6) = 1100
    GRDBILL.ColWidth(7) = 1100
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 3
    GRDBILL.ColAlignment(6) = 6
    GRDBILL.ColAlignment(7) = 1
    
    OptPetty.Visible = False
    CmdCunterSales.Visible = False
    CmdCounterReg.Visible = False
    'OptAll.Visible = False
    cmdwoprint.Value = False

    If frmLogin.rs!Level <> "0" Then
        GRDTranx.ColWidth(16) = 0
        GRDTranx.ColWidth(17) = 0
    Else
        GRDTranx.ColWidth(16) = 1200
        GRDTranx.ColWidth(17) = 900
    End If
    
    If frmLogin.rs!Level = "5" Then
        CmdCunterSales.Visible = True
        CmdCounterReg.Visible = True
        Frame1.Visible = False
    End If
    Call fillcombo
    
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 11130
    'Me.Height = 10125
    Me.Left = 0
    Me.Top = 0
    ACT_FLAG = True
    PHY_FLAG = True
    AREA_FLAG = True
    CAT_FLAG = True
    AGNT_FLAG = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close
    If PHY_FLAG = False Then PHY_REC.Close
    If CAT_FLAG = False Then CAT_REC.Close
    If AREA_FLAG = False Then AREA_REC.Close
    If AGNT_FLAG = False Then AGNT_REC.Close
    
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

Private Sub GRDTranx_DblClick()
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    
    If frmLogin.rs!Level = "5" Then Exit Sub
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
                
    Select Case Trim(GRDTranx.TextMatrix(GRDTranx.Row, 1))
        Case "HI"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDTranx.TextMatrix(GRDTranx.Row, 15)) Then Exit Sub
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                If IsFormLoaded(frmsales) <> True Then
                    frmsales.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    frmsales.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    frmsales.Show
                    frmsales.SetFocus
                    Call frmsales.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMSALES1) <> True Then
                    FRMSALES1.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMSALES1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMSALES1.Show
                    FRMSALES1.SetFocus
                    Call FRMSALES1.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMSALES2) <> True Then
                    FRMSALES2.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMSALES2.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMSALES2.Show
                    FRMSALES2.SetFocus
                    Call FRMSALES2.txtBillNo_KeyDown(13, 0)
                End If
            Else
                If SALESLT_FLAG = "Y" Then
                    If IsFormLoaded(FRMGSTRSM1) <> True Then
                        FRMGSTRSM1.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM1.Show
                        FRMGSTRSM1.SetFocus
                        Call FRMGSTRSM1.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTRSM2) <> True Then
                        FRMGSTRSM2.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM2.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM2.Show
                        FRMGSTRSM2.SetFocus
                        Call FRMGSTRSM2.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTRSM3) <> True Then
                        FRMGSTRSM3.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM3.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM3.Show
                        FRMGSTRSM3.SetFocus
                        Call FRMGSTRSM3.txtBillNo_KeyDown(13, 0)
                    End If
                Else
                    If Val(GRDTranx.TextMatrix(GRDTranx.Row, 5)) = 0 Then Cancelbill_flag = True
                    If IsFormLoaded(FRMGSTR) <> True Then
                        FRMGSTR.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR.Show
                        FRMGSTR.SetFocus
                        Call FRMGSTR.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTR1) <> True Then
                        FRMGSTR1.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR1.Show
                        FRMGSTR1.SetFocus
                        Call FRMGSTR1.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTR2) <> True Then
                        FRMGSTR2.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR2.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR2.Show
                        FRMGSTR2.SetFocus
                        Call FRMGSTR2.txtBillNo_KeyDown(13, 0)
                    End If
                    Cancelbill_flag = False
                End If
            End If
        Case "GI"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDTranx.TextMatrix(GRDTranx.Row, 15)) Then Exit Sub
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                Exit Sub
            Else
                If IsFormLoaded(FRMGST) <> True Then
                    FRMGST.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGST.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGST.Show
                    FRMGST.SetFocus
                    Call FRMGST.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMGST1) <> True Then
                    FRMGST1.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGST1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGST1.Show
                    FRMGST1.SetFocus
                    Call FRMGST1.txtBillNo_KeyDown(13, 0)
                End If
            End If
        Case "WO"
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.rows = 1 Then Exit Sub
            LBLBILLNO.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            LBLBILLAMT.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 5), "0.00")
            LBLDISC.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 6), "0.00")
            LBLNETAMT.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 7), "0.00")
             
            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILE WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "  AND TRX_TYPE = '" & Trim(GRDTranx.TextMatrix(GRDTranx.Row, 1)) & "' AND TRX_YEAR = '" & Trim(GRDTranx.TextMatrix(GRDTranx.Row, 15)) & "'", db, adOpenStatic, adLockReadOnly
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
                GRDBILL.TextMatrix(i, 7) = IIf(IsNull(RSTTRXFILE!REF_NO), "", RSTTRXFILE!REF_NO)
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing

            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    End Select
End Sub

'Private Sub TMPDELETE_Click()
'    If GRDTranx.Rows = 1 Then Exit Sub
'    If MsgBox("Are You Sure You want to Delete PRINT_BILL NO." & "*** " & GRDTranx.TextMatrix(GRDTranx.Row, 2) & " ****", vbYesNo, "DELETING BILL....") = vbNo Then Exit Sub
'    DB.Execute ("DELETE from SALESREG WHERE VCH_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 2) & " AND (TRX_TYPE='GI' OR TRX_TYPE='SI'  TRX_TYPE='SI')")
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
    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
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
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!Address & ", " & RSTCOMPANY!HO_NAME, 140)
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(48) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(48) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(48) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "SALES SUMMARY FOR THE PERIOD"
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "FROM " & DTFROM.Value & " TO " & DTTo.Value
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Set RSTTRXFILE = New ADODB.Recordset
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(13) & RepeatString("-", 59)
    Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SN", 3) & Space(2) & _
            AlignLeft("INV DATE", 8) & Space(10) & _
            AlignLeft("INV AMT", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Print #1, Space(12) & RepeatString("-", 59)
    SN = 0
    RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        SN = SN + 1
        Print #1, Chr(27) & Chr(71) & Space(5) & Chr(14) & Chr(15) & AlignRight(str(SN), 4) & ". " & Space(1) & _
            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
            AlignRight(Format(Round(RSTTRXFILE!VCH_AMOUNT, 0), "0.00"), 16)
        'Print #1, Chr(13)
        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
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
    'MsgBox "Report file generated at " & Rptpath & "Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Sub

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Function ReportREGISTER()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    Dim rstSUBfile As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim SN As Integer
    Dim TRXTOTAL As Double
    
    SN = 0
    TRXTOTAL = 0
    '//NOTE : Report file name should never contain blank space.
    db.Execute "delete From SALESREG2"
    
    On Error GoTo CLOSEFILE
    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
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
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!Address & ", " & RSTCOMPANY!HO_NAME, 140)
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(48) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(48) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(48) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "SALES REGSITER FOR THE PERIOD"
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "FROM " & DTFROM.Value & " TO " & DTTo.Value
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Set RSTTRXFILE = New ADODB.Recordset
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(13) & RepeatString("-", 59)
    Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SN", 3) & Space(2) & _
            AlignLeft("INV DATE", 8) & Space(10) & _
            AlignLeft("INV AMT", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Print #1, Space(12) & RepeatString("-", 59)
    SN = 0
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG2", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTSALEREG.Properties("Update Criteria").Value = adCriteriaKey
    'RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_NO", DB, adOpenStatic,adLockReadOnly
    RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE='SV' OR TRX_TYPE='SI' OR TRX_TYPE='RI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        SN = SN + 1
        CMDDISPLAY.Tag = ""
        If RSTTRXFILE!SLSM_CODE = "A" Then
            CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(Round((RSTTRXFILE!DISCOUNT * RSTTRXFILE!VCH_AMOUNT) / 100, 2), "0.00"))
        ElseIf RSTTRXFILE!SLSM_CODE = "P" Then
            CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(RSTTRXFILE!DISCOUNT, "0.00"))
        End If
        cmdview.Tag = ""
        cmdview.Tag = IIf(IsNull(RSTTRXFILE!ADD_AMOUNT), "", RSTTRXFILE!ADD_AMOUNT)
        'SLIPAMT = SLIPAMT + RSTTRXFILE!VCH_AMOUNT - (Val(CMDDISPLAY.Tag) + Val(cmdview.Tag))
        Print #1, Chr(27) & Chr(71) & Space(5) & Chr(14) & Chr(15) & AlignRight(str(SN), 4) & ". " & Space(1) & _
            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
            AlignRight(Format(Round(RSTTRXFILE!VCH_AMOUNT - (Val(CMDDISPLAY.Tag) + Val(cmdview.Tag)), 0), "0.00"), 16)
        'Print #1, Chr(13)
        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
        
        RSTSALEREG.AddNew
        RSTSALEREG!VCH_NO = RSTTRXFILE!VCH_NO
        RSTSALEREG!TRX_TYPE = "SI"
        RSTSALEREG!VCH_DATE = RSTTRXFILE!VCH_DATE
        RSTSALEREG!VCH_AMOUNT = RSTTRXFILE!VCH_AMOUNT
        RSTSALEREG!PAYAMOUNT = 0 ' TRXFILE!PAY_AMOUNT
        RSTSALEREG!ACT_NAME = "Sales"
        RSTSALEREG!ACT_CODE = "111001"
        RSTSALEREG!DISCOUNT = 0 'rstTRANX!DISCOUNT
        RSTSALEREG.Update
        
        RSTTRXFILE.MoveNext
    Loop
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
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
    'MsgBox "Report file generated at " & Rptpath & "Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub LBLTOTAL_DblClick(index As Integer)
    If CmdCunterSales.Visible = True Then
        CmdCunterSales.Visible = False
        CmdCounterReg.Visible = False
    Else
        CmdCunterSales.Visible = True
        CmdCounterReg.Visible = True
    End If

    If frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4" Then Exit Sub
    If OptPetty.Visible = True Then
        OptPetty.Visible = False
        CmdCunterSales.Visible = False
        CmdCounterReg.Visible = False
        'OptAll.Visible = False
    Else
        OptPetty.Visible = True
        CmdCunterSales.Visible = True
        CmdCounterReg.Visible = True
        'OptAll.Visible = True
    End If
End Sub

Private Sub OPTCUSTOMER_Click()
    TXTDEALER.SetFocus
End Sub

Private Sub OPTCUSTOMER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub

Private Sub TxtAgent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCustomercode_Change()
    If frmLogin.rs!Level <> "0" Then Exit Sub
    If txtCustomercode.Text = "*101*" Then
        cmdwoprint.Visible = True
        chkunbill.Value = 1
        chkunbill.Visible = True
    Else
        cmdwoprint.Visible = False
        chkunbill.Value = 1
        chkunbill.Visible = False
    End If
End Sub

Private Sub txtCustomercode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub txtCustomercode_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCustomerName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmDDisplay_Click
        
    End Select
End Sub

Private Sub txtCustomerName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_GotFocus()
    OPTCUSTOMER.Value = True
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
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
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.Text = DataList2.Text
    GRDTranx.rows = 1
    LBLTRXTOTAL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
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
    TXTDEALER.Text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub TXTDEALER2_Change()
    
    On Error GoTo ERRHAND
    If flagchange2.Caption <> "1" Then
        If PHY_FLAG = True Then
            PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            PHY_FLAG = False
        Else
            PHY_REC.Close
            PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            PHY_FLAG = False
        End If
        If (PHY_REC.EOF And PHY_REC.BOF) Then
            LBLDEALER2.Caption = ""
        Else
            LBLDEALER2.Caption = PHY_REC!MANUFACTURER
        End If
        Set Me.DataList1.RowSource = PHY_REC
        DataList1.ListField = "MANUFACTURER"
        DataList1.BoundColumn = "MANUFACTURER"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.Text)
End Sub

Private Sub TXTDEALER2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList1.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList1.SetFocus
    End Select

End Sub

Private Sub TXTDEALER2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_Click()
        
    TXTDEALER2.Text = DataList1.Text
    LBLDEALER2.Caption = TXTDEALER2.Text

End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER2.Text) = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Category From List", vbOKOnly, "Category List..."
                DataList1.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            TXTDEALER2.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_GotFocus()
    flagchange2.Caption = 1
    TXTDEALER2.Text = LBLDEALER2.Caption
    DataList1.Text = TXTDEALER2.Text
    Call DataList1_Click
End Sub

Private Sub DataList1_LostFocus()
     flagchange2.Caption = ""
End Sub

'Private Sub TXTDEALER3_Change()
'
'    On Error GoTo ErrHand
'    If flagchange3.Caption <> "1" Then
'        If CAT_FLAG = True Then
'            CAT_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER3.text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
'            CAT_FLAG = False
'        Else
'            CAT_REC.Close
'            CAT_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER3.text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
'            CAT_FLAG = False
'        End If
'        If (CAT_REC.EOF And CAT_REC.BOF) Then
'            LBLDEALER3.Caption = ""
'        Else
'            LBLDEALER3.Caption = CAT_REC!Category
'        End If
'        Set Me.DataList3.RowSource = CAT_REC
'        DataList3.ListField = "CATEGORY"
'        DataList3.BoundColumn = "CATEGORY"
'    End If
'    Exit Sub
'ErrHand:
'    MsgBox err.Description
'
'End Sub
'
'Private Sub TXTDEALER3_GotFocus()
'    TXTDEALER3.SelStart = 0
'    TXTDEALER3.SelLength = Len(TXTDEALER3.text)
'End Sub
'
'Private Sub TXTDEALER3_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyReturn, 40
'            If DataList3.VisibleCount = 0 Then Exit Sub
'            'lbladdress.Caption = ""
'            DataList3.SetFocus
'    End Select
'
'End Sub
'
'Private Sub TXTDEALER3_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
'            KeyAscii = 0
'    End Select
'End Sub
'
'Private Sub DataList3_Click()
'
'    TXTDEALER3.text = DataList3.text
'    LBLDEALER3.Caption = TXTDEALER3.text
'
'End Sub
'
'Private Sub DataList3_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyReturn
'            If Trim(TXTDEALER3.text) = "" Then Exit Sub
'            If IsNull(DataList3.SelectedItem) Then
'                MsgBox "Select Category From List", vbOKOnly, "Category List..."
'                DataList3.SetFocus
'                Exit Sub
'            End If
'            CMDDISPLAY.SetFocus
'        Case vbKeyEscape
'            TXTDEALER3.SetFocus
'    End Select
'End Sub
'
'Private Sub DataList3_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
'            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub
'
'Private Sub DataList3_GotFocus()
'    flagchange3.Caption = 1
'    TXTDEALER3.text = LBLDEALER3.Caption
'    DataList3.text = TXTDEALER3.text
'    Call DataList3_Click
'End Sub
'
'Private Sub DataList3_LostFocus()
'     flagchange3.Caption = ""
'End Sub

Private Sub TXTDEALER4_Change()
    
    On Error GoTo ERRHAND
    If flagchange4.Caption <> "1" Then
        If AREA_FLAG = True Then
            AREA_REC.Open "Select DISTINCT AREA From CUSTMAST WHERE AREA Like '" & TXTDEALER4.Text & "%' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
            AREA_FLAG = False
        Else
            AREA_REC.Close
            AREA_REC.Open "Select DISTINCT AREA From CUSTMAST WHERE AREA Like '" & TXTDEALER4.Text & "%' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
            AREA_FLAG = False
        End If
        If (AREA_REC.EOF And AREA_REC.BOF) Then
            lbldealer4.Caption = ""
        Else
            lbldealer4.Caption = AREA_REC!Area
        End If
        Set Me.DataList4.RowSource = AREA_REC
        DataList4.ListField = "AREA"
        DataList4.BoundColumn = "AREA"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTDEALER4_GotFocus()
    TXTDEALER4.SelStart = 0
    TXTDEALER4.SelLength = Len(TXTDEALER4.Text)
End Sub

Private Sub TXTDEALER4_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList4.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList4.SetFocus
    End Select

End Sub

Private Sub TXTDEALER4_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList4_Click()
        
    TXTDEALER4.Text = DataList4.Text
    lbldealer4.Caption = TXTDEALER4.Text

End Sub

Private Sub DataList4_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER4.Text) = "" Then Exit Sub
            If IsNull(DataList4.SelectedItem) Then
                MsgBox "Select Area From List", vbOKOnly, "Area List..."
                DataList4.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            TXTDEALER4.SetFocus
    End Select
End Sub

Private Sub DataList4_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList4_GotFocus()
    flagchange4.Caption = 1
    TXTDEALER4.Text = lbldealer4.Caption
    DataList4.Text = TXTDEALER4.Text
    Call DataList4_Click
End Sub

Private Sub DataList4_LostFocus()
     flagchange4.Caption = ""
End Sub


Private Function fillcombo()
    txtCustomerName.Clear
    On Error GoTo ERRHAND
    Dim rstfillcombo As ADODB.Recordset
    Set rstfillcombo = New ADODB.Recordset
    rstfillcombo.Open "Select DISTINCT BILL_NAME From TRXMAST ORDER BY BILL_NAME", db, adOpenStatic, adLockReadOnly
    Do Until rstfillcombo.EOF
        If Not IsNull(rstfillcombo!BILL_NAME) Then txtCustomerName.AddItem (rstfillcombo!BILL_NAME)
        rstfillcombo.MoveNext
    Loop
    rstfillcombo.Close
    Set rstfillcombo = Nothing
    
    LstCategory.Clear
    Set rstfillcombo = New ADODB.Recordset
    rstfillcombo.Open "select * from category where category <> '' and not isnull(category) ORDER BY category", db, adOpenForwardOnly
    Do Until rstfillcombo.EOF
        LstCategory.AddItem Trim(rstfillcombo!Category)
        rstfillcombo.MoveNext
    Loop
    rstfillcombo.Close
    Set rstfillcombo = Nothing

    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub Report_Generate(search_type As String)
    Dim i As Integer
    
    On Error GoTo ERRHAND
    If search_type = "M" Then
        ReportNameVar = Rptpath & "RPTMONWISE"
    ElseIf search_type = "D" Then
        ReportNameVar = Rptpath & "RPTDAILYWISE"
    ElseIf search_type = "Y" Then
        ReportNameVar = Rptpath & "RPTYEARWISE"
    Else
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If chkunbill.Value = 0 Or chkunbill.Visible = False Then
        Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE} = 'SV' OR {TRXMAST.TRX_TYPE} = 'HI' OR {TRXMAST.TRX_TYPE} = 'GI' OR {TRXMAST.TRX_TYPE} = 'VI' OR {TRXMAST.TRX_TYPE} = 'RI' OR {TRXMAST.TRX_TYPE} = 'SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE} = 'SV' OR {TRXMAST.TRX_TYPE} = 'HI' OR {TRXMAST.TRX_TYPE} = 'GI' OR {TRXMAST.TRX_TYPE} = 'WO' OR {TRXMAST.TRX_TYPE} = 'VI' OR {TRXMAST.TRX_TYPE} = 'RI' OR {TRXMAST.TRX_TYPE} = 'SI') AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    End If
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            'Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            If Report.Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub TXTDEALER5_Change()
    
    On Error GoTo ERRHAND
    If flagchange5.Caption <> "1" Then
        If AGNT_FLAG = True Then
            AGNT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='911')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            AGNT_FLAG = False
        Else
            AGNT_REC.Close
            AGNT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='911')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            AGNT_FLAG = False
        End If
        If (AGNT_REC.EOF And AGNT_REC.BOF) Then
            lbldealer5.Caption = ""
        Else
            lbldealer5.Caption = AGNT_REC!ACT_NAME
        End If
        Set Me.DataList5.RowSource = AGNT_REC
        DataList5.ListField = "ACT_NAME"
        DataList5.BoundColumn = "ACT_CODE"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTDEALER5_GotFocus()
    TXTDEALER5.SelStart = 0
    TXTDEALER5.SelLength = Len(TXTDEALER5.Text)
End Sub

Private Sub TXTDEALER5_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList5.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList5.SetFocus
    End Select

End Sub

Private Sub TXTDEALER5_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList5_Click()
        
    TXTDEALER5.Text = DataList5.Text
    lbldealer5.Caption = TXTDEALER5.Text

End Sub

Private Sub DataList5_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER5.Text) = "" Then Exit Sub
            If IsNull(DataList5.SelectedItem) Then
                MsgBox "Select Category From List", vbOKOnly, "Category List..."
                DataList5.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            TXTDEALER5.SetFocus
    End Select
End Sub

Private Sub DataList5_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList5_GotFocus()
    flagchange5.Caption = 1
    TXTDEALER5.Text = lbldealer5.Caption
    DataList5.Text = TXTDEALER5.Text
    Call DataList5_Click
End Sub

Private Sub DataList5_LostFocus()
     flagchange5.Caption = ""
End Sub

