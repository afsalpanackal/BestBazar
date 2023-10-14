VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMRAWMIX2 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROCESS"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12420
   Icon            =   "FRMRAWMIX2.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12420
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   -135
      TabIndex        =   9
      Top             =   -15
      Width           =   12555
      Begin VB.Frame Frame1 
         Caption         =   "Final Products"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   210
         TabIndex        =   66
         Top             =   4350
         Width           =   12330
         Begin VB.TextBox TXTsample2 
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
            Height          =   315
            Left            =   45
            TabIndex        =   68
            Top             =   300
            Visible         =   0   'False
            Width           =   1110
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales2 
            Height          =   1770
            Left            =   30
            TabIndex        =   67
            Top             =   210
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   3122
            _Version        =   393216
            Rows            =   1
            Cols            =   11
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
            AllowUserResizing=   1
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
      End
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
         TabIndex        =   31
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Used Products"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3450
         Left            =   210
         TabIndex        =   16
         Top             =   945
         Width           =   12345
         Begin VB.TextBox TXTsample 
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
            Height          =   315
            Left            =   4650
            TabIndex        =   58
            Top             =   1725
            Visible         =   0   'False
            Width           =   1110
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   3210
            Left            =   30
            TabIndex        =   8
            Top             =   180
            Width           =   12285
            _ExtentX        =   21669
            _ExtentY        =   5662
            _Version        =   393216
            Rows            =   1
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
            AllowUserResizing=   1
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
      End
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   465
         Left            =   12630
         TabIndex        =   30
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
         BackColor       =   &H00F2DBFB&
         Height          =   2475
         Left            =   210
         TabIndex        =   17
         Top             =   6270
         Width           =   10575
         Begin VB.TextBox TxtPack 
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
            Left            =   5985
            MaxLength       =   7
            TabIndex        =   64
            Top             =   1140
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtBarcode 
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
            Left            =   5535
            MaxLength       =   7
            TabIndex        =   59
            Top             =   1515
            Visible         =   0   'False
            Width           =   3105
         End
         Begin VB.CommandButton CmDCancel 
            Caption         =   "&Cancel"
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
            Left            =   5655
            TabIndex        =   57
            Top             =   1920
            Width           =   1125
         End
         Begin VB.TextBox TxttaxMRP 
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
            Left            =   9825
            MaxLength       =   7
            TabIndex        =   54
            Top             =   1140
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox TXTRETAIL 
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
            Left            =   6765
            MaxLength       =   7
            TabIndex        =   50
            Top             =   1140
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtWS 
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
            Height          =   390
            Left            =   7755
            MaxLength       =   7
            TabIndex        =   49
            Top             =   1110
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox txtvanrate 
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
            Height          =   375
            Left            =   8790
            MaxLength       =   7
            TabIndex        =   48
            Top             =   1125
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox TxtResult 
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
            Left            =   4530
            MaxLength       =   7
            TabIndex        =   41
            Top             =   1140
            Visible         =   0   'False
            Width           =   960
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
            TabIndex        =   38
            Top             =   2460
            Visible         =   0   'False
            Width           =   795
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
            TabIndex        =   37
            Top             =   1095
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox TXTPRODUCT2 
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
            Height          =   390
            Left            =   75
            TabIndex        =   0
            Top             =   450
            Width           =   4440
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
            Height          =   390
            Left            =   4530
            MaxLength       =   7
            TabIndex        =   1
            Top             =   450
            Width           =   1245
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
            Height          =   450
            Left            =   3255
            TabIndex        =   4
            Top             =   1920
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
            Height          =   450
            Left            =   7005
            TabIndex        =   6
            Top             =   1920
            Width           =   1140
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
            Height          =   450
            Left            =   2055
            TabIndex        =   3
            Top             =   1920
            Width           =   1125
         End
         Begin VB.CommandButton CmdDelete 
            Caption         =   "&Delete"
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
            Left            =   825
            TabIndex        =   2
            Top             =   1920
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
            Left            =   8700
            TabIndex        =   22
            Top             =   1515
            Visible         =   0   'False
            Width           =   690
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
            TabIndex        =   21
            Top             =   2775
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
            TabIndex        =   20
            Top             =   2775
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox TxtActqty 
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
            Left            =   1185
            TabIndex        =   19
            Top             =   2505
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
            TabIndex        =   18
            Top             =   1800
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
            Height          =   450
            Left            =   4455
            TabIndex        =   5
            Top             =   1920
            Width           =   1125
         End
         Begin MSDataListLib.DataList DataList1 
            Height          =   1035
            Left            =   75
            TabIndex        =   43
            Top             =   855
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   1826
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
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   5
            Left            =   5985
            TabIndex        =   65
            Top             =   870
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label lblpack 
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
            Height          =   360
            Left            =   5505
            TabIndex        =   56
            Top             =   1140
            Visible         =   0   'False
            Width           =   465
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
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   12
            Left            =   9825
            TabIndex        =   55
            Top             =   870
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Retail"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   24
            Left            =   6765
            TabIndex        =   53
            Top             =   870
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "W.Sale"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   27
            Left            =   7755
            TabIndex        =   52
            Top             =   870
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Van Rate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   32
            Left            =   8790
            TabIndex        =   51
            Top             =   870
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lBLpRODUCT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   5805
            TabIndex        =   47
            Top             =   465
            Visible         =   0   'False
            Width           =   4725
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Output Product"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   300
            Index           =   4
            Left            =   5805
            TabIndex        =   46
            Top             =   150
            Visible         =   0   'False
            Width           =   4725
         End
         Begin VB.Label flagchange2 
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   7710
            TabIndex        =   45
            Top             =   915
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbldealer2 
            BackColor       =   &H00FAF2F1&
            Height          =   315
            Left            =   8370
            TabIndex        =   44
            Top             =   1800
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Output Qty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   3
            Left            =   4530
            TabIndex        =   42
            Top             =   870
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Mixture Name"
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
            Height          =   285
            Index           =   9
            Left            =   75
            TabIndex        =   29
            Top             =   150
            Width           =   4440
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "No. of Mix"
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
            Height          =   285
            Index           =   10
            Left            =   4530
            TabIndex        =   28
            Top             =   150
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Barcode"
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
            Height          =   330
            Index           =   15
            Left            =   4545
            TabIndex        =   27
            Top             =   1530
            Visible         =   0   'False
            Width           =   960
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
            TabIndex        =   26
            Top             =   2790
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Act. Qty"
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
            Left            =   45
            TabIndex        =   25
            Top             =   2520
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
            TabIndex        =   24
            Top             =   2190
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
            TabIndex        =   23
            Top             =   2790
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00F2DBFB&
         Height          =   990
         Left            =   210
         TabIndex        =   10
         Top             =   -30
         Width           =   12345
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9495
            ScaleHeight     =   240
            ScaleWidth      =   1335
            TabIndex        =   63
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9510
            ScaleHeight     =   240
            ScaleWidth      =   855
            TabIndex        =   62
            Top             =   870
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7680
            ScaleHeight     =   240
            ScaleWidth      =   1965
            TabIndex        =   61
            Top             =   600
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7680
            ScaleHeight     =   240
            ScaleWidth      =   1800
            TabIndex        =   60
            Top             =   870
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox TXTREMARKS 
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
            Left            =   3990
            MaxLength       =   100
            TabIndex        =   39
            Top             =   555
            Width           =   2670
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
            Left            =   1590
            TabIndex        =   7
            Top             =   150
            Visible         =   0   'False
            Width           =   885
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   345
            Left            =   1545
            TabIndex        =   32
            Top             =   570
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
            Caption         =   "REMARKS"
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
            Left            =   3045
            TabIndex        =   40
            Top             =   600
            Width           =   900
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Prodn Date"
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
            Left            =   105
            TabIndex        =   33
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Production No."
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
            TabIndex        =   15
            Top             =   165
            Width           =   1440
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
            Left            =   3390
            TabIndex        =   14
            Top             =   180
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
            Left            =   3990
            TabIndex        =   13
            Top             =   150
            Width           =   1335
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
            Left            =   5445
            TabIndex        =   12
            Top             =   150
            Width           =   1230
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
            Left            =   1590
            TabIndex        =   11
            Top             =   150
            Width           =   900
         End
      End
   End
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   -15
      TabIndex        =   36
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11355
      TabIndex        =   35
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   34
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMRAWMIX2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim MIX_ITEM As New ADODB.Recordset
Dim MIX_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_STOCK As Integer
Dim M_EDIT As Boolean
Dim EDIT_INV, OLD_INV As Boolean

Private Sub CMDADD_Click()
    Dim rststock As ADODB.Recordset
    'Dim RSTMINQTY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Long
    
    If DataList1.BoundText = "" Then
        MsgBox "Select Mixture from the List", , "PROCESS"
        TXTPRODUCT2.SetFocus
        Exit Sub
    End If
    
    If Val(TXTQTY.text) = 0 Then
        MsgBox "Enter the number of mixture", , "PROCESS"
        TXTQTY.SetFocus
        Exit Sub
    End If
    
    If Val(TxtResult.text) = 0 Then
        MsgBox "Output Qty could not be Zero", , "PROCESS"
        TxtResult.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    grdsales.FixedRows = 0
    grdsales.rows = 1
        
    i = 1
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM  TRXFORMULASUB WHERE FOR_NO = " & DataList1.BoundText & " AND TRX_TYPE='PR' ", db, adOpenStatic, adLockReadOnly, adCmdText
    With RSTITEMMAST
        Do Until .EOF
            grdsales.rows = grdsales.rows + 1
            grdsales.FixedRows = 1
            grdsales.TextMatrix(i, 0) = i
            grdsales.TextMatrix(i, 1) = RSTITEMMAST!ITEM_NAME
            grdsales.TextMatrix(i, 4) = RSTITEMMAST!ITEM_CODE
            If UCase(RSTITEMMAST!Category) = "SERVICE CHARGE" Then
                grdsales.TextMatrix(i, 2) = ""
                grdsales.TextMatrix(i, 3) = ""
                grdsales.TextMatrix(i, 5) = ""
                grdsales.TextMatrix(i, 6) = ""
                grdsales.TextMatrix(i, 7) = RSTITEMMAST!QTY
            Else
                grdsales.TextMatrix(i, 3) = 1
                grdsales.TextMatrix(i, 5) = IIf(IsNull(RSTITEMMAST!LOOSE_PACK), "1", RSTITEMMAST!LOOSE_PACK)
                grdsales.TextMatrix(i, 2) = RSTITEMMAST!QTY * Val(grdsales.TextMatrix(i, 5))
                grdsales.TextMatrix(i, 6) = IIf(IsNull(RSTITEMMAST!PACK_TYPE), "", RSTITEMMAST!PACK_TYPE)
            End If
            grdsales.TextMatrix(i, 8) = IIf(IsNull(RSTITEMMAST!Category), "", RSTITEMMAST!Category)
            grdsales.TextMatrix(i, 9) = IIf(IsNull(RSTITEMMAST!WASTE_QTY), 0, RSTITEMMAST!WASTE_QTY) * Val(TXTQTY)
            i = i + 1
            .MoveNext
        Loop
    End With
    Set RSTITEMMAST = Nothing
    
    grdsales2.FixedRows = 0
    grdsales2.rows = 1
    i = 1
    Dim rstformula As ADODB.Recordset
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM  TRXFORMULAMAST WHERE FOR_NO = " & DataList1.BoundText & " AND TRX_TYPE='PR' ", db, adOpenStatic, adLockReadOnly, adCmdText
    With RSTITEMMAST
        Do Until .EOF
            grdsales2.rows = grdsales2.rows + 1
            grdsales2.FixedRows = 1
            grdsales2.TextMatrix(i, 0) = i
        
            grdsales2.TextMatrix(i, 1) = IIf(IsNull(RSTITEMMAST!ITEM_CODE), "", RSTITEMMAST!ITEM_CODE)
            grdsales2.TextMatrix(i, 2) = IIf(IsNull(RSTITEMMAST!ITEM_NAME), "", RSTITEMMAST!ITEM_NAME)
            grdsales2.TextMatrix(i, 3) = Val(TxtResult.text) 'IIf(IsNull(RSTITEMMAST!QTY), "1", RSTITEMMAST!QTY) * Val(TxtResult.Text)
            grdsales2.TextMatrix(i, 4) = IIf(IsNull(RSTITEMMAST!LOOSE_PACK), "", RSTITEMMAST!LOOSE_PACK)
            grdsales2.TextMatrix(i, 5) = IIf(IsNull(RSTITEMMAST!PACK_TYPE), "", RSTITEMMAST!PACK_TYPE)
                
            Set rstformula = New ADODB.Recordset
            rstformula.Open "select * from ITEMMAST where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstformula.EOF Or rstformula.BOF) Then
                grdsales2.TextMatrix(i, 6) = IIf(IsNull(rstformula!P_RETAIL), "", rstformula!P_RETAIL)
                grdsales2.TextMatrix(i, 7) = IIf(IsNull(rstformula!P_WS), "", rstformula!P_WS)
                grdsales2.TextMatrix(i, 8) = IIf(IsNull(rstformula!P_VAN), "", rstformula!P_VAN)
                grdsales2.TextMatrix(i, 9) = IIf(IsNull(rstformula!SALES_TAX), "", rstformula!SALES_TAX)
                grdsales2.TextMatrix(i, 10) = IIf(IsNull(rstformula!LOOSE_PACK), "1", rstformula!LOOSE_PACK)
            End If
            rstformula.Close
            Set rstformula = Nothing
    
            i = i + 1
            .MoveNext
        Loop
    End With
    Set RSTITEMMAST = Nothing
    
    cmdadd.Enabled = False
    ''CmdDelete.Enabled = False
    cmdexit.Enabled = False
    M_EDIT = False
    'Call COSTCALCULATION
    'grdsales.TopRow = grdsales.Rows - 1

    cmdRefresh.Enabled = True
Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If grdsales.rows <= 1 Then Exit Sub
            If grdsales2.rows <= 1 Then Exit Sub
            cmdadd.Enabled = False
            cmdRefresh.Enabled = True
            cmdRefresh.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub cmdcancel_Click()
    If MsgBox("Are you sure you want to cancel?", vbYesNo, "PROCESS") = vbNo Then Exit Sub
    Call cancel_bill
End Sub

Private Function cancel_bill()
    On Error GoTo ErrHand
    Dim rstBILL As ADODB.Recordset
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PC'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
        
    TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
    LBLDATE.Caption = Date
    lbltime.Caption = Time
    'LBLTOTALCOST.Caption = ""
    grdsales.rows = 1
    grdsales2.rows = 1
    M_EDIT = False
    EDIT_INV = False
    TXTPRODUCT2.text = ""
    TXTQTY.text = ""
    lBLpRODUCT.Caption = ""
    TXTITEMCODE.text = ""
    TxtResult.text = ""
    TxtBarcode.text = ""
    LblPack.Caption = ""
    txtretail.text = ""
    Txtpack.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    TxttaxMRP.text = ""
    
    cmdRefresh.Enabled = False
    cmdexit.Enabled = True
    CMDPRINT.Enabled = False
    cmdexit.Enabled = True
    FRMEHEAD.Enabled = True
    OLD_INV = False
    TXTPRODUCT2.SetFocus
    'LBLITEMCOST.Caption = ""
    TXTQTY.Tag = ""
    Exit Function
ErrHand:
    MsgBox err.Description
End Function

Private Sub CmdDelete_Click()
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstTRXMAST As ADODB.Recordset

    If grdsales.rows <= 1 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE THE ENTIRE PRODUCTION", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    If OLD_INV = False Then
        Call cancel_bill
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'    With RSTTRXFILE
'        If Not (.EOF And .BOF) Then
'            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
'            !ISSUE_QTY = !ISSUE_QTY + Val(TxtResult.Text)
'            !FREE_QTY = 0
'            !ISSUE_VAL = 0
'            !CLOSE_QTY = !CLOSE_QTY - Val(TxtResult.Text)
'            !CLOSE_VAL = 0
'            RSTTRXFILE.Update
'        End If
'    End With
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing

    For i = 1 To grdsales.rows - 1
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        With RSTTRXFILE
'            If Not (.EOF And .BOF) Then
'                !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(i, 2)) '/ Val(grdsales.TextMatrix(i, 5)))
'                !FREE_QTY = 0
'                !ISSUE_VAL = 0
'                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(i, 2)) '/ Val(grdsales.TextMatrix(i, 5)))
'                !CLOSE_VAL = 0
'                !LOOSE_PACK = Val(grdsales.TextMatrix(i, 5))
'                !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
'                RSTTRXFILE.Update
'            End If
'        End With
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
                
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                !LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
                !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
                !PD_NO = Val(txtBillNo.text)
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    Next i
    
    db.Execute "delete From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete From TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete FROM RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & ""
    
    Dim RSTITEMMAST, rststock As ADODB.Recordset
    Dim INWARD, OUTWARD As Double
    INWARD = 0
    OUTWARD = 0
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & TXTITEMCODE.text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        INWARD = 0
        OUTWARD = 0
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Do Until rststock.EOF
            INWARD = INWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY)) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
            
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='PC' OR TRX_TYPE='RM' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockOptimistic, adCmdText
        'rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Do Until rststock.EOF
            OUTWARD = OUTWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)) + IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
            OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        
        RSTITEMMAST!CLOSE_QTY = INWARD - OUTWARD
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    For i = 1 To grdsales.rows - 1
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTITEMMAST.EOF
            INWARD = 0
            OUTWARD = 0
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            Do Until rststock.EOF
                INWARD = INWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY)) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            Set rststock = New ADODB.Recordset
            'rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='PC' OR TRX_TYPE='RM' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until rststock.EOF
                OUTWARD = OUTWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)) + IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
                OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            RSTITEMMAST!CLOSE_QTY = INWARD - OUTWARD
            RSTITEMMAST.Update
            RSTITEMMAST.MoveNext
        Loop
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    Next i
    
    Call cancel_bill
    Screen.MousePointer = vbNormal
    Exit Sub
'    For i = Val(TXTSLNO.Text) - 1 To grdsales.Rows - 2
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = i
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = grdsales.TextMatrix(i + 1, 1)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = grdsales.TextMatrix(i + 1, 2)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = grdsales.TextMatrix(i + 1, 3)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = grdsales.TextMatrix(i + 1, 4)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(i + 1, 6)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = grdsales.TextMatrix(i + 1, 5)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = grdsales.TextMatrix(i + 1, 7)
'    Next i
'    grdsales.Rows = grdsales.Rows - 1
    
    'Call COSTCALCULATION
    
    TXTITEMCODE.text = ""
    TXTVCHNO.text = ""
    TXTLINENO.text = ""
    TXTUNIT.text = ""
    TXTQTY.text = ""
    cmdadd.Enabled = False
    'CmdDelete.Enabled = False
    cmdexit.Enabled = False
    M_EDIT = False
    EDIT_INV = True
    If grdsales.rows = 1 Then
'        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    
    If grdsales.rows = 1 Then Exit Sub
    
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
    End If
    Call Generateprint
    
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    Dim Num As Currency
    
    On Error GoTo ErrHand
    
    db.Execute "delete From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "PC"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = "GENERAL"
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 2))
        RSTTRXFILE!WASTE_QTY = Val(grdsales.TextMatrix(i, 9))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = 0
        RSTTRXFILE!PTR = 0
        RSTTRXFILE!SALES_PRICE = 0
        RSTTRXFILE!SALES_TAX = 0
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 3)
        RSTTRXFILE!VCH_DESC = "Issued to      Press"
        RSTTRXFILE!REF_NO = ""
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = ""
        RSTTRXFILE!MFGR = ""
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = 0
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!P_RETAIL = 0
        RSTTRXFILE!P_RETAILWOTAX = 0
        RSTTRXFILE!SALE_1_FLAG = ""
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "" 'DataList2.BoundText
        
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Call ReportGeneratION
    ReportNameVar = Rptpath & "rptRAWBILL"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "( {TRXFILE.TRX_TYPE}='PC' AND {TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & " )"
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
    End If
    Set Printer = Printers(barcodeprinter)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        'If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & DataList2.Text & "'"
'        If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & lbladdress.Caption & "'"
'        If CRXFormulaField.Name = "{@DLNO2}" Then CRXFormulaField.Text = "'" & LBLDLNO2.Caption & "'"
'        If CRXFormulaField.Name = "{@DLNO}" Then CRXFormulaField.Text = "'" & lbldlno.Caption & "'"
'        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
'        If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'TAX INVOICE FORM 8H/8B/8'"
'        If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & lbltin.Caption & "'"
'        If lblcredit.Caption = "0" Then
'            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'CASH'"
'        Else
'            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'" & txtcrdays.Text & "'" & "' Days'"
'        End If
    Next
    frmreport.Caption = "BILL"
    Call GENERATEREPORT
    
    ''cmdRefresh.SetFocus
'
    
    cmdexit.Enabled = False
    'TXTQTY.Enabled = False
    
    ''rptPRINT.Action = 1
    Exit Function
ErrHand:
    MsgBox err.Description
End Function

Private Sub cmdRefresh_Click()
    
   ' If grdsales.Rows = 1 Then GoTo SKIP
    
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
    End If
    
    If DataList1.BoundText = "" Then
        MsgBox "Please select the output Product", vbOKOnly, "PROCESS"
        Exit Sub
    End If
    If Val(Txtpack.text) = 0 Then Txtpack.text = 1
    
    If grdsales.rows <= 1 Then
        MsgBox "Please add a Process", vbOKOnly, "PROCESS"
        Exit Sub
    End If
    
    If grdsales2.rows <= 1 Then
        MsgBox "Please add a Process", vbOKOnly, "PROCESS"
        Exit Sub
    End If
'    Dim i As Long
'    If MDIMAIN.StatusBar.Panels(6).Text = "Y" Then
'        If Trim(TxtBarcode.Text) = "" Then TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(txtretail.Text)
'        If MsgBox("Do you want to Print Barcode Labels", vbYesNo, "Production.....") = vbYes Then
'            i = Val(InputBox("Enter number of lables to be print", "No. of labels..", Val(TxtResult.Text)))
'            If i > 0 Then
'                If i > 0 And MDIMAIN.barcode_profile.Caption = 0 Then
'                    Call print_3labels(i, Trim(TxtBarcode.Text), "")
'                Else
'                    Call print_labels(i, Trim(TxtBarcode.Text), "")
'                End If
'            End If
'        End If
'    End If
    
    Call AppendSale
    
'    Me.Enabled = False
'    FRMDEBIT.Show
    
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
'            TXTPRODUCT.Text = ""
'            TXTQTY.Text = ""
'            TXTITEMCODE.Text = ""
'            TXTVCHNO.Text = ""
'            TXTLINENO.Text = ""
'            TXTTRXTYPE.Text = ""
'            TXTUNIT.Text = ""
'            TXTPRODUCT.SetFocus
'            TXTQTY.Enabled = False
'            CMDMODIFY.Enabled = False
'            'CmdDelete.Enabled = False
    End Select
End Sub

Private Sub Form_Activate()
    If TXTPRODUCT2.Enabled = True Then TXTPRODUCT2.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
    If txtremarks.Enabled = True Then txtremarks.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PC'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
        
    MIX_FLAG = True
    ACT_FLAG = True
    LBLDATE.Caption = Date
    lbltime.Caption = Time
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 3600
    grdsales.ColWidth(2) = 1000
    'grdsales.ColWidth(3) = 0
    grdsales.ColWidth(4) = 0
    grdsales.ColWidth(5) = 0
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM NAME"
    grdsales.TextArray(2) = "QTY"
    grdsales.TextArray(3) = "PACK"
    grdsales.TextArray(4) = "" '"ITEM CODE"
    grdsales.TextArray(5) = "" '"Loose Pack"
    grdsales.TextArray(6) = "Pack Type"
    grdsales.TextArray(7) = "Amount"
    grdsales.TextArray(8) = "Category"
    grdsales.TextArray(9) = "Waste"
    
    
    grdsales2.ColWidth(0) = 400
    grdsales2.ColWidth(1) = 0
    grdsales2.ColWidth(2) = 2500
    grdsales2.ColWidth(3) = 1100
    grdsales2.ColWidth(4) = 1000
    grdsales2.ColWidth(5) = 1500
    grdsales2.ColWidth(6) = 1000
    grdsales2.ColWidth(7) = 1000
    grdsales2.ColWidth(8) = 1000
    grdsales2.ColWidth(9) = 1000
    grdsales2.ColWidth(10) = 1000
    
    'grdsales.ColWidth(3) = 0
    'grdsales.ColWidth(4) = 0
    'grdsales.ColWidth(5) = 0
    
    grdsales2.TextArray(0) = "Sl"
    grdsales2.TextArray(1) = "Item Code"
    grdsales2.TextArray(2) = "Item Name"
    grdsales2.TextArray(3) = "Qty"
    grdsales2.TextArray(4) = "Pack"
    grdsales2.TextArray(5) = "UOM"
    grdsales2.TextArray(6) = "RT"
    grdsales2.TextArray(7) = "WS"
    grdsales2.TextArray(8) = "VN"
    grdsales2.TextArray(9) = "TAX"
    grdsales2.TextArray(10) = "Pack"
    
    'TXTQTY.Enabled = False
    'CmdDelete.Enabled = False
    CMDPRINT.Enabled = False
    
    CLOSEALL = 1
    M_EDIT = False
    EDIT_INV = False
'    Me.Width = 11700
'    Me.Height = 10185
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If MIX_FLAG = False Then MIX_ITEM.Close
        If ACT_FLAG = False Then ACT_REC.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdsales_Click()
    On Error Resume Next
    TXTsample.Visible = False
    TXTsample2.Visible = False
    grdsales.SetFocus
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If grdsales.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            'If OLD_INV = True Then Exit Sub
            Select Case grdsales.Col
                Case 2
                    TXTsample.Visible = True
                    TXTsample.Top = grdsales.CellTop + 130
                    TXTsample.Left = grdsales.CellLeft + 50
                    TXTsample.Width = grdsales.CellWidth
                    TXTsample.Height = grdsales.CellHeight
                    TXTsample.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TXTsample.SetFocus
                Case 9
                    TXTsample.Visible = True
                    TXTsample.Top = grdsales.CellTop + 130
                    TXTsample.Left = grdsales.CellLeft + 50
                    TXTsample.Width = grdsales.CellWidth
                    TXTsample.Height = grdsales.CellHeight
                    TXTsample.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TXTsample.SetFocus
            End Select
    End Select
End Sub

Private Sub grdsales_Scroll()
    TXTsample.Visible = False
    TXTsample2.Visible = False
    grdsales.SetFocus
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim TRXMAST As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    
    Dim i As Long
    Dim n As Integer
    Dim M As Integer

    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            cmdexit.Enabled = True
            OLD_INV = False
            If Val(txtBillNo.text) = 0 Then Exit Sub
            grdsales.rows = 1
            grdsales2.rows = 1
            i = 0
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "Select * From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until TRXMAST.EOF
                i = i + 1
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = TRXMAST!ITEM_NAME
                grdsales.TextMatrix(i, 4) = TRXMAST!ITEM_CODE
                grdsales.TextMatrix(i, 8) = IIf(IsNull(TRXMAST!Category), "", TRXMAST!Category)
                grdsales.TextMatrix(i, 9) = IIf(IsNull(TRXMAST!WASTE_QTY), 0, TRXMAST!WASTE_QTY)
                If UCase(grdsales.TextMatrix(i, 8)) <> "SERVICE CHARGE" Then
                    grdsales.TextMatrix(i, 2) = TRXMAST!QTY
                    grdsales.TextMatrix(i, 3) = 1 'TRXMAST!UNIT
                    grdsales.TextMatrix(i, 5) = IIf(IsNull(TRXMAST!LOOSE_PACK), 1, TRXMAST!LOOSE_PACK)
                    grdsales.TextMatrix(i, 6) = TRXMAST!PACK_TYPE
                Else
                    grdsales.TextMatrix(i, 7) = IIf(IsNull(TRXMAST!ITEM_COST), 0, TRXMAST!ITEM_COST)
                End If
                TRXMAST.MoveNext
            Loop
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            i = 0
            Dim rstformula As ADODB.Recordset
            Dim RSTTRXFILE As ADODB.Recordset
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * FROM RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until RSTTRXFILE.EOF
                i = i + 1
                grdsales2.rows = grdsales2.rows + 1
                grdsales2.FixedRows = 1
                grdsales2.TextMatrix(i, 0) = i
                grdsales2.TextMatrix(i, 1) = IIf(IsNull(RSTTRXFILE!ITEM_CODE), "", RSTTRXFILE!ITEM_CODE)
                grdsales2.TextMatrix(i, 2) = IIf(IsNull(RSTTRXFILE!ITEM_NAME), "", RSTTRXFILE!ITEM_NAME)
                grdsales2.TextMatrix(i, 3) = IIf(IsNull(RSTTRXFILE!QTY), "1", RSTTRXFILE!QTY)
                grdsales2.TextMatrix(i, 4) = IIf(IsNull(RSTTRXFILE!LOOSE_PACK), "", RSTTRXFILE!LOOSE_PACK)
                grdsales2.TextMatrix(i, 5) = IIf(IsNull(RSTTRXFILE!PACK_TYPE), "", RSTTRXFILE!PACK_TYPE)
                grdsales2.TextMatrix(i, 9) = IIf(IsNull(RSTTRXFILE!SALES_TAX), "", RSTTRXFILE!SALES_TAX)
                
                Set rstformula = New ADODB.Recordset
                rstformula.Open "select * from ITEMMAST where ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (rstformula.EOF Or rstformula.BOF) Then
                    grdsales2.TextMatrix(i, 6) = IIf(IsNull(rstformula!P_RETAIL), "", rstformula!P_RETAIL)
                    grdsales2.TextMatrix(i, 7) = IIf(IsNull(rstformula!P_WS), "", rstformula!P_WS)
                    grdsales2.TextMatrix(i, 8) = IIf(IsNull(rstformula!P_VAN), "", rstformula!P_VAN)
                    grdsales2.TextMatrix(i, 10) = IIf(IsNull(rstformula!LOOSE_PACK), "1", rstformula!LOOSE_PACK)
                End If
                rstformula.Close
                Set rstformula = Nothing
                
                TXTINVDATE.text = Format(RSTTRXFILE!VCH_DATE, "DD/MM/YYYY")
                LBLDATE.Caption = Format(Date, "DD/MM/YYYY")
                lbltime.Caption = Time
                lBLpRODUCT.Caption = IIf(IsNull(RSTTRXFILE!ITEM_NAME), "", RSTTRXFILE!ITEM_NAME)
                TXTITEMCODE.text = IIf(IsNull(RSTTRXFILE!ITEM_CODE), "", RSTTRXFILE!ITEM_CODE)
                TXTPRODUCT2.text = IIf(IsNull(RSTTRXFILE!FORM_NAME), "", RSTTRXFILE!FORM_NAME)
                TXTQTY.text = IIf(IsNull(RSTTRXFILE!FORM_QTY), "", RSTTRXFILE!FORM_QTY)
                txtretail.text = IIf(IsNull(RSTTRXFILE!P_RETAIL), "", RSTTRXFILE!P_RETAIL)
                Txtpack.text = IIf(IsNull(RSTTRXFILE!LOOSE_PACK), "1", RSTTRXFILE!LOOSE_PACK)
                If Val(Txtpack.text) = 0 Then Txtpack.text = 1
                txtWS.text = IIf(IsNull(RSTTRXFILE!P_WS), "", RSTTRXFILE!P_WS)
                txtvanrate.text = IIf(IsNull(RSTTRXFILE!P_VAN), "", RSTTRXFILE!P_VAN)
                TxttaxMRP.text = IIf(IsNull(RSTTRXFILE!SALES_TAX), "", RSTTRXFILE!SALES_TAX)
                TxtResult.text = IIf(IsNull(RSTTRXFILE!QTY), "", Format(RSTTRXFILE!QTY / Val(Txtpack.text), "0.00"))
                LblPack.Caption = IIf(IsNull(RSTTRXFILE!PACK_TYPE), "", RSTTRXFILE!PACK_TYPE)
                TxtBarcode.text = IIf(IsNull(RSTTRXFILE!BARCODE), "", RSTTRXFILE!BARCODE)
                
                cmdexit.Enabled = False
                OLD_INV = True
                cmdadd.Enabled = False
                cmdRefresh.Enabled = False
                
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
                
        
            'LBLTIME.Caption = IIf(IsNull(TRXMAST!CFORM_NO), Time, TRXMAST!CFORM_NO)
            
            LBLBILLNO.Caption = Val(txtBillNo.text)
            
            'Call COSTCALCULATION
            
            txtBillNo.Visible = False
            
            If OLD_INV = True Then
                If grdsales.rows > 1 Then
                    cmdRefresh.Enabled = True
                    cmdRefresh.SetFocus
                Else
                    txtremarks.SetFocus
                End If
            End If
    End Select
    DataList1.text = TXTPRODUCT2.text
    Call DataList1_Click
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTBILLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBillNo_LostFocus()
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long

    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PC'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        If Val(txtBillNo.text) > i Then
            MsgBox "The last bill No. is " & i, vbCritical, "BILL..."
            txtBillNo.Visible = True
            txtBillNo.SetFocus
            Exit Sub
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
      
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MIN(VCH_NO) From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PC'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
        If Val(txtBillNo.text) < i Then
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

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.text = "  /  /    " Then
                TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
                txtremarks.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
                txtremarks.SetFocus
            End If
        Case vbKeyEscape
            'Exit Sub
            'If EDIT_INV = True Then Exit Sub
            txtBillNo.Visible = True
            txtBillNo.SetFocus
    End Select
End Sub

Private Sub TXTINVDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.text)
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTQTY.text) = 0 Then Exit Sub
            i = 0
'            Set RSTTRXFILE = New ADODB.Recordset
'            RSTTRXFILE.Open "SELECT BAL_QTY  FROM RTRXFILE WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND VCH_NO = " & Val(TXTVCHNO.Text) & " AND LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockReadOnly
'            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
'                If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
'                i = RSTTRXFILE!BAL_QTY
'            End If
'            RSTTRXFILE.Close
'            Set RSTTRXFILE = Nothing
'
'            If Val(TXTQTY.Text) = 0 Then Exit Sub
'            If i > 0 Then
'                If Val(TXTQTY.Text) > i Then
'                    MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
'                    TXTQTY.SelStart = 0
'                    TXTQTY.SelLength = Len(TXTQTY.Text)
'                    Exit Sub
'                End If
'            End If
'SKIP:
            'TxtResult.SetFocus
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            TXTPRODUCT2.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    TXTQTY.text = Format(TXTQTY.text, ".00")
    If Val(Txtpack.text) = 0 Then Txtpack.text = 1
    TxtResult.text = Format(Val(TXTQTY) * Val(TxtActqty.text), "0.00")
End Sub

'Private Function COSTCALCULATION()
'    Dim RSTCOST As ADODB.Recordset
'    Dim COST As Double
'    Dim N As Integer
'    'Dim RSTITEMMAST As ADODB.Recordset
'
'     'LBLTOTALCOST.Caption = ""
'     'LBLPROFIT.Caption = ""
'        COST = 0
'    On Error GoTo eRRHAND
'    For N = 1 To grdsales.Rows - 1
'        Set RSTCOST = New ADODB.Recordset
'        RSTCOST.Open "SELECT ITEM_COST FROM RTRXFILE WHERE TRX_TYPE = '" & Trim(grdsales.TextMatrix(N, 7)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(N, 5)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(N, 6)) & "", db, adOpenStatic, adLockReadOnly, adCmdText
'        Do Until RSTCOST.EOF
'            'COST = COST + (RSTCOST!ITEM_COST) * Val(grdsales.TextMatrix(N, 3))
'            RSTCOST.MoveNext
'        Loop
'        RSTCOST.Close
'        Set RSTCOST = Nothing
'    Next N
'
'    'LBLTOTALCOST.Caption = Round(COST, 2)
'    'LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - COST, 2)
'
'    Exit Function
'
'eRRHAND:
'    MsgBox Err.Description
'End Function

Private Function AppendSale()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTP_RATE As ADODB.Recordset
    Dim RSTITEMMAST, RSTRTRXFILE, rststock As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim i, M_DATA As Double
    Dim TRXVALUE As Double
    
    Dim DAY_DATE As String
    Dim MONTH_DATE As String
    Dim YEAR_DATE As String
    Dim E_DATE As Date
    i = 0
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    
    ''db.RollbackTrans
    db.BeginTrans
    db.Execute "delete From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete From TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & ""

    
    E_DATE = Format(TXTINVDATE.text, "MM/DD/YYYY")
    If Day(E_DATE) <= 12 Then
        DAY_DATE = Format(Month(E_DATE), "00")
        MONTH_DATE = Format(Day(E_DATE), "00")
        YEAR_DATE = Format(Year(E_DATE), "0000")
        E_DATE = DAY_DATE & "/" & MONTH_DATE & "/" & YEAR_DATE
    End If
    E_DATE = Format(E_DATE, "MM/DD/YYYY")
    
    
    Dim rstTRXMAST As ADODB.Recordset
    Dim ITEMCOST As Double
    ITEMCOST = 0
    For i = 1 To grdsales.rows - 1
        If UCase(grdsales.TextMatrix(i, 8)) <> "SERVICE CHARGE" Then
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
            With rstTRXMAST
                If Not (.EOF And .BOF) Then
                    !ISSUE_QTY = !ISSUE_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                    !FREE_QTY = 0
                    !ISSUE_VAL = 0
                    !CLOSE_QTY = !CLOSE_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                    !CLOSE_VAL = 0
                    '!LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
                    !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
                    ITEMCOST = ITEMCOST + IIf(IsNull(!ITEM_COST), 0, !ITEM_COST * Val(grdsales.TextMatrix(i, 2)) + Val(grdsales.TextMatrix(i, 9)))
                    rstTRXMAST.Update
                End If
            End With
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
                    
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "SELECT *  FROM RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' and ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
            With rstTRXMAST
                If Not (.EOF And .BOF) Then
                    If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                    !ISSUE_QTY = !ISSUE_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                    If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                    !BAL_QTY = !BAL_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                    '!LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
                    !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
                    !PD_NO = Val(txtBillNo.text)
                    rstTRXMAST.Update
                End If
            End With
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
        Else
            ITEMCOST = ITEMCOST + Val(grdsales.TextMatrix(i, 7))
        End If
    Next i
    'ITEMCOST = Round(ITEMCOST / (Val(TxtResult.Text) * Val(TxtPack.Text)), 3)
    '''ITEMCOST = Round(ITEMCOST * 100 / (Val(TxttaxMRP.Text) + 100), 3)
    
    Dim itemnetcost As Double
    Dim totqty As Double
    itemnetcost = 0
    totqty = 0
    For i = 1 To grdsales2.rows - 1
        totqty = totqty + Val(grdsales2.TextMatrix(i, 3))
    Next i
    
    For i = 1 To grdsales2.rows - 1
        If totqty <> 0 Then itemnetcost = Round(ITEMCOST / totqty, 2)
        M_DATA = 0
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE= '" & grdsales2.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        'RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
            RSTRTRXFILE.AddNew
            RSTRTRXFILE!TRX_TYPE = "PC"
            RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTRTRXFILE!VCH_NO = Val(txtBillNo.text)
            RSTRTRXFILE!LINE_NO = i
            RSTRTRXFILE!ITEM_CODE = grdsales2.TextMatrix(i, 1)
            RSTRTRXFILE!QTY = Round(Val(grdsales2.TextMatrix(i, 3)) * Val(grdsales2.TextMatrix(i, 4)), 3)
            RSTRTRXFILE!BAL_QTY = Round(Val(grdsales2.TextMatrix(i, 3)) * Val(grdsales2.TextMatrix(i, 4)), 3)
            
            RSTRTRXFILE!ITEM_COST = itemnetcost
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT *  From ITEMMAST WHERE ITEM_CODE = '" & grdsales2.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With rststock
                If Not (.EOF And .BOF) Then
                    rststock!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
                    !CLOSE_QTY = !CLOSE_QTY + Round(Val(grdsales2.TextMatrix(i, 3)) * Val(grdsales2.TextMatrix(i, 4)), 3)
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    !CLOSE_VAL = 0
                    
                    !RCPT_QTY = !RCPT_QTY + Round(Val(grdsales2.TextMatrix(i, 3)) * Val(grdsales2.TextMatrix(i, 4)), 3)
                    If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                    !RCPT_VAL = 0 ' !RCPT_VAL + Val(grdsales2.TextMatrix(Val(TXTSLNO.Text), 13))
                    
                    If Val(grdsales2.TextMatrix(i, 6)) <> 0 Then
                        !P_RETAIL = Val(grdsales2.TextMatrix(i, 6))
                        !P_CRTN = Val(grdsales2.TextMatrix(i, 6)) / Val(grdsales2.TextMatrix(i, 4))
                    End If
                    If Val(grdsales2.TextMatrix(i, 7)) <> 0 Then
                        !P_WS = Val(grdsales2.TextMatrix(i, 7))
                        !P_LWS = Val(grdsales2.TextMatrix(i, 7)) / Val(grdsales2.TextMatrix(i, 4))
                    End If
                    If Val(grdsales2.TextMatrix(i, 8)) <> 0 Then !P_VAN = Val(grdsales2.TextMatrix(i, 8))
                    If Val(grdsales2.TextMatrix(i, 10)) > 1 Then !LOOSE_PACK = Val(grdsales2.TextMatrix(i, 10))
                    If Val(grdsales2.TextMatrix(i, 9)) <> 0 Then !SALES_TAX = Val(grdsales2.TextMatrix(i, 9))
                    !CRTN_PACK = 1
                    !ITEM_COST = itemnetcost
                    RSTRTRXFILE!MFGR = !MANUFACTURER
                    rststock.Update
                End If
            End With
            rststock.Close
            Set rststock = Nothing
            
        Else
            M_DATA = Round(Val(grdsales2.TextMatrix(i, 3)) * Val(grdsales2.TextMatrix(i, 4)), 3)
            M_DATA = M_DATA - (RSTRTRXFILE!QTY - RSTRTRXFILE!BAL_QTY)
            RSTRTRXFILE!BAL_QTY = M_DATA
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT *  From ITEMMAST WHERE ITEM_CODE = '" & grdsales2.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With rststock
                If Not (.EOF And .BOF) Then
                    rststock!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
                    !CLOSE_QTY = !CLOSE_QTY - RSTRTRXFILE!QTY
                    !CLOSE_QTY = !CLOSE_QTY + (Round(Val(grdsales2.TextMatrix(i, 3)) * Val(grdsales2.TextMatrix(i, 4)), 3))
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    !CLOSE_VAL = 0 '!CLOSE_VAL + Val(grdsales2.TextMatrix(Val(TXTSLNO.Text), 13))
                    
                    !RCPT_QTY = !RCPT_QTY - RSTRTRXFILE!QTY
                    !RCPT_QTY = !RCPT_QTY + (Round(Val(grdsales2.TextMatrix(i, 3)) * Val(grdsales2.TextMatrix(i, 4)), 3))
                    If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                    !RCPT_VAL = 0 '!RCPT_VAL + Val(grdsales2.TextMatrix(Val(TXTSLNO.Text), 13))
                    !ITEM_COST = itemnetcost
                    RSTRTRXFILE!MFGR = !MANUFACTURER
                    If Val(grdsales2.TextMatrix(i, 6)) <> 0 Then
                        !P_RETAIL = Val(grdsales2.TextMatrix(i, 6))
                        !P_CRTN = Val(grdsales2.TextMatrix(i, 6)) / Val(grdsales2.TextMatrix(i, 4))
                    End If
                    If Val(grdsales2.TextMatrix(i, 7)) <> 0 Then
                        !P_WS = Val(grdsales2.TextMatrix(i, 7))
                        !P_LWS = Val(grdsales2.TextMatrix(i, 7)) / Val(grdsales2.TextMatrix(i, 4))
                    End If
                    If Val(grdsales2.TextMatrix(i, 8)) <> 0 Then !P_VAN = Val(grdsales2.TextMatrix(i, 8))
                    If Val(grdsales2.TextMatrix(i, 10)) > 1 Then !LOOSE_PACK = Val(grdsales2.TextMatrix(i, 10))
                    If Val(grdsales2.TextMatrix(i, 9)) <> 0 Then !SALES_TAX = Val(grdsales2.TextMatrix(i, 9))
                    !CRTN_PACK = 1
                    rststock.Update
                End If
            End With
            rststock.Close
            Set rststock = Nothing
            RSTRTRXFILE!QTY = (Round(Val(grdsales2.TextMatrix(i, 3)) * Val(grdsales2.TextMatrix(i, 4)), 3))
        End If
        RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTRTRXFILE!ITEM_NAME = grdsales2.TextMatrix(i, 2)
        RSTRTRXFILE!FORM_CODE = DataList1.BoundText
        RSTRTRXFILE!FORM_QTY = Val(TXTQTY.text)
        RSTRTRXFILE!FORM_NAME = DataList1.text
        RSTRTRXFILE!TRX_TOTAL = 0 'Val(grdsales2.TextMatrix(Val(TXTSLNO.Text), 13))
        RSTRTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "dd/mm/yyyy")
        RSTRTRXFILE!BARCODE = Trim(TxtBarcode.text)
        RSTRTRXFILE!P_RETAIL = Val(grdsales2.TextMatrix(i, 6))
        RSTRTRXFILE!LOOSE_PACK = Val(grdsales2.TextMatrix(i, 10))
        RSTRTRXFILE!P_WS = Val(grdsales2.TextMatrix(i, 7))
        RSTRTRXFILE!P_VAN = Val(grdsales2.TextMatrix(i, 8))
        RSTRTRXFILE!SALES_TAX = Val(grdsales2.TextMatrix(i, 9))
        RSTRTRXFILE!CRTN_PACK = 1
        RSTRTRXFILE!P_CRTN = Val(grdsales2.TextMatrix(i, 6)) / Val(grdsales2.TextMatrix(i, 4))
        RSTRTRXFILE!P_LWS = Val(grdsales2.TextMatrix(i, 7)) / Val(grdsales2.TextMatrix(i, 4))
        'RSTRTRXFILE!LOOSE_PACK = 1
        RSTRTRXFILE!LINE_DISC = 1 ' Val(grdsales2.TextMatrix(Val(TXTSLNO.Text), 5))
        RSTRTRXFILE!P_DISC = 0 'Val(grdsales2.TextMatrix(Val(TXTSLNO.Text), 17))
        RSTRTRXFILE!MRP = 0 'Val(grdsales2.TextMatrix(Val(TXTSLNO.Text), 6))
        RSTRTRXFILE!PTR = 0 ' Val(grdsales2.TextMatrix(Val(TXTSLNO.Text), 9))
        RSTRTRXFILE!SALES_PRICE = 0
        RSTRTRXFILE!Category = "RAW"
        RSTRTRXFILE!UNIT = 1 'Val(grdsales2.TextMatrix(Val(TXTSLNO.Text), 4))
        RSTRTRXFILE!REF_NO = "" 'Trim(grdsales2.TextMatrix(Val(TXTSLNO.Text), 11))
        RSTRTRXFILE!CST = 0
            
        RSTRTRXFILE!SCHEME = 0
        'RSTRTRXFILE!EXP_DATE = Null
        RSTRTRXFILE!FREE_QTY = 0
        RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTRTRXFILE!C_USER_ID = "SM"
        RSTRTRXFILE!check_flag = "V"
        RSTRTRXFILE.Update
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing
   Next i
    TRXVALUE = 0
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "Select * From TRXSUB ", db, adOpenStatic, adLockOptimistic, adCmdText
'    For i = 1 To grdsales.Rows - 1
'        RSTTRXFILE.AddNew
'        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
'        RSTTRXFILE!TRX_TYPE = "PC"
'        RSTTRXFILE!LINE_NO = i
'        RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(i, 5) = "", 0, grdsales.TextMatrix(i, 5))
'        RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(i, 6) = "", 0, grdsales.TextMatrix(i, 6))
'        RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(i, 7) = "", "PC", grdsales.TextMatrix(i, 7))
'        RSTTRXFILE!QTY = grdsales.TextMatrix(i, 3)
'        RSTTRXFILE.Update
'    Next i
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    
    Dim RSTITEMCOST As ADODB.Recordset
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "PC"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = Trim(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!WASTE_QTY = Val(grdsales.TextMatrix(i, 9))
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 1)
        If UCase(grdsales.TextMatrix(i, 8)) <> "SERVICE CHARGE" Then
            RSTTRXFILE!QTY = (Val(grdsales.TextMatrix(i, 2))) '* Val(grdsales.TextMatrix(i, 5)))
            RSTTRXFILE!LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
            RSTTRXFILE!PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
            RSTTRXFILE!LOOSE_FLAG = "L"
            RSTTRXFILE!ITEM_COST = 0
            Set RSTITEMCOST = New ADODB.Recordset
            RSTITEMCOST.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            With RSTITEMCOST
                If Not (.EOF And .BOF) Then
                    RSTTRXFILE!ITEM_COST = RSTITEMCOST!ITEM_COST
                End If
            End With
            RSTITEMCOST.Close
            Set RSTITEMCOST = Nothing
        Else
            RSTTRXFILE!QTY = 0
            RSTTRXFILE!LOOSE_PACK = 0
            RSTTRXFILE!PACK_TYPE = ""
            RSTTRXFILE!LOOSE_FLAG = ""
            RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(i, 7))
        End If
        RSTTRXFILE!MRP = 0
        RSTTRXFILE!PTR = 0
        RSTTRXFILE!SALES_PRICE = 0
        RSTTRXFILE!P_RETAIL = 0
        RSTTRXFILE!P_RETAILWOTAX = 0
        RSTTRXFILE!SALES_TAX = 0
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 3)
        RSTTRXFILE!VCH_DESC = "Issued to      Factory" '& Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = ""
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = ""
        RSTTRXFILE!MFGR = ""
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = 0
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "" 'DataList2.BoundText
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PC' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!TRX_TYPE = "PC"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    End If
    RSTTRXFILE!VCH_AMOUNT = 0
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = "10111"
    RSTTRXFILE!ACT_NAME = "PRESS"
    RSTTRXFILE!DISCOUNT = 0
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = 0
    RSTTRXFILE!REF_NO = ""
    RSTTRXFILE!SLSM_CODE = ""
    RSTTRXFILE!PAY_AMOUNT = 0
    RSTTRXFILE!check_flag = ""
    RSTTRXFILE!POST_FLAG = ""
    RSTTRXFILE!CFORM_NO = lbltime.Caption
    RSTTRXFILE!REMARKS = ""
    RSTTRXFILE!DISC_PERS = 0
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!BILL_NAME = DataList1.text
    RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!C_USER_ID = "SM"
    
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    i = 0
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PC'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    Dim INWARD, OUTWARD, BALQTY, DIFFQTY As Double
    
    For i = 1 To grdsales.rows - 1
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTITEMMAST.EOF
            INWARD = 0
            OUTWARD = 0
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            Do Until rststock.EOF
                INWARD = INWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                INWARD = INWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) ' * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='PC' OR TRX_TYPE='RM' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockReadOnly
            Do Until rststock.EOF
                OUTWARD = OUTWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)) + IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
                'OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            '=============
            
            BALQTY = 0
            db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY <0"
            If Round(INWARD - OUTWARD, 2) = 0 Then
                db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY >0"
            Else
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT SUM(BAL_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY > 0", db, adOpenForwardOnly
                If Not (rststock.EOF And rststock.BOF) Then
                    BALQTY = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
                End If
                rststock.Close
                Set rststock = Nothing
            End If
        
            If Round(INWARD - OUTWARD, 2) < BALQTY Then
                DIFFQTY = BALQTY - (Round(INWARD - OUTWARD, 2))
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ", db, adOpenStatic, adLockOptimistic, adCmdText
                Do Until rststock.EOF
                    If DIFFQTY - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) >= 0 Then
                        DIFFQTY = DIFFQTY - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY)
                        rststock!BAL_QTY = 0
                        rststock.Update
                    Else
                        rststock!BAL_QTY = IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) - DIFFQTY
                        DIFFQTY = 0
                        rststock.Update
                    End If
                    If DIFFQTY <= 0 Then Exit Do
                    rststock.MoveNext
                Loop
                rststock.Close
                Set rststock = Nothing
            ElseIf Round(INWARD - OUTWARD, 2) > BALQTY Then
                DIFFQTY = Round((INWARD - OUTWARD), 2) - BALQTY
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockOptimistic, adCmdText
                Do Until rststock.EOF
                    If DIFFQTY <= IIf(IsNull(rststock!QTY), 0, rststock!QTY) - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) Then
                        rststock!BAL_QTY = IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) + DIFFQTY
                        DIFFQTY = 0
                    Else
                        If Not rststock!BAL_QTY = IIf(IsNull(rststock!QTY), 0, rststock!QTY) Then
                            rststock!BAL_QTY = IIf(IsNull(rststock!QTY), 0, rststock!QTY)
                            DIFFQTY = DIFFQTY - IIf(IsNull(rststock!QTY), 0, rststock!QTY)
                        End If
                    End If
                    rststock.Update
                    If DIFFQTY <= 0 Then Exit Do
                    rststock.MoveNext
                Loop
                rststock.Close
                Set rststock = Nothing
                'MsgBox ""
            End If
            
            '============
            
            RSTITEMMAST!CLOSE_QTY = INWARD - OUTWARD
            RSTITEMMAST!RCPT_QTY = INWARD
            RSTITEMMAST!ISSUE_QTY = OUTWARD
            RSTITEMMAST.Update
            RSTITEMMAST.MoveNext
        Loop
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    Next i
    
    For i = 1 To grdsales2.rows - 1
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & grdsales2.TextMatrix(i, 1) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTITEMMAST.EOF
            INWARD = 0
            OUTWARD = 0
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            Do Until rststock.EOF
                INWARD = INWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) ' * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                INWARD = INWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='PC' OR TRX_TYPE='RM' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until rststock.EOF
                'OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                OUTWARD = OUTWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)) + IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
                OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            RSTITEMMAST!CLOSE_QTY = INWARD - OUTWARD
            RSTITEMMAST!RCPT_QTY = INWARD
            RSTITEMMAST!ISSUE_QTY = OUTWARD
            RSTITEMMAST.Update
            RSTITEMMAST.MoveNext
        Loop
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    Next i
    db.CommitTrans
    
    TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
    LBLDATE.Caption = Date
    lbltime.Caption = Time
    'LBLTOTALCOST.Caption = ""
    grdsales.rows = 1
    grdsales2.rows = 1
    M_EDIT = False
    EDIT_INV = False
    TXTPRODUCT2.text = ""
    TXTQTY.text = ""
    lBLpRODUCT.Caption = ""
    TXTITEMCODE.text = ""
    TxtResult.text = ""
    TxtBarcode.text = ""
    LblPack.Caption = ""
    txtretail.text = ""
    Txtpack.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    TxttaxMRP.text = ""
    
    cmdRefresh.Enabled = False
    cmdexit.Enabled = True
    CMDPRINT.Enabled = False
    cmdexit.Enabled = True
    FRMEHEAD.Enabled = True
    TXTPRODUCT2.SetFocus
    'LBLITEMCOST.Caption = ""
    TXTQTY.Tag = ""
    OLD_INV = False
    Screen.MousePointer = vbNormal
    Exit Function
ErrHand:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Function

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
    Open App.Path & "\Report.PRN" For Output As #1 '//Report file Creation
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open App.Path & "\Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ErrHand
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
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenForwardOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!Address & ", " & RSTCOMPANY!HO_NAME, 140)
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
    'MsgBox "Report file generated at " & App.Path & "\Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub txtremarks_GotFocus()
    txtremarks.SelStart = 0
    txtremarks.SelLength = Len(txtremarks.text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            FRMEHEAD.Enabled = False
            TXTPRODUCT2.SetFocus
        Case vbKeyEscape
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
    End Select
End Sub

Private Sub TXTREMARKS_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtProduct2_Change()
    Dim rstCharge As ADODB.Recordset
    On Error GoTo ErrHand
    If flagchange2.Caption <> "1" Then
        If MIX_FLAG = True Then
            MIX_ITEM.Open "select DISTINCT FOR_NO, FOR_NAME from TRXFORMULAMAST where FOR_NAME Like '" & TXTPRODUCT2.text & "%' AND TRX_TYPE='PR' ORDER BY FOR_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            MIX_FLAG = False
        Else
            MIX_ITEM.Close
            MIX_ITEM.Open "select DISTINCT FOR_NO, FOR_NAME from TRXFORMULAMAST  where FOR_NAME Like '" & TXTPRODUCT2.text & "%' AND TRX_TYPE='PR' ORDER BY FOR_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            MIX_FLAG = False
        End If
        If (MIX_ITEM.EOF And MIX_ITEM.BOF) Then
            LBLDEALER2.Caption = ""
        Else
            LBLDEALER2.Caption = MIX_ITEM!FOR_NAME
            'TxtActqty.Text = IIf(IsNull(MIX_ITEM!QTY), "", MIX_ITEM!QTY)
            'TXTITEMCODE.Text = IIf(IsNull(MIX_ITEM!ITEM_CODE), "", MIX_ITEM!ITEM_CODE)
        End If
        Set DataList1.RowSource = MIX_ITEM
        DataList1.ListField = "FOR_NAME"
        DataList1.BoundColumn = "FOR_NO"
       
    End If
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TxtProduct2_GotFocus()
    TXTPRODUCT2.SelStart = 0
    TXTPRODUCT2.SelLength = Len(TXTPRODUCT2.text)
End Sub

Private Sub TxtProduct2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList1.VisibleCount = 0 Then Exit Sub
            DataList1.Enabled = True
            DataList1.SetFocus
    End Select
End Sub

Private Sub TxtProduct2_KeyPress(KeyAscii As Integer)
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
    TXTPRODUCT2.text = DataList1.text
    LBLDEALER2.Caption = TXTPRODUCT2.text
    
    Dim rstformula As ADODB.Recordset
    On Error GoTo ErrHand
    
    If DataList1.BoundText = "" Then Exit Sub
    Set rstformula = New ADODB.Recordset
    rstformula.Open "select * from TRXFORMULAMAST where FOR_NO = " & DataList1.BoundText & " and TRX_TYPE='PR'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstformula.EOF Or rstformula.BOF) Then
        TxtActqty.text = IIf(IsNull(rstformula!QTY), "", rstformula!QTY)
        Txtpack.text = IIf(IsNull(rstformula!LOOSE_PACK), "", rstformula!LOOSE_PACK)
        LblPack.Caption = IIf(IsNull(rstformula!PACK_TYPE), "", rstformula!PACK_TYPE)
        TXTITEMCODE.text = IIf(IsNull(rstformula!ITEM_CODE), "", rstformula!ITEM_CODE)
        lBLpRODUCT.Caption = IIf(IsNull(rstformula!ITEM_NAME), "", rstformula!ITEM_NAME)
    Else
        TxtActqty.text = ""
        Txtpack.text = ""
        TXTITEMCODE.text = ""
        lBLpRODUCT.Caption = ""
        LblPack.Caption = ""
    End If
    rstformula.Close
    Set rstformula = Nothing
    
    Set rstformula = New ADODB.Recordset
    rstformula.Open "select * from ITEMMAST where ITEM_CODE = '" & TXTITEMCODE.text & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstformula.EOF Or rstformula.BOF) Then
        txtretail.text = IIf(IsNull(rstformula!P_RETAIL), "", rstformula!P_RETAIL)
        Txtpack.text = IIf(IsNull(rstformula!LOOSE_PACK), "1", rstformula!LOOSE_PACK)
        txtWS.text = IIf(IsNull(rstformula!P_WS), "", rstformula!P_WS)
        txtvanrate.text = IIf(IsNull(rstformula!P_VAN), "", rstformula!P_VAN)
        TxttaxMRP.text = IIf(IsNull(rstformula!SALES_TAX), "", rstformula!SALES_TAX)
    End If
    rstformula.Close
    Set rstformula = Nothing
    
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            If DataList1.text = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Mixture from the List", vbOKOnly, "PROCESS"
                DataList1.SetFocus
                Exit Sub
            End If
            
            'TXTPRODUCT2.Enabled = False
            'DataList1.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
            
        Case vbKeyEscape
            TXTPRODUCT2.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_GotFocus()
    flagchange2.Caption = 1
    TXTPRODUCT2.text = LBLDEALER2.Caption
    DataList1.text = TXTPRODUCT2.text
    Call DataList1_Click
End Sub

Private Sub DataList1_LostFocus()
    flagchange2.Caption = ""
    
    Dim rstformula As ADODB.Recordset
    On Error GoTo ErrHand
    If DataList1.BoundText = "" Then Exit Sub
    
    Set rstformula = New ADODB.Recordset
    rstformula.Open "select * from TRXFORMULAMAST where FOR_NO = " & DataList1.BoundText & " and TRX_TYPE='PR'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstformula.EOF Or rstformula.BOF) Then
        TxtActqty.text = IIf(IsNull(rstformula!QTY), "", rstformula!QTY)
        Txtpack.text = IIf(IsNull(rstformula!LOOSE_PACK), "", rstformula!LOOSE_PACK)
        LblPack.Caption = IIf(IsNull(rstformula!PACK_TYPE), "", rstformula!PACK_TYPE)
        TXTITEMCODE.text = IIf(IsNull(rstformula!ITEM_CODE), "", rstformula!ITEM_CODE)
        lBLpRODUCT.Caption = IIf(IsNull(rstformula!ITEM_NAME), "", rstformula!ITEM_NAME)
    Else
        TxtActqty.text = ""
        Txtpack.text = ""
        TXTITEMCODE.text = ""
        lBLpRODUCT.Caption = ""
        LblPack.Caption = ""
    End If
    rstformula.Close
    Set rstformula = Nothing
    
    
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub


Private Sub TxtResult_GotFocus()
    TxtResult.SelStart = 0
    TxtResult.SelLength = Len(TxtResult.text)

End Sub

Private Sub TxtResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtResult.text) = 0 Then Exit Sub
            txtretail.SetFocus
         Case vbKeyEscape
            TXTQTY.SetFocus
    End Select

End Sub

Private Sub TxtResult_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub TXTRETAIL_GotFocus()
    txtretail.SelStart = 0
    txtretail.SelLength = Len(txtretail.text)
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtWS.SetFocus
         Case vbKeyEscape
            TxtResult.SetFocus
    End Select
End Sub

Private Sub TXTRETAIL_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRETAIL_LostFocus()
    txtretail.text = Format(txtretail.text, "0.00")
End Sub

Private Sub TXTsample_LostFocus()
    TXTsample.Visible = False
End Sub

Private Sub txtws_GotFocus()
    txtWS.SelStart = 0
    txtWS.SelLength = Len(txtWS.text)
End Sub

Private Sub txtws_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtvanrate.SetFocus
         Case vbKeyEscape
            txtretail.SetFocus
    End Select
End Sub

Private Sub txtws_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtws_LostFocus()
    txtWS.text = Format(txtWS.text, "0.00")
End Sub

Private Sub txtvanrate_GotFocus()
    txtvanrate.SelStart = 0
    txtvanrate.SelLength = Len(txtvanrate.text)
End Sub

Private Sub txtvanrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxttaxMRP.SetFocus
         Case vbKeyEscape
            txtWS.SetFocus
    End Select
End Sub

Private Sub txtvanrate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtvanrate_LostFocus()
    txtvanrate.text = Format(txtvanrate.text, "0.00")
End Sub

Private Sub TxttaxMRP_GotFocus()
    TxttaxMRP.SelStart = 0
    TxttaxMRP.SelLength = Len(TxttaxMRP.text)
End Sub

Private Sub TxttaxMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If OLD_INV = False Then
                cmdadd.Enabled = True
                cmdadd.SetFocus
            Else
                cmdadd.Enabled = False
                cmdRefresh.Enabled = False
            End If
         Case vbKeyEscape
            txtvanrate.SetFocus
    End Select
End Sub

Private Sub TxttaxMRP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxttaxMRP_LostFocus()
    TxttaxMRP.text = Format(TxttaxMRP.text, "0.00")
End Sub

Private Function ResetStock()
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    For i = 1 To grdsales.rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                !ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                !FREE_QTY = 0
                !ISSUE_VAL = 0
                !CLOSE_QTY = !CLOSE_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                !CLOSE_VAL = 0
                '!LOOSE_PACK = Val(grdsales.TextMatrix(i, 5))
                !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND PD_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                !LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
                !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
    
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    Next i
    
End Function

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 2  ' QTY
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = TXTsample.text
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                Case 9  ' WASTAGE
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = TXTsample.text
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            grdsales.SetFocus
    End Select
        Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case grdsales.Col
        Case 2
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

Private Function print_labels(i As Long, BAR_LABEL As String, P_RATE As String)
    Dim wid As Single
    Dim hgt As Single
    
    On Error GoTo ErrHand
    
'    Dim P, PNAME
'    Dim printerfound As Boolean
'    printerfound = False
'    For Each P In Printers
'        PNAME = P.DeviceName
'        If UCase(Right(PNAME, 16)) Like "BAR CODE PRINTER" Then
'            Set Printer = P
'            printerfound = True
'            Exit For
'        End If
'    Next P
'    If printerfound = False Then
'        MsgBox ("Printer not found. Please correct the printer name")
'        Exit Function
'    End If
    
    'i = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3))
    
    Picture1.Cls
    Picture1.Picture = Nothing
    Picture1.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
    Picture1.CurrentY = 0 'Y2 + 0.25 * Th
    Picture1.FontName = "MS Sans Serif"
    Picture1.FontSize = 7
    Picture1.FontBold = True
    Picture1.Print Trim(MDIMAIN.StatusBar.Panels(5).text) 'COMP NAME
    
    Picture2.Cls
    Picture2.Picture = Nothing
    Picture2.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
    Picture2.CurrentY = 0 'Y2 + 0.25 * Th
    Picture2.FontName = "MS Sans Serif"
    Picture2.FontSize = 6
    Picture2.FontBold = False
    Picture2.Print Trim(TXTPRODUCT2.text) 'ITEM NAME
        
    If Val(txtretail.text) <> 0 Then
        Picture5.Cls
        Picture5.Picture = Nothing
        Picture5.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
        Picture5.CurrentY = 0 'Y2 + 0.25 * Th
        Picture2.FontName = "Arial"
        Picture2.FontSize = 6
        Picture2.FontBold = True
        Picture5.Print "Price: " & Format(Val(txtretail.text), "0.00")
    End If
    
'    If Val(TXTRATE.Text) > 0 And Val(TXTRETAIL.Text) < Val(TXTRATE.Text) Then
'        Picture6.Cls
'        Picture6.Picture = Nothing
'        Picture6.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'        Picture6.CurrentY = 0 'Y2 + 0.25 * Th
'        Picture2.FontName = "Arial"
'        Picture2.FontSize = 6
'        Picture2.FontBold = True
'        Picture6.Print "MRP  : " & Format(Val(TXTRATE.Text), "0.00")
'    End If
    

'    Picture3.Cls
'    Picture3.Picture = Nothing
'    Picture3.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'    Picture3.CurrentY = 0 'Y2 + 0.25 * Th
'    Picture3.FontName = "barcode font"
'    Picture3.FontSize = 14
'    Picture3.FontBold = False
'    Picture3.Print BAR_LABEL
    
    Do Until i <= 0
        Picture1.ScaleMode = vbPixels
        Picture2.ScaleMode = vbPixels
        Picture5.ScaleMode = vbPixels
        Picture6.ScaleMode = vbPixels
        
        '=========
        Select Case MDIMAIN.barcode_profile.Caption
            Case 1
                Printer.PaintPicture Picture2.Image, 900, 800 ', wid, hgt  'Item Name      'SREEDEVI, PARTHAN
                Printer.PaintPicture Picture2.Image, 3150, 800 ', wid, hgt  'Item Name
        
                Printer.PaintPicture Picture5.Image, 900, 960 ', wid, hgt  ' Price
                Printer.PaintPicture Picture5.Image, 3150, 960 ', wid, hgt
        
                Printer.PaintPicture Picture1.Image, 900, 1160 ', wid, hgt  ' Comp Name
                Printer.PaintPicture Picture1.Image, 3150, 1160 ', wid, hgt
            Case 2
                Printer.PaintPicture Picture2.Image, 200, 800 ', wid, hgt  'Item Name       'IHIJABI, NRS
                Printer.PaintPicture Picture2.Image, 2400, 800 ', wid, hgt  'Item Name
        
                Printer.PaintPicture Picture5.Image, 200, 960 ', wid, hgt  ' Price
                Printer.PaintPicture Picture5.Image, 2400, 960 ', wid, hgt
        
                Printer.PaintPicture Picture1.Image, 200, 1160 ', wid, hgt  ' Comp Name
                Printer.PaintPicture Picture1.Image, 2400, 1160 ', wid, hgt
            Case 3
                Printer.PaintPicture Picture2.Image, 200, 800 ', wid, hgt  'Item Name       'NUNU
                Printer.PaintPicture Picture2.Image, 3000, 800 ', wid, hgt  'Item Name
        
                Printer.PaintPicture Picture5.Image, 200, 960 ', wid, hgt  ' Price
                Printer.PaintPicture Picture5.Image, 3000, 960 ', wid, hgt
        
                Printer.PaintPicture Picture1.Image, 200, 1160 ', wid, hgt  ' Comp Name
                Printer.PaintPicture Picture1.Image, 3000, 1160 ', wid, hgt
            Case Else
                Printer.PaintPicture Picture2.Image, 200, 800 ', wid, hgt  'Item Name       'soubhagya
                Printer.PaintPicture Picture2.Image, 3200, 800 ', wid, hgt  'Item Name
        
                Printer.PaintPicture Picture5.Image, 200, 960 ', wid, hgt  ' Price
                Printer.PaintPicture Picture5.Image, 3200, 960 ', wid, hgt
        
                Printer.PaintPicture Picture1.Image, 200, 1160 ', wid, hgt  ' Comp Name
                Printer.PaintPicture Picture1.Image, 3200, 1160 ', wid, hgt
        End Select
            
        Printer.FontName = "Arial"
        'Printer.FontName = "barcode font"
        Printer.FontSize = 5
        Printer.FontBold = False
        Printer.Print ""
        
        Printer.FontName = "IDAutomationHC39M"
        'Printer.FontName = "barcode font"
        Printer.FontSize = 24
        Printer.FontBold = False
        Dim bar_space As Integer
        If Len(BAR_LABEL) > 13 Then
            bar_space = 0
            Printer.FontSize = 6
        ElseIf Len(BAR_LABEL) >= 12 Then
            bar_space = 0
            Printer.FontSize = 7
        Else
            Select Case MDIMAIN.barcode_profile.Caption
                Case 1
                    bar_space = 9 - Len(BAR_LABEL) 'Parthan
                Case 2
                    bar_space = 9 - Len(BAR_LABEL) 'ihijabi, NRS
                Case 3
                    bar_space = 12 - Len(BAR_LABEL) 'NUNU
                Case Else
                    bar_space = 13 - Len(BAR_LABEL) 'soubhagya
            End Select
            Printer.FontSize = 11
        End If
        Select Case MDIMAIN.barcode_profile.Caption
            Case 1
                Printer.Print "    (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")" ' parthan
            Case 2
                Printer.Print " (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")" 'ihijabi
            Case 3
                Printer.Print " (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")" 'NUNU
            Case Else
                Printer.Print " (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")" 'NUNU
        End Select
        
'        'Picture1.ScaleMode = vbPixels
'        Picture5.ScaleMode = vbPixels
'        Picture6.ScaleMode = vbPixels
        ' Finish printing.
        Printer.EndDoc
        i = i - 2
    Loop
    
    Exit Function
ErrHand:
    MsgBox err.Description
End Function

Private Function print_3labels(i As Long, BAR_LABEL As String, P_RATE As String)
    Dim wid As Single
    Dim hgt As Single
    
    On Error GoTo ErrHand
    
'    Dim P, PNAME
'    Dim printerfound As Boolean
'    printerfound = False
'    For Each P In Printers
'        PNAME = P.DeviceName
'        If UCase(Right(PNAME, 16)) Like "BAR CODE PRINTER" Then
'            Set Printer = P
'            printerfound = True
'            Exit For
'        End If
'    Next P
'    If printerfound = False Then
'        MsgBox ("Printer not found. Please correct the printer name")
'        Exit Function
'    End If
    
    'i = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3))
    
    Picture1.Cls
    Picture1.Picture = Nothing
    Picture1.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
    Picture1.CurrentY = 0 'Y2 + 0.25 * Th
    Picture1.FontName = "MS Sans Serif"
    Picture1.FontSize = 7
    Picture1.FontBold = True
    Picture1.Print Trim(MDIMAIN.StatusBar.Panels(5).text) 'COMP NAME
    
    Picture2.Cls
    Picture2.Picture = Nothing
    Picture2.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
    Picture2.CurrentY = 0 'Y2 + 0.25 * Th
    Picture2.FontName = "MS Sans Serif"
    Picture2.FontSize = 6
    Picture2.FontBold = False
    Picture2.Print Trim(TXTPRODUCT2.text) 'ITEM NAME
        
    If Val(txtretail.text) <> 0 Then
        Picture5.Cls
        Picture5.Picture = Nothing
        Picture5.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
        Picture5.CurrentY = 0 'Y2 + 0.25 * Th
        Picture2.FontName = "Arial"
        Picture2.FontSize = 6
        Picture2.FontBold = True
        Picture5.Print "Price: " & Format(Val(txtretail.text), "0.00")
    End If
    
'    If Val(TXTRATE.Text) > 0 And Val(TXTRETAIL.Text) < Val(TXTRATE.Text) Then
'        Picture6.Cls
'        Picture6.Picture = Nothing
'        Picture6.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'        Picture6.CurrentY = 0 'Y2 + 0.25 * Th
'        Picture2.FontName = "Arial"
'        Picture2.FontSize = 6
'        Picture2.FontBold = True
'        Picture6.Print "MRP  : " & Format(Val(TXTRATE.Text), "0.00")
'    End If
    

'    Picture3.Cls
'    Picture3.Picture = Nothing
'    Picture3.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'    Picture3.CurrentY = 0 'Y2 + 0.25 * Th
'    Picture3.FontName = "barcode font"
'    Picture3.FontSize = 14
'    Picture3.FontBold = False
'    Picture3.Print BAR_LABEL
    
    Do Until i <= 0
        Picture1.ScaleMode = vbPixels
        Picture2.ScaleMode = vbPixels
        Picture5.ScaleMode = vbPixels
        Picture6.ScaleMode = vbPixels
        
        Printer.PaintPicture Picture2.Image, 200, 830 ', wid, hgt  'Item Name
        Printer.PaintPicture Picture2.Image, 2000, 830 ', wid, hgt  'Item Name
        Printer.PaintPicture Picture2.Image, 4000, 830 ', wid, hgt  'Item Name
        
        Printer.PaintPicture Picture5.Image, 200, 1040 ', wid, hgt  ' Price
        Printer.PaintPicture Picture5.Image, 2000, 1040 ', wid, hgt
        Printer.PaintPicture Picture5.Image, 4000, 1040 ', wid, hgt
        
        'Printer.PaintPicture Picture6.Image, 2000, 950 ', wid, hgt 'MRP
        'Printer.PaintPicture Picture6.Image, 3500, 600 ', wid, hgt 'MRP
        
        Printer.PaintPicture Picture1.Image, 200, 1240 ', wid, hgt  ' Comp Name
        Printer.PaintPicture Picture1.Image, 2000, 1240 ', wid, hgt
        Printer.PaintPicture Picture1.Image, 4000, 1240 ', wid, hgt
        
        
        
        Printer.FontName = "Arial"
        'Printer.FontName = "barcode font"
        Printer.FontSize = 5
        Printer.FontBold = False
        Printer.Print ""
        
        Printer.FontName = "IDAutomationHC39M"
        'Printer.FontName = "barcode font"
        Printer.FontSize = 24
        Printer.FontBold = False
        Dim bar_space As Integer
        If Len(BAR_LABEL) > 13 Then
            bar_space = 0
            Printer.FontSize = 6
        ElseIf Len(BAR_LABEL) >= 12 Then
            bar_space = 0
            Printer.FontSize = 7
        Else
            bar_space = 7 - Len(BAR_LABEL)
            Printer.FontSize = 11
        End If
        
        'Printer.Print " (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")"
        'Printer.Print " (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")"
        Printer.Print "(" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")"
'        'Picture1.ScaleMode = vbPixels
'        Picture5.ScaleMode = vbPixels
'        Picture6.ScaleMode = vbPixels
        ' Finish printing.
        Printer.EndDoc
        i = i - 3
    Loop
    
    Exit Function
ErrHand:
    MsgBox err.Description
End Function

Private Sub Txtpack_GotFocus()
    If Val(Txtpack.text) = 0 Then Txtpack.text = 1
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.text)
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtWS.SetFocus
         Case vbKeyEscape
            TxtResult.SetFocus
    End Select
End Sub

Private Sub Txtpack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtPack_LostFocus()
    Txtpack.text = Format(Txtpack.text, "0.00")
End Sub

Private Sub grdsales2_Click()
    On Error Resume Next
    TXTsample.Visible = False
    TXTsample2.Visible = False
    grdsales2.SetFocus
End Sub

Private Sub grdsales2_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If grdsales2.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            'If OLD_INV = True Then Exit Sub
            Select Case grdsales2.Col
                Case 3, 4, 6, 7, 8, 9, 10
                    TXTsample2.Visible = True
                    TXTsample2.Top = grdsales2.CellTop + 130
                    TXTsample2.Left = grdsales2.CellLeft + 50
                    TXTsample2.Width = grdsales2.CellWidth
                    TXTsample2.Height = grdsales2.CellHeight
                    TXTsample2.text = grdsales2.TextMatrix(grdsales2.Row, grdsales2.Col)
                    TXTsample2.SetFocus
            End Select
    End Select
End Sub

Private Sub grdsales2_Scroll()
    TXTsample.Visible = False
    TXTsample2.Visible = False
    grdsales2.SetFocus
End Sub

Private Sub TXTsample2_GotFocus()
    TXTsample2.SelStart = 0
    TXTsample2.SelLength = Len(TXTsample2.text)
End Sub

Private Sub TXTsample2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales2.Col
                Case 3, 4, 6, 7, 8, 9, 10
                    grdsales2.TextMatrix(grdsales2.Row, grdsales2.Col) = TXTsample2.text
                    grdsales2.Enabled = True
                    TXTsample2.Visible = False
                    grdsales2.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample2.Visible = False
            grdsales2.SetFocus
    End Select
        Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample2_KeyPress(KeyAscii As Integer)
    Select Case grdsales2.Col
        Case 2
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
