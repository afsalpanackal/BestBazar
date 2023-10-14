VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMFormula 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productiom Formula Master"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11340
   Icon            =   "FrmFormula.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   11340
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   450
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   7005
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   2970
         Left            =   15
         TabIndex        =   18
         Top             =   30
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   5239
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
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
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   8475
      Left            =   -180
      TabIndex        =   22
      Top             =   0
      Width           =   11520
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00FFC0FF&
         Height          =   615
         Left            =   210
         TabIndex        =   24
         Top             =   -75
         Width           =   11310
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
            Left            =   1485
            TabIndex        =   0
            Top             =   195
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox TxtFormula 
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
            Left            =   4125
            MaxLength       =   100
            TabIndex        =   1
            Top             =   180
            Width           =   3660
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
            Left            =   1485
            TabIndex        =   27
            Top             =   210
            Width           =   900
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Formula Code"
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
            TabIndex        =   26
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name of Mixture"
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
            Left            =   2475
            TabIndex        =   25
            Top             =   225
            Width           =   1605
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
         TabIndex        =   23
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   465
         Left            =   11565
         TabIndex        =   38
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0FF&
         Height          =   5025
         Left            =   210
         TabIndex        =   28
         Top             =   450
         Width           =   11310
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0FF&
            Height          =   6840
            Left            =   11475
            TabIndex        =   29
            Top             =   135
            Visible         =   0   'False
            Width           =   1815
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
               TabIndex        =   36
               Top             =   2970
               Visible         =   0   'False
               Width           =   1515
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
               TabIndex        =   35
               Top             =   3255
               Visible         =   0   'False
               Width           =   1440
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
               TabIndex        =   34
               Top             =   3555
               Visible         =   0   'False
               Width           =   1440
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
               TabIndex        =   33
               Top             =   4035
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
               TabIndex        =   32
               Top             =   4605
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
               TabIndex        =   31
               Top             =   3795
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
               TabIndex        =   30
               Top             =   4365
               Visible         =   0   'False
               Width           =   1395
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4830
            Left            =   45
            TabIndex        =   37
            Top             =   135
            Width           =   11250
            _ExtentX        =   19844
            _ExtentY        =   8520
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
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
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0FF&
         Height          =   3090
         Left            =   210
         TabIndex        =   39
         Top             =   5385
         Width           =   11310
         Begin VB.CommandButton Cmdcancel 
            Caption         =   "Delete &Formula "
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9930
            TabIndex        =   64
            Top             =   855
            Width           =   1320
         End
         Begin VB.TextBox Txtwaste 
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
            Left            =   7800
            MaxLength       =   7
            TabIndex        =   62
            Top             =   435
            Width           =   1095
         End
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
            Left            =   7185
            MaxLength       =   7
            TabIndex        =   60
            Top             =   1620
            Width           =   765
         End
         Begin VB.TextBox Los_Pack 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   4815
            MaxLength       =   7
            TabIndex        =   4
            Top             =   435
            Width           =   705
         End
         Begin VB.ComboBox combopack 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            ItemData        =   "FrmFormula.frx":08CA
            Left            =   5520
            List            =   "FrmFormula.frx":0901
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   435
            Width           =   1140
         End
         Begin VB.ComboBox CmbPack 
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
            ItemData        =   "FrmFormula.frx":0963
            Left            =   5955
            List            =   "FrmFormula.frx":099A
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1605
            Width           =   1200
         End
         Begin VB.TextBox TXTPRODUCT2 
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
            Left            =   90
            MaxLength       =   75
            TabIndex        =   11
            Top             =   1605
            Width           =   4920
         End
         Begin VB.TextBox TxtqtySer 
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
            Left            =   5025
            MaxLength       =   7
            TabIndex        =   13
            Top             =   1605
            Width           =   930
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
            Left            =   5415
            TabIndex        =   15
            Top             =   855
            Width           =   1125
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
            TabIndex        =   46
            Top             =   3060
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
            Left            =   7890
            TabIndex        =   45
            Top             =   3150
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
            TabIndex        =   44
            Top             =   2760
            Visible         =   0   'False
            Width           =   690
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
            TabIndex        =   43
            Top             =   3210
            Visible         =   0   'False
            Width           =   690
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
            Left            =   1275
            TabIndex        =   8
            Top             =   855
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
            Left            =   2475
            TabIndex        =   9
            Top             =   855
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
            Left            =   6615
            TabIndex        =   16
            Top             =   855
            Width           =   1140
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
            Left            =   3675
            TabIndex        =   10
            Top             =   855
            Width           =   1125
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
            Height          =   360
            Left            =   6675
            MaxLength       =   7
            TabIndex        =   6
            Top             =   435
            Width           =   1095
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
            Height          =   360
            Left            =   465
            TabIndex        =   3
            Top             =   435
            Width           =   4335
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
            Height          =   360
            Left            =   60
            TabIndex        =   2
            Top             =   435
            Width           =   390
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
            Left            =   75
            TabIndex        =   7
            Top             =   855
            Width           =   1125
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
            Left            =   11460
            MaxLength       =   6
            TabIndex        =   42
            Top             =   1095
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox TxtCategory 
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
            TabIndex        =   41
            Top             =   780
            Visible         =   0   'False
            Width           =   795
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
            TabIndex        =   40
            Top             =   3225
            Visible         =   0   'False
            Width           =   690
         End
         Begin MSDataListLib.DataList DataList1 
            Height          =   1035
            Left            =   90
            TabIndex        =   12
            Top             =   1980
            Width           =   4920
            _ExtentX        =   8678
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
            Caption         =   "Wastage"
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
            Index           =   4
            Left            =   7800
            TabIndex        =   63
            Top             =   150
            Width           =   1095
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
            Height          =   270
            Index           =   5
            Left            =   7185
            TabIndex        =   61
            Top             =   1350
            Width           =   750
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
            Height          =   285
            Index           =   3
            Left            =   4815
            TabIndex        =   59
            Top             =   150
            Width           =   1830
         End
         Begin VB.Label Label1 
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
            Height          =   315
            Index           =   16
            Left            =   90
            TabIndex        =   58
            Top             =   1350
            Width           =   4920
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
            Height          =   300
            Index           =   2
            Left            =   5025
            TabIndex        =   57
            Top             =   1350
            Width           =   2130
         End
         Begin VB.Label lbldealer2 
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   8385
            TabIndex        =   56
            Top             =   3885
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange2 
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   7680
            TabIndex        =   55
            Top             =   3840
            Visible         =   0   'False
            Width           =   495
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
            TabIndex        =   54
            Top             =   2775
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
            TabIndex        =   53
            Top             =   3075
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
            TabIndex        =   52
            Top             =   2805
            Visible         =   0   'False
            Width           =   1080
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
            TabIndex        =   51
            Top             =   2760
            Visible         =   0   'False
            Width           =   1080
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
            Height          =   285
            Index           =   10
            Left            =   6675
            TabIndex        =   50
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Raw Materials used"
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
            Height          =   300
            Index           =   9
            Left            =   465
            TabIndex        =   49
            Top             =   150
            Width           =   4335
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
            Height          =   285
            Index           =   8
            Left            =   60
            TabIndex        =   48
            Top             =   150
            Width           =   390
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
            TabIndex        =   47
            Top             =   2775
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   -15
      TabIndex        =   21
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11355
      TabIndex        =   20
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   19
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim TMPREC As New ADODB.Recordset
Dim TMPFLAG As Boolean
Dim ACT_REC As New ADODB.Recordset

Dim ACT_FLAG As Boolean
Dim MIX_ITEM As New ADODB.Recordset
Dim MIX_FLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean

Dim CLOSEALL As Integer
Dim M_STOCK As Integer
Dim M_EDIT As Boolean
Dim EDIT_INV As Boolean

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If UCase(txtcategory.text) <> "SERVICE CHARGE" Then
                If CmbPack.ListIndex = -1 Then
                    MsgBox "Please select the Pack", vbOKOnly, "Formula Mixture"
                    CmbPack.Enabled = True
                    CmbPack.SetFocus
                    Exit Sub
                End If
            End If
            Txtpack.Enabled = True
            Txtpack.SetFocus
         Case vbKeyEscape
            TxtqtySer.Enabled = True
            TxtqtySer.SetFocus
    End Select
End Sub

Private Sub CMDADD_Click()
    
    On Error GoTo ErrHand
'    If Val(TXTQTY.Text) = 0 Then
'        MsgBox "Please enter the Qty", vbOKOnly, "Formula Mixture"
'        TXTQTY.SetFocus
'        Exit Sub
'    End If
'
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then Txtwaste.text = 0
    If UCase(txtcategory.text) <> "SERVICE CHARGE" Then
        If Val(Los_Pack.text) = 0 Then
            MsgBox "Please enter the pack", vbOKOnly, "Formula Mixture"
            Los_Pack.SetFocus
            Exit Sub
        End If
        
        If combopack.ListIndex = -1 Then
            MsgBox "Please select the type for the pack", vbOKOnly, "Formula Mixture"
            combopack.Enabled = True
            combopack.SetFocus
            Exit Sub
        End If
    End If
       
    If Trim(TXTITEMCODE.text) = DataList1.BoundText Then
        MsgBox "Can't add same item in Raw materials used.", , "Formula Creation"
        Exit Sub
    End If
    
    If grdsales.rows <= Val(TXTSLNO.text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.text), 0) = Val(TXTSLNO.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 1) = Trim(TXTITEMCODE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 2) = Trim(TXTPRODUCT.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 3) = Val(TXTQTY.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 4) = Trim(txtcategory.text)
    If Val(Los_Pack.text) = 0 Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 5) = "1"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.text), 5) = Val(Los_Pack.text)
    End If
    grdsales.TextMatrix(Val(TXTSLNO.text), 6) = combopack.text
    grdsales.TextMatrix(Val(TXTSLNO.text), 7) = Val(Txtwaste.text)
    TXTSLNO.text = grdsales.rows
    TXTPRODUCT.text = ""
    combopack.ListIndex = -1
    Los_Pack.text = ""
    TXTITEMCODE.text = ""
    TXTVCHNO.text = ""
    TXTLINENO.text = ""
    TXTTRXTYPE.text = ""
    TXTUNIT.text = ""
    TXTQTY.text = ""
    Txtwaste.text = ""
    Label1(10).Caption = "Qty"
    cmdcancel.Enabled = True
    cmdRefresh.Enabled = True
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    M_EDIT = False
    EDIT_INV = True
    
    TXTSLNO.SetFocus
    'grdsales.TopRow = grdsales.Rows - 1
Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub cmdcancel_Click()
    
    On Error GoTo ErrHand
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE THE FORMULA!!!!!", vbYesNo + vbDefaultButton2, "DELETE!!!") = vbNo Then Exit Sub
    db.Execute "delete From TRXFORMULAMAST WHERE TRX_TYPE='FR' AND FOR_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete FROM TRXFORMULASUB WHERE TRX_TYPE='FR' AND FOR_NO = " & Val(txtBillNo.text) & ""

    
    Dim rstBILL As ADODB.Recordset
    Dim i As Integer
    i = 0
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(FOR_NO) FROM TRXFORMULAMAST WHERE TRX_TYPE = 'FR'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    M_EDIT = False
    EDIT_INV = False
    cmdRefresh.Enabled = False
    CMDEXIT.Enabled = True
    CMDPRINT.Enabled = False
    CMDEXIT.Enabled = True
    TXTSLNO.Enabled = False
    'FRMEHEAD.Enabled = True
    TXTQTY.Tag = ""
    'txtremarks.Text = ""
    TxtFormula.text = ""
    TXTPRODUCT2.text = ""
    DataList1.BoundText = ""
    TxtqtySer.text = ""
    Txtpack.text = ""
    CmbPack.ListIndex = -1
    cmdcancel.Enabled = False
    TxtFormula.SetFocus


    MsgBox "DELETED SUCCESSFULLY", , "EzBiz"
    Exit Sub
ErrHand:
    
End Sub

Private Sub CmdDelete_Click()
    Dim i As Long
    
    For i = Val(TXTSLNO.text) To grdsales.rows - 2
        grdsales.TextMatrix(i, 0) = i
        grdsales.TextMatrix(i, 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(i, 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(i, 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(i, 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(i, 6) = grdsales.TextMatrix(i + 1, 6)
        grdsales.TextMatrix(i, 5) = grdsales.TextMatrix(i + 1, 5)
    Next i
    grdsales.rows = grdsales.rows - 1
    
    TXTSLNO.text = Val(grdsales.rows)
    TXTPRODUCT.text = ""
    TXTITEMCODE.text = ""
    combopack.ListIndex = -1
    Los_Pack.text = ""
    TXTVCHNO.text = ""
    TXTLINENO.text = ""
    TXTTRXTYPE.text = ""
    TXTUNIT.text = ""
    TXTQTY.text = ""
    Txtwaste.text = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    cmdcancel.Enabled = True
    cmdRefresh.Enabled = True
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    EDIT_INV = True
    If grdsales.rows = 1 Then
'        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    
End Sub

Private Sub CmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            TXTQTY.text = ""
            Txtwaste.text = ""
            TXTITEMCODE.text = ""
            combopack.ListIndex = -1
            Los_Pack.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TXTUNIT.text = ""
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.text) >= grdsales.rows Then Exit Sub
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = True
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
    
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            TXTQTY.text = ""
            Txtwaste.text = ""
            TXTITEMCODE.text = ""
            combopack.ListIndex = -1
            Los_Pack.text = ""
            txtcategory.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TXTUNIT.text = ""
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub CmdPrint_Click()
    Exit Sub
    If grdsales.rows = 1 Then Exit Sub

    Call Generateprint
    
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    Dim Num As Currency
    
    On Error GoTo ErrHand
    
    db.Execute "delete FROM TRXFORMULASUB WHERE TRX_TYPE='FR' AND FOR_NO = " & Val(txtBillNo.text) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFORMULASUB", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "FR"
        RSTTRXFILE!FOR_NO = Val(txtBillNo.text)
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = Trim(grdsales.TextMatrix(i, 4))
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!WASTE_QTY = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "" 'DataList2.BoundText
               
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Call ReportGeneratION
    
    ReportNameVar = Rptpath & "rptRAWBILL"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "( {TRXFILE.TRX_TYPE}='FR' AND {TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & " )"
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
    
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    
    ''rptPRINT.Action = 1
    Exit Function
ErrHand:
    MsgBox err.Description
End Function

Private Sub CMDPRINT_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            TXTQTY.text = ""
            Txtwaste.text = ""
            TXTITEMCODE.text = ""
            combopack.ListIndex = -1
            Los_Pack.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TXTUNIT.text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub cmdRefresh_Click()
    
   ' If grdsales.Rows = 1 Then GoTo SKIP
    
    On Error GoTo ErrHand
    
    If grdsales.rows <= 1 Then
        Call AppendSale
        Exit Sub
    End If
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFORMULAMAST WHERE FOR_NAME= '" & Trim(TXTPRODUCT2.text) & "' AND FOR_NO <> " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        MsgBox "Already exists", vbOKOnly, "Formula Creation"
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        Exit Sub
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    On Error GoTo ErrHand
    If DataList1.BoundText = "" Then
        MsgBox "Please Select a Product", , "Formula Creation"
        TXTPRODUCT2.SetFocus
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To grdsales.rows - 1
        If Trim(grdsales.TextMatrix(i, 1)) = DataList1.BoundText Then
            MsgBox "Can't add same item in Raw materials used.", , "Formula Creation"
            Exit Sub
        End If
    Next i
    
    If Val(TxtqtySer.text) = 0 Then
        MsgBox "Please Enter Qty for the Product", , "Formula Creation"
        TxtqtySer.SetFocus
        Exit Sub
    End If
    
    If Val(Txtpack.text) = 0 Then Txtpack.text = 1
        
    If CmbPack.ListIndex = -1 Then
        MsgBox "Select Pack", vbOKOnly, "Mixture Creation"
        CmbPack.SetFocus
        Exit Sub
    End If


    Call AppendSale
'    Me.Enabled = False
'    FRMDEBIT.Show
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            TXTQTY.text = ""
            Txtwaste.text = ""
            TXTITEMCODE.text = ""
            combopack.ListIndex = -1
            Los_Pack.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TXTUNIT.text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub combopack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If combopack.ListIndex = -1 Then Exit Sub
            combopack.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            Los_Pack.Enabled = True
            combopack.Enabled = False
            Los_Pack.SetFocus
    End Select
End Sub

Private Sub Form_Activate()
    If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(FOR_NO) FROM TRXFORMULAMAST WHERE TRX_TYPE = 'FR'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    ACT_FLAG = True
    grdsales.ColWidth(0) = 500
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 5000
    grdsales.ColWidth(3) = 1200
    'grdsales.ColWidth(4) = 0
    'grdsales.ColWidth(5) = 0
    
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY /AMT"
    grdsales.TextArray(4) = "CATEGORY"
    grdsales.TextArray(5) = "Unit"
    grdsales.TextArray(6) = "Pack"
    grdsales.TextArray(7) = "Waste"

    PHYFLAG = True
    TMPFLAG = True
    MIX_FLAG = True
    ITEM_FLAG = True
    
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    TXTSLNO.text = 1

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
        If PHYFLAG = False Then PHY.Close
        If TMPFLAG = False Then TMPREC.Close
        If MIX_FLAG = False Then MIX_ITEM.Close
        If ITEM_FLAG = False Then PHY_ITEM.Close
        If ACT_FLAG = False Then ACT_REC.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDPOPUPITEM_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RSTtax As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            TXTPRODUCT.text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.text = GRDPOPUPITEM.Columns(0)
            txtcategory.text = IIf(IsNull(GRDPOPUPITEM.Columns(4)), "", GRDPOPUPITEM.Columns(4))
            Los_Pack.text = IIf(IsNull(GRDPOPUPITEM.Columns(2)), "", GRDPOPUPITEM.Columns(2))
            On Error Resume Next
            combopack.text = IIf(IsNull(GRDPOPUPITEM.Columns(3)), "", GRDPOPUPITEM.Columns(3))
            On Error GoTo ErrHand
            i = 0
            For i = 1 To grdsales.rows - 1
                If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
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
            'TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            'TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
            Else
                Los_Pack.Enabled = True
                Los_Pack.SetFocus
            End If
            Exit Sub
        Case vbKeyEscape
            TXTQTY.text = ""
            Txtwaste.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TXTUNIT.text = ""
            
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim TRXSUB As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    
    Dim i As Long
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtBillNo.text) = 0 Then Exit Sub
            
            cmdcancel.Enabled = True
            grdsales.rows = 1
            
            Set TRXSUB = New ADODB.Recordset
            TRXSUB.Open "Select * From TRXFORMULAMAST WHERE TRX_TYPE='FR' AND FOR_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
            If Not (TRXSUB.EOF And TRXSUB.BOF) Then
                TxtFormula.text = TRXSUB!FOR_NAME
                'txtremarks.Text = IIf(IsNull(TRXSUB!REMARKS), "", TRXSUB!REMARKS)
                TXTPRODUCT2.text = IIf(IsNull(TRXSUB!ITEM_NAME), "", TRXSUB!ITEM_NAME)
                TxtqtySer.text = IIf(IsNull(TRXSUB!QTY), "", TRXSUB!QTY)
                Txtpack.text = IIf(IsNull(TRXSUB!LOOSE_PACK), "1", TRXSUB!LOOSE_PACK)
                On Error Resume Next
                CmbPack.text = IIf(IsNull(TRXSUB!PACK_TYPE), CmbPack.ListIndex - 1, TRXSUB!PACK_TYPE)
                On Error GoTo ErrHand
                i = 0
                Set TRXFILE = New ADODB.Recordset
                TRXFILE.Open "Select * FROM TRXFORMULASUB WHERE TRX_TYPE='FR' AND FOR_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO ", db, adOpenStatic, adLockReadOnly
                Do Until TRXFILE.EOF
                    i = i + 1
                    grdsales.rows = grdsales.rows + 1
                    grdsales.FixedRows = 1
                    grdsales.TextMatrix(i, 0) = i
                    grdsales.TextMatrix(i, 1) = TRXFILE!ITEM_CODE
                    grdsales.TextMatrix(i, 2) = TRXFILE!ITEM_NAME
                    grdsales.TextMatrix(i, 3) = TRXFILE!QTY
                    grdsales.TextMatrix(i, 4) = IIf(IsNull(TRXFILE!Category), "", TRXFILE!Category)
                    grdsales.TextMatrix(i, 5) = IIf(IsNull(TRXFILE!LOOSE_PACK), "1", TRXFILE!LOOSE_PACK)
                    grdsales.TextMatrix(i, 6) = IIf(IsNull(TRXFILE!PACK_TYPE), "", TRXFILE!PACK_TYPE)
                    grdsales.TextMatrix(i, 7) = IIf(IsNull(TRXFILE!WASTE_QTY), 0, TRXFILE!WASTE_QTY)
                    TRXFILE.MoveNext
                Loop
                TRXFILE.Close
                Set TRXFILE = Nothing
            End If
            TRXSUB.Close
            Set TRXSUB = Nothing
            
            LBLBILLNO.Caption = Val(txtBillNo.text)
            
            TXTSLNO.text = grdsales.rows
            txtBillNo.Visible = False
            TXTSLNO.Enabled = True
            
            If grdsales.rows > 1 Then
                TXTSLNO.SetFocus
            Else
                TXTSLNO.Enabled = False
                TxtFormula.SetFocus
            End If
            DataList1.text = TXTPRODUCT2.text
            TXTPRODUCT2.text = DataList1.text
            LBLDEALER2.Caption = TXTPRODUCT2.text
    End Select

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
    TRXMAST.Open "Select MAX(FOR_NO) FROM TRXFORMULASUB WHERE TRX_TYPE = 'FR'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        If Val(txtBillNo.text) > i Then
            MsgBox "The last No. is " & i, vbCritical, "BILL..."
            txtBillNo.Visible = True
            txtBillNo.SetFocus
            Exit Sub
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
      
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MIN(FOR_NO) FROM TRXFORMULASUB WHERE TRX_TYPE = 'FR'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
        If Val(txtBillNo.text) < i Then
            MsgBox "This Starting No. is " & i, vbCritical, "BILL..."
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

Private Sub TXTPRODUCT_GotFocus()
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.text)
    Label1(10).Caption = "Qty"
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim RSTZEROSTOCK As ADODB.Recordset
    Dim RSTBALQTY As ADODB.Recordset
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case 106
            If TXTQTY.Tag <> "" Then
                TXTPRODUCT.text = Trim(TXTQTY.Tag)
                TXTPRODUCT.SelStart = 0
                TXTPRODUCT.SelLength = Len(TXTPRODUCT.text)
            End If
        Case vbKeyReturn
            M_STOCK = 0
            If Trim(TXTPRODUCT.text) = "" Then Exit Sub
            CmdDelete.Enabled = False
            TXTQTY.text = ""
            Txtwaste.text = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
           
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                'PHY.Open "Select ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHY.Open "Select ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                'PHY.Open "Select ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHY.Open "Select ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                txtcategory.text = IIf(IsNull(grdtmp.Columns(4)), "", grdtmp.Columns(4))
                TXTITEMCODE.text = grdtmp.Columns(0)
                TXTPRODUCT.text = grdtmp.Columns(1)
                Los_Pack.text = IIf(IsNull(grdtmp.Columns(2)), "", grdtmp.Columns(2))
                On Error Resume Next
                combopack.text = IIf(IsNull(grdtmp.Columns(3)), combopack.ListIndex = -1, grdtmp.Columns(3))
                On Error GoTo ErrHand
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i
                
                If PHY.RecordCount = 1 Then
                    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                        TXTPRODUCT.Enabled = False
                        TXTQTY.Enabled = True
                        TXTQTY.SetFocus
                    Else
                        TXTPRODUCT.Enabled = False
                        Los_Pack.Enabled = True
                        Los_Pack.SetFocus
                    End If
                    Exit Sub
                End If
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
        
            CmdDelete.Enabled = False
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
        
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.text)
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        Label1(10).Caption = "Amount"
    Else
        Label1(10).Caption = "Qty"
    End If
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.text) = 0 And UCase(txtcategory.text) <> "SERVICE CHARGE" Then
                MsgBox "Please enter the qty", vbOKOnly, "Formula Master"
                Exit Sub
            End If
            If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                TXTQTY.Enabled = False
                cmdadd.Enabled = True
                cmdadd.SetFocus
            Else
                TXTQTY.Enabled = False
                Txtwaste.Enabled = True
                Txtwaste.SetFocus
            End If
         Case vbKeyEscape
            'If M_EDIT = True Then Exit Sub
            If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                If M_EDIT = True Then Exit Sub
                TXTQTY.Enabled = False
                TXTPRODUCT.Enabled = True
                TXTPRODUCT.SetFocus
            Else
                TXTQTY.Enabled = False
                combopack.Enabled = True
                combopack.SetFocus
            End If
            Exit Sub
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TXTUNIT.text = ""
            
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    TXTQTY.text = Format(TXTQTY.text, ".000")
End Sub

Private Sub TxtqtySer_GotFocus()
    TxtqtySer.SelStart = 0
    TxtqtySer.SelLength = Len(TxtqtySer.text)
End Sub

Private Sub TxtqtySer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmbPack.Enabled = True
            CmbPack.SetFocus
         Case vbKeyEscape
            TXTPRODUCT2.Enabled = True
            TXTPRODUCT2.SetFocus
    End Select
End Sub

Private Sub TxtqtySer_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtpack_GotFocus()
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.text)
    If Val(Txtpack.text) = 0 Then Txtpack.text = 1
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdRefresh.Enabled = True
            cmdRefresh.SetFocus
         Case vbKeyEscape
            CmbPack.Enabled = True
            CmbPack.SetFocus
    End Select
End Sub

Private Sub Txtpack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.text)
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.text) = 0 Then
                TXTSLNO.text = ""
                TXTPRODUCT.text = ""
                TXTQTY.text = ""
                Txtwaste.text = ""
                TXTITEMCODE.text = ""
                combopack.ListIndex = -1
                
                Los_Pack.text = ""
                TXTVCHNO.text = ""
                TXTLINENO.text = ""
                TXTTRXTYPE.text = ""
                TXTUNIT.text = ""
                TXTSLNO.text = grdsales.rows
                CmdDelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.text) >= grdsales.rows Then
                TXTSLNO.text = grdsales.rows
                CmdDelete.Enabled = False
                CMDMODIFY.Enabled = False
            End If
            If Val(TXTSLNO.text) < grdsales.rows Then
                TXTSLNO.text = grdsales.TextMatrix(Val(TXTSLNO.text), 0)
                TXTITEMCODE.text = grdsales.TextMatrix(Val(TXTSLNO.text), 1)
                TXTPRODUCT.text = grdsales.TextMatrix(Val(TXTSLNO.text), 2)
                TXTQTY.text = grdsales.TextMatrix(Val(TXTSLNO.text), 3)
                txtcategory.text = grdsales.TextMatrix(Val(TXTSLNO.text), 4)
                Los_Pack.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
                On Error Resume Next
                combopack.text = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 6))
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.text = Val(grdsales.rows)
                TXTPRODUCT.text = ""
                TXTITEMCODE.text = ""
                combopack.ListIndex = -1
                
                Los_Pack.text = ""
                TXTVCHNO.text = ""
                TXTLINENO.text = ""
                txtcategory.text = ""
                TXTTRXTYPE.text = ""
                TXTUNIT.text = ""
                TXTQTY.text = ""
                Txtwaste.text = ""
                
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.rows > 1 Then
                TXTSLNO.Enabled = False
                CMDPRINT.Enabled = True
                cmdRefresh.Enabled = True
                cmdRefresh.SetFocus
            Else
                TXTSLNO.Enabled = False
                'FRMEHEAD.Enabled = True
                TxtFormula.Enabled = True
                TxtFormula.SetFocus
            End If
    End Select
End Sub

Private Sub TXTSLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case vbKeyTab
            KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Function FILL_ITEMGRID()
    FRMEMAIN.Enabled = False
    FRMEITEM.Visible = True
    Set GRDPOPUPITEM.DataSource = Nothing
    
    
    If ITEM_FLAG = True Then
        'PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, CATEGORY  From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, CATEGORY  From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.text & "%' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, CATEGORY  From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.text & "%' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
     End If

    Set GRDPOPUPITEM.DataSource = PHY_ITEM
    'GRDPOPUPITEM.RowHeight = 250
    GRDPOPUPITEM.Columns(0).Visible = False
    GRDPOPUPITEM.Columns(1).Caption = "ITEM NAME"
    GRDPOPUPITEM.Columns(1).Width = 4800
    GRDPOPUPITEM.SetFocus
End Function

Private Function AppendSale()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTP_RATE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim i As Double
    Dim TRXVALUE As Double
    
    i = 0
    On Error GoTo ErrHand
    db.BeginTrans
    db.Execute "delete From TRXFORMULAMAST WHERE TRX_TYPE='FR' AND FOR_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete FROM TRXFORMULASUB WHERE TRX_TYPE='FR' AND FOR_NO = " & Val(txtBillNo.text) & ""
    
    If grdsales.rows <= 1 Then GoTo SKIP
    Dim ITEMCOST As Double
    Dim rstTRXMAST As ADODB.Recordset
    ITEMCOST = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFORMULASUB", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "FR"
        RSTTRXFILE!FOR_NO = Val(txtBillNo.text)
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!FOR_NAME = DataList1.BoundText
        RSTTRXFILE!Category = Trim(grdsales.TextMatrix(i, 4))
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!WASTE_QTY = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!OWN_QTY = Val(TxtqtySer.text)
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 5))
        RSTTRXFILE!PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "" 'DataList2.BoundText
        
        'ITEMCOST = 0
        Set rstTRXMAST = New ADODB.Recordset
        rstTRXMAST.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        With rstTRXMAST
            If Not (.EOF And .BOF) Then
                'ITEMCOST = ITEMCOST + IIf(IsNull(!ITEM_COST), 0, !ITEM_COST * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)))
                ITEMCOST = ITEMCOST + IIf(IsNull(!ITEM_COST), 0, !ITEM_COST) / (Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5)))
            End If
        End With
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
        
        RSTTRXFILE.Update
    Next i
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
'    Dim RSTITEMCOST As ADODB.Recordset
'    Set RSTITEMCOST = New ADODB.Recordset
'    RSTITEMCOST.Open "Select * FROM ITEMMAST WHERE ITEM_CODE = '" & DataList1.BoundText & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
'    If Not (RSTITEMCOST.EOF And RSTITEMCOST.BOF) Then
'        RSTITEMCOST!ITEM_COST = ITEMCOST
'        RSTITEMCOST.Update
'    End If
'    RSTITEMCOST.Close
'    Set RSTITEMCOST = Nothing
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFORMULAMAST WHERE TRX_TYPE='FR' AND FOR_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!FOR_NO = Val(txtBillNo.text)
        RSTTRXFILE!TRX_TYPE = "FR"
    End If
    RSTTRXFILE!LINE_NO = 1
    RSTTRXFILE!FOR_NAME = Trim(TXTPRODUCT2.text)
    RSTTRXFILE!ITEM_CODE = DataList1.BoundText
    RSTTRXFILE!ITEM_NAME = Trim(TXTPRODUCT2.text)
    RSTTRXFILE!QTY = Val(TxtqtySer.text)
    RSTTRXFILE!LOOSE_PACK = Val(Txtpack.text)
    RSTTRXFILE!PACK_TYPE = CmbPack.text
    'RSTTRXFILE!REMARKS = Trim(txtremarks.Text)
    RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!C_USER_ID = "SM"
    
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.CommitTrans
SKIP:
    i = 0
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(FOR_NO) FROM TRXFORMULAMAST WHERE TRX_TYPE = 'FR'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    M_EDIT = False
    EDIT_INV = False
    cmdRefresh.Enabled = False
    CMDEXIT.Enabled = True
    CMDPRINT.Enabled = False
    CMDEXIT.Enabled = True
    TXTSLNO.Enabled = False
    'FRMEHEAD.Enabled = True
    TXTQTY.Tag = ""
    'txtremarks.Text = ""
    TxtFormula.text = ""
    TXTPRODUCT2.text = ""
    DataList1.BoundText = ""
    TxtqtySer.text = ""
    Txtpack.text = ""
    CmbPack.ListIndex = -1
    cmdcancel.Enabled = False
    TxtFormula.SetFocus
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
'    RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY FOR_NO", Conn, adOpenForwardOnly
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

Private Sub TxtFormula_GotFocus()
    TxtFormula.SelStart = 0
    TxtFormula.SelLength = Len(TxtFormula.text)
End Sub

Private Sub TxtFormula_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'FRMEHEAD.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            If EDIT_INV = True Then Exit Sub
            txtBillNo.Visible = True
            txtBillNo.SetFocus
    End Select
End Sub

Private Sub TxtFormula_KeyPress(KeyAscii As Integer)
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
            MIX_ITEM.Open "select ITEM_CODE, ITEM_NAME from ITEMMAST where ITEM_NAME Like '%" & TXTPRODUCT2.text & "%' AND ucase(CATEGORY) = 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            MIX_FLAG = False
        Else
            MIX_ITEM.Close
            MIX_ITEM.Open "select ITEM_CODE, ITEM_NAME from ITEMMAST  where ITEM_NAME Like '%" & TXTPRODUCT2.text & "%' AND ucase(CATEGORY) = 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            MIX_FLAG = False
        End If
        If (MIX_ITEM.EOF And MIX_ITEM.BOF) Then
            LBLDEALER2.Caption = ""
        Else
            LBLDEALER2.Caption = MIX_ITEM!ITEM_NAME
        End If
        Set DataList1.RowSource = MIX_ITEM
        DataList1.ListField = "ITEM_NAME"
        DataList1.BoundColumn = "ITEM_CODE"
       
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
            If DataList1.VisibleCount = 0 Then
                If MsgBox("Item not exists!!! Do You want to add this item?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
                Dim RSTITEMMAST As ADODB.Recordset
                TXTPRODUCT.Tag = ""
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
                If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                    If IsNull(RSTITEMMAST.Fields(0)) Then
                        TXTPRODUCT.Tag = 1
                    Else
                        TXTPRODUCT.Tag = Val(RSTITEMMAST.Fields(0)) + 1
                    End If
                End If
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & TXTPRODUCT.Tag & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                db.BeginTrans
                RSTITEMMAST.AddNew
                'RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
                RSTITEMMAST!ITEM_CODE = TXTPRODUCT.Tag
                RSTITEMMAST!ITEM_NAME = Trim(TXTPRODUCT2.text)
                RSTITEMMAST!Category = "OWN"
                RSTITEMMAST!UNIT = 1
                RSTITEMMAST!MANUFACTURER = "GENERAL"
                RSTITEMMAST!DEAD_STOCK = "N"
                RSTITEMMAST!REMARKS = ""
                RSTITEMMAST!REORDER_QTY = 1
                RSTITEMMAST!PACK_TYPE = "Nos"
                RSTITEMMAST!FULL_PACK = "Nos"
                RSTITEMMAST!BIN_LOCATION = ""
                RSTITEMMAST!MRP = 0
                RSTITEMMAST!PTR = 0
                RSTITEMMAST!CST = 0
                RSTITEMMAST!OPEN_QTY = 0
                RSTITEMMAST!OPEN_VAL = 0
                RSTITEMMAST!RCPT_QTY = 0
                RSTITEMMAST!RCPT_VAL = 0
                RSTITEMMAST!ISSUE_QTY = 0
                RSTITEMMAST!ISSUE_VAL = 0
                RSTITEMMAST!CLOSE_QTY = 0
                RSTITEMMAST!CLOSE_VAL = 0
                RSTITEMMAST!DAM_QTY = 0
                RSTITEMMAST!DAM_VAL = 0
                RSTITEMMAST!DISC = 0
                RSTITEMMAST!SALES_TAX = 0
                RSTITEMMAST!ITEM_COST = 0
                RSTITEMMAST!P_RETAIL = 0
                RSTITEMMAST!P_WS = 0
                RSTITEMMAST!CRTN_PACK = 1
                RSTITEMMAST!P_CRTN = 0
                RSTITEMMAST!LOOSE_PACK = 1
                RSTITEMMAST!UN_BILL = "N"
                RSTITEMMAST.Update
                db.CommitTrans
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                TXTITEMCODE.text = TXTPRODUCT.Tag
                Call TxtProduct2_Change
                'frmitemmaster.Show
                'frmitemmaster.TXTITEM.Text = Trim(TXTPRODUCT2.Text)
                'frmitemmaster.LBLLP.Caption = "P"
                'MsgBox "Item not found!!!!", , "EzBiz"
                Exit Sub
            End If
            DataList1.Enabled = True
            DataList1.SetFocus
    End Select
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans

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
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            If DataList1.text = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Product from the List", vbOKOnly, "Mixture Creation"
                DataList1.SetFocus
                Exit Sub
            End If
            'TXTPRODUCT2.Enabled = False
            'DataList1.Enabled = False
            TxtqtySer.Enabled = True
            TxtqtySer.SetFocus
            
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
End Sub

Private Sub OptLoose_Click()
    On Error Resume Next
    TXTQTY.SetFocus
End Sub

Private Sub OptLoose_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub OptNormal_Click()
    On Error Resume Next
    TXTQTY.SetFocus
End Sub

Private Sub OptNormal_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub Los_Pack_GotFocus()
    Los_Pack.SelStart = 0
    Los_Pack.SelLength = Len(Los_Pack.text)
End Sub

Private Sub Los_Pack_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            If Val(Los_Pack.text) = 0 Then Exit Sub
            Los_Pack.Enabled = False
            combopack.Enabled = True
            combopack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TXTUNIT.text = ""
            
            TXTPRODUCT.Enabled = True
            Los_Pack.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub Los_Pack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Los_Pack_LostFocus()
    Los_Pack.text = Format(Los_Pack.text, ".000")
End Sub

Private Sub Txtwaste_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtwaste.text) >= Val(TXTQTY.text) * Val(Los_Pack.text) Then
                MsgBox "Wastage could not be greater than qty", vbOKOnly, "Formula Master"
                Txtwaste.SetFocus
                Exit Sub
            End If
            Txtwaste.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            'If M_EDIT = True Then Exit Sub
            Txtwaste.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub Txtwaste_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtwaste_LostFocus()
    Txtwaste.text = Format(Txtwaste.text, ".000")
End Sub
