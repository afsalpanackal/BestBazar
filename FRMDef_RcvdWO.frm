VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMDef_RcvdWO 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goods Under Warranty Received"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13560
   ControlBox      =   0   'False
   Icon            =   "FRMDef_RcvdWO.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   13560
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   1440
      TabIndex        =   62
      Top             =   3870
      Visible         =   0   'False
      Width           =   7425
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   2835
         Left            =   75
         TabIndex        =   63
         Top             =   105
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   5001
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
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   1440
      TabIndex        =   58
      Top             =   3195
      Visible         =   0   'False
      Width           =   7440
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2535
         Left            =   90
         TabIndex        =   59
         Top             =   360
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   4471
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
         Index           =   2
         Left            =   3135
         TabIndex        =   61
         Top             =   105
         Width           =   4215
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
         Index           =   1
         Left            =   90
         TabIndex        =   60
         Top             =   105
         Width           =   3045
      End
   End
   Begin VB.Frame FRMEGRDBILL 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   735
      TabIndex        =   50
      Top             =   4110
      Visible         =   0   'False
      Width           =   8820
      Begin MSDataGridLib.DataGrid GRDPOPUPBILL 
         Height          =   2535
         Left            =   90
         TabIndex        =   51
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4471
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
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   "BILL DETAILS FOR THE ITEM"
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
         Width           =   2685
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
         Left            =   2775
         TabIndex        =   52
         Top             =   105
         Width           =   5970
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   -150
      TabIndex        =   13
      Top             =   15
      Width           =   10560
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00C0FFC0&
         Height          =   2235
         Left            =   210
         TabIndex        =   14
         Top             =   -105
         Width           =   10305
         Begin VB.OptionButton OPTCUSTOMER 
            BackColor       =   &H00C0FFC0&
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   105
            TabIndex        =   57
            Top             =   735
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton OPTSELF 
            BackColor       =   &H00C0FFC0&
            Caption         =   "SELF"
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
            Height          =   255
            Left            =   1605
            TabIndex        =   56
            Top             =   735
            Width           =   1395
         End
         Begin MSDataListLib.DataCombo cmbinv 
            Height          =   330
            Left            =   165
            TabIndex        =   48
            Top             =   1845
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   255
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
            Left            =   165
            TabIndex        =   0
            Top             =   1095
            Width           =   4575
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
            Left            =   510
            TabIndex        =   11
            Top             =   225
            Width           =   975
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   345
            Left            =   2160
            TabIndex        =   37
            Top             =   225
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
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
            Left            =   150
            TabIndex        =   45
            Top             =   1440
            Width           =   4560
            _ExtentX        =   8043
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
         Begin VB.Label LBLSALEFLAG 
            Height          =   540
            Left            =   9105
            TabIndex        =   64
            Top             =   300
            Visible         =   0   'False
            Width           =   750
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
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   22
            Left            =   4785
            TabIndex        =   42
            Top             =   1260
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
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   21
            Left            =   4860
            TabIndex        =   41
            Top             =   1785
            Width           =   360
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
            Left            =   5625
            TabIndex        =   40
            Top             =   1755
            Width           =   2400
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
            Left            =   5610
            TabIndex        =   39
            Top             =   1050
            Width           =   4635
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   1605
            TabIndex        =   38
            Top             =   255
            Width           =   525
         End
         Begin VB.Label LblInvoice 
            BackStyle       =   0  'Transparent
            Caption         =   "No."
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
            Index           =   0
            Left            =   135
            TabIndex        =   18
            Top             =   255
            Width           =   375
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
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   4965
            TabIndex        =   17
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
            Left            =   5580
            TabIndex        =   16
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
            Left            =   6795
            TabIndex        =   15
            Top             =   225
            Width           =   1110
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
         TabIndex        =   36
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Height          =   6060
         Left            =   210
         TabIndex        =   19
         Top             =   2070
         Width           =   10320
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   5730
            Left            =   90
            TabIndex        =   12
            Top             =   270
            Width           =   10155
            _ExtentX        =   17912
            _ExtentY        =   10107
            _Version        =   393216
            Rows            =   1
            Cols            =   13
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
            SelectionMode   =   1
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
         Left            =   11100
         TabIndex        =   35
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
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00C0FFC0&
         Height          =   1365
         Left            =   210
         TabIndex        =   20
         Top             =   8070
         Width           =   10335
         Begin VB.TextBox TxtPack 
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
            Left            =   345
            MaxLength       =   6
            TabIndex        =   54
            Top             =   1020
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.CommandButton cmdview 
            Caption         =   "&VIEW"
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
            Left            =   6600
            TabIndex        =   49
            Top             =   810
            Width           =   1125
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
            Left            =   9180
            TabIndex        =   43
            Top             =   810
            Width           =   420
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
            Height          =   465
            Left            =   660
            TabIndex        =   5
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
            TabIndex        =   1
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
            Left            =   630
            TabIndex        =   2
            Top             =   450
            Width           =   4065
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
            Left            =   4710
            MaxLength       =   7
            TabIndex        =   3
            Top             =   450
            Width           =   765
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
            Left            =   4245
            TabIndex        =   8
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
            Left            =   7800
            TabIndex        =   10
            Top             =   810
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
            Left            =   3060
            TabIndex        =   7
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
            Left            =   1845
            TabIndex        =   6
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
            TabIndex        =   25
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
            Left            =   5490
            MaxLength       =   15
            TabIndex        =   4
            Top             =   450
            Width           =   3405
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            Left            =   5430
            TabIndex        =   9
            Top             =   810
            Width           =   1125
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
            Height          =   225
            Index           =   0
            Left            =   345
            TabIndex        =   55
            Top             =   795
            Visible         =   0   'False
            Width           =   525
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
            Left            =   45
            TabIndex        =   34
            Top             =   225
            Width           =   570
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            Caption         =   " Product Name"
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
            Left            =   630
            TabIndex        =   33
            Top             =   225
            Width           =   4065
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
            Left            =   4710
            TabIndex        =   32
            Top             =   225
            Width           =   765
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
            TabIndex        =   31
            Top             =   1260
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            Caption         =   " Serial No."
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
            Left            =   5490
            TabIndex        =   30
            Top             =   225
            Width           =   3405
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   1275
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
   Begin MSDataListLib.DataCombo CMBDISTI 
      Height          =   1020
      Left            =   6660
      TabIndex        =   44
      Top             =   1275
      Visible         =   0   'False
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   1773
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      ForeColor       =   255
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11655
      TabIndex        =   47
      Top             =   2610
      Width           =   495
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11445
      TabIndex        =   46
      Top             =   3255
      Width           =   1620
   End
End
Attribute VB_Name = "FRMDef_RcvdWO"
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
Dim PHY_BILL As New ADODB.Recordset
Dim BILL_FLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean
Dim INV_FLAG As Boolean
Dim INV_REC As New ADODB.Recordset
Dim PHY_BATCH As New ADODB.Recordset
Dim BATCH_FLAG As Boolean

Dim CLOSEALL As Integer
Dim M_STOCK As Double
Dim EDIT_BILL As Boolean
Dim M_EDIT As Boolean
Dim B_FLAG As Boolean
Dim CN_FLAG As Boolean
Dim M_DELETE As Boolean

Private Sub cmbinv_Change()
    txtBillNo.Text = cmbinv.Text
    Call VIEWGRID
    FRMEMASTER.Enabled = True
End Sub

Private Sub CMDADD_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim I As Integer
    
    On Error GoTo eRRhAND
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Val(TxtPack.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Trim(TXTTRXTYPE.Text)
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = Trim(TXTCOMAMT.Text)
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMASTWO WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db2, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = IIf(IsNull(RSTTRXFILE!MANUFACTURER), "", Trim(RSTTRXFILE!MANUFACTURER))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = "N"
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Val(TXTQTY.Tag)
    
    If OPTSELF.Value = True Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 6) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                '!ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
                !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILEWO WHERE RTRXFILEWO.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 7)) & "' AND RTRXFILEWO.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)) & " AND RTRXFILEWO.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) & " AND BAL_QTY > 0", db2, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        LBLSALEFLAG = "Y"
    Else
        LBLSALEFLAG = "N"
    End If
    
    CN_FLAG = False
    
SKIP:
    'Call STOCKADJUST
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TxtPack.Text = ""
    
    TXTQTY.Text = ""
    txtBatch.Text = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    M_EDIT = True
    TXTSLNO.SetFocus
    'grdsales.TopRow = grdsales.Rows - 1
Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim I As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    On Error GoTo eRRhAND
    If OPTSELF.Value = True Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 6) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILEWO WHERE RTRXFILEWO.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 7)) & "' AND RTRXFILEWO.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)) & " AND RTRXFILEWO.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) & " AND BAL_QTY > 0", db2, adOpenStatic, adLockOptimistic, adCmdText
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
    End If
    
    For I = Val(TXTSLNO.Text) - 1 To grdsales.Rows - 2
        grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = I
        grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = grdsales.TextMatrix(I + 1, 1)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = grdsales.TextMatrix(I + 1, 2)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = grdsales.TextMatrix(I + 1, 3)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = grdsales.TextMatrix(I + 1, 4)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = grdsales.TextMatrix(I + 1, 5)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(I + 1, 6)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = grdsales.TextMatrix(I + 1, 7)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = grdsales.TextMatrix(I + 1, 8)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = grdsales.TextMatrix(I + 1, 9)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = grdsales.TextMatrix(I + 1, 10)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = grdsales.TextMatrix(I + 1, 11)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = grdsales.TextMatrix(I + 1, 12)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = grdsales.TextMatrix(I + 1, 13)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = grdsales.TextMatrix(I + 1, 14)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = grdsales.TextMatrix(I + 1, 15)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = grdsales.TextMatrix(I + 1, 16)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = grdsales.TextMatrix(I + 1, 17)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = grdsales.TextMatrix(I + 1, 18)
    Next I
    grdsales.Rows = grdsales.Rows - 1
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TxtPack.Text = ""
    TXTQTY.Text = ""
    txtBatch.Text = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = True
    If grdsales.Rows = 1 Then
'        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    M_DELETE = True
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub CMDEXIT_Click()
    If CMDEXIT.Caption = "E&XIT" Then
        CLOSEALL = 0
        Unload Me
    Else
        FRMEMASTER.Enabled = True
        txtBillNo.Enabled = True
        txtBillNo.SetFocus
        CMDEXIT.Caption = "E&XIT"
        TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
        DataList2.Text = ""
        lbladdress.Caption = ""
        lbltin.Caption = ""
        LBLDATE.Caption = Date
        LBLTIME.Caption = Time
        grdsales.Rows = 1
        TXTSLNO.Text = 1
        M_EDIT = False
        cmdRefresh.Enabled = False
        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        CMDEXIT.Enabled = True
        TXTQTY.Tag = ""
        cmdview.Enabled = True
        LblInvoice(0).Top = 240
        TXTDEALER.Top = 600
        DataList2.Top = 945
        cmbinv.Visible = False
        TXTDEALER.Text = ""
        TXTDEALER.Enabled = False
        DataList2.Enabled = False
        TXTINVDATE.Enabled = False
    End If
End Sub

Private Sub CMDMODIFY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    On Error GoTo eRRhAND
    If OPTSELF.Value = True Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 6) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILEWO WHERE RTRXFILEWO.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 7)) & "' AND RTRXFILEWO.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)) & " AND RTRXFILEWO.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) & " AND BAL_QTY > 0", db2, adOpenStatic, adLockOptimistic, adCmdText
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
    End If
    
    
    TXTQTY.Tag = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = True
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TxtPack.Text = ""
            
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            txtBatch.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub cmdPrint_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim I As Integer
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
    
    Exit Sub
    db2.Execute "delete * From TRXFILEWO"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFILEWO", db2, adOpenStatic, adLockOptimistic, adCmdText
    For I = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        Set TRXMAST = New ADODB.Recordset
        TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMASTWO WHERE ITEMMASTWO.ITEM_CODE = '" & Trim(grdsales.TextMatrix(I, 12)) & "'", db2, adOpenStatic, adLockReadOnly
        If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
            RSTTRXFILE!MFGR = TRXMAST!MANUFACTURER
        End If
        TRXMAST.Close
        Set TRXMAST = Nothing
        
        RSTTRXFILE!TRX_TYPE = "DN"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = I
        RSTTRXFILE!CATEGORY = "MEDICINE"
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(I, 12)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(I, 2)
        RSTTRXFILE!QTY = grdsales.TextMatrix(I, 3)
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = grdsales.TextMatrix(I, 5)
        RSTTRXFILE!PTR = grdsales.TextMatrix(I, 6)
        RSTTRXFILE!SALES_PRICE = grdsales.TextMatrix(I, 6)
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(I, 8)
        RSTTRXFILE!UNIT = 1
        RSTTRXFILE!VCH_DESC = "Issued to     " & DataList2.Text
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(I, 9)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(I, 11)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
GOSKIP:
    Next I

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    
    txtBatch.Enabled = False

    
End Sub

Private Sub cmdRefresh_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTTRXFILESUB As ADODB.Recordset
    Dim I As Double
    
    Dim DAY_DATE As String
    Dim MONTH_DATE As String
    Dim YEAR_DATE As String
    Dim E_DATE As Date
    
    'If grdsales.Rows = 1 Then GoTo SKIP
    
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
    
    If (OPTCUSTOMER.Value = True And IsNull(DataList2.SelectedItem)) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
        DataList2.SetFocus
        Exit Sub
    End If
    
    I = 0
    On Error GoTo eRRhAND
    
    db2.Execute "delete * From WAR_TRXFILE WHERE VCH_NO = " & Val(txtBillNo.Text) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From WAR_TRXFILE", db2, adOpenStatic, adLockOptimistic, adCmdText
    For I = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = I
        RSTTRXFILE!TRX_TYPE = "CN"
        RSTTRXFILE!CATEGORY = "MEDICINE"
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(I, 1)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(I, 2)
        RSTTRXFILE!QTY = grdsales.TextMatrix(I, 3)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(I, 4)
        If OPTCUSTOMER.Value = True Then
            RSTTRXFILE!VCH_DESC = "Received from " & Trim(DataList2.Text)
        Else
            RSTTRXFILE!VCH_DESC = "Issued from   Stock"
        End If
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(I, 5)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        If OPTCUSTOMER.Value = True Then
            RSTTRXFILE!ACT_CODE = DataList2.BoundText
            RSTTRXFILE!ACT_NAME = DataList2.Text
        Else
            RSTTRXFILE!ACT_CODE = "111111"
            RSTTRXFILE!ACT_NAME = "SELF"
        End If
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!CHECK_FLAG = "N"

        RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(I, 7) = "", 0, grdsales.TextMatrix(I, 7))
        RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(I, 8) = "", 0, grdsales.TextMatrix(I, 8))
        RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(I, 9) = "", "WN", grdsales.TextMatrix(I, 9))
        RSTTRXFILE!ISSUEQTY = Val(grdsales.TextMatrix(I, 12))
        'RSTTRXFILE!COM_AMT = IIf(grdsales.TextMatrix(i, 19) = "", 0, grdsales.TextMatrix(i, 19))
        
        Set RSTTRXFILESUB = New ADODB.Recordset
        RSTTRXFILESUB.Open "Select * From RTRXFILEWO WHERE VCH_NO = " & RSTTRXFILE!R_VCH_NO & " AND LINE_NO = " & RSTTRXFILE!R_LINE_NO & " AND TRX_TYPE = '" & RSTTRXFILE!R_TRX_TYPE & "' ", db2, adOpenStatic, adLockReadOnly
        If Not (RSTTRXFILESUB.EOF And RSTTRXFILESUB.BOF) Then
            Select Case RSTTRXFILESUB!TRX_TYPE
                Case "PI"
                    RSTTRXFILE!DIST_NAME = IIf(IsNull(RSTTRXFILESUB!VCH_DESC), "", Mid(RSTTRXFILESUB!VCH_DESC, 15))
                    RSTTRXFILE!BILL_NO = IIf(IsNull(RSTTRXFILESUB!PINV), "", RSTTRXFILESUB!PINV)
                    RSTTRXFILE!BILL_DATE = IIf(IsNull(RSTTRXFILESUB!VCH_DATE), Null, RSTTRXFILESUB!VCH_DATE)
                Case "XX", "OP"
                    RSTTRXFILE!DIST_NAME = "OPENIING STOCK"
            End Select
        End If
        RSTTRXFILESUB.Close
        Set RSTTRXFILESUB = Nothing
        
        RSTTRXFILE.Update
    Next I

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    txtBillNo.Text = 1
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select MAX(Val(VCH_NO)) From WAR_TRXFILE", db2, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        txtBillNo.Text = IIf(IsNull(RSTTRXFILE.Fields(0)), 1, RSTTRXFILE.Fields(0) + 1)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
SKIP:
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    lbladdress.Caption = ""
    lbltin.Caption = ""
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    LBLSALEFLAG.Caption = ""
    OPTCUSTOMER.Value = True
    cmdRefresh.Enabled = False
    CMDEXIT.Enabled = True
    CMDPRINT.Enabled = False
    CMDEXIT.Enabled = True
    TXTSLNO.Enabled = True
    TXTDEALER.Enabled = True
    TXTDEALER.SetFocus
    TXTQTY.Tag = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    cmdview.Enabled = True
    M_DELETE = False
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub cmdview_Click()
    LblInvoice(0).Top = 1400
    TXTDEALER.Top = 225
    DataList2.Top = 570
    cmbinv.Visible = True
    TXTDEALER.Text = ""
    CMDEXIT.Caption = "CANCEL"
    TXTDEALER.Enabled = True
    DataList2.Enabled = True
    cmdRefresh.Enabled = False
    TXTDEALER.SetFocus
    cmdview.Enabled = False
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo eRRhAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(VCH_NO)) From WAR_TRXFILE", db2, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    BATCH_FLAG = True
    ACT_FLAG = True
    Call FILLCOMBO
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 5000
    grdsales.ColWidth(3) = 900
    grdsales.ColWidth(4) = 0
    grdsales.ColWidth(5) = 2500
    grdsales.ColWidth(6) = 0
    grdsales.ColWidth(7) = 0
    grdsales.ColWidth(8) = 0
    grdsales.ColWidth(9) = 0
    grdsales.ColWidth(10) = 0
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 0
    
    grdsales.TextArray(0) = "Sl No"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "Product Description"
    grdsales.TextArray(3) = "Qty"
    grdsales.TextArray(4) = "PACK"
    grdsales.TextArray(5) = "Serial No"
    grdsales.TextArray(6) = "ITEM CODE"
    grdsales.TextArray(7) = "Vch No"
    grdsales.TextArray(8) = "Line No"
    grdsales.TextArray(8) = "Trx Type"
    grdsales.TextArray(10) = "MFGR"
    grdsales.TextArray(11) = "FLAG"
    grdsales.TextArray(12) = "ISSUE QTY"
    
    grdsales.ColAlignment(0) = 4
    grdsales.ColAlignment(2) = 1
    grdsales.ColAlignment(3) = 4
    grdsales.ColAlignment(5) = 1
    
    PHYFLAG = True
    TMPFLAG = True
    BILL_FLAG = True
    ITEM_FLAG = True
    Me.Top = 0
    INV_FLAG = True
    M_DELETE = False
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    
    txtBatch.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    TXTSLNO.Text = 1
    TXTSLNO.Enabled = True
    txtBillNo.Enabled = False
    CLOSEALL = 1
    M_EDIT = False
    Me.Width = 11100
    Me.Height = 10000
    Me.Left = 0

    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If TMPFLAG = False Then TMPREC.Close
        If BILL_FLAG = False Then PHY_BILL.Close
        If ITEM_FLAG = False Then PHY_ITEM.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If INV_FLAG = False Then INV_REC.Close
        If BATCH_FLAG = False Then PHY_BATCH.Close
        If MDIMAIN.PCTMENU.Visible = True Then
            MDIMAIN.PCTMENU.Enabled = True
            MDIMAIN.PCTMENU.SetFocus
        Else
            MDIMAIN.pctmenu2.Enabled = True
            MDIMAIN.pctmenu2.SetFocus
        End If
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDPOPUPBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
              '0 VCH_NO
              '1 VCH_DATE
               '2 UNIT
               '3 QTY
               '4 ITEM_NAME
               '5 REF_NO
               '6 R_VCH_NO
               '7 R_TRX_TYPE
               '8 R_LINE_NO

            
            TXTQTY.Text = ""
            TxtPack.Text = GRDPOPUPBILL.Columns(2)
            TXTQTY.Text = GRDPOPUPBILL.Columns(3)
            TXTQTY.Tag = Val(TXTQTY.Text)
            txtBatch.Text = GRDPOPUPBILL.Columns(5)
            TXTVCHNO.Text = GRDPOPUPBILL.Columns(6)
            TXTTRXTYPE.Text = GRDPOPUPBILL.Columns(7)
            TXTLINENO.Text = GRDPOPUPBILL.Columns(8)
        
            'TXTCOMAMT.Text = GRDPOPUPBILL.Columns(15)
            Set GRDPOPUPBILL.DataSource = Nothing
            
            FRMEGRDBILL.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TxtPack.Text = ""
            
            Set GRDPOPUPBILL.DataSource = Nothing
            FRMEGRDBILL.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
        
    End Select
End Sub

Private Sub GRDPOPUPITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim I As Integer
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            For I = 1 To grdsales.Rows - 1
                If Trim(grdsales.TextMatrix(I, 12)) = Trim(TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists.... Do yo want to add this item", vbYesNo, "Goods Under Warranty") = vbNo Then
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
            Next I
            
            If OPTCUSTOMER.Value = True Then
                Call FILLBILLDB
                If B_FLAG = True Then
                    Call FILL_BILLGRID
                Else
                    FRMEITEM.Visible = False
                    FRMEMAIN.Enabled = True
                    If MsgBox("This Item has not been sold to " & DataList2.Text & " this Year... Do You Want to Continue...?", vbYesNo, "Goods Under Warranty") = vbYes Then
                        TXTPRODUCT.Enabled = False
                        TXTQTY.Enabled = True
                        TXTQTY.SetFocus
                    Else
                        TXTPRODUCT.Enabled = True
                        TXTPRODUCT.SetFocus
                    End If
                End If
            Else
                Call Check_Stock
            End If
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TxtPack.Text = ""
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
            
    End Select
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub OPTCUSTOMER_Click()
    If LBLSALEFLAG.Caption = "Y" Then
        MsgBox "CANNOT CHANGE TO SELF SINCE ALREADY SAVED...", , "Warranty Replacement.."
        OPTSELF.SetFocus
        Exit Sub
    End If
    TXTDEALER.Visible = True
    DataList2.Visible = True
    DataList2.Enabled = True
    TXTDEALER.Enabled = True
    TXTDEALER.SetFocus
End Sub

Private Sub OPTSELF_Click()
    If LBLSALEFLAG.Caption = "N" Then
        MsgBox "CANNOT CHANGE TO SELF SINCE ALREADY SAVED...", , "Warranty Replacement.."
        OPTCUSTOMER.SetFocus
        Exit Sub
    End If
    TXTDEALER.Visible = False
    DataList2.Visible = False
    
End Sub

Private Sub OPTSELF_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTINVDATE.Enabled = False
            DataList2.Enabled = False
            TXTDEALER.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            'TXTDEALER.Enabled = False
            'DataList2.Enabled = True
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
    End Select
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(txtBatch.Text) = "" Then Exit Sub
            txtBatch.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            txtBatch.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
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
Dim TRXMAST As ADODB.Recordset
Dim RSTDN As ADODB.Recordset

Dim E_Bill As String
Dim I As Integer
On Error GoTo eRRhAND
Select Case KeyCode
    Case vbKeyReturn
        If Val(txtBillNo.Text) = 0 Then Exit Sub
        grdsales.Rows = 1
        I = 0
        EDIT_BILL = False
        LBLSALEFLAG = ""
        Set RSTDN = New ADODB.Recordset
        RSTDN.Open "Select * From WAR_TRXFILE WHERE VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db2, adOpenStatic, adLockReadOnly
        Do Until RSTDN.EOF
            I = I + 1
            LBLDATE.Caption = Format(RSTDN!VCH_DATE, "DD/MM/YYYY")
            LBLTIME.Caption = Time
            grdsales.Rows = grdsales.Rows + 1
            grdsales.FixedRows = 1
            grdsales.TextMatrix(I, 0) = I
            grdsales.TextMatrix(I, 1) = RSTDN!ITEM_CODE
            grdsales.TextMatrix(I, 2) = RSTDN!ITEM_NAME
            grdsales.TextMatrix(I, 3) = RSTDN!QTY
            grdsales.TextMatrix(I, 4) = Val(RSTDN!UNIT)
            grdsales.TextMatrix(I, 5) = IIf(IsNull(RSTDN!REF_NO), "", RSTDN!REF_NO)
            grdsales.TextMatrix(I, 6) = RSTDN!ITEM_CODE
            grdsales.TextMatrix(I, 7) = RSTDN!R_VCH_NO
            grdsales.TextMatrix(I, 8) = RSTDN!R_LINE_NO
            grdsales.TextMatrix(I, 9) = RSTDN!R_TRX_TYPE
            If RSTDN!ACT_CODE <> "111111" Then
                TXTDEALER.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
                LBLSALEFLAG = "N"
                OPTCUSTOMER.Value = True
            Else
                TXTDEALER.Text = ""
                LBLSALEFLAG = "Y"
                OPTSELF.Value = True
            End If
            'DataList2.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
            TXTINVDATE.Text = IIf(IsNull(RSTDN!VCH_DATE), Date, RSTDN!VCH_DATE)
            
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMASTWO WHERE ITEMMASTWO.ITEM_CODE = '" & Trim(RSTDN!ITEM_CODE) & "'", db2, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                grdsales.TextMatrix(I, 10) = IIf(IsNull(TRXMAST!MANUFACTURER), "", Trim(TRXMAST!MANUFACTURER))
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            grdsales.TextMatrix(I, 11) = RSTDN!CHECK_FLAG
            grdsales.TextMatrix(I, 12) = RSTDN!ISSUEQTY
            If RSTDN!CHECK_FLAG = "Y" Then EDIT_BILL = True
            RSTDN.MoveNext
        Loop
        RSTDN.Close
        Set RSTDN = Nothing
        
        TXTSLNO.Text = grdsales.Rows
        txtBillNo.Enabled = False
        TXTSLNO.Enabled = True
        
        If EDIT_BILL = True Then
            CMDEXIT.Caption = "CANCEL"
            FRMEMASTER.Enabled = False
            TXTSLNO.Enabled = False
            cmdview.Enabled = False
            CMDEXIT.SetFocus
            'TXTSLNO.SetFocus
        Else
            cmdview.Enabled = True
            CMDEXIT.Caption = "E&XIT"
            TXTINVDATE.Enabled = True
            If OPTCUSTOMER.Value = True Then
                DataList2.Enabled = True
                TXTDEALER.Enabled = True
                TXTDEALER.SetFocus
            Else
                OPTSELF.SetFocus
            End If
        End If
    
End Select
    
    'DataList2.BoundText = DataList2.TextMatrix(grdSTOCKLESS.Row, 1)
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    
    Exit Sub
eRRhAND:
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
    Dim TRXDN As ADODB.Recordset
    Dim I As Integer
    Dim n As Integer
    
    I = 1
    n = 1
    Set TRXDN = New ADODB.Recordset
    TRXDN.Open "Select MAX(Val(VCH_NO)) From WAR_TRXFILE", db2, adOpenStatic, adLockReadOnly
    If Not (TRXDN.EOF And TRXDN.BOF) Then
        I = IIf(IsNull(TRXDN.Fields(0)), 1, TRXDN.Fields(0) + 1)
        If Val(txtBillNo.Text) > I Then txtBillNo.Text = I
    End If
    TRXDN.Close
    Set TRXDN = Nothing
    
    txtBillNo.Enabled = False
    'Call TXTBILLNO_KeyDown(13, 0)
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
                TXTDEALER.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                TXTINVDATE.Enabled = False
                If OPTCUSTOMER.Value = True Then
                    TXTDEALER.Enabled = True
                    DataList2.Enabled = True
                    TXTDEALER.SetFocus
                Else
                    OPTSELF.SetFocus
                End If
            End If
        Case vbKeyEscape
            TXTINVDATE.Enabled = False
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

Private Sub TxtPack_GotFocus()
    TxtPack.SelStart = 0
    TxtPack.SelLength = Len(TxtPack.Text)
End Sub

Private Sub TxtPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TxtPack.Text) = 0 Then Exit Sub
        
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TxtPack.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TxtPack.Text = ""
            TXTPRODUCT.Enabled = True
            TxtPack.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TxtPack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim I As Integer
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTP_RATE As ADODB.Recordset

'    On Error GoTo eRRhAND
    Select Case KeyCode
        Case 106
            If TXTQTY.Tag <> "" Then
                TXTPRODUCT.Text = Trim(TXTQTY.Tag)
                TXTPRODUCT.SelStart = 0
                TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
            End If
        Case vbKeyReturn
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            CmdDelete.Enabled = False
            'If NONSTOCK = True Then GoTo SKIP
            TxtPack.Text = ""
            TXTQTY.Text = ""
            
            txtBatch.Text = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
           
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For I = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(I, 12)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                Next I
                                
                If OPTCUSTOMER.Value = True Then
                    Call FILLBILLDB
                    If B_FLAG = True Then
                        Call FILL_BILLGRID
                    Else
                        FRMEITEM.Visible = False
                        FRMEMAIN.Enabled = True
                        If MsgBox("This Item has not been sold to " & DataList2.Text & " this Year... Do You Want to Continue...?", vbYesNo, "Goods Under Warranty") = vbYes Then
                            TXTPRODUCT.Enabled = False
                            TXTQTY.Enabled = True
                            TXTQTY.SetFocus
                        Else
                            TXTPRODUCT.Enabled = True
                            TXTPRODUCT.SetFocus
                        End If
                    End If
                Else
                    Call CheckStockBatch
                End If
                
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
            txtBatch.Enabled = False
            CmdDelete.Enabled = False
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            txtBatch.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
eRRhAND:
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

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim I As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If OPTCUSTOMER.Value = True And Val(TXTQTY.Text) > Val(TXTQTY.Tag) Then
            
                If (MsgBox("Sold Qty is only .. " & Val(TXTQTY.Tag) & "...Do you want to Continue", vbYesNo, "SALES RETURN") = vbNo) Then
                    TXTQTY.SelStart = 0
                    TXTQTY.SelLength = Len(TXTQTY.Text)
                    Exit Sub
                End If
            Else
                I = 0
                 Set RSTTRXFILE = New ADODB.Recordset
                 RSTTRXFILE.Open "SELECT BAL_QTY  FROM RTRXFILEWO WHERE RTRXFILEWO.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILEWO.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILEWO.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILEWO.LINE_NO = " & Val(TXTLINENO.Text) & "", db2, adOpenStatic, adLockReadOnly
                 If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                     If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                     I = RSTTRXFILE!BAL_QTY
                 End If
                 RSTTRXFILE.Close
                 Set RSTTRXFILE = Nothing
                 
                Set RSTTRXFILE = Nothing
                 'If Val(TXTQTY.Text) = 0 Then Exit Sub
                 If I > 0 Then
                     If Val(TXTQTY.Text) > I Then
                         If (MsgBox("AVAILABLE STOCK IS  " & I & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                             'MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                             TXTQTY.SelStart = 0
                             TXTQTY.SelLength = Len(TXTQTY.Text)
                             Exit Sub
                         End If
                     End If
                 End If
            End If
            
            TXTQTY.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TxtPack.Text = ""
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
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
    cmdview.Enabled = False
    DataList2.Enabled = False
    TXTDEALER.Enabled = False
    TXTINVDATE.Enabled = False
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.Text) = 0 Then
                TXTSLNO.Text = ""
                TXTPRODUCT.Text = ""
                TXTQTY.Text = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TxtPack.Text = ""
                
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
                TXTSLNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 0)
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 3)
                TxtPack.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                'TXTCOMAMT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 19)
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                txtBatch.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            txtBatch.Enabled = False
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TxtPack.Text = ""
                TXTQTY.Text = ""
                txtBatch.Text = ""
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.Rows > 1 Then
                CMDPRINT.Enabled = True
                cmdRefresh.Enabled = True
                CMDPRINT.SetFocus
            Else
                FRMEMASTER.Enabled = True
                TXTINVDATE.Enabled = True
                If OPTCUSTOMER.Value = True Then
                    TXTDEALER.Enabled = True
                    DataList2.Enabled = True
                    TXTDEALER.SetFocus
                Else
                    OPTSELF.SetFocus
                End If
            End If
            If M_DELETE = True Then
                cmdRefresh.Enabled = True
                cmdRefresh.SetFocus
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

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Format(Month(DateIn), "00") & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Function FILL_ITEMGRID()
    FRMEMAIN.Enabled = False
    FRMEITEM.Visible = True
    'Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    'FRMEGRDTMP.Visible = False
    
    
    If ITEM_FLAG = True Then
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY] From ITEMMASTWO  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY] From ITEMMASTWO  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    End If

    Set GRDPOPUPITEM.DataSource = PHY_ITEM
    'GRDPOPUPITEM.RowHeight = 250
    GRDPOPUPITEM.Columns(0).Visible = False
    GRDPOPUPITEM.Columns(1).Caption = "ITEM NAME"
    GRDPOPUPITEM.Columns(1).Width = 3800
    GRDPOPUPITEM.Columns(2).Caption = "QTY"
    GRDPOPUPITEM.Columns(2).Width = 1300
    GRDPOPUPITEM.SetFocus
End Function

Function FILL_BILLGRID()
                    
    FRMEMAIN.Enabled = False
    FRMEITEM.Visible = False
    FRMEGRDBILL.Visible = True
    Set GRDPOPUPBILL.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    
    If BILL_FLAG = True Then
        PHY_BILL.Open "Select VCH_NO, VCH_DATE, UNIT, QTY, ITEM_NAME, REF_NO, R_VCH_NO, R_TRX_TYPE, R_LINE_NO From DE_RET_DETAILS ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
        BILL_FLAG = False
    Else
        PHY_BILL.Close
        PHY_BILL.Open "Select VCH_NO, VCH_DATE, UNIT, QTY, ITEM_NAME, REF_NO, R_VCH_NO, R_TRX_TYPE, R_LINE_NO From DE_RET_DETAILS ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
        BILL_FLAG = False
    End If
    
    Set GRDPOPUPBILL.DataSource = PHY_BILL
    
    GRDPOPUPBILL.Columns(0).Caption = "BILL NO."
    GRDPOPUPBILL.Columns(1).Caption = "BILL DATE"
    GRDPOPUPBILL.Columns(2).Caption = "UNIT"
    GRDPOPUPBILL.Columns(3).Caption = "QTY"
    GRDPOPUPBILL.Columns(4).Caption = "ITEM"
    GRDPOPUPBILL.Columns(5).Caption = "Serial No"
    '10- R_VCH NO
    '11- R_TYPE
    '12 - R_LINE NO
    
    GRDPOPUPBILL.Columns(0).Width = 900
    GRDPOPUPBILL.Columns(1).Width = 1150
    GRDPOPUPBILL.Columns(2).Width = 0
    GRDPOPUPBILL.Columns(3).Width = 900
    GRDPOPUPBILL.Columns(4).Width = 3000
    GRDPOPUPBILL.Columns(5).Width = 2500
    GRDPOPUPBILL.Columns(6).Width = 0
    GRDPOPUPBILL.Columns(7).Width = 0
    GRDPOPUPBILL.Columns(8).Width = 0
    'GRDPOPUPBILL.Columns(9).Width = 1200
    'GRDPOPUPBILL.Columns(10).Width = 0
    'GRDPOPUPBILL.Columns(11).Width = 0
    
    GRDPOPUPBILL.SetFocus
    LBLHEAD(0).Caption = GRDPOPUPBILL.Columns(4).Text
    LBLHEAD(2).Visible = True
    LBLHEAD(0).Visible = True
    
End Function

Private Sub FILLCOMBO()
    On Error GoTo eRRhAND
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If ACT_FLAG = True Then
        ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    Else
        ACT_REC.Close
        ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    End If
    
    Set Me.CMBDISTI.RowSource = ACT_REC
    CMBDISTI.ListField = "ACT_NAME"
    CMBDISTI.BoundColumn = "ACT_CODE"
    Screen.MousePointer = vbNormal
    Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub


Private Sub TXTDEALER_Change()
    On Error GoTo eRRhAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
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
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
    OPTCUSTOMER.Value = True
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            lbladdress.Caption = ""
            lbltin.Caption = ""
            DataList2.Enabled = True
            DataList2.SetFocus
        Case vbKeyEscape
            'TXTDEALER.Enabled = False
            'DataList2.Enabled = True
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
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

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
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
            'FRMEMASTER.Enabled = False
            DataList2.Enabled = False
            TXTDEALER.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTDEALER.Enabled = True
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

Private Sub DataList2_Click()
    Dim rstCustomer As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    On Error GoTo eRRhAND

    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select ADDRESS, DL_NO, KGST from [ACTMAST]  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db2, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = IIf(IsNull(rstCustomer!ADDRESS), "", Trim(rstCustomer!ADDRESS))
        'lbldlno.Caption = Trim(rstCustomer!DL_NO)
        lbltin.Caption = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
    Else
        lbladdress.Caption = ""
        lbltin.Caption = ""
    End If
    Call FILLINVOICE
    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text
    cmbinv.Text = ""
    If TXTDEALER.Top = 225 Then
        grdsales.FixedRows = 0
        grdsales.Rows = 1
    End If
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
End Sub


Private Function FILLINVOICE()
    On Error GoTo eRRhAND
    
    Screen.MousePointer = vbHourglass
    Set cmbinv.DataSource = Nothing
    If INV_FLAG = True Then
        INV_REC.Open "Select DISTINCT VCH_NO From WAR_TRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_NO", db2, adOpenStatic, adLockReadOnly
        INV_FLAG = False
    Else
        INV_REC.Close
        INV_REC.Open "Select DISTINCT VCH_NO From WAR_TRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_NO", db2, adOpenStatic, adLockReadOnly
        INV_FLAG = False
    End If
    
    Set Me.cmbinv.RowSource = INV_REC
    cmbinv.ListField = "VCH_NO"
    cmbinv.BoundColumn = "VCH_NO"
    Screen.MousePointer = vbNormal
    Exit Function

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
End Sub

Private Function VIEWGRID()
    Exit Function
    Dim TRXMAST As ADODB.Recordset
    Dim RSTDN As ADODB.Recordset
    
    Dim E_Bill As String
    Dim I As Integer
    On Error GoTo eRRhAND
        If Val(txtBillNo.Text) = 0 Then Exit Function
        grdsales.Rows = 1
        I = 0
        Set RSTDN = New ADODB.Recordset
        RSTDN.Open "Select * From WAR_TRXFILE WHERE VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db2, adOpenStatic, adLockReadOnly
        Do Until RSTDN.EOF
            I = I + 1
            LBLDATE.Caption = Format(RSTDN!VCH_DATE, "DD/MM/YYYY")
            LBLTIME.Caption = Time
            grdsales.Rows = grdsales.Rows + 1
            grdsales.FixedRows = 1
            grdsales.TextMatrix(I, 0) = I
            grdsales.TextMatrix(I, 1) = RSTDN!ITEM_CODE
            grdsales.TextMatrix(I, 2) = RSTDN!ITEM_NAME
            grdsales.TextMatrix(I, 3) = RSTDN!QTY
            grdsales.TextMatrix(I, 4) = Val(RSTDN!UNIT)
            
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMASTWO WHERE ITEMMASTWO.ITEM_CODE = '" & Trim(RSTDN!ITEM_CODE) & "'", db2, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                grdsales.TextMatrix(I, 16) = Trim(TRXMAST!MANUFACTURER)
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            grdsales.TextMatrix(I, 5) = Format(RSTDN!MRP, ".000")
            grdsales.TextMatrix(I, 6) = Format(RSTDN!SALES_PRICE, ".000")
            grdsales.TextMatrix(I, 7) = 0 'DISC
            grdsales.TextMatrix(I, 8) = Val(RSTDN!SALES_TAX)
            grdsales.TextMatrix(I, 9) = RSTDN!REF_NO
            grdsales.TextMatrix(I, 10) = ""
            grdsales.TextMatrix(I, 11) = Format(Val(RSTDN!TRX_TOTAL), ".000")
            
            grdsales.TextMatrix(I, 12) = RSTDN!ITEM_CODE
            grdsales.TextMatrix(I, 13) = RSTDN!R_VCH_NO
            grdsales.TextMatrix(I, 14) = RSTDN!R_LINE_NO
            grdsales.TextMatrix(I, 15) = RSTDN!R_TRX_TYPE
            TXTDEALER.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
            'DataList2.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
            TXTINVDATE.Text = IIf(IsNull(RSTDN!VCH_DATE), Date, RSTDN!VCH_DATE)
            
            RSTDN.MoveNext
        Loop
        RSTDN.Close
        Set RSTDN = Nothing
        
        TXTSLNO.Text = grdsales.Rows
        Exit Function
eRRhAND:
    MsgBox Err.Description

End Function

Private Function FILLBILLDB()
    Dim TRXFILE As ADODB.Recordset
    Dim TRXFILESUB As ADODB.Recordset
    Dim TRXBILL As ADODB.Recordset
    
    Dim n As Integer
    Dim M As Integer
    
    B_FLAG = False
    db2.Execute "delete * From DE_RET_DETAILS"
    Set TRXFILE = New ADODB.Recordset
    TRXFILE.Open "Select * From TRXFILEWO WHERE (TRX_TYPE='RI' OR TRX_TYPE='SI') AND CST <>2 AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "'", db2, adOpenStatic, adLockReadOnly
    Do Until TRXFILE.EOF
        Set TRXFILESUB = New ADODB.Recordset
        TRXFILESUB.Open "Select * From TRXSUB WHERE VCH_NO = " & TRXFILE!VCH_NO & " AND LINE_NO = " & TRXFILE!LINE_NO & "", db2, adOpenStatic, adLockReadOnly

        If Not (TRXFILESUB.EOF And TRXFILESUB.BOF) Then
            Set TRXBILL = New ADODB.Recordset
            TRXBILL.Open "SELECT *  FROM DE_RET_DETAILS", db2, adOpenStatic, adLockOptimistic, adCmdText
            B_FLAG = True
            TRXBILL.AddNew
            TRXBILL!VCH_NO = TRXFILESUB!VCH_NO
            TRXBILL!TRX_TYPE = TRXFILESUB!TRX_TYPE
            TRXBILL!LINE_NO = TRXFILESUB!LINE_NO
            TRXBILL!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
            TRXBILL!UNIT = TRXFILE!UNIT
            TRXBILL!QTY = TRXFILE!QTY
            TRXBILL!ITEM_NAME = TRXFILE!ITEM_NAME
            TRXBILL!REF_NO = TRXFILE!REF_NO
            TRXBILL!R_VCH_NO = TRXFILESUB!R_VCH_NO
            TRXBILL!R_TRX_TYPE = TRXFILESUB!R_TRX_TYPE
            TRXBILL!R_LINE_NO = TRXFILESUB!R_LINE_NO
            'TRXBILL!COM_AMT = TRXFILE!COM_AMT
            
            TRXBILL.Update
            TRXBILL.Close
            Set TRXBILL = Nothing
        End If
        TRXFILESUB.Close
        Set TRXFILESUB = Nothing
        TRXFILE.MoveNext
    Loop
    TRXFILE.Close
    Set TRXFILE = Nothing

    Set GRDPOPUPITEM.DataSource = Nothing
End Function

Private Function CheckStockBatch()
    Dim I As Integer
    Dim RSTBALQTY As ADODB.Recordset
    Dim MINUSFLAG, NONSTOCKFLAG  As Boolean
    
    M_STOCK = 0
    If Trim(TXTPRODUCT.Text) = "" Then Exit Function
            CmdDelete.Enabled = False
            TXTQTY.Text = ""
            txtBatch.Text = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
           
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For I = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(I, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                            Exit Function
                        Else
                            Exit For
                        End If
                    End If
                Next I
                Set grdtmp.DataSource = Nothing
                If TMPFLAG = True Then
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, LINE_DISC, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN From RTRXFILEWO  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                Else
                    TMPREC.Close
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, LINE_DISC, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN From RTRXFILEWO  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                End If
                
                Set grdtmp.DataSource = TMPREC
                If TMPREC.RecordCount = 1 Then
                    'TXTQTY.Text = grdtmp.Columns(2)
                    txtBatch.Text = grdtmp.Columns(7)
                    
                    TXTVCHNO.Text = grdtmp.Columns(9)
                    TXTLINENO.Text = grdtmp.Columns(10)
                    TXTTRXTYPE.Text = grdtmp.Columns(11)
                    TXTUNIT.Text = grdtmp.Columns(6)
                                        
                    TXTPRODUCT.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                    Exit Function
                ElseIf TMPREC.RecordCount = 0 Then
                    Set RSTBALQTY = New ADODB.Recordset
                    RSTBALQTY.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
                    With RSTBALQTY
                        If Not (.EOF And .BOF) Then
                            M_STOCK = !CLOSE_QTY
                        End If
                    End With
                    RSTBALQTY.Close
                    Set RSTBALQTY = Nothing
            
                    TXTQTY.Text = 0
                    I = 0
                    If (MsgBox("AVAILABLE STOCK IS  " & I & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                        TXTPRODUCT.Enabled = True
                        TXTQTY.Enabled = False
                        TXTPRODUCT.SelStart = 0
                        TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
                        TXTPRODUCT.SetFocus
                        Exit Function
                    Else
                        MINUSFLAG = True
                    End If
                    NONSTOCKFLAG = True
                ElseIf TMPREC.RecordCount > 1 Then
                    Call FILL_BATCHGRID
                    Exit Function
                End If
                Exit Function
            ElseIf PHY.RecordCount > 1 Then
                'FRMSUB.grdsub.Columns(0).Visible = True
                'FRMSUB.grdsub.Columns(1).Caption = "ITEM NAME"
                'FRMSUB.grdsub.Columns(1).Width = 3200
                'FRMSUB.grdsub.Columns(2).Caption = "QTY"
                'FRMSUB.grdsub.Columns(2).Width = 1300
                Call FILL_ITEMGRID
            End If
End Function

Private Function Check_Stock()
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim NONSTOCKFLAG As Boolean
    Dim MINUSFLAG As Boolean
    Dim I As Integer

        NONSTOCKFLAG = False
        MINUSFLAG = False
        M_STOCK = Val(GRDPOPUPITEM.Columns(2))
        TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
        TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
        I = 0
        If M_STOCK <= 0 Then
            If (MsgBox("AVAILABLE STOCK IS  " & M_STOCK & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                Exit Function
            Else
                MINUSFLAG = True
            End If
            NONSTOCKFLAG = True
        End If
        For I = 1 To grdsales.Rows - 1
            If Trim(grdsales.TextMatrix(I, 1)) = Trim(TXTITEMCODE.Text) Then
                If MsgBox("This Item Already exists.... Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then
                    Set GRDPOPUPITEM.DataSource = Nothing
                    FRMEITEM.Visible = False
                    FRMEMAIN.Enabled = True
                    TXTPRODUCT.Enabled = True
                    TXTQTY.Enabled = False
                    TXTPRODUCT.SetFocus
                    Exit Function
                Else
                    Exit For
                End If
            End If
        Next I
        Set GRDPOPUPITEM.DataSource = Nothing
        If ITEM_FLAG = True Then
            If NONSTOCKFLAG = True Then
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, REF_NO, LINE_DISC, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG  From RTRXFILEWO  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
            Else
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, REF_NO, LINE_DISC, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG  From RTRXFILEWO  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
            End If
            ITEM_FLAG = False
        Else
            PHY_ITEM.Close
            If NONSTOCKFLAG = True Then
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, REF_NO, LINE_DISC, VCH_NO, LINE_NO, TRX_TYPE CHECK_FLAG From RTRXFILEWO  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
            Else
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, REF_NO, LINE_DISC, VCH_NO, LINE_NO, TRX_TYPE, CHECK_FLAG From RTRXFILEWO  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
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
            Exit Function
        End If
        If PHY_ITEM.RecordCount = 1 Or MINUSFLAG = True Then
            'TXTQTY.Text = GRDPOPUPITEM.Columns(2)
            'TXTTAX.Text = 0 'GRDPOPUPITEM.Columns(4)
            txtBatch.Text = GRDPOPUPITEM.Columns(6)
            
            TXTUNIT.Text = GRDPOPUPITEM.Columns(4)
            TXTVCHNO.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(5), "")
            TXTLINENO.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(6), "")
            TXTTRXTYPE.Text = IIf((NONSTOCKFLAG = False), GRDPOPUPITEM.Columns(7), "")
                        
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
            Exit Function
        ElseIf PHY_ITEM.RecordCount > 1 And MINUSFLAG = False Then
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Call FILL_BATCHGRID
        End If
    Exit Function
eRRhAND:
    MsgBox Err.Description
End Function

Function FILL_BATCHGRID()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, LINE_DISC, CHECK_FLAG From RTRXFILEWO  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, LINE_DISC, CHECK_FLAG From RTRXFILEWO  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "Serial No."
    GRDPOPUP.Columns(1).Caption = "Qty"
    GRDPOPUP.Columns(2).Caption = ""
    GRDPOPUP.Columns(4).Caption = "VCH No"
    GRDPOPUP.Columns(5).Caption = "Line No"
    GRDPOPUP.Columns(6).Caption = "Trx Type"
    
    GRDPOPUP.Columns(0).Width = 2500
    GRDPOPUP.Columns(1).Width = 900
    
    GRDPOPUP.Columns(2).Visible = False
    GRDPOPUP.Columns(3).Visible = False
    GRDPOPUP.Columns(4).Visible = False
    GRDPOPUP.Columns(5).Visible = False
    GRDPOPUP.Columns(6).Visible = False
    GRDPOPUP.Columns(7).Visible = False
    GRDPOPUP.Columns(8).Visible = False
    
    FRMEGRDTMP.Enabled = True
    GRDPOPUP.Enabled = True
    GRDPOPUP.SetFocus
    LBLHEAD(2).Caption = GRDPOPUP.Columns(3).Text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
End Function

Private Sub GRDPOPUP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTtax As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            'TXTQTY.Text = GRDPOPUP.Columns(1)
            txtBatch.Text = GRDPOPUP.Columns(0)
            TXTVCHNO.Text = GRDPOPUP.Columns(4)
            TXTLINENO.Text = GRDPOPUP.Columns(5)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(6)
            TXTUNIT.Text = GRDPOPUP.Columns(7)
            TxtPack.Text = GRDPOPUP.Columns(7)
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

