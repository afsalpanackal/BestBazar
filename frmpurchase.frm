VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmpurchase 
   Caption         =   "PURCHASE ENTRY"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   Icon            =   "frmpurchase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid grdPURCHASE 
      Height          =   1800
      Left            =   5385
      TabIndex        =   37
      Top             =   15
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   3175
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&REFRESH"
      Height          =   435
      Left            =   3750
      TabIndex        =   15
      Top             =   7170
      Width           =   1200
   End
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "E&XIT"
      Height          =   420
      Left            =   5040
      TabIndex        =   16
      Top             =   7170
      Width           =   1170
   End
   Begin VB.Frame FRMEMASTER 
      BackColor       =   &H00FFC0C0&
      Height          =   1845
      Left            =   -15
      TabIndex        =   21
      Top             =   -75
      Width           =   5370
      Begin VB.TextBox TXTINVDATE 
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
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1500
         Width           =   1260
      End
      Begin VB.TextBox TXTINVOICE 
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
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1485
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo CMBDISTI 
         Height          =   1230
         Left            =   1320
         TabIndex        =   1
         Top             =   135
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   2170
         _Version        =   393216
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
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
         Index           =   5
         Left            =   90
         TabIndex        =   24
         Top             =   435
         Width           =   1005
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
         Left            =   2550
         TabIndex        =   23
         Top             =   1485
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   60
         TabIndex        =   22
         Top             =   1485
         Width           =   1215
      End
   End
   Begin VB.Frame FRMEPURCHASE 
      BackColor       =   &H00FFC0C0&
      Height          =   8715
      Left            =   -15
      TabIndex        =   0
      Top             =   1740
      Width           =   9615
      Begin VB.Frame FRMEGRDTMP 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   2865
         Left            =   1920
         TabIndex        =   39
         Top             =   795
         Visible         =   0   'False
         Width           =   5910
         Begin MSDataGridLib.DataGrid grdtmp 
            Height          =   2490
            Left            =   135
            TabIndex        =   40
            Top             =   210
            Width           =   5670
            _ExtentX        =   10001
            _ExtentY        =   4392
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
      Begin MSDataGridLib.DataGrid GRDPURLIST 
         Height          =   3960
         Left            =   45
         TabIndex        =   36
         Top             =   165
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   6985
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame FRMFIELDS 
         BackColor       =   &H00FFC0C0&
         Height          =   1170
         Left            =   75
         TabIndex        =   25
         Top             =   4110
         Width           =   9480
         Begin VB.TextBox TXTITEM 
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
            Left            =   870
            MaxLength       =   10
            TabIndex        =   38
            Top             =   285
            Width           =   3930
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
            Left            =   5265
            MaxLength       =   3
            TabIndex        =   4
            Top             =   240
            Width           =   555
         End
         Begin VB.TextBox TXTEXPDATE 
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
            Left            =   1815
            MaxLength       =   10
            TabIndex        =   8
            Top             =   765
            Width           =   1350
         End
         Begin VB.TextBox TXTMRP 
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
            Left            =   3705
            MaxLength       =   8
            TabIndex        =   9
            Top             =   780
            Width           =   675
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
            Height          =   315
            Left            =   6390
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   510
         End
         Begin VB.TextBox TXTBATCH 
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
            Left            =   7725
            MaxLength       =   15
            TabIndex        =   6
            Top             =   240
            Width           =   1020
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
            Height          =   315
            Left            =   4860
            MaxLength       =   8
            TabIndex        =   10
            Top             =   765
            Width           =   690
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
            Height          =   315
            Left            =   6090
            MaxLength       =   3
            TabIndex        =   11
            Top             =   750
            Width           =   540
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   330
            Left            =   840
            TabIndex        =   7
            Top             =   765
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
            Caption         =   "ITEM"
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
            Left            =   135
            TabIndex        =   34
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "EXPIRY"
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
            Height          =   270
            Index           =   0
            Left            =   90
            TabIndex        =   33
            Top             =   825
            Width           =   750
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "TAX"
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
            Height          =   270
            Index           =   6
            Left            =   5610
            TabIndex        =   32
            Top             =   780
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   7
            Left            =   3255
            TabIndex        =   31
            Top             =   765
            Width           =   435
         End
         Begin VB.Label LBLITEM 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   870
            TabIndex        =   30
            Top             =   285
            Width           =   3930
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "QTY"
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
            Left            =   4845
            TabIndex        =   29
            Top             =   270
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   9
            Left            =   5880
            TabIndex        =   28
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "BATCH"
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
            Index           =   10
            Left            =   7035
            TabIndex        =   27
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   11
            Left            =   4440
            TabIndex        =   26
            Top             =   765
            Width           =   390
         End
      End
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00FFC0C0&
         Height          =   765
         Left            =   45
         TabIndex        =   35
         Top             =   5205
         Width           =   6285
         Begin VB.CommandButton CMDCANCEL 
            Caption         =   "&CANCEL"
            Height          =   420
            Left            =   1275
            TabIndex        =   13
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton CMDDELETE 
            Caption         =   "&DELETE"
            Height          =   420
            Left            =   2460
            TabIndex        =   14
            Top             =   240
            Width           =   1170
         End
         Begin VB.CommandButton CMDSAVE 
            Caption         =   "&SAVE"
            Height          =   435
            Left            =   60
            TabIndex        =   12
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Label LBLVCH 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   8595
         TabIndex        =   20
         Top             =   5415
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "VCH NO"
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
         Left            =   7635
         TabIndex        =   19
         Top             =   5475
         Width           =   960
      End
      Begin VB.Label LBLLINE 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   8610
         TabIndex        =   18
         Top             =   5880
         Width           =   870
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LINE NO"
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
         Left            =   7665
         TabIndex        =   17
         Top             =   5940
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmpurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_FLAG As Boolean
Dim PHY_FLAG As Boolean
Dim ITEM_FLAG As Boolean


Dim ACT_REC As New ADODB.Recordset
Dim PHY_REC As New ADODB.Recordset
Dim ITEM_REC As New ADODB.Recordset

Dim CLOSEALL As Integer

Private Sub CMBDISTI_Click(Area As Integer)
    Dim rstTMP As ADODB.Recordset
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "Select Max(Val(VCH_NO)) From RTRXFILE Where TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (rstTMP.EOF And rstTMP.BOF) Then
        LBLVCH.Caption = rstTMP.Fields(0) + 1
    End If
    rstTMP.Close
    Set rstTMP = Nothing
    LBLLINE.Caption = 1
    TXTINVDATE.Text = Date
    Call FILLGRID

End Sub

Private Sub CMBDISTI_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMBDISTI.Text = "" Then Exit Sub
            If IsNull(CMBDISTI.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "PURCHASE"
                CMBDISTI.SetFocus
                Exit Sub
            End If
            TXTINVOICE.SetFocus
                        
    End Select
End Sub

Private Sub CMDCANCEL_Click()
    TXTUNIT = ""
    LBLITEM.Caption = ""
    TXTQTY = ""
    TXTBATCH = ""
    TXTEXPDATE = ""
    TXTMRP = ""
    TXTPTR = ""
    TXTTAX = ""
    TXTFREE = ""
    TXTEXPIRY.Text = "  /  "
    
    If grdPURCHASE.ApproxCount < 1 Then
        FRMEMASTER.Enabled = True
        FRMEPURCHASE.Enabled = False
        Call FILLCOMBO
    Else
        FRMEMASTER.Enabled = False
        FRMEPURCHASE.Enabled = True
        FRMECONTROLS.Enabled = False
        grdPURCHASE.Enabled = True
        grdPURCHASE.SetFocus
       
    End If
    
    

End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    TXTUNIT = ""
    LBLITEM.Caption = ""
    TXTQTY = ""
    TXTBATCH = ""
    TXTEXPDATE = ""
    TXTMRP = ""
    TXTPTR = ""
    TXTTAX = ""
    TXTFREE = ""
    TXTEXPIRY.Text = "  /  "
    TXTINVDATE.Text = ""
    TXTINVOICE.Text = ""
    
    FRMEMASTER.Enabled = True
    FRMEPURCHASE.Enabled = False
       
End Sub

Private Sub CMDSAVE_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTPRODLINK As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    
    Dim M_DATA As Integer
    Dim E_DATA As Integer
    Dim n As Integer
    
    On Error GoTo ErrHand
    
    If Val(TXTQTY.Text) = 0 Then
        MsgBox "ENTER THE QTY", vbOKOnly, "PURCHASE"
        TXTQTY.SetFocus
        Exit Sub
    End If
    
    If Val(TXTUNIT.Text) = 0 Then
        MsgBox "ENTER THE UNIT", vbOKOnly, "PURCHASE"
        TXTUNIT.SetFocus
        Exit Sub
    End If
    
    If TXTBATCH.Text = "" Then
        MsgBox "ENTER THE BATCH", vbOKOnly, "PURCHASE"
        TXTBATCH.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(TXTEXPDATE.Text) Then
        MsgBox "ENTER A VALID DATE FOR EXPIRY", vbOKOnly, "PURCHASE"
        TXTEXPDATE.SetFocus
        Exit Sub
    End If
    
    If Val(TXTMRP.Text) = 0 Then
        MsgBox "ENTER THE MRP", vbOKOnly, "PURCHASE"
        TXTMRP.SetFocus
        Exit Sub
    End If
    
    If Val(TXTPTR.Text) = 0 Then
        MsgBox "ENTER THE PTR", vbOKOnly, "PURCHASE"
        TXTPTR.SetFocus
        Exit Sub
    End If
    Set RSTPRODLINK = New ADODB.Recordset
    RSTPRODLINK.Open "SELECT * from [PRODLINK] WHERE PRODLINK.ITEM_CODE = '" & grdPURCHASE.Columns(4) & "' AND ACT_CODE='" & CMBDISTI.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTPRODLINK.EOF And RSTPRODLINK.BOF) Then
        RSTPRODLINK!ITEM_COST = Val(TXTPTR.Text)
        RSTPRODLINK!MRP = Val(TXTMRP.Text)
        RSTPRODLINK!PTR = Val(TXTPTR.Text)
        RSTPRODLINK!SALES_PRICE = Val(TXTMRP.Text)
        RSTPRODLINK!SALES_TAX = Val(TXTTAX.Text)
        RSTPRODLINK!UNIT = Val(TXTUNIT.Text)
        RSTPRODLINK!REMARKS = Val(TXTUNIT.Text)
        RSTPRODLINK!ORD_QTY = 0
        RSTPRODLINK!CST = 0
        RSTPRODLINK!CHECK_FLAG = "Y"
        RSTPRODLINK.Update
    End If
    RSTPRODLINK.Close
    Set RSTPRODLINK = Nothing
    
    M_DATA = 0
    E_DATA = 0
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * from [ITEMMAST] WHERE ITEMMAST.ITEM_CODE = '" & grdPURCHASE.Columns(4) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        M_DATA = RSTITEMMAST!RCPT_QTY
        RSTITEMMAST!RCPT_QTY = M_DATA + (Val(TXTQTY.Text) * Val(TXTUNIT.Text))
        E_DATA = (RSTITEMMAST!OPEN_QTY + RSTITEMMAST!RCPT_QTY) - RSTITEMMAST!ISSUE_QTY
        RSTITEMMAST!CLOSE_QTY = E_DATA
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    n = Val(LBLVCH.Caption)
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * from [RTRXFILE]", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTRTRXFILE.AddNew
    RSTRTRXFILE!TRX_TYPE = "PI"
    RSTRTRXFILE!VCH_NO = Val(LBLVCH.Caption)
    RSTRTRXFILE!VCH_DATE = Trim(TXTEXPDATE)
    RSTRTRXFILE!LINE_NO = Val(LBLLINE.Caption)
    RSTRTRXFILE!CATEGORY = "MEDICINE"
    RSTRTRXFILE!ITEM_CODE = Trim(grdPURCHASE.Columns(4))
    RSTRTRXFILE!ITEM_NAME = Trim(grdPURCHASE.Columns(1))
    RSTRTRXFILE!QTY = Val(grdPURCHASE.Columns(7))
    RSTRTRXFILE!ITEM_COST = 0
    RSTRTRXFILE!MRP = Val(TXTMRP.Text)
    RSTRTRXFILE!PTR = Val(TXTPTR.Text)
    RSTRTRXFILE!SALES_PRICE = Val(TXTMRP.Text)
    RSTRTRXFILE!SALES_TAX = Val(TXTTAX.Text)
    RSTRTRXFILE!UNIT = Val(TXTUNIT.Text)
    RSTRTRXFILE!VCH_DESC = "Received From " & Trim(CMBDISTI.Text)
    RSTRTRXFILE!REF_NO = Trim(TXTBATCH.Text)
    RSTRTRXFILE!ISSUE_QTY = 0
    RSTRTRXFILE!CST = 0
    RSTRTRXFILE!BAL_QTY = Val(TXTQTY.Text) * Val(TXTUNIT.Text)
    RSTRTRXFILE!TRX_TOTAL = RSTRTRXFILE!MRP * Val(TXTQTY.Text)
    RSTRTRXFILE!LINE_DISC = 0
    RSTRTRXFILE!SCHEME = 0
    RSTRTRXFILE!EXP_DATE = Trim(TXTEXPDATE.Text)
    RSTRTRXFILE!FREE_QTY = 0
    RSTRTRXFILE!CREATE_DATE = Date
    RSTRTRXFILE!C_USER_ID = "SM"
    RSTRTRXFILE!M_USER_ID = Trim(grdPURCHASE.Columns(5))
    RSTRTRXFILE!CHECK_FLAG = ""
    RSTRTRXFILE!PINV = Trim(TXTINVOICE.Text)
    RSTRTRXFILE.Update

    
    
    
    RSTRTRXFILE.Close
    
    Set RSTRTRXFILE = Nothing
    
    db2.Execute ("Delete from [TmpOrderlist] where ORCODE = '" & Trim(grdPURCHASE.Columns(0)) & "'")
    TXTUNIT = ""
    LBLITEM.Caption = ""
    TXTQTY = ""
    TXTBATCH = ""
    TXTEXPDATE = ""
    TXTMRP = ""
    TXTPTR = ""
    TXTTAX = ""
    TXTFREE = ""
    TXTEXPIRY.Text = "  /  "
    Call FILLGRID
    MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "PURCHASE"
    LBLLINE.Caption = Val(LBLLINE.Caption) + 1
    Screen.MousePointer = vbNormal
    
    If grdPURCHASE.ApproxCount < 1 Then
        FRMEMASTER.Enabled = True
        FRMEPURCHASE.Enabled = False
        Call FILLCOMBO
    Else
        FRMEMASTER.Enabled = False
        FRMEPURCHASE.Enabled = True
        FRMECONTROLS.Enabled = False
        grdPURCHASE.Enabled = True
        grdPURCHASE.SetFocus
       
    End If
    
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub Form_Load()
    ACT_FLAG = True
    PHY_FLAG = True
    ITEM_FLAG = True
    CLOSEALL = 1
    Call FILLCOMBO
    FRMEPURCHASE.Enabled = False
    Me.Width = 11400
    Me.Height = 8300
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub FILLGRID()
        
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    
    Set grdPURCHASE.DataSource = Nothing
    If PHY_FLAG = True Then
        PHY_REC.Open "select * from [TmpOrderlist]  WHERE Dist_Code = '" & CMBDISTI.BoundText & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
        PHY_FLAG = False
    Else
        PHY_REC.Close
        PHY_REC.Open "select * from [TmpOrderlist]  WHERE Dist_Code = '" & CMBDISTI.BoundText & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
        PHY_FLAG = False
    End If
    
    
    Set grdPURCHASE.DataSource = PHY_REC
    
    grdPURCHASE.Columns(0).Visible = False
    grdPURCHASE.Columns(1).Caption = "ITEM NAME"
    grdPURCHASE.Columns(1).Width = 2400
    grdPURCHASE.Columns(2).Visible = False
    grdPURCHASE.Columns(3).Visible = False
    grdPURCHASE.Columns(4).Visible = False
    grdPURCHASE.Columns(5).Visible = False
    grdPURCHASE.Columns(6).Visible = False
    grdPURCHASE.Columns(7).Caption = "QTY"
    grdPURCHASE.Columns(7).Width = 800
    
    
    
    grdPURCHASE.RowHeight = 250
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub FILLCOMBO()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If ACT_FLAG = True Then
        ACT_REC.Open "select Distinct Or_Distrib, Dist_Code from [TmpOrderlist] ORDER BY Or_Distrib", db2, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    Else
        ACT_REC.Close
        ACT_REC.Open "select Distinct Or_Distrib, Dist_Code from [TmpOrderlist] ORDER BY Or_Distrib", db2, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    End If
    
    Set Me.CMBDISTI.RowSource = ACT_REC
    CMBDISTI.ListField = "Or_Distrib"
    CMBDISTI.BoundColumn = "Dist_Code"
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If ACT_FLAG = False Then ACT_REC.Close
        If PHY_FLAG = False Then PHY_REC.Close
        If ITEM_FLAG = False Then ITEM_REC.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.Height = 15555
        'FrmCrimedata.Enabled = True
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GrdPURCHASE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTUNIT As ADODB.Recordset
    On Error GoTo ErrHand
        
    Select Case KeyCode
        Case vbKeyReturn
            Set RSTUNIT = New ADODB.Recordset
            RSTUNIT.Open "SELECT UNIT FROM PRODLINK WHERE ITEM_CODE='" & grdPURCHASE.Columns(4) & "' AND ACT_CODE='" & CMBDISTI.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not (RSTUNIT.EOF And RSTUNIT.BOF) Then
                TXTUNIT = RSTUNIT!UNIT
            End If
            
            RSTUNIT.Close
            Set RSTUNIT = Nothing
            
            LBLITEM.Caption = grdPURCHASE.Columns(1)
            TXTQTY = grdPURCHASE.Columns(7)
            TXTBATCH = ""
            TXTEXPDATE = ""
            TXTMRP = ""
            TXTPTR = ""
            TXTTAX = 0
            TXTFREE = ""
            TXTEXPIRY.Text = "  /  "
            
            grdPURCHASE.Enabled = False
            FRMFIELDS.Enabled = True
            FRMECONTROLS.Enabled = True
            
            TXTQTY.SetFocus
                        
    End Select
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub


Private Sub MaskEdBox1_Change()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub TXTBATCH_GotFocus()
    TXTBATCH.SelStart = 0
    TXTBATCH.SelLength = Len(TXTBATCH.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTBATCH.Text) = "" Then Exit Sub
            TXTEXPIRY.SetFocus
                        
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("-")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTEXPDATE_GotFocus()
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
End Sub


Private Sub TXTEXPDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Not IsDate(TXTBATCH.Text) Then Exit Sub
            TXTMRP.SetFocus
                        
    End Select
End Sub

Private Sub TXTEXPDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
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
            
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
            TXTEXPDATE.SetFocus
                        
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    Dim M_DATE As Date
    Dim D As Integer
    Dim M As Integer
    Dim Y As Integer
    
    
    If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then
        TXTEXPDATE.Text = ""
        Exit Sub
    End If
    If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then
        TXTEXPDATE.Text = ""
        Exit Sub
    End If
    
    If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then
        TXTEXPDATE.Text = ""
        Exit Sub
    End If
    
    M = Val(Mid(TXTEXPIRY.Text, 1, 2))
    Y = Val(Right(TXTEXPIRY.Text, 2))
    Y = 2000 + Y
    M_DATE = "01" & "/" & M & "/" & Y
    D = LastDayOfMonth(M_DATE)
    M_DATE = D & "/" & M & "/" & Y
    TXTEXPDATE.Text = M_DATE
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTBATCH.Text)
    
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.Text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.Text = "" Then Exit Sub
            If TXTINVOICE.Text = "" Then
                TXTINVOICE.SetFocus
                Exit Sub
            End If
            
            FRMEPURCHASE.Enabled = True
            FRMEMASTER.Enabled = False
            grdPURCHASE.Enabled = True
            FRMFIELDS.Enabled = False
            FRMECONTROLS.Enabled = False
            
                        
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

Private Sub TXTINVOICE_GotFocus()
    TXTINVOICE.SelStart = 0
    TXTINVOICE.SelLength = Len(TXTINVOICE.Text)
End Sub

Private Sub TXTINVOICE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVOICE.Text = "" Then Exit Sub
            TXTINVDATE.SetFocus
                        
    End Select
End Sub

Private Sub TXTINVOICE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("-")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTMRP_GotFocus()
    TXTMRP.SelStart = 0
    TXTMRP.SelLength = Len(TXTMRP.Text)
End Sub

Private Sub TXTMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTMRP.Text) = 0 Then Exit Sub
            TXTPTR.SetFocus
                        
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

Private Sub TXTPTR_GotFocus()
    TXTPTR.SelStart = 0
    TXTPTR.SelLength = Len(TXTPTR.Text)
End Sub

Private Sub TXTPTR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTPTR.Text) = 0 Then Exit Sub
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

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 Then Exit Sub
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

Private Sub TXTTAX_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
End Sub

Private Sub TXTTAX_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMDSAVE.SetFocus
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

Private Sub TXTUNIT_GotFocus()
    TXTUNIT.SelStart = 0
    TXTUNIT.SelLength = Len(TXTUNIT.Text)
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.Text) = 0 Then Exit Sub
            TXTBATCH.SetFocus
                        
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

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function


Private Sub FILLitemcombo()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set cmbItem.DataSource = Nothing
    If ITEM_FLAG = True Then
        ITEM_REC.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME] From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.cmbItem.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        ITEM_REC.Close
        ITEM_REC.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME] From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.cmbItem.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    End If
    
    Set Me.cmbItem.RowSource = ITEM_REC
    cmbItem.ListField = "ITEM_NAME"
    cmbItem.BoundColumn = "ITEM_CODE"
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub
