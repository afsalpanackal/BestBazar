VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE FROM CUSTOMERS"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18585
   ControlBox      =   0   'False
   Icon            =   "FrmLPC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10995
   ScaleWidth      =   18585
   Begin VB.CommandButton CmdDeleteAll 
      Caption         =   "&Cancel Bill"
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
      Left            =   17070
      TabIndex        =   172
      Top             =   6900
      Width           =   1335
   End
   Begin VB.CheckBox Chkcancel 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Cancel Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   17085
      TabIndex        =   171
      Top             =   6630
      Width           =   1320
   End
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   2280
      TabIndex        =   116
      Top             =   3345
      Visible         =   0   'False
      Width           =   14820
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   2310
         Left            =   30
         TabIndex        =   117
         Top             =   270
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4075
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   13
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
            Size            =   9
            Charset         =   0
            Weight          =   400
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
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Index           =   2
         Left            =   3795
         TabIndex        =   119
         Top             =   15
         Width           =   11010
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   " PREVIOUS RATES FOR THE ITEM "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   1
         Left            =   30
         TabIndex        =   118
         Top             =   15
         Width           =   3780
      End
   End
   Begin VB.TextBox txtBillNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1260
      TabIndex        =   75
      Top             =   90
      Width           =   885
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
      Height          =   435
      Left            =   4395
      TabIndex        =   47
      Top             =   7755
      Width           =   1155
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   2280
      TabIndex        =   62
      Top             =   2010
      Visible         =   0   'False
      Width           =   12420
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   3900
         Left            =   15
         TabIndex        =   63
         Top             =   15
         Width           =   12390
         _ExtentX        =   21855
         _ExtentY        =   6879
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
            Weight          =   400
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
   Begin VB.Frame Fram 
      BackColor       =   &H00D2E4F4&
      Caption         =   "Frame1"
      Height          =   11040
      Left            =   -135
      TabIndex        =   48
      Top             =   -90
      Width           =   18690
      Begin VB.ComboBox Cmbbarcode 
         Height          =   315
         ItemData        =   "FrmLPC.frx":030A
         Left            =   14640
         List            =   "FrmLPC.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   176
         Top             =   1155
         Width           =   2865
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<<&Previous"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12945
         TabIndex        =   152
         Top             =   135
         Width           =   1125
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Next>>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12945
         TabIndex        =   151
         Top             =   525
         Width           =   1125
      End
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00D2E4F4&
         Height          =   1575
         Left            =   150
         TabIndex        =   65
         Top             =   0
         Width           =   12795
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
            Left            =   1245
            TabIndex        =   101
            Top             =   540
            Width           =   3735
         End
         Begin VB.TextBox TXTLASTBILL 
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
            Left            =   12840
            TabIndex        =   73
            Top             =   135
            Visible         =   0   'False
            Width           =   885
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
            Height          =   315
            Left            =   6045
            MaxLength       =   150
            TabIndex        =   108
            Top             =   810
            Width           =   4815
         End
         Begin VB.TextBox TXTDATE 
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
            Height          =   300
            Left            =   2820
            MaxLength       =   10
            TabIndex        =   71
            Top             =   210
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
            Left            =   6450
            MaxLength       =   20
            TabIndex        =   104
            Top             =   135
            Width           =   2445
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   6450
            TabIndex        =   106
            Top             =   465
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   255
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
            Left            =   1245
            TabIndex        =   102
            Top             =   885
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
         Begin MSDataListLib.DataCombo CMBPO 
            Height          =   1425
            Left            =   10875
            TabIndex        =   142
            Top             =   135
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   2514
            _Version        =   393216
            Appearance      =   0
            Style           =   1
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
         Begin MSMask.MaskEdBox TXTRCVDATE 
            Height          =   315
            Left            =   9075
            TabIndex        =   167
            Top             =   465
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   255
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
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "RCVD DATE"
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
            Left            =   7950
            TabIndex        =   168
            Top             =   495
            Width           =   1110
         End
         Begin MSForms.ComboBox CMBDISTRICT 
            Height          =   375
            Left            =   6045
            TabIndex        =   109
            Top             =   1140
            Width           =   4110
            VariousPropertyBits=   746604571
            ForeColor       =   255
            MaxLength       =   30
            DisplayStyle    =   3
            Size            =   "7250;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            DropButtonStyle =   0
            BorderColor     =   255
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "GODOWN"
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
            Left            =   5070
            TabIndex        =   159
            Top             =   1185
            Width           =   1290
         End
         Begin VB.Label lblcredit 
            Height          =   525
            Left            =   12960
            TabIndex        =   90
            Top             =   645
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "P.O No."
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
            Left            =   10155
            TabIndex        =   74
            Top             =   135
            Width           =   705
         End
         Begin VB.Label INVDATE 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   1
            Left            =   5070
            TabIndex        =   72
            Top             =   840
            Width           =   1290
         End
         Begin VB.Label INVDATE 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   0
            Left            =   2205
            TabIndex        =   70
            Top             =   210
            Width           =   780
         End
         Begin VB.Label Label1 
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
            Height          =   300
            Index           =   0
            Left            =   165
            TabIndex        =   69
            Top             =   210
            Width           =   870
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
            Left            =   5070
            TabIndex        =   68
            Top             =   165
            Width           =   1215
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
            Left            =   5070
            TabIndex        =   67
            Top             =   495
            Width           =   1335
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
            Left            =   150
            TabIndex        =   66
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4440
         Left            =   150
         TabIndex        =   164
         Top             =   1485
         Width           =   18555
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
            Height          =   290
            Left            =   285
            TabIndex        =   165
            Top             =   1215
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4335
            Left            =   15
            TabIndex        =   166
            Top             =   90
            Width           =   18510
            _ExtentX        =   32650
            _ExtentY        =   7646
            _Version        =   393216
            Rows            =   1
            Cols            =   42
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
            AllowUserResizing=   3
            Appearance      =   0
            GridLineWidth   =   2
         End
      End
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00D2E4F4&
         Height          =   5160
         Left            =   150
         TabIndex        =   49
         Top             =   5835
         Width           =   18480
         Begin VB.TextBox TxtTotalexp 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   450
            Left            =   11325
            MaxLength       =   15
            TabIndex        =   160
            Top             =   2370
            Width           =   1350
         End
         Begin VB.TextBox TxtNetrate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   16140
            MaxLength       =   11
            TabIndex        =   16
            Top             =   465
            Width           =   960
         End
         Begin VB.TextBox TxTfree 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   9075
            MaxLength       =   8
            TabIndex        =   7
            Top             =   465
            Width           =   495
         End
         Begin VB.ComboBox cmbfull 
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
            ItemData        =   "FrmLPC.frx":030E
            Left            =   7500
            List            =   "FrmLPC.frx":0360
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   495
            Width           =   825
         End
         Begin VB.CommandButton CmdLabels 
            Caption         =   "Print &Labels"
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
            Left            =   17070
            TabIndex        =   153
            Top             =   1605
            Width           =   1335
         End
         Begin VB.TextBox TxtCustDisc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   9885
            MaxLength       =   7
            TabIndex        =   30
            Top             =   1140
            Width           =   1140
         End
         Begin VB.CommandButton CmdPrint 
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
            Height          =   420
            Left            =   17055
            TabIndex        =   148
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox TxtCessPer 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   12705
            MaxLength       =   7
            TabIndex        =   34
            Top             =   1140
            Width           =   645
         End
         Begin VB.TextBox txtCess 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   13365
            MaxLength       =   7
            TabIndex        =   35
            Top             =   1140
            Width           =   915
         End
         Begin VB.TextBox txtHSN 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   60
            MaxLength       =   15
            TabIndex        =   17
            Top             =   1125
            Width           =   960
         End
         Begin VB.TextBox TxtBarcode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   480
            MaxLength       =   20
            TabIndex        =   1
            Top             =   480
            Width           =   1785
         End
         Begin VB.TextBox TxtLWRate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   8865
            MaxLength       =   7
            TabIndex        =   29
            Top             =   1140
            Width           =   1005
         End
         Begin VB.TextBox TxtTrDisc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   33
            Top             =   1140
            Width           =   855
         End
         Begin VB.TextBox TxtCSTper 
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
            Left            =   8070
            MaxLength       =   7
            TabIndex        =   41
            Top             =   4215
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox TxtExDuty 
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
            Left            =   7110
            MaxLength       =   7
            TabIndex        =   40
            Top             =   4215
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox TxtExpense 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   3570
            MaxLength       =   7
            TabIndex        =   20
            Top             =   1140
            Width           =   840
         End
         Begin VB.TextBox txtcategory 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   2280
            TabIndex        =   2
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox TxtWarranty 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   14295
            MaxLength       =   4
            TabIndex        =   36
            Top             =   1140
            Width           =   315
         End
         Begin VB.ComboBox CmbWrnty 
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
            ItemData        =   "FrmLPC.frx":03FC
            Left            =   14625
            List            =   "FrmLPC.frx":0406
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1155
            Width           =   840
         End
         Begin VB.TextBox TxtRetailPercent 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Left            =   4425
            MaxLength       =   7
            TabIndex        =   22
            Top             =   1545
            Width           =   945
         End
         Begin VB.TextBox txtWsalePercent 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Left            =   5385
            MaxLength       =   7
            TabIndex        =   24
            Top             =   1545
            Width           =   870
         End
         Begin VB.TextBox txtSchPercent 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Left            =   6270
            MaxLength       =   7
            TabIndex        =   26
            Top             =   1545
            Width           =   900
         End
         Begin VB.ComboBox CmbPack 
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
            ItemData        =   "FrmLPC.frx":0417
            Left            =   9585
            List            =   "FrmLPC.frx":0469
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   495
            Width           =   780
         End
         Begin VB.TextBox Los_Pack 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   7050
            MaxLength       =   7
            TabIndex        =   4
            Top             =   480
            Width           =   435
         End
         Begin VB.TextBox Txtgrossamt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   17115
            MaxLength       =   10
            TabIndex        =   15
            Top             =   465
            Width           =   1305
         End
         Begin VB.TextBox txtvanrate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   6270
            MaxLength       =   7
            TabIndex        =   25
            Top             =   1140
            Width           =   900
         End
         Begin VB.TextBox txtcrtnpack 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   7185
            MaxLength       =   7
            TabIndex        =   27
            Top             =   1140
            Width           =   705
         End
         Begin VB.TextBox TxtComper 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   11040
            MaxLength       =   7
            TabIndex        =   31
            Top             =   1140
            Width           =   720
         End
         Begin VB.TextBox TxtComAmt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   11775
            MaxLength       =   7
            TabIndex        =   32
            Top             =   1140
            Width           =   915
         End
         Begin VB.TextBox txtcrtn 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   7905
            MaxLength       =   7
            TabIndex        =   28
            Top             =   1140
            Width           =   945
         End
         Begin VB.TextBox txtWS 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   5385
            MaxLength       =   7
            TabIndex        =   23
            Top             =   1140
            Width           =   870
         End
         Begin VB.TextBox TXTRETAIL 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   4425
            MaxLength       =   7
            TabIndex        =   21
            Top             =   1140
            Width           =   945
         End
         Begin VB.TextBox txtPD 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   1035
            MaxLength       =   7
            TabIndex        =   18
            Top             =   1140
            Width           =   750
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
            Left            =   14520
            MaxLength       =   6
            TabIndex        =   97
            Top             =   3840
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtprofit 
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
            Height          =   300
            Left            =   15555
            MaxLength       =   7
            TabIndex        =   95
            Top             =   4005
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox txtaddlamt 
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
            Left            =   7815
            TabIndex        =   92
            Top             =   2205
            Width           =   1050
         End
         Begin VB.TextBox txtcramt 
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
            Left            =   6735
            TabIndex        =   91
            Top             =   2205
            Width           =   1050
         End
         Begin VB.TextBox TXTDISCAMOUNT 
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
            Left            =   5745
            TabIndex        =   84
            Top             =   2205
            Width           =   945
         End
         Begin VB.OptionButton OPTNET 
            BackColor       =   &H00D2E4F4&
            Caption         =   "NET"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   15735
            TabIndex        =   82
            Top             =   2985
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00D2E4F4&
            Caption         =   "Tax on MRP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   15720
            TabIndex        =   80
            Top             =   2505
            Width           =   1410
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00D2E4F4&
            Caption         =   "GSTAX %"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   15735
            TabIndex        =   81
            Top             =   2745
            Width           =   1395
         End
         Begin VB.TextBox TxttaxMRP 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   15420
            MaxLength       =   7
            TabIndex        =   14
            Top             =   465
            Width           =   705
         End
         Begin VB.TextBox Txtpack 
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
            Left            =   3375
            MaxLength       =   7
            TabIndex        =   76
            Top             =   3075
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox TXTPTR 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   11295
            MaxLength       =   11
            TabIndex        =   10
            Top             =   465
            Width           =   990
         End
         Begin VB.TextBox TXTRATE 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   10365
            MaxLength       =   7
            TabIndex        =   9
            Top             =   465
            Width           =   915
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
            Height          =   435
            Left            =   60
            TabIndex        =   43
            Top             =   2010
            Width           =   1095
         End
         Begin VB.TextBox TXTSLNO 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   45
            TabIndex        =   0
            Top             =   480
            Width           =   420
         End
         Begin VB.TextBox TXTPRODUCT 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   3420
            TabIndex        =   3
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox TXTQTY 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   8325
            MaxLength       =   8
            TabIndex        =   6
            Top             =   465
            Width           =   735
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
            Height          =   435
            Left            =   2280
            TabIndex        =   45
            Top             =   2010
            Width           =   1035
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
            Height          =   435
            Left            =   1185
            TabIndex        =   44
            Top             =   2010
            Width           =   1065
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
            Height          =   360
            Left            =   60
            TabIndex        =   51
            Top             =   3075
            Visible         =   0   'False
            Width           =   3300
         End
         Begin VB.TextBox txtBatch 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   12300
            MaxLength       =   35
            TabIndex        =   11
            Top             =   465
            Width           =   1905
         End
         Begin VB.TextBox TXTUNIT 
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
            Height          =   300
            Left            =   12945
            TabIndex        =   50
            Top             =   3945
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdRefresh 
            BackColor       =   &H00000080&
            Caption         =   "&Save"
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
            Height          =   435
            Left            =   3375
            TabIndex        =   46
            Top             =   2010
            Width           =   975
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   375
            Left            =   14220
            TabIndex        =   12
            Top             =   480
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   661
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
         Begin MSMask.MaskEdBox TXTEXPDATE 
            Height          =   375
            Left            =   14220
            TabIndex        =   13
            Top             =   480
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   661
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
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00D7F4F1&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   7185
            TabIndex        =   112
            Top             =   1455
            Width           =   2565
            Begin VB.OptionButton OptComAmt 
               BackColor       =   &H00D7F4F1&
               Caption         =   "Co&mm Amt"
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
               Left            =   1200
               TabIndex        =   39
               Top             =   195
               Width           =   1335
            End
            Begin VB.OptionButton OptComper 
               BackColor       =   &H00D7F4F1&
               Caption         =   "C&omm %"
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
               Left            =   15
               TabIndex        =   38
               Top             =   180
               Value           =   -1  'True
               Width           =   1155
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00D7F4F1&
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   6345
            TabIndex        =   125
            Top             =   2475
            Width           =   2565
            Begin VB.TextBox TxtInsurance 
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
               Left            =   1575
               TabIndex        =   127
               Top             =   510
               Width           =   945
            End
            Begin VB.TextBox TxtCST 
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
               Left            =   1575
               TabIndex        =   126
               Top             =   150
               Width           =   945
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Insurance Amt"
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
               Index           =   37
               Left            =   75
               TabIndex        =   129
               Top             =   525
               Width           =   1470
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "TCS %"
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
               Index           =   36
               Left            =   90
               TabIndex        =   128
               Top             =   195
               Width           =   1050
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00D7F4F1&
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   870
            TabIndex        =   120
            Top             =   1530
            Width           =   2070
            Begin VB.OptionButton optdiscper 
               BackColor       =   &H00D7F4F1&
               Caption         =   "D&isc %"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   15
               TabIndex        =   122
               Top             =   120
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton Optdiscamt 
               BackColor       =   &H00D7F4F1&
               Caption         =   "Di&sc Amt"
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
               Left            =   930
               TabIndex        =   121
               Top             =   135
               Width           =   1125
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00D7F4F1&
            Height          =   2415
            Left            =   8970
            TabIndex        =   135
            Top             =   2880
            Visible         =   0   'False
            Width           =   3855
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   2295
               Left            =   15
               Top             =   105
               Width           =   3825
            End
         End
         Begin VB.Label LBLNET 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00004080&
            Height          =   450
            Left            =   14190
            TabIndex        =   175
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label LBLGROSSAMT 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00008000&
            Height          =   450
            Left            =   13005
            TabIndex        =   174
            Top             =   1725
            Width           =   1470
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "GROSS AMT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   59
            Left            =   13005
            TabIndex        =   173
            Top             =   1500
            Width           =   1470
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL QTY"
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
            Height          =   225
            Index           =   58
            Left            =   12900
            TabIndex        =   170
            Top             =   2160
            Width           =   1080
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblqty 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00004080&
            Height          =   450
            Left            =   12690
            TabIndex        =   169
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label LBLEXP 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00800000&
            Height          =   450
            Left            =   15885
            TabIndex        =   163
            Top             =   1725
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL EXP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   57
            Left            =   15900
            TabIndex        =   162
            Top             =   1500
            Width           =   1140
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Enter total expenses here"
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
            Height          =   480
            Index           =   56
            Left            =   9780
            TabIndex        =   161
            Top             =   2355
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TAX AMOUNT"
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
            Height          =   210
            Index           =   55
            Left            =   14460
            TabIndex        =   158
            Top             =   1500
            Width           =   1470
            WordWrap        =   -1  'True
         End
         Begin VB.Label LBLTOTALTAX 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00008000&
            Height          =   450
            Left            =   14490
            TabIndex        =   157
            Top             =   1725
            Width           =   1380
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
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   54
            Left            =   16140
            TabIndex        =   156
            Top             =   195
            Width           =   960
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
            Height          =   285
            Index           =   53
            Left            =   9585
            TabIndex        =   155
            Top             =   195
            Width           =   750
         End
         Begin VB.Label LBLPRE 
            Height          =   330
            Left            =   13275
            TabIndex        =   154
            Top             =   3780
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Cust Disc%"
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
            Index           =   52
            Left            =   9885
            TabIndex        =   150
            Top             =   885
            Width           =   1140
         End
         Begin VB.Label lblcategory 
            Height          =   345
            Left            =   15780
            TabIndex        =   149
            Top             =   2475
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Cess%"
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
            Index           =   51
            Left            =   12705
            TabIndex        =   147
            Top             =   885
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Adl. Cess"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   49
            Left            =   13365
            TabIndex        =   146
            Top             =   885
            Width           =   915
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Item Code /"
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
            Height          =   285
            Index           =   40
            Left            =   2280
            TabIndex        =   145
            Top             =   195
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "HSN Code"
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
            Index           =   48
            Left            =   60
            TabIndex        =   144
            Top             =   885
            Width           =   960
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
            ForeColor       =   &H008080FF&
            Height          =   285
            Index           =   47
            Left            =   480
            TabIndex        =   143
            Top             =   195
            Width           =   1785
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L. W. Rate"
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
            Index           =   46
            Left            =   8865
            TabIndex        =   141
            Top             =   885
            Width           =   1005
         End
         Begin VB.Label LblGross 
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
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   4170
            TabIndex        =   140
            Top             =   3045
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Gross"
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
            Index           =   45
            Left            =   4170
            TabIndex        =   139
            Top             =   2790
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Scheme"
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
            Index           =   44
            Left            =   1800
            TabIndex        =   138
            Top             =   885
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "CST %"
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
            Index           =   43
            Left            =   8070
            TabIndex        =   137
            Top             =   3960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Ex Duty%"
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
            Index           =   42
            Left            =   7110
            TabIndex        =   136
            Top             =   3960
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lbltaxamount 
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
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2670
            TabIndex        =   19
            Top             =   1155
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Expense"
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
            Index           =   41
            Left            =   3585
            TabIndex        =   134
            Top             =   885
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Warranty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   39
            Left            =   14295
            TabIndex        =   131
            Top             =   885
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "% of   Profit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   390
            Index           =   38
            Left            =   3570
            TabIndex        =   130
            Top             =   1545
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Product Code"
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
            Height          =   285
            Index           =   35
            Left            =   60
            TabIndex        =   124
            Top             =   2790
            Visible         =   0   'False
            Width           =   3300
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Loose Pack"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   285
            Index           =   34
            Left            =   7050
            TabIndex        =   123
            Top             =   195
            Width           =   1260
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Gross Amt"
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
            Height          =   285
            Index           =   33
            Left            =   17115
            TabIndex        =   115
            Top             =   195
            Width           =   1305
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "V. Rate"
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
            Left            =   6270
            TabIndex        =   114
            Top             =   885
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L. Pack"
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
            Index           =   31
            Left            =   7185
            TabIndex        =   113
            Top             =   885
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Comi %"
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
            Index           =   30
            Left            =   11040
            TabIndex        =   111
            Top             =   885
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Comi Amt"
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
            Index           =   29
            Left            =   11775
            TabIndex        =   110
            Top             =   885
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L. R. Rate"
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
            Index           =   28
            Left            =   7905
            TabIndex        =   107
            Top             =   885
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "W. Rate"
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
            Left            =   5385
            TabIndex        =   105
            Top             =   885
            Width           =   870
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
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   26
            Left            =   15555
            TabIndex        =   103
            Top             =   3765
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Disc"
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
            Index           =   25
            Left            =   1035
            TabIndex        =   98
            Top             =   885
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "R. Rate"
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
            Left            =   4425
            TabIndex        =   96
            Top             =   885
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TCS Amt"
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
            Index           =   23
            Left            =   7845
            TabIndex        =   94
            Top             =   1965
            Width           =   1020
         End
         Begin VB.Label Label1 
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
            Height          =   300
            Index           =   22
            Left            =   6735
            TabIndex        =   93
            Top             =   1965
            Width           =   1140
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
            ForeColor       =   &H00000080&
            Height          =   210
            Index           =   21
            Left            =   11595
            TabIndex        =   89
            Top             =   1500
            Width           =   1185
            WordWrap        =   -1  'True
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   450
            Left            =   11385
            TabIndex        =   88
            Top             =   1725
            Width           =   1605
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Disc. Amt"
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
            Index           =   19
            Left            =   5745
            TabIndex        =   87
            Top             =   1980
            Width           =   915
         End
         Begin VB.Label lbltotalwodiscount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   450
            Left            =   9765
            TabIndex        =   86
            Top             =   1725
            Width           =   1605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PURCHASE AMT"
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
            Height          =   225
            Index           =   6
            Left            =   9780
            TabIndex        =   85
            Top             =   1500
            Width           =   1620
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "FREE"
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
            Height          =   285
            Index           =   17
            Left            =   9075
            TabIndex        =   83
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Tax Amt"
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
            Index           =   13
            Left            =   2670
            TabIndex        =   79
            Top             =   885
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "GSTax%"
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
            Height          =   285
            Index           =   12
            Left            =   15420
            TabIndex        =   78
            Top             =   195
            Width           =   705
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
            Height          =   285
            Index           =   4
            Left            =   3375
            TabIndex        =   77
            Top             =   2790
            Visible         =   0   'False
            Width           =   780
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
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   2
            Left            =   11295
            TabIndex        =   64
            Top             =   195
            Width           =   990
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
            ForeColor       =   &H008080FF&
            Height          =   285
            Index           =   8
            Left            =   45
            TabIndex        =   61
            Top             =   195
            Width           =   420
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
            ForeColor       =   &H008080FF&
            Height          =   285
            Index           =   9
            Left            =   3525
            TabIndex        =   60
            Top             =   195
            Width           =   3510
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
            ForeColor       =   &H008080FF&
            Height          =   285
            Index           =   10
            Left            =   8325
            TabIndex        =   59
            Top             =   195
            Width           =   735
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
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   11
            Left            =   10365
            TabIndex        =   58
            Top             =   195
            Width           =   915
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
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   14
            Left            =   15465
            TabIndex        =   57
            Top             =   885
            Width           =   1530
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
            Left            =   11085
            TabIndex        =   56
            Top             =   3885
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
            ForeColor       =   &H008080FF&
            Height          =   285
            Index           =   16
            Left            =   14220
            TabIndex        =   55
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Batch / Serial No."
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
            Index           =   7
            Left            =   12300
            TabIndex        =   54
            Top             =   210
            Width           =   1905
         End
         Begin VB.Label LBLSUBTOTAL 
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
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   15465
            TabIndex        =   42
            Top             =   1125
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Sell Unit"
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
            Index           =   20
            Left            =   12945
            TabIndex        =   53
            Top             =   3660
            Visible         =   0   'False
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
            Left            =   13275
            TabIndex        =   52
            Top             =   3615
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Add:"
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
         Height          =   360
         Index           =   61
         Left            =   14130
         TabIndex        =   179
         Top             =   120
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbladdress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   825
         Left            =   14595
         TabIndex        =   178
         Top             =   120
         Width           =   4035
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode Printer"
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
         Height          =   255
         Index           =   60
         Left            =   14640
         TabIndex        =   177
         Top             =   915
         Width           =   1620
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This month"
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
         Height          =   360
         Index           =   50
         Left            =   12945
         TabIndex        =   133
         Top             =   900
         Width           =   1680
         WordWrap        =   -1  'True
      End
      Begin VB.Label LBLmonth 
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   12945
         TabIndex        =   132
         Top             =   1110
         Width           =   1680
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   135
         TabIndex        =   100
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   705
         TabIndex        =   99
         Top             =   45
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmLPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytData() As Byte
Dim ACT_PO As New ADODB.Recordset
Dim PO_FLAG As Boolean
Dim PHY As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim ACT_FLAG As Boolean
Dim PHY_CODE As New ADODB.Recordset
Dim PHYCODE_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT, M_ADD, OLD_BILL As Boolean
Dim PHY_PRERATE As New ADODB.Recordset
Dim PRERATE_FLAG As Boolean
Dim PONO As String
Dim CHANGE_FLAG As Boolean
Dim BARCODE_FLAG As Boolean
Dim ADDCLICK As Boolean
Dim BARPRINTER As String

Private Sub Cmbbarcode_Click()
    BARPRINTER = Cmbbarcode.ListIndex
End Sub

Private Sub CMBDISTRICT_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            FRMECONTROLS.Enabled = True
            If CMBPO.VisibleCount = 0 Then
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            Else
                CMBPO.SetFocus
            End If
        Case vbKeyEscape
            TXTREMARKS.SetFocus
    End Select
End Sub

Private Sub cmbfull_GotFocus()
    cmbfull.BackColor = &H98F3C1
End Sub

Private Sub cmbfull_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If cmbfull.ListIndex = -1 Then cmbfull.ListIndex = 0
            TXTQTY.SetFocus
         Case vbKeyEscape
            'TXTUNIT.Text = ""
            Los_Pack.SetFocus
    End Select
End Sub

Private Sub cmbfull_LostFocus()
    cmbfull.BackColor = vbWhite
End Sub

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = "1"
            TXTRATE.SetFocus
         Case vbKeyEscape
            'TXTUNIT.Text = ""
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub CMBPO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMBPO.VisibleCount <> 0 And CMBPO.BoundText = "" Then
                If (MsgBox("Are you sure you want to continue without selecting the Purchase Order No.? !!!!", vbYesNo, "EzBiz") = vbNo) Then
                    CMBPO.SetFocus
                    Exit Sub
                End If
            End If
            If CMBPO.VisibleCount = 0 Then CMBPO.Text = ""
            If CMBPO.Text <> "" And CMBPO.MatchedWithList = False Then
                MsgBox "Please select a valid PO No. from the list", vbOKOnly, "EzBiz"
                On Error Resume Next
                CMBPO.SetFocus
                Exit Sub
            End If
            FRMECONTROLS.Enabled = True
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTREMARKS.SetFocus
    End Select
End Sub

Private Sub CMBPO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMDADD_Click()
    
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
    ADDCLICK = True
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
    If Val(Los_Pack.Text) = 1 Then
         TxtLWRate.Text = Val(txtWS.Text)
         txtcrtn.Text = Val(TXTRETAIL.Text)
    End If
    
    If Val(TXTQTY.Text) = 0 Then
        MsgBox "Please enter the Qty", vbOKOnly, "EzBiz"
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
        Exit Sub
    End If
    If Val(TXTPTR.Text) = 0 Then
        MsgBox "Please enter the Price", vbOKOnly, "EzBiz"
        TXTPTR.SetFocus
        Exit Sub
    End If
        
    If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
        If optdiscper.value = True Then
            txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxTfree.Text)), 3)
            TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) + ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
        Else
            txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxTfree.Text)), 3)
            TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) + ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
        End If
    Else
        If optdiscper.value = True Then
            txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxTfree.Text)), 3)
            TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
        Else
            txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxTfree.Text)), 3)
            TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
        End If
    End If
    
    If Val(TXTRATE.Text) <> 0 And Val(TXTRATE.Text) < Val(TXTPTR.Tag) Then
        MsgBox "MRP less than cost", vbOKOnly, "Purchase....."
        TXTRATE.SetFocus
        Exit Sub
    End If
    
    If Val(TXTRATE.Text) <> 0 And Val(TXTRETAIL.Text) <> 0 And Val(TXTRETAIL.Text) > Val(TXTRATE.Text) Then
        MsgBox "Retail Price greater than MRP", vbOKOnly, "EzBiz"
        TXTRETAIL.SetFocus
        Exit Sub
    End If
    
    If Val(TXTRATE.Text) <> 0 And Val(txtWS.Text) <> 0 And Val(txtWS.Text) > Val(TXTRATE.Text) Then
        MsgBox "WS Price greater than MRP", vbOKOnly, "EzBiz"
        txtWS.SetFocus
        Exit Sub
    End If
    
    If Val(TXTRATE.Text) <> 0 And Val(txtvanrate.Text) <> 0 And Val(txtvanrate.Text) > Val(TXTRATE.Text) Then
        MsgBox "VAN Price greater than MRP", vbOKOnly, "EzBiz"
        txtvanrate.SetFocus
        Exit Sub
    End If
    
    If Val(TXTRETAIL.Text) <> 0 And Val(TXTRETAIL.Text) < Val(TXTPTR.Tag) Then
        If MsgBox("Retail Price less than cost. Are you sure?", vbYesNo + vbDefaultButton2, "Purchase.....") = vbNo Then
            TXTRETAIL.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtWS.Text) <> 0 And Val(txtWS.Text) < Val(TXTPTR.Tag) Then
        If MsgBox("WS Price less than cost. Are you sure?", vbYesNo + vbDefaultButton2, "Purchase.....") = vbNo Then
            txtWS.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtvanrate.Text) <> 0 And Val(txtvanrate.Text) < Val(TXTPTR.Tag) Then
        If MsgBox("Van Price less than cost. Are you sure?", vbYesNo + vbDefaultButton2, "Purchase.....") = vbNo Then
            txtvanrate.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(TXTRETAIL.Text) <> 0 And Val(txtcrtn.Text) <> 0 And Val(TXTRETAIL.Text) < Val(txtcrtn.Text) Then
        MsgBox "Retail Price less than Loose Price", vbOKOnly, "EzBiz"
        TXTRETAIL.SetFocus
        Exit Sub
    End If
    
    If Val(txtWS.Text) <> 0 And Val(TxtLWRate.Text) <> 0 And Val(txtWS.Text) < Val(TxtLWRate.Text) Then
        MsgBox "WS Price less than Loose Price", vbOKOnly, "EzBiz"
        txtWS.SetFocus
        Exit Sub
    End If
    
    'Call TXTPTR_LostFocus
    Call TXTQTY_LostFocus
    'Call Txtgrossamt_LostFocus
    Call txtPD_LostFocus
    Call txtcrtn_GotFocus
    Call TxtLWRate_GotFocus
    
    ADDCLICK = False
    txtcrtn.BackColor = vbWhite
    TxtLWRate.BackColor = vbWhite
    TXTRATE.BackColor = vbWhite
    
    Dim I As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim M_DATA As Long
    
    M_DATA = 0
    Txtpack.Text = 1
    If MDIMAIN.StatusBar.Panels(6).Text = "Y" Then
        If Trim(TxtBarcode.Text) = "" Or Trim(TXTITEMCODE.Text) = Left(Trim(TxtBarcode.Text), Len(Trim(TXTITEMCODE.Text))) Then '(Trim(TxtBarcode.Text) = Trim(TXTITEMCODE.Text) & Val(LBLPRE.Caption)) Then
            TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text)
'            If Trim(Txtsize.Text) = "" Then
'                TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text) & Trim(cmbcolor.Text)
'            Else
'
'            End If
'        Else
'            If Trim(Txtsize.Text) = "" Then
'                TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text) & Trim(cmbcolor.Text)
'            Else
'                TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text) & Left(Trim(Txtsize.Text), 2)
'            End If
        End If
    End If
    
    'If MDIMAIN.StatusBar.Panels(6).Text = "Y" Then If Trim(TxtBarcode.Text) = "" Then TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text)
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text) + Val(TxTfree.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = 1 'Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Val(Los_Pack.Text) ' 1 'Val(TxtPack.Text)
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Round(Val(TXTRATE.Text) / Val(Los_Pack.Text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTRATE.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Format(Round(((Val(LblGross.Caption) / (Val(Los_Pack.Text) * (Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)) * Val(Los_Pack.Text))))), 4), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Format(Round(Val(TXTPTR.Text) / Val(Los_Pack.Text), 4), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Format((Val(txtprofit.Text)), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = IIf(Val(TxttaxMRP.Text) = 0, "", Format(Val(TxttaxMRP.Text), ".00")) 'TAX
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = IIf(Trim(TXTEXPDATE.Text) = "/  /", "", TXTEXPDATE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Val(TxTfree.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = Val(txtPD.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = Format(Val(TXTRETAIL.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = Format(Val(txtWS.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = Format(Val(txtvanrate.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 26) = Format(Val(Txtgrossamt.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = Format(Val(txtcrtn.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 37) = Format(Val(TxtLWRate.Text), ".000")
    If OptComAmt.value = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 21) = ""
        grdsales.TextMatrix(Val(TXTSLNO.Text), 22) = Format(Val(TxtComAmt.Text), ".00")
        grdsales.TextMatrix(Val(TXTSLNO.Text), 23) = "A"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.Text), 21) = Format(Val(TxtComper.Text), ".00")
        grdsales.TextMatrix(Val(TXTSLNO.Text), 22) = ""
        grdsales.TextMatrix(Val(TXTSLNO.Text), 23) = "P"
    End If
    If optdiscper.value = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 27) = "P"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.Text), 27) = "A"
    End If
    grdsales.TextMatrix(Val(TXTSLNO.Text), 28) = Format(Val(Los_Pack.Text), ".00")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 29) = Trim(CmbPack.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 30) = Val(TxtWarranty.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 31) = Trim(CmbWrnty.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = Val(TxtExpense.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 33) = Val(TxtExDuty.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 34) = Val(TxtCSTper.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 35) = Val(TxtTrDisc.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 36) = Val(LblGross.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 38) = Trim(TxtBarcode.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 39) = Val(txtCess.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 40) = Val(TxtCessPer.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = Format(Val(txtcrtnpack.Text), ".000")
    If Val(TxttaxMRP.Text) = 0 Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = "N"
    Else
        If OPTTaxMRP.value = True Then
            grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = "M"
        ElseIf OPTVAT.value = True Then
            grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = "V"
        End If
    End If
    
    If M_EDIT = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16))
    Else
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Val(TXTSLNO.Text)
    End If
    
    On Error GoTo ErrHand
    'If OLD_BILL = False Then Call checklastbill
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    If OLD_BILL = False And Val(txtBillNo.Text) <> 1 Then
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE VCH_NO= (SELECT MAX(VCH_NO) FROM TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'PI')", db, adOpenStatic, adLockOptimistic, adCmdText
        txtBillNo.Text = RSTTRXFILE!VCH_NO + 1
        db.BeginTrans
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "PI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = txtBillNo.Text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE.Update
    Else
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "PI"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
            RSTTRXFILE!VCH_NO = txtBillNo.Text
            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE.Update
        End If
    End If
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!TRX_TYPE = "PI"
        RSTRTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTRTRXFILE!line_no = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16))
        RSTRTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1))
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
        RSTRTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))

        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
'                If UCase(rststock!CATEGORY) = "CUTSHEET" Then
'                Else
                .Properties("Update Criteria").value = adCriteriaKey
                !ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                '!CLOSE_VAL = !CLOSE_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !CLOSE_VAL = Round(!ITEM_COST * !CLOSE_QTY, 3)
                !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                '!RCPT_VAL = !RCPT_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !RCPT_VAL = Round(!ITEM_COST * !RCPT_QTY, 3)
            
                !MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
                If Trim(txtHSN.Text) <> "" Then !Remarks = Trim(txtHSN.Text)
                If cmbfull.ListIndex <> -1 Then !FULL_PACK = cmbfull.Text
                !CUST_DISC = Val(TxtCustDisc.Text)
                If !PRICE_CHANGE = "Y" And Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) <> 0 Then
                    db.Execute "Update RTRXFILE set P_RETAIL = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) & " WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "' AND BAL_QTY >0 "
                End If
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 38)) <> "" Then !BARCODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 38))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)) <> 0 Then !P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) <> 0 Then !P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37)) <> 0 Then !P_LWS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) <> 0 Then !P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) ' / Val(Los_Pack.Text), 3)
                'If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) <> 0 Then
                !CESS_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) ' / Val(Los_Pack.Text), 3)
                'If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) <> 0 Then
                !CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) ' / Val(Los_Pack.Text), 3)
                
                '!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))) <> 0 Then !CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))

                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23)) = "A" Then
                    !COM_FLAG = "A"
                    !COM_PER = 0
                    !COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22))
                Else
                    !COM_FLAG = "P"
                    !COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21))
                    !COM_AMT = 0
                End If
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))) >= 5 Then !SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
                '!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
                !CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))
                !LOOSE_PACK = Val(Los_Pack.Text)
                !PACK_TYPE = Trim(CmbPack.Text)
                !WARRANTY = Val(TxtWarranty.Text)
                !WARRANTY_TYPE = Trim(CmbWrnty.Text)
                RSTRTRXFILE!MFGR = !MANUFACTURER
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        
    Else
        M_DATA = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
        M_DATA = M_DATA - (RSTRTRXFILE!QTY - RSTRTRXFILE!BAL_QTY)
        RSTRTRXFILE!BAL_QTY = M_DATA
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").value = adCriteriaKey
                '!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
                !ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
                !CLOSE_QTY = !CLOSE_QTY - RSTRTRXFILE!QTY
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                '!CLOSE_VAL = !CLOSE_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !CLOSE_VAL = Round(!ITEM_COST * !CLOSE_QTY, 3)
                
                !RCPT_QTY = !RCPT_QTY - RSTRTRXFILE!QTY
                !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                '!RCPT_VAL =  !RCPT_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !RCPT_VAL = Round(!ITEM_COST * !RCPT_QTY, 3)
                
                If Trim(txtHSN.Text) <> "" Then !Remarks = Trim(txtHSN.Text)
                If cmbfull.ListIndex <> -1 Then !FULL_PACK = cmbfull.Text
                !CUST_DISC = Val(TxtCustDisc.Text)
            
                If !PRICE_CHANGE = "Y" And Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) <> 0 Then
                    db.Execute "Update RTRXFILE set P_RETAIL = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) & " WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "' AND BAL_QTY >0 "
                End If
                
                !MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 38)) <> "" Then !BARCODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 38))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)) <> 0 Then !P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) <> 0 Then !P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37)) <> 0 Then !P_LWS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) <> 0 Then !P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) ' / Val(Los_Pack.Text), 3)
                'If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) <> 0 Then
                !CESS_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) ' / Val(Los_Pack.Text), 3)
                'If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) <> 0 Then
                !CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) ' / Val(Los_Pack.Text), 3)

                '!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))) <> 0 Then !CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
                                    
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23)) = "A" Then
                    !COM_FLAG = "A"
                    !COM_PER = 0
                    !COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22))
                Else
                    !COM_FLAG = "P"
                    !COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21))
                    !COM_AMT = 0
                End If
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))) >= 5 Then !SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
                '!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
                !CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))
                !LOOSE_PACK = Val(Los_Pack.Text)
                !PACK_TYPE = Trim(CmbPack.Text)
                !WARRANTY = Val(TxtWarranty.Text)
                !WARRANTY_TYPE = Trim(CmbWrnty.Text)
                RSTRTRXFILE!MFGR = !MANUFACTURER
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
    End If
    
    RSTRTRXFILE!Category = Trim(lblcategory.Caption)
    RSTRTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
    RSTRTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "dd/mm/yyyy")
    RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2))
    RSTRTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
    RSTRTRXFILE!ITEM_COST_PRICE = Round(Val(TXTPTR.Text), 3)
    RSTRTRXFILE!ITEM_NET_COST_PRICE = Round((Val(LBLSUBTOTAL.Caption) / TXTQTY.Text) + Val(TxtExpense.Text), 3)
    
    RSTRTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
    RSTRTRXFILE!P_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
    RSTRTRXFILE!MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
    RSTRTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 9))
    RSTRTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
    RSTRTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
    RSTRTRXFILE!P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19))
    RSTRTRXFILE!P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
    RSTRTRXFILE!P_LWS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37))
    RSTRTRXFILE!CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
    RSTRTRXFILE!P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25))
    RSTRTRXFILE!GROSS_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 26))
    RSTRTRXFILE!BARCODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 38))
    RSTRTRXFILE!CESS_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39))
    RSTRTRXFILE!CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40))
    If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23)) = "A" Then
        RSTRTRXFILE!COM_FLAG = "A"
        RSTRTRXFILE!COM_PER = 0
        RSTRTRXFILE!COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22))
    Else
        RSTRTRXFILE!COM_FLAG = "P"
        RSTRTRXFILE!COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21))
        RSTRTRXFILE!COM_AMT = 0
    End If
    RSTRTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
    RSTRTRXFILE!LOOSE_PACK = Val(Los_Pack.Text)
    RSTRTRXFILE!PACK_TYPE = Trim(CmbPack.Text)
    RSTRTRXFILE!WARRANTY = Val(TxtWarranty.Text)
    RSTRTRXFILE!WARRANTY_TYPE = Trim(CmbWrnty.Text)
    RSTRTRXFILE!EXPENSE = Val(TxtExpense.Text)
    RSTRTRXFILE!EXDUTY = Val(TxtExDuty.Text)
    RSTRTRXFILE!CSTPER = Val(TxtCSTper.Text)
    RSTRTRXFILE!TR_DISC = Val(TxtTrDisc.Text)
    
    RSTRTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
    'RSTRTRXFILE!VCH_DESC = "Received From " & DataList2.Text
    RSTRTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
    'RSTRTRXFILE!ISSUE_QTY = 0
    RSTRTRXFILE!CST = 0
    If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 27)) = "P" Then
        RSTRTRXFILE!DISC_FLAG = "P"
    Else
        RSTRTRXFILE!DISC_FLAG = "A"
    End If
    RSTRTRXFILE!SCHEME = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))
    'RSTRTRXFILE!EXP_DATE = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = "", Null, Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "dd/mm/yyyy"))
    If IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 12)) Then
        RSTRTRXFILE!EXP_DATE = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = "", Null, Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "dd/mm/yyyy"))
    End If
    RSTRTRXFILE!FREE_QTY = 0
    RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!C_USER_ID = "SM"
    RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))
    
    'RSTRTRXFILE!M_USER_ID = DataList2.BoundText
    ''''RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))  'MODE OF TAX
    'RSTRTRXFILE!PINV = Trim(TXTINVOICE.Text)
    RSTRTRXFILE.Update
    db.CommitTrans
    RSTRTRXFILE.Close
    
    M_DATA = 0
    Set RSTRTRXFILE = Nothing
           
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    Dim GROSSVAL As Double
    For I = 1 To grdsales.Rows - 1
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(I, 13)), ".00")
        LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(I, 32)), ".00")
        lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(I, 3)), ".00")
        GROSSVAL = (Val(grdsales.TextMatrix(I, 9)) * IIf(Val(grdsales.TextMatrix(I, 5)) = 0, 1, Val(grdsales.TextMatrix(I, 5)))) * (Val(grdsales.TextMatrix(I, 3)) - Val(grdsales.TextMatrix(I, 14)))
        If Trim(grdsales.TextMatrix(I, 27)) = "P" Then
            LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - (GROSSVAL * Val(grdsales.TextMatrix(I, 17)) / 100)) * Val(grdsales.TextMatrix(I, 10)) / 100, 2))
        ElseIf Trim(grdsales.TextMatrix(I, 27)) = "A" Then
            LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - Val(grdsales.TextMatrix(I, 17))) * Val(grdsales.TextMatrix(I, 10)) / 100, 2))
        End If
        LBLGROSSAMT.Caption = Val(LBLGROSSAMT.Caption) + Val(grdsales.TextMatrix(I, 8)) * Val(grdsales.TextMatrix(I, 3)) * Val(grdsales.TextMatrix(I, 5))
    Next I
    
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    If Roundflag = True Then
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    Else
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 2), "0.00")
    End If
    LBLNET.Caption = Val(LBLGROSSAMT.Caption) + Val(LBLTOTALTAX.Caption)
    
    If MDIMAIN.StatusBar.Panels(6).Text = "Y" Then
        If MsgBox("Do you want to Print Barcode Labels now?", vbYesNo + vbDefaultButton2, "Purchase.....") = vbYes Then
            I = Val(InputBox("Enter number of lables to be print", "No. of labels..", Val(TXTQTY.Text) + Val(TxTfree.Text)))
'            If MDIMAIN.barcode_profile.Caption = 0 Then
'                If i > 0 Then Call print_3labels(i, Trim(TxtBarcode.Text), Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2)), Val(TXTRATE.Text), Val(txtretail.Text))
'                '(i As Long, BAR_LABEL As String, itemname As String, itemmrp As Double, itemprice As Double)
'            Else
'                If i > 0 Then Call print_labels(i, Trim(TxtBarcode.Text), Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2)), Val(TXTRATE.Text), Val(txtretail.Text))
'            End If
            Dim M, N As Integer
            db.Execute "Delete from barprint"
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From barprint", db, adOpenStatic, adLockOptimistic, adCmdText
            For M = 1 To I
                RSTTRXFILE.AddNew
                RSTTRXFILE!BARCODE = "*" & grdsales.TextMatrix(Val(TXTSLNO.Text), 38) & "*"
                RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2))
                RSTTRXFILE!item_Price = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
                RSTTRXFILE!item_MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
                If IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 12)) Then
                    RSTTRXFILE!expdate = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "dd/mm/yyyy")
                    If IsDate(TXTINVDATE.Text) Then
                        RSTTRXFILE!pckdate = Format(TXTINVDATE.Text, "dd/mm/yyyy")
                    End If
                End If
                RSTTRXFILE!item_color = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
                RSTTRXFILE!Remarks = ""
                RSTTRXFILE!COMP_NAME = Trim(MDIMAIN.StatusBar.Panels(5).Text)
                RSTTRXFILE.Update
            Next M
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            If BARPRINTER = barcodeprinter Then
                ReportNameVar = Rptpath & "Rptbarprn"
            Else
                ReportNameVar = Rptpath & "Rptbarprn1"
            End If
            Set Report = crxApplication.OpenReport(ReportNameVar, 1)
            Set CRXFormulaFields = Report.FormulaFields
        
            For N = 1 To Report.Database.Tables.COUNT
                Report.Database.Tables.Item(N).SetLogOnInfo strConnection
                If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                    Set oRs = New ADODB.Recordset
                    Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(N).Name & " ")
                    Report.Database.SetDataSource oRs, 3, N
                    Set oRs = Nothing
                End If
            Next N
            
            Set Printer = Printers(BARPRINTER)
            Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
            Report.DiscardSavedData
            Report.VerifyOnEveryPrint = True
            Report.PrintOut (False)
            Set CRXFormulaFields = Nothing
            Set crxApplication = Nothing
            Set Report = Nothing
        Else
            If BARCODE_FLAG = False Then grdsales.TextMatrix(Val(TXTSLNO.Text), 41) = Val(TXTQTY.Text) + Val(TxTfree.Text) 'Val(InputBox("Enter number of lables to be print", "No. of labels..", Val(TXTQTY.Text) + Val(TxtFree.Text)))
        End If
    End If
    BARCODE_FLAG = False
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TxtBarcode.Text = ""
    TXTPTR.Text = ""
    TxtNetrate.Text = ""
    Txtgrossamt.Text = ""
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.Text = ""
    TxttaxMRP.Text = ""
    TxtExDuty.Text = ""
    TxtCSTper.Text = ""
    TxtTrDisc.Text = ""
    TxtCustDisc.Text = ""
    TxtCessPer.Text = ""
    txtCess.Text = ""
    txtPD.Text = ""
    TxtExpense.Text = ""
    txtprofit.Text = ""
    TXTRETAIL.Text = ""
    TxtRetailPercent.Text = ""
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    TxtLWRate.Text = ""
    txtcrtnpack.Text = ""
    TXTRATE.Text = ""
    TxtComAmt.Text = ""
    TxtComper.Text = ""
    txtmrpbt.Text = ""
    txtBatch.Text = ""
    txtHSN.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    lblcategory.Caption = ""
    LBLPRE.Caption = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    'optnet.value = True
    'OptComper.value = True
    M_ADD = True
    Chkcancel.value = 0
    OLD_BILL = True
    'txtcategory.Enabled = True
    txtBillNo.Enabled = False
    FRMEGRDTMP.Visible = False
    cmdRefresh.Enabled = True
    Los_Pack.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxttaxMRP.Enabled = False
    TxtExDuty.Enabled = False
    TxtTrDisc.Enabled = False
    TxtCessPer.Enabled = False
    txtCess.Enabled = False
    TxtCSTper.Enabled = False
    txtPD.Enabled = False
    TxtExpense.Enabled = False
    TXTRETAIL.Enabled = False
    TxtRetailPercent.Enabled = False
    txtWS.Enabled = False
    txtWsalePercent.Enabled = False
    txtvanrate.Enabled = False
    txtSchPercent.Enabled = False
    txtcrtnpack.Enabled = False
    txtcrtn.Enabled = False
    TxtLWRate.Enabled = False
    TxtCustDisc.Enabled = False
    TxtComper.Enabled = False
    TxtComAmt.Enabled = False
    OptComper.Enabled = False
    OptComAmt.Enabled = False
    cmdadd.Enabled = False
    txtBatch.Enabled = False
    txtHSN.Enabled = False
    TxtWarranty.Enabled = False
    CmbWrnty.Enabled = False
    TXTEXPIRY.Visible = False
    TXTEXPDATE.Enabled = False
    txtcategory.Enabled = True
    TXTPRODUCT.Enabled = True
    TxtBarcode.Enabled = True
    Txtgrossamt.Enabled = False
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
    If M_EDIT = True Then
        TXTSLNO.Enabled = True
        txtcategory.Enabled = False
        grdsales.SetFocus
    Else
        If grdsales.Rows >= 11 Then grdsales.TopRow = grdsales.Rows - 1
        TxtBarcode.SetFocus
    End If
    M_EDIT = False
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number = -2147206461 Then
        MsgBox Err.Description
    ElseIf Err.Number <> -2147168237 Then
        MsgBox Err.Description
        On Error Resume Next
        db.RollbackTrans
    Else
        On Error Resume Next
        db.RollbackTrans
    End If
    
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TxtLWRate.SetFocus
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim I As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHand
    db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & ""
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    With rststock
        If Not (.EOF And .BOF) Then
            .Properties("Update Criteria").value = adCriteriaKey
            !RCPT_QTY = !RCPT_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
            If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
            !RCPT_VAL = !RCPT_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
            
            !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
            rststock.Update
        End If
    End With
    db.CommitTrans
    rststock.Close
    Set rststock = Nothing
    
    I = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(LINE_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        I = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND  TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTRTRXFILE.EOF
        RSTRTRXFILE!line_no = I
        I = I + 1
        RSTRTRXFILE.Update
        RSTRTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    I = 1
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND  TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTRTRXFILE.EOF
        RSTRTRXFILE!line_no = I
        I = I + 1
        RSTRTRXFILE.Update
        RSTRTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    grdsales.Rows = 1
    I = 0
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    LBLTOTALTAX.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    grdsales.Rows = 1
    Dim GROSSVAL As Double
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTRTRXFILE.EOF
        grdsales.Rows = grdsales.Rows + 1
        grdsales.FixedRows = 1
        I = I + 1
        
        grdsales.TextMatrix(I, 0) = I
        grdsales.TextMatrix(I, 1) = RSTRTRXFILE!ITEM_CODE
        grdsales.TextMatrix(I, 2) = RSTRTRXFILE!ITEM_NAME
        grdsales.TextMatrix(I, 3) = Val(RSTRTRXFILE!QTY) / Val(RSTRTRXFILE!LINE_DISC)
        grdsales.TextMatrix(I, 4) = RSTRTRXFILE!UNIT
        grdsales.TextMatrix(I, 5) = RSTRTRXFILE!LINE_DISC
        grdsales.TextMatrix(I, 6) = Format(RSTRTRXFILE!MRP, ".000")
        grdsales.TextMatrix(I, 7) = Format(RSTRTRXFILE!SALES_PRICE, ".000")
        grdsales.TextMatrix(I, 8) = Format(RSTRTRXFILE!ITEM_COST, ".000")
        grdsales.TextMatrix(I, 9) = Format(RSTRTRXFILE!PTR, ".000")
        grdsales.TextMatrix(I, 10) = IIf(Val(RSTRTRXFILE!SALES_TAX) = 0, "", Format(RSTRTRXFILE!SALES_TAX, ".00"))
        grdsales.TextMatrix(I, 11) = RSTRTRXFILE!REF_NO
        grdsales.TextMatrix(I, 12) = Format(RSTRTRXFILE!EXP_DATE, "DD/MM/YYYY")
        grdsales.TextMatrix(I, 13) = Format(RSTRTRXFILE!TRX_TOTAL, ".000")
        grdsales.TextMatrix(I, 14) = IIf(IsNull(RSTRTRXFILE!SCHEME), "", RSTRTRXFILE!SCHEME)
        grdsales.TextMatrix(I, 15) = IIf(IsNull(RSTRTRXFILE!CHECK_FLAG), "N", RSTRTRXFILE!CHECK_FLAG)
        grdsales.TextMatrix(I, 16) = RSTRTRXFILE!line_no
        grdsales.TextMatrix(I, 17) = IIf(IsNull(RSTRTRXFILE!P_DISC), 0, RSTRTRXFILE!P_DISC)
        grdsales.TextMatrix(I, 18) = IIf(IsNull(RSTRTRXFILE!P_RETAIL), 0, RSTRTRXFILE!P_RETAIL)
        grdsales.TextMatrix(I, 19) = IIf(IsNull(RSTRTRXFILE!P_WS), 0, RSTRTRXFILE!P_WS)
        grdsales.TextMatrix(I, 20) = IIf(IsNull(RSTRTRXFILE!P_CRTN), 0, RSTRTRXFILE!P_CRTN)
        grdsales.TextMatrix(I, 37) = IIf(IsNull(RSTRTRXFILE!P_LWS), 0, RSTRTRXFILE!P_LWS)
        If RSTRTRXFILE!COM_FLAG = "A" Then
            grdsales.TextMatrix(I, 21) = 0
            grdsales.TextMatrix(I, 22) = IIf(IsNull(RSTRTRXFILE!COM_AMT), 0, RSTRTRXFILE!COM_AMT)
            grdsales.TextMatrix(I, 23) = "A"
        Else
            grdsales.TextMatrix(I, 21) = IIf(IsNull(RSTRTRXFILE!COM_PER), 0, RSTRTRXFILE!COM_PER)
            grdsales.TextMatrix(I, 22) = 0
            grdsales.TextMatrix(I, 23) = "P"
        End If
        GROSSVAL = (Val(grdsales.TextMatrix(I, 9)) * IIf(Val(grdsales.TextMatrix(I, 5)) = 0, 1, Val(grdsales.TextMatrix(I, 5)))) * (Val(grdsales.TextMatrix(I, 3)) - Val(grdsales.TextMatrix(I, 14)))
        If RSTRTRXFILE!DISC_FLAG = "P" Then
            grdsales.TextMatrix(I, 27) = "P"
            LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - (GROSSVAL * Val(grdsales.TextMatrix(I, 17)) / 100)) * Val(grdsales.TextMatrix(I, 10)) / 100, 2))
        Else
            grdsales.TextMatrix(I, 27) = "A"
            LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - Val(grdsales.TextMatrix(I, 17))) * Val(grdsales.TextMatrix(I, 10)) / 100, 2))
        End If
        grdsales.TextMatrix(I, 24) = IIf(IsNull(RSTRTRXFILE!CRTN_PACK), 0, RSTRTRXFILE!CRTN_PACK)
        grdsales.TextMatrix(I, 25) = IIf(IsNull(RSTRTRXFILE!P_VAN), 0, RSTRTRXFILE!P_VAN)
        grdsales.TextMatrix(I, 26) = IIf(IsNull(RSTRTRXFILE!GROSS_AMT), 0, RSTRTRXFILE!GROSS_AMT)
        grdsales.TextMatrix(I, 28) = IIf(IsNull(RSTRTRXFILE!LOOSE_PACK), 1, RSTRTRXFILE!LOOSE_PACK)
        grdsales.TextMatrix(I, 29) = IIf(IsNull(RSTRTRXFILE!PACK_TYPE), "Nos", RSTRTRXFILE!PACK_TYPE)
        grdsales.TextMatrix(I, 30) = IIf(IsNull(RSTRTRXFILE!WARRANTY), "", RSTRTRXFILE!WARRANTY)
        grdsales.TextMatrix(I, 31) = IIf(IsNull(RSTRTRXFILE!WARRANTY_TYPE), "", RSTRTRXFILE!WARRANTY_TYPE)
        grdsales.TextMatrix(I, 32) = IIf(IsNull(RSTRTRXFILE!EXPENSE), "", RSTRTRXFILE!EXPENSE)
        grdsales.TextMatrix(I, 33) = IIf(IsNull(RSTRTRXFILE!EXDUTY), "", RSTRTRXFILE!EXDUTY)
        grdsales.TextMatrix(I, 34) = IIf(IsNull(RSTRTRXFILE!CSTPER), "", RSTRTRXFILE!CSTPER)
        grdsales.TextMatrix(I, 35) = IIf(IsNull(RSTRTRXFILE!TR_DISC), "", RSTRTRXFILE!TR_DISC)
        grdsales.TextMatrix(I, 36) = IIf(IsNull(RSTRTRXFILE!GROSS_AMOUNT), "", RSTRTRXFILE!GROSS_AMOUNT)
        grdsales.TextMatrix(I, 38) = IIf(IsNull(RSTRTRXFILE!BARCODE), "", RSTRTRXFILE!BARCODE)
        grdsales.TextMatrix(I, 39) = IIf(IsNull(RSTRTRXFILE!CESS_AMT), "", RSTRTRXFILE!CESS_AMT)
        grdsales.TextMatrix(I, 40) = IIf(IsNull(RSTRTRXFILE!CESS_PER), "", RSTRTRXFILE!CESS_PER)
        LBLGROSSAMT.Caption = Val(LBLGROSSAMT.Caption) + Val(grdsales.TextMatrix(I, 8)) * Val(RSTRTRXFILE!QTY)
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(I, 13)), ".00")
        LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(I, 32)), ".00")
        lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(I, 3)), ".00")
        'TXTDEALER.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
        
        'TXTINVDATE.Text = Format(RSTRTRXFILE!VCH_DATE, "DD/MM/YYYY")
        'TXTREMARKS.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
        'TXTINVOICE.Text = IIf(IsNull(RSTRTRXFILE!PINV), "", RSTRTRXFILE!PINV)
        RSTRTRXFILE.MoveNext
    Loop
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    If Roundflag = True Then
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    Else
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 2), "0.00")
    End If
    LBLNET.Caption = Val(LBLGROSSAMT.Caption) + Val(LBLTOTALTAX.Caption)
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TxtBarcode.Text = ""
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.Text = ""
    TxttaxMRP.Text = ""
    TxtExDuty.Text = ""
    TxtCSTper.Text = ""
    TxtTrDisc.Text = ""
    TxtCustDisc.Text = ""
    TxtCessPer.Text = ""
    txtCess.Text = ""
    txtPD.Text = ""
    TxtExpense.Text = ""
    txtprofit.Text = ""
    TXTRETAIL.Text = ""
    TxtRetailPercent.Text = ""
    
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    TxtLWRate.Text = ""
    txtcrtnpack.Text = ""
    TXTRATE.Text = ""
    TxtComAmt.Text = ""
    TxtComper.Text = ""
    txtmrpbt.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    txtBatch.Text = ""
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    lblcategory.Caption = ""
    LBLPRE.Caption = ""
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_ADD = True
    OLD_BILL = True
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number <> -2147168237 Then
        MsgBox Err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub CmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            Txtpack.Text = 1 '""
            Los_Pack.Text = ""
            CmbPack.ListIndex = -1
            TxtWarranty.Text = ""
            CmbWrnty.ListIndex = -1
            TxTfree.Text = ""
            TxttaxMRP.Text = ""
            TxtExDuty.Text = ""
            TxtCSTper.Text = ""
            TxtTrDisc.Text = ""
            TxtCustDisc.Text = ""
            TxtCessPer.Text = ""
            txtCess.Text = ""
            txtPD.Text = ""
            TxtExpense.Text = ""
            txtprofit.Text = ""
            TXTRETAIL.Text = ""
            TxtRetailPercent.Text = ""
            
            txtWsalePercent.Text = ""
            txtSchPercent.Text = ""
            txtWS.Text = ""
            txtvanrate.Text = ""
            Txtgrossamt.Text = ""
            txtcrtn.Text = ""
            TxtLWRate.Text = ""
            txtcrtnpack.Text = ""
            TXTRATE.Text = ""
            TxtComAmt.Text = ""
            TxtComper.Text = ""
            txtmrpbt.Text = ""
            TXTITEMCODE.Text = ""
            TxtBarcode.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            lbltaxamount.Caption = ""
            lblcategory.Caption = ""
            LBLPRE.Caption = ""
            TXTEXPDATE.Text = "  /  /    "
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
        
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
            M_EDIT = False
    End Select
End Sub

Private Sub CmdDeleteAll_Click()
    Dim I As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    On Error GoTo ErrHand
    If Chkcancel.value = 0 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE ALL", vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then Exit Sub
    
    TXTDEALER.Text = ""
    DataList2.BoundText = ""
    TXTINVOICE.Text = ""
    CMBDISTRICT.Text = ""
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    TXTRCVDATE.Text = "  /  /    "
    TXTDATE.Text = Format(Date, "DD/MM/YYYY")
    lbladdress.Caption = ""
    TXTREMARKS.Text = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    LBLTOTAL.Caption = ""
    TXTDISCAMOUNT.Text = ""
    TxtTotalexp.Text = ""
    txtaddlamt.Text = ""
        
    db.Execute "delete  From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " "
    db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'PY' AND INV_TRX_TYPE = 'PI' "
    db.Execute "delete FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'CR' AND INV_TRX_TYPE = 'PI'"
    'db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'PI'"
    For I = 1 To grdsales.Rows - 1
        db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(I, 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(I, 16)) & ""
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(I, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        With rststock
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").value = adCriteriaKey
                !RCPT_QTY = !RCPT_QTY - Val(grdsales.TextMatrix(I, 3)) * Val(grdsales.TextMatrix(I, 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                !RCPT_VAL = !RCPT_VAL - Val(grdsales.TextMatrix(I, 13))
                
                !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(I, 3)) * Val(grdsales.TextMatrix(I, 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(I, 13))
                rststock.Update
            End If
        End With
        db.CommitTrans
        rststock.Close
        Set rststock = Nothing
    Next I
    
    grdsales.FixedRows = 0
    grdsales.Rows = 1
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    cmdRefresh.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.Text = TXTLASTBILL.Text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTRCVDATE.Text = "  /  /    "
    TXTINVOICE.Text = ""
    CMBDISTRICT.Text = ""
    lbladdress.Caption = ""
    TXTREMARKS.Text = ""
    TXTSLNO.Text = ""
    TXTITEMCODE.Text = ""
    TxtBarcode.Text = ""
    TXTPRODUCT.Text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.Text = ""
    TxttaxMRP.Text = ""
    TxtExDuty.Text = ""
    TxtCSTper.Text = ""
    TxtTrDisc.Text = ""
    TxtCustDisc.Text = ""
    TxtCessPer.Text = ""
    txtCess.Text = ""
    txtPD.Text = ""
    TxtExpense.Text = ""
    txtprofit.Text = ""
    TXTRETAIL.Text = ""
    TxtRetailPercent.Text = ""
    
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    TxtLWRate.Text = ""
    txtcrtnpack.Text = ""
    txtBatch.Text = ""
    txtHSN.Text = ""
    TXTRATE.Text = ""
    txtmrpbt.Text = ""
    TXTPTR.Text = ""
    TxtNetrate.Text = ""
    Txtgrossamt.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    lblcategory.Caption = ""
    LBLPRE.Caption = ""
    txtaddlamt.Text = ""
    txtcramt.Text = ""
    TxtInsurance.Text = ""
    TxtCST.Text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    TXTDISCAMOUNT.Text = ""
    TxtTotalexp.Text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    grdsales.Rows = 1
    CMDEXIT.Enabled = True
    OptComper.value = True
    txtBillNo.SetFocus
    M_ADD = False
    M_EDIT = False
    OLD_BILL = False
    LBLmonth.Caption = "0.00"
    Chkcancel.value = 0
    Call CLEAR_COMBO
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "EzBiz"
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number <> -2147168237 Then
        MsgBox Err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdLabels_Click()
    Dim N, sl, M As Long
    If grdsales.Rows <= 1 Then Exit Sub
    'If grdsales.Cols = 20 Then Exit Sub
    
    On Error GoTo ErrHand
    db.Execute "Delete from barprint"
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From barprint", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    sl = Val(InputBox("Enter the Serial No. from which to be Print", "Label Printing", 1))
    If sl = 0 Then Exit Sub
    For N = sl To grdsales.Rows - 1
        For M = 1 To Val(grdsales.TextMatrix(N, 41))
            RSTTRXFILE.AddNew
            RSTTRXFILE!BARCODE = "*" & grdsales.TextMatrix(N, 38) & "*"
            RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(N, 2))
            RSTTRXFILE!item_Price = Val(grdsales.TextMatrix(N, 18))
            RSTTRXFILE!item_MRP = Val(grdsales.TextMatrix(N, 6))
            RSTTRXFILE!item_color = Trim(grdsales.TextMatrix(N, 11))
            If IsDate(grdsales.TextMatrix(N, 12)) Then
                RSTTRXFILE!expdate = Format(grdsales.TextMatrix(N, 12), "dd/mm/yyyy")
                If IsDate(TXTINVDATE.Text) Then
                    RSTTRXFILE!pckdate = Format(TXTINVDATE.Text, "dd/mm/yyyy")
                End If
            End If
            RSTTRXFILE!Remarks = ""
            RSTTRXFILE!COMP_NAME = Trim(MDIMAIN.StatusBar.Panels(5).Text)
            RSTTRXFILE.Update
        Next M
'            Select Case (MsgBox("Do you want to print Label for " & grdsales.TextMatrix(N, 2), vbYesNoCancel, "Label Printing!!!"))
'                Case vbYes
'                    'grdsales.TextMatrix(N, 5)
''                    Picture5.Tag = ""
''                    Picture5.Cls
''                    Picture5.Picture = Nothing
''                    Picture5.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
''                    Picture5.CurrentY = 0 'Y2 + 0.25 * Th
''                    Picture5.Print Picture5.Tag & " " & Picture4.Tag
'
'                    Picture5.Cls
'                    Picture5.Picture = Nothing
'                    Picture5.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                    Picture5.CurrentY = 0 'Y2 + 0.25 * Th
'                    Picture5.Print "PRICE: " & Format(grdsales.TextMatrix(N, 5), "0.00")
'
'                    Picture6.Cls
'                    Picture6.Picture = Nothing
'                    Picture6.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                    Picture6.CurrentY = 0 'Y2 + 0.25 * Th
'                    Picture6.Print "MRP  : " & Format(grdsales.TextMatrix(N, 7), "0.00")
'
'                    Picture1.Cls
'                    Picture1.Picture = Nothing
'                    Picture1.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                    Picture1.CurrentY = 0 'Y2 + 0.25 * Th
'                    Picture1.Print Mid(Trim(grdsales.TextMatrix(N, 2)), 1, 11) & " MRP: " & Format(grdsales.TextMatrix(N, 7), "0.00")
'
'                    Dim i As Long
'                    i = Val(InputBox("Enter number of lables to be print", "No. of labels..", grdsales.TextMatrix(N, 41)))
'                    'i = Val(grdsales.TextMatrix(N, 41))
'                    If i <= 0 Then Exit Sub
'                    If MDIMAIN.barcode_profile.Caption = 0 Then
'                        If i > 0 Then Call print_3labels(i, Trim(grdsales.TextMatrix(N, 38)), Trim(grdsales.TextMatrix(N, 2)), Val(grdsales.TextMatrix(N, 6)), Val(grdsales.TextMatrix(N, 18)))
'                        'grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
'                        '(i As Long, BAR_LABEL As String, itemname As String, itemmrp As Double, itemprice As Double)
'                    Else
'                        If i > 0 Then Call print_labels(i, Trim(grdsales.TextMatrix(N, 38)), Trim(grdsales.TextMatrix(N, 2)), Val(grdsales.TextMatrix(N, 6)), Val(grdsales.TextMatrix(N, 18)))
'                        'If i > 0 Then Call print_labels(i, Trim(txtBarcode.Text), "")
'                    End If
'                    'Call print_labels(Val(grdsales.TextMatrix(N, 3)))
'                Case vbCancel
'                    Exit For
'                Case vbNo
'
'            End Select
        Next N
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        db.CommitTrans
        
  
    Dim I As Long
     If BARPRINTER = barcodeprinter Then
        ReportNameVar = Rptpath & "Rptbarprn"
    Else
        ReportNameVar = Rptpath & "Rptbarprn1"
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields

    For N = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(N).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(N).Name & " ")
            Report.Database.SetDataSource oRs, 3, N
            Set oRs = Nothing
        End If
    Next N
    Set Printer = Printers(BARPRINTER)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    Report.PrintOut (False)
    Set CRXFormulaFields = Nothing
    Set crxApplication = Nothing
    Set Report = Nothing
            
'    For i = 1 To Report.Database.Tables.COUNT
'        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
'    Next i
'    Report.DiscardSavedData
'    frmreport.Caption = "BARCODE"
'    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
        
Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number = -2147206461 Then
        MsgBox Err.Description
    ElseIf Err.Number <> -2147168237 Then
        MsgBox Err.Description
        On Error Resume Next
        db.RollbackTrans
    Else
        On Error Resume Next
        db.RollbackTrans
    End If
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
'    Los_Pack.Enabled = True
'    CmbPack.Enabled = True
'    TXTQTY.Enabled = True
'    TxtFree.Enabled = True
'    TXTRATE.Enabled = True
'    TXTPTR.Enabled = True
'    TxttaxMRP.Enabled = True
'    TxtExDuty.Enabled = True
'    TxtTrDisc.Enabled = True
'    TxtCessPer.Enabled = True
'    txtCess.Enabled = True
'    TxtCSTper.Enabled = True
'    txtPD.Enabled = True
'    TxtExpense.Enabled = True
'    TXTRETAIL.Enabled = True
'    TxtRetailPercent.Enabled = True
'    txtWS.Enabled = True
'    txtWsalePercent.Enabled = True
'    txtvanrate.Enabled = True
'    txtSchPercent.Enabled = True
'    txtcrtnpack.Enabled = True
'    txtcrtn.Enabled = True
'    TxtLWRate.Enabled = True
'    TxtComper.Enabled = True
'    TxtComAmt.Enabled = True
'    cmdadd.Enabled = True
'    Txtgrossamt.Enabled = True
'    txtBatch.Enabled = True
'    txtHSN.Enabled = True
'    TxtWarranty.Enabled = True
'    CmbWrnty.Enabled = True
'    TXTEXPIRY.Visible = False
'    TXTEXPDATE.Enabled = True
'    TxtBarcode.Enabled = False
    If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
        Los_Pack.Text = 1
        TXTQTY.Text = 1
        TxTfree.Text = ""
        TXTRATE.Text = ""
        TXTPTR.Enabled = True
        TXTPTR.SetFocus
    Else
        Los_Pack.Enabled = True
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
    End If
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            Txtpack.Text = 1 '""
            Los_Pack.Text = ""
            CmbPack.ListIndex = -1
            TxtWarranty.Text = ""
            CmbWrnty.ListIndex = -1
            TxTfree.Text = ""
            TxttaxMRP.Text = ""
            TxtExDuty.Text = ""
            TxtCSTper.Text = ""
            TxtTrDisc.Text = ""
            TxtCustDisc.Text = ""
            TxtCessPer.Text = ""
            txtCess.Text = ""
            txtPD.Text = ""
            TxtExpense.Text = ""
            txtprofit.Text = ""
            TXTRETAIL.Text = ""
            TxtRetailPercent.Text = ""
            
            txtWsalePercent.Text = ""
            txtSchPercent.Text = ""
            txtWS.Text = ""
            txtvanrate.Text = ""
            Txtgrossamt.Text = ""
            txtcrtn.Text = ""
            TxtLWRate.Text = ""
            txtcrtnpack.Text = ""
            TXTRATE.Text = ""
            TxtComAmt.Text = ""
            TxtComper.Text = ""
            txtmrpbt.Text = ""
            TXTITEMCODE.Text = ""
            TxtBarcode.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            lbltaxamount.Caption = ""
            lblcategory.Caption = ""
            LBLPRE.Caption = ""
            TXTEXPDATE.Text = "  /  /    "
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
        
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
            M_EDIT = False
    End Select
End Sub

Private Sub CmdPrint_Click()
    Dim I As Long
    
    On Error GoTo ErrHand
     
    Screen.MousePointer = vbHourglass
    db.Execute "delete from TEMPTRXFILE"
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    For I = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "LP"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!line_no = I
        'RSTTRXFILE!Category = "" 'grdsales.TextMatrix(i, 25)
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(I, 1)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(I, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(I, 3))
        
        
        RSTTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(I, 13))
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "dd/mm/yyyy")
        RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(I, 2))
        RSTTRXFILE!ITEM_COST = Format(Round(Val(grdsales.TextMatrix(I, 9)) * Val(grdsales.TextMatrix(I, 5)), 4), "0.0000") 'Val(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(I, 17))
        If Trim(grdsales.TextMatrix(I, 27)) = "P" Then
            RSTTRXFILE!Category = "P"
        ElseIf Trim(grdsales.TextMatrix(I, 27)) = "A" Then
            RSTTRXFILE!Category = "A"
        End If
                
        
        'RSTTRXFILE!P_DISC = Val(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(I, 6))
        RSTTRXFILE!PTR = Format(Round(Val(grdsales.TextMatrix(I, 9)) * Val(grdsales.TextMatrix(I, 5)), 4), "0.0000")
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(I, 7))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(I, 18))
        RSTTRXFILE!P_RETAILWOTAX = Round(Val(grdsales.TextMatrix(I, 9)) * Val(grdsales.TextMatrix(I, 5)), 4)
        'RSTTRXFILE!P_WS = Val(grdsales.TextMatrix(i, 19))
        'RSTTRXFILE!P_CRTN = Val(grdsales.TextMatrix(i, 20))
        'RSTTRXFILE!CRTN_PACK = Val(grdsales.TextMatrix(i, 24))
        'RSTTRXFILE!P_VAN = Val(grdsales.TextMatrix(i, 25))
        'RSTTRXFILE!GROSS_AMT = Val(grdsales.TextMatrix(i, 26))
        RSTTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(I, 10))
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(I, 28))
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(I, 29)
        RSTTRXFILE!WARRANTY = Val(grdsales.TextMatrix(I, 30))
        RSTTRXFILE!WARRANTY_TYPE = Trim(grdsales.TextMatrix(I, 31))
        RSTTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(I, 4))
        If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
            RSTTRXFILE!VCH_DESC = "Y"
        Else
            RSTTRXFILE!VCH_DESC = "N"
        End If
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(I, 11))
        'RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!SCHEME = Val(grdsales.TextMatrix(I, 14))
        If IsDate(grdsales.TextMatrix(I, 12)) Then
            RSTTRXFILE!EXP_DATE = Format(grdsales.TextMatrix(I, 12), "dd/mm/yyyy")
        End If
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(I, 15))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
        RSTTRXFILE.Update
    Next I
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompAddress4, CompAddress5, CompTin, CompCST, DL, ML, DL1, DL2, INV_TERMS, BANK_DET, PAN_NO As String
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
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
        INV_TERMS = IIf(IsNull(RSTCOMPANY!INV_TERMS) Or RSTCOMPANY!INV_TERMS = "", "", RSTCOMPANY!INV_TERMS)
        BANK_DET = IIf(IsNull(RSTCOMPANY!bank_details) Or RSTCOMPANY!bank_details = "", "", RSTCOMPANY!bank_details)
        PAN_NO = IIf(IsNull(RSTCOMPANY!PAN_NO) Or RSTCOMPANY!PAN_NO = "", "", RSTCOMPANY!PAN_NO)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
              
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        DL1 = IIf(IsNull(RSTCOMPANY!DL_NO), "", Trim(RSTCOMPANY!DL_NO))
        DL2 = IIf(IsNull(RSTCOMPANY!Remarks), "", Trim(RSTCOMPANY!Remarks))
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
            
    ReportNameVar = Rptpath & "rptLP"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields
    
    For I = 1 To Report.OpenSubreport("RPTBILL1.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTBILL1.rpt").Database.Tables(I).SetLogOnInfo strConnection
    Next I
    For I = 1 To Report.OpenSubreport("RPTBILL2.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTBILL2.rpt").Database.Tables(I).SetLogOnInfo strConnection
    Next I
    For I = 1 To Report.OpenSubreport("RPTBILL3.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTBILL3.rpt").Database.Tables(I).SetLogOnInfo strConnection
    Next I
    For I = 1 To 3
        'Set CRXFormulaFields = Report.FormulaFields
        Report.OpenSubreport("RPTBILL" & I & ".rpt").RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.Text) & ")"
        Report.OpenSubreport("RPTBILL" & I & ".rpt").DiscardSavedData
        Report.OpenSubreport("RPTBILL" & I & ".rpt").VerifyOnEveryPrint = True
        Set CRXFormulaFields = Report.OpenSubreport("RPTBILL" & I & ".rpt").FormulaFields
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
            If CRXFormulaField.Name = "{@DL1}" Then CRXFormulaField.Text = "'" & DL1 & "'"
            If CRXFormulaField.Name = "{@DL2}" Then CRXFormulaField.Text = "'" & DL2 & "'"
            If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.Text = "'" & INV_TERMS & "'"
            If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.Text = "'" & BANK_DET & "'"
            If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.Text = "'" & PAN_NO & "'"
            If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & Trim(TXTDEALER.Text) & "'"
            If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.Text = "'" & Trim(TXTDEALER.Text) & "'"
'            If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.Text = "'" & Trim(lbladdress.Caption) & "'"
            If CRXFormulaField.Name = "{DLNO2}" Then CRXFormulaField.Text = "'" & DL1 & "'"
            If CRXFormulaField.Name = "{DLNO}" Then CRXFormulaField.Text = "'" & DL2 & "'"
            'If CRXFormulaField.Name = "{@Area}" Then CRXFormulaField.Text = "'" & Trim(TXTAREA.Text) & "'"
            'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Tag), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
            If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
    '        If Tax_Print = False Then
    '            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
    '        End If
            'If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
            If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & TXTINVOICE.Text & "'"
            If CRXFormulaField.Name = "{@VCH_NO}" Then CRXFormulaField.Text = "'" & Format(Trim(txtBillNo.Text), "0000") & "' "
            'If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
            'If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.Text = "'" & Trim(TxtOrder.Text) & "'"
    '            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLTOTAL.Caption), 0), "0.00") & "'"
        Next
    Next I
        
    'Preview
    frmreport.Caption = "BILL"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number <> -2147168237 Then
        MsgBox Err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub cmdRefresh_Click()
    If CMBPO.VisibleCount = 0 Then CMBPO.Text = ""
    If CMBPO.VisibleCount <> 0 And CMBPO.BoundText = "" Then
        If (MsgBox("Are you sure you want to save the Purchase Bill without selecting the Purchase Order No.? !!!!", vbYesNo, "EzBiz") = vbNo) Then Exit Sub
    End If
    If CMBPO.Text <> "" And CMBPO.MatchedWithList = False Then
        MsgBox "Please select a valid PO No. from the list", vbOKOnly, "EzBiz"
        On Error Resume Next
        CMBPO.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(TXTRCVDATE.Text) Then
        TXTRCVDATE.Text = TXTINVDATE.Text
    End If
    If (DateValue(TXTINVDATE.Text) < DateValue(MDIMAIN.DTFROM.value)) Or (DateValue(TXTINVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.value))) Then
        'db.Execute "delete from Users"
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
'    If (DateValue(TXTRCVDATE.Text) < DateValue(MDIMAIN.DTFROM.value)) Or (DateValue(TXTRCVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.value))) Then
'        'db.Execute "delete from Users"
'        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
'        TXTINVDATE.SetFocus
'        Exit Sub
'    End If
    
    If DateValue(TXTRCVDATE.Text) < DateValue(TXTINVDATE.Text) Then
        'db.Execute "delete from Users"
        MsgBox "Goods received date could not be less than Invice date", vbOKOnly, "EzBiz"
        TXTRCVDATE.SetFocus
        Exit Sub
    End If
    
    If DateValue(TXTRCVDATE.Text) <> DateValue(TXTINVDATE.Text) Then
        If (MsgBox("Invice date & Rcvd date are different. Are you sure?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo) Then
            TXTRCVDATE.SetFocus
            Exit Sub
        End If
    End If
    
    BARCODE_FLAG = False
    On Error GoTo ErrHand
    If grdsales.Rows <= 1 Then
        lblcredit.Caption = "0"
        Call appendpurchase
    Else
        If IsNull(DataList2.SelectedItem) Then
            MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
            DataList2.SetFocus
            Exit Sub
        End If
        If TXTINVOICE.Text = "" Then
            MsgBox "Enter Supplier Invoice No.", vbOKOnly, "EzBiz"
            Exit Sub
        End If
        If Not IsDate(TXTINVDATE.Text) Then
            MsgBox "Enter Supplier Invoice Date", vbOKOnly, "EzBiz"
            Exit Sub
        End If
        'Me.Enabled = False
        'MDIMAIN.cmdpurchase.Enabled = False
        'Set creditbill = Me
        'frmCREDIT.Show
        Call appendpurchase
    End If
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub cmdRefresh_GotFocus()
    FRMEGRDTMP.Visible = False
End Sub


Private Sub Command4_Click()
    If CMDEXIT.Enabled = False Then Exit Sub
    If Val(txtBillNo.Text) = 1 Then Exit Sub
    txtBillNo.Text = Val(txtBillNo.Text) - 1
    
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    cmdRefresh.Enabled = False
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTRCVDATE.Text = "  /  /    "
    TXTINVOICE.Text = ""
    CMBDISTRICT.Text = ""
    lbladdress.Caption = ""
    TXTREMARKS.Text = ""
    TXTSLNO.Text = ""
    TXTITEMCODE.Text = ""
    TxtBarcode.Text = ""
    TXTPRODUCT.Text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.Text = ""
    TxttaxMRP.Text = ""
    TxtExDuty.Text = ""
    TxtCSTper.Text = ""
    TxtTrDisc.Text = ""
    TxtCustDisc.Text = ""
    TxtCessPer.Text = ""
    txtCess.Text = ""
    txtPD.Text = ""
    TxtExpense.Text = ""
    txtprofit.Text = ""
    TXTRETAIL.Text = ""
    TxtRetailPercent.Text = ""
    
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    TxtLWRate.Text = ""
    txtcrtnpack.Text = ""
    txtBatch.Text = ""
    txtHSN.Text = ""
    TXTRATE.Text = ""
    txtmrpbt.Text = ""
    TXTPTR.Text = ""
    TxtNetrate.Text = ""
    
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    lblcategory.Caption = ""
    LBLPRE.Caption = ""
    txtaddlamt.Text = ""
    txtcramt.Text = ""
    TxtInsurance.Text = ""
    TxtCST.Text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    TXTDISCAMOUNT.Text = ""
    TxtTotalexp.Text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    grdsales.Rows = 1
    CMDEXIT.Enabled = True
    OptComper.value = True
    M_ADD = False
    OLD_BILL = False
    LBLmonth.Caption = "0.00"
    
    Chkcancel.value = 0
    Call TXTBILLNO_KeyDown(13, 0)
End Sub

Private Sub Command5_Click()
    If CMDEXIT.Enabled = False Then Exit Sub
    Dim rstBILL As ADODB.Recordset
    Dim lastbillno As Double
    On Error GoTo ErrHand
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        lastbillno = IIf(IsNull(rstBILL.Fields(0)), 0, rstBILL.Fields(0))
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    If Val(txtBillNo.Text) > lastbillno Then Exit Sub
    txtBillNo.Text = Val(txtBillNo.Text) + 1
    
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    cmdRefresh.Enabled = False
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTRCVDATE.Text = "  /  /    "
    TXTINVOICE.Text = ""
    CMBDISTRICT.Text = ""
    lbladdress.Caption = ""
    TXTREMARKS.Text = ""
    TXTSLNO.Text = ""
    TXTITEMCODE.Text = ""
    TxtBarcode.Text = ""
    TXTPRODUCT.Text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.Text = ""
    TxttaxMRP.Text = ""
    TxtExDuty.Text = ""
    TxtCSTper.Text = ""
    TxtTrDisc.Text = ""
    TxtCustDisc.Text = ""
    TxtCessPer.Text = ""
    txtCess.Text = ""
    txtPD.Text = ""
    TxtExpense.Text = ""
    txtprofit.Text = ""
    TXTRETAIL.Text = ""
    TxtRetailPercent.Text = ""
    
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    TxtLWRate.Text = ""
    txtcrtnpack.Text = ""
    txtBatch.Text = ""
    txtHSN.Text = ""
    TXTRATE.Text = ""
    txtmrpbt.Text = ""
    TXTPTR.Text = ""
    TxtNetrate.Text = ""
    
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    lblcategory.Caption = ""
    LBLPRE.Caption = ""
    txtaddlamt.Text = ""
    txtcramt.Text = ""
    TxtInsurance.Text = ""
    TxtCST.Text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    TXTDISCAMOUNT.Text = ""
    TxtTotalexp.Text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    grdsales.Rows = 1
    CMDEXIT.Enabled = True
    OptComper.value = True
    M_ADD = False
    OLD_BILL = False
    LBLmonth.Caption = "0.00"
    
    Chkcancel.value = 0
    Call TXTBILLNO_KeyDown(13, 0)
    Exit Sub
ErrHand:
    MsgBox Err.Description, , "EzBiz"
End Sub

Private Sub Form_Activate()
    'On Error GoTo ErrHand
    On Error Resume Next
    If txtBillNo.Enabled = True Then txtBillNo.SetFocus
    'if TXTDEALER.Enabled=True then TXTDEALER.SetFocus
'    Exit Sub
'ErrHand:
'    If Err.Number = 5 Then Exit Sub
'    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
        
    Dim P
    For Each P In Printers
        Cmbbarcode.AddItem (P.DeviceName)
    Next P
    
    Dim ObjFso
    Dim StrFileName
    Dim ObjFile
    If FileExists(App.Path & "\BillPrint") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\BillPrint")  'Reading from the file
        On Error Resume Next
        ObjFile.ReadLine
        ObjFile.ReadLine
        ObjFile.ReadLine
        Cmbbarcode.ListIndex = ObjFile.ReadLine
        Err.Clear
        On Error GoTo ErrHand
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    BARPRINTER = barcodeprinter
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    Call CLEAR_COMBO
    
    
    M_EDIT = False
    ACT_FLAG = True
    PO_FLAG = True
    PRERATE_FLAG = True
    OLD_BILL = False
    grdsales.ColWidth(0) = 500
    grdsales.ColWidth(1) = 1000
    grdsales.ColWidth(2) = 4000
    grdsales.ColWidth(3) = 1000
    grdsales.ColWidth(4) = 0 ' 800
    grdsales.ColWidth(5) = 0 '800
    grdsales.ColWidth(6) = 1200
    grdsales.ColWidth(7) = 0 '800
    grdsales.ColWidth(8) = 800
    grdsales.ColWidth(9) = 800
    grdsales.ColWidth(10) = 1000
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 0 '1100
    grdsales.ColWidth(13) = 1700 '1100
    grdsales.ColWidth(16) = 0
    grdsales.ColWidth(14) = 800
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(17) = 800
    grdsales.ColWidth(18) = 800
    grdsales.ColWidth(19) = 800
    grdsales.ColWidth(20) = 800
    grdsales.ColWidth(37) = 800
    grdsales.ColWidth(21) = 0
    grdsales.ColWidth(22) = 0
    grdsales.ColWidth(23) = 0
    grdsales.ColWidth(24) = 700
    grdsales.ColWidth(25) = 0
    grdsales.ColWidth(26) = 1100
    grdsales.ColWidth(27) = 0
    grdsales.ColWidth(28) = 1100
    grdsales.ColWidth(29) = 0
    grdsales.ColWidth(30) = 0
    grdsales.ColWidth(31) = 0
    
    grdsales.ColAlignment(0) = 4
    grdsales.ColAlignment(1) = 1
    grdsales.ColAlignment(2) = 1
    grdsales.ColAlignment(3) = 4
    grdsales.ColAlignment(4) = 4
    grdsales.ColAlignment(9) = 4
    grdsales.ColAlignment(10) = 4
    grdsales.ColAlignment(5) = 7
    grdsales.ColAlignment(6) = 7
    grdsales.ColAlignment(7) = 7
    grdsales.ColAlignment(8) = 7
    grdsales.ColAlignment(11) = 7
    grdsales.ColAlignment(17) = 7
    grdsales.ColAlignment(18) = 7
    grdsales.ColAlignment(19) = 7
    grdsales.ColAlignment(20) = 7
    grdsales.ColAlignment(37) = 7
    grdsales.ColAlignment(21) = 7
    grdsales.ColAlignment(22) = 7
    grdsales.ColAlignment(26) = 7
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "TOTAL QTY"
    grdsales.TextArray(4) = "UNIT"
    grdsales.TextArray(5) = "" '"PACK"
    grdsales.TextArray(6) = "MRP"
    grdsales.TextArray(7) = "PTS"
    grdsales.TextArray(8) = "COST"
    grdsales.TextArray(9) = "RATE"
    grdsales.TextArray(10) = "TAX %"
    grdsales.TextArray(11) = "SERIAL NO"
    grdsales.TextArray(12) = "EXPIRY"
    grdsales.TextArray(13) = "SUB TOTAL"
    grdsales.TextArray(14) = "FREE"
    grdsales.TextArray(15) = "TAX MODE"
    grdsales.TextArray(16) = "Line No"
    grdsales.TextArray(17) = "Disc"
    grdsales.TextArray(18) = "RT Price"
    grdsales.TextArray(19) = "WS Price"
    grdsales.TextArray(20) = "L. R.Price"
    grdsales.TextArray(37) = "L. W.Price"
    grdsales.TextArray(21) = "Comm %"
    grdsales.TextArray(22) = "Comm Amt"
    grdsales.TextArray(23) = "Comm Flag"
    grdsales.TextArray(24) = "L. Pck"
    grdsales.TextArray(25) = "Van Rate"
    grdsales.TextArray(26) = "GROSS AMOUNT"
    grdsales.TextArray(27) = "DISC_FLAG"
    grdsales.TextArray(28) = "PACK"
    grdsales.TextArray(32) = "Expense"
    grdsales.TextArray(33) = "Ex. Duty %"
    grdsales.TextArray(34) = "CST %"
    grdsales.TextArray(35) = "Trade Disc"
    grdsales.TextArray(36) = "Gross"
    grdsales.TextArray(38) = "Barcode"
    grdsales.TextArray(39) = "Cess Rate"
    grdsales.TextArray(40) = "Cess %"
    grdsales.TextArray(41) = "Labels"
    
    PHYFLAG = True
    PHYCODE_FLAG = True
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    'TXTDATE.Text = Date
    TXTEXPDATE.Enabled = False
    txtBatch.Enabled = False
    txtHSN.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    TXTUNIT.Enabled = False
    TXTSLNO.Text = 1
    TXTSLNO.Enabled = True
    FRMECONTROLS.Enabled = False
    FRMEMASTER.Enabled = False
    CLOSEALL = 1
    lblcredit.Caption = "1"
    TXTDEALER.Text = ""
    M_ADD = False
    
    'Me.Width = 15135
    'Me.Height = 9660
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If PHYCODE_FLAG = False Then PHY_CODE.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If PO_FLAG = False Then ACT_PO.Close
        If PRERATE_FLAG = False Then PHY_PRERATE.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdsales_Click()
    TXTsample.Visible = False
End Sub

Private Sub grdsales_DblClick()
    If grdsales.Rows <= 1 Then Exit Sub
    If M_EDIT = True Then Exit Sub
    TXTSLNO.Text = grdsales.TextMatrix(grdsales.Row, 0)
    Call TXTSLNO_KeyDown(13, 0)
    CMDMODIFY_Click
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 41
                        If grdsales.Cols = 20 Then Exit Sub
                        TXTsample.MaxLength = 3
                        TXTsample.Visible = True
                        TXTsample.Top = grdsales.CellTop + 100
                        TXTsample.Left = grdsales.CellLeft '+ 50
                        TXTsample.Width = grdsales.CellWidth
                        TXTsample.Height = grdsales.CellHeight
                        TXTsample.Text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                        TXTsample.SetFocus
            End Select
        Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            If txtBillNo.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then Exit Sub
            If Not IsDate(TXTINVDATE.Text) Then Exit Sub
            If TXTQTY.Enabled = True Then Exit Sub
            If Los_Pack.Enabled = True Then Exit Sub
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyDelete
            If grdsales.Rows <= 1 Then Exit Sub
            If M_EDIT = True Then Exit Sub
            If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(grdsales.Row, 2) & """", vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then
                grdsales.SetFocus
                Exit Sub
            End If
            
            Dim I As Long
            Dim rststock As ADODB.Recordset
            Dim RSTRTRXFILE As ADODB.Recordset
            Dim rstMaxNo As ADODB.Recordset
            
            Screen.MousePointer = vbHourglass
            On Error GoTo ErrHand
            db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(grdsales.Row, 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(grdsales.Row, 16)) & ""
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            db.BeginTrans
            With rststock
                If Not (.EOF And .BOF) Then
                    .Properties("Update Criteria").value = adCriteriaKey
                    !RCPT_QTY = !RCPT_QTY - Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 5))
                    If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                    !RCPT_VAL = !RCPT_VAL - Val(grdsales.TextMatrix(grdsales.Row, 13))
                    
                    !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 5))
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(grdsales.Row, 13))
                    rststock.Update
                End If
            End With
            db.CommitTrans
            rststock.Close
            Set rststock = Nothing
            
            I = 0
            Set rstMaxNo = New ADODB.Recordset
            rstMaxNo.Open "Select MAX(LINE_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockReadOnly
            If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
                I = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
            End If
            rstMaxNo.Close
            Set rstMaxNo = Nothing
            
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND  TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
            db.BeginTrans
            Do Until RSTRTRXFILE.EOF
                RSTRTRXFILE!line_no = I
                I = I + 1
                RSTRTRXFILE.Update
                RSTRTRXFILE.MoveNext
            Loop
            db.CommitTrans
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            I = 1
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND  TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
            db.BeginTrans
            Do Until RSTRTRXFILE.EOF
                RSTRTRXFILE!line_no = I
                I = I + 1
                RSTRTRXFILE.Update
                RSTRTRXFILE.MoveNext
            Loop
            db.CommitTrans
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            grdsales.Rows = 1
            I = 0
            LBLTOTAL.Caption = ""
            lbltotalwodiscount = ""
            LBLTOTALTAX.Caption = ""
            LBLGROSSAMT.Caption = ""
            LBLEXP.Caption = ""
            lblqty.Caption = ""
            grdsales.Rows = 1
            Dim GROSSVAL As Double
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE.EOF
                grdsales.Rows = grdsales.Rows + 1
                grdsales.FixedRows = 1
                I = I + 1
                
                grdsales.TextMatrix(I, 0) = I
                grdsales.TextMatrix(I, 1) = RSTRTRXFILE!ITEM_CODE
                grdsales.TextMatrix(I, 2) = RSTRTRXFILE!ITEM_NAME
                grdsales.TextMatrix(I, 3) = Val(RSTRTRXFILE!QTY) / Val(RSTRTRXFILE!LINE_DISC)
                grdsales.TextMatrix(I, 4) = RSTRTRXFILE!UNIT
                grdsales.TextMatrix(I, 5) = RSTRTRXFILE!LINE_DISC
                grdsales.TextMatrix(I, 6) = Format(RSTRTRXFILE!MRP, ".000")
                grdsales.TextMatrix(I, 7) = Format(RSTRTRXFILE!SALES_PRICE, ".000")
                grdsales.TextMatrix(I, 8) = Format(RSTRTRXFILE!ITEM_COST, ".000")
                grdsales.TextMatrix(I, 9) = Format(RSTRTRXFILE!PTR, ".000")
                grdsales.TextMatrix(I, 10) = IIf(Val(RSTRTRXFILE!SALES_TAX) = 0, "", Format(RSTRTRXFILE!SALES_TAX, ".00"))
                grdsales.TextMatrix(I, 11) = RSTRTRXFILE!REF_NO
                grdsales.TextMatrix(I, 12) = Format(RSTRTRXFILE!EXP_DATE, "DD/MM/YYYY")
                grdsales.TextMatrix(I, 13) = Format(RSTRTRXFILE!TRX_TOTAL, ".000")
                grdsales.TextMatrix(I, 14) = IIf(IsNull(RSTRTRXFILE!SCHEME), "", RSTRTRXFILE!SCHEME)
                grdsales.TextMatrix(I, 15) = IIf(IsNull(RSTRTRXFILE!CHECK_FLAG), "N", RSTRTRXFILE!CHECK_FLAG)
                grdsales.TextMatrix(I, 16) = RSTRTRXFILE!line_no
                grdsales.TextMatrix(I, 17) = IIf(IsNull(RSTRTRXFILE!P_DISC), 0, RSTRTRXFILE!P_DISC)
                grdsales.TextMatrix(I, 18) = IIf(IsNull(RSTRTRXFILE!P_RETAIL), 0, RSTRTRXFILE!P_RETAIL)
                grdsales.TextMatrix(I, 19) = IIf(IsNull(RSTRTRXFILE!P_WS), 0, RSTRTRXFILE!P_WS)
                grdsales.TextMatrix(I, 20) = IIf(IsNull(RSTRTRXFILE!P_CRTN), 0, RSTRTRXFILE!P_CRTN)
                grdsales.TextMatrix(I, 37) = IIf(IsNull(RSTRTRXFILE!P_LWS), 0, RSTRTRXFILE!P_LWS)
                If RSTRTRXFILE!COM_FLAG = "A" Then
                    grdsales.TextMatrix(I, 21) = 0
                    grdsales.TextMatrix(I, 22) = IIf(IsNull(RSTRTRXFILE!COM_AMT), 0, RSTRTRXFILE!COM_AMT)
                    grdsales.TextMatrix(I, 23) = "A"
                Else
                    grdsales.TextMatrix(I, 21) = IIf(IsNull(RSTRTRXFILE!COM_PER), 0, RSTRTRXFILE!COM_PER)
                    grdsales.TextMatrix(I, 22) = 0
                    grdsales.TextMatrix(I, 23) = "P"
                End If
                GROSSVAL = (Val(grdsales.TextMatrix(I, 9)) * IIf(Val(grdsales.TextMatrix(I, 5)) = 0, 1, Val(grdsales.TextMatrix(I, 5)))) * (Val(grdsales.TextMatrix(I, 3)) - Val(grdsales.TextMatrix(I, 14)))
                If RSTRTRXFILE!DISC_FLAG = "P" Then
                    grdsales.TextMatrix(I, 27) = "P"
                    LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - (GROSSVAL * Val(grdsales.TextMatrix(I, 17)) / 100)) * Val(grdsales.TextMatrix(I, 10)) / 100, 2))
                Else
                    grdsales.TextMatrix(I, 27) = "A"
                    LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - Val(grdsales.TextMatrix(I, 17))) * Val(grdsales.TextMatrix(I, 10)) / 100, 2))
                End If
                grdsales.TextMatrix(I, 24) = IIf(IsNull(RSTRTRXFILE!CRTN_PACK), 0, RSTRTRXFILE!CRTN_PACK)
                grdsales.TextMatrix(I, 25) = IIf(IsNull(RSTRTRXFILE!P_VAN), 0, RSTRTRXFILE!P_VAN)
                grdsales.TextMatrix(I, 26) = IIf(IsNull(RSTRTRXFILE!GROSS_AMT), 0, RSTRTRXFILE!GROSS_AMT)
                grdsales.TextMatrix(I, 28) = IIf(IsNull(RSTRTRXFILE!LOOSE_PACK), 1, RSTRTRXFILE!LOOSE_PACK)
                grdsales.TextMatrix(I, 29) = IIf(IsNull(RSTRTRXFILE!PACK_TYPE), "Nos", RSTRTRXFILE!PACK_TYPE)
                grdsales.TextMatrix(I, 30) = IIf(IsNull(RSTRTRXFILE!WARRANTY), "", RSTRTRXFILE!WARRANTY)
                grdsales.TextMatrix(I, 31) = IIf(IsNull(RSTRTRXFILE!WARRANTY_TYPE), "", RSTRTRXFILE!WARRANTY_TYPE)
                grdsales.TextMatrix(I, 32) = IIf(IsNull(RSTRTRXFILE!EXPENSE), "", RSTRTRXFILE!EXPENSE)
                grdsales.TextMatrix(I, 33) = IIf(IsNull(RSTRTRXFILE!EXDUTY), "", RSTRTRXFILE!EXDUTY)
                grdsales.TextMatrix(I, 34) = IIf(IsNull(RSTRTRXFILE!CSTPER), "", RSTRTRXFILE!CSTPER)
                grdsales.TextMatrix(I, 35) = IIf(IsNull(RSTRTRXFILE!TR_DISC), "", RSTRTRXFILE!TR_DISC)
                grdsales.TextMatrix(I, 36) = IIf(IsNull(RSTRTRXFILE!GROSS_AMOUNT), "", RSTRTRXFILE!GROSS_AMOUNT)
                grdsales.TextMatrix(I, 38) = IIf(IsNull(RSTRTRXFILE!BARCODE), "", RSTRTRXFILE!BARCODE)
                grdsales.TextMatrix(I, 39) = IIf(IsNull(RSTRTRXFILE!CESS_AMT), "", RSTRTRXFILE!CESS_AMT)
                grdsales.TextMatrix(I, 40) = IIf(IsNull(RSTRTRXFILE!CESS_PER), "", RSTRTRXFILE!CESS_PER)
                lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(I, 13)), ".00")
                LBLGROSSAMT.Caption = Val(LBLGROSSAMT.Caption) + Val(grdsales.TextMatrix(I, 8)) * Val(RSTRTRXFILE!QTY)
                LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(I, 32)), ".00")
                lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(I, 3)), ".00")
                'TXTDEALER.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
                
                'TXTINVDATE.Text = Format(RSTRTRXFILE!VCH_DATE, "DD/MM/YYYY")
                'TXTREMARKS.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
                'TXTINVOICE.Text = IIf(IsNull(RSTRTRXFILE!PINV), "", RSTRTRXFILE!PINV)
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
            If Roundflag = True Then
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
            Else
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 2), "0.00")
            End If
            LBLNET.Caption = Val(LBLGROSSAMT.Caption) + Val(LBLTOTALTAX.Caption)
            TXTSLNO.Text = Val(grdsales.Rows)
            
            M_ADD = True
            OLD_BILL = True
            grdsales.SetFocus
            Screen.MousePointer = vbNormal
        End Select
        Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number <> -2147168237 Then
        MsgBox Err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub grdsales_Scroll()
    TXTsample.Visible = False
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim I As Long
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            On Error Resume Next
            TXTITEMCODE.Text = grdtmp.Columns(0)
            TXTPRODUCT.Text = grdtmp.Columns(1)
            On Error Resume Next
            Set Image1.DataSource = PHY
            If IsNull(PHY!PHOTO) Then
                Frame6.Visible = False
                Set Image1.DataSource = Nothing
                bytData = ""
            Else
                If Err.Number = 545 Then
                    Frame6.Visible = False
                    Set Image1.DataSource = Nothing
                    bytData = ""
                Else
                    Frame6.Visible = True
                    Set Image1.DataSource = PHY 'setting image1�s datasource
                    Image1.DataField = "PHOTO"
                    bytData = PHY!PHOTO
                End If
            End If
            On Error GoTo ErrHand
            For I = 1 To grdsales.Rows - 1
                If Trim(grdsales.TextMatrix(I, 1)) = Trim(TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists in Line No. " & I & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                    Exit For
                End If
            Next I
            
            Set RSTRXFILE = New ADODB.Recordset
            'RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly, adCmdText
            RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE = 'PI' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                'RSTRXFILE.MoveLast
                TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
                If IsNull(RSTRXFILE!LINE_DISC) Then
                    Txtpack.Text = ""
                Else
                    Txtpack.Text = RSTRXFILE!LINE_DISC
                End If
                Txtpack.Text = 1
                TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                If IsNull(RSTRXFILE!REF_NO) Then
                    txtBatch.Text = ""
                Else
                    txtBatch.Text = RSTRXFILE!REF_NO
                End If
                TXTEXPIRY.Text = IIf(IsDate(RSTRXFILE!EXP_DATE), Format(RSTRXFILE!EXP_DATE, "MM/YY"), "  /  ")
                If IsNull(RSTRXFILE!MRP) Then
                    TXTRATE.Text = ""
                Else
                    TXTRATE.Text = IIf(IsNull(RSTRXFILE!MRP), "", Format(Round(Val(RSTRXFILE!MRP) * Val(Los_Pack.Text), 2), ".000"))
                End If
                If IsNull(RSTRXFILE!MRP_BT) Then
                    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105
                Else
                    txtmrpbt.Text = Format(Val(RSTRXFILE!MRP_BT), ".000")
                End If
                If IsNull(RSTRXFILE!PTR) Then
                    TXTPTR.Text = ""
                Else
                    TXTPTR.Text = Format(Round(Val(RSTRXFILE!PTR) * Val(Los_Pack.Text), 3), ".000")
                End If
                If IsNull(RSTRXFILE!P_RETAIL) Then
                    TXTRETAIL.Text = ""
                Else
                    TXTRETAIL.Text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                End If
                'TXTPTR.Text = IIf(IsNull(RSTRXFILE!PTR), "", Format(Round(Val(RSTRXFILE!PTR), 2), ".000"))
                'txtretail.Text = IIf(IsNull(RSTRXFILE!P_RETAIL), "", Format(Round(Val(RSTRXFILE!P_RETAIL) * Val(Los_Pack.Text), 2), ".000"))
                If IsNull(RSTRXFILE!P_WS) Then
                    txtWS.Text = ""
                Else
                    txtWS.Text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                End If
                If IsNull(RSTRXFILE!P_VAN) Then
                    txtvanrate.Text = ""
                Else
                    txtvanrate.Text = Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000")
                End If
                If IsNull(RSTRXFILE!P_CRTN) Then
                    txtcrtn.Text = ""
                Else
                    txtcrtn.Text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                End If
                If IsNull(RSTRXFILE!P_LWS) Then
                    TxtLWRate.Text = ""
                Else
                    TxtLWRate.Text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                End If
                If IsNull(RSTRXFILE!CRTN_PACK) Then
                    txtcrtnpack.Text = ""
                Else
                    txtcrtnpack.Text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                End If
                If IsNull(RSTRXFILE!SALES_PRICE) Then
                    txtprofit.Text = ""
                Else
                    txtprofit.Text = Format(Round(Val(RSTRXFILE!SALES_PRICE), 2), ".000")
                End If
                If IsNull(RSTRXFILE!SALES_TAX) Then
                    TxttaxMRP.Text = ""
                Else
                    TxttaxMRP.Text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                End If
                If IsNull(RSTRXFILE!EXDUTY) Then
                    TxtExDuty.Text = ""
                Else
                    TxtExDuty.Text = Format(Val(RSTRXFILE!EXDUTY), ".00")
                End If
                If IsNull(RSTRXFILE!CSTPER) Then
                    TxtCSTper.Text = ""
                Else
                    TxtCSTper.Text = Format(Val(RSTRXFILE!CSTPER), ".00")
                End If
                If IsNull(RSTRXFILE!TR_DISC) Then
                    TxtTrDisc.Text = ""
                Else
                    TxtTrDisc.Text = Format(Val(RSTRXFILE!TR_DISC), ".00")
                End If
                If IsNull(RSTRXFILE!CESS_AMT) Then
                    txtCess.Text = ""
                Else
                    txtCess.Text = Format(Val(RSTRXFILE!CESS_AMT), ".00")
                End If
                If IsNull(RSTRXFILE!CESS_PER) Then
                    TxtCessPer.Text = ""
                Else
                    TxtCessPer.Text = Format(Val(RSTRXFILE!CESS_PER), ".00")
                End If
                TxtWarranty.Text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                If RSTRXFILE!COM_FLAG = "A" Then
                    TxtComAmt.Text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                    OptComAmt.value = True
                Else
                    TxtComper.Text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                    OptComper.value = True
                End If
                On Error Resume Next
                CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                CmbWrnty.Text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                On Error GoTo ErrHand
                
                ''TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                If RSTRXFILE!CHECK_FLAG = "M" Then
                    OPTTaxMRP.value = True
                ElseIf RSTRXFILE!CHECK_FLAG = "V" Then
                    OPTVAT.value = True
                Else
                    OPTNET.value = True
                End If
            Else
                TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                Txtpack.Text = 1
                Los_Pack.Text = 1
                TxtWarranty.Text = ""
                On Error Resume Next
                CmbPack.Text = "Nos"
                CmbWrnty.ListIndex = -1
                On Error GoTo ErrHand
                
                TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                txtBatch.Text = ""
                TXTEXPIRY.Text = "  /  "
                TXTRATE.Text = ""
                txtmrpbt.Text = ""
                TXTPTR.Text = ""
                TxtNetrate.Text = ""
                TXTRETAIL.Text = ""
                txtWS.Text = ""
                txtvanrate.Text = ""
                txtcrtn.Text = ""
                TxtLWRate.Text = ""
                txtcrtnpack.Text = ""
                txtprofit.Text = ""
                TxttaxMRP.Text = ""
                Los_Pack.Text = "1"
                TxtWarranty.Text = ""
                On Error Resume Next
                CmbPack.Text = "Nos"
                CmbWrnty.ListIndex = -1
                On Error GoTo ErrHand
                OPTVAT.value = True
            End If
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing
            
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With RSTRXFILE
                If Not (.EOF And .BOF) Then
                    lblcategory.Caption = IIf(IsNull(RSTRXFILE!Category), "", RSTRXFILE!Category)
                    If IsNull(RSTRXFILE!P_RETAIL) Then
                        TXTRETAIL.Text = ""
                    Else
                        TXTRETAIL.Text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.Text = ""
                    Else
                        TxttaxMRP.Text = Format(Round(Val(RSTRXFILE!SALES_TAX), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_WS) Then
                        txtWS.Text = ""
                    Else
                        txtWS.Text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_VAN) Then
                        txtvanrate.Text = ""
                    Else
                        txtvanrate.Text = Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000")
                    End If
                    If RSTRXFILE!COM_FLAG = "A" Then
                        TxtComAmt.Text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                        OptComAmt.value = True
                    Else
                        TxtComper.Text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                        OptComper.value = True
                    End If
                    If IsNull(RSTRXFILE!P_CRTN) Then
                        txtcrtn.Text = ""
                    Else
                        txtcrtn.Text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_LWS) Then
                        TxtLWRate.Text = ""
                    Else
                        TxtLWRate.Text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!CRTN_PACK) Then
                        txtcrtnpack.Text = ""
                    Else
                        txtcrtnpack.Text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.Text = ""
                    Else
                        TxttaxMRP.Text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                    End If
                End If
            End With
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing
            
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Fram.Enabled = True
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            txtcategory.Enabled = False
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                Los_Pack.Text = 1
                TXTQTY.Text = 1
                TxTfree.Text = ""
                TXTRATE.Text = ""
                TXTPTR.Enabled = True
                TXTPTR.SetFocus
            Else
                Los_Pack.Enabled = True
                Los_Pack.SetFocus
            End If
            'TxtPack.Enabled = True
            'TxtPack.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TxTfree.Text = ""
            Fram.Enabled = True
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Los_Pack_LostFocus()
    Call CHANGEBOXCOLOR(Los_Pack, False)
End Sub

Private Sub OptComper_LostFocus()
    cmbfull.BackColor = vbWhite
End Sub

Private Sub Optdiscamt_Click()
    Call TxttaxMRP_LostFocus
End Sub

Private Sub optdiscper_Click()
    Call TxttaxMRP_LostFocus
End Sub

Private Sub OPTNET_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxttaxMRP.Text) <> 0 Then
                If OPTTaxMRP.value = False And OPTVAT.value = False Then
                'If OPTVAT.Value = False Then
                    MsgBox "Tax should be Zero ....", vbOKOnly, "EzBiz"
                    TxttaxMRP.Enabled = True
                    TxttaxMRP.SetFocus
                    Exit Sub
                End If
            End If
            If TxttaxMRP.Enabled = True Then
                'TxttaxMRP.Enabled = False
                txtPD.Enabled = True
                txtPD.SetFocus
            ElseIf cmdadd.Enabled = True Then
                cmdadd.SetFocus
            End If
        Case vbKeyEscape
'            TxttaxMRP.Enabled = True
'            TxttaxMRP.SetFocus
    End Select
End Sub

Private Sub OPTNET_LostFocus()
    OPTNET.BackColor = vbWhite
End Sub

Private Sub OPTTaxMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
'            If Val(TxttaxMRP.Text) <> 0 Then
'                If OPTTaxMRP.Value = False And OPTVAT.Value = False Then
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "EzBiz"
'                    Exit Sub
'                End If
'            End If
            If TxttaxMRP.Enabled = True Then
                'TxttaxMRP.Enabled = False
                txtPD.Enabled = True
                txtPD.SetFocus
            ElseIf cmdadd.Enabled = True Then
                cmdadd.SetFocus
            End If
        Case vbKeyEscape
'            TxttaxMRP.Enabled = True
'            TxttaxMRP.SetFocus
    End Select
End Sub

Private Sub OPTTaxMRP_LostFocus()
    OPTTaxMRP.BackColor = vbWhite
End Sub

Private Sub OPTVAT_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
        Case vbKeyReturn
'            If Val(TxttaxMRP.Text) <> 0 Then
'                If OPTTaxMRP.Value = False And OPTVAT.Value = False Then
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "EzBiz"
'                    Exit Sub
'                End If
'            End If
            If TxttaxMRP.Enabled = True Then
                'TxttaxMRP.Enabled = False
                txtPD.Enabled = True
                txtPD.SetFocus
            ElseIf cmdadd.Enabled = True Then
                cmdadd.SetFocus
            End If
        Case vbKeyEscape
'            TxttaxMRP.Enabled = True
'            TxttaxMRP.SetFocus
    End Select

End Sub

Private Sub OPTVAT_LostFocus()
    OPTVAT.BackColor = vbWhite
End Sub

Private Sub txtbarcode_GotFocus()
    Call CHANGEBOXCOLOR(TxtBarcode, True)
    Call CHANGEBOXCOLOR(txtcrtn, False)
    
    TxtBarcode.SelStart = 0
    TxtBarcode.SelLength = Len(TxtBarcode.Text)
    FRMEGRDTMP.Visible = False
    TXTSLNO.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxttaxMRP.Enabled = False
    TxtExDuty.Enabled = False
    TxtTrDisc.Enabled = False
    TxtCessPer.Enabled = False
    txtCess.Enabled = False
    TxtCSTper.Enabled = False
    txtPD.Enabled = False
    TxtExpense.Enabled = False
    TXTRETAIL.Enabled = False
    TxtRetailPercent.Enabled = False
    
    txtWS.Enabled = False
    txtWsalePercent.Enabled = False
    txtvanrate.Enabled = False
    txtSchPercent.Enabled = False
    txtcrtnpack.Enabled = False
    txtcrtn.Enabled = False
    TxtLWRate.Enabled = False
    TxtCustDisc.Enabled = False
    TxtComper.Enabled = False
    TxtComAmt.Enabled = False
    cmdadd.Enabled = False
    Txtgrossamt.Enabled = False
    txtBatch.Enabled = False
    txtHSN.Enabled = False
    TxtWarranty.Enabled = False
    CmbWrnty.Enabled = False
    TXTEXPIRY.Visible = False
    TXTEXPDATE.Enabled = False
    TxtBarcode.Enabled = True
End Sub

Private Sub TxtBarcode_LostFocus()
    Call CHANGEBOXCOLOR(TxtBarcode, False)
End Sub

Private Sub TXTBATCH_GotFocus()
    Call CHANGEBOXCOLOR(txtBatch, True)
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                TxttaxMRP.Enabled = True
                TxttaxMRP.SetFocus
            Else
                TXTEXPIRY.Visible = True
                TXTEXPIRY.SetFocus
            End If
        Case vbKeyEscape
            TXTPTR.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBatch_LostFocus()
    Call CHANGEBOXCOLOR(txtBatch, False)
End Sub

Private Sub TXTBILLNO_GotFocus()
    Call CHANGEBOXCOLOR(txtBillNo, True)
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
    'txtBillNo.ForeColor = &HFFFF&
End Sub

Public Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTDIST As ADODB.Recordset
    Dim RSTTRNSMAST As ADODB.Recordset
    Dim I As Long

    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtBillNo.Text) > 200 Then Exit Sub
'            Set rsttrxmast = New ADODB.Recordset
'            rsttrxmast.Open "Select MAX(VCH_NO) From TRXMAST", db, adOpenForwardOnly
'            If Not (rsttrxmast.EOF And rsttrxmast.BOF) Then
'                i = IIf(IsNull(rsttrxmast.Fields(0)), 1, rsttrxmast.Fields(0))
'                If i > 3100 Then
'                    rsttrxmast.Close
'                    Set rsttrxmast = Nothing
'                    Exit Sub
'                End If
'            End If
'            rsttrxmast.Close
'            Set rsttrxmast = Nothing
            Chkcancel.value = 0
            grdsales.Rows = 1
            I = 0
            PONO = ""
            CMBPO.Text = ""
            LBLTOTAL.Caption = ""
            lbltotalwodiscount = ""
            LBLTOTALTAX.Caption = ""
            LBLGROSSAMT.Caption = ""
            LBLEXP.Caption = ""
            lblqty.Caption = ""
            Dim GROSSVAL As Double
            grdsales.Rows = 1
            OLD_BILL = False
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until rstTRXMAST.EOF
                grdsales.Rows = grdsales.Rows + 1
                grdsales.FixedRows = 1
                I = I + 1
                
                grdsales.TextMatrix(I, 0) = I
                grdsales.TextMatrix(I, 1) = rstTRXMAST!ITEM_CODE
                grdsales.TextMatrix(I, 2) = rstTRXMAST!ITEM_NAME
                grdsales.TextMatrix(I, 3) = Val(rstTRXMAST!QTY) / Val(rstTRXMAST!LINE_DISC)
                grdsales.TextMatrix(I, 4) = rstTRXMAST!UNIT
                grdsales.TextMatrix(I, 5) = rstTRXMAST!LINE_DISC
                grdsales.TextMatrix(I, 6) = Format(rstTRXMAST!MRP, ".000")
                grdsales.TextMatrix(I, 7) = Format(rstTRXMAST!SALES_PRICE, ".00000")
                grdsales.TextMatrix(I, 8) = Format(rstTRXMAST!ITEM_COST, ".00000")
                grdsales.TextMatrix(I, 9) = Format(rstTRXMAST!PTR, ".0000")
                grdsales.TextMatrix(I, 10) = IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
                grdsales.TextMatrix(I, 11) = IIf(IsNull(rstTRXMAST!REF_NO), "", rstTRXMAST!REF_NO)
                grdsales.TextMatrix(I, 12) = IIf(IsNull(rstTRXMAST!EXP_DATE), "", Format(rstTRXMAST!EXP_DATE, "DD/MM/YYYY"))
                grdsales.TextMatrix(I, 13) = Format(rstTRXMAST!TRX_TOTAL, ".000")
                grdsales.TextMatrix(I, 14) = IIf(IsNull(rstTRXMAST!SCHEME), "", rstTRXMAST!SCHEME)
                grdsales.TextMatrix(I, 15) = IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
                grdsales.TextMatrix(I, 16) = rstTRXMAST!line_no
                grdsales.TextMatrix(I, 17) = IIf(IsNull(rstTRXMAST!P_DISC), 0, rstTRXMAST!P_DISC)
                grdsales.TextMatrix(I, 18) = IIf(IsNull(rstTRXMAST!P_RETAIL), 0, rstTRXMAST!P_RETAIL)
                grdsales.TextMatrix(I, 19) = IIf(IsNull(rstTRXMAST!P_WS), 0, rstTRXMAST!P_WS)
                grdsales.TextMatrix(I, 20) = IIf(IsNull(rstTRXMAST!P_CRTN), 0, rstTRXMAST!P_CRTN)
                grdsales.TextMatrix(I, 37) = IIf(IsNull(rstTRXMAST!P_LWS), 0, rstTRXMAST!P_LWS)
                
                If rstTRXMAST!COM_FLAG = "A" Then
                    grdsales.TextMatrix(I, 21) = ""
                    grdsales.TextMatrix(I, 22) = IIf(IsNull(rstTRXMAST!COM_AMT), 0, rstTRXMAST!COM_AMT)
                    grdsales.TextMatrix(I, 23) = "A"
                Else
                    grdsales.TextMatrix(I, 21) = IIf(IsNull(rstTRXMAST!COM_PER), 0, rstTRXMAST!COM_PER)
                    grdsales.TextMatrix(I, 22) = ""
                    grdsales.TextMatrix(I, 23) = "P"
                End If
                grdsales.TextMatrix(I, 24) = IIf(IsNull(rstTRXMAST!CRTN_PACK), 0, rstTRXMAST!CRTN_PACK)
                grdsales.TextMatrix(I, 25) = IIf(IsNull(rstTRXMAST!P_VAN), 0, rstTRXMAST!P_VAN)
                grdsales.TextMatrix(I, 26) = IIf(IsNull(rstTRXMAST!GROSS_AMT), 0, Format(rstTRXMAST!GROSS_AMT, "0.00"))
                GROSSVAL = (Val(grdsales.TextMatrix(I, 9)) * IIf(Val(grdsales.TextMatrix(I, 5)) = 0, 1, Val(grdsales.TextMatrix(I, 5)))) * (Val(grdsales.TextMatrix(I, 3)) - Val(grdsales.TextMatrix(I, 14)))
                If rstTRXMAST!DISC_FLAG = "P" Then
                    grdsales.TextMatrix(I, 27) = "P"
                    LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - (GROSSVAL * Val(grdsales.TextMatrix(I, 17)) / 100)) * Val(grdsales.TextMatrix(I, 10)) / 100, 2))
                Else
                    grdsales.TextMatrix(I, 27) = "A"
                    LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - Val(grdsales.TextMatrix(I, 17))) * Val(grdsales.TextMatrix(I, 10)) / 100, 2))
                End If
                grdsales.TextMatrix(I, 28) = IIf(IsNull(rstTRXMAST!LOOSE_PACK), 1, rstTRXMAST!LOOSE_PACK)
                grdsales.TextMatrix(I, 29) = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)
                grdsales.TextMatrix(I, 30) = IIf(IsNull(rstTRXMAST!WARRANTY), "", rstTRXMAST!WARRANTY)
                grdsales.TextMatrix(I, 31) = IIf(IsNull(rstTRXMAST!WARRANTY_TYPE), "", rstTRXMAST!WARRANTY_TYPE)
                grdsales.TextMatrix(I, 32) = IIf(IsNull(rstTRXMAST!EXPENSE), "", rstTRXMAST!EXPENSE)
                grdsales.TextMatrix(I, 33) = IIf(IsNull(rstTRXMAST!EXDUTY), "", rstTRXMAST!EXDUTY)
                grdsales.TextMatrix(I, 34) = IIf(IsNull(rstTRXMAST!CSTPER), "", rstTRXMAST!CSTPER)
                grdsales.TextMatrix(I, 35) = IIf(IsNull(rstTRXMAST!TR_DISC), "", rstTRXMAST!TR_DISC)
                grdsales.TextMatrix(I, 36) = IIf(IsNull(rstTRXMAST!GROSS_AMOUNT), "", rstTRXMAST!GROSS_AMOUNT)
                grdsales.TextMatrix(I, 38) = IIf(IsNull(rstTRXMAST!BARCODE), "", rstTRXMAST!BARCODE)
                grdsales.TextMatrix(I, 39) = IIf(IsNull(rstTRXMAST!CESS_AMT), "", rstTRXMAST!CESS_AMT)
                grdsales.TextMatrix(I, 40) = IIf(IsNull(rstTRXMAST!CESS_PER), "", rstTRXMAST!CESS_PER)
                grdsales.TextMatrix(I, 41) = Val(grdsales.TextMatrix(I, 3))
                LBLGROSSAMT.Caption = Val(LBLGROSSAMT.Caption) + Val(grdsales.TextMatrix(I, 8)) * Val(rstTRXMAST!QTY)
                lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(I, 13)), ".00")
                LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(I, 32)), ".00")
                lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(I, 3)), ".00")
                TXTDEALER.Text = IIf(IsNull(rstTRXMAST!VCH_DESC), "", Mid(rstTRXMAST!VCH_DESC, 15))
                PONO = IIf(IsNull(rstTRXMAST!PO_NO), "", rstTRXMAST!PO_NO)
                On Error Resume Next
                TXTINVDATE.Text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                TXTRCVDATE.Text = Format(rstTRXMAST!RCVD_DATE, "DD/MM/YYYY")
                OLD_BILL = True
                On Error GoTo ErrHand
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            LBLNET.Caption = Val(LBLGROSSAMT.Caption) + Val(LBLTOTALTAX.Caption)
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                TXTDISCAMOUNT.Text = IIf(IsNull(rstTRXMAST!DISCOUNT), "", Format(rstTRXMAST!DISCOUNT, ".00"))
                txtaddlamt.Text = IIf(IsNull(rstTRXMAST!ADD_AMOUNT), "", Format(rstTRXMAST!ADD_AMOUNT, ".00"))
                txtcramt.Text = IIf(IsNull(rstTRXMAST!DISC_PERS), "", Format(rstTRXMAST!DISC_PERS, ".00"))
                TxtCST.Text = IIf(IsNull(rstTRXMAST!CST_PER), "", Format(rstTRXMAST!CST_PER, ".00"))
                TxtInsurance.Text = IIf(IsNull(rstTRXMAST!INS_PER), "", Format(rstTRXMAST!INS_PER, ".00"))
                'If rstTRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                lblcredit.Caption = "1"
                TXTREMARKS.Text = IIf(IsNull(rstTRXMAST!Remarks), "", rstTRXMAST!Remarks)
                On Error Resume Next
                If grdsales.Rows <= 1 Then
                    TXTINVDATE.Text = "  /  /    "
                    TXTRCVDATE.Text = "  /  /    "
                    TXTDATE.Text = Format(Date, "DD/MM/YYYY")
                Else
                    TXTINVDATE.Text = IIf(IsDate(rstTRXMAST!VCH_DATE), Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY"), "  /  /    ")
                    TXTRCVDATE.Text = IIf(IsDate(rstTRXMAST!RCVD_DATE), Format(rstTRXMAST!RCVD_DATE, "DD/MM/YYYY"), TXTINVDATE.Text)
                    TXTDATE.Text = Format(rstTRXMAST!CREATE_DATE, "DD/MM/YYYY")
                End If
                On Error GoTo ErrHand
                TXTINVOICE.Text = IIf(IsNull(rstTRXMAST!PINV), "", rstTRXMAST!PINV)
                CMBDISTRICT.Text = IIf(IsNull(rstTRXMAST!TRX_GODOWN), "", rstTRXMAST!TRX_GODOWN)
                'OLD_BILL = True
            Else
                TXTDATE.Text = Format(Date, "DD/MM/YYYY")
                'OLD_BILL = False
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            ''''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
            'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
            If Roundflag = True Then
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
            Else
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 2), "0.00")
            End If
            
            TXTSLNO.Text = grdsales.Rows
            TXTSLNO.Enabled = True
            txtBillNo.Enabled = False
            FRMEMASTER.Enabled = True
            If grdsales.Rows > 1 Then
                FRMEMASTER.Enabled = True
                FRMECONTROLS.Enabled = True
                cmdRefresh.Enabled = True
                cmdRefresh.SetFocus
            Else
                TXTDEALER.SetFocus
            End If
            
'            Set RSTTRNSMAST = New ADODB.Recordset
'            RSTTRNSMAST.Open "Select CHECK_FLAG From TRANSMAST WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
'            If Not (RSTTRNSMAST.EOF Or RSTTRNSMAST.BOF) Then
'                If RSTTRNSMAST!CHECK_FLAG = "Y" Then FRMEMASTER.Enabled = False
'            End If
'            RSTTRNSMAST.Close
'            Set RSTTRNSMAST = Nothing
    
    End Select
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    CMBPO.Text = PONO
    Exit Sub
ErrHand:
    MsgBox Err.Description
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
    Call CHANGEBOXCOLOR(txtBillNo, False)
    If Val(txtBillNo.Text) = 0 Or Val(txtBillNo.Text) > Val(TXTLASTBILL.Text) Then txtBillNo.Text = TXTLASTBILL.Text
    'txtBillNo.BackColor = &HFFFFFF
    'txtBillNo.ForeColor = &H0&
End Sub

Private Sub txtcategory_Change()
    Dim RSTRXFILE As ADODB.Recordset
    Dim I As Long
    On Error GoTo ErrHand
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, MRP, ITEM_COST, SALES_TAX, ITEM_NET_COST, P_RETAIL, P_WS From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '" & Me.txtcategory.Text & "%') AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHYFLAG = False
         Else
             PHY.Close
             'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, MRP, ITEM_COST, SALES_TAX, ITEM_NET_COST, P_RETAIL, P_WS From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '" & Me.txtcategory.Text & "%') AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         End If
         
        Set grdtmp.DataSource = PHY
        
        If PHY.RecordCount > 0 Then
            FRMEGRDTMP.Visible = True
        Else
            FRMEGRDTMP.Visible = False
            Exit Sub
        End If
'        grdtmp.Columns(0).Visible = True
        grdtmp.Columns(0).Caption = "ITEM CODE"
        grdtmp.Columns(0).Width = 1300
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(1).Width = 5000
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(2).Width = 900
        grdtmp.Columns(2).Caption = "QTY"
        grdtmp.Columns(3).Width = 900
        grdtmp.Columns(3).Caption = "MRP"
        grdtmp.Columns(4).Width = 950
        grdtmp.Columns(4).Caption = "COST"
        grdtmp.Columns(5).Caption = "TAX%"
        grdtmp.Columns(5).Width = 800
        grdtmp.Columns(6).Caption = "NET COST"
        grdtmp.Columns(6).Width = 950
        grdtmp.Columns(7).Caption = "R. PRICE"
        grdtmp.Columns(7).Width = 950
        grdtmp.Columns(8).Caption = "W. PRICE"
        grdtmp.Columns(8).Width = 950

        Exit Sub
ErrHand:
        MsgBox Err.Description
End Sub

Private Sub txtcategory_GotFocus()
    Call CHANGEBOXCOLOR(txtcategory, True)
    Call CHANGEBOXCOLOR(txtcrtn, False)
    Call CHANGEBOXCOLOR(TxtLWRate, False)
    
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.Text)
    FRMEGRDTMP.Visible = False
    'TXTSLNO.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxttaxMRP.Enabled = False
    TxtExDuty.Enabled = False
    TxtTrDisc.Enabled = False
    TxtCessPer.Enabled = False
    txtCess.Enabled = False
    TxtCSTper.Enabled = False
    txtPD.Enabled = False
    TxtExpense.Enabled = False
    TXTRETAIL.Enabled = False
    TxtRetailPercent.Enabled = False
    
    txtWS.Enabled = False
    txtWsalePercent.Enabled = False
    txtvanrate.Enabled = False
    txtSchPercent.Enabled = False
    txtcrtnpack.Enabled = False
    txtcrtn.Enabled = False
    TxtLWRate.Enabled = False
    TxtCustDisc.Enabled = False
    TxtComper.Enabled = False
    TxtComAmt.Enabled = False
    cmdadd.Enabled = False
    Txtgrossamt.Enabled = False
    txtBatch.Enabled = False
    txtHSN.Enabled = False
    TxtWarranty.Enabled = False
    CmbWrnty.Enabled = False
    TXTEXPIRY.Visible = False
    TXTEXPDATE.Enabled = False
    TxtBarcode.Enabled = True
    TXTPRODUCT.Enabled = True
End Sub

Private Sub txtcategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
        Case vbKeyReturn
            txtcategory.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            TxtBarcode.Enabled = True
            TxtBarcode.SetFocus
            Exit Sub
            TXTSLNO.Enabled = True
            txtcategory.Enabled = False
            TXTSLNO.SetFocus
    End Select
End Sub

Private Sub txtcategory_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcategory_LostFocus()
    Call CHANGEBOXCOLOR(txtcategory, False)
End Sub

Private Sub TxtCSTper_LostFocus()
    Call CHANGEBOXCOLOR(TxtCSTper, False)
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TXTDEALER_LostFocus()
    Call CHANGEBOXCOLOR(TXTDEALER, False)
End Sub

Private Sub TxtExDuty_LostFocus()
    Call CHANGEBOXCOLOR(TxtExDuty, False)
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TXTEXPDATE_GotFocus()
    TXTEXPDATE.BackColor = &H98F3C1
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
End Sub

Private Sub TXTEXPDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Len(Trim(TXTEXPDATE.Text)) = 4 Then GoTo SKID
            If Not IsDate(TXTEXPDATE.Text) Then Exit Sub
            If DateDiff("d", Date, TXTEXPDATE.Text) < 0 Then
                MsgBox "Item Expired....", vbOKOnly, "EzBiz"
                TXTEXPDATE.SelStart = 0
                TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 60 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days", vbOKOnly, "EzBiz"
                TXTEXPDATE.SelStart = 0
                TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 180 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "EzBiz") = vbNo Then
                    TXTEXPDATE.SelStart = 0
                    TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
                    TXTEXPDATE.SetFocus
                    Exit Sub
                End If
            End If
SKID:
            TxttaxMRP.Enabled = True
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
            TxttaxMRP.SetFocus
        Case vbKeyEscape
            If TXTEXPDATE.Text = "  /  /    " Then GoTo SKIP
            If Not IsDate(TXTEXPDATE.Text) Then Exit Sub
SKIP:
            txtBatch.Enabled = True
            TXTEXPDATE.Enabled = False
            TXTEXPIRY.Visible = False
            txtBatch.SetFocus
    End Select
End Sub

Private Sub TXTEXPDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKey0 To vbKey9, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTEXPDATE_LostFocus()
    'Call CHANGEBOXCOLOR(txtBillNo, False)
    TXTEXPDATE.BackColor = vbWhite
    TXTEXPDATE.Text = Format(TXTEXPDATE.Text, "DD/MM/YYYY")
    If IsDate(TXTEXPDATE.Text) Then TXTEXPIRY.Text = Format(TXTEXPDATE.Text, "MM/YY")
End Sub

Private Sub TxtExpense_LostFocus()
    Call CHANGEBOXCOLOR(TxtExpense, False)
End Sub

Private Sub TxtFree_GotFocus()
    Call CHANGEBOXCOLOR(TxTfree, True)
    TxTfree.SelStart = 0
    TxTfree.SelLength = Len(TxTfree.Text)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRATE.SetFocus
        Case vbKeyEscape
            TXTQTY.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtFree_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtFree_LostFocus()
    Call CHANGEBOXCOLOR(TxTfree, False)
    If Val(TxTfree.Text) = 0 Then TxTfree.Text = 0
    TxTfree.Text = Format(TxTfree.Text, "0.00")
End Sub

Private Sub TxtHSN_LostFocus()
    Call CHANGEBOXCOLOR(txtHSN, False)
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.BackColor = &H98F3C1
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.Text = "  /  /    " Then
                TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTINVDATE.SetFocus
                Exit Sub
            End If
            
            If (DateValue(TXTINVDATE.Text) < DateValue(MDIMAIN.DTFROM.value)) Or (DateValue(TXTINVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.value))) Then
                'db.Execute "delete from Users"
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTINVDATE.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
            End If
            
        Case vbKeyEscape
            TXTINVOICE.SetFocus
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

Private Sub TXTINVDATE_LostFocus()
    TXTINVDATE.BackColor = vbWhite
    If Not IsDate(TXTRCVDATE.Text) And IsDate(TXTINVDATE.Text) Then
       TXTRCVDATE.Text = Format(TXTINVDATE, "DD/MM/YYYY")
    End If
End Sub

Private Sub TXTINVOICE_GotFocus()
    Call CHANGEBOXCOLOR(TXTINVOICE, True)
    TXTINVOICE.BackColor = &H98F3C1
    TXTINVOICE.SelStart = 0
    TXTINVOICE.SelLength = Len(TXTINVOICE.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTINVOICE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVOICE.Text = "" Then
                MsgBox "Please enter the Invoice Number", vbOKOnly, "EzBiz"
                TXTINVOICE.SetFocus
                Exit Sub
            End If
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND PINV = '" & Trim(TXTINVOICE.Text) & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                MsgBox "You have already entered this Invoice number for " & Trim(DataList2.Text) & " as Computer Bill No. " & rstTRXMAST!VCH_NO, vbOKOnly, "EzBiz"
                TXTINVOICE.SetFocus
            Else
                TXTINVDATE.SetFocus
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
        Case vbKeyEscape
            DataList2.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TXTINVOICE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTINVOICE_LostFocus()
    Call CHANGEBOXCOLOR(TXTINVOICE, False)
End Sub

Private Sub Txtpack_GotFocus()
    Call CHANGEBOXCOLOR(Txtpack, True)
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.Text)
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtpack.Text) = 0 Then Exit Sub
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            Txtpack.Enabled = False
            CmbPack.Enabled = True
            CmbPack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            Txtpack.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
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

Private Sub TXTPRODUCT_Change()
    Dim RSTRXFILE As ADODB.Recordset
    Dim I As Long
    On Error GoTo ErrHand
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, MRP, ITEM_COST, SALES_TAX, ITEM_NET_COST, P_RETAIL, P_WS From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.Text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHYFLAG = False
         Else
             PHY.Close
             PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, MRP, ITEM_COST, SALES_TAX, ITEM_NET_COST, P_RETAIL, P_WS From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.Text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         End If
         
        Set grdtmp.DataSource = PHY
        
        If PHY.RecordCount > 0 Then
            FRMEGRDTMP.Visible = True
        Else
            FRMEGRDTMP.Visible = False
            Exit Sub
        End If
        grdtmp.Columns(0).Caption = "ITEM CODE"
        grdtmp.Columns(0).Width = 1300
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(1).Width = 5000
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(2).Width = 900
        grdtmp.Columns(2).Caption = "QTY"
        grdtmp.Columns(3).Width = 900
        grdtmp.Columns(3).Caption = "MRP"
        grdtmp.Columns(4).Width = 950
        grdtmp.Columns(4).Caption = "COST"
        grdtmp.Columns(5).Caption = "TAX%"
        grdtmp.Columns(5).Width = 800
        grdtmp.Columns(6).Caption = "NET COST"
        grdtmp.Columns(6).Width = 950
        grdtmp.Columns(7).Caption = "R. PRICE"
        grdtmp.Columns(7).Width = 950
        grdtmp.Columns(8).Caption = "W. PRICE"
        grdtmp.Columns(8).Width = 950
        Exit Sub
ErrHand:
        MsgBox Err.Description
                
End Sub

Private Sub TXTPRODUCT_GotFocus()
    Call CHANGEBOXCOLOR(TXTPRODUCT, True)
    Call CHANGEBOXCOLOR(txtcrtn, False)
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
    If Trim(TXTPRODUCT.Text) <> "" Or Trim(txtcategory.Text) <> "" Then Call TXTPRODUCT_Change
    'TXTSLNO.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxttaxMRP.Enabled = False
    TxtExDuty.Enabled = False
    TxtTrDisc.Enabled = False
    TxtCessPer.Enabled = False
    txtCess.Enabled = False
    TxtCSTper.Enabled = False
    txtPD.Enabled = False
    TxtExpense.Enabled = False
    TXTRETAIL.Enabled = False
    TxtRetailPercent.Enabled = False
    
    txtWS.Enabled = False
    txtWsalePercent.Enabled = False
    txtvanrate.Enabled = False
    txtSchPercent.Enabled = False
    txtcrtnpack.Enabled = False
    txtcrtn.Enabled = False
    TxtLWRate.Enabled = False
    TxtCustDisc.Enabled = False
    TxtComper.Enabled = False
    TxtComAmt.Enabled = False
    cmdadd.Enabled = False
    Txtgrossamt.Enabled = False
    txtBatch.Enabled = False
    txtHSN.Enabled = False
    TxtWarranty.Enabled = False
    CmbWrnty.Enabled = False
    TXTEXPIRY.Visible = False
    TXTEXPDATE.Enabled = False
    TxtBarcode.Enabled = True
    txtcategory.Enabled = True
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    Dim I As Long
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
        Case vbKeyReturn
            On Error Resume Next
            TXTITEMCODE.Text = ""
            TXTITEMCODE.Text = grdtmp.Columns(0)
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If TXTITEMCODE.Text <> "" Then
                Call TxtItemcode_KeyDown(13, 0)
                Exit Sub
            End If
            If Trim(txtcategory.Text) = "" Then
                If MsgBox("Item not exists!!! Do You want to add this item?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
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
                RSTITEMMAST!ITEM_NAME = Trim(TXTPRODUCT.Text)
                RSTITEMMAST!Category = "GENERAL"
                RSTITEMMAST!UNIT = 1
                RSTITEMMAST!MANUFACTURER = "GENERAL"
                RSTITEMMAST!DEAD_STOCK = "N"
                RSTITEMMAST!Remarks = ""
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
                If PC_FLAG = "Y" Then
                    RSTITEMMAST!PRICE_CHANGE = "Y"
                Else
                    RSTITEMMAST!PRICE_CHANGE = "N"
                End If
                RSTITEMMAST.Update
                db.CommitTrans
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                TXTITEMCODE.Text = TXTPRODUCT.Tag
                
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT CATEGORY FROM CATEGORY where CATEGORY = 'GENERAL'", db, adOpenStatic, adLockOptimistic, adCmdText
                If (rststock.EOF And rststock.BOF) Then
                    rststock.AddNew
                    rststock!Category = "GENERAL"
                    rststock.Update
                End If
                rststock.Close
                Set rststock = Nothing


                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT MANUFACTURER FROM MANUFACT where MANUFACTURER = 'GENERAL'", db, adOpenStatic, adLockOptimistic, adCmdText
                If (rststock.EOF And rststock.BOF) Then
                    rststock.AddNew
                    rststock!MANUFACTURER = "GENERAL"
                    rststock.Update
                End If
                rststock.Close
                Set rststock = Nothing
                    
                Call TxtItemcode_KeyDown(13, 0)
                'frmitemmaster.Show
                'frmitemmaster.TXTITEM.Text = Trim(TXTPRODUCT.Text)
                'frmitemmaster.LBLLP.Caption = "P"
                'MsgBox "Item not found!!!!", , "EzBiz"
                Exit Sub
            Else
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME = '" & Trim(TXTPRODUCT.Text) & "' ", db, adOpenForwardOnly
                If Not (RSTITEMMAST.EOF Or RSTITEMMAST.BOF) Then
                    MsgBox "Item Name already exists with Item Code " & RSTITEMMAST!ITEM_CODE, , "EzBiz"
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
                    Exit Sub
                End If
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                        
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(txtcategory.Text) & "' ", db, adOpenForwardOnly
                If Not (RSTITEMMAST.EOF Or RSTITEMMAST.BOF) Then
                    If MsgBox("Item Code exists for " & RSTITEMMAST!ITEM_NAME & " Do You want to add this item with a system generated Item Code?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        Exit Sub
                    Else
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
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
                        RSTITEMMAST!ITEM_NAME = Trim(TXTPRODUCT.Text)
                        RSTITEMMAST!Category = "GENERAL"
                        RSTITEMMAST!UNIT = 1
                        RSTITEMMAST!MANUFACTURER = "GENERAL"
                        RSTITEMMAST!DEAD_STOCK = "N"
                        RSTITEMMAST!Remarks = ""
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
                        If PC_FLAG = "Y" Then
                            RSTITEMMAST!PRICE_CHANGE = "Y"
                        Else
                            RSTITEMMAST!PRICE_CHANGE = "N"
                        End If
                        RSTITEMMAST.Update
                        db.CommitTrans
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        TXTITEMCODE.Text = TXTPRODUCT.Tag
                        txtcategory.Text = TXTPRODUCT.Tag
                        Call TxtItemcode_KeyDown(13, 0)
                        Exit Sub
                    End If
                Else
                    If MsgBox("Are you sure you want to add this item with this Item Code?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        Exit Sub
                    Else
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        
                        db.BeginTrans
                        Set RSTITEMMAST = New ADODB.Recordset
                        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(txtcategory.Text) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
                        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                            RSTITEMMAST.AddNew
                            RSTITEMMAST!ITEM_CODE = Trim(txtcategory.Text)
                        End If
                        'RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
                        
                        RSTITEMMAST!ITEM_NAME = Trim(TXTPRODUCT.Text)
                        RSTITEMMAST!Category = "GENERAL"
                        RSTITEMMAST!UNIT = 1
                        RSTITEMMAST!MANUFACTURER = "GENERAL"
                        RSTITEMMAST!DEAD_STOCK = "N"
                        RSTITEMMAST!Remarks = ""
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
                        If PC_FLAG = "Y" Then
                            RSTITEMMAST!PRICE_CHANGE = "Y"
                        Else
                            RSTITEMMAST!PRICE_CHANGE = "N"
                        End If
                        RSTITEMMAST.Update
                        db.CommitTrans
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        TXTITEMCODE.Text = Trim(txtcategory.Text)
                        Call TxtItemcode_KeyDown(13, 0)
                        Exit Sub
                    End If
                End If
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                'Call TxtItemcode_KeyDown(13, 0)
            End If
            Exit Sub
            'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If Trim(TXTPRODUCT.Text) = "" Then
                txtcategory.Enabled = True
                txtcategory.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False
                
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            
            If PHY.RecordCount = 0 Then
                If MsgBox("Item not exists!!! Do You want to add this item?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
                frmitemmaster.Show
                frmitemmaster.TXTITEM.Text = Trim(TXTPRODUCT.Text)
                'MsgBox "Item not found!!!!", , "EzBiz"
                Exit Sub
            End If
            
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For I = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(I, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & I & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                        Exit For
                    End If
                Next I

                Set RSTRXFILE = New ADODB.Recordset
                'RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE = 'PI' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    'RSTRXFILE.MoveLast
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.Text = 1
                    Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
                    TxtWarranty.Text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.Text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo ErrHand
                    
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.Text = IIf(IsNull(RSTRXFILE!REF_NO), "", RSTRXFILE!REF_NO)
                    TXTEXPIRY.Text = IIf(IsDate(RSTRXFILE!EXP_DATE), Format(RSTRXFILE!EXP_DATE, "MM/YY"), "  /  ")
                    If (IsNull(RSTRXFILE!MRP)) Then
                        TXTRATE.Text = ""
                    Else
                        TXTRATE.Text = Format(Round(Val(RSTRXFILE!MRP) * Val(Los_Pack.Text), 2), ".000")
                    End If
                    If (IsNull(RSTRXFILE!MRP_BT)) Then
                        txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105
                    Else
                        txtmrpbt.Text = Val(TXTRATE.Text)
                    End If
                    If IsNull(RSTRXFILE!PTR) Then
                        TXTPTR.Text = ""
                    Else
                        TXTPTR.Text = Format(Round(Val(RSTRXFILE!PTR) * Val(Los_Pack.Text), 3), ".000")
                        'TXTPTR.Text = Format(Round(Val(RSTRXFILE!PTR), 3), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_RETAIL) Then
                        TXTRETAIL.Text = ""
                    Else
                        TXTRETAIL.Text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_WS) Then
                        txtWS.Text = ""
                    Else
                        txtWS.Text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_VAN) Then
                        txtvanrate.Text = ""
                    Else
                        txtvanrate.Text = Format(Round(Val(RSTRXFILE!P_VAN) * Val(Los_Pack.Text), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_CRTN) Then
                        txtcrtn.Text = ""
                    Else
                        txtcrtn.Text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_LWS) Then
                        TxtLWRate.Text = ""
                    Else
                        TxtLWRate.Text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!CRTN_PACK) Then
                        txtcrtnpack.Text = ""
                    Else
                        txtcrtnpack.Text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_PRICE) Then
                        txtprofit.Text = ""
                    Else
                        txtprofit.Text = Format(Round(Val(RSTRXFILE!SALES_PRICE), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.Text = ""
                    Else
                        TxttaxMRP.Text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                    End If
                    If IsNull(RSTRXFILE!EXDUTY) Then
                        TxtExDuty.Text = ""
                    Else
                        TxtExDuty.Text = Format(Val(RSTRXFILE!EXDUTY), ".00")
                    End If
                    If IsNull(RSTRXFILE!CSTPER) Then
                        TxtCSTper.Text = ""
                    Else
                        TxtCSTper.Text = Format(Val(RSTRXFILE!CSTPER), ".00")
                    End If
                    If IsNull(RSTRXFILE!TR_DISC) Then
                        TxtTrDisc.Text = ""
                    Else
                        TxtTrDisc.Text = Format(Val(RSTRXFILE!TR_DISC), ".00")
                    End If
                    If IsNull(RSTRXFILE!CESS_AMT) Then
                        txtCess.Text = ""
                    Else
                        txtCess.Text = Format(Val(RSTRXFILE!CESS_AMT), ".00")
                    End If
                    If IsNull(RSTRXFILE!CESS_PER) Then
                        TxtCessPer.Text = ""
                    Else
                        TxtCessPer.Text = Format(Val(RSTRXFILE!CESS_PER), ".00")
                    End If
                    Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
                    TxtWarranty.Text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                    If RSTRXFILE!COM_FLAG = "A" Then
                        TxtComAmt.Text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                        OptComAmt.value = True
                    Else
                        TxtComper.Text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                        OptComper.value = True
                    End If
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.Text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo ErrHand
                
                    'TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                    If RSTRXFILE!CHECK_FLAG = "M" Then
                        OPTTaxMRP.value = True
                    ElseIf RSTRXFILE!CHECK_FLAG = "V" Then
                        OPTVAT.value = True
                    Else
                        OPTNET.value = True
                    End If
                Else
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = 1
                    Los_Pack.Text = 1
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.Text = ""
                    txtHSN.Text = ""
                    TXTEXPIRY.Text = "  /  "
                    TXTRATE.Text = ""
                    txtmrpbt.Text = ""
                    TXTPTR.Text = ""
                    TxtNetrate.Text = ""
                    TXTRETAIL.Text = ""
                    txtWS.Text = ""
                    txtvanrate.Text = ""
                    txtcrtn.Text = ""
                    TxtLWRate.Text = ""
                    txtcrtnpack.Text = ""
                    txtprofit.Text = ""
                    TxttaxMRP.Text = ""
                    TxtExDuty.Text = ""
                    TxtCSTper.Text = ""
                    TxtTrDisc.Text = ""
                    TxtCustDisc.Text = ""
                    TxtCessPer.Text = ""
                    txtCess.Text = ""
                    Los_Pack.Text = "1"
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    OPTVAT.value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                
                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    TXTITEMCODE.Enabled = False
                    Los_Pack.Enabled = True
                    Los_Pack.SetFocus
                    'TxtPack.Enabled = True
                    'TxtPack.SetFocus
                    Exit Sub
                End If
            ElseIf PHY.RecordCount > 1 Then
                FRMEGRDTMP.Visible = True
                Fram.Enabled = False
                grdtmp.Columns(0).Visible = False
                grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
                grdtmp.Columns(1).Width = 4700
                'grdtmp.Columns(2).Visible = False
                grdtmp.Columns(17).Caption = "QTY"
                grdtmp.Columns(17).Width = 1300
                grdtmp.Columns(2).Visible = False
                grdtmp.Columns(3).Visible = False
                grdtmp.Columns(4).Visible = False
                grdtmp.Columns(5).Visible = False
                grdtmp.Columns(6).Visible = False
                grdtmp.Columns(7).Visible = False
                grdtmp.Columns(8).Visible = False
                grdtmp.Columns(9).Visible = False
                grdtmp.Columns(10).Visible = False
                grdtmp.Columns(11).Visible = False
                grdtmp.Columns(12).Visible = False
                grdtmp.Columns(13).Visible = False
                grdtmp.Columns(14).Visible = False
                grdtmp.Columns(15).Visible = False
                grdtmp.Columns(16).Visible = False
                grdtmp.Columns(18).Visible = False
                grdtmp.Columns(19).Visible = False
                grdtmp.Columns(20).Visible = False
                grdtmp.Columns(21).Visible = False
                grdtmp.Columns(22).Visible = False
                grdtmp.Columns(23).Visible = False
                grdtmp.Columns(24).Visible = False
                grdtmp.Columns(25).Visible = False
                grdtmp.Columns(26).Visible = False
                grdtmp.Columns(27).Visible = False
                grdtmp.Columns(28).Visible = False
                grdtmp.Columns(29).Visible = False
                grdtmp.Columns(30).Visible = False
                grdtmp.Columns(31).Visible = False
                grdtmp.Columns(32).Visible = False
                grdtmp.Columns(33).Visible = False
                grdtmp.Columns(34).Visible = False
                grdtmp.Columns(35).Visible = False
                grdtmp.Columns(36).Visible = False
                grdtmp.Columns(37).Visible = False
                grdtmp.Columns(38).Visible = False
                grdtmp.Columns(39).Visible = False
                grdtmp.Columns(40).Visible = False
                grdtmp.Columns(41).Visible = False
                grdtmp.Columns(42).Visible = False
                grdtmp.Columns(43).Visible = False
                grdtmp.Columns(44).Visible = False
                grdtmp.Columns(45).Visible = False
                grdtmp.Columns(46).Visible = False
                grdtmp.SetFocus
            End If
            
        Case vbKeyEscape
            txtcategory.Enabled = True
            'TXTPRODUCT.Enabled = False
            txtcategory.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number <> -2147168237 Then
        MsgBox Err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPRODUCT_LostFocus()
    Call CHANGEBOXCOLOR(TXTPRODUCT, False)
End Sub

Private Sub TXTPTR_GotFocus()
    Call CHANGEBOXCOLOR(TXTPTR, True)
    TXTPTR.SelStart = 0
    TXTPTR.SelLength = Len(TXTPTR.Text)
    Call FILL_PREVIIOUSRATE
    
    TXTRATE.Enabled = True
    TXTPTR.Enabled = True
    TxtNetrate.Enabled = True
    TxttaxMRP.Enabled = True
    TxtExDuty.Enabled = True
    TxtTrDisc.Enabled = True
    TxtCessPer.Enabled = True
    txtCess.Enabled = True
    TxtCSTper.Enabled = True
    txtPD.Enabled = True
    TxtExpense.Enabled = True
    TXTRETAIL.Enabled = True
    TxtRetailPercent.Enabled = True
    txtWS.Enabled = True
    txtWsalePercent.Enabled = True
    txtvanrate.Enabled = True
    txtSchPercent.Enabled = True
    txtcrtnpack.Enabled = True
    txtcrtn.Enabled = True
    TxtLWRate.Enabled = True
    TxtCustDisc.Enabled = True
    TxtComper.Enabled = True
    TxtComAmt.Enabled = True
    cmdadd.Enabled = True
    Txtgrossamt.Enabled = True
    txtBatch.Enabled = True
    txtHSN.Enabled = True
    TxtWarranty.Enabled = True
    CmbWrnty.Enabled = True
    TXTEXPDATE.Enabled = True
    End Sub

Private Sub TXTPTR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" And M_EDIT = True Then Exit Sub
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                Set GRDPRERATE.DataSource = Nothing
                fRMEPRERATE.Visible = False
                TXTPRODUCT.Enabled = True
                TXTPRODUCT.SetFocus
            Else
                TXTRATE.SetFocus
            End If
        Case 116
            Call FILL_PREVIIOUSRATE
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TXTPTR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPTR_LostFocus()
    Call CHANGEBOXCOLOR(TXTPTR, False)
    'tXTptrdummy.Text = Format(Val(TXTPTR.Text) / Val(TXTUNIT.Text), ".000")
    Txtgrossamt.Text = Val(TXTPTR.Text) * Val(TXTQTY.Text)
    TXTPTR.Text = Format(TXTPTR.Text, ".0000")
    'TxtNetrate.Text = Round(Val(TXTPTR.Text) + Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100, 4)
    Call TxttaxMRP_LostFocus
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            TxtRetailPercent.Text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtWsalePercent.Text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtSchPercent.Text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtSchPercent_LostFocus
    End If
    
    'TXTRETAIL.Text = Round(Val(txtmrpbt.Text) * 0.8, 2)
'    txtretail.Text = Format(Round(Val(TXTRATE.Text) - (Val(txtmrpbt.Text) * 20 / 100), 3), ".000")
'    txtprofit.Text = Format(Round(Val(txtretail.Text) - Val(txtretail.Text) * 10 / 100, 3), ".000")
End Sub

Private Sub TXTQTY_GotFocus()
    Call CHANGEBOXCOLOR(TXTQTY, True)
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
    If cmbfull.ListIndex = -1 Then cmbfull.Text = "Nos"
    If Val(Los_Pack.Text) = 1 Then CmbPack.Text = cmbfull.Text
    FRMEGRDTMP.Visible = False
    CmbPack.Enabled = True
    cmbfull.Enabled = True
    Los_Pack.Enabled = True
    TXTQTY.Enabled = True
    TxTfree.Enabled = True
    TXTRATE.Enabled = True
    TXTPTR.Enabled = True
    TxtNetrate.Enabled = True
    TxttaxMRP.Enabled = True
    TxtExDuty.Enabled = True
    TxtTrDisc.Enabled = True
    TxtCessPer.Enabled = True
    txtCess.Enabled = True
    TxtCSTper.Enabled = True
    txtPD.Enabled = True
    TxtExpense.Enabled = True
    TXTRETAIL.Enabled = True
    TxtRetailPercent.Enabled = True
    txtWS.Enabled = True
    txtWsalePercent.Enabled = True
    txtvanrate.Enabled = True
    txtSchPercent.Enabled = True
    txtcrtnpack.Enabled = True
    txtcrtn.Enabled = True
    TxtLWRate.Enabled = True
    TxtCustDisc.Enabled = True
    TxtComper.Enabled = True
    TxtComAmt.Enabled = True
    cmdadd.Enabled = True
    Txtgrossamt.Enabled = True
    txtBatch.Enabled = True
    txtHSN.Enabled = True
    TxtWarranty.Enabled = True
    CmbWrnty.Enabled = True
    TXTEXPIRY.Visible = False
    TXTEXPDATE.Enabled = True
    TxtBarcode.Enabled = False
    txtcategory.Enabled = False
    TXTPRODUCT.Enabled = False
    
    Dim rststock As ADODB.Recordset
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    With rststock
        If Not (.EOF And .BOF) Then
            txtHSN.Text = IIf(IsNull(rststock!Remarks), "", rststock!Remarks)
            TxtCustDisc.Text = IIf(IsNull(rststock!CUST_DISC), "", rststock!CUST_DISC)
            On Error Resume Next
            If cmbfull.ListIndex = -1 Then cmbfull.Text = IIf(IsNull(rststock!FULL_PACK), 0, rststock!FULL_PACK)
            On Error GoTo ErrHand
        Else
            txtHSN.Text = ""
            TxtCustDisc.Text = ""
            On Error Resume Next
            If cmbfull.ListIndex = -1 Then cmbfull.Text = CmbPack.Text
            On Error GoTo ErrHand
        End If
    End With
    rststock.Close
    Set rststock = Nothing
    
    If Trim(TxtBarcode.Text) = "" Then
        Set rststock = New ADODB.Recordset
        rststock.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
        If Not (rststock.EOF Or rststock.BOF) Then
            TxtBarcode.Text = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
        End If
        rststock.Close
        Set rststock = Nothing
    End If


    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            TxTfree.Enabled = True
            TxTfree.SetFocus
        Case vbKeyEscape
            cmbfull.Enabled = True
            cmbfull.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TXTQTY, False)
    TXTQTY.Text = Format(TXTQTY.Text, ".00")
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTPTR.Text), 2)), ".000")
    LblGross.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTPTR.Text), 2)), ".000")
    Call TXTPTR_LostFocus
End Sub

Private Sub TXTRATE_Change()
'    If Val(TXTRATE.Text) > 0 Then TXTRETAIL.Text = Val(TXTRATE.Text)

'    If val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTRATE.Text) > 0 Then
'        TXTRETAIL.Text = Val(TXTRATE.Text)
'    End If
'    If val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTRATE.Text) > 0 Then
'        txtWS.Text = Val(TXTRATE.Text)
'    End If
'    If val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTRATE.Text) > 0 Then
'        txtvanrate.Text = Val(TXTRATE.Text)
'    End If
End Sub

Private Sub TXTRATE_GotFocus()
    Call CHANGEBOXCOLOR(TXTRATE, True)
    TXTRATE.SelStart = 0
    TXTRATE.SelLength = Len(TXTRATE.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TXTRATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTRATE.Text) = 0 Then Exit Sub
            TXTPTR.SetFocus
         Case vbKeyEscape
            TxTfree.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TXTRATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRATE_LostFocus()
    Call CHANGEBOXCOLOR(TXTRATE, False)
    TXTRATE.Text = Format(TXTRATE.Text, ".000")
    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105 '(100 + Val(TxttaxMRP.Text))
End Sub

Private Sub txtremarks_GotFocus()
    Call CHANGEBOXCOLOR(TXTREMARKS, True)
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If txtBillNo.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Please select the supplier from the list", vbOKOnly, "EzBiz"
                TXTDEALER.SetFocus
                Exit Sub
            End If
            'If TXTINVOICE.Text = "" Then Exit Sub
            If Not IsDate(TXTINVDATE.Text) Then Exit Sub
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND PINV = '" & Trim(TXTINVOICE.Text) & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                MsgBox "You have already entered this Invoice number for " & Trim(DataList2.Text) & " as Computer Bill No. " & rstTRXMAST!VCH_NO, vbOKOnly, "EzBiz"
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                TXTINVOICE.SetFocus
                Exit Sub
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            CMBDISTRICT.SetFocus
        Case vbKeyEscape
            TXTINVDATE.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
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

Private Sub TXTREMARKS_LostFocus()
    Call CHANGEBOXCOLOR(TXTREMARKS, False)
End Sub

Private Sub TxtRetailPercent_GotFocus()
    Call CHANGEBOXCOLOR(TxtRetailPercent, True)
    TxtRetailPercent.SelStart = 0
    TxtRetailPercent.SelLength = Len(TxtRetailPercent.Text)
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.Text) <> 0 Then Label1(38).Caption = "Disc% on MRP"
End Sub

Private Sub TxtRetailPercent_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            txtWS.SetFocus
         Case vbKeyEscape
            TXTRETAIL.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtRetailPercent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtRetailPercent_LostFocus()
    Call CHANGEBOXCOLOR(TxtRetailPercent, False)
    On Error Resume Next
'    If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) + ((Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) + ((Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    Else
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    End If
''    If Val(TXTRATE.Text) = 0 Then
''        txtretail.Text = Round((Val(TXTPTR.Tag) * Val(TxtRetailPercent.Text) / 100) + Val(TXTPTR.Tag), 0)
''    Else
''        'txtretail.Text = Round(Val(TXTRATE.Text) / 1.12, 2) - (Round(Val(TXTRATE.Text) / 1.12, 2) * Val(TxtRetailPercent.Text) / 100)
''        txtretail.Text = Round(Val(TXTRATE.Text) * 100 / (Val(TxtRetailPercent.Text) + 100), 0)
''    End If
    
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.Text) <> 0 Then
        If Val(TxtRetailPercent.Text) > 0 Then
            TXTRETAIL.Text = Round(Val(TXTRATE.Text) - (Val(TXTRATE.Text) * Val(TxtRetailPercent.Text) / 100), 2)
        Else
            TXTRETAIL.Text = Val(TXTRATE.Text)
        End If
        Call TXTRETAIL_LostFocus
    Else
        If MDIMAIN.lblgst.Caption <> "R" Then
            If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
                TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption)) + ((Val(TxtExpense.Text)))), 4)
            Else
                TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
            End If
        Else
            If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
                TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.Text)))), 4)
            Else
                TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
            End If
        End If
        If MDIMAIN.lblgst.Caption <> "R" Then
            TXTRETAIL.Text = Round((Val(TXTPTR.Tag) * Val(TxtRetailPercent.Text) / 100) + Val(TXTPTR.Tag), 2)
        Else
            If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
                TXTRETAIL.Text = (Val(TXTPTR.Tag) * Val(TxtRetailPercent.Text) / 100) + Val(TXTPTR.Tag)
                TXTRETAIL.Text = Round(Val(TXTRETAIL.Text) + (Val(TXTRETAIL.Text) * (Val(TxttaxMRP.Text) + Val(TxtCessPer.Text)) / 100), 2)
                'TXTRETAIL.Tag = Round(Val(TXTRETAIL.Text) * 100 / ((Val(TxttaxMRP.Text)) + 100), 2)
            Else
                TXTRETAIL.Text = Round((Val(TXTPTR.Tag) * Val(TxtRetailPercent.Text) / 100) + Val(TXTPTR.Tag), 2)
            End If
        End If
    End If
    
    TXTRETAIL.Text = Format(Val(TXTRETAIL.Text), "0.00")
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub
'
Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                  Case 41  'BARCODE COUNT
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Trim(TXTsample.Text)
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
    MsgBox Err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case grdsales.Col
        Case 3, 5, 6, 7, 8, 9, 10
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 41
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub txtSchPercent_GotFocus()
    Call CHANGEBOXCOLOR(txtSchPercent, True)
    txtSchPercent.SelStart = 0
    txtSchPercent.SelLength = Len(txtSchPercent.Text)
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.Text) <> 0 Then Label1(38).Caption = "Disc% on MRP"
End Sub

Private Sub txtSchPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcrtnpack.SetFocus
         Case vbKeyEscape
            txtvanrate.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtSchPercent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSchPercent_LostFocus()
    Call CHANGEBOXCOLOR(txtSchPercent, False)
    On Error Resume Next
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.Text) <> 0 Then
        If Val(txtSchPercent.Text) > 0 Then
            txtvanrate.Text = Round(Val(TXTRATE.Text) - (Val(TXTRATE.Text) * Val(txtSchPercent.Text) / 100), 2)
        Else
            txtvanrate.Text = Val(TXTRATE.Text)
        End If
        Call txtvanrate_LostFocus
    Else
        If MDIMAIN.lblgst.Caption <> "R" Then
            If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
                TXTPTR.Tag = Val(LBLSUBTOTAL.Caption)
            Else
                TXTPTR.Tag = Round(Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.Text) + Val(TxTfree.Text), 4)
            End If
        Else
            If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
                TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.Text)))), 4)
            Else
                TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
            End If
        End If
        If MDIMAIN.lblgst.Caption <> "R" Then
            txtvanrate.Text = Round((Val(TXTPTR.Tag) * Val(txtSchPercent.Text) / 100) + Val(TXTPTR.Tag), 2)
        Else
            If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
                txtvanrate.Text = (Val(TXTPTR.Tag) * Val(txtSchPercent.Text) / 100) + Val(TXTPTR.Tag)
                txtvanrate.Text = Round(Val(txtvanrate.Text) + (Val(txtvanrate.Text) * (Val(TxttaxMRP.Text) + Val(TxtCessPer.Text)) / 100), 2)
            Else
                txtvanrate.Text = Round((Val(TXTPTR.Tag) * Val(txtSchPercent.Text) / 100) + Val(TXTPTR.Tag), 2)
            End If
        End If
    End If
    
    txtvanrate.Text = Format(Val(txtvanrate.Text), "0.00")
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub TXTSLNO_GotFocus()
    Call CHANGEBOXCOLOR(TXTSLNO, True)
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    TxtBarcode.Enabled = False
    txtcategory.Enabled = False
    TXTPRODUCT.Enabled = False
    
    Los_Pack.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxttaxMRP.Enabled = False
    TxtExDuty.Enabled = False
    TxtTrDisc.Enabled = False
    TxtCessPer.Enabled = False
    txtCess.Enabled = False
    TxtCSTper.Enabled = False
    txtPD.Enabled = False
    TxtExpense.Enabled = False
    TXTRETAIL.Enabled = False
    TxtRetailPercent.Enabled = False
    
    txtWS.Enabled = False
    txtWsalePercent.Enabled = False
    txtvanrate.Enabled = False
    txtSchPercent.Enabled = False
    txtcrtnpack.Enabled = False
    txtcrtn.Enabled = False
    TxtLWRate.Enabled = False
    TxtCustDisc.Enabled = False
    TxtComper.Enabled = False
    TxtComAmt.Enabled = False
    cmdadd.Enabled = False
    Txtgrossamt.Enabled = False
    txtBatch.Enabled = False
    txtHSN.Enabled = False
    TxtWarranty.Enabled = False
    CmbWrnty.Enabled = False
    TXTEXPIRY.Visible = False
    TXTEXPDATE.Enabled = False
    
    
    BARCODE_FLAG = False
    Set grdtmp.DataSource = Nothing
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.Text) = 0 Then
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
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 1)
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))
                TXTUNIT.Text = 1 'grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                Txtpack.Text = 1 'grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                'TXTRATE.Text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)), 2), "0.000")
                TXTRATE.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)), "0.000")
                TXTPTR.Text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 9)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)), 4), "0.0000")
                txtprofit.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7)), "0.00")
                TXTRETAIL.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)), "0.00")
                LBLPRE.Caption = Val(TXTRETAIL.Text)
                txtWS.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)), "0.00")
                txtvanrate.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)), "0.00")
                Txtgrossamt.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 26)), "0.00")
                txtcrtn.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)), "0.00")
                TxtLWRate.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37)), "0.00")
                txtcrtnpack.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)), "0.00")
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23)) = "A" Then
                    OptComAmt.value = True
                    TxtComper.Text = ""
                    TxtComAmt.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22)), "0.00")
                Else
                    OptComper.value = True
                    TxtComper.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21)), "0.00")
                    TxtComAmt.Text = ""
                End If
                
                'TXTPTR.Text = Format((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))) * Val(Los_Pack.Text), "0.000")

                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 11)
                TXTEXPDATE.Text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 12)), grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "  /  /    ")
                TXTEXPIRY.Text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 12)), Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "mm/yy"), "  /  ")
                'LBLSUBTOTAL.Caption = Format(Val(TXTQTY.Text) * (Val(TXTPTR.Text) + Val(lbltaxamount.Caption)), ".000")
                If Optdiscamt.value = True Then
                    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption) - Val(txtPD.Text), ".000")
                    LblGross.Caption = Format(Val(Txtgrossamt.Text) - Val(txtPD.Text), ".000")
                Else
                    LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100), ".000")
                    LblGross.Caption = Format(Val(Txtgrossamt.Text) - (Val(Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100)), ".000")
                End If
                TxTfree.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TxttaxMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105 '(100 + Val(TxttaxMRP.Text))
                txtPD.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) = "V" Then
                    OPTVAT.value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) = "M" Then
                    OPTTaxMRP.value = True
                Else
                    OPTNET.value = True
                End If
                
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 27)) = "P" Then
                    optdiscper.value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 27)) = "A" Then
                    Optdiscamt.value = True
                End If
                On Error Resume Next
                Los_Pack.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 28))
                CmbPack.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 29)
                TxtWarranty.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 30)
                CmbWrnty.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 31)
                TxtExpense.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32))
                TxtExDuty.Text = "" 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 33))
                TxtCSTper.Text = "" 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 34))
                TxtTrDisc.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 35))
                LblGross.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 36))
                TxtBarcode.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 38))
                txtCess.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39))
                TxtCessPer.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40))
                TxtNetrate.Text = Round(Val(TXTPTR.Text) + Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100, 4)
                FRMEGRDTMP.Visible = False
                Err.Clear
                
                On Error GoTo ErrHand
                Dim rststock As ADODB.Recordset
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                With rststock
                    If Not (.EOF And .BOF) Then
                        lblcategory.Caption = IIf(IsNull(rststock!Category), "", rststock!Category)
                        On Error Resume Next
                        cmbfull.Text = IIf(IsNull(rststock!FULL_PACK), 0, rststock!FULL_PACK)
                        Err.Clear
                        On Error GoTo ErrHand
                    Else
                        lblcategory.Caption = ""
                    End If
                End With
                rststock.Close
                Set rststock = Nothing
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTITEMCODE.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            'TXTPRODUCT.Enabled = False
            TxtBarcode.Enabled = True
            TxtBarcode.SetFocus
            Exit Sub
            txtcategory.Enabled = True
            txtcategory.SetFocus
            'TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TxtBarcode.Text = ""
                TXTQTY.Text = ""
                Txtpack.Text = 1 '""
                Los_Pack.Text = ""
                CmbPack.ListIndex = -1
                TxtWarranty.Text = ""
                CmbWrnty.ListIndex = -1
                TxTfree.Text = ""
                TxttaxMRP.Text = ""
                TxtExDuty.Text = ""
                TxtCSTper.Text = ""
                TxtTrDisc.Text = ""
                TxtCustDisc.Text = ""
                TxtCessPer.Text = ""
                txtCess.Text = ""
                txtPD.Text = ""
                TxtExpense.Text = ""
                txtprofit.Text = ""
                TXTRETAIL.Text = ""
                TxtRetailPercent.Text = ""
                
                txtWsalePercent.Text = ""
                txtSchPercent.Text = ""
                txtWS.Text = ""
                txtvanrate.Text = ""
                Txtgrossamt.Text = ""
                txtcrtn.Text = ""
                TxtLWRate.Text = ""
                txtcrtnpack.Text = ""
                OptComper.value = True
                TXTRATE.Text = ""
                TxtComAmt.Text = ""
                TxtComper.Text = ""
                txtmrpbt.Text = ""
                LBLSUBTOTAL.Caption = ""
                LblGross.Caption = ""
                lbltaxamount.Caption = ""
                lblcategory.Caption = ""
                TXTEXPDATE.Text = "  /  /    "
                TXTEXPIRY.Text = "  /  "
                txtBatch.Text = ""
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            Else
                cmdRefresh.Enabled = True
                cmdRefresh.SetFocus
            End If
'''            If M_ADD = False Then
'''                FRMECONTROLS.Enabled = False
'''                FRMEMASTER.Enabled = False
'''                cmdRefresh.Enabled = False
'''                txtBillNo.Enabled = True
'''                txtBillNo.SetFocus
'''                Exit Sub
'''            End If
            
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
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

Private Sub TXTEXPIRY_GotFocus()
    'Call CHANGEBOXCOLOR(TXTEXPIRY, True)
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
Dim M_DATE As Date
Dim D As Integer
Dim M As Integer
Dim Y As Integer
    Select Case KeyCode
        Case vbKeyReturn
            If Len(Trim(TXTEXPIRY.Text)) = 1 Then GoTo SKIP
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
            
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then
                TXTEXPDATE.Text = "  /  /    "
                Exit Sub
            End If
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then
                TXTEXPDATE.Text = "  /  /    "
                Exit Sub
            End If
            
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then
                TXTEXPDATE.Text = "  /  /    "
                Exit Sub
            End If
            
            M = Val(Mid(TXTEXPIRY.Text, 1, 2))
            Y = Val(Right(TXTEXPIRY.Text, 2))
            Y = 2000 + Y
            M_DATE = "01" & "/" & M & "/" & Y
            D = LastDayOfMonth(M_DATE)
            M_DATE = D & "/" & M & "/" & Y
            TXTEXPDATE.Text = Format(M_DATE, "dd/mm/yyyy")
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 0 Then
                MsgBox "Item Expired....", vbOKOnly, "EzBiz"
                TXTEXPDATE.Text = "  /  /    "
                TXTEXPIRY.SelStart = 0
                TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
                TXTEXPIRY.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 60 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days", vbOKOnly, "EzBiz"
                TXTEXPDATE.Text = "  /  /    "
                TXTEXPIRY.SelStart = 0
                TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
                TXTEXPIRY.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 180 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "EzBiz") = vbNo Then
                    TXTEXPDATE.Text = "  /  /    "
                    TXTEXPIRY.SelStart = 0
                    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
                    TXTEXPIRY.SetFocus
                    Exit Sub
                End If
            End If
SKIP:
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            txtBatch.Enabled = True
            TXTEXPDATE.Enabled = False
            txtBatch.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    'Call CHANGEBOXCOLOR(TXTEXPIRY, False)
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
    TXTEXPIRY.Visible = False
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Sub TXTSLNO_LostFocus()
    Call CHANGEBOXCOLOR(TXTSLNO, False)
End Sub

Private Sub TxttaxMRP_GotFocus()
    Call CHANGEBOXCOLOR(TxttaxMRP, True)
    TxttaxMRP.SelStart = 0
    TxttaxMRP.SelLength = Len(TxttaxMRP.Text)
End Sub

Private Sub TxttaxMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxttaxMRP.Text) <> 0 And OPTNET.value = True Then
                OPTVAT.value = True
                OPTVAT.SetFocus
                Exit Sub
            End If
            If Trim(txtHSN.Text) = "" Then
                txtHSN.Enabled = True
                txtHSN.SetFocus
            Else
                txtPD.SetFocus
            End If
         Case vbKeyEscape
            TXTEXPDATE.Enabled = True
            TXTEXPDATE.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TxttaxMRP, False)
    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / (100 + Val(TxttaxMRP.Text))
    Txtgrossamt.Text = Val(TXTPTR.Text) * Val(TXTQTY.Text)
    Txtgrossamt.Tag = Val(Txtgrossamt.Text) + (Val(Txtgrossamt.Text) * Val(TxtExDuty.Text) / 100)
    Txtgrossamt.Tag = Val(Txtgrossamt.Tag) + (Val(Txtgrossamt.Text) * Val(TxtCSTper.Text) / 100)
    'Txtgrossamt.Tag = Val(Txtgrossamt.Tag) + Val(txtCess.Text)
'    If Val(TxttaxMRP.Text) = 0 Then
'
'        TxttaxMRP.Text = 0
'        lbltaxamount.Caption = 0
'        lbltaxamount.Caption = ""
'        If optdiscper.value = True Then
'            LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag)) - Val(Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)
'            LblGross.Caption = (Val(Txtgrossamt.Tag)) - Val(Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)
'        Else
'            LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) - Val(TxtPD.Text))
'            LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(TxtPD.Text))
'        End If
'    Else
'        If OPTTaxMRP.value = True Then
'            lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100
'            If optdiscper.value = True Then
'                LBLSUBTOTAL.Caption = (Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption)
'                LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - (Val(LBLSUBTOTAL.Caption) * Val(TxtPD.Text) / 100)
'            Else
'                LBLSUBTOTAL.Caption = (Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption)
'                LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - Val(TxtPD.Text)
'            End If
'            LblGross.Caption = LBLSUBTOTAL.Caption
        If OPTVAT.value = True Then
           If optdiscper.value = True Then
                lbltaxamount.Tag = (Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(txtPD.Text) / 100))
                lbltaxamount.Caption = Round((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                'lbltaxamount.Caption = Round((Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                'LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)
                'LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - (Val(LBLSUBTOTAL.Caption) - Val(lbltaxamount.Caption)) * Val(TxtTrDisc.Text) / 100
                LBLSUBTOTAL.Caption = Round(((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.Text) / 100))) + Val(lbltaxamount.Caption), 2)
                LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.Text) / 100))
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.Text) / 100)
            Else
                lbltaxamount.Tag = Val(Txtgrossamt.Tag) - Val(txtPD.Text)
                lbltaxamount.Caption = Round((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                'LBLSUBTOTAL.Caption = Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption) - Val(TxtPD.Text)
                LBLSUBTOTAL.Caption = Round(((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.Text) / 100))) + Val(lbltaxamount.Caption), 2)
                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.Text)
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.Text) / 100)
            End If
            LBLSUBTOTAL.Caption = LBLSUBTOTAL.Caption + (Val(LblGross.Caption) * Val(TxtCessPer.Text) / 100)
            LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) + (Val(txtCess.Text) * Val(TXTQTY.Text))
        Else
            TxttaxMRP.Text = 0
            If optdiscper.value = True Then
                lbltaxamount.Tag = (Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(txtPD.Text) / 100))
                lbltaxamount.Caption = Round((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                'lbltaxamount.Caption = Round((Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                'LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)
                'LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - (Val(LBLSUBTOTAL.Caption) - Val(lbltaxamount.Caption)) * Val(TxtTrDisc.Text) / 100
                LBLSUBTOTAL.Caption = Round(((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.Text) / 100))) + Val(lbltaxamount.Caption), 2)
                LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.Text) / 100))
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.Text) / 100)
            Else
                lbltaxamount.Tag = Val(Txtgrossamt.Tag) - Val(txtPD.Text)
                lbltaxamount.Caption = Round((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                'LBLSUBTOTAL.Caption = Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption) - Val(TxtPD.Text)
                LBLSUBTOTAL.Caption = Round(((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.Text) / 100))) + Val(lbltaxamount.Caption), 2)
                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.Text)
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.Text) / 100)
            End If
            LBLSUBTOTAL.Caption = LBLSUBTOTAL.Caption + (Val(LblGross.Caption) * Val(TxtCessPer.Text) / 100)
            LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) + (Val(txtCess.Text) * Val(TXTQTY.Text))
        End If
'    End If
    'LBLSUBTOTAL.Caption = Round(Val(LBLSUBTOTAL.Caption) + Val(txtCess.Text), 2)
    'TxtNetrate.Text = Round(Val(TXTPTR.Text) + Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100, 4)
    LBLSUBTOTAL.Caption = Format(Round(LBLSUBTOTAL.Caption, 3), "0.00")
    LblGross.Caption = Format(LblGross.Caption, "0.00")
    TxttaxMRP.Text = Format(TxttaxMRP.Text, "0.00")
    lbltaxamount.Caption = Format(lbltaxamount.Caption, "0.00")
    If Val(TXTQTY.Text) > 0 Then TxtNetrate.Text = Val(LBLSUBTOTAL) / Val(TXTQTY.Text)
    
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            TxtRetailPercent.Text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtWsalePercent.Text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtSchPercent.Text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtSchPercent_LostFocus
    End If
    
End Sub

Private Sub TxtTotalexp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            Call find_small_number
    End Select
End Sub

Private Sub TxtTotalexp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtTrDisc_LostFocus()
    Call CHANGEBOXCOLOR(TxtTrDisc, False)
    Call TxttaxMRP_LostFocus
    'If Val(TXTQTY.Text) <> 0 Then TxtNetrate.Text = Val(LBLSUBTOTAL) / Val(TXTQTY.Text)
End Sub

Private Sub TXTUNIT_GotFocus()
    Call CHANGEBOXCOLOR(TXTUNIT, True)
    TXTUNIT.SelStart = 0
    TXTUNIT.SelLength = Len(TXTUNIT.Text)
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.Text) = 0 Then Exit Sub
            
            TXTUNIT.Enabled = False
            Txtpack.Enabled = True
            Txtpack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTQTY.Text = ""
            TxTfree.Text = ""
            TxttaxMRP.Text = ""
            TxtExDuty.Text = ""
            TxtCSTper.Text = ""
            TxtTrDisc.Text = ""
            TxtCustDisc.Text = ""
            TxtCessPer.Text = ""
            txtCess.Text = ""
            txtprofit.Text = ""
            TXTRETAIL.Text = ""
            TxtRetailPercent.Text = ""
            
            txtWsalePercent.Text = ""
            txtSchPercent.Text = ""
            txtWS.Text = ""
            txtvanrate.Text = ""
            Txtgrossamt.Text = ""
            txtcrtn.Text = ""
            TxtLWRate.Text = ""
            txtcrtnpack.Text = ""
            txtPD.Text = ""
            TxtExpense.Text = ""
            txtBatch.Text = ""
            TXTRATE.Text = ""
            txtmrpbt.Text = ""
            TXTPTR.Text = ""
            TxtNetrate.Text = ""
            Txtgrossamt.Text = ""
            TXTEXPDATE.Text = "  /  /    "
            TXTEXPIRY.Text = "  /  "
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            lbltaxamount.Caption = ""
            lblcategory.Caption = ""
            TXTPRODUCT.Enabled = True
            TXTUNIT.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTUNIT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDISCAMOUNT_LostFocus()
    Call CHANGEBOXCOLOR(TXTDISCAMOUNT, False)
    Dim DISC As Currency
    
    On Error GoTo ErrHand
    If (TXTDISCAMOUNT.Text = "") Then
        DISC = 0
    Else
        DISC = TXTDISCAMOUNT.Text
    End If
    If grdsales.Rows = 1 Then
        TXTDISCAMOUNT.Text = "0"
    ElseIf Val(TXTDISCAMOUNT.Text) > Val(lbltotalwodiscount.Caption) Then
'        MsgBox "Discount Amount More than Bill Amount", , "PURCHASE..."
'        TXTDISCAMOUNT.SelStart = 0
'        TXTDISCAMOUNT.SelLength = Len(TXTDISCAMOUNT.Text)
'        TXTDISCAMOUNT.SetFocus
'        Exit Sub
    End If
    TXTDISCAMOUNT.Text = Format(TXTDISCAMOUNT.Text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    If Roundflag = True Then
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
            Else
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 2), "0.00")
            End If
    ''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    TXTDISCAMOUNT.SetFocus
End Sub

Private Sub TXTDISCAMOUNT_GotFocus()
    Call CHANGEBOXCOLOR(TXTDISCAMOUNT, True)
    TXTDISCAMOUNT.SelStart = 0
    TXTDISCAMOUNT.SelLength = Len(TXTDISCAMOUNT.Text)
End Sub

Private Sub TXTDISCAMOUNT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDISCAMOUNT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            'If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            'If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            'If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTRETAIL.Enabled = True Then TXTRETAIL.SetFocus
            If txtWS.Enabled = True Then txtWS.SetFocus
            If txtcrtn.Enabled = True Then txtcrtn.SetFocus
            If TxtLWRate.Enabled = True Then TxtLWRate.SetFocus
            If txtcrtnpack.Enabled = True Then txtcrtnpack.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Public Sub appendpurchase()
    
    Dim rstMaxRec As ADODB.Recordset
    Dim RSTLINK As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    Dim M_DATA As Double
    Dim I As Long
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    
    'If OLD_BILL = False Then Call checklastbill
    Set RSTTRXFILE = New ADODB.Recordset
    If OLD_BILL = False And Val(txtBillNo.Text) <> 1 Then
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE VCH_NO= (SELECT MAX(VCH_NO) FROM TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'PI')", db, adOpenStatic, adLockOptimistic, adCmdText
        txtBillNo.Text = RSTTRXFILE!VCH_NO + 1
        db.BeginTrans
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "PI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = txtBillNo.Text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    Else
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "PI"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
            RSTTRXFILE!VCH_NO = txtBillNo.Text
            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        End If
    End If
    If Not IsDate(TXTDATE.Text) Then TXTDATE.Text = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!CREATE_DATE = Format(TXTDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    If IsDate(TXTRCVDATE.Text) Then
        RSTTRXFILE!RCVD_DATE = Format(TXTRCVDATE.Text, "DD/MM/YYYY")
    End If
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = Trim(DataList2.Text)
    RSTTRXFILE!VCH_AMOUNT = Val(lbltotalwodiscount.Caption)
    RSTTRXFILE!NET_AMOUNT = Val(LBLTOTAL.Caption)
    RSTTRXFILE!DISCOUNT = Val(TXTDISCAMOUNT.Text)
    RSTTRXFILE!ADD_AMOUNT = Val(txtaddlamt.Text)
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!OPEN_PAY = 0
    RSTTRXFILE!PAY_AMOUNT = 0
    RSTTRXFILE!REF_NO = ""
    RSTTRXFILE!SLSM_CODE = "CS"
    RSTTRXFILE!CHECK_FLAG = "N"
    'If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
    RSTTRXFILE!POST_FLAG = "N"
    RSTTRXFILE!CFORM_NO = ""
    RSTTRXFILE!CFORM_DATE = Date
    RSTTRXFILE!Remarks = Trim(TXTREMARKS.Text)
    RSTTRXFILE!DISC_PERS = Val(txtcramt.Text)
    RSTTRXFILE!CST_PER = Val(TxtCST.Text)
    RSTTRXFILE!INS_PER = Val(TxtInsurance.Text)
    RSTTRXFILE!LETTER_NO = 0
    RSTTRXFILE!LETTER_DATE = Date
    RSTTRXFILE!INV_MSGS = ""
    RSTTRXFILE!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!PINV = Trim(TXTINVOICE.Text)
    RSTTRXFILE!TRX_GODOWN = Trim(CMBDISTRICT.Text)
    RSTTRXFILE.Update
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    'db.Execute "delete From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    'db.Execute "delete FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'CR' AND INV_TRX_TYPE = 'PI'"
    db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'PY' AND INV_TRX_TYPE = 'PI' "
    db.Execute "delete FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'CR' AND INV_TRX_TYPE = 'PI'"
    
    'db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'PI'"
    If grdsales.Rows = 1 Then GoTo SKIP
            
    I = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'PY' AND '" & Year(MDIMAIN.DTFROM.value) & "'", db, adOpenForwardOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        I = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    Dim BillNO As Long
    BillNO = 1
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(REC_NO) From DBTPYMT WHERE TRX_TYPE = 'PY' AND '" & Year(MDIMAIN.DTFROM.value) & "'", db, adOpenForwardOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        BillNO = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    'If lblcredit.Caption = "1" Then
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'PY' AND INV_TRX_TYPE = 'PI'", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!TRX_TYPE = "PY"
            RSTITEMMAST!INV_TRX_TYPE = "PI"
            RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        End If
        RSTITEMMAST!CR_NO = I
        RSTITEMMAST!REC_NO = BillNO
        RSTITEMMAST!RCPT_AMT = Val(LBLTOTAL.Caption)
        RSTITEMMAST!RCPT_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = DataList2.Text
        RSTITEMMAST!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTITEMMAST!INV_AMT = 0
        'RSTITEMMAST!INV_NO = 0
        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
        RSTITEMMAST!BANK_FLAG = "N"
        RSTITEMMAST!B_TRX_TYPE = Null
        'RSTITEMMAST!B_TRX_NO = Null
        RSTITEMMAST!B_BILL_TRX_TYPE = Null
        RSTITEMMAST!B_TRX_YEAR = Null
        RSTITEMMAST!BANK_CODE = Null
        RSTITEMMAST!BANK_NAME = ""
        RSTITEMMAST!C_TRX_TYPE = Null
        'RSTITEMMAST!C_REC_NO = Null
        RSTITEMMAST!C_INV_TRX_TYPE = Null
        RSTITEMMAST!C_INV_TYPE = Null
        'RSTITEMMAST!RCPT_AMOUNT = 0
        'RSTITEMMAST!CHECK_FLAG = "N"
        'RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption)
        'RSTITEMMAST!PINV = Trim(TXTINVOICE.Text)
        
        RSTITEMMAST.Update
        db.CommitTrans
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    
    
    'End If
        
'    For i = 1 To grdsales.Rows - 1
'        Set RSTLINK = New ADODB.Recordset
'        RSTLINK.Open "SELECT * FROM PRODLINK WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "' AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        If (RSTLINK.EOF And RSTLINK.BOF) Then
'            RSTLINK.AddNew
'            RSTLINK!ITEM_CODE = grdsales.TextMatrix(i, 1)
'            RSTLINK!ITEM_NAME = grdsales.TextMatrix(i, 2)
'            RSTLINK!RQTY = grdsales.TextMatrix(i, 3)
'            RSTLINK!ITEM_COST = grdsales.TextMatrix(i, 8)
'            RSTLINK!MRP = grdsales.TextMatrix(i, 6)
'            RSTLINK!PTR = grdsales.TextMatrix(i, 9)
'            RSTLINK!SALES_PRICE = grdsales.TextMatrix(i, 7)
'            RSTLINK!SALES_TAX = Val(grdsales.TextMatrix(i, 10))
'            RSTLINK!UNIT = grdsales.TextMatrix(i, 4)
'            RSTLINK!Remarks = grdsales.TextMatrix(i, 4)
'            RSTLINK!ORD_QTY = 0
'            RSTLINK!CST = 0
'            RSTLINK!ACT_CODE = DataList2.BoundText
'            RSTLINK!CREATE_DATE = Format(Date, "dd/mm/yyyy")
'            RSTLINK!C_USER_ID = ""
'            RSTLINK!MODIFY_DATE = Format(Date, "dd/mm/yyyy")
'            RSTLINK!M_USER_ID = ""
'            RSTLINK!CHECK_FLAG = "Y"
'            RSTLINK!SITEM_CODE = ""
'
'            RSTLINK.Update
'        End If
'        RSTLINK.Close
'        Set RSTLINK = Nothing
'    Next i
'
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(Trim(TXTINVDATE.Text), "dd/mm/yyyy")
        RSTTRXFILE!VCH_DESC = "Received From " & DataList2.Text
        RSTTRXFILE!PINV = Trim(TXTINVOICE.Text)
        RSTTRXFILE!TRX_GODOWN = Trim(CMBDISTRICT.Text)
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        If CMBPO.Text <> "" Then
            RSTTRXFILE!PO_NO = IIf(CMBPO.Text = "", Null, CMBPO.Text)
        End If
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
'    If lblcredit.Caption = "0" Then
'        i = 0
'        Set rstMaxRec = New ADODB.Recordset
'        rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
'        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
'            i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
'        End If
'        rstMaxRec.Close
'        Set rstMaxRec = Nothing
'
'        Set RSTITEMMAST = New ADODB.Recordset
'        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'PI'", db, adOpenStatic, adLockOptimistic, adCmdText
'        db.BeginTrans
'        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'            RSTITEMMAST.AddNew
'            RSTITEMMAST!rec_no = i + 1
'            RSTITEMMAST!INV_TYPE = "PY"
'            RSTITEMMAST!INV_TRX_TYPE = "PI"
'            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'            RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
'        End If
'        If lblcredit.Caption = "0" Then
'            RSTITEMMAST!TRX_TYPE = "DR"
'        Else
'            RSTITEMMAST!TRX_TYPE = "CR"
'        End If
'        RSTITEMMAST!TRX_TYPE = "CR"
'        RSTITEMMAST!act_code = DataList2.BoundText
'        RSTITEMMAST!act_name = Trim(DataList2.Text)
'        RSTITEMMAST!AMOUNT = Val(LBLTOTAL.Caption)
'        RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
'        RSTITEMMAST!CHECK_FLAG = "P"
'        RSTITEMMAST.Update
'        db.CommitTrans
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
'    End If

SKIP:
    
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    cmdRefresh.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.Text = TXTLASTBILL.Text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTRCVDATE.Text = "  /  /    "
    TXTINVOICE.Text = ""
    CMBDISTRICT.Text = ""
    lbladdress.Caption = ""
    TXTREMARKS.Text = ""
    TXTSLNO.Text = ""
    TXTITEMCODE.Text = ""
    TxtBarcode.Text = ""
    TXTPRODUCT.Text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.Text = ""
    TxttaxMRP.Text = ""
    TxtExDuty.Text = ""
    TxtCSTper.Text = ""
    TxtTrDisc.Text = ""
    TxtCustDisc.Text = ""
    TxtCessPer.Text = ""
    txtCess.Text = ""
    txtPD.Text = ""
    TxtExpense.Text = ""
    txtprofit.Text = ""
    TXTRETAIL.Text = ""
    TxtRetailPercent.Text = ""
    
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    TxtLWRate.Text = ""
    txtcrtnpack.Text = ""
    txtBatch.Text = ""
    txtHSN.Text = ""
    TXTRATE.Text = ""
    txtmrpbt.Text = ""
    TXTPTR.Text = ""
    TxtNetrate.Text = ""
    Txtgrossamt.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    lblcategory.Caption = ""
    LBLPRE.Caption = ""
    txtaddlamt.Text = ""
    txtcramt.Text = ""
    TxtInsurance.Text = ""
    TxtCST.Text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    TXTDISCAMOUNT.Text = ""
    TxtTotalexp.Text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    grdsales.Rows = 1
    CMDEXIT.Enabled = True
    OptComper.value = True
    txtBillNo.SetFocus
    M_ADD = False
    M_EDIT = False
    OLD_BILL = False
    LBLmonth.Caption = "0.00"
    Chkcancel.value = 0
    Call CLEAR_COMBO
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "EzBiz"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number = 7 Then
        MsgBox "Select Supplier from the list", vbOKOnly, "EzBiz"
    Else
        Screen.MousePointer = vbNormal
        If Err.Number <> -2147168237 Then
            MsgBox Err.Description
        End If
        On Error Resume Next
        db.RollbackTrans
    End If
End Sub


Private Sub txtaddlamt_GotFocus()
    Call CHANGEBOXCOLOR(txtaddlamt, True)
    txtaddlamt.SelStart = 0
    txtaddlamt.SelLength = Len(txtaddlamt.Text)
End Sub

Private Sub txtaddlamt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtaddlamt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            'If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            'If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            'If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTRETAIL.Enabled = True Then TXTRETAIL.SetFocus
            If txtWS.Enabled = True Then txtWS.SetFocus
            If txtcrtn.Enabled = True Then txtcrtn.SetFocus
            If TxtLWRate.Enabled = True Then TxtLWRate.SetFocus
            If txtcrtnpack.Enabled = True Then txtcrtnpack.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub txtaddlamt_LostFocus()
    Call CHANGEBOXCOLOR(txtaddlamt, False)
    Dim DISC As Currency
    
    On Error GoTo ErrHand
    If (txtaddlamt.Text = "") Then
        DISC = 0
    Else
        DISC = txtaddlamt.Text
    End If
    If grdsales.Rows = 1 Then
        txtaddlamt.Text = "0"
    ElseIf Val(txtaddlamt.Text) > Val(lbltotalwodiscount.Caption) Then
        MsgBox "Discount Amount More than Bill Amount", , "PURCHASE..."
        txtaddlamt.SelStart = 0
        txtaddlamt.SelLength = Len(txtaddlamt.Text)
        txtaddlamt.SetFocus
        Exit Sub
    End If
    txtaddlamt.Text = Format(txtaddlamt.Text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    If Roundflag = True Then
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    Else
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 2), "0.00")
    End If
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    txtaddlamt.SetFocus
End Sub

Private Sub txtcramt_GotFocus()
    Call CHANGEBOXCOLOR(txtcramt, True)
    txtcramt.SelStart = 0
    txtcramt.SelLength = Len(txtcramt.Text)
End Sub

Private Sub txtcramt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcramt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            'If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            'If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            'If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTRETAIL.Enabled = True Then TXTRETAIL.SetFocus
            If txtWS.Enabled = True Then txtWS.SetFocus
            If txtcrtn.Enabled = True Then txtcrtn.SetFocus
            If TxtLWRate.Enabled = True Then TxtLWRate.SetFocus
            If txtcrtnpack.Enabled = True Then txtcrtnpack.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub txtcramt_LostFocus()
    Call CHANGEBOXCOLOR(txtcramt, False)
    Dim DISC As Currency
    
    On Error GoTo ErrHand
    If (txtcramt.Text = "") Then
        DISC = 0
    Else
        DISC = txtcramt.Text
    End If
    If grdsales.Rows = 1 Then
        txtcramt.Text = "0"
    ElseIf Val(txtcramt.Text) > Val(lbltotalwodiscount.Caption) Then
        MsgBox "Credit Note Amount More than Bill Amount", , "PURCHASE..."
        txtcramt.SelStart = 0
        txtcramt.SelLength = Len(txtcramt.Text)
        txtcramt.SetFocus
        Exit Sub
    End If
    txtcramt.Text = Format(txtcramt.Text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 2), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    txtcramt.SetFocus
End Sub

Private Sub OPTTaxMRP_GotFocus()
    OPTTaxMRP.BackColor = &H98F3C1
    'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100
    'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
    lbltaxamount.Caption = ((Val(TXTRATE.Text) * (Val(TXTQTY.Text) + Val(TxTfree.Text)) * 55 / 100)) * Val(TxttaxMRP.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
    LblGross.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)), ".000")
            
'    If optdiscper.Value = True Then
'        lbltaxamount.Caption = Round((Val(Txtgrossamt.Text) - (Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
'        LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100), ".000")
'    Else
'        lbltaxamount.Caption = Round((Val(Txtgrossamt.Text) - Val(txtPD.Text)) * Val(TxttaxMRP.Text) / 100, 2)
'        LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption) - Val(txtPD.Text), ".000")
'    End If
End Sub

Private Sub OPTVAT_GotFocus()
    OPTVAT.BackColor = &H98F3C1
    'lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text) + Val(TxtFree.Text))
    If optdiscper.value = True Then
        lbltaxamount.Caption = Round((Val(Txtgrossamt.Text) - (Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
        LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100), ".000")
        LblGross.Caption = Format(Val(Txtgrossamt.Text) - Val(Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100), ".000")
    Else
        lbltaxamount.Caption = Round((Val(Txtgrossamt.Text) - Val(txtPD.Text)) * Val(TxttaxMRP.Text) / 100, 2)
        LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption) - Val(txtPD.Text), ".000")
        LblGross.Caption = Format(Val(Txtgrossamt.Text) - Val(txtPD.Text), ".000")
    End If
End Sub

Private Sub OPTNET_GotFocus()
    OPTNET.BackColor = &H98F3C1
    lbltaxamount.Caption = ""
    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text), ".000")
    LblGross.Caption = Format(Val(Txtgrossamt.Text), ".000")
End Sub

Private Sub txtprofit_GotFocus()
    Call CHANGEBOXCOLOR(txtprofit, True)
    txtprofit.SelStart = 0
    txtprofit.SelLength = Len(txtprofit.Text)
End Sub

Private Sub txtprofit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtprofit.Enabled = False
            TXTRETAIL.Enabled = True
            TXTRETAIL.SetFocus
         Case vbKeyEscape
            txtprofit.Enabled = False
            txtPD.Enabled = True
            txtPD.SetFocus
    End Select
End Sub

Private Sub txtprofit_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtprofit_LostFocus()
    Call CHANGEBOXCOLOR(txtprofit, False)
    txtprofit.Text = Format(txtprofit.Text, "0.00")
End Sub

Private Sub txtPD_GotFocus()
    Call CHANGEBOXCOLOR(txtPD, True)
    txtPD.SelStart = 0
    txtPD.SelLength = Len(txtPD.Text)
End Sub

Private Sub txtPD_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
'            txtPD.Enabled = False
'            cmdadd.Enabled = True
'            cmdadd.SetFocus
'            Exit Sub
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                Call CMDADD_Click
            Else
                TxtTrDisc.SetFocus
            End If
         Case vbKeyEscape
            TxttaxMRP.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtPD_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPD_LostFocus()
    Call CHANGEBOXCOLOR(txtPD, False)
    Call TxttaxMRP_LostFocus
    
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            TxtRetailPercent.Text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtWsalePercent.Text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtSchPercent.Text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtSchPercent_LostFocus
    End If
    'If Val(TXTQTY.Text) <> 0 Then TxtNetrate.Text = Val(LBLSUBTOTAL) / Val(TXTQTY.Text)
'    If optdiscper.Value = True Then
'        txtPD.Tag = ((Val(LBLSUBTOTAL.Caption) - Val(lbltaxamount.Caption)) * Val(txtPD.Text) / 100)
'    Else
'        txtPD.Tag = ((Val(LBLSUBTOTAL.Caption) - Val(lbltaxamount.Caption)) * Val(txtPD.Text) / 100)
'        lbltaxamount.Caption = (Val(Txtgrossamt.Text) - Val(txtPD.Text)) * Val(TxttaxMRP.Text) / 100
'    End If
'
'    LBLSUBTOTAL.Caption = Format(Val(LBLSUBTOTAL.Caption) - Val(txtPD.Tag), ".000")
'    txtPD.Text = Format(txtPD.Text, "0.00")
End Sub


Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TXTDEALER_GotFocus()
    Call CHANGEBOXCOLOR(TXTDEALER, True)
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
    If DataList2.BoundText = "" Then Call TXTDEALER_Change
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Trim(TXTDEALER.Text) = "" Then Exit Sub
            If DataList2.VisibleCount = 0 Then
                If MsgBox("No such supplier exists. Do you want to create a new supplier", vbYesNo, "EzBiz") = vbNo Then
                    TXTDEALER.SetFocus
                    Exit Sub
                Else
                    frmsuppliermast.Show
                    frmsuppliermast.SetFocus
                    frmsuppliermast.txtsupplier.Text = Trim(TXTDEALER.Text)
                    Exit Sub
                End If
            End If
            DataList2.SetFocus
        Case vbKeyEscape
            If M_ADD = False Then
                FRMECONTROLS.Enabled = False
                FRMEMASTER.Enabled = False
                txtBillNo.Enabled = True
                txtBillNo.SetFocus
            End If
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

Private Sub DataList2_Click()
    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text
    Call FILLCOMBO
    On Error GoTo ErrHand
    Dim rstCustomer As ADODB.Recordset
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address)) & Chr(13) & "GSTIN: " & IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
    Else
        lbladdress.Caption = ""
    End If
        
    'TXTDEALER.Text = lbldealer.Caption
    'LBL.Caption = ""
    Exit Sub
ErrHand:
    MsgBox Err.Description, , "EzBiz"
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "Purchase Bill..."
                DataList2.SetFocus
                Exit Sub
            End If
            TXTINVOICE.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
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
    'Call CHANGEBOXCOLOR(DataList2, True)
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub DataList2_LostFocus()
    'Call CHANGEBOXCOLOR(txtBillNo, False)
    DataList2.BackColor = vbWhite
    flagchange.Caption = ""
End Sub

Private Sub TXTRETAIL_GotFocus()
    If Val(TXTRETAIL.Text) = 0 And Val(TXTRATE.Text) <> 0 Then TXTRETAIL.Text = Val(TXTRATE.Text)
    If Val(TXTRETAIL.Text) = 0 Then TXTRETAIL.Text = ""
    Call CHANGEBOXCOLOR(TXTRETAIL, True)
    Call FILL_PREVIIOUSRATE
    TXTRETAIL.SelStart = 0
    TXTRETAIL.SelLength = Len(TXTRETAIL.Text)
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTRETAIL.Text) = 0 Then
                TxtRetailPercent.SetFocus
            Else
                txtWS.SetFocus
            End If
         Case vbKeyEscape
            txtPD.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TXTRETAIL, False)
    On Error Resume Next
    TXTRETAIL.Text = Format(TXTRETAIL.Text, "0.00")
'    If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) + ((Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) + ((Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    Else
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    End If
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MDIMAIN.lblgst.Caption <> "R" Then
        If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption)) + ((Val(TxtExpense.Text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
        End If
    Else
        If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
            TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.Text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
        End If
    End If
    If MDIMAIN.lblgst.Caption <> "R" Then
        TXTRETAIL.Tag = TXTRETAIL.Text
    Else
        'TXTRETAIL.Tag = (Val(TXTRETAIL.Text) - Val(txtCess.Text)) / (1 + ((Val(TxttaxMRP.Text)) / 100) + (Val(TxtCessPer.Text) / 100))
        'TXTRETAIL.Tag = Round(Val(TXTRETAIL.Tag) + (Val(TXTRETAIL.Tag) * Val(TxttaxMRP.Text) / 100), 4)
        If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
            'txtretail.Tag = Val(txtretail.Text) * 100 / ((Val(TxttaxMRP.Text) + Val(TxtCessPer.Text)) + 100)
            'txtretail.Tag = Val(txtretail.Tag) * 100 / ((Val(TxtCessPer.Text)) + 100)
            'txtretail.Tag = Round(Val(txtretail.Tag) - Val(txtCess.Text), 2)
            
            TXTRETAIL.Tag = (Val(TXTRETAIL.Text) - Val(txtCess.Text)) / (1 + ((Val(TxttaxMRP.Text)) / 100) + (Val(TxtCessPer.Text) / 100))
            txtCess.Tag = Round(Val(TXTRETAIL.Tag) + (Val(TXTRETAIL.Tag) * Val(TxttaxMRP.Text) / 100), 4)
            TXTRETAIL.Tag = Round(Val(txtCess.Tag) * 100 / ((Val(TxttaxMRP.Text)) + 100), 4)
            
'            TXTRETAILNOTAX.Text = (Val(txtNetrate.Text) - Val(TxtCessAmt.Text)) / (1 + ((Val(TXTTAX.Text) + Val(TxtKFC.Caption)) / 100) + (Val(TxtCessPer.Text) / 100))
'            txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + (Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100), 4)
'            TXTRETAILNOTAX.Text = Round(Val(txtretail.Text) * 100 / ((Val(TXTTAX.Text)) + 100), 4)
        
        Else
            TXTRETAIL.Tag = TXTRETAIL.Text
        End If
    End If
    If Val(Val(TXTPTR.Tag)) <> 0 Then
        TxtRetailPercent.Text = Round(((Val(TXTRETAIL.Tag) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
        TxtRetailPercent.Text = Format(Val(TxtRetailPercent.Text), "0.00")
    Else
         TxtRetailPercent.Text = Round(((Val(TXTRETAIL.Tag) - Val(TXTPTR.Tag)) * 100), 2)
        TxtRetailPercent.Text = Format(Val(TxtRetailPercent.Text), "0.00")
    End If
    
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub TxtWarranty_LostFocus()
    Call CHANGEBOXCOLOR(TxtWarranty, False)
End Sub

Private Sub txtws_GotFocus()
    Call CHANGEBOXCOLOR(txtWS, True)
    txtWS.SelStart = 0
    txtWS.SelLength = Len(txtWS.Text)
End Sub

Private Sub txtws_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtWS.Text) = 0 Then
                txtWsalePercent.SetFocus
            Else
                txtvanrate.SetFocus
            End If
         Case vbKeyEscape
            TXTRETAIL.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(txtWS, False)
    On Error Resume Next
    txtWS.Text = Format(txtWS.Text, "0.00")
'    If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) + ((Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) + ((Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    Else
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    End If
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MDIMAIN.lblgst.Caption <> "R" Then
        If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption)) + ((Val(TxtExpense.Text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
        End If
    Else
        If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
            TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.Text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
        End If
    End If
    If MDIMAIN.lblgst.Caption <> "R" Then
        txtWS.Tag = txtWS.Text
    Else
        'txtws.Tag = (Val(txtws.Text) - Val(txtCess.Text)) / (1 + ((Val(TxttaxMRP.Text)) / 100) + (Val(TxtCessPer.Text) / 100))
        'txtws.Tag = Round(Val(txtws.Tag) + (Val(txtws.Tag) * Val(TxttaxMRP.Text) / 100), 4)
        If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
            'TxtWS.Tag = Round(Val(TxtWS.Text) * 100 / ((Val(TxttaxMRP.Text)) + 100), 2)
            txtWS.Tag = (Val(txtWS.Text) - Val(txtCess.Text)) / (1 + ((Val(TxttaxMRP.Text)) / 100) + (Val(TxtCessPer.Text) / 100))
            txtCess.Tag = Round(Val(txtWS.Tag) + (Val(txtWS.Tag) * Val(TxttaxMRP.Text) / 100), 4)
            txtWS.Tag = Round(Val(txtCess.Tag) * 100 / ((Val(TxttaxMRP.Text)) + 100), 4)
        Else
            txtWS.Tag = txtWS.Text
        End If
    End If
    If Val(Val(TXTPTR.Tag)) <> 0 Then
        txtWsalePercent.Text = Round(((Val(txtWS.Tag) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
        txtWsalePercent.Text = Format(Val(txtWsalePercent.Text), "0.00")
    Else
        txtWsalePercent.Text = Round(((Val(txtWS.Tag) - Val(TXTPTR.Tag)) * 100), 2)
        txtWsalePercent.Text = Format(Val(txtWsalePercent.Text), "0.00")
    End If
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub txtcrtn_GotFocus()
    Call CHANGEBOXCOLOR(txtcrtn, True)
    If Val(txtcrtnpack.Text) = 0 Then txtcrtnpack.Text = "1"
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = "1"
    If Val(Los_Pack.Text) = 1 Then
        txtcrtn.Text = Format(Val(TXTRETAIL.Text), "0.00")
        txtcrtnpack.Text = "1"
    Else
        If Val(txtcrtn.Text) = 0 Then
            If Val(txtcrtnpack.Text) = 1 Then
                txtcrtn.Text = Format(Round(Val(TXTRETAIL.Text) / Val(Los_Pack.Text), 2), "0.00")
            Else
                txtcrtn.Text = Format(Round((Val(TXTRETAIL.Text) / Val(Los_Pack.Text)) * Val(txtcrtnpack.Text), 2), "0.00")
            End If
        End If
    End If
    
    txtcrtn.SelStart = 0
    txtcrtn.SelLength = Len(txtcrtn.Text)
End Sub

Private Sub txtcrtn_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtcrtn.Text) <> 0 And Val(txtcrtnpack.Text) = 0 Then
                MsgBox "Please enter the Pack Qty for Loose Qty", vbOKOnly, "EzBiz"
                txtcrtnpack.SetFocus
                Exit Sub
            End If
            If Val(Los_Pack.Text) = 1 Then
                txtcrtn.Text = Format(Val(TXTRETAIL.Text), "0.00")
                txtcrtnpack.Text = "1"
            End If
           TxtLWRate.SetFocus
         Case vbKeyEscape
            txtcrtnpack.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtcrtn_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcrtn_LostFocus()
    Call CHANGEBOXCOLOR(txtcrtn, False)
    txtcrtn.Text = Format(txtcrtn.Text, "0.00")
End Sub

Private Sub TxtComper_GotFocus()
    Call CHANGEBOXCOLOR(TxtComper, True)
    TxtComper.SelStart = 0
    TxtComper.SelLength = Len(TxtComper.Text)
    OptComper.value = True
End Sub

Private Sub TxtComper_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn
            TxtCessPer.SetFocus
        Case vbKeyEscape
            TxtCustDisc.SetFocus
        Case vbKeyDown
'            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
'            If Val(TXTQTY.Text) = 0 Then Exit Sub
'            If Val(TXTPTR.Text) = 0 Then Exit Sub
'            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtComper_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtComper_LostFocus()
    Call CHANGEBOXCOLOR(TxtComper, False)
    TxtComper.Text = Format(TxtComper.Text, "0.00")
End Sub

Private Sub TxtComAmt_GotFocus()
    Call CHANGEBOXCOLOR(TxtComAmt, True)
    TxtComAmt.SelStart = 0
    TxtComAmt.SelLength = Len(TxtComAmt.Text)
    OptComAmt.value = True
End Sub

Private Sub TxtComAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtCessPer.SetFocus
        Case vbKeyEscape
            TxtCustDisc.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtComAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtComAmt_LostFocus()
    Call CHANGEBOXCOLOR(TxtComAmt, False)
    TxtComAmt.Text = Format(TxtComAmt.Text, "0.00")
End Sub

Private Sub OptComAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TxtComAmt.Enabled = True Then
                TxtComAmt.SetFocus
            ElseIf cmdadd.Enabled = True Then
                cmdadd.SetFocus
            End If
        Case vbKeyEscape
            If TxtComAmt.Enabled = True Then TxtComAmt.SetFocus
'            TxtComAmt.Enabled = True
'            TxtComAmt.SetFocus
    End Select
End Sub

Private Sub OptComAmt_GotFocus()
    Call CHANGEBOXCOLOR(TxtComper, True)
    TxtComper.Text = ""
    TxtComAmt.Enabled = True
    TxtComper.Enabled = False
    TxtComAmt.SetFocus
End Sub

Private Sub OptComper_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TxtComper.Enabled = True Then
                TxtComper.SetFocus
            ElseIf cmdadd.Enabled = True Then
                cmdadd.SetFocus
            End If
        Case vbKeyEscape
            If TxtComper.Enabled = True Then TxtComper.SetFocus
'            TxtComper.Enabled = True
'            TxtComper.SetFocus
    End Select
End Sub

Private Sub OptComper_GotFocus()
    cmbfull.BackColor = &H98F3C1
    TxtComAmt.Text = ""
    TxtComAmt.Enabled = False
    TxtComper.Enabled = True
    TxtComper.SetFocus
End Sub

Private Sub txtcrtnpack_GotFocus()
    Call CHANGEBOXCOLOR(txtcrtnpack, True)
    If Val(Los_Pack.Text) = 1 Then
        txtcrtn.Text = Format(Val(TXTRETAIL.Text), "0.00")
        txtcrtnpack.Text = "1"
    End If
    txtcrtnpack.SelStart = 0
    txtcrtnpack.SelLength = Len(txtcrtnpack.Text)
End Sub

Private Sub txtcrtnpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtcrtnpack.Text) = 0 Then txtcrtnpack.Text = "1"
            txtcrtn.SetFocus
         Case vbKeyEscape
            txtvanrate.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtcrtnpack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcrtnpack_LostFocus()
    Call CHANGEBOXCOLOR(txtcrtnpack, False)
    txtcrtnpack.Text = Format(txtcrtnpack.Text, "0.00")
End Sub

Private Sub txtvanrate_GotFocus()
    Call CHANGEBOXCOLOR(txtvanrate, True)
    txtvanrate.SelStart = 0
    txtvanrate.SelLength = Len(txtvanrate.Text)
End Sub

Private Sub txtvanrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtvanrate.Text) = 0 Then
                txtSchPercent.SetFocus
            Else
                txtcrtnpack.SetFocus
            End If
         Case vbKeyEscape
            txtWS.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(txtvanrate, False)
    On Error Resume Next
    txtvanrate.Text = Format(txtvanrate.Text, "0.00")
'    If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) + ((Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) + ((Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    Else
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    End If
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MDIMAIN.lblgst.Caption <> "R" Then
        If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption)) + ((Val(TxtExpense.Text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
        End If
    Else
        If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
            TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.Text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
        End If
    End If
    If MDIMAIN.lblgst.Caption <> "R" Then
        txtvanrate.Tag = txtvanrate.Text
    Else
        'txtvanrate.Tag = (Val(txtvanrate.Text) - Val(txtCess.Text)) / (1 + ((Val(TxttaxMRP.Text)) / 100) + (Val(TxtCessPer.Text) / 100))
        'txtvanrate.Tag = Round(Val(txtvanrate.Tag) + (Val(txtvanrate.Tag) * Val(TxttaxMRP.Text) / 100), 4)
        If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
            'txtvanrate.Tag = Round(Val(txtvanrate.Text) * 100 / ((Val(TxttaxMRP.Text)) + 100), 2)
            txtvanrate.Tag = (Val(txtvanrate.Text) - Val(txtCess.Text)) / (1 + ((Val(TxttaxMRP.Text)) / 100) + (Val(TxtCessPer.Text) / 100))
            txtCess.Tag = Round(Val(txtvanrate.Tag) + (Val(txtvanrate.Tag) * Val(TxttaxMRP.Text) / 100), 4)
            txtvanrate.Tag = Round(Val(txtCess.Tag) * 100 / ((Val(TxttaxMRP.Text)) + 100), 4)
        Else
            txtvanrate.Tag = txtvanrate.Text
        End If
    End If
    If Val(Val(TXTPTR.Tag)) <> 0 Then
        txtSchPercent.Text = Round(((Val(txtvanrate.Tag) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
        txtSchPercent.Text = Format(Val(txtSchPercent.Text), "0.00")
    Else
        txtSchPercent.Text = Round(((Val(txtvanrate.Tag) - Val(TXTPTR.Tag)) * 100), 2)
        txtSchPercent.Text = Format(Val(txtSchPercent.Text), "0.00")
    End If
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub Txtgrossamt_GotFocus()
    Call CHANGEBOXCOLOR(Txtgrossamt, True)
    Txtgrossamt.SelStart = 0
    Txtgrossamt.SelLength = Len(Txtgrossamt.Text)
End Sub

Private Sub Txtgrossamt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(Txtgrossamt.Text) = 0 Then Exit Sub
            If Trim(txtHSN.Text) = "" Then
                txtHSN.Enabled = True
                txtHSN.SetFocus
            Else
                txtPD.SetFocus
            End If
        Case vbKeyEscape
            TxttaxMRP.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub Txtgrossamt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtgrossamt_LostFocus()
    Call CHANGEBOXCOLOR(Txtgrossamt, False)
    If Val(Txtgrossamt.Text) <> 0 Then
        Txtgrossamt.Text = Format(Txtgrossamt.Text, ".000")
        If Val(TXTQTY.Text) <> 0 Then
            TXTPTR.Text = Format(Round(Val(Txtgrossamt.Text) / Val(TXTQTY.Text), 4), "0.0000")
        ElseIf Val(TXTPTR.Text) <> 0 Then
            TXTQTY.Text = Format(Round(Val(Txtgrossamt.Text) / Val(TXTPTR.Text), 4), "0.0000")
        End If
    End If
    Call TxttaxMRP_LostFocus
    
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            TxtRetailPercent.Text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtWsalePercent.Text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtSchPercent.Text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtSchPercent_LostFocus
    End If
    
End Sub

Function FILL_PREVIIOUSRATE()
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, ITEM_COST_PRICE, ITEM_NET_COST_PRICE, P_RETAIL, P_WS, VCH_NO, VCH_DATE, VCH_DESC  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE = 'PI' OR TRX_TYPE = 'PW' OR TRX_TYPE = 'LP') ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, ITEM_COST_PRICE, ITEM_NET_COST_PRICE, P_RETAIL, P_WS, VCH_NO, VCH_DATE, VCH_DESC  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE = 'PI' OR TRX_TYPE = 'PW' OR TRX_TYPE = 'LP') ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    End If
    
    If PHY_PRERATE.RecordCount > 0 Then
        'Fram.Enabled = False
        fRMEPRERATE.Visible = True
        Set GRDPRERATE.DataSource = PHY_PRERATE
        
'        Select Case PHY_PRERATE!TRX_TYPE
'            Case "CN"
'                GRDSTOCK.TextMatrix(i, 3) = "SALES RETURN"
'                GRDSTOCK.TextMatrix(i, 4) = Mid(rststock!VCH_DESC, 15)
'            Case "XX", "OP"
'                GRDSTOCK.TextMatrix(i, 3) = "OPENING STOCK"
'            Case Else
'                GRDSTOCK.TextMatrix(i, 3) = "Purchase"
'                GRDSTOCK.TextMatrix(i, 4) = Mid(rststock!VCH_DESC, 15)
'        End Select
        GRDPRERATE.Columns(0).Caption = "TYPR"
        GRDPRERATE.Columns(1).Caption = "ITEM CODE"
        GRDPRERATE.Columns(2).Caption = "ITEM NAME"
        GRDPRERATE.Columns(3).Caption = "PACK"
        GRDPRERATE.Columns(4).Caption = "UNIT"
        GRDPRERATE.Columns(5).Caption = "COST"
        GRDPRERATE.Columns(6).Caption = "NET COST"
        GRDPRERATE.Columns(7).Caption = "RT PRICE"
        GRDPRERATE.Columns(8).Caption = "WS PRICE"
        GRDPRERATE.Columns(9).Caption = "BILL NO."
        GRDPRERATE.Columns(10).Caption = "BILL DATE"
        GRDPRERATE.Columns(11).Caption = "RECEIVED FROM"
    
        GRDPRERATE.Columns(0).Visible = False
        GRDPRERATE.Columns(1).Visible = False
        GRDPRERATE.Columns(2).Width = 0
        GRDPRERATE.Columns(3).Width = 600
        GRDPRERATE.Columns(4).Width = 600
        GRDPRERATE.Columns(5).Width = 1200
        GRDPRERATE.Columns(6).Width = 1200
        GRDPRERATE.Columns(7).Width = 1100
        GRDPRERATE.Columns(8).Width = 1300
        GRDPRERATE.Columns(9).Width = 1100
        GRDPRERATE.Columns(10).Width = 1300
        GRDPRERATE.Columns(11).Width = 4500
        
        
        'GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(2).Text
    End If
    
End Function

Private Sub GRDPRERATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            'FRMEMAIN.Enabled = True
            TXTPTR.Enabled = True
            TXTPTR.SetFocus
    End Select
End Sub

Private Sub Los_Pack_GotFocus()
    Call CHANGEBOXCOLOR(Los_Pack, True)
    Los_Pack.SelStart = 0
    Los_Pack.SelLength = Len(Los_Pack.Text)
    FRMEGRDTMP.Visible = False
    CmbPack.Enabled = True
    cmbfull.Enabled = True
    TXTQTY.Enabled = True
    TxTfree.Enabled = True
    TXTRATE.Enabled = True
    TXTPTR.Enabled = True
    TxttaxMRP.Enabled = True
    TxtExDuty.Enabled = True
    TxtTrDisc.Enabled = True
    TxtCessPer.Enabled = True
    txtCess.Enabled = True
    TxtCSTper.Enabled = True
    txtPD.Enabled = True
    TxtExpense.Enabled = True
    TXTRETAIL.Enabled = True
    TxtRetailPercent.Enabled = True
    
    txtWS.Enabled = True
    txtWsalePercent.Enabled = True
    txtvanrate.Enabled = True
    txtSchPercent.Enabled = True
    txtcrtnpack.Enabled = True
    txtcrtn.Enabled = True
    TxtLWRate.Enabled = True
    TxtCustDisc.Enabled = True
    TxtComper.Enabled = True
    TxtComAmt.Enabled = True
    cmdadd.Enabled = True
    Txtgrossamt.Enabled = True
    txtBatch.Enabled = True
    txtHSN.Enabled = True
    TxtWarranty.Enabled = True
    CmbWrnty.Enabled = True
    TXTEXPIRY.Visible = False
    TXTEXPDATE.Enabled = True
    TxtBarcode.Enabled = False
    txtcategory.Enabled = False
    TXTPRODUCT.Enabled = False
    
    Dim rststock As ADODB.Recordset
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    With rststock
        If Not (.EOF And .BOF) Then
            On Error Resume Next
            cmbfull.Text = IIf(IsNull(rststock!FULL_PACK), 0, rststock!FULL_PACK)
            On Error GoTo ErrHand
        Else
            On Error Resume Next
            cmbfull.Text = CmbPack.Text
            On Error GoTo ErrHand
        End If
    End With
    rststock.Close
    Set rststock = Nothing
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Los_Pack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            cmbfull.SetFocus
         Case vbKeyEscape
             If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            Los_Pack.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
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

Private Sub TxtItemcode_GotFocus()
    Call CHANGEBOXCOLOR(TXTITEMCODE, True)
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim I As Long
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
        
            If Trim(TXTITEMCODE.Text) = "" Then
                TXTPRODUCT.Enabled = True
                TXTPRODUCT.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False
            
            Set grdtmp.DataSource = Nothing
            If PHYCODE_FLAG = True Then
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            Else
                PHY_CODE.Close
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            End If
            
            Set grdtmp.DataSource = PHY_CODE
            
            If PHY_CODE.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "EzBiz"
                Exit Sub
            End If
            
            If PHY_CODE.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                lblcategory.Caption = IIf(IsNull(PHY_CODE!Category), "", PHY_CODE!Category)
                On Error Resume Next
                Set Image1.DataSource = PHY
                If IsNull(PHY!PHOTO) Then
                    Frame6.Visible = False
                    Set Image1.DataSource = Nothing
                    bytData = ""
                Else
                    If Err.Number = 545 Then
                        Frame6.Visible = False
                        Set Image1.DataSource = Nothing
                        bytData = ""
                    Else
                        Frame6.Visible = True
                        Set Image1.DataSource = PHY 'setting image1�s datasource
                        Image1.DataField = "PHOTO"
                        bytData = PHY!PHOTO
                    End If
                End If
                On Error GoTo ErrHand
                For I = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(I, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & I & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                        Exit For
                    End If
                Next I

                Set RSTRXFILE = New ADODB.Recordset
                'RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                'RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE = 'PI' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    'RSTRXFILE.MoveLast
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.Text = 1
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.Text = IIf(IsNull(RSTRXFILE!REF_NO), "", RSTRXFILE!REF_NO)
                    TXTEXPIRY.Text = IIf(IsDate(RSTRXFILE!EXP_DATE), Format(RSTRXFILE!EXP_DATE, "MM/YY"), "  /  ")
                    Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
                    If (IsNull(RSTRXFILE!MRP)) Then
                        TXTRATE.Text = ""
                    Else
                        TXTRATE.Text = Format(Round(Val(RSTRXFILE!MRP) * Val(Los_Pack.Text), 2), ".000")
                    End If
                    If (IsNull(RSTRXFILE!MRP_BT)) Then
                        txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105
                    Else
                        txtmrpbt.Text = Val(TXTRATE.Text)
                    End If
                    If IsNull(RSTRXFILE!PTR) Then
                        TXTPTR.Text = ""
                    Else
                        TXTPTR.Text = Format(Round(Val(RSTRXFILE!PTR) * Val(Los_Pack.Text), 3), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_RETAIL) Then
                        TXTRETAIL.Text = ""
                    Else
                        TXTRETAIL.Text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_WS) Then
                        txtWS.Text = ""
                    Else
                        txtWS.Text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_VAN) Then
                        txtvanrate.Text = ""
                    Else
                        txtvanrate.Text = Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_CRTN) Then
                        txtcrtn.Text = ""
                    Else
                        txtcrtn.Text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_LWS) Then
                        TxtLWRate.Text = ""
                    Else
                        TxtLWRate.Text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!CRTN_PACK) Then
                        txtcrtnpack.Text = ""
                    Else
                        txtcrtnpack.Text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_PRICE) Then
                        txtprofit.Text = ""
                    Else
                        txtprofit.Text = Format(Round(Val(RSTRXFILE!SALES_PRICE), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.Text = ""
                    Else
                        TxttaxMRP.Text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                    End If
                    If IsNull(RSTRXFILE!EXDUTY) Then
                        TxtExDuty.Text = ""
                    Else
                        TxtExDuty.Text = Format(Val(RSTRXFILE!EXDUTY), ".00")
                    End If
                    If IsNull(RSTRXFILE!CSTPER) Then
                        TxtCSTper.Text = ""
                    Else
                        TxtCSTper.Text = Format(Val(RSTRXFILE!CSTPER), ".00")
                    End If
                    If IsNull(RSTRXFILE!TR_DISC) Then
                        TxtTrDisc.Text = ""
                    Else
                        TxtTrDisc.Text = Format(Val(RSTRXFILE!TR_DISC), ".00")
                    End If
                    If IsNull(RSTRXFILE!CESS_AMT) Then
                        txtCess.Text = ""
                    Else
                        txtCess.Text = Format(Val(RSTRXFILE!CESS_AMT), ".00")
                    End If
                    If IsNull(RSTRXFILE!CESS_PER) Then
                        txtCess.Text = ""
                    Else
                        txtCess.Text = Format(Val(RSTRXFILE!CESS_PER), ".00")
                    End If
                    TxtWarranty.Text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                    If RSTRXFILE!COM_FLAG = "A" Then
                        TxtComAmt.Text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                        OptComAmt.value = True
                    Else
                        TxtComper.Text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                        OptComper.value = True
                    End If
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.Text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo ErrHand
                    
                    'TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                    If RSTRXFILE!CHECK_FLAG = "M" Then
                        OPTTaxMRP.value = True
                    ElseIf RSTRXFILE!CHECK_FLAG = "V" Then
                        OPTVAT.value = True
                    Else
                        OPTNET.value = True
                    End If
                Else
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = 1
                    Los_Pack.Text = 1
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.Text = ""
                    txtHSN.Text = ""
                    TXTEXPIRY.Text = "  /  "
                    TXTRATE.Text = ""
                    txtmrpbt.Text = ""
                    TXTPTR.Text = ""
                    TxtNetrate.Text = ""
                    TXTRETAIL.Text = ""
                    txtWS.Text = ""
                    txtvanrate.Text = ""
                    txtcrtn.Text = ""
                    TxtLWRate.Text = ""
                    txtcrtnpack.Text = ""
                    txtprofit.Text = ""
                    TxttaxMRP.Text = ""
                    TxtExDuty.Text = ""
                    TxtCSTper.Text = ""
                    TxtTrDisc.Text = ""
                    TxtCustDisc.Text = ""
                    TxtCessPer.Text = ""
                    txtCess.Text = ""
                    Los_Pack.Text = "1"
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    OPTVAT.value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                
                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                With RSTRXFILE
                    If Not (.EOF And .BOF) Then
                        If IsNull(RSTRXFILE!P_RETAIL) Then
                            TXTRETAIL.Text = ""
                        Else
                            TXTRETAIL.Text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                        End If
                        If IsNull(RSTRXFILE!P_WS) Then
                            txtWS.Text = ""
                        Else
                            txtWS.Text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                        End If
                        If IsNull(RSTRXFILE!P_VAN) Then
                            txtvanrate.Text = ""
                        Else
                            txtvanrate.Text = Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000")
                        End If
                        If RSTRXFILE!COM_FLAG = "A" Then
                            TxtComAmt.Text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                            OptComAmt.value = True
                        Else
                            TxtComper.Text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                            OptComper.value = True
                        End If
                        If IsNull(RSTRXFILE!P_CRTN) Then
                            txtcrtn.Text = ""
                        Else
                            txtcrtn.Text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                        End If
                        If IsNull(RSTRXFILE!P_LWS) Then
                            TxtLWRate.Text = ""
                        Else
                            TxtLWRate.Text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                        End If
                        If IsNull(RSTRXFILE!CRTN_PACK) Then
                            txtcrtnpack.Text = ""
                        Else
                            txtcrtnpack.Text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                        End If
                    End If
                End With
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                
                If PHY_CODE.RecordCount = 1 Then
                    If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                        TXTITEMCODE.Enabled = False
                        TXTPRODUCT.Enabled = False
                        txtcategory.Enabled = False
                        Los_Pack.Text = 1
                        TXTQTY.Text = 1
                        TxTfree.Text = ""
                        TXTRATE.Text = ""
                        TXTPTR.Enabled = True
                        TXTPTR.SetFocus
                        'TxtPack.Enabled = True
                        'TxtPack.SetFocus
                    Else
                        TXTITEMCODE.Enabled = False
                        TXTPRODUCT.Enabled = False
                        Los_Pack.Enabled = True
                        Los_Pack.SetFocus
                        'TxtPack.Enabled = True
                        'TxtPack.SetFocus
                    End If
                    Exit Sub
                End If
            ElseIf PHY_CODE.RecordCount > 1 Then
                FRMEGRDTMP.Visible = True
                Fram.Enabled = False
                grdtmp.Columns(0).Visible = False
                grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
                grdtmp.Columns(1).Width = 4700
                'grdtmp.Columns(2).Visible = False
                grdtmp.Columns(2).Caption = "QTY"
                grdtmp.Columns(2).Width = 1300
                grdtmp.SetFocus
            End If
            
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TxtItemcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub txtcst_GotFocus()
    Call CHANGEBOXCOLOR(TxtCST, True)
    TxtCST.SelStart = 0
    TxtCST.SelLength = Len(TxtCST.Text)
End Sub

Private Sub TxtCST_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            'If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            'If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            'If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTRETAIL.Enabled = True Then TXTRETAIL.SetFocus
            If txtWS.Enabled = True Then txtWS.SetFocus
            If txtcrtn.Enabled = True Then txtcrtn.SetFocus
            If TxtLWRate.Enabled = True Then TxtLWRate.SetFocus
            If txtcrtnpack.Enabled = True Then txtcrtnpack.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub TxtCST_LostFocus()
    Call CHANGEBOXCOLOR(TxtCST, False)
    Dim DISC As Currency
    
    On Error GoTo ErrHand
    If (TxtCST.Text = "") Then
        DISC = 0
    Else
        DISC = TxtCST.Text
    End If
    If grdsales.Rows = 1 Then
        TxtCST.Text = "0"
        Exit Sub
    End If
    TxtCST.Text = Format(TxtCST.Text, ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 2), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtCST.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtCST.SetFocus
End Sub

Private Sub TxtInsurance_GotFocus()
    Call CHANGEBOXCOLOR(TxtInsurance, True)
    TxtInsurance.SelStart = 0
    TxtInsurance.SelLength = Len(TxtInsurance.Text)
End Sub

Private Sub TxtInsurance_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtInsurance_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            'If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            'If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            'If txtBatch.Enabled = True Then txtBatch.SetFocus
            If TXTRETAIL.Enabled = True Then TXTRETAIL.SetFocus
            If txtWS.Enabled = True Then txtWS.SetFocus
            If txtcrtn.Enabled = True Then txtcrtn.SetFocus
            If TxtLWRate.Enabled = True Then TxtLWRate.SetFocus
            If txtcrtnpack.Enabled = True Then txtcrtnpack.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub TxtInsurance_LostFocus()
    Call CHANGEBOXCOLOR(TxtInsurance, False)
    Dim DISC As Currency
    
    On Error GoTo ErrHand
    If (TxtInsurance.Text = "") Then
        DISC = 0
    Else
        DISC = TxtInsurance.Text
    End If
    If grdsales.Rows = 1 Then
        TxtInsurance.Text = "0"
        Exit Sub
    End If
    TxtInsurance.Text = Format(TxtInsurance.Text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + (Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(txtcst.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 2), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtInsurance.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtInsurance.SetFocus
End Sub

Private Sub txtWsalePercent_GotFocus()
    Call CHANGEBOXCOLOR(txtWsalePercent, True)
    txtWsalePercent.SelStart = 0
    txtWsalePercent.SelLength = Len(txtWsalePercent.Text)
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.Text) <> 0 Then Label1(38).Caption = "Disc% on MRP"
End Sub

Private Sub txtWsalePercent_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            txtvanrate.SetFocus
         Case vbKeyEscape
            txtWS.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtWsalePercent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtWsalePercent_LostFocus()
    Call CHANGEBOXCOLOR(txtWsalePercent, False)
    On Error Resume Next
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.Text) <> 0 Then
        If Val(txtWsalePercent.Text) > 0 Then
            txtWS.Text = Round(Val(TXTRATE.Text) - (Val(TXTRATE.Text) * Val(txtWsalePercent.Text) / 100), 2)
        Else
            txtWS.Text = Val(TXTRATE.Text)
        End If
        Call txtws_LostFocus
    Else
        If MDIMAIN.lblgst.Caption <> "R" Then
            If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
                TXTPTR.Tag = Val(LBLSUBTOTAL.Caption)
            Else
                TXTPTR.Tag = Round(Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.Text) + Val(TxTfree.Text), 4)
            End If
        Else
            If Val(TXTQTY.Text) + Val(TxTfree.Text) = 0 Then
                TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.Text)))), 4)
            Else
                TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.Text) + Val(TxTfree.Text)))) + ((Val(TxtExpense.Text) / (Val(TXTQTY.Text) + Val(TxTfree.Text))))), 4)
            End If
        End If
        If MDIMAIN.lblgst.Caption <> "R" Then
            txtWS.Text = Round((Val(TXTPTR.Tag) * Val(txtWsalePercent.Text) / 100) + Val(TXTPTR.Tag), 2)
        Else
            If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
                txtWS.Text = (Val(TXTPTR.Tag) * Val(txtWsalePercent.Text) / 100) + Val(TXTPTR.Tag)
                txtWS.Text = Round(Val(txtWS.Text) + (Val(txtWS.Text) * (Val(TxttaxMRP.Text) + Val(TxtCessPer.Text)) / 100), 2)
            Else
                txtWS.Text = Round((Val(TXTPTR.Tag) * Val(txtWsalePercent.Text) / 100) + Val(TXTPTR.Tag), 2)
            End If
        End If
    End If
    
    txtWS.Text = Format(Val(txtWS.Text), "0.00")
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub TxtWarranty_GotFocus()
    Call CHANGEBOXCOLOR(TxtWarranty, True)
    TxtWarranty.SelStart = 0
    TxtWarranty.SelLength = Len(TxtWarranty.Text)
End Sub

Private Sub TxtWarranty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.Text) = 0 Then
                cmdadd.SetFocus
            Else
                CmbWrnty.SetFocus
            End If
         Case vbKeyEscape
            txtCess.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtWarranty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CmbWrnty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.Text) <> 0 And CmbWrnty.ListIndex = -1 Then
                MsgBox "Please select the Warranty Period", , "EzBiz"
                CmbWrnty.SetFocus
                Exit Sub
            End If
            If Val(TxtWarranty.Text) = 0 Then CmbWrnty.ListIndex = -1
            cmdadd.SetFocus
         Case vbKeyEscape
            TxtWarranty.SetFocus
    End Select
End Sub

Private Function checklastbill()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ErrHand
    
    Dim BillNO As Double
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'PI'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        BillNO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    If Val(txtBillNo.Text) >= BillNO Then
        txtBillNo.Text = BillNO
    End If
Exit Function
ErrHand:
    MsgBox Err.Description
End Function

Private Sub TxtExpense_GotFocus()
    Call CHANGEBOXCOLOR(TxtExpense, True)
    TxtExpense.SelStart = 0
    TxtExpense.SelLength = Len(TxtExpense.Text)
End Sub

Private Sub TxtExpense_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRETAIL.SetFocus
         Case vbKeyEscape
            TxtTrDisc.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtExpense_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtExDuty_GotFocus()
    Call CHANGEBOXCOLOR(TxtExDuty, True)
    TxtExDuty.SelStart = 0
    TxtExDuty.SelLength = Len(TxtExDuty.Text)
End Sub

Private Sub TxtExDuty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtCSTper.SetFocus
         Case vbKeyEscape
            txtcrtn.SetFocus
    End Select
End Sub

Private Sub TxtExDuty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCSTper_GotFocus()
    Call CHANGEBOXCOLOR(TxtCSTper, True)
    TxtCSTper.SelStart = 0
    TxtCSTper.SelLength = Len(TxtCSTper.Text)
End Sub

Private Sub TxtCSTper_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRETAIL.SetFocus
         Case vbKeyEscape
            TxtExDuty.SetFocus
    End Select
End Sub

Private Sub TxtCSTper_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtTrDisc_GotFocus()
    Call CHANGEBOXCOLOR(TxtTrDisc, True)
    TxtTrDisc.SelStart = 0
    TxtTrDisc.SelLength = Len(TxtTrDisc.Text)
End Sub

Private Sub TxtTrDisc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtExpense.SetFocus
        Case vbKeyEscape
            txtPD.SetFocus
'            Frame1.Enabled = True
'            If OptComper.value = True Then
'                TxtComper.SetFocus
'            Else
'                TxtComAmt.SetFocus
'            End If
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtTrDisc_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLWRate_GotFocus()
    Call CHANGEBOXCOLOR(TxtLWRate, True)
    On Error Resume Next
    If Val(txtcrtnpack.Text) = 0 Then txtcrtnpack.Text = "1"
    If Val(Los_Pack.Text) = 1 Then
        TxtLWRate.Text = Format(Val(txtWS.Text), "0.00")
        txtcrtnpack.Text = "1"
    Else
        If Val(TxtLWRate.Text) = 0 Then
            If Val(txtcrtnpack.Text) = 1 Then
                TxtLWRate.Text = Format(Round(Val(txtWS.Text) / Val(Los_Pack.Text), 2), "0.00")
            Else
                TxtLWRate.Text = Format(Round((Val(txtWS.Text) / Val(Los_Pack.Text)) * Val(txtcrtnpack.Text), 2), "0.00")
            End If
        End If
    End If
    
    TxtLWRate.SelStart = 0
    TxtLWRate.SelLength = Len(TxtLWRate.Text)
End Sub

Private Sub TxtLWRate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtLWRate.Text) <> 0 And Val(txtcrtnpack.Text) = 0 Then
                MsgBox "Please enter the Pack Qty for Loose Qty", vbOKOnly, "EzBiz"
                txtcrtnpack.SetFocus
                Exit Sub
            End If
            If Val(Los_Pack.Text) = 1 Then
                TxtLWRate.Text = Format(Val(txtWS.Text), "0.00")
                txtcrtnpack.Text = "1"
            End If
            TxtCustDisc.SetFocus
         Case vbKeyEscape
            txtcrtn.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtLWRate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLWRate_LostFocus()
    Call CHANGEBOXCOLOR(TxtLWRate, False)
    TxtLWRate.Text = Format(TxtLWRate.Text, "0.00")
End Sub

Private Function FILLCOMBO()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set CMBPO.DataSource = Nothing
    If PO_FLAG = True Then
        ACT_PO.Open "Select VCH_NO, ACT_CODE from POMAST  WHERE (VCH_NO = " & Val(PONO) & " OR (ISNULL(STATUS) OR STATUS = 'N')) AND ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_DATE ASC", db, adOpenStatic, adLockReadOnly, adCmdText
        PO_FLAG = False
    Else
        ACT_PO.Close
        ACT_PO.Open "Select VCH_NO, ACT_CODE from POMAST  WHERE (VCH_NO = " & Val(PONO) & " OR (ISNULL(STATUS) OR STATUS = 'N')) AND ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_DATE ASC", db, adOpenStatic, adLockReadOnly, adCmdText
        PO_FLAG = False
    End If
    
    Set Me.CMBPO.RowSource = ACT_PO
    CMBPO.ListField = "VCH_NO"
    CMBPO.BoundColumn = "ACT_CODE"
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TxtBarcode.Text) = "" Then
                txtcategory.Enabled = True
                txtcategory.SetFocus
                Exit Sub
            End If
            
            Set rstTRXMAST = New ADODB.Recordset
            'MFG_REC.Open "SELECT DISTINCT CATEGORY FROM ITEMMAST RIGHT JOIN RTRXFILE ON ITEMMAST.ITEM_CODE = RTRXFILE.ITEM_CODE WHERE RTRXFILE.BAL_QTY > 0 ORDER BY ITEMMAST.MANUFACTURER", db, adOpenForwardOnly ' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y')
            'rstTRXMAST.Open "Select * From RTRXFILE LEFT JOIN ON ITEMMAST.ITEM_CODE = RTRXFILE.ITEM_CODE WHERE RTRXFILE.BARCODE= '" & Trim(txtBarcode.Text) & "' AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') ORDER BY VCH_NO ", db, adOpenStatic, adLockReadOnly
            'WHERE RTRXFILE.BARCODE= '" & Trim(txtBarcode.Text) & "' AND ITEMMAST.UN_BILL <> 'Y' ORDER BY VCH_NO
            rstTRXMAST.Open "Select * From RTRXFILE LEFT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE RTRXFILE.BARCODE= '" & Trim(TxtBarcode.Text) & "' AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                CHANGE_FLAG = True
                TXTITEMCODE.Text = IIf(IsNull(rstTRXMAST!ITEM_CODE), "", rstTRXMAST!ITEM_CODE)
                TXTPRODUCT.Text = IIf(IsNull(rstTRXMAST!ITEM_NAME), "", rstTRXMAST!ITEM_NAME)
                CHANGE_FLAG = False
                TXTUNIT.Text = 1 'IIf(IsNull(rstTRXMAST!UNIT), "", rstTRXMAST!UNIT)
                
                Set RSTITEM = New ADODB.Recordset
                RSTITEM.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & rstTRXMAST!TRX_YEAR & "' AND TRX_TYPE= '" & rstTRXMAST!TRX_TYPE & "' AND VCH_NO = " & rstTRXMAST!VCH_NO & " AND LINE_NO = " & rstTRXMAST!line_no & "", db, adOpenStatic, adLockReadOnly
                If Not (RSTITEM.EOF Or RSTITEM.BOF) Then
                    Txtpack.Text = IIf(IsNull(RSTITEM!LINE_DISC), "", RSTITEM!LINE_DISC)
                    Txtpack.Text = 1
                    Los_Pack.Text = IIf(IsNull(RSTITEM!LOOSE_PACK), "1", RSTITEM!LOOSE_PACK)
                    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
                    TxtWarranty.Text = IIf(IsNull(RSTITEM!WARRANTY), "", RSTITEM!WARRANTY)
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(RSTITEM!PACK_TYPE), "Nos", RSTITEM!PACK_TYPE)
                    CmbWrnty.Text = IIf(IsNull(RSTITEM!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTITEM!WARRANTY_TYPE)
                    'cmbcolor.Text = IIf(IsNull(rstITEM!ITEM_COLOR), CmbWrnty.ListIndex = -1, rstITEM!ITEM_COLOR)
                    On Error GoTo ErrHand
                    'Txtsize.Text = IIf(IsNull(rstITEM!ITEM_SIZE), "", rstITEM!ITEM_SIZE)
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(rstITEM!EXP_DATE), "  /  /    ", Format(rstITEM!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.Text = IIf(IsNull(RSTITEM!REF_NO), "", RSTITEM!REF_NO)
                    TXTEXPIRY.Text = IIf(IsDate(RSTITEM!EXP_DATE), Format(RSTITEM!EXP_DATE, "MM/YY"), "  /  ")
                    TXTRATE.Text = IIf(IsNull(RSTITEM!MRP), "", Format(Round(Val(RSTITEM!MRP) * Val(Los_Pack.Text), 2), ".000"))
                    If (IsNull(RSTITEM!MRP_BT)) Then
                        txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105
                    Else
                        txtmrpbt.Text = Val(TXTRATE.Text)
                    End If
                    If IsNull(RSTITEM!PTR) Then
                        TXTPTR.Text = ""
                    Else
                        TXTPTR.Text = Format(Round(Val(RSTITEM!PTR) * Val(Los_Pack.Text), 2), ".000")
                    End If
                    If IsNull(RSTITEM!P_RETAIL) Then
                        TXTRETAIL.Text = ""
                    Else
                        TXTRETAIL.Text = Format(Round(Val(RSTITEM!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(RSTITEM!P_WS) Then
                        txtWS.Text = ""
                    Else
                        txtWS.Text = Format(Round(Val(RSTITEM!P_WS), 2), ".000")
                    End If
                    If IsNull(RSTITEM!P_VAN) Then
                        txtvanrate.Text = ""
                    Else
                        txtvanrate.Text = Format(Round(Val(RSTITEM!P_VAN), 2), ".000")
                    End If
                    If IsNull(RSTITEM!P_CRTN) Then
                        txtcrtn.Text = ""
                    Else
                        txtcrtn.Text = Format(Round(Val(RSTITEM!P_CRTN), 2), ".000")
                    End If
                    If IsNull(RSTITEM!CRTN_PACK) Then
                        txtcrtnpack.Text = ""
                    Else
                        txtcrtnpack.Text = Format(Round(Val(RSTITEM!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(RSTITEM!SALES_PRICE) Then
                        txtprofit.Text = ""
                    Else
                        txtprofit.Text = Format(Round(Val(RSTITEM!SALES_PRICE), 2), ".000")
                    End If
                    If IsNull(RSTITEM!SALES_TAX) Then
                        TxttaxMRP.Text = ""
                    Else
                        TxttaxMRP.Text = Format(Val(RSTITEM!SALES_TAX), ".00")
                    End If
                    Los_Pack.Text = IIf(IsNull(RSTITEM!LOOSE_PACK), "1", RSTITEM!LOOSE_PACK)
                    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
                    TxtWarranty.Text = IIf(IsNull(RSTITEM!WARRANTY), "", RSTITEM!WARRANTY)
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(RSTITEM!PACK_TYPE), "Nos", RSTITEM!PACK_TYPE)
                    CmbWrnty.Text = IIf(IsNull(RSTITEM!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTITEM!WARRANTY_TYPE)
                    On Error GoTo ErrHand
                    txtPD.Text = IIf(IsNull(RSTITEM!P_DISC), "", RSTITEM!P_DISC)
                    Select Case RSTITEM!DISC_FLAG
                        Case "P"
                            optdiscper.value = True
                        Case "A"
                            Optdiscamt.value = True
                    End Select
                    'TxttaxMRP.Text = IIf(IsNull(rstITEM!SALES_TAX), "", Format(Val(rstITEM!SALES_TAX), ".00"))
                    If RSTITEM!CHECK_FLAG = "M" Then
                        OPTTaxMRP.value = True
                    ElseIf RSTITEM!CHECK_FLAG = "V" Then
                        OPTVAT.value = True
                    Else
                        OPTNET.value = True
                    End If
                End If
                RSTITEM.Close
                Set RSTITEM = Nothing
                
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                'txtbarcode.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
            Else
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                
                Set rstTRXMAST = New ADODB.Recordset
                rstTRXMAST.Open "Select * From ITEMMAST WHERE BARCODE= '" & Trim(TxtBarcode.Text) & "' AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
                If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                    CHANGE_FLAG = True
                    TXTITEMCODE.Text = IIf(IsNull(rstTRXMAST!ITEM_CODE), "", rstTRXMAST!ITEM_CODE)
                    TXTPRODUCT.Text = IIf(IsNull(rstTRXMAST!ITEM_NAME), "", rstTRXMAST!ITEM_NAME)
                    CHANGE_FLAG = False
                    TXTUNIT.Text = 1 'IIf(IsNull(rstTRXMAST!UNIT), "", rstTRXMAST!UNIT)
                    Txtpack.Text = 1
                    Los_Pack.Text = IIf(IsNull(rstTRXMAST!LOOSE_PACK), "1", rstTRXMAST!LOOSE_PACK)
                    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
                    TxtWarranty.Text = IIf(IsNull(rstTRXMAST!WARRANTY), "", rstTRXMAST!WARRANTY)
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)
                    CmbWrnty.Text = IIf(IsNull(rstTRXMAST!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, rstTRXMAST!WARRANTY_TYPE)
                    'cmbcolor.Text = IIf(IsNull(rstTRXMAST!ITEM_COLOR), CmbWrnty.ListIndex = -1, rstTRXMAST!ITEM_COLOR)
                    On Error GoTo ErrHand
                    'Txtsize.Text = IIf(IsNull(rstTRXMAST!ITEM_SIZE), "", rstTRXMAST!ITEM_SIZE)
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(rstTRXMAST!EXP_DATE), "  /  /    ", Format(rstTRXMAST!EXP_DATE, "DD/MM/YYYY"))
                    'txtBatch.Text = IIf(IsNull(rstTRXMAST!REF_NO), "", rstTRXMAST!REF_NO)
                    'TXTEXPIRY.Text = IIf(IsDate(rstTRXMAST!EXP_DATE), Format(rstTRXMAST!EXP_DATE, "MM/YY"), "  /  ")
                    TXTRATE.Text = IIf(IsNull(rstTRXMAST!MRP), "", Format(Round(Val(rstTRXMAST!MRP) * Val(Los_Pack.Text), 2), ".000"))
                    If IsNull(rstTRXMAST!ITEM_COST) Then
                        TXTPTR.Text = ""
                    Else
                        TXTPTR.Text = Format(Round(Val(rstTRXMAST!ITEM_COST) * Val(Los_Pack.Text), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!P_RETAIL) Then
                        TXTRETAIL.Text = ""
                    Else
                        TXTRETAIL.Text = Format(Round(Val(rstTRXMAST!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!P_WS) Then
                        txtWS.Text = ""
                    Else
                        txtWS.Text = Format(Round(Val(rstTRXMAST!P_WS), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!P_VAN) Then
                        txtvanrate.Text = ""
                    Else
                        txtvanrate.Text = Format(Round(Val(rstTRXMAST!P_VAN), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!P_CRTN) Then
                        txtcrtn.Text = ""
                    Else
                        txtcrtn.Text = Format(Round(Val(rstTRXMAST!P_CRTN), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!CRTN_PACK) Then
                        txtcrtnpack.Text = ""
                    Else
                        txtcrtnpack.Text = Format(Round(Val(rstTRXMAST!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!SALES_PRICE) Then
                        txtprofit.Text = ""
                    Else
                        txtprofit.Text = Format(Round(Val(rstTRXMAST!SALES_PRICE), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!SALES_TAX) Then
                        TxttaxMRP.Text = ""
                    Else
                        TxttaxMRP.Text = Format(Val(rstTRXMAST!SALES_TAX), ".00")
                    End If
                    Los_Pack.Text = IIf(IsNull(rstTRXMAST!LOOSE_PACK), "1", rstTRXMAST!LOOSE_PACK)
                    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
                    TxtWarranty.Text = IIf(IsNull(rstTRXMAST!WARRANTY), "", rstTRXMAST!WARRANTY)
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)
                    CmbWrnty.Text = IIf(IsNull(rstTRXMAST!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, rstTRXMAST!WARRANTY_TYPE)
                    On Error GoTo ErrHand
                    txtPD.Text = IIf(IsNull(rstTRXMAST!CUST_DISC), "", rstTRXMAST!CUST_DISC)
                    optdiscper.value = True
                    OPTVAT.value = True
                    'TxttaxMRP.Text = IIf(IsNull(rstTRXMAST!SALES_TAX), "", Format(Val(rstTRXMAST!SALES_TAX), ".00"))
                    
                    rstTRXMAST.Close
                    Set rstTRXMAST = Nothing
                    'txtbarcode.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                Else
                    rstTRXMAST.Close
                    Set rstTRXMAST = Nothing
                    TxtBarcode.Enabled = False
                    txtcategory.Enabled = True
                    txtcategory.SetFocus
                End If
            End If
            If Trim(TxtBarcode.Text) = "" Then
                BARCODE_FLAG = False
            Else
                BARCODE_FLAG = True
            End If
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyMultiply, vbKeyInsert
            If grdsales.Rows <= 1 Then Exit Sub
            If Trim(grdsales.TextMatrix(grdsales.Rows - 1, 38)) = "" Then
                TXTITEMCODE.Text = grdsales.TextMatrix(grdsales.Rows - 1, 1)
                TXTPRODUCT.Text = grdsales.TextMatrix(grdsales.Rows - 1, 2)
                Call TXTPRODUCT_KeyDown(13, 0)
            Else
                TxtBarcode.Text = Trim(grdsales.TextMatrix(grdsales.Rows - 1, 38))
            End If
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCess_GotFocus()
    Call CHANGEBOXCOLOR(txtCess, True)
    txtCess.SelStart = 0
    txtCess.SelLength = Len(txtCess.Text)
End Sub

Private Sub txtCess_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.SetFocus
        Case vbKeyEscape
            TxtCessPer.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtCess_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCess_LostFocus()
    Call CHANGEBOXCOLOR(txtCess, False)
    Call TxttaxMRP_LostFocus
End Sub

Private Function print_labels(I As Long, BAR_LABEL As String, itemname As String, itemmrp As Double, itemprice As Double)
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
        
    
    Exit Function
ErrHand:
    MsgBox Err.Description
End Function

Private Sub TxtCessPer_GotFocus()
    Call CHANGEBOXCOLOR(TxtCessPer, True)
    TxtCessPer.SelStart = 0
    TxtCessPer.SelLength = Len(TxtCessPer.Text)
End Sub

Private Sub TxtCessPer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtCess.SetFocus
        Case vbKeyEscape
            Frame1.Enabled = True
            If OptComper.value = True Then
                TxtComper.SetFocus
            Else
                TxtComAmt.SetFocus
            End If

        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtCessPer_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCessPer_LostFocus()
    Call CHANGEBOXCOLOR(TxtCessPer, False)
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TxtHSN_GotFocus()
    Call CHANGEBOXCOLOR(txtHSN, True)
    txtHSN.SelStart = 0
    txtHSN.SelLength = Len(txtHSN.Text)
End Sub

Private Sub TxtHSN_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtHSN.Text) = "" And MDIMAIN.lblgst.Caption <> "N" And Trim(UCase(lblcategory.Caption)) <> "SERVICE CHARGE" Then
                If MsgBox("HSN Code not entered. Are you sure?", vbYesNo + vbDefaultButton2, "PURCHASE ENTRY") = vbNo Then Exit Sub
            End If
            txtPD.Enabled = True
            txtPD.SetFocus
         Case vbKeyEscape
            TxttaxMRP.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtHSN_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCustDisc_GotFocus()
    Call CHANGEBOXCOLOR(TxtCustDisc, True)
    TxtCustDisc.SelStart = 0
    TxtCustDisc.SelLength = Len(TxtCustDisc.Text)
End Sub

Private Sub TxtCustDisc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Frame1.Enabled = True
            If OptComper.value = True Then
                TxtComper.SetFocus
            Else
                TxtComAmt.SetFocus
            End If
        Case vbKeyEscape
            TxtLWRate.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtCustDisc_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCustDisc_LostFocus()
    Call CHANGEBOXCOLOR(TxtCustDisc, False)
    TxtCustDisc.Text = Format(TxtCustDisc.Text, "0.00")
End Sub


Private Sub txtNetrate_GotFocus()
    Call CHANGEBOXCOLOR(TxtNetrate, True)
    TxtNetrate.SelStart = 0
    TxtNetrate.SelLength = Len(TxtNetrate.Text)
End Sub

Private Sub txtNetrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TxtNetrate.Text) = 0 Then Exit Sub
            If Trim(txtHSN.Text) = "" Then
                txtHSN.Enabled = True
                txtHSN.SetFocus
            Else
                txtPD.SetFocus
            End If
        Case vbKeyEscape
            TxttaxMRP.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtNetrate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNetrate_LostFocus()
    Call CHANGEBOXCOLOR(TxtNetrate, False)
    If Val(TxtNetrate.Text) <> 0 Then
        TxtNetrate.Text = Format(TxtNetrate.Text, ".00")
        TXTPTR.Tag = Format(Round(Val(TxtNetrate.Text) * 100 / (100 - Val(txtPD.Text)), 4), "0.0000")
        TXTPTR.Text = Format(Round(Val(TXTPTR.Tag) * 100 / (Val(TxttaxMRP.Text) + 100), 4), "0.0000")
    End If
    Call TxttaxMRP_LostFocus
    
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            TxtRetailPercent.Text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtWsalePercent.Text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then
            txtSchPercent.Text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.Text) > 0 Then Call txtSchPercent_LostFocus
    End If
    
End Sub

Private Sub CHANGEBOXCOLOR(BOX As TextBox, texton As Boolean)
    If texton Then
        BOX.BackColor = &H98F3C1
    Else
        BOX.BackColor = vbWhite
    End If
End Sub

Private Function CLEAR_COMBO()
    CMBDISTRICT.Clear
    Dim rstfillcombo As ADODB.Recordset
    
    On Error GoTo ErrHand
    Set rstfillcombo = New ADODB.Recordset
    rstfillcombo.Open "Select DISTINCT TRX_GODOWN From TRANSMAST ORDER BY TRX_GODOWN", db, adOpenStatic, adLockReadOnly
    Do Until rstfillcombo.EOF
        If Not IsNull(rstfillcombo!TRX_GODOWN) Then CMBDISTRICT.AddItem (rstfillcombo!TRX_GODOWN)
        rstfillcombo.MoveNext
    Loop
    rstfillcombo.Close
    Set rstfillcombo = Nothing
    Exit Function
ErrHand:
    MsgBox Err.Description
End Function

Private Function find_small_number()
    Dim I As Integer
    Dim sum_ary As Double
    Dim grossamt As Double
    
    On Error GoTo ErrHand
    sum_ary = 0
    grossamt = 0
    For I = 1 To grdsales.Rows - 1
        'If Aray(i) < sn Then sn = Aray(i)
        sum_ary = sum_ary + Val(grdsales.TextMatrix(I, 3))
    Next I
    
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    For I = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(I, 32) = Round((Val(TxtTotalexp.Text) / sum_ary) * Val(grdsales.TextMatrix(I, 3)), 3)
        grossamt = Round((Val(grdsales.TextMatrix(I, 3)) - Val(grdsales.TextMatrix(I, 14))) * (Val(grdsales.TextMatrix(I, 9)) * Val(grdsales.TextMatrix(I, 5))), 3)
        'grdsales.TextMatrix(i, 8) = Format(Round(((grossamt / (Val(grdsales.TextMatrix(i, 5)) * (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))))) + ((Val(grdsales.TextMatrix(i, 32)) / Val(grdsales.TextMatrix(i, 5))))), 4), ".0000")
        grdsales.TextMatrix(I, 8) = Format(Round(((grossamt / (Val(grdsales.TextMatrix(I, 5)) * (Val(grdsales.TextMatrix(I, 3)) + Val(grdsales.TextMatrix(I, 14))))) + ((Val(grdsales.TextMatrix(I, 32)) / ((Val(grdsales.TextMatrix(I, 3)) + Val(grdsales.TextMatrix(I, 14))) * Val(grdsales.TextMatrix(I, 5)))))), 4), ".0000")
        LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(I, 32)), ".00")
        lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(I, 3)), ".00")
        db.Execute "Update RTRXFILE set ITEM_COST = " & Val(grdsales.TextMatrix(I, 8)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(I, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(I, 16)) & "  "
        db.Execute "Update RTRXFILE set EXPENSE = " & Val(grdsales.TextMatrix(I, 32)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(I, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(I, 16)) & "  "
       
    Next I
    Exit Function
ErrHand:
    'MsgBox "Smallest Number is: " & sn
End Function

Private Sub TXTRCVDATE_GotFocus()
    If Not IsDate(TXTRCVDATE.Text) And IsDate(TXTINVDATE.Text) Then
       TXTRCVDATE.Text = Format(TXTINVDATE, "DD/MM/YYYY")
    End If
    TXTRCVDATE.BackColor = &H98F3C1
    TXTRCVDATE.SelStart = 0
    TXTRCVDATE.SelLength = Len(TXTRCVDATE.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTRCVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTRCVDATE.Text = "  /  /    " Then
                TXTRCVDATE.Text = Format(Date, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTRCVDATE.Text) Then
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTRCVDATE.SetFocus
                Exit Sub
            End If
            
'            If (DateValue(TXTRCVDATE.Text) < DateValue(MDIMAIN.DTFROM.value)) Or (DateValue(TXTRCVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.value))) Then
'                'db.Execute "delete from Users"
'                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
'                TXTRCVDATE.SetFocus
'                Exit Sub
'            End If
            If Not IsDate(TXTRCVDATE.Text) Then
                TXTRCVDATE.SetFocus
            Else
                TXTRCVDATE.Text = Format(TXTRCVDATE.Text, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
            End If
        Case vbKeyEscape
            TXTINVOICE.SetFocus
    End Select
End Sub

Private Sub TXTRCVDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRCVDATE_LostFocus()
    TXTRCVDATE.BackColor = vbWhite
End Sub
