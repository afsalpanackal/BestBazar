VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExpiryRet 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expiry Return by the Customer"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   19590
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   2415
      TabIndex        =   92
      Top             =   1650
      Visible         =   0   'False
      Width           =   12570
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   3840
         Left            =   30
         TabIndex        =   93
         Top             =   390
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6773
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
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Index           =   2
         Left            =   3795
         TabIndex        =   95
         Top             =   15
         Width           =   8715
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   "SOLD RATES FOR THE ITEM "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Index           =   1
         Left            =   30
         TabIndex        =   94
         Top             =   15
         Width           =   3780
      End
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
      Height          =   315
      Left            =   1275
      TabIndex        =   42
      Top             =   165
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
      Height          =   450
      Left            =   5490
      TabIndex        =   11
      Top             =   7740
      Width           =   1200
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   2100
      TabIndex        =   26
      Top             =   2070
      Visible         =   0   'False
      Width           =   8070
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   3810
         Left            =   15
         TabIndex        =   27
         Top             =   15
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   6720
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
   Begin VB.Frame Fram 
      BackColor       =   &H00D0ECF2&
      Height          =   9750
      Left            =   -120
      TabIndex        =   12
      Top             =   -45
      Width           =   19665
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
         Height          =   465
         Left            =   15150
         TabIndex        =   123
         Top             =   6375
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
         Height          =   390
         Left            =   15150
         TabIndex        =   122
         Top             =   5970
         Width           =   1320
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00D0ECF2&
         Caption         =   "Actual Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1605
         Left            =   9510
         TabIndex        =   116
         Top             =   -15
         Width           =   4305
         Begin VB.TextBox TxtPhone 
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
            Left            =   1245
            MaxLength       =   35
            TabIndex        =   118
            Top             =   1260
            Width           =   3000
         End
         Begin VB.TextBox TXTTIN 
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
            Left            =   1245
            MaxLength       =   35
            TabIndex        =   117
            Top             =   915
            Width           =   3000
         End
         Begin MSForms.TextBox TxtBillAddress 
            Height          =   690
            Left            =   45
            TabIndex        =   121
            Top             =   210
            Width           =   4200
            VariousPropertyBits=   -1400879077
            MaxLength       =   150
            BorderStyle     =   1
            Size            =   "7408;1217"
            SpecialEffect   =   0
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            Left            =   90
            TabIndex        =   120
            Top             =   1290
            Width           =   660
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "GST No"
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
            Index           =   41
            Left            =   75
            TabIndex        =   119
            Top             =   960
            Width           =   660
         End
      End
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00D0ECF2&
         Height          =   1575
         Left            =   150
         TabIndex        =   31
         Top             =   15
         Width           =   9345
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
            TabIndex        =   73
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
            Left            =   9225
            TabIndex        =   40
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
            Height          =   300
            Left            =   6630
            MaxLength       =   100
            TabIndex        =   37
            Top             =   1230
            Width           =   2670
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
            Left            =   3720
            MaxLength       =   10
            TabIndex        =   36
            Top             =   210
            Width           =   1260
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   6630
            TabIndex        =   39
            Top             =   885
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
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
            Left            =   1245
            TabIndex        =   74
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
         Begin VB.Label lbltype 
            Caption         =   "ER"
            Height          =   375
            Left            =   5535
            TabIndex        =   115
            Top             =   195
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblcredit 
            Height          =   525
            Left            =   9480
            TabIndex        =   61
            Top             =   645
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "LAST BILL"
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
            Left            =   7965
            TabIndex        =   41
            Top             =   150
            Visible         =   0   'False
            Width           =   1215
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
            Left            =   5280
            TabIndex        =   38
            Top             =   1260
            Width           =   1290
         End
         Begin VB.Label INVDATE 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   0
            Left            =   2595
            TabIndex        =   35
            Top             =   210
            Width           =   1110
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "NO."
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
            TabIndex        =   34
            Top             =   210
            Width           =   870
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Return Date"
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
            Left            =   5280
            TabIndex        =   33
            Top             =   945
            Width           =   1335
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   5
            Left            =   150
            TabIndex        =   32
            Top             =   600
            Width           =   1005
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdsales 
         Height          =   4275
         Left            =   135
         TabIndex        =   85
         Top             =   1590
         Width           =   19395
         _ExtentX        =   34211
         _ExtentY        =   7541
         _Version        =   393216
         Rows            =   1
         Cols            =   32
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
      End
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00D0ECF2&
         Height          =   3645
         Left            =   135
         TabIndex        =   13
         Top             =   5790
         Width           =   15000
         Begin VB.TextBox txtcategory 
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
            Height          =   360
            Left            =   2085
            TabIndex        =   124
            Top             =   480
            Width           =   1260
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Print"
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
            Left            =   3420
            TabIndex        =   9
            Top             =   1980
            Width           =   1020
         End
         Begin VB.TextBox TxtWarranty 
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
            Left            =   13485
            MaxLength       =   4
            TabIndex        =   113
            Top             =   2610
            Visible         =   0   'False
            Width           =   795
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
            Left            =   14280
            Style           =   2  'Dropdown List
            TabIndex        =   112
            Top             =   2610
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox TxtRetailPercent 
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
            Height          =   405
            Left            =   6735
            MaxLength       =   7
            TabIndex        =   110
            Top             =   3600
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox txtWsalePercent 
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
            Height          =   405
            Left            =   7815
            MaxLength       =   7
            TabIndex        =   109
            Top             =   3600
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtSchPercent 
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
            Height          =   405
            Left            =   12645
            MaxLength       =   7
            TabIndex        =   108
            Top             =   2940
            Visible         =   0   'False
            Width           =   1275
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
            ItemData        =   "FrmExpiryRet.frx":0000
            Left            =   8970
            List            =   "FrmExpiryRet.frx":0034
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   480
            Width           =   1200
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
            Left            =   8250
            MaxLength       =   7
            TabIndex        =   99
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox Txtgrossamt 
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
            Height          =   375
            Left            =   5910
            MaxLength       =   7
            TabIndex        =   90
            Top             =   1140
            Width           =   1560
         End
         Begin VB.TextBox txtvanrate 
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
            Left            =   12645
            MaxLength       =   7
            TabIndex        =   88
            Top             =   2610
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox txtcrtnpack 
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
            Left            =   12975
            MaxLength       =   7
            TabIndex        =   86
            Top             =   2610
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox TxtComper 
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
            Left            =   8910
            MaxLength       =   7
            TabIndex        =   82
            Top             =   3225
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox TxtComAmt 
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
            Left            =   9720
            MaxLength       =   7
            TabIndex        =   81
            Top             =   3225
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox txtcrtn 
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
            Left            =   11910
            MaxLength       =   7
            TabIndex        =   79
            Top             =   2610
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtWS 
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
            Height          =   375
            Left            =   7815
            MaxLength       =   7
            TabIndex        =   77
            Top             =   3210
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox TXTRETAIL 
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
            Left            =   6735
            MaxLength       =   7
            TabIndex        =   75
            Top             =   3225
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox txtPD 
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
            Height          =   375
            Left            =   2085
            MaxLength       =   7
            TabIndex        =   69
            Top             =   1140
            Width           =   960
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
            TabIndex        =   68
            Top             =   2505
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
            Left            =   14970
            MaxLength       =   7
            TabIndex        =   66
            Top             =   3255
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
            Height          =   405
            Left            =   11025
            TabIndex        =   63
            Top             =   2040
            Width           =   1095
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
            Height          =   420
            Left            =   12165
            TabIndex        =   62
            Top             =   2025
            Width           =   1095
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
            Height          =   405
            Left            =   9855
            TabIndex        =   55
            Top             =   2040
            Width           =   1095
         End
         Begin VB.OptionButton OPTNET 
            BackColor       =   &H00D0ECF2&
            Caption         =   "NET"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3945
            TabIndex        =   52
            Top             =   1575
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.TextBox TxtFree 
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
            Left            =   12630
            MaxLength       =   7
            TabIndex        =   53
            Top             =   3090
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00D0ECF2&
            Caption         =   "Tax on MRP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   5235
            TabIndex        =   50
            Top             =   1515
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00D0ECF2&
            Caption         =   "TAX %"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2685
            TabIndex        =   51
            Top             =   1575
            Width           =   1140
         End
         Begin VB.TextBox TxttaxMRP 
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
            Left            =   1245
            MaxLength       =   7
            TabIndex        =   46
            Top             =   1155
            Width           =   810
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
            Left            =   5925
            MaxLength       =   7
            TabIndex        =   44
            Top             =   3255
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   60
            MaxLength       =   7
            TabIndex        =   30
            Top             =   1155
            Width           =   1170
         End
         Begin VB.TextBox TXTRATE 
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
            Left            =   13620
            MaxLength       =   7
            TabIndex        =   29
            Top             =   480
            Width           =   1050
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
            Height          =   450
            Left            =   75
            TabIndex        =   6
            Top             =   1980
            Width           =   1125
         End
         Begin VB.TextBox TXTSLNO 
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
            Left            =   45
            TabIndex        =   0
            Top             =   480
            Width           =   540
         End
         Begin VB.TextBox TXTPRODUCT 
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
            Left            =   3360
            TabIndex        =   1
            Top             =   480
            Width           =   4875
         End
         Begin VB.TextBox TXTQTY 
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
            Left            =   10170
            MaxLength       =   9
            TabIndex        =   2
            Top             =   480
            Width           =   915
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
            Left            =   2355
            TabIndex        =   8
            Top             =   1980
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
            Height          =   450
            Left            =   1245
            TabIndex        =   7
            Top             =   1980
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
            Left            =   600
            TabIndex        =   15
            Top             =   480
            Width           =   1470
         End
         Begin VB.TextBox txtBatch 
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
            Left            =   11100
            MaxLength       =   15
            TabIndex        =   4
            Top             =   480
            Width           =   1275
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
            TabIndex        =   14
            Top             =   3195
            Visible         =   0   'False
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
            Height          =   450
            Left            =   4455
            TabIndex        =   10
            Top             =   1980
            Width           =   975
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   360
            Left            =   12390
            TabIndex        =   3
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
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
            Height          =   360
            Left            =   12390
            TabIndex        =   43
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
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
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   9315
            TabIndex        =   103
            Top             =   3645
            Visible         =   0   'False
            Width           =   2595
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
               TabIndex        =   105
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
               TabIndex        =   104
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
               TabIndex        =   107
               Top             =   525
               Width           =   1470
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
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
               Height          =   300
               Index           =   36
               Left            =   90
               TabIndex        =   106
               Top             =   195
               Width           =   1050
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00D0ECF2&
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   75
            TabIndex        =   96
            Top             =   1455
            Width           =   2580
            Begin VB.OptionButton optdiscper 
               BackColor       =   &H00D0ECF2&
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
               Height          =   240
               Left            =   60
               TabIndex        =   98
               Top             =   150
               Value           =   -1  'True
               Width           =   1170
            End
            Begin VB.OptionButton Optdiscamt 
               BackColor       =   &H00D0ECF2&
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
               Left            =   1365
               TabIndex        =   97
               Top             =   150
               Width           =   1155
            End
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
            Left            =   2085
            TabIndex        =   125
            Top             =   195
            Width           =   1290
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
            Height          =   285
            Index           =   39
            Left            =   13485
            TabIndex        =   114
            Top             =   2355
            Visible         =   0   'False
            Width           =   2130
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
            Height          =   405
            Index           =   38
            Left            =   5745
            TabIndex        =   111
            Top             =   3600
            Visible         =   0   'False
            Width           =   960
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
            Left            =   600
            TabIndex        =   102
            Top             =   195
            Width           =   1470
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3060
            TabIndex        =   49
            Top             =   1140
            Width           =   1305
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Pack"
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
            Left            =   8250
            TabIndex        =   100
            Top             =   195
            Width           =   1905
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
            Height          =   255
            Index           =   33
            Left            =   5910
            TabIndex        =   91
            Top             =   885
            Width           =   1560
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Scheme Rate"
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
            Left            =   12645
            TabIndex        =   89
            Top             =   2355
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Case Pack"
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
            Left            =   12975
            TabIndex        =   87
            Top             =   2355
            Visible         =   0   'False
            Width           =   1020
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
            Left            =   8910
            TabIndex        =   84
            Top             =   2970
            Visible         =   0   'False
            Width           =   795
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
            Left            =   9720
            TabIndex        =   83
            Top             =   2970
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Case Rate"
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
            Left            =   11910
            TabIndex        =   80
            Top             =   2355
            Visible         =   0   'False
            Width           =   1035
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
            Index           =   27
            Left            =   7815
            TabIndex        =   78
            Top             =   2970
            Visible         =   0   'False
            Width           =   1065
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
            Left            =   14970
            TabIndex        =   76
            Top             =   3015
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   25
            Left            =   2085
            TabIndex        =   70
            Top             =   885
            Width           =   960
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
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   24
            Left            =   6735
            TabIndex        =   67
            Top             =   2970
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Addnl Amt"
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
            Left            =   11055
            TabIndex        =   65
            Top             =   1770
            Width           =   1050
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
            Left            =   12180
            TabIndex        =   64
            Top             =   1770
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
            Height          =   225
            Index           =   21
            Left            =   11820
            TabIndex        =   60
            Top             =   885
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
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   585
            Left            =   11595
            TabIndex        =   59
            Top             =   1125
            Width           =   1695
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
            Left            =   9930
            TabIndex        =   58
            Top             =   1770
            Width           =   990
         End
         Begin VB.Label lbltotalwodiscount 
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
            Height          =   585
            Left            =   9870
            TabIndex        =   57
            Top             =   1125
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "RETURN AMOUNT"
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
            Height          =   240
            Index           =   6
            Left            =   9900
            TabIndex        =   56
            Top             =   885
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
            Left            =   12630
            TabIndex        =   54
            Top             =   2805
            Visible         =   0   'False
            Width           =   810
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
            Height          =   255
            Index           =   13
            Left            =   3060
            TabIndex        =   48
            Top             =   885
            Width           =   1305
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
            Left            =   1245
            TabIndex        =   47
            Top             =   885
            Width           =   810
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
            Left            =   5925
            TabIndex        =   45
            Top             =   2970
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
            Left            =   45
            TabIndex        =   28
            Top             =   885
            Width           =   1170
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
            TabIndex        =   25
            Top             =   195
            Width           =   540
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
            Left            =   3360
            TabIndex        =   24
            Top             =   195
            Width           =   4875
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
            Left            =   10170
            TabIndex        =   23
            Top             =   210
            Width           =   915
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
            Height          =   285
            Index           =   11
            Left            =   13620
            TabIndex        =   22
            Top             =   195
            Width           =   1050
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
            Left            =   4380
            TabIndex        =   21
            Top             =   885
            Width           =   1500
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
            TabIndex        =   20
            Top             =   3135
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
            Left            =   12390
            TabIndex        =   19
            Top             =   195
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Batch No."
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
            Index           =   7
            Left            =   11100
            TabIndex        =   18
            Top             =   195
            Width           =   1275
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
            Height          =   375
            Left            =   4380
            TabIndex        =   5
            Top             =   1140
            Width           =   1500
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
            TabIndex        =   17
            Top             =   2925
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
            TabIndex        =   16
            Top             =   2865
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   135
         TabIndex        =   72
         Top             =   1575
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   705
         TabIndex        =   71
         Top             =   1575
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmExpiryRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim ACT_FLAG As Boolean
Dim PHY_CODE As New ADODB.Recordset
Dim PHYCODE_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT, M_ADD As Boolean
Dim PHY_PRERATE As New ADODB.Recordset
Dim PRERATE_FLAG As Boolean
Dim CHANGE_FLAG As Boolean

Private Sub CmbPack_GotFocus()
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = "1"
            CmbPack.Enabled = False
            TXTQTY.Enabled = True
            Call FILL_PREVIIOUSRATE
            'TXTQTY.SetFocus
         Case vbKeyEscape
            'TXTUNIT.Text = ""
            CmbPack.Enabled = False
            Los_Pack.Enabled = True
            Los_Pack.SetFocus
    End Select
End Sub

Private Sub CMDADD_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim M_DATA As Long

    M_DATA = 0
    
    Txtpack.text = 1
    If grdsales.rows <= Val(TXTSLNO.text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.text), 0) = Val(TXTSLNO.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 1) = Trim(TXTITEMCODE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 2) = Trim(TXTPRODUCT.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 3) = Val(TXTQTY.text) + Val(TXTFREE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 4) = 1 'Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 5) = Val(Los_Pack.text) ' 1 'Val(TxtPack.Text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 6) = Format(Round(Val(TXTRATE.text) / Val(Los_Pack.text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 8) = Format(Round((Val(LBLSUBTOTAL.Caption) / (Val(Los_Pack.text) * TXTQTY.text)), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 9) = Format(Round(Val(TXTPTR.text) / Val(Los_Pack.text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 7) = Format((Val(txtprofit.text)), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 10) = IIf(Val(TxttaxMRP.text) = 0, "", Format(Val(TxttaxMRP.text), ".00")) 'TAX
    grdsales.TextMatrix(Val(TXTSLNO.text), 11) = Trim(txtBatch.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 12) = IIf(Trim(TXTEXPDATE.text) = "/  /", "", TXTEXPDATE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 13) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 14) = Val(TXTFREE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 17) = Val(txtPD.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 26) = Format(Val(Txtgrossamt.text), ".00")
    If optdiscper.Value = True Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 27) = "P"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.text), 27) = "A"
    End If
    grdsales.TextMatrix(Val(TXTSLNO.text), 28) = Format(Val(Los_Pack.text), ".00")
    grdsales.TextMatrix(Val(TXTSLNO.text), 29) = Trim(CmbPack.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 30) = Val(TxtWarranty.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 31) = Trim(CmbWrnty.text)
    If Val(TxttaxMRP.text) = 0 Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 15) = "N"
    Else
        If OPTTaxMRP.Value = True Then
            grdsales.TextMatrix(Val(TXTSLNO.text), 15) = "M"
        ElseIf OPTVAT.Value = True Then
            grdsales.TextMatrix(Val(TXTSLNO.text), 15) = "V"
        End If
    End If

    If M_EDIT = True Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 16) = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16))
    Else
        grdsales.TextMatrix(Val(TXTSLNO.text), 16) = Val(TXTSLNO.text)
    End If
    
'    If OLD_BILL = False Then Call checklastbill
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!TRX_TYPE = "ER"
        RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTRTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTRTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16))
        RSTRTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1))
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
    End If
    RSTRTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 13))
    RSTRTRXFILE!VCH_DATE = Format(TXTINVDATE, "dd/mm/yyyy")
    RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 2))
    RSTRTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
    RSTRTRXFILE!P_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 17))
    RSTRTRXFILE!MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6))
    RSTRTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 9))
    RSTRTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 7))
    RSTRTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18))
    RSTRTRXFILE!P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 19))
    RSTRTRXFILE!P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))
    RSTRTRXFILE!CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24))
    RSTRTRXFILE!P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 25))
    RSTRTRXFILE!GROSS_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 26))
    If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 23)) = "A" Then
        RSTRTRXFILE!COM_FLAG = "A"
        RSTRTRXFILE!COM_PER = 0
        RSTRTRXFILE!COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 22))
    Else
        RSTRTRXFILE!COM_FLAG = "P"
        RSTRTRXFILE!COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 21))
        RSTRTRXFILE!COM_AMT = 0
    End If
    RSTRTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))
    RSTRTRXFILE!LOOSE_PACK = Val(Los_Pack.text)
    RSTRTRXFILE!PACK_TYPE = Trim(CmbPack.text)
    RSTRTRXFILE!WARRANTY = Val(TxtWarranty.text)
    RSTRTRXFILE!WARRANTY_TYPE = Trim(CmbWrnty.text)
    RSTRTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
    'RSTRTRXFILE!VCH_DESC = "Received From " & DataList2.Text
    RSTRTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 11))
    'RSTRTRXFILE!ISSUE_QTY = 0
    RSTRTRXFILE!CST = 0
    If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 27)) = "P" Then
        RSTRTRXFILE!DISC_FLAG = "P"
    Else
        RSTRTRXFILE!DISC_FLAG = "A"
    End If
    RSTRTRXFILE!SCHEME = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 14))
    RSTRTRXFILE!EXP_DATE = IIf(grdsales.TextMatrix(Val(TXTSLNO.text), 12) = "", Null, Format(grdsales.TextMatrix(Val(TXTSLNO.text), 12), "dd/mm/yyyy"))
    RSTRTRXFILE!FREE_QTY = 0
    RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!C_USER_ID = "SM"
    RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15))

    'RSTRTRXFILE!M_USER_ID = DataList2.BoundText
    ''''RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))  'MODE OF TAX
    'RSTRTRXFILE!PINV = Trim(TXTINVOICE.Text)
    RSTRTRXFILE.Update
    RSTRTRXFILE.Close

    M_DATA = 0
    Set RSTRTRXFILE = Nothing


    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    For i = 1 To grdsales.rows - 1
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
    Next i
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    TXTSLNO.text = grdsales.rows
    TXTPRODUCT.text = ""

    TXTITEMCODE.text = ""
    TXTPTR.text = ""
    Txtgrossamt.text = ""
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    txtPD.text = ""
    txtprofit.text = ""
    txtretail.text = ""
    TxtRetailPercent.text = ""
    txtWsalePercent.text = ""
    txtSchPercent.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    Txtgrossamt.text = ""
    txtcrtn.text = ""
    txtcrtnpack.text = ""
    TXTRATE.text = ""
    TxtComAmt.text = ""
    TxtComper.text = ""
    txtmrpbt.text = ""
    txtBatch.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    optnet.Value = True
    M_EDIT = False
    M_ADD = True
    cmdRefresh.Enabled = True
    CMDPRINT.Enabled = True
    TXTITEMCODE.Enabled = True
    TXTITEMCODE.SetFocus
    txtBillNo.Enabled = False

    If grdsales.rows >= 18 Then grdsales.TopRow = grdsales.rows - 1

End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset

    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub

    On Error GoTo ErrHand
    db.Execute "delete  From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & ""
    i = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(LINE_NO) From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing

    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTRTRXFILE.EOF
        RSTRTRXFILE!LINE_NO = i
        i = i + 1
        RSTRTRXFILE.Update
        RSTRTRXFILE.MoveNext
    Loop
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing

    i = 1
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTRTRXFILE.EOF
        RSTRTRXFILE!LINE_NO = i
        i = i + 1
        RSTRTRXFILE.Update
        RSTRTRXFILE.MoveNext
    Loop
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing

    grdsales.rows = 1
    i = 0
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    grdsales.rows = 1
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "Select * From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTRTRXFILE.EOF
        grdsales.rows = grdsales.rows + 1
        grdsales.FixedRows = 1
        i = i + 1

        grdsales.TextMatrix(i, 0) = i
        grdsales.TextMatrix(i, 1) = RSTRTRXFILE!ITEM_CODE
        grdsales.TextMatrix(i, 2) = RSTRTRXFILE!ITEM_NAME
        grdsales.TextMatrix(i, 3) = Val(RSTRTRXFILE!QTY) / Val(RSTRTRXFILE!LINE_DISC)
        grdsales.TextMatrix(i, 4) = RSTRTRXFILE!UNIT
        grdsales.TextMatrix(i, 5) = RSTRTRXFILE!LINE_DISC
        grdsales.TextMatrix(i, 6) = Format(RSTRTRXFILE!MRP, ".000")
        grdsales.TextMatrix(i, 7) = Format(RSTRTRXFILE!SALES_PRICE, ".000")
        grdsales.TextMatrix(i, 8) = Format(RSTRTRXFILE!ITEM_COST, ".000")
        grdsales.TextMatrix(i, 9) = Format(RSTRTRXFILE!PTR, ".000")
        grdsales.TextMatrix(i, 10) = IIf(Val(RSTRTRXFILE!SALES_TAX) = 0, "", Format(RSTRTRXFILE!SALES_TAX, ".00"))
        grdsales.TextMatrix(i, 11) = RSTRTRXFILE!REF_NO
        grdsales.TextMatrix(i, 12) = Format(RSTRTRXFILE!EXP_DATE, "DD/MM/YYYY")
        grdsales.TextMatrix(i, 13) = Format(RSTRTRXFILE!TRX_TOTAL, ".000")
        grdsales.TextMatrix(i, 14) = IIf(IsNull(RSTRTRXFILE!SCHEME), "", RSTRTRXFILE!SCHEME)
        grdsales.TextMatrix(i, 15) = IIf(IsNull(RSTRTRXFILE!CHECK_FLAG), "N", RSTRTRXFILE!CHECK_FLAG)
        grdsales.TextMatrix(i, 16) = RSTRTRXFILE!LINE_NO
        grdsales.TextMatrix(i, 17) = IIf(IsNull(RSTRTRXFILE!P_DISC), 0, RSTRTRXFILE!P_DISC)
        grdsales.TextMatrix(i, 18) = IIf(IsNull(RSTRTRXFILE!P_RETAIL), 0, RSTRTRXFILE!P_RETAIL)
        grdsales.TextMatrix(i, 19) = IIf(IsNull(RSTRTRXFILE!P_WS), 0, RSTRTRXFILE!P_WS)
        grdsales.TextMatrix(i, 20) = IIf(IsNull(RSTRTRXFILE!P_CRTN), 0, RSTRTRXFILE!P_CRTN)
        If RSTRTRXFILE!COM_FLAG = "A" Then
            grdsales.TextMatrix(i, 21) = 0
            grdsales.TextMatrix(i, 22) = IIf(IsNull(RSTRTRXFILE!COM_AMT), 0, RSTRTRXFILE!COM_AMT)
            grdsales.TextMatrix(i, 23) = "A"
        Else
            grdsales.TextMatrix(i, 21) = IIf(IsNull(RSTRTRXFILE!COM_PER), 0, RSTRTRXFILE!COM_PER)
            grdsales.TextMatrix(i, 22) = 0
            grdsales.TextMatrix(i, 23) = "P"
        End If
        If RSTRTRXFILE!DISC_FLAG = "P" Then
            grdsales.TextMatrix(i, 27) = "P"
        Else
            grdsales.TextMatrix(i, 27) = "A"
        End If
        grdsales.TextMatrix(i, 24) = IIf(IsNull(RSTRTRXFILE!CRTN_PACK), 0, RSTRTRXFILE!CRTN_PACK)
        grdsales.TextMatrix(i, 25) = IIf(IsNull(RSTRTRXFILE!P_VAN), 0, RSTRTRXFILE!P_VAN)
        grdsales.TextMatrix(i, 26) = IIf(IsNull(RSTRTRXFILE!GROSS_AMT), 0, RSTRTRXFILE!GROSS_AMT)
        grdsales.TextMatrix(i, 28) = IIf(IsNull(RSTRTRXFILE!LOOSE_PACK), 1, RSTRTRXFILE!LOOSE_PACK)
        grdsales.TextMatrix(i, 29) = IIf(IsNull(RSTRTRXFILE!PACK_TYPE), "Nos", RSTRTRXFILE!PACK_TYPE)
        grdsales.TextMatrix(i, 30) = IIf(IsNull(RSTRTRXFILE!WARRANTY), "", RSTRTRXFILE!WARRANTY)
        grdsales.TextMatrix(i, 31) = IIf(IsNull(RSTRTRXFILE!WARRANTY_TYPE), "", RSTRTRXFILE!WARRANTY_TYPE)
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
        'TXTDEALER.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)

        'TXTINVDATE.Text = Format(RSTRTRXFILE!VCH_DATE, "DD/MM/YYYY")
        'TXTREMARKS.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
        'TXTINVOICE.Text = IIf(IsNull(RSTRTRXFILE!PINV), "", RSTRTRXFILE!PINV)
        RSTRTRXFILE.MoveNext
    Loop
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing

    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")

    TXTSLNO.text = Val(grdsales.rows)
    TXTPRODUCT.text = ""
    TXTITEMCODE.text = ""
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    txtPD.text = ""
    txtprofit.text = ""
    txtretail.text = ""
    TxtRetailPercent.text = ""
    txtWsalePercent.text = ""
    txtSchPercent.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    Txtgrossamt.text = ""
    txtcrtn.text = ""
    txtcrtnpack.text = ""
    TXTRATE.text = ""
    TxtComAmt.text = ""
    TxtComper.text = ""
    txtmrpbt.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    txtBatch.text = ""
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_ADD = True
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub CmdDeleteAll_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    On Error GoTo ErrHand
    If Chkcancel.Value = 0 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE ALL", vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then Exit Sub
    
    grdsales.FixedRows = 0
    grdsales.rows = 1
    Call appendpurchase
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDMODIFY_Click()

    If Val(TXTSLNO.text) >= grdsales.rows Then Exit Sub

    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTQTY.Enabled = True
    TXTQTY.SetFocus

End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            TXTQTY.text = ""
            Txtpack.text = 1 '""
            Los_Pack.text = ""
            CmbPack.ListIndex = -1
            TxtWarranty.text = ""
            CmbWrnty.ListIndex = -1
            TXTFREE.text = ""
            TxttaxMRP.text = ""
            txtPD.text = ""
            txtprofit.text = ""
            txtretail.text = ""
            TxtRetailPercent.text = ""
            txtWsalePercent.text = ""
            txtSchPercent.text = ""
            txtWS.text = ""
            txtvanrate.text = ""
            Txtgrossamt.text = ""
            txtcrtn.text = ""
            txtcrtnpack.text = ""
            TXTRATE.text = ""
            TxtComAmt.text = ""
            TxtComper.text = ""
            txtmrpbt.text = ""
            TXTITEMCODE.text = ""
            LBLSUBTOTAL.Caption = ""
            lbltaxamount.Caption = ""
            TXTEXPDATE.text = "  /  /    "
            TXTEXPIRY.text = "  /  "
            txtBatch.text = ""

            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            txtcategory.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            TXTEXPIRY.Visible = False
            txtBatch.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
            M_EDIT = False
    End Select
End Sub

Private Sub CmdPrint_Click()
    
   If grdsales.rows = 1 Then Exit Sub
    'If Month(MDIMAIN.DTFROM.value) >= 4 And Year(MDIMAIN.DTFROM.value) >= 2021 Then Exit Sub
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        DataList2.SetFocus
        Exit Sub
    End If
    
    If MDIMAIN.StatusBar.Panels(8).text <> "Y" Then
        Call Generateprint
    Else
        Call ReportGeneratION_estimate
        On Error GoTo CLOSEFILE
    Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
    End If
    On Error GoTo ErrHand
    
    Print #1, "TYPE " & Rptpath & "Report.txt > PRN"
    Print #1, "EXIT"
    Close #1
    
    '//HERE write the proper path where your command.com file exist
    Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & Rptpath & "REPO.BAT N", vbHide
    cmdRefresh.SetFocus
    'Call cmdRefresh_Click
    Screen.MousePointer = vbNormal
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    Dim b As Integer
    Dim Num As Currency

    On Error GoTo ErrHand
    b = 0
    
    db.Execute "delete From TEMPTRXFILE "
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "ER"
        'RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        
        
        RSTTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(i, 13))
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE, "dd/mm/yyyy")
        RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
        RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 17))
        'RSTTRXFILE!P_DISC = Val(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 9)) + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!P_RETAIL = (Val(grdsales.TextMatrix(i, 9)) + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)) * Val(grdsales.TextMatrix(i, 28))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 28))   ''+ (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)
        'RSTTRXFILE!P_WS = Val(grdsales.TextMatrix(i, 19))
        'RSTTRXFILE!P_CRTN = Val(grdsales.TextMatrix(i, 20))
        'RSTTRXFILE!CRTN_PACK = Val(grdsales.TextMatrix(i, 24))
        'RSTTRXFILE!P_VAN = Val(grdsales.TextMatrix(i, 25))
        'RSTTRXFILE!GROSS_AMT = Val(grdsales.TextMatrix(i, 26))
        RSTTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 28))
        RSTTRXFILE!PACK_TYPE = "Nos" 'Trim(CmbPack.Text)
        RSTTRXFILE!WARRANTY = Val(TxtWarranty.text)
        RSTTRXFILE!WARRANTY_TYPE = Trim(CmbWrnty.text)
        RSTTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(I, 4))
        'RSTTRXFILE!VCH_DESC = "Received From " & DataList2.Text
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 11))
        'RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!SCHEME = Val(grdsales.TextMatrix(i, 14))
        If IsDate(grdsales.TextMatrix(i, 12)) Then
            RSTTRXFILE!EXP_DATE = Format(grdsales.TextMatrix(i, 12), "dd/mm/yyyy")
        End If
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 15))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
        RSTTRXFILE.Update
    Next i
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    'lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption), 0)) - Val(Round(Val(LBLTOTAL.Caption), 2)), 2)
    'Num = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0))
    
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompTin, CompCST As String
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        CompName = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
        CompAddress1 = IIf(IsNull(RSTCOMPANY!Address), "", RSTCOMPANY!Address)
        CompAddress2 = IIf(IsNull(RSTCOMPANY!HO_NAME), "", RSTCOMPANY!HO_NAME)
        If Trim(CompAddress2) = "" Then
            CompAddress2 = "Ph: " & IIf(IsNull(RSTCOMPANY!TEL_NO), "", RSTCOMPANY!TEL_NO) & IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", ", " & RSTCOMPANY!FAX_NO) & _
                        IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!FAX_NO)
        Else
            CompAddress3 = "Ph: " & IIf(IsNull(RSTCOMPANY!TEL_NO), "", RSTCOMPANY!TEL_NO) & IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", ", " & RSTCOMPANY!FAX_NO) & _
                        IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!FAX_NO)
        End If
        CompTin = IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", "", "GSTIN No. " & RSTCOMPANY!CST)
        CompCST = IIf(IsNull(RSTCOMPANY!DL_NO) Or RSTCOMPANY!DL_NO = "", "", "CST No. " & RSTCOMPANY!DL_NO)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Sleep (300)
    ReportNameVar = Rptpath & "rptSR"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.RecordSelectionFormula = "( {TRXFILE.TRX_TYPE}='ER' AND {TRXFILE.VCH_NO}= " & Val(txtBillNo.Text) & " )"
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    Report.DiscardSavedData
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.text = "'" & CompName & "'"
        If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.text = "'" & CompAddress1 & "'"
        If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.text = "'" & CompAddress2 & "'"
        If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.text = "'" & CompAddress3 & "'"
        If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
        If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
        If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.text = "'" & TXTDEALER.text & "'"
        If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
        If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.text = "'" & TXTTIN.text & "'"
        If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.text = "'" & TxtPhone.text & "'"
'        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
'        If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
'        If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & TxtPhone.Text & "'"
'        If CRXFormulaField.Name = "{@VCH_NO}" Then CRXFormulaField.Text = "'" & Format(Trim(txtBillNo.text), bill_for) & "' & "'"
'        If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
'        If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = "'" & Format(Val(LBLDISCAMT.Caption), "0.00") & "'"
'        If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
'        If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.Text = "'" & Trim(TxtCode.Text) & "'"
'        If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.Text = "'" & Format(Val(txtOutstanding.Text), "0.00") & "'"
'
'        'If CRXFormulaField.Name = "{@unit}" Then CRXFormulaField.Text = "'" & Trim(lblunit.Caption) & "'"
'        If Trim(TXTTIN.Text) = "" Then
'            If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'TAX INVOICE FORM 8B'"
'        Else
'            If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'TAX INVOICE FORM 8'"
'        End If
    Next
    frmreport.Caption = "EXPIRY RETURN"
    Call GENERATEREPORT

    CMDEXIT.Enabled = False

    Exit Function
ErrHand:
    MsgBox err.Description
End Function


Private Sub cmdRefresh_Click()
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
    If (DateValue(TXTINVDATE.text) < DateValue(MDIMAIN.DTFROM.Value)) Or (DateValue(TXTINVDATE.text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.Value))) Then
        'db.Execute "delete from Users"
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If

     If IsNull(DataList2.SelectedItem) Then
         MsgBox "Select Customer From List", vbOKOnly, "Sales Return..."
         DataList2.SetFocus
         Exit Sub
     End If
     If Not IsDate(TXTINVDATE.text) Then
         MsgBox "Enter Returned Date", vbOKOnly, "EXPIRY RETURN"
         Exit Sub
     End If
     Call appendpurchase

End Sub

Private Sub Form_Activate()
    On Error GoTo ErrHand
    txtBillNo.SetFocus
    Exit Sub
ErrHand:
    If err.Number = 5 Then Exit Sub
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    'txtBillNo.Visible = False
    'cmditemcreate.Visible = False
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set TRXMAST = New ADODB.Recordset
    
    TRXMAST.Open "Select MAX(VCH_NO) From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' ", db, adOpenStatic, adLockReadOnly
    'TRXMAST.Open "Select MAX(VCH_NO) From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'ER'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    ACT_FLAG = True
    PRERATE_FLAG = True
    grdsales.ColWidth(0) = 500
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2700
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
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(17) = 800
    grdsales.ColWidth(18) = 800
    grdsales.ColWidth(19) = 800
    grdsales.ColWidth(20) = 800
    grdsales.ColWidth(21) = 800
    grdsales.ColWidth(22) = 1000
    grdsales.ColWidth(23) = 0
    grdsales.ColWidth(26) = 1700
    grdsales.ColWidth(27) = 0
    grdsales.ColWidth(28) = 0
    grdsales.ColWidth(29) = 0

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
    grdsales.TextArray(11) = "Batch NO"
    grdsales.TextArray(12) = "EXPIRY"
    grdsales.TextArray(13) = "SUB TOTAL"
    grdsales.TextArray(14) = "FREE"
    grdsales.TextArray(15) = "TAX MODE"
    grdsales.TextArray(16) = "Line No"
    grdsales.TextArray(17) = "Disc"
    grdsales.TextArray(18) = "RT Price"
    grdsales.TextArray(19) = "WS Price"
    grdsales.TextArray(20) = "Cartn Price"
    grdsales.TextArray(21) = "Comm %"
    grdsales.TextArray(22) = "Comm Amt"
    grdsales.TextArray(23) = "Comm Flag"
    grdsales.TextArray(24) = "Cnt Pck"
    grdsales.TextArray(25) = "Van Rate"
    grdsales.TextArray(26) = "GROSS AMOUNT"
    grdsales.TextArray(27) = "DISC_FLAG"

    PHYFLAG = True
    PHYCODE_FLAG = True
    TXTPRODUCT.Enabled = False
    txtcategory.Enabled = False
    TXTITEMCODE.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTDATE.text = Date
    TXTEXPDATE.Enabled = False
    txtBatch.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    TXTUNIT.Enabled = False
    TXTSLNO.text = 1
    TXTSLNO.Enabled = True
    FRMECONTROLS.Enabled = False
    FRMEMASTER.Enabled = False
    CLOSEALL = 1
    TXTDEALER.text = ""
    M_ADD = False
    'Me.Width = 15135
    'Me.Height = 9660
    Me.Left = 0
    Me.Top = 0
    
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If grdsales.rows <= 1 Then db.Execute "delete From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & ""
        If PHYFLAG = False Then PHY.Close
        If PHYCODE_FLAG = False Then PHY_CODE.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If PRERATE_FLAG = False Then PHY_PRERATE.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long

    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn

    
            TXTITEMCODE.text = grdtmp.Columns(0)
            TXTPRODUCT.text = grdtmp.Columns(1)
            For i = 1 To grdsales.rows - 1
                If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                    If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                End If
            Next i

            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "Select * From  rtntrxfile  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' ORDER BY CREATE_DATE", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                RSTRXFILE.MoveLast
                TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                If IsNull(RSTRXFILE!LINE_DISC) Then
                    Txtpack.text = ""
                Else
                    Txtpack.text = RSTRXFILE!LINE_DISC
                End If
                Txtpack.text = 1
                Los_Pack.text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                On Error Resume Next
                CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                CmbWrnty.text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                On Error GoTo ErrHand
            Else
                TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                Txtpack.text = 1
                Los_Pack.text = 1
                TxtWarranty.text = ""
                On Error Resume Next
                CmbPack.text = "Nos"
                CmbWrnty.ListIndex = -1
                TXTEXPDATE.text = IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                On Error GoTo ErrHand
                txtBatch.text = ""
                TXTEXPIRY.text = "  /  "
                TXTRATE.text = ""
                txtmrpbt.text = ""
                TXTPTR.text = ""
                txtretail.text = ""
                txtWS.text = ""
                txtvanrate.text = ""
                txtcrtn.text = ""
                txtcrtnpack.text = ""
                txtprofit.text = ""
                TxttaxMRP.text = "5"
                Los_Pack.text = "1"
                TxtWarranty.text = ""
                On Error Resume Next
                CmbPack.text = "Nos"
                CmbWrnty.ListIndex = -1
                On Error GoTo ErrHand
                OPTVAT.Value = True
            End If
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing

            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Fram.Enabled = True
            TXTPRODUCT.Enabled = False
            txtcategory.Enabled = False
            TXTITEMCODE.Enabled = False
            Los_Pack.Enabled = True
            Los_Pack.SetFocus
            'TxtPack.Enabled = True
            'TxtPack.SetFocus
        Case vbKeyEscape
            TXTQTY.text = ""
            TXTFREE.text = ""
            Fram.Enabled = True
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
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
            If Val(TxttaxMRP.text) <> 0 Then
                If OPTTaxMRP.Value = False And OPTVAT.Value = False Then
                'If OPTVAT.Value = False Then
                    MsgBox "Tax should be Zero ....", vbOKOnly, "Opening Balance"
                    TxttaxMRP.Enabled = True
                    TxttaxMRP.SetFocus
                    Exit Sub
                End If
            End If
            If TxttaxMRP.Enabled = True Then
                TxttaxMRP.Enabled = False
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

Private Sub OPTTaxMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
'            If Val(TxttaxMRP.Text) <> 0 Then
'                If OPTTaxMRP.Value = False And OPTVAT.Value = False Then
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "EXPIRY RETURN"
'                    Exit Sub
'                End If
'            End If
            If TxttaxMRP.Enabled = True Then
                TxttaxMRP.Enabled = False
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

Private Sub OPTVAT_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
        Case vbKeyReturn
'            If Val(TxttaxMRP.Text) <> 0 Then
'                If OPTTaxMRP.Value = False And OPTVAT.Value = False Then
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "EXPIRY RETURN"
'                    Exit Sub
'                End If
'            End If
            If TxttaxMRP.Enabled = True Then
                TxttaxMRP.Enabled = False
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

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(txtBatch.Text) = "" Then Exit Sub
            txtBatch.Enabled = False
            TXTRATE.Enabled = True
            TXTEXPIRY.Visible = True
            TXTEXPIRY.Enabled = True
            TXTEXPIRY.SetFocus
        Case vbKeyEscape
            TXTQTY.Enabled = True
            txtBatch.Enabled = False
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTDIST As ADODB.Recordset
    Dim RSTTRNSMAST As ADODB.Recordset
    Dim i As Long

    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            grdsales.rows = 1
            i = 0
            LBLTOTAL.Caption = ""
            lbltotalwodiscount = ""
            grdsales.rows = 1
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until rstTRXMAST.EOF
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                i = i + 1

                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = rstTRXMAST!ITEM_CODE
                grdsales.TextMatrix(i, 2) = rstTRXMAST!ITEM_NAME
                grdsales.TextMatrix(i, 3) = Val(rstTRXMAST!QTY) / Val(rstTRXMAST!LINE_DISC)
                grdsales.TextMatrix(i, 4) = rstTRXMAST!UNIT
                grdsales.TextMatrix(i, 5) = rstTRXMAST!LINE_DISC
                grdsales.TextMatrix(i, 6) = Format(rstTRXMAST!MRP, ".000")
                grdsales.TextMatrix(i, 7) = Format(rstTRXMAST!SALES_PRICE, ".000")
                grdsales.TextMatrix(i, 8) = Format(rstTRXMAST!ITEM_COST, ".000")
                grdsales.TextMatrix(i, 9) = Format(rstTRXMAST!PTR, ".000")
                grdsales.TextMatrix(i, 10) = IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
                grdsales.TextMatrix(i, 11) = IIf(IsNull(rstTRXMAST!REF_NO), "", rstTRXMAST!REF_NO)
                grdsales.TextMatrix(i, 12) = IIf(IsNull(rstTRXMAST!EXP_DATE), "", Format(rstTRXMAST!EXP_DATE, "DD/MM/YYYY"))
                grdsales.TextMatrix(i, 13) = Format(rstTRXMAST!TRX_TOTAL, ".000")
                grdsales.TextMatrix(i, 14) = IIf(IsNull(rstTRXMAST!SCHEME), "", rstTRXMAST!SCHEME)
                grdsales.TextMatrix(i, 15) = IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
                grdsales.TextMatrix(i, 16) = rstTRXMAST!LINE_NO
                grdsales.TextMatrix(i, 17) = IIf(IsNull(rstTRXMAST!P_DISC), 0, rstTRXMAST!P_DISC)
                grdsales.TextMatrix(i, 18) = IIf(IsNull(rstTRXMAST!P_RETAIL), 0, rstTRXMAST!P_RETAIL)
                grdsales.TextMatrix(i, 19) = IIf(IsNull(rstTRXMAST!P_WS), 0, rstTRXMAST!P_WS)
                grdsales.TextMatrix(i, 20) = IIf(IsNull(rstTRXMAST!P_CRTN), 0, rstTRXMAST!P_CRTN)
                If rstTRXMAST!COM_FLAG = "A" Then
                    grdsales.TextMatrix(i, 21) = ""
                    grdsales.TextMatrix(i, 22) = IIf(IsNull(rstTRXMAST!COM_AMT), 0, rstTRXMAST!COM_AMT)
                    grdsales.TextMatrix(i, 23) = "A"
                Else
                    grdsales.TextMatrix(i, 21) = IIf(IsNull(rstTRXMAST!COM_PER), 0, rstTRXMAST!COM_PER)
                    grdsales.TextMatrix(i, 22) = ""
                    grdsales.TextMatrix(i, 23) = "P"
                End If
                grdsales.TextMatrix(i, 24) = IIf(IsNull(rstTRXMAST!CRTN_PACK), 0, rstTRXMAST!CRTN_PACK)
                grdsales.TextMatrix(i, 25) = IIf(IsNull(rstTRXMAST!P_VAN), 0, rstTRXMAST!P_VAN)
                grdsales.TextMatrix(i, 26) = IIf(IsNull(rstTRXMAST!GROSS_AMT), 0, Format(rstTRXMAST!GROSS_AMT, "0.00"))
                If rstTRXMAST!DISC_FLAG = "P" Then
                    grdsales.TextMatrix(i, 27) = "P"
                Else
                    grdsales.TextMatrix(i, 27) = "A"
                End If
                grdsales.TextMatrix(i, 28) = IIf(IsNull(rstTRXMAST!LOOSE_PACK), 1, rstTRXMAST!LOOSE_PACK)
                grdsales.TextMatrix(i, 29) = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)

                grdsales.TextMatrix(i, 30) = IIf(IsNull(rstTRXMAST!WARRANTY), "", rstTRXMAST!WARRANTY)
                grdsales.TextMatrix(i, 31) = IIf(IsNull(rstTRXMAST!WARRANTY_TYPE), "", rstTRXMAST!WARRANTY_TYPE)

                lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
                TXTDEALER.text = IIf(IsNull(rstTRXMAST!VCH_DESC), "", Mid(rstTRXMAST!VCH_DESC, 15))
                TXTINVDATE.text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                'TXTREMARKS.Text = IIf(IsNull(rstTRXMAST!VCH_DESC), "", Mid(rstTRXMAST!VCH_DESC, 15))
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing

            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
            If (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                rstTRXMAST.AddNew
                rstTRXMAST!VCH_NO = Val(txtBillNo.text)
                rstTRXMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                rstTRXMAST!TRX_TYPE = "ER"
                rstTRXMAST.Update
            Else
                TXTREMARKS.text = IIf(IsNull(rstTRXMAST!REMARKS), "", rstTRXMAST!REMARKS)
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing

            ''''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
            'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
            LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")

            TXTSLNO.text = grdsales.rows
            TXTSLNO.Enabled = True
            txtBillNo.Enabled = False
            FRMEMASTER.Enabled = True
            If i > 0 Or (Val(txtBillNo.text) < Val(TXTLASTBILL.text)) Then
                FRMEMASTER.Enabled = True
                FRMECONTROLS.Enabled = True
                cmdRefresh.Enabled = True
                CMDPRINT.Enabled = True
                CMDPRINT.SetFocus
            Else
                TXTDEALER.SetFocus
            End If

            Set RSTTRNSMAST = New ADODB.Recordset
            RSTTRNSMAST.Open "Select CHECK_FLAG From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRNSMAST.EOF Or RSTTRNSMAST.BOF) Then
                If RSTTRNSMAST!CHECK_FLAG = "Y" Then FRMEMASTER.Enabled = False
            End If
            RSTTRNSMAST.Close
            Set RSTTRNSMAST = Nothing
        
    End Select
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
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
    If Val(txtBillNo.text) = 0 Or Val(txtBillNo.text) > Val(TXTLASTBILL.text) Then txtBillNo.text = TXTLASTBILL.text
End Sub

Private Sub txtcategory_Change()
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHand
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select * From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.text & "%') AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         Else
             PHY.Close
             'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHY.Open "Select * From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.text & "%') AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         End If
         
        Set grdtmp.DataSource = PHY
        
        If PHY.RecordCount > 0 Then
            FRMEGRDTMP.Visible = True
        Else
            FRMEGRDTMP.Visible = False
            Exit Sub
        End If
        grdtmp.Columns(0).Visible = True
        grdtmp.Columns(0).Caption = "CODE"
        grdtmp.Columns(0).Width = 1900
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(1).Width = 5800
        'grdtmp.Columns(2).Visible = False
        grdtmp.Columns(17).Caption = "QTY"
        grdtmp.Columns(17).Width = 1200
        grdtmp.Columns(2).Visible = False
        grdtmp.Columns(3).Visible = False
        grdtmp.Columns(4).Visible = False
        grdtmp.Columns(5).Visible = False
        'grdtmp.Columns(6).Visible = False
        grdtmp.Columns(6).Width = 1200
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
        grdtmp.Columns(47).Visible = False
        Exit Sub
ErrHand:
        MsgBox err.Description
End Sub

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.text)
    Call txtcategory_Change
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
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
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

Private Sub TXTEXPDATE_GotFocus()
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
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
    TXTEXPDATE.text = Format(TXTEXPDATE.text, "DD/MM/YYYY")
    If TXTEXPDATE.text <> "  /  /    " Then TXTEXPIRY.text = Format(TXTEXPDATE.text, "MM/YY")
End Sub

Private Sub TxtFree_GotFocus()
    TXTFREE.SelStart = 0
    TXTFREE.SelLength = Len(TXTFREE.text)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTFREE.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
            TXTQTY.Enabled = True
            TXTFREE.Enabled = False
            TXTQTY.SetFocus
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
    If Val(TXTFREE.text) = 0 Then TXTFREE.text = 0
    TXTFREE.text = Format(TXTFREE.text, "0.00")
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.text = "  /  /    " Then
                TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.text) Then
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTINVDATE.SetFocus
                Exit Sub
            End If
            
            If (DateValue(TXTINVDATE.text) < DateValue(MDIMAIN.DTFROM.Value)) Or (DateValue(TXTINVDATE.text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.Value))) Then
                'db.Execute "delete from Users"
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTINVDATE.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.text) Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
            End If
        Case vbKeyEscape
            TXTREMARKS.SetFocus
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

Private Sub Txtpack_GotFocus()
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.text)
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtpack.text) = 0 Then Exit Sub
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            Txtpack.Enabled = False
            CmbPack.Enabled = True
            CmbPack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            Txtpack.Enabled = False
            txtcategory.Enabled = True
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
    Dim i As Long
    On Error GoTo ErrHand
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select * From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         Else
             PHY.Close
             'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHY.Open "Select * From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         End If
         
        Set grdtmp.DataSource = PHY
        
        If PHY.RecordCount > 0 Then
            FRMEGRDTMP.Visible = True
        Else
            FRMEGRDTMP.Visible = False
            Exit Sub
        End If
        grdtmp.Columns(0).Visible = True
        grdtmp.Columns(0).Caption = "CODE"
        grdtmp.Columns(0).Width = 1900
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(1).Width = 5800
        'grdtmp.Columns(2).Visible = False
        grdtmp.Columns(17).Caption = "QTY"
        grdtmp.Columns(17).Width = 1200
        grdtmp.Columns(2).Visible = False
        grdtmp.Columns(3).Visible = False
        grdtmp.Columns(4).Visible = False
        grdtmp.Columns(5).Visible = False
        'grdtmp.Columns(6).Visible = False
        grdtmp.Columns(6).Width = 1200
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
        grdtmp.Columns(47).Visible = False
        Exit Sub
ErrHand:
        MsgBox err.Description
                
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
        Case vbKeyReturn

            If Trim(TXTPRODUCT.text) = "" Then Exit Sub
'            If Trim(TXTPRODUCT.Text) = "" Then
'                TXTITEMCODE.Enabled = True
'                TXTITEMCODE.SetFocus
'                Exit Sub
'            End If
            CmdDelete.Enabled = False

            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If

            Set grdtmp.DataSource = PHY

            If PHY.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "EXPIRY RETURN"
                Exit Sub
            End If

            If PHY.RecordCount = 1 Then
                TXTITEMCODE.text = grdtmp.Columns(0)
                TXTPRODUCT.text = grdtmp.Columns(1)
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From  rtntrxfile  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' ORDER BY CREATE_DATE", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    RSTRXFILE.MoveLast
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.text = 1
                    Los_Pack.text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    TxtWarranty.text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo ErrHand
                Else
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = 1
                    Los_Pack.text = 1
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand

                    TXTEXPDATE.text = "  /  /    "
                    txtBatch.text = ""
                    TXTEXPIRY.text = "  /  "
                    TXTRATE.text = ""
                    txtmrpbt.text = ""
                    TXTPTR.text = ""
                    txtretail.text = ""
                    txtWS.text = ""
                    txtvanrate.text = ""
                    txtcrtn.text = ""
                    txtcrtnpack.text = ""
                    txtprofit.text = ""
                    TxttaxMRP.text = "5"
                    Los_Pack.text = "1"
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    OPTVAT.Value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing

                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    txtcategory.Enabled = False
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
            TXTITEMCODE.Enabled = False
            'TXTPRODUCT.Enabled = False
            txtcategory.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            txtcategory.SetFocus
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

Private Sub TXTPTR_GotFocus()
    TXTPTR.SelStart = 0
    TXTPTR.SelLength = Len(TXTPTR.text)
End Sub

Private Sub TXTPTR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTPTR.text) = 0 Then Exit Sub
            TxttaxMRP.Enabled = True
            TXTPTR.Enabled = False
            TxttaxMRP.SetFocus
        Case vbKeyEscape
            TXTPTR.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case 116
            Call FILL_PREVIIOUSRATE
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
    'tXTptrdummy.Text = Format(Val(TXTPTR.Text) / Val(TXTUNIT.Text), ".000")
    Txtgrossamt.text = Val(TXTPTR.text) * Val(TXTQTY.text)
    TXTPTR.text = Format(TXTPTR.text, ".000")
    'TXTRETAIL.Text = Round(Val(txtmrpbt.Text) * 0.8, 2)
'    txtretail.Text = Format(Round(Val(TXTRATE.Text) - (Val(txtmrpbt.Text) * 20 / 100), 3), ".000")
'    txtprofit.Text = Format(Round(Val(txtretail.Text) - Val(txtretail.Text) * 10 / 100, 3), ".000")
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.text)
    'Call FILL_PREVIIOUSRATE
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.text) = 0 Then Exit Sub
            TXTQTY.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
'            TXTQTY.Text = ""
'            TXTFREE.Text = ""
'            TxttaxMRP.Text = ""
'            txtprofit.Text = ""
'            txtretail.Text = ""
'            txtWS.Text = ""
'            txtvanrate.Text = ""
'            Txtgrossamt.Text = ""
'            txtcrtn.Text = ""
'            txtcrtnpack.Text = ""
'            txtPD.Text = ""
'            txtBatch.Text = ""
'            TXTRATE.Text = ""
'            txtmrpbt.Text = ""
'            TXTPTR.Text = ""
'            Txtgrossamt.Text = ""
'            LBLSUBTOTAL.Caption = ""
'            lbltaxamount.Caption = ""
            TXTQTY.Enabled = False
            CmbPack.Enabled = True
            CmbPack.SetFocus
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
    TXTQTY.text = Format(TXTQTY.text, ".00")
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Round(Val(TXTPTR.text), 2)), ".000")
End Sub

Private Sub TXTRATE_GotFocus()
    TXTRATE.SelStart = 0
    TXTRATE.SelLength = Len(TXTRATE.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TXTRATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTRATE.Text) = 0 Then Exit Sub
            TXTRATE.Enabled = False
            TXTPTR.Enabled = True
            TXTPTR.SetFocus
         Case vbKeyEscape
            TXTRATE.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
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
    TXTRATE.text = Format(TXTRATE.text, ".000")
    txtmrpbt.text = 100 * Val(TXTRATE.text) / 105 '(100 + Val(TxttaxMRP.Text))
End Sub

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If txtBillNo.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then Exit Sub
            'If TXTINVOICE.Text = "" Then Exit Sub
            FRMECONTROLS.Enabled = True
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTINVDATE.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
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

Private Sub TxtRetailPercent_GotFocus()
    TxtRetailPercent.SelStart = 0
    TxtRetailPercent.SelLength = Len(TxtRetailPercent.text)
End Sub

Private Sub TxtRetailPercent_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            txtretail.Enabled = False
            TxtRetailPercent.Enabled = False
            txtWS.Enabled = True
            txtWsalePercent.Enabled = True
            txtWS.SetFocus
            'TXTRETAIL.SetFocus
         Case vbKeyEscape
            txtretail.SetFocus
    End Select
End Sub

Private Sub TxtRetailPercent_LostFocus()
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - (Val(TXTPTR.text) * Val(txtPD.text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - Val(txtPD.text)
    End If
    txtretail.text = Round((Val(TXTPTR.Tag) * Val(TxtRetailPercent.text) / 100) + Val(TXTPTR.Tag), 2)
    'txtretail.Text = Round(((Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) + Val(TXTPTR.Text)) * Val(TxtRetailPercent.Text) / 100 + ((Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) + Val(TXTPTR.Text)), 2)
    txtretail.text = Format(Val(txtretail.text), "0.0000")
End Sub

Private Sub txtSchPercent_GotFocus()
    txtSchPercent.SelStart = 0
    txtSchPercent.SelLength = Len(txtSchPercent.text)
End Sub

Private Sub txtSchPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtvanrate.Enabled = False
            txtSchPercent.Enabled = False
            'Frame1.Enabled = True
         Case vbKeyEscape
            txtvanrate.SetFocus
    End Select
End Sub

Private Sub txtSchPercent_LostFocus()
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - (Val(TXTPTR.text) * Val(txtPD.text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - Val(txtPD.text)
    End If
    txtvanrate.text = Round((Val(TXTPTR.Tag) * Val(txtSchPercent.text) / 100) + Val(TXTPTR.Tag), 2)
    txtvanrate.text = Format(Val(txtvanrate.text), "0.000")
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.text)
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.text) = 0 Then
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
                TXTQTY.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 14))
                TXTUNIT.text = 1 'grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                Txtpack.text = 1 'grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTRATE.text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5)), 2), "0.000")
                TXTPTR.text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 9)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5)), 2), "0.000")
                txtprofit.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 7)), "0.00")
                txtretail.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18)), "0.00")
                txtWS.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 19)), "0.00")
                txtvanrate.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 25)), "0.00")
                Txtgrossamt.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 26)), "0.00")
                txtcrtn.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20)), "0.00")
                txtcrtnpack.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24)), "0.00")
                'TXTPTR.Text = Format((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))) * Val(Los_Pack.Text), "0.000")

                txtBatch.text = grdsales.TextMatrix(Val(TXTSLNO.text), 11)
                TXTEXPDATE.text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 12)), grdsales.TextMatrix(Val(TXTSLNO.text), 12), "  /  /    ")
                TXTEXPIRY.text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 12)), Format(grdsales.TextMatrix(Val(TXTSLNO.text), 12), "mm/yy"), "  /  ")
                'LBLSUBTOTAL.Caption = Format(Val(TXTQTY.Text) * (Val(TXTPTR.Text) + Val(lbltaxamount.Caption)), ".000")
                If OptDiscAmt.Value = True Then
                    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text) + Val(lbltaxamount.Caption) - Val(txtPD.text), ".000")
                Else
                    LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100), ".000")
                End If
                TXTFREE.text = grdsales.TextMatrix(Val(TXTSLNO.text), 14)
                TxttaxMRP.text = grdsales.TextMatrix(Val(TXTSLNO.text), 10)
                txtmrpbt.text = 100 * Val(TXTRATE.text) / 105 '(100 + Val(TxttaxMRP.Text))
                txtPD.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 17))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) = "V" Then
                    OPTVAT.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) = "M" Then
                    OPTTaxMRP.Value = True
                Else
                    optnet.Value = True
                End If

                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 27)) = "P" Then
                    optdiscper.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 27)) = "A" Then
                    OptDiscAmt.Value = True
                End If
                On Error Resume Next
                Los_Pack.text = grdsales.TextMatrix(Val(TXTSLNO.text), 28)
                CmbPack.text = grdsales.TextMatrix(Val(TXTSLNO.text), 29)
                TxtWarranty.text = grdsales.TextMatrix(Val(TXTSLNO.text), 30)
                CmbWrnty.text = grdsales.TextMatrix(Val(TXTSLNO.text), 31)

                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                txtcategory.Enabled = False
                TXTITEMCODE.Enabled = False
                TXTQTY.Enabled = False
                TXTRATE.Enabled = False
                TXTEXPDATE.Enabled = False
                txtBatch.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            TXTPRODUCT.Enabled = True
            txtcategory.Enabled = True
            txtcategory.SetFocus
            'TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.text = Val(grdsales.rows)
                TXTPRODUCT.text = ""
                TXTITEMCODE.text = ""
                TXTQTY.text = ""
                Txtpack.text = 1 '""
                Los_Pack.text = ""
                CmbPack.ListIndex = -1
                TxtWarranty.text = ""
                CmbWrnty.ListIndex = -1
                TXTFREE.text = ""
                TxttaxMRP.text = ""
                txtPD.text = ""
                txtprofit.text = ""
                txtretail.text = ""
                TxtRetailPercent.text = ""
                txtWsalePercent.text = ""
                txtSchPercent.text = ""
                txtWS.text = ""
                txtvanrate.text = ""
                Txtgrossamt.text = ""
                txtcrtn.text = ""
                txtcrtnpack.text = ""
                TXTRATE.text = ""
                TxtComAmt.text = ""
                TxtComper.text = ""
                txtmrpbt.text = ""
                LBLSUBTOTAL.Caption = ""
                lbltaxamount.Caption = ""
                TXTEXPDATE.text = "  /  /    "
                TXTEXPIRY.text = "  /  "
                txtBatch.text = ""
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            Else
                cmdRefresh.Enabled = True
                CMDPRINT.Enabled = True
                CMDPRINT.SetFocus
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
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
Dim M_DATE As Date
Dim D As Integer
Dim M As Integer
Dim Y As Integer
    Select Case KeyCode
        Case vbKeyReturn
            'If Len(Trim(TXTEXPIRY.Text)) = 1 Then GoTo SKIP
            If Val(Mid(TXTEXPIRY.text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.text, 4, 5)) = 0 Then Exit Sub

            If Val(Mid(TXTEXPIRY.text, 1, 2)) = 0 Then
                TXTEXPDATE.text = "  /  /    "
                Exit Sub
            End If
            If Val(Mid(TXTEXPIRY.text, 4, 5)) = 0 Then
                TXTEXPDATE.text = "  /  /    "
                Exit Sub
            End If

            If Val(Mid(TXTEXPIRY.text, 1, 2)) > 12 Then
                TXTEXPDATE.text = "  /  /    "
                Exit Sub
            End If

            M = Val(Mid(TXTEXPIRY.text, 1, 2))
            Y = Val(Right(TXTEXPIRY.text, 2))
            Y = 2000 + Y
            M_DATE = "01" & "/" & M & "/" & Y
            D = LastDayOfMonth(M_DATE)
            M_DATE = D & "/" & M & "/" & Y
            TXTEXPDATE.text = Format(M_DATE, "dd/mm/yyyy")
            
'            If lbltype.Caption <> "ER" Then
'                If DateDiff("d", Date, TXTEXPDATE.Text) < 0 Then
'                    MsgBox "Item Expired....", vbOKOnly, "Sales Return.."
'                    TXTEXPDATE.Text = "  /  /    "
'                    TXTEXPIRY.SelStart = 0
'                    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
'                    TXTEXPIRY.SetFocus
'                    Exit Sub
'                End If
'
'                If DateDiff("d", Date, TXTEXPDATE.Text) < 60 Then
'                    MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days", vbOKOnly, "Sales Return.."
'                    TXTEXPDATE.Text = "  /  /    "
'                    TXTEXPIRY.SelStart = 0
'                    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
'                    TXTEXPIRY.SetFocus
'                    Exit Sub
'                End If
'
'                If DateDiff("d", Date, TXTEXPDATE.Text) < 180 Then
'                    If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "Sales Return..") = vbNo Then
'                        TXTEXPDATE.Text = "  /  /    "
'                        TXTEXPIRY.SelStart = 0
'                        TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
'                        TXTEXPIRY.SetFocus
'                        Exit Sub
'                    End If
'                End If
'            End If
SKIP:
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
        Case vbKeyEscape
            TXTRATE.Enabled = False
            TXTEXPIRY.Visible = False
            txtBatch.Enabled = True
            TXTEXPDATE.Enabled = False
            txtBatch.SetFocus

    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(txtBatch.text)
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

Private Sub TxttaxMRP_GotFocus()
    TxttaxMRP.SelStart = 0
    TxttaxMRP.SelLength = Len(TxttaxMRP.text)
End Sub

Private Sub TxttaxMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxttaxMRP.text) <> 0 And optnet.Value = True Then
                OPTVAT.Value = True
                OPTVAT.SetFocus
                Exit Sub
            End If
            TxttaxMRP.Enabled = False
            txtPD.Enabled = True
            txtPD.SetFocus
         Case vbKeyEscape
            TxttaxMRP.Enabled = False
            TXTPTR.Enabled = True
            TXTPTR.SetFocus
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
    txtmrpbt.text = 100 * Val(TXTRATE.text) / (100 + Val(TxttaxMRP.text))
    If Val(TxttaxMRP.text) = 0 Then
        TxttaxMRP.text = 0
        lbltaxamount.Caption = 0
        lbltaxamount.Caption = ""
        If optdiscper.Value = True Then
            LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.text)) - Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100), ".000")
        Else
            LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text) - Val(txtPD.text), ".000")
        End If

    Else
        If OPTTaxMRP.Value = True Then
            If optdiscper.Value = True Then
                lbltaxamount.Caption = (Val(txtmrpbt.text) - (Val(TXTRATE.text) * Val(txtPD.text) / 100)) * (Val(TXTQTY.text)) * Val(TxttaxMRP.text) / 100
                LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption), ".000")
            Else
                lbltaxamount.Caption = (Val(txtmrpbt.text) - Val(txtPD.text)) * (Val(TXTQTY.text)) * Val(TxttaxMRP.text) / 100
                LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption), ".000")
            End If
        ElseIf OPTVAT.Value = True Then
            'lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text) + Val(TxtFree.Text))
            'lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text))

            If optdiscper.Value = True Then
                lbltaxamount.Caption = Round((Val(Txtgrossamt.text) - (Val(Txtgrossamt.text) * Val(txtPD.text) / 100)) * Val(TxttaxMRP.text) / 100, 2)
                LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100), ".000")
            Else
                lbltaxamount.Caption = Round((Val(Txtgrossamt.text) - Val(txtPD.text)) * Val(TxttaxMRP.text) / 100, 2)
                LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text) + Val(lbltaxamount.Caption) - Val(txtPD.text), ".000")
            End If
            'LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.Text)) + Val(lbltaxamount.Caption), ".000")
        Else
            lbltaxamount.Caption = ""
            LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text), ".000")
        End If
    End If

    TxttaxMRP.text = Format(TxttaxMRP.text, "0.00")
    lbltaxamount.Caption = Format(lbltaxamount.Caption, "0.00")
End Sub

Private Sub TXTUNIT_GotFocus()
    TXTUNIT.SelStart = 0
    TXTUNIT.SelLength = Len(TXTUNIT.text)
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.text) = 0 Then Exit Sub

            TXTUNIT.Enabled = False
            Txtpack.Enabled = True
            Txtpack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTQTY.text = ""
            TXTFREE.text = ""
            TxttaxMRP.text = ""
            txtprofit.text = ""
            txtretail.text = ""
            TxtRetailPercent.text = ""
            txtWsalePercent.text = ""
            txtSchPercent.text = ""
            txtWS.text = ""
            txtvanrate.text = ""
            Txtgrossamt.text = ""
            txtcrtn.text = ""
            txtcrtnpack.text = ""
            txtPD.text = ""
            txtBatch.text = ""
            TXTRATE.text = ""
            txtmrpbt.text = ""
            TXTPTR.text = ""
            Txtgrossamt.text = ""
            TXTEXPDATE.text = "  /  /    "
            TXTEXPIRY.text = "  /  "
            LBLSUBTOTAL.Caption = ""
            lbltaxamount.Caption = ""
            TXTPRODUCT.Enabled = True
            txtcategory.Enabled = True
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
    Dim DISC As Currency

    On Error GoTo ErrHand
    If (TXTDISCAMOUNT.text = "") Then
        DISC = 0
    Else
        DISC = TXTDISCAMOUNT.text
    End If
    If grdsales.rows = 1 Then
        TXTDISCAMOUNT.text = "0"
    ElseIf Val(TXTDISCAMOUNT.text) > Val(lbltotalwodiscount.Caption) Then
        MsgBox "Discount Amount More than Bill Amount", , "SALES..."
        TXTDISCAMOUNT.SelStart = 0
        TXTDISCAMOUNT.SelLength = Len(TXTDISCAMOUNT.text)
        TXTDISCAMOUNT.SetFocus
        Exit Sub
    End If
    TXTDISCAMOUNT.text = Format(TXTDISCAMOUNT.text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    ''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    TXTDISCAMOUNT.SetFocus
End Sub

Private Sub TXTDISCAMOUNT_GotFocus()
    TXTDISCAMOUNT.SelStart = 0
    TXTDISCAMOUNT.SelLength = Len(TXTDISCAMOUNT.text)
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
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub appendpurchase()

    Dim rstMaxRec As ADODB.Recordset
    Dim RSTLINK As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    Dim TRXTYPE, INVTRXTYPE, INVTYPE As String
    Dim RECNO, INVNO As Long
    
    Dim M_DATA As Double
    Dim i As Long

    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass

    db.Execute "delete From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE= 'DB' AND INV_TRX_TYPE = 'DN'"
    
    If grdsales.rows = 1 Then
        db.Execute "delete FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE= 'ER' AND INV_TRX_TYPE = 'ER'"
        GoTo SKIP
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!TRX_TYPE = "ER"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = Trim(DataList2.text)
        RSTTRXFILE!VCH_AMOUNT = Val(lbltotalwodiscount.Caption)
        RSTTRXFILE!DISCOUNT = Val(TXTDISCAMOUNT.text)
        RSTTRXFILE!ADD_AMOUNT = Val(txtaddlamt.text)
        RSTTRXFILE!ROUNDED_OFF = 0
        RSTTRXFILE!OPEN_PAY = 0
        RSTTRXFILE!PAY_AMOUNT = 0
        RSTTRXFILE!REF_NO = ""
        RSTTRXFILE!SLSM_CODE = "CS"
        RSTTRXFILE!CHECK_FLAG = "N"
        RSTTRXFILE!POST_FLAG = "N"
        RSTTRXFILE!CFORM_NO = ""
        RSTTRXFILE!CFORM_DATE = Date
        RSTTRXFILE!REMARKS = Trim(TXTREMARKS.text)
        RSTTRXFILE!DISC_PERS = Val(txtcramt.text)
        RSTTRXFILE!CST_PER = Val(TxtCST.text)
        RSTTRXFILE!INS_PER = Val(TxtInsurance.text)
        RSTTRXFILE!LETTER_NO = 0
        RSTTRXFILE!LETTER_DATE = Date
        RSTTRXFILE!INV_MSGS = ""
        RSTTRXFILE!CREATE_DATE = Format(TXTDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!MODIFY_DATE = Format(TXTDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    i = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(CR_NO) From DBTPYMT", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing

    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'ER' AND INV_TRX_TYPE = 'ER'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!TRX_TYPE = "ER"
        RSTITEMMAST!INV_TRX_TYPE = "ER"
        RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTITEMMAST!CR_NO = i
        RSTITEMMAST!INV_NO = Val(txtBillNo.text)
        RSTITEMMAST!INV_AMT = 0
    End If
    RSTITEMMAST!INV_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTITEMMAST!RCPT_AMT = Val(LBLTOTAL.Caption)
    RSTITEMMAST!CHECK_FLAG = "N"
    RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption) - RSTITEMMAST!RCPT_AMT
    RSTITEMMAST!ACT_CODE = DataList2.BoundText
    RSTITEMMAST!ACT_NAME = DataList2.text
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
        
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(Trim(TXTINVDATE.text), "dd/mm/yyyy")
        RSTTRXFILE!VCH_DESC = "Received From " & Mid(DataList2.text, 1, 80)
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

SKIP:
    Set rstMaxNo = New ADODB.Recordset
    'rstMaxNo.Open "Select MAX(VCH_NO) From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= 'ER'", db, adOpenStatic, adLockReadOnly
    rstMaxNo.Open "Select MAX(VCH_NO) From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' ", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing

    grdsales.rows = 1
    TXTSLNO.text = 1
    cmdRefresh.Enabled = False
    CMDPRINT.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.text = TXTLASTBILL.text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.text = "  /  /    "
    TXTREMARKS.text = ""
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TXTPRODUCT.text = ""
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    txtPD.text = ""
    txtprofit.text = ""
    txtretail.text = ""
    TxtRetailPercent.text = ""
    txtWsalePercent.text = ""
    txtSchPercent.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    Txtgrossamt.text = ""
    txtcrtn.text = ""
    txtcrtnpack.text = ""
    txtBatch.text = ""
    TXTRATE.text = ""
    txtmrpbt.text = ""
    TXTPTR.text = ""
    Txtgrossamt.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    txtaddlamt.text = ""
    txtcramt.text = ""
    TxtInsurance.text = ""
    txtcategory.text = ""
    TxtCST.text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    TXTDISCAMOUNT.text = ""
    flagchange.Caption = ""
    TXTDEALER.text = ""
    lbldealer.Caption = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
    txtBillNo.Visible = True
    Chkcancel.Value = 0
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "Sales Return ENTRY"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If err.Number = 7 Then
        MsgBox "Select Customer from the list", vbOKOnly, "EXPIRY RETURN"
    Else
        MsgBox err.Description
    End If
End Sub


Private Sub txtaddlamt_GotFocus()
    txtaddlamt.SelStart = 0
    txtaddlamt.SelLength = Len(txtaddlamt.text)
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
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub txtaddlamt_LostFocus()
    Dim DISC As Currency

    On Error GoTo ErrHand
    If (txtaddlamt.text = "") Then
        DISC = 0
    Else
        DISC = txtaddlamt.text
    End If
    If grdsales.rows = 1 Then
        txtaddlamt.text = "0"
    ElseIf Val(txtaddlamt.text) > Val(lbltotalwodiscount.Caption) Then
        MsgBox "Discount Amount More than Bill Amount", , "Sales Return..."
        txtaddlamt.SelStart = 0
        txtaddlamt.SelLength = Len(txtaddlamt.text)
        txtaddlamt.SetFocus
        Exit Sub
    End If
    txtaddlamt.text = Format(txtaddlamt.text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    txtaddlamt.SetFocus
End Sub

Private Sub txtcramt_GotFocus()
    txtcramt.SelStart = 0
    txtcramt.SelLength = Len(txtcramt.text)
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
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub txtcramt_LostFocus()
    Dim DISC As Currency

    On Error GoTo ErrHand
    If (txtcramt.text = "") Then
        DISC = 0
    Else
        DISC = txtcramt.text
    End If
    If grdsales.rows = 1 Then
        txtcramt.text = "0"
    ElseIf Val(txtcramt.text) > Val(lbltotalwodiscount.Caption) Then
        MsgBox "Credit Note Amount More than Bill Amount", , "Sales Return..."
        txtcramt.SelStart = 0
        txtcramt.SelLength = Len(txtcramt.text)
        txtcramt.SetFocus
        Exit Sub
    End If
    txtcramt.text = Format(txtcramt.text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    txtcramt.SetFocus
End Sub

Private Sub OPTTaxMRP_GotFocus()
    If optdiscper.Value = True Then
        lbltaxamount.Caption = (Val(txtmrpbt.text) - (Val(TXTRATE.text) * Val(txtPD.text) / 100)) * (Val(TXTQTY.text)) * Val(TxttaxMRP.text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption), ".000")

        'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
        'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
    Else
        lbltaxamount.Caption = (Val(txtmrpbt.text) - Val(txtPD.text)) * (Val(TXTQTY.text)) * Val(TxttaxMRP.text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption), ".000")
    End If
End Sub

Private Sub OPTVAT_GotFocus()
    'lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text) + Val(TxtFree.Text))
    If optdiscper.Value = True Then
        lbltaxamount.Caption = Round((Val(Txtgrossamt.text) - (Val(Txtgrossamt.text) * Val(txtPD.text) / 100)) * Val(TxttaxMRP.text) / 100, 2)
        LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100), ".000")
    Else
        lbltaxamount.Caption = Round((Val(Txtgrossamt.text) - Val(txtPD.text)) * Val(TxttaxMRP.text) / 100, 2)
        LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text) + Val(lbltaxamount.Caption) - Val(txtPD.text), ".000")
    End If
End Sub

Private Sub OPTNET_GotFocus()
    lbltaxamount.Caption = ""
    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text), ".000")
End Sub

Private Sub txtprofit_GotFocus()
    txtprofit.SelStart = 0
    txtprofit.SelLength = Len(txtprofit.text)
End Sub

Private Sub txtprofit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtprofit.Enabled = False
            txtretail.Enabled = True
            txtretail.SetFocus
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
    txtprofit.text = Format(txtprofit.text, "0.00")
End Sub

Private Sub txtPD_GotFocus()
    txtPD.SelStart = 0
    txtPD.SelLength = Len(txtPD.text)
End Sub

Private Sub txtPD_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtPD.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            txtPD.Enabled = False
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
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
    Call TxttaxMRP_LostFocus

    If Val(txtretail.text) = 0 Then txtretail.text = Format(Round(Val(txtmrpbt.text) - (Val(txtmrpbt.text) * 20 / 100), 3), ".000")
    If Val(txtWS.text) = 0 Then txtWS.text = Format(Round(Val(txtretail.text) - (Val(txtretail.text) * 10 / 100), 3), ".000")

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
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
    MsgBox err.Description
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
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
    TXTDEALER.text = DataList2.text
    lbldealer.Caption = TXTDEALER.text
    On Error GoTo ErrHand
    Dim rstCustomer As ADODB.Recordset
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        'If TxtBillName.Text = "" Then TxtBillName.Text = DataList2.Text
        'TxtBillName.Text = DataList2.Text
        If TxtBillAddress.text = "" Then TxtBillAddress.text = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
        'TxtBillAddress.Text = IIf(IsNull(rstCustomer!ADDRESS), "", Trim(rstCustomer!ADDRESS))
        TXTTIN.text = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
        TxtPhone.text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO))
        'TXTAREA.Text = IIf(IsNull(rstCustomer!Area), "", Trim(rstCustomer!Area))
        'TxtDL1.Text = IIf(IsNull(rstCustomer!DL_NO), "", Trim(rstCustomer!DL_NO))
        'TxtDL2.Text = IIf(IsNull(rstCustomer!REMARKS), "", Trim(rstCustomer!REMARKS))
        'TxtCST.Text = IIf(IsNull(rstCustomer!CST), "", Trim(rstCustomer!CST))
    End If
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Sales Return..."
                DataList2.SetFocus
                Exit Sub
            End If
            TXTINVDATE.SetFocus
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
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub TXTRETAIL_GotFocus()
    txtretail.SelStart = 0
    txtretail.SelLength = Len(txtretail.text)
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtretail.text) = 0 Then
                TxtRetailPercent.SetFocus
            Else
                txtretail.Enabled = False
                TxtRetailPercent.Enabled = False
                txtWS.Enabled = True
                txtWsalePercent.Enabled = True
                txtWS.SetFocus
            End If
            Exit Sub
            If Val(txtretail.text) = 0 Then
                TxtRetailPercent.SetFocus
                Exit Sub
            End If
            txtretail.Enabled = False
            TxtRetailPercent.Enabled = False
            'cmdadd.Enabled = True
            'cmdadd.SetFocus
            txtWS.Enabled = True
            txtWsalePercent.Enabled = True
            txtWS.SetFocus
         Case vbKeyEscape
            txtretail.Enabled = False
            TxtRetailPercent.Enabled = False
            txtPD.Enabled = True
            txtPD.SetFocus
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
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - (Val(TXTPTR.text) * Val(txtPD.text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - Val(txtPD.text)
    End If
    TxtRetailPercent.text = Round(((Val(txtretail.text) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
    TxtRetailPercent.text = Format(Val(TxtRetailPercent.text), "0.00")
End Sub

Private Sub txtws_GotFocus()
    txtWS.SelStart = 0
    txtWS.SelLength = Len(txtWS.text)
End Sub

Private Sub txtws_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtWS.text) = 0 Then
                txtWsalePercent.SetFocus
            Else
                txtWS.Enabled = False
                txtWsalePercent.Enabled = False
                cmdadd.Enabled = True
                cmdadd.SetFocus
            End If
            Exit Sub
            If Val(txtWS.text) = 0 Then
                txtWsalePercent.SetFocus
                Exit Sub
            End If
            txtWS.Enabled = False
            txtWsalePercent.Enabled = False
         Case vbKeyEscape
            txtWS.Enabled = False
            txtWsalePercent.Enabled = False
            txtretail.Enabled = True
            TxtRetailPercent.Enabled = True
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
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - (Val(TXTPTR.text) * Val(txtPD.text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - Val(txtPD.text)
    End If
    txtWsalePercent.text = Round(((Val(txtWS.text) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
    txtWsalePercent.text = Format(Val(txtWsalePercent.text), "0.00")
End Sub

Private Sub txtcrtn_GotFocus()
    txtcrtn.SelStart = 0
    txtcrtn.SelLength = Len(txtcrtn.text)
End Sub

Private Sub txtcrtn_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcrtn.Enabled = False
            txtcrtnpack.Enabled = True
            txtcrtnpack.SetFocus
         Case vbKeyEscape
            txtcrtn.Enabled = False
            txtWS.Enabled = True
            txtSchPercent.Enabled = True
            txtWS.SetFocus
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
    txtcrtn.text = Format(txtcrtn.text, "0.00")
End Sub

Private Sub txtcrtnpack_GotFocus()
    If Val(txtcrtn.text) = 0 Then txtcrtnpack.text = 0
    txtcrtnpack.SelStart = 0
    txtcrtnpack.SelLength = Len(txtcrtnpack.text)
End Sub

Private Sub txtcrtnpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtcrtn.text) <> 0 And Val(txtcrtnpack.text) = 0 Then
                MsgBox "Please enter the Pack Qty for Carton", vbOKOnly, "EXPIRY RETURN"
                txtcrtnpack.SetFocus
                Exit Sub
            End If
            If Val(txtcrtn.text) = 0 And Val(txtcrtnpack.text) <> 0 Then
                MsgBox "Please enter the Rate for Carton", vbOKOnly, "EXPIRY RETURN"
                txtcrtnpack.Enabled = False
                txtcrtn.Enabled = True
                txtcrtn.SetFocus
                Exit Sub
            End If
            txtcrtnpack.Enabled = False
            
         Case vbKeyEscape
            txtcrtnpack.Enabled = False
            txtcrtn.Enabled = True
            txtcrtn.SetFocus
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
    txtcrtnpack.text = Format(txtcrtnpack.text, "0.00")
End Sub

Private Sub txtvanrate_GotFocus()
    txtvanrate.SelStart = 0
    txtvanrate.SelLength = Len(txtvanrate.text)
End Sub

Private Sub txtvanrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtvanrate.text) = 0 Then
                txtSchPercent.SetFocus
            Else
                txtvanrate.Enabled = False
                txtSchPercent.Enabled = False
                TxtWarranty.Enabled = True
                CmbWrnty.Enabled = True
                TxtWarranty.SetFocus
            End If
            Exit Sub
            If Val(txtvanrate.text) = 0 Then
                txtSchPercent.SetFocus
                Exit Sub
            End If
            txtvanrate.Enabled = False
            txtSchPercent.Enabled = False
            TxtWarranty.Enabled = True
            CmbWrnty.Enabled = True
            TxtWarranty.SetFocus
            'txtcrtn.Enabled = True
            'txtcrtn.SetFocus
         Case vbKeyEscape
            txtvanrate.Enabled = False
            txtSchPercent.Enabled = False
            txtWS.Enabled = True
            txtWsalePercent.Enabled = True
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
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - (Val(TXTPTR.text) * Val(txtPD.text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - Val(txtPD.text)
    End If
    txtSchPercent.text = Round(((Val(txtvanrate.text) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
    txtSchPercent.text = Format(Val(txtSchPercent.text), "0.00")
End Sub

Private Sub Txtgrossamt_GotFocus()
    Txtgrossamt.SelStart = 0
    Txtgrossamt.SelLength = Len(Txtgrossamt.text)
End Sub

Private Sub Txtgrossamt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(Txtgrossamt.Text) = 0 Then Exit Sub
            TXTRATE.Enabled = True
            Txtgrossamt.Enabled = False
            TXTRATE.SetFocus
        Case vbKeyEscape
            Txtgrossamt.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
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
    If Val(Txtgrossamt.text) <> 0 Then
        Txtgrossamt.text = Format(Txtgrossamt.text, ".000")
        TXTPTR.text = Format(Round(Val(Txtgrossamt.text) / Val(TXTQTY.text), 1), "0.00")
    End If
    Call TxttaxMRP_LostFocus
End Sub

Function FILL_PREVIIOUSRATE()
    Set GRDPRERATE.DataSource = Nothing

    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, VCH_NO, LINE_NO  From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, VCH_NO, LINE_NO  From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    End If

    If PHY_PRERATE.RecordCount > 0 Then
        'Fram.Enabled = False
        fRMEPRERATE.Visible = True
        Set GRDPRERATE.DataSource = PHY_PRERATE
        
        GRDPRERATE.Columns(0).Caption = "TYPE"
        GRDPRERATE.Columns(1).Caption = "ITEM CODE"
        GRDPRERATE.Columns(2).Caption = "ITEM NAME"
        GRDPRERATE.Columns(3).Caption = "BILL DATE"
        GRDPRERATE.Columns(4).Caption = "SOLD QTY"
        GRDPRERATE.Columns(5).Caption = "MRP"
        GRDPRERATE.Columns(6).Caption = "SOLD RATE"
        GRDPRERATE.Columns(7).Caption = "TAX"
        GRDPRERATE.Columns(8).Caption = "TAX MODE"
        GRDPRERATE.Columns(9).Caption = "COST"
        GRDPRERATE.Columns(10).Caption = "BATCH"
        GRDPRERATE.Columns(11).Caption = "EXPIRY"
        GRDPRERATE.Columns(12).Caption = "VCH NO"
        GRDPRERATE.Columns(13).Caption = "LINE NO"

        GRDPRERATE.Columns(0).Visible = False
        GRDPRERATE.Columns(1).Visible = False
        GRDPRERATE.Columns(2).Width = 3500
        GRDPRERATE.Columns(3).Width = 1400
        GRDPRERATE.Columns(4).Width = 1200
        GRDPRERATE.Columns(5).Width = 1200
        GRDPRERATE.Columns(6).Width = 1200
        GRDPRERATE.Columns(7).Width = 1200
        GRDPRERATE.Columns(8).Width = 1300
        GRDPRERATE.Columns(9).Width = 1300
        GRDPRERATE.Columns(10).Width = 1300
        GRDPRERATE.Columns(11).Width = 1300
        GRDPRERATE.Columns(12).Width = 500
        GRDPRERATE.Columns(13).Width = 500


        'GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(2).text
        GRDPRERATE.SetFocus
    Else
        Set GRDPRERATE.DataSource = Nothing
    
        If PRERATE_FLAG = True Then
            PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, VCH_NO, LINE_NO  From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND (TRX_TYPE = 'GI' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
            PRERATE_FLAG = False
        Else
            PHY_PRERATE.Close
            PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, VCH_NO, LINE_NO  From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND (TRX_TYPE = 'GI' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
            PRERATE_FLAG = False
        End If
    
        If PHY_PRERATE.RecordCount > 0 Then
            'Fram.Enabled = False
            fRMEPRERATE.Visible = True
            Set GRDPRERATE.DataSource = PHY_PRERATE
            
            GRDPRERATE.Columns(0).Caption = "TYPE"
            GRDPRERATE.Columns(1).Caption = "ITEM CODE"
            GRDPRERATE.Columns(2).Caption = "ITEM NAME"
            GRDPRERATE.Columns(3).Caption = "BILL DATE"
            GRDPRERATE.Columns(4).Caption = "SOLD QTY"
            GRDPRERATE.Columns(5).Caption = "MRP"
            GRDPRERATE.Columns(6).Caption = "SOLD RATE"
            GRDPRERATE.Columns(7).Caption = "TAX"
            GRDPRERATE.Columns(8).Caption = "TAX MODE"
            GRDPRERATE.Columns(9).Caption = "COST"
            GRDPRERATE.Columns(10).Caption = "BATCH"
            GRDPRERATE.Columns(11).Caption = "EXPIRY"
            GRDPRERATE.Columns(12).Caption = "VCH NO"
            GRDPRERATE.Columns(13).Caption = "LINE NO"
    
            GRDPRERATE.Columns(0).Visible = False
            GRDPRERATE.Columns(1).Visible = False
            GRDPRERATE.Columns(2).Width = 3500
            GRDPRERATE.Columns(3).Width = 1400
            GRDPRERATE.Columns(4).Width = 1200
            GRDPRERATE.Columns(5).Width = 1200
            GRDPRERATE.Columns(6).Width = 1200
            GRDPRERATE.Columns(7).Width = 1200
            GRDPRERATE.Columns(8).Width = 1300
            GRDPRERATE.Columns(9).Width = 1300
            GRDPRERATE.Columns(10).Width = 1300
            GRDPRERATE.Columns(11).Width = 1300
            GRDPRERATE.Columns(12).Width = 500
            GRDPRERATE.Columns(13).Width = 500
    
    
            'GRDPRERATE.SetFocus
            LBLHEAD(2).Caption = GRDPRERATE.Columns(2).text
            GRDPRERATE.SetFocus
        Else
            'If MsgBox("This Item has not been sold to " & DataList2.Text & " Yet!! Do You Want to Continue...?", vbYesNo, "SALES RETURN..") = vbYes Then
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
            'Else
            '    TXTQTY.Enabled = False
            '    CmbPack.Enabled = True
            '    CmbPack.SetFocus
            'End If
        End If
    End If

End Function

Private Sub GRDPRERATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTEXPDATE.text = IIf(IsNull(GRDPRERATE.Columns(11)) Or GRDPRERATE.Columns(11) = "", "  /  /    ", Format(GRDPRERATE.Columns(11), "DD/MM/YYYY"))
            txtBatch.text = IIf(IsNull(GRDPRERATE.Columns(10)), "", GRDPRERATE.Columns(10))
            TXTEXPIRY.text = IIf(IsNull(GRDPRERATE.Columns(11)) Or GRDPRERATE.Columns(11) = "", "  /  ", Format(GRDPRERATE.Columns(11), "MM/YY"))
            TXTRATE.text = IIf(IsNull(GRDPRERATE.Columns(5)), "", Format(Round(Val(GRDPRERATE.Columns(5)) * Val(Los_Pack.text), 2), ".000"))
            txtmrpbt.text = 100 * Val(TXTRATE.text) / 105
            TXTPTR.text = IIf(IsNull(GRDPRERATE.Columns(6)), "", Format(Round(Val(GRDPRERATE.Columns(6)) * Val(Los_Pack.text), 2), ".000"))
            'TxttaxMRP.Text = IIf(IsNull(GRDPRERATE.Columns(7)), "", Format(Val(GRDPRERATE.Columns(7)), ".00"))
            optnet.Value = True
'            If GRDPRERATE.Columns(8) = "M" Then
'                OPTTaxMRP.value = True
'            ElseIf GRDPRERATE.Columns(8) = "V" Then
'                OPTVAT.value = True
'            Else
'                optnet.value = True
'            End If
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Fram.Enabled = True
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Fram.Enabled = True
            TXTQTY.Enabled = True
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
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            Los_Pack.Enabled = False
            CmbPack.Enabled = True
            CmbPack.SetFocus
         Case vbKeyEscape
             If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            Los_Pack.Enabled = False
            TXTPRODUCT.Enabled = True
            txtcategory.Enabled = True
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub Los_Pack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtItemcode_GotFocus()
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn

            If Trim(TXTITEMCODE.text) = "" Then
                txtcategory.Enabled = True
                TXTPRODUCT.Enabled = True
                txtcategory.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False

            Set grdtmp.DataSource = Nothing
            If PHYCODE_FLAG = True Then
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            Else
                PHY_CODE.Close
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            End If

            Set grdtmp.DataSource = PHY_CODE

            If PHY_CODE.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "EXPIRY RETURN"
                Exit Sub
            End If

            If PHY_CODE.RecordCount = 1 Then
                TXTITEMCODE.text = grdtmp.Columns(0)
                TXTPRODUCT.text = grdtmp.Columns(1)
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From  rtntrxfile  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' ORDER BY CREATE_DATE", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    RSTRXFILE.MoveLast
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.text = 1
                Else
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = 1
                    Los_Pack.text = 1
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand

                    TXTEXPDATE.text = "  /  /    "
                    txtBatch.text = ""
                    TXTEXPIRY.text = "  /  "
                    TXTRATE.text = ""
                    txtmrpbt.text = ""
                    TXTPTR.text = ""
                    txtretail.text = ""
                    txtWS.text = ""
                    txtvanrate.text = ""
                    txtcrtn.text = ""
                    txtcrtnpack.text = ""
                    txtprofit.text = ""
                    TxttaxMRP.text = "5"
                    Los_Pack.text = "1"
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    OPTVAT.Value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing

                If PHY_CODE.RecordCount = 1 Then
                    TXTITEMCODE.Enabled = False
                    TXTPRODUCT.Enabled = False
                    txtcategory.Enabled = False
                    Los_Pack.Enabled = True
                    Los_Pack.SetFocus
                    'TxtPack.Enabled = True
                    'TxtPack.SetFocus
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
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
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
    TxtCST.SelStart = 0
    TxtCST.SelLength = Len(TxtCST.text)
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
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub TxtCST_LostFocus()
    Dim DISC As Currency

    On Error GoTo ErrHand
    If (TxtCST.text = "") Then
        DISC = 0
    Else
        DISC = TxtCST.text
    End If
    If grdsales.rows = 1 Then
        TxtCST.text = "0"
        Exit Sub
    End If
    TxtCST.text = Format(TxtCST.text, ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtCST.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtCST.SetFocus
End Sub

Private Sub TxtInsurance_GotFocus()
    TxtInsurance.SelStart = 0
    TxtInsurance.SelLength = Len(TxtInsurance.text)
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
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub TxtInsurance_LostFocus()
    Dim DISC As Currency

    On Error GoTo ErrHand
    If (TxtInsurance.text = "") Then
        DISC = 0
    Else
        DISC = TxtInsurance.text
    End If
    If grdsales.rows = 1 Then
        TxtInsurance.text = "0"
        Exit Sub
    End If
    TxtInsurance.text = Format(TxtInsurance.text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + (Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(txtcst.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtInsurance.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtInsurance.SetFocus
End Sub

Private Sub txtWsalePercent_GotFocus()
    txtWsalePercent.SelStart = 0
    txtWsalePercent.SelLength = Len(txtWsalePercent.text)
End Sub

Private Sub txtWsalePercent_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            txtWS.Enabled = False
            txtWsalePercent.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            txtWS.SetFocus
    End Select
End Sub

Private Sub txtWsalePercent_LostFocus()
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - (Val(TXTPTR.text) * Val(txtPD.text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) - Val(txtPD.text)
    End If
    txtWS.text = Round((Val(TXTPTR.Tag) * Val(txtWsalePercent.text) / 100) + Val(TXTPTR.Tag), 2)
    txtWS.text = Format(Val(txtWS.text), "0.000")
End Sub

Private Sub TxtWarranty_GotFocus()
    TxtWarranty.SelStart = 0
    TxtWarranty.SelLength = Len(TxtWarranty.text)
End Sub

Private Sub TxtWarranty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.text) = 0 Then
                TxtWarranty.Enabled = False
                CmbWrnty.Enabled = False
                cmdadd.Enabled = True
                cmdadd.SetFocus
            Else
                CmbWrnty.Enabled = True
                CmbWrnty.SetFocus
            End If
         Case vbKeyEscape
            TxtWarranty.Enabled = False
            CmbWrnty.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
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
            If Val(TxtWarranty.text) <> 0 And CmbWrnty.ListIndex = -1 Then
                MsgBox "Please select the Warranty Period", , "EXPIRY RETURN"
                CmbWrnty.SetFocus
                Exit Sub
            End If
            If Val(TxtWarranty.text) = 0 Then CmbWrnty.ListIndex = -1
            TxtWarranty.Enabled = False
            CmbWrnty.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            TxtWarranty.SetFocus
    End Select
End Sub

Private Function checklastbill()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ErrHand

    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From  rtntrxfile WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'ER' ", db, adOpenStatic, adLockReadOnly
    'rstBILL.Open "Select MAX(VCH_NO) From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= 'ER'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing

Exit Function
ErrHand:
    MsgBox err.Description
End Function

Private Function ReportGeneratION_estimate()
    
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open Rptpath & "Report.txt" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "Report.txt" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ErrHand
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold


    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
        Print #1, AlignLeft("RETURN", 25)
        Print #1, RepeatString("-", 80)
        Print #1, AlignLeft("Sl", 2) & Space(1) & _
                AlignLeft("Comm Code", 14) & Space(1) & _
                AlignLeft("Description", 35) & _
                AlignLeft("Qty", 4) & Space(3) & _
                AlignLeft("Rate", 10) & Space(3) & _
                AlignLeft("Amount", 12) '& _
                Chr(27) & Chr(72)  '//Bold Ends
    
        Print #1, RepeatString("-", 80)
    
        For i = 1 To grdsales.rows - 1
            Print #1, AlignLeft(Val(i), 3) & _
                Space(15) & AlignLeft(grdsales.TextMatrix(i, 2), 34) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 4) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 18)), 2), "0.00"), 9) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 13)), "0.00"), 13) '& _
                Chr(27) & Chr(72)  '//Bold Ends
        Next i
    
        Print #1, AlignRight("-------------", 80)
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 57) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, AlignRight("Round off ", 65) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
        Print #1, Chr(13)
        Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(LBLTOTAL.Caption), "####.00")), 9)
        Num = CCur(Round(LBLTOTAL.Caption, 0))
        Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & ")", 80)
        Print #1, RepeatString("-", 80)
        'Print #1, Chr(27) & Chr(71) & Chr(0)
        If Trim(TXTTIN.text) <> "" Then
            Print #1, "Certified that all the particulars shown in the above Tax Invoice are true and correct"
            Print #1, "and that my/our Registration under KVAT ACT 2003 is valid as on the date of this bill"
            Print #1, RepeatString("-", 80)
        End If
        'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** THANK YOU ****", 40)
    

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

ErrHand:
    MsgBox err.Description
End Function

