VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMGSTBR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BRANCH SALE"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18495
   Icon            =   "FrmGSTBR.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11100
   ScaleWidth      =   18495
   Visible         =   0   'False
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   4365
      TabIndex        =   81
      Top             =   3540
      Visible         =   0   'False
      Width           =   9555
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   2655
         Left            =   15
         TabIndex        =   82
         Top             =   390
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   16
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
         Left            =   3630
         TabIndex        =   84
         Top             =   15
         Width           =   5910
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   " PREVIOUS RATES FOR THE ITEM "
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
         Left            =   15
         TabIndex        =   83
         Top             =   15
         Width           =   3615
      End
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   1875
      TabIndex        =   58
      Top             =   3390
      Visible         =   0   'False
      Width           =   10950
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2835
         Left            =   30
         TabIndex        =   61
         Top             =   360
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   5001
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   23
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
         Height          =   315
         Index           =   9
         Left            =   15
         TabIndex        =   60
         Top             =   30
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
         Height          =   315
         Index           =   0
         Left            =   3060
         TabIndex        =   59
         Top             =   30
         Visible         =   0   'False
         Width           =   7860
      End
   End
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3225
      Left            =   1860
      TabIndex        =   62
      Top             =   3375
      Visible         =   0   'False
      Width           =   10935
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   3165
         Left            =   45
         TabIndex        =   63
         Top             =   45
         Width           =   10860
         _ExtentX        =   19156
         _ExtentY        =   5583
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   23
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
   Begin MSDataGridLib.DataGrid grdtmp 
      Height          =   3300
      Left            =   90
      TabIndex        =   87
      Top             =   3285
      Visible         =   0   'False
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   5821
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "Arial"
         Size            =   11.25
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
      Left            =   1485
      TabIndex        =   0
      Top             =   30
      Width           =   630
   End
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   11130
      Left            =   -150
      TabIndex        =   51
      Top             =   -15
      Width           =   18660
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
         Left            =   11430
         MaxLength       =   15
         TabIndex        =   64
         Top             =   10635
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00FCE4F1&
         Height          =   2430
         Left            =   210
         TabIndex        =   52
         Top             =   -90
         Width           =   18435
         Begin VB.TextBox txtphsearch 
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   5310
            TabIndex        =   242
            Top             =   465
            Width           =   1560
         End
         Begin VB.CommandButton CMDPre 
            Caption         =   "<<Pre&vious (Ctrl + Left)"
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
            Left            =   16755
            TabIndex        =   213
            Top             =   390
            Width           =   1410
         End
         Begin VB.CommandButton cmdnxt 
            Caption         =   "&Next>>      (Ctrl + Right)"
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
            Left            =   16755
            TabIndex        =   212
            Top             =   930
            Width           =   1410
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
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   11550
            MaxLength       =   35
            TabIndex        =   210
            Top             =   765
            Width           =   2745
         End
         Begin VB.TextBox TxtOrder 
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
            Height          =   585
            Left            =   11805
            MaxLength       =   300
            TabIndex        =   12
            Top             =   1110
            Width           =   2490
         End
         Begin VB.TextBox TXTTYPE 
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
            Height          =   300
            Left            =   8145
            TabIndex        =   9
            Top             =   1845
            Width           =   630
         End
         Begin VB.ComboBox cmbtype 
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
            ItemData        =   "FrmGSTBR.frx":030A
            Left            =   8850
            List            =   "FrmGSTBR.frx":031A
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1845
            Width           =   1620
         End
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   11370
            MaxLength       =   35
            TabIndex        =   14
            Top             =   2055
            Width           =   2925
         End
         Begin VB.TextBox TxtCode 
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   90
            TabIndex        =   3
            Top             =   465
            Width           =   1470
         End
         Begin VB.TextBox TxtVehicle 
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
            Left            =   11805
            MaxLength       =   35
            TabIndex        =   13
            Top             =   1710
            Width           =   2490
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EDF0F3&
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
            Height          =   660
            Left            =   10635
            TabIndex        =   85
            Top             =   90
            Width           =   3705
            Begin VB.Label lbladdress 
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
               ForeColor       =   &H00FF0000&
               Height          =   435
               Left            =   45
               TabIndex        =   11
               Top             =   180
               Width           =   3615
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00EDF0F3&
            Caption         =   "Billing Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1710
            Left            =   6885
            TabIndex        =   76
            Top             =   90
            Width           =   3735
            Begin VB.TextBox TxtBillName 
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
               Height          =   330
               Left            =   30
               MaxLength       =   100
               TabIndex        =   7
               Top             =   225
               Width           =   3645
            End
            Begin MSForms.TextBox TxtBillAddress 
               Height          =   1095
               Left            =   30
               TabIndex        =   8
               Top             =   570
               Width           =   3645
               VariousPropertyBits=   -1400879077
               ForeColor       =   16711680
               MaxLength       =   150
               BorderStyle     =   1
               Size            =   "6429;1931"
               SpecialEffect   =   0
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
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
            Height          =   300
            Left            =   930
            TabIndex        =   2
            Top             =   1275
            Width           =   645
         End
         Begin VB.TextBox TXTDEALER 
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   1590
            TabIndex        =   4
            Top             =   465
            Width           =   3705
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   300
            Left            =   3255
            TabIndex        =   1
            Top             =   150
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
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
            Height          =   870
            Left            =   1590
            TabIndex        =   5
            Top             =   840
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   1535
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo CMBDISTI 
            Height          =   1845
            Left            =   14355
            TabIndex        =   15
            Top             =   420
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   3254
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
         Begin MSDataListLib.DataCombo CMBBRNCH 
            Height          =   330
            Left            =   1590
            TabIndex        =   155
            Top             =   2055
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   16711680
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
         Begin VB.Label lblOB 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   16755
            TabIndex        =   245
            Top             =   1860
            Width           =   1635
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Old Balance"
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
            Index           =   68
            Left            =   16905
            TabIndex        =   244
            Top             =   1635
            Width           =   1305
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "4.   MRP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   165
            Index           =   66
            Left            =   9150
            TabIndex        =   243
            Top             =   2160
            Width           =   675
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4650
            TabIndex        =   219
            Top             =   150
            Width           =   930
         End
         Begin VB.Label lblsuf 
            BackColor       =   &H00DFF9F5&
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
            Height          =   270
            Left            =   2085
            TabIndex        =   218
            Top             =   150
            Width           =   510
         End
         Begin VB.Label lblpre 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DFF9F5&
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
            Height          =   270
            Left            =   945
            TabIndex        =   217
            Top             =   150
            Width           =   450
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cr. Days"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   45
            TabIndex        =   75
            Top             =   1290
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "GSTin No."
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
            Left            =   10635
            TabIndex        =   211
            Top             =   795
            Width           =   870
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DL Nos."
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
            Left            =   14430
            TabIndex        =   167
            Top             =   1875
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblIGST 
            BackColor       =   &H00FEF1FC&
            Height          =   285
            Left            =   5715
            TabIndex        =   166
            Top             =   2100
            Width           =   690
         End
         Begin VB.Label lblsubdealer 
            Height          =   405
            Left            =   60
            TabIndex        =   162
            Top             =   2475
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Office"
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
            Index           =   60
            Left            =   45
            TabIndex        =   156
            Top             =   2085
            Width           =   1440
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Order No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   55
            Left            =   10635
            TabIndex        =   154
            Top             =   1350
            Width           =   1110
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "3.   VP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   165
            Index           =   54
            Left            =   8430
            TabIndex        =   93
            Top             =   2175
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "2.   WS"
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
            Height          =   165
            Index           =   51
            Left            =   7635
            TabIndex        =   91
            Top             =   2175
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "1.   RT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   165
            Index           =   39
            Left            =   6885
            TabIndex        =   90
            Top             =   2160
            Width           =   585
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Billing Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   50
            Left            =   6945
            TabIndex        =   89
            Top             =   1860
            Width           =   1110
         End
         Begin MSForms.ComboBox TXTAREA 
            Height          =   315
            Left            =   1590
            TabIndex        =   6
            Top             =   1725
            Width           =   5280
            VariousPropertyBits=   746604571
            ForeColor       =   16711680
            MaxLength       =   200
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "9313;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            DropButtonStyle =   0
            BorderColor     =   0
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   47
            Left            =   10635
            TabIndex        =   88
            Top             =   1710
            Width           =   1110
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
            Index           =   35
            Left            =   10635
            TabIndex        =   86
            Top             =   2040
            Width           =   660
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Area / Delivery"
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
            Index           =   40
            Left            =   60
            TabIndex        =   80
            Top             =   1725
            Width           =   1545
         End
         Begin VB.Label Label1 
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
            Height          =   330
            Index           =   3
            Left            =   14355
            TabIndex        =   77
            Top             =   165
            Width           =   720
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
            Left            =   45
            TabIndex        =   66
            Top             =   870
            Width           =   1260
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   2745
            TabIndex        =   65
            Top             =   150
            Width           =   570
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
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   56
            Top             =   150
            Width           =   780
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   1
            Left            =   105
            TabIndex        =   55
            Top             =   495
            Visible         =   0   'False
            Width           =   1170
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
            Height          =   315
            Left            =   1590
            TabIndex        =   54
            Top             =   930
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label LBLBILLNO 
            Alignment       =   2  'Center
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1425
            TabIndex        =   53
            Top             =   135
            Width           =   645
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FCE4F1&
         ForeColor       =   &H008080FF&
         Height          =   4335
         Left            =   210
         TabIndex        =   57
         Top             =   2220
         Width           =   18390
         Begin VB.CommandButton cmdratechange 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            Caption         =   "Rate Change"
            Height          =   450
            Left            =   17505
            MaskColor       =   &H008080FF&
            Style           =   1  'Graphical
            TabIndex        =   240
            Top             =   3825
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.CommandButton CmdReset 
            Appearance      =   0  'Flat
            BackColor       =   &H00CBEFDE&
            Caption         =   "Reset"
            Height          =   450
            Left            =   17505
            MaskColor       =   &H008080FF&
            Style           =   1  'Graphical
            TabIndex        =   239
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FCE4F1&
            Height          =   4275
            Left            =   14220
            TabIndex        =   170
            Top             =   30
            Width           =   3285
            Begin VB.TextBox lblcomamt 
               Alignment       =   2  'Center
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
               Height          =   360
               Left            =   840
               TabIndex        =   223
               Top             =   2730
               Width           =   870
            End
            Begin VB.TextBox txtcommper 
               Alignment       =   2  'Center
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
               Height          =   360
               Left            =   45
               TabIndex        =   222
               Top             =   2730
               Width           =   615
            End
            Begin VB.CommandButton CmdTax 
               Appearance      =   0  'Flat
               BackColor       =   &H000040C0&
               Caption         =   "&Tax Print"
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
               Left            =   2430
               MaskColor       =   &H008080FF&
               Style           =   1  'Graphical
               TabIndex        =   179
               Top             =   4755
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.TextBox TxtTaxPrint 
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
               Height          =   420
               Left            =   1920
               TabIndex        =   178
               Top             =   4770
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.CommandButton cmdreturn 
               Appearance      =   0  'Flat
               BackColor       =   &H000080FF&
               Caption         =   "&Exchange"
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
               Left            =   1935
               MaskColor       =   &H008080FF&
               Style           =   1  'Graphical
               TabIndex        =   177
               Top             =   3780
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.TextBox Txthandle 
               Appearance      =   0  'Flat
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
               Left            =   1980
               TabIndex        =   176
               Top             =   3120
               Width           =   1230
            End
            Begin VB.TextBox lblhandle 
               BorderStyle     =   0  'None
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
               Left            =   75
               TabIndex        =   175
               Text            =   "Handling Charge"
               Top             =   3165
               Width           =   1875
            End
            Begin VB.TextBox lblFrieght 
               BorderStyle     =   0  'None
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
               Left            =   75
               TabIndex        =   174
               Text            =   "Frieght Charge"
               Top             =   3495
               Width           =   1875
            End
            Begin VB.TextBox TxtFrieght 
               Appearance      =   0  'Flat
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
               Height          =   315
               Left            =   1980
               TabIndex        =   173
               Top             =   3465
               Width           =   1230
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
               Height          =   435
               Left            =   60
               Style           =   1  'Graphical
               TabIndex        =   172
               Top             =   3795
               Visible         =   0   'False
               Width           =   1860
            End
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
               Left            =   1695
               MaskColor       =   &H008080FF&
               Style           =   1  'Graphical
               TabIndex        =   171
               Top             =   5175
               Width           =   1530
            End
            Begin VB.Label LBLGROSSAMT 
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
               ForeColor       =   &H80000008&
               Height          =   405
               Left            =   30
               TabIndex        =   250
               Top             =   300
               Width           =   1710
            End
            Begin VB.Label Label2 
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   675
               TabIndex        =   224
               Top             =   2805
               Width           =   165
            End
            Begin VB.Label LBLNETCOST 
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
               Height          =   405
               Left            =   1785
               TabIndex        =   206
               Top             =   1485
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label LBLNETPROFIT 
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
               Height          =   405
               Left            =   1785
               TabIndex        =   205
               Top             =   2085
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "COMM AMOUNT"
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
               Index           =   53
               Left            =   45
               TabIndex        =   204
               Top             =   2475
               Width           =   1575
            End
            Begin VB.Label lblActAmt 
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
               Height          =   450
               Left            =   60
               TabIndex        =   203
               Top             =   3450
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label LBLRETAMT 
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
               Height          =   405
               Left            =   45
               TabIndex        =   202
               Top             =   2085
               Width           =   1710
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   49
               Left            =   45
               TabIndex        =   201
               Top             =   1860
               Width           =   1575
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "PROFIT AMT"
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
               Height          =   450
               Index           =   45
               Left            =   1755
               TabIndex        =   200
               Top             =   1860
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label LblProfitAmt 
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
               Height          =   405
               Left            =   3195
               TabIndex        =   199
               Top             =   2085
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "PROFIT %"
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
               Height          =   375
               Index           =   44
               Left            =   1755
               TabIndex        =   198
               Top             =   2475
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label LblProfitPerc 
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
               Height          =   405
               Left            =   1755
               TabIndex        =   197
               Top             =   2700
               Visible         =   0   'False
               Width           =   1440
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
               ForeColor       =   &H80000008&
               Height          =   405
               Left            =   1470
               TabIndex        =   196
               Top             =   3375
               Visible         =   0   'False
               Width           =   1545
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   6
               Left            =   45
               TabIndex        =   195
               Top             =   90
               Width           =   1485
            End
            Begin VB.Label lblnetamount 
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
               ForeColor       =   &H000000FF&
               Height          =   405
               Left            =   45
               TabIndex        =   194
               Top             =   885
               Width           =   1710
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   23
               Left            =   45
               TabIndex        =   193
               Top             =   675
               Width           =   1440
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   25
               Left            =   1755
               TabIndex        =   192
               Top             =   90
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
               Height          =   405
               Left            =   1785
               TabIndex        =   191
               Top             =   885
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
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   405
               Left            =   3195
               TabIndex        =   190
               Top             =   1485
               Visible         =   0   'False
               Width           =   1440
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
               Height          =   495
               Left            =   285
               TabIndex        =   189
               Top             =   4605
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ITEM NET COST"
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
               Height          =   375
               Index           =   27
               Left            =   1755
               TabIndex        =   188
               Top             =   1260
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   28
               Left            =   285
               TabIndex        =   187
               Top             =   4665
               Visible         =   0   'False
               Width           =   1440
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
               Height          =   495
               Left            =   195
               TabIndex        =   186
               Top             =   4260
               Visible         =   0   'False
               Width           =   1440
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   31
               Left            =   165
               TabIndex        =   185
               Top             =   4350
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   4
               Left            =   45
               TabIndex        =   184
               Top             =   1260
               Width           =   1500
            End
            Begin VB.Label LBLDISCAMT 
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
               Height          =   405
               Left            =   45
               TabIndex        =   183
               Top             =   1485
               Width           =   1710
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Com Amt"
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
               Index           =   36
               Left            =   1740
               TabIndex        =   182
               Top             =   4515
               Visible         =   0   'False
               Width           =   1455
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
               Height          =   400
               Left            =   1785
               TabIndex        =   181
               Top             =   300
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL PROFIT"
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
               Left            =   1755
               TabIndex        =   180
               Top             =   675
               Visible         =   0   'False
               Width           =   1425
            End
         End
         Begin VB.TextBox TXTsample 
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
            Left            =   11520
            MaxLength       =   10
            TabIndex        =   94
            Top             =   270
            Visible         =   0   'False
            Width           =   1350
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
            Left            =   7950
            TabIndex        =   74
            Top             =   5805
            Visible         =   0   'False
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
            Left            =   13260
            TabIndex        =   73
            Top             =   3885
            Width           =   930
         End
         Begin VB.OptionButton OPTDISCPERCENT 
            BackColor       =   &H0080C0FF&
            Caption         =   "Disc %"
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
            Left            =   11175
            TabIndex        =   72
            Top             =   3885
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton OptDiscAmt 
            BackColor       =   &H0080C0FF&
            Caption         =   "Disc Amt"
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
            Left            =   12120
            TabIndex        =   71
            Top             =   3885
            Width           =   1125
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   3690
            Left            =   30
            TabIndex        =   16
            Top             =   150
            Width           =   14145
            _ExtentX        =   24950
            _ExtentY        =   6509
            _Version        =   393216
            Rows            =   1
            Cols            =   47
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   420
            BackColor       =   16050128
            BackColorFixed  =   12320767
            ForeColorFixed  =   255
            FocusRect       =   2
            AllowUserResizing=   1
            Appearance      =   0
            GridLineWidth   =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "RT"
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
            Index           =   21
            Left            =   75
            TabIndex        =   238
            Top             =   3900
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "WS"
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
            Index           =   22
            Left            =   1605
            TabIndex        =   237
            Top             =   3900
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "VP"
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
            Index           =   30
            Left            =   3180
            TabIndex        =   236
            Top             =   3900
            Width           =   585
         End
         Begin VB.Label lblretail 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   675
            TabIndex        =   235
            Top             =   3900
            Width           =   855
         End
         Begin VB.Label lblwsale 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2235
            TabIndex        =   234
            Top             =   3900
            Width           =   855
         End
         Begin VB.Label lblvan 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   3855
            TabIndex        =   233
            Top             =   3900
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L.R.Price"
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
            Index           =   42
            Left            =   6285
            TabIndex        =   232
            Top             =   3900
            Width           =   885
         End
         Begin VB.Label lblcrtnpack 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7275
            TabIndex        =   231
            Top             =   3900
            Width           =   645
         End
         Begin VB.Label lblcase 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   9015
            TabIndex        =   230
            Top             =   3900
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "per"
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
            Index           =   52
            Left            =   9900
            TabIndex        =   229
            Top             =   3900
            Width           =   405
         End
         Begin VB.Label lblLWPrice 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   10410
            TabIndex        =   228
            Top             =   3900
            Width           =   660
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L.W.Price"
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
            Index           =   63
            Left            =   8010
            TabIndex        =   227
            Top             =   3900
            Width           =   900
         End
         Begin VB.Label LBLMRP 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   5415
            TabIndex        =   226
            Top             =   3900
            Width           =   810
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
            Height          =   300
            Index           =   67
            Left            =   4680
            TabIndex        =   225
            Top             =   3900
            Width           =   630
         End
         Begin VB.Label lblOr_Pack 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9210
            TabIndex        =   92
            Top             =   4335
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "on Pack"
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
            Index           =   34
            Left            =   9405
            TabIndex        =   79
            Top             =   4920
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "C/S"
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
            Index           =   33
            Left            =   7965
            TabIndex        =   78
            Top             =   4920
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FCE4F1&
         ForeColor       =   &H008080FF&
         Height          =   4365
         Left            =   210
         TabIndex        =   95
         Top             =   6465
         Width           =   18450
         Begin VB.CommandButton CmdExport 
            Caption         =   "Export "
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
            Left            =   13860
            TabIndex        =   249
            Top             =   1530
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtretaildummy 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   16455
            MaxLength       =   6
            TabIndex        =   220
            Top             =   1095
            Width           =   630
         End
         Begin VB.TextBox TxtDiscAmt 
            Appearance      =   0  'Flat
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
            Height          =   450
            Left            =   14985
            MaxLength       =   7
            TabIndex        =   32
            Top             =   375
            Width           =   885
         End
         Begin VB.TextBox txtNetrate 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   17445
            TabIndex        =   214
            Top             =   1095
            Width           =   945
         End
         Begin VB.TextBox TxtSpec 
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
            Height          =   1050
            Left            =   30
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   208
            Top             =   1770
            Width           =   5025
         End
         Begin VB.TextBox TrxRYear 
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
            TabIndex        =   207
            Top             =   4095
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox lblunit 
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
            Height          =   450
            Left            =   7320
            MaxLength       =   5
            TabIndex        =   22
            Top             =   375
            Width           =   390
         End
         Begin VB.TextBox TxtCessAmt 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   16500
            MaxLength       =   5
            TabIndex        =   34
            Top             =   450
            Width           =   630
         End
         Begin VB.TextBox TxtCessPer 
            Appearance      =   0  'Flat
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
            Height          =   450
            Left            =   15885
            MaxLength       =   5
            TabIndex        =   33
            Top             =   375
            Width           =   600
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   450
            Left            =   10770
            TabIndex        =   27
            Top             =   375
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   794
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
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
         Begin VB.PictureBox picUnchecked 
            Height          =   285
            Left            =   16080
            Picture         =   "FrmGSTBR.frx":032F
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   160
            Top             =   3105
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.PictureBox picChecked 
            Height          =   285
            Left            =   15435
            Picture         =   "FrmGSTBR.frx":0671
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   159
            Top             =   3075
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CheckBox CHKSELECT 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   375
            Left            =   3495
            TabIndex        =   158
            Top             =   3285
            Width           =   1545
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Thermal Print"
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
            Left            =   7380
            TabIndex        =   40
            Top             =   855
            Width           =   870
         End
         Begin VB.TextBox txtPrintname 
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
            Height          =   450
            Left            =   1155
            TabIndex        =   36
            Top             =   1065
            Width           =   3900
         End
         Begin VB.CheckBox chkTerms 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Terms && Conditions"
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
            Left            =   5085
            TabIndex        =   49
            Top             =   1275
            Width           =   2145
         End
         Begin VB.TextBox Terms1 
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
            Left            =   5070
            MaxLength       =   255
            TabIndex        =   50
            Top             =   1515
            Width           =   7065
         End
         Begin VB.CommandButton Command2 
            Height          =   435
            Left            =   15690
            TabIndex        =   123
            Top             =   3330
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox Txtrcvd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   13530
            MaxLength       =   7
            TabIndex        =   47
            Top             =   1020
            Width           =   1290
         End
         Begin VB.TextBox txtcommi 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   14280
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   35
            Top             =   2835
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Print Commission"
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
            Left            =   17730
            TabIndex        =   122
            Top             =   2370
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtName1 
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
            Height          =   435
            Left            =   2055
            MaxLength       =   15
            TabIndex        =   19
            Top             =   390
            Width           =   1065
         End
         Begin VB.TextBox TxtCN 
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
            Height          =   450
            Left            =   8460
            MaxLength       =   6
            TabIndex        =   121
            Top             =   3660
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton CMDPRINT 
            Caption         =   "&PRINT -A4"
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
            Left            =   8280
            TabIndex        =   41
            Top             =   855
            Width           =   990
         End
         Begin VB.CommandButton CmdPrintA5 
            Caption         =   "PRINT -A&5"
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
            Left            =   9270
            TabIndex        =   42
            Top             =   855
            Width           =   990
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
            Height          =   435
            Left            =   30
            MaxLength       =   30
            TabIndex        =   120
            Top             =   1080
            Width           =   450
         End
         Begin VB.TextBox TxtWarranty_type 
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
            Height          =   435
            Left            =   495
            MaxLength       =   30
            TabIndex        =   119
            Top             =   1080
            Width           =   645
         End
         Begin VB.CommandButton cmdRefresh 
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
            Height          =   405
            Left            =   11385
            TabIndex        =   43
            Top             =   855
            Width           =   885
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
            Left            =   14745
            TabIndex        =   118
            Top             =   3585
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
            TabIndex        =   117
            Top             =   3660
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
            TabIndex        =   116
            Top             =   3600
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
            TabIndex        =   115
            Top             =   3600
            Visible         =   0   'False
            Width           =   690
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
            Height          =   450
            Left            =   7725
            MaxLength       =   15
            TabIndex        =   23
            Top             =   375
            Width           =   990
         End
         Begin VB.TextBox TXTITEMCODE 
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
            Height          =   450
            Left            =   510
            TabIndex        =   18
            Top             =   375
            Width           =   1530
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
            Left            =   5850
            TabIndex        =   38
            Top             =   855
            Width           =   750
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
            Left            =   6630
            TabIndex        =   39
            Top             =   855
            Width           =   750
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
            Left            =   12285
            TabIndex        =   44
            Top             =   855
            Width           =   885
         End
         Begin VB.TextBox TXTDISC 
            Appearance      =   0  'Flat
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
            Height          =   450
            Left            =   14415
            MaxLength       =   5
            TabIndex        =   31
            Top             =   375
            Width           =   555
         End
         Begin VB.TextBox TXTTAX 
            Alignment       =   2  'Center
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
            Height          =   450
            Left            =   11685
            MaxLength       =   4
            TabIndex        =   28
            Top             =   375
            Width           =   570
         End
         Begin VB.TextBox TXTQTY 
            Appearance      =   0  'Flat
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
            Height          =   450
            Left            =   8730
            MaxLength       =   8
            TabIndex        =   24
            Top             =   375
            Width           =   750
         End
         Begin VB.TextBox TXTPRODUCT 
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
            Height          =   435
            Left            =   3135
            TabIndex        =   20
            Top             =   390
            Width           =   3675
         End
         Begin VB.TextBox TXTSLNO 
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
            Height          =   450
            Left            =   30
            TabIndex        =   17
            Top             =   375
            Width           =   465
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
            Height          =   405
            Left            =   5085
            TabIndex        =   37
            Top             =   855
            Width           =   750
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
            Height          =   450
            Left            =   9960
            MaxLength       =   6
            TabIndex        =   26
            Top             =   375
            Width           =   795
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
            Height          =   405
            Left            =   13185
            TabIndex        =   114
            Top             =   855
            Width           =   300
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
            Height          =   405
            Left            =   90
            TabIndex        =   113
            Top             =   2820
            Visible         =   0   'False
            Width           =   420
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
            Height          =   405
            Left            =   555
            TabIndex        =   112
            Top             =   2790
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.TextBox TXTFREE 
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
            Height          =   450
            Left            =   9495
            MaxLength       =   7
            TabIndex        =   25
            Top             =   375
            Width           =   450
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
            Left            =   12795
            MaxLength       =   6
            TabIndex        =   111
            Top             =   4290
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00EDF0F3&
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
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   12285
            TabIndex        =   110
            Top             =   2250
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00EDF0F3&
            Caption         =   "TAX %"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   12270
            TabIndex        =   109
            Top             =   1995
            Width           =   945
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
            Left            =   1395
            MaxLength       =   6
            TabIndex        =   107
            Top             =   3180
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.OptionButton optnet 
            BackColor       =   &H00EDF0F3&
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
            Height          =   210
            Left            =   13230
            TabIndex        =   108
            Top             =   2010
            Width           =   720
         End
         Begin VB.TextBox TXTRETAILNOTAX 
            Appearance      =   0  'Flat
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
            Height          =   450
            Left            =   12270
            MaxLength       =   9
            TabIndex        =   29
            Top             =   375
            Width           =   1050
         End
         Begin VB.TextBox txtretail 
            Appearance      =   0  'Flat
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
            Height          =   450
            Left            =   13335
            MaxLength       =   9
            TabIndex        =   30
            Top             =   375
            Width           =   1080
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
            Left            =   3120
            MaxLength       =   6
            TabIndex        =   106
            Top             =   4035
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox TxtRetailmode 
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
            Left            =   5070
            MaxLength       =   6
            TabIndex        =   105
            Top             =   4050
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtcategory 
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
            Height          =   435
            Left            =   510
            MaxLength       =   15
            TabIndex        =   104
            Top             =   3615
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox TXTAPPENDQTY 
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
            Left            =   13125
            MaxLength       =   8
            TabIndex        =   103
            Top             =   4425
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox TXTFREEAPPEND 
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
            Left            =   15045
            MaxLength       =   8
            TabIndex        =   102
            Top             =   3915
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox TXTAPPENDTOTAL 
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
            Left            =   15045
            MaxLength       =   8
            TabIndex        =   101
            Top             =   3555
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox txtappendcomm 
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
            Left            =   15015
            MaxLength       =   8
            TabIndex        =   100
            Top             =   3750
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmddeleteall 
            Caption         =   "Cancel Bill"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   12270
            TabIndex        =   46
            Top             =   1545
            Width           =   1215
         End
         Begin VB.TextBox txtOutstanding 
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
            Left            =   120
            TabIndex        =   99
            Top             =   3600
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.CheckBox Chkcancel 
            Appearance      =   0  'Flat
            BackColor       =   &H00400000&
            Caption         =   "Cancel Bill"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   12285
            TabIndex        =   45
            Top             =   1305
            Width           =   1200
         End
         Begin VB.TextBox LblPack 
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
            Height          =   450
            Left            =   6825
            MaxLength       =   8
            TabIndex        =   21
            Top             =   375
            Width           =   480
         End
         Begin VB.Frame FrmeType 
            BackColor       =   &H0080C0FF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   720
            Left            =   10260
            TabIndex        =   96
            Top             =   750
            Width           =   1110
            Begin VB.OptionButton OptNormal 
               BackColor       =   &H000080FF&
               Caption         =   "&Full Pack"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   30
               TabIndex        =   98
               Top             =   135
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton OptLoose 
               BackColor       =   &H000080FF&
               Caption         =   "&Loose"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   45
               TabIndex        =   97
               Top             =   405
               Width           =   1020
            End
         End
         Begin MSFlexGridLib.MSFlexGrid GRDRECEIPT 
            Height          =   1125
            Left            =   0
            TabIndex        =   124
            Top             =   3405
            Visible         =   0   'False
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   1984
            _Version        =   393216
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   450
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            GridLineWidth   =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid grdEXPIRYLIST 
            Height          =   1845
            Left            =   5070
            TabIndex        =   157
            Top             =   1815
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   3254
            _Version        =   393216
            Rows            =   1
            Cols            =   9
            FixedRows       =   0
            RowHeightMin    =   300
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            BackColorBkg    =   12632256
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            FillStyle       =   1
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "F5 - Print"
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
            Height          =   180
            Index           =   76
            Left            =   15015
            TabIndex        =   248
            Top             =   1545
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "F6 - Save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   77
            Left            =   16035
            TabIndex        =   247
            Top             =   1545
            Width           =   1260
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "F3 - Tender Cash"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   180
            Index           =   78
            Left            =   15015
            TabIndex        =   246
            Top             =   1770
            Width           =   1650
         End
         Begin VB.Label LBLUNBILL 
            Height          =   225
            Left            =   15120
            TabIndex        =   241
            Top             =   2685
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblcomtype 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   17145
            TabIndex        =   221
            Top             =   1185
            Width           =   210
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Disc Amt"
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
            Index           =   65
            Left            =   14985
            TabIndex        =   216
            Top             =   150
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Rate Calc"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   225
            Index           =   64
            Left            =   17445
            TabIndex        =   215
            Top             =   855
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Product Specifications"
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
            Index           =   59
            Left            =   30
            TabIndex        =   209
            Top             =   1530
            Width           =   5025
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Add Cess Amt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Index           =   62
            Left            =   16500
            TabIndex        =   169
            Top             =   150
            Width           =   630
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
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   29
            Left            =   15885
            TabIndex        =   168
            Top             =   150
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Expiry"
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
            Index           =   61
            Left            =   10770
            TabIndex        =   165
            Top             =   150
            Width           =   900
         End
         Begin VB.Label lblbarcode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   30
            TabIndex        =   164
            Top             =   3645
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.Label lblactqty 
            Height          =   375
            Left            =   4200
            TabIndex        =   163
            Top             =   1560
            Width           =   735
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
            Height          =   450
            Left            =   2250
            TabIndex        =   153
            Top             =   3840
            Visible         =   0   'False
            Width           =   1560
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
            Index           =   58
            Left            =   9495
            TabIndex        =   152
            Top             =   150
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Name to be Printed"
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
            Index           =   38
            Left            =   1155
            TabIndex        =   151
            Top             =   840
            Width           =   3900
         End
         Begin VB.Label lblbalance 
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
            ForeColor       =   &H000000C0&
            Height          =   480
            Left            =   14835
            TabIndex        =   48
            Top             =   1020
            Width           =   1290
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bal. Amt"
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
            Index           =   57
            Left            =   14835
            TabIndex        =   150
            Top             =   795
            Width           =   780
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Rcvd Cash"
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
            Index           =   56
            Left            =   13515
            TabIndex        =   149
            Top             =   795
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Comm"
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
            Index           =   46
            Left            =   16455
            TabIndex        =   148
            Top             =   855
            Width           =   975
         End
         Begin VB.Label LBLTYPE 
            Caption         =   "GI"
            Height          =   330
            Left            =   4230
            TabIndex        =   147
            Top             =   2910
            Visible         =   0   'False
            Width           =   720
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
            Left            =   4860
            TabIndex        =   146
            Top             =   3675
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
            Left            =   13605
            TabIndex        =   145
            Top             =   3585
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
            Left            =   6540
            TabIndex        =   144
            Top             =   3900
            Visible         =   0   'False
            Width           =   1080
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
            TabIndex        =   143
            Top             =   3690
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label LBLSUBTOTAL 
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
            Height          =   450
            Left            =   17145
            TabIndex        =   142
            Top             =   375
            Width           =   1260
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Serial No."
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
            Left            =   7725
            TabIndex        =   141
            Top             =   150
            Width           =   990
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
            Left            =   1155
            TabIndex        =   140
            Top             =   4425
            Visible         =   0   'False
            Width           =   1080
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
            Left            =   17145
            TabIndex        =   139
            Top             =   150
            Width           =   1260
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
            Left            =   14430
            TabIndex        =   138
            Top             =   150
            Width           =   540
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
            Left            =   11685
            TabIndex        =   137
            Top             =   150
            Width           =   570
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
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   11
            Left            =   13335
            TabIndex        =   136
            Top             =   150
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
            Height          =   225
            Index           =   10
            Left            =   8730
            TabIndex        =   135
            Top             =   150
            Width           =   750
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
            Left            =   2055
            TabIndex        =   134
            Top             =   150
            Width           =   4755
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
            Left            =   30
            TabIndex        =   133
            Top             =   150
            Width           =   465
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
            Left            =   9960
            TabIndex        =   132
            Top             =   150
            Width           =   795
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
            TabIndex        =   131
            Top             =   3435
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label Lblprice 
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
            Index           =   30
            Left            =   12270
            TabIndex        =   130
            Top             =   150
            Width           =   1050
         End
         Begin VB.Label lblP_Rate 
            Caption         =   "0"
            Height          =   390
            Left            =   13200
            TabIndex        =   129
            Top             =   3690
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Barcode / Code"
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
            Index           =   43
            Left            =   510
            TabIndex        =   128
            Top             =   150
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Warranty"
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
            Index           =   48
            Left            =   30
            TabIndex        =   127
            Top             =   840
            Width           =   1110
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Unit"
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
            Left            =   7320
            TabIndex        =   126
            Top             =   150
            Width           =   390
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
            Index           =   37
            Left            =   6825
            TabIndex        =   125
            Top             =   150
            Width           =   480
         End
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   13155
      TabIndex        =   67
      Top             =   3090
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1217
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grdcount 
      Height          =   5145
      Left            =   0
      TabIndex        =   161
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   9075
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      FillStyle       =   1
      SelectionMode   =   1
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
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   -15
      TabIndex        =   70
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11355
      TabIndex        =   69
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   68
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMGSTBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BR_FLAG As Boolean
Dim BR_CODE As New ADODB.Recordset
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim PHYCODE As New ADODB.Recordset
Dim PHYCODEFLAG As Boolean
Dim TMPREC As New ADODB.Recordset
Dim TMPFLAG As Boolean
Dim ACT_REC As New ADODB.Recordset
Dim ACT_AGNT As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim AGNT_FLAG As Boolean
Dim PHY_BATCH As New ADODB.Recordset
Dim BATCH_FLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean
Dim PHY_PRERATE As New ADODB.Recordset
Dim PRERATE_FLAG As Boolean
Dim SERIAL_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_STOCK As Double
Dim cr_days As Boolean
Dim M_ADD, M_EDIT, NEW_BILL As Boolean
Dim OLD_BILL As Boolean
Dim Small_Print, Dos_Print, ST_PRINT, Tax_Print As Boolean
Dim CHANGE_ADDRESS, CHANGE_NAME As Boolean
Dim item_change As Boolean
Dim thermalprn As Boolean
Dim mrpplus As Boolean
Dim preclick As Boolean
Dim cancel_bill As Boolean

Private Sub CMBDISTI_GotFocus()
    CMBDISTI.BackColor = &H98F3C1
End Sub

Private Sub CMBDISTI_LostFocus()
    CMBDISTI.BackColor = vbWhite
End Sub

Private Sub cmbtype_Click()
    If cmbtype.ListIndex = 0 Then TXTTYPE.text = 1
    If cmbtype.ListIndex = 1 Then TXTTYPE.text = 2
    If cmbtype.ListIndex = 2 Then TXTTYPE.text = 3
    If cmbtype.ListIndex = 3 Then TXTTYPE.text = 4
End Sub

Private Sub cmbtype_GotFocus()
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub cmdadd_GotFocus()
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    If UCase(txtcategory.text) = "PARDHA" Or UCase(txtcategory.text) = "CLOTHES" Then
        If Val(TXTRETAILNOTAX.text) < 1000 Then
            TXTTAX.text = "5"
        Else
            TXTTAX.text = "12"
        End If
    End If
    If Val(txtretail.text) = 0 Then
        Call TXTRETAILNOTAX_LostFocus
    End If
    If Val(TXTRETAILNOTAX.text) = 0 Then
        Call TXTRETAIL_LostFocus
    End If
    
    If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
        Call TXTRETAILNOTAX_LostFocus
    Else
        If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
            Call TXTRETAILNOTAX_LostFocus
        Else
            Call TXTRETAIL_LostFocus
        End If
    End If

    Call TXTDISC_LostFocus
End Sub

Private Sub CmdDeleteAll_Click()
    Dim i As Long
    Dim n As Long
    If Chkcancel.Value = 0 Then Exit Sub
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rststock As ADODB.Recordset
'    If grdsales.Rows = 1 Then Exit Sub
'    If MsgBox("ARE YOU SURE YOU WANT TO CANCEL THE BILL!!!!!", vbYesNo + vbDefaultButton2, "DELETE!!!") = vbNo Then
'        Chkcancel.value = 0
'        Exit Sub
'    End If
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    If grdsales.rows > 1 Then
        If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
            MsgBox "Permission Denied", vbOKOnly, "Sales"
            Exit Sub
        End If
        If MsgBox("ARE YOU SURE YOU WANT TO CANCEL THE BILL!!!!!", vbYesNo + vbDefaultButton2, "DELETE!!!") = vbNo Then
            Chkcancel.Value = 0
            Exit Sub
        End If
    End If
    
    'db.Execute "delete From RTRXFILEVAN WHERE TRX_TYPE='CN' AND VCH_NO = " & Val(TxtCN.Text) & ""
    
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND BILL_NO = " & Val(txtBillNo.Text) & " AND BILL_TRX_TYPE = 'GI' AND TRX_TYPE = 'GI'", db, adOpenStatic, adLockOptimistic, adCmdText
'    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'        RSTTRXFILE!CHECK_FLAG = "N"
'        RSTTRXFILE!BILL_NO = Null
'        RSTTRXFILE!BILL_TRX_TYPE = Null
'        RSTTRXFILE!BILL_DATE = Null
'        RSTTRXFILE.Update
'    End If
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    
    db.Execute "delete From TRXSUBVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " "
    db.Execute "delete From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " "
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE TRX_TYPE='GI' AND VCH_NO = " & Val(TxtCN.text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
    With RSTTRXFILE
        Do Until .EOF
            If Not (UCase(RSTTRXFILE!Category) = "SERVICES" Or UCase(RSTTRXFILE!Category) = "SELF") Then
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                With rststock
                    If Not (.EOF And .BOF) Then
                        .Properties("Update Criteria").Value = adCriteriaKey
                        '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20)))
                        !ISSUE_QTY = !ISSUE_QTY + RSTTRXFILE!QTY
                        If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                        !ISSUE_VAL = !ISSUE_VAL + RSTTRXFILE!TRX_TOTAL
                        !CLOSE_QTY = !CLOSE_QTY - RSTTRXFILE!QTY
                        If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                        !CLOSE_VAL = !CLOSE_VAL - RSTTRXFILE!TRX_TOTAL
                        rststock.Update
                    End If
                End With
                rststock.Close
                Set rststock = Nothing
            End If
        RSTTRXFILE.MoveNext
        Loop
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    db.Execute "delete From RTRXFILEVAN WHERE TRX_TYPE='GI' AND VCH_NO = " & Val(TxtCN.text) & ""
    For n = 1 To grdsales.rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(n, 13) & "' AND CATEGORY <> 'SERVICES' AND CATEGORY <> 'SERVICE CHARGE' AND CATEGORY <> 'SELF'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").Value = adCriteriaKey
                '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20)))
                !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(n, 3)) * Val(grdsales.TextMatrix(n, 27)), 3)
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(n, 20)) * Val(grdsales.TextMatrix(n, 27)), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(n, 12))
                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20))) * Val(grdsales.TextMatrix(n, 27)), 3)
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(n, 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
           
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE TRX_TYPE = '" & Trim(grdsales.TextMatrix(n, 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(n, 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(n, 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(n, 43)) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20))) * Val(grdsales.TextMatrix(n, 27)), 3)
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20))) * Val(grdsales.TextMatrix(n, 27)), 3)
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        If grdsales.TextMatrix(n, 19) = "DN" Then
            Set RSTTRXFILE = New ADODB.Recordset
            'RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND TRX_TYPE = 'GI' AND VCH_NO = " & Val(grdsales.TextMatrix(n, 35)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(n, 37)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
            RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE TRX_TYPE = 'SI' AND VCH_NO = " & Val(grdsales.TextMatrix(n, 35)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(n, 37)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
            With RSTTRXFILE
                If Not (.EOF And .BOF) Then
                    RSTTRXFILE!check_flag = "N"
                    RSTTRXFILE!BILL_NO = 0 ' Null
                    'RSTTRXFILE!BILL_DATE = Null
                    RSTTRXFILE.Update
                End If
            End With
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        End If
    Next n
    grdsales.FixedRows = 0
    grdsales.rows = 1
    LBLTOTAL.Caption = ""
    LBLGROSSAMT.Caption = ""
    lblnetamount.Caption = ""
    TXTTOTALDISC.text = ""
    LBLRETAMT.Caption = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
    Txthandle.text = ""
    TxtFrieght.text = ""
    Call AppendSale
    Chkcancel.Value = 0
End Sub

Private Sub CMDDOS_Click()
    Chkcancel.Value = 0
    If grdsales.rows = 1 Then Exit Sub
    
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMASTVAN", db, adOpenForwardOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
'        If i > 3000 Then
'            TRXMAST.Close
'            Set TRXMAST = Nothing
'            Exit Sub
'        End If
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
'    If Not IsDate(TXTINVDATE.Text) Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    Else
'        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'    End If
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
    
'    If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.Text <> "" Then
'        MsgBox "Select Agent From List", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        CMBDISTI.SetFocus
'        Exit Sub
'    End If
            
'    If Trim(TXTAREA.Text) = "" Then
'        MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTAREA.SetFocus
'        Exit Sub
'    End If
    
    'If Val(txtcrdays.Text) = 0 Or DataList2.BoundText = "130000" Then
    Small_Print = False
    Dos_Print = True
    Set creditbill = Me
    If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
        CMDEXIT.Enabled = False
        Me.lblcredit.Caption = "0"
        Me.Generateprint
    Else
        Me.Enabled = False
        FRMDEBITRT.Show
    End If
End Sub

Private Sub CMDDOS_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            'TXTPRODUCT.Text = ""
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            txtappendcomm.text = ""
            TXTAPPENDTOTAL.text = ""
            txtretail.text = ""
            txtNetrate.text = ""
            txtBatch.text = ""
            TxtWarranty.text = ""
            TxtWarranty_type.text = ""
            TXTTAX.text = ""
            TXTRETAILNOTAX.text = ""
            TXTSALETYPE.text = ""
            TXTFREE.text = ""
            optnet.Value = True
            TxtMRP.text = ""
            txtmrpbt.text = ""
            txtretaildummy.text = ""
            lblcomtype.Caption = ""
            txtcommi.text = ""
            TxtRetailmode.text = ""
            
            TXTDISC.text = ""
            TxtDiscAmt.text = ""
            TxtCessAmt.text = ""
            TxtCessPer.text = ""
            LBLUNBILL.Caption = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TrxRYear.text = ""
            TXTUNIT.text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtDiscAmt.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub CmdPrintA5_Click()
    
    Chkcancel.Value = 0
    If grdsales.rows = 1 Then Exit Sub
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    thermalprn = False
    Tax_Print = False
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    'If Month(Date) >= 5 And Year(Date) >= 2021 Then Exit Sub
'    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then
'        'db.Execute "delete From USERS "
'        Exit Sub
'    End If
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMASTVAN", db, adOpenForwardOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
'        If i > 3000 Then
'            TRXMAST.Close
'            Set TRXMAST = Nothing
'            Exit Sub
'        End If
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
'    If Not IsDate(TXTINVDATE.Text) Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    Else
'        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'    End If
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
    
    If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.text <> "" Then
        MsgBox "Select Agent From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        CMBDISTI.SetFocus
        Exit Sub
    End If
            
'    If Trim(TXTAREA.Text) = "" Then
'        MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTAREA.SetFocus
'        Exit Sub
'    End If
    
    'If Val(txtcrdays.Text) = 0 Or DataList2.BoundText = "130000" Then
    Small_Print = True
    Dos_Print = False
    Set creditbill = Me
    If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
        CMDEXIT.Enabled = False
        Me.lblcredit.Caption = "0"
        Me.Generateprint
    Else
        Me.Enabled = False
        FRMDEBITRT.Show
    End If
End Sub

Private Sub CmdPrintA5_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            'TXTPRODUCT.Text = ""
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            txtappendcomm.text = ""
            TXTAPPENDTOTAL.text = ""
            txtretail.text = ""
            txtNetrate.text = ""
            txtBatch.text = ""
            TxtWarranty.text = ""
            TxtWarranty_type.text = ""
            TXTTAX.text = ""
            TXTRETAILNOTAX.text = ""
            TXTSALETYPE.text = ""
            TXTFREE.text = ""
            optnet.Value = True
            TxtMRP.text = ""
            txtmrpbt.text = ""
            txtretaildummy.text = ""
            lblcomtype.Caption = ""
            txtcommi.text = ""
            TxtRetailmode.text = ""
            
            TXTDISC.text = ""
            TxtDiscAmt.text = ""
            TxtCessAmt.text = ""
            TxtCessPer.text = ""
            LBLUNBILL.Caption = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TrxRYear.text = ""
            TXTUNIT.text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtDiscAmt.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub cmdratechange_Click()
    If grdsales.rows <= 1 Then Exit Sub
    Select Case cmbtype.ListIndex
        Case 0
            If (MsgBox("Are you sure you want to change the whole rate to Retail?", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then Exit Sub
        Case 1
            If (MsgBox("Are you sure you want to change the whole rate to Wholesale?", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then Exit Sub
        Case 2
            If (MsgBox("Are you sure you want to change the whole rate to VP?", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then Exit Sub
        Case 3
            If (MsgBox("Are you sure you want to change the whole rate to MRP?", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then Exit Sub
        Case Else
            MsgBox "Please select a valid Rate type from the list", , "SALES"
            On Error Resume Next
            cmbtype.SetFocus
            Exit Sub
    End Select
    On Error GoTo ERRHAND
     
    Dim i As Integer
    Dim Newprice As Double
    Dim RSTTRXFILE As ADODB.Recordset
    For i = 1 To grdsales.rows - 1
        Newprice = 0
        
        Select Case cmbtype.ListIndex
            Case 1
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
                With RSTTRXFILE
                    If Not (.EOF And .BOF) Then
                        If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                            Newprice = IIf(IsNull(RSTTRXFILE!P_WS), 0, RSTTRXFILE!P_WS)
                        Else
                            If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!P_WS), 0, RSTTRXFILE!P_WS)
                            Else
                                Newprice = IIf(IsNull(RSTTRXFILE!P_LWS), 0, RSTTRXFILE!P_LWS)
                            End If
                        End If
                    End If
                End With
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                If Newprice = 0 Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!P_WS), 0, RSTTRXFILE!P_WS)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_WS), 0, RSTTRXFILE!P_WS)
                                Else
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_LWS), 0, RSTTRXFILE!P_LWS)
                                End If
                            End If
                        End If
                    End With
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
            Case 3
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
                With RSTTRXFILE
                    If Not (.EOF And .BOF) Then
                        If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                            Newprice = IIf(IsNull(RSTTRXFILE!MRP), 0, RSTTRXFILE!MRP)
                        Else
                            If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!MRP), 0, RSTTRXFILE!MRP)
                            Else
                                Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                            End If
                        End If
                    End If
                End With
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                If Newprice = 0 Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!MRP), 0, RSTTRXFILE!MRP)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!MRP), 0, RSTTRXFILE!MRP)
                                Else
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                                End If
                            End If
                        End If
                    End With
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
                    
                If Newprice = 0 Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                                Else
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                                End If
                            End If
                        End If
                    End With
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    If Newprice = 0 Then
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
                        With RSTTRXFILE
                            If Not (.EOF And .BOF) Then
                                If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                                Else
                                    If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                        Newprice = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                                    Else
                                        Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                                    End If
                                End If
                            End If
                        End With
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    End If
                End If
            Case 2
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
                With RSTTRXFILE
                    If Not (.EOF And .BOF) Then
                        If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                            Newprice = IIf(IsNull(RSTTRXFILE!P_VAN), 0, RSTTRXFILE!P_VAN)
                        Else
                            If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!P_VAN), 0, RSTTRXFILE!P_VAN)
                            Else
                                Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                            End If
                        End If
                    End If
                End With
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                If Newprice = 0 Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!P_VAN), 0, RSTTRXFILE!P_VAN)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_VAN), 0, RSTTRXFILE!P_VAN)
                                Else
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                                End If
                            End If
                        End If
                    End With
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
            Case Else
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
                With RSTTRXFILE
                    If Not (.EOF And .BOF) Then
                        If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                            Newprice = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                        Else
                            If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                            Else
                                Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                            End If
                        End If
                    End If
                End With
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                If Newprice = 0 Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                                Else
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                                End If
                            End If
                        End If
                    End With
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
        End Select
        If Newprice = 0 Then GoTo SKIP
        
        TXTDISC.Tag = 0
        If UCase(grdsales.TextMatrix(i, 25)) = "PARDHA" Or UCase(grdsales.TextMatrix(i, 25)) = "CLOTHES" Then
            If Val(grdsales.TextMatrix(i, 6)) < 1000 Then
               grdsales.TextMatrix(i, 9) = "5"
            Else
                grdsales.TextMatrix(i, 9) = "12"
            End If
        End If
        'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
        
        If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
            grdsales.TextMatrix(grdsales.Row, 7) = Format(Round(Newprice + Newprice * Val(grdsales.TextMatrix(grdsales.Row, 9)) / 100, 4), "0.0000")
            grdsales.TextMatrix(grdsales.Row, 21) = Format(Round(Newprice + Newprice * Val(grdsales.TextMatrix(grdsales.Row, 9)) / 100, 4), "0.0000")
            grdsales.TextMatrix(i, 6) = Format(Newprice, "0.000")
        Else
            grdsales.TextMatrix(i, 6) = Format(Round(Newprice * 100 / (Val(grdsales.TextMatrix(i, 9)) + 100), 4), "0.0000")
            grdsales.TextMatrix(i, 22) = Format(Round(Newprice * 100 / (Val(grdsales.TextMatrix(i, 9)) + 100), 4), "0.0000")
            grdsales.TextMatrix(i, 7) = Format(Newprice, "0.000")
        End If
        
        If UCase(grdsales.TextMatrix(i, 25)) = "SERVICE CHARGE" Then
            TXTDISC.Tag = Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
            grdsales.TextMatrix(i, 12) = Format(Round(Val(grdsales.TextMatrix(i, 7)) - Val(TXTDISC.Tag), 4), ".0000")
            grdsales.TextMatrix(i, 34) = Format(Round(Val(grdsales.TextMatrix(i, 6)) - Val(TXTDISC.Tag), 4), ".0000")
        Else
            TXTDISC.Tag = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
            grdsales.TextMatrix(i, 12) = Format(Round((Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 7))) - Val(TXTDISC.Tag), 4), ".0000")
            grdsales.TextMatrix(i, 34) = Format(Round((Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 6))) - Val(TXTDISC.Tag), 4), ".0000")
        End If
        
        
        TXTDISC.Tag = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        db.BeginTrans
        db.Execute "Update TRXFILEVAN set SALES_PRICE = " & Val(grdsales.TextMatrix(i, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(i, 7)) & ", PTR = " & Val(grdsales.TextMatrix(i, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(i, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(i, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & ""
        db.CommitTrans
SKIP:
    Next i
    LBLTOTAL.Caption = ""
    LBLGROSSAMT.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.text = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
        Else
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
        End If
    Next i
    
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
    If Val(lblnetamount.Caption) = 0 Then
        txtcommper.text = 0
    Else
        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
    End If
    Call COSTCALCULATION
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
    Else
        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
    End If
    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"
    
                    
                    
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CMDRESET_Click()
    
    If grdsales.rows <= 1 Then Exit Sub
    If MsgBox("Are you sure you want to reset the items?", vbYesNo, "BILL..") = vbNo Then Exit Sub
    Dim RSTUNBILL As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ERRHAND
    For i = 1 To grdsales.rows - 1
        Set RSTUNBILL = New ADODB.Recordset
        RSTUNBILL.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        With RSTUNBILL
            If Not (.EOF And .BOF) Then
                grdsales.TextMatrix(i, 46) = IIf(IsNull(RSTUNBILL!UN_BILL), "N", RSTUNBILL!UN_BILL)
                db.Execute "Update TRXFILEVAN set UN_BILL = '" & Trim(grdsales.TextMatrix(i, 46)) & "' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & ""
            End If
        End With
        RSTUNBILL.Close
        Set RSTUNBILL = Nothing
    Next i
    
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub cmdreturn_Click()
    If DataList2.BoundText = "" Then Exit Sub
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    Set creditbill = Me
    Enabled = False
    M_ADD = True
    MDIMAIN.Enabled = False
    FRMCRDTNOTE.LBLCUSTOMER.Caption = DataList2.BoundText
    FRMCRDTNOTE.Tag = "Y"
    FRMCRDTNOTE.Show
End Sub

Private Sub cmdreturn_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            'If TxtName1.Enabled = True Then TxtName1.SetFocus
            'If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            'If txtretail.Enabled = True Then txtretail.SetFocus
            'If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
            'If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            'If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            'If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
            'If CMDPRINT.Enabled = True Then CMDPRINT.SetFocus
            If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
            If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
            If txtBillNo.Visible = True Then txtBillNo.SetFocus
    End Select
End Sub

Private Sub CMDSALERETURN_Click()
    Dim RSTMFGR As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ERRHAND
    'If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM SALERETURN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y'", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        grdsales.rows = grdsales.rows + 1
        grdsales.FixedRows = 1
        grdsales.TextMatrix(grdsales.rows - 1, 0) = grdsales.rows
        grdsales.TextMatrix(grdsales.rows - 1, 1) = RSTTRXFILE!ITEM_CODE
        grdsales.TextMatrix(grdsales.rows - 1, 2) = RSTTRXFILE!ITEM_NAME
        grdsales.TextMatrix(grdsales.rows - 1, 3) = RSTTRXFILE!QTY
        grdsales.TextMatrix(grdsales.rows - 1, 4) = RSTTRXFILE!UNIT
        grdsales.TextMatrix(grdsales.rows - 1, 5) = Format(RSTTRXFILE!MRP, ".000")
        grdsales.TextMatrix(grdsales.rows - 1, 6) = Format(RSTTRXFILE!PTR, ".000")
        grdsales.TextMatrix(grdsales.rows - 1, 7) = Format(RSTTRXFILE!SALES_PRICE, ".000")
        grdsales.TextMatrix(grdsales.rows - 1, 8) = Format(RSTTRXFILE!LINE_DISC, ".00")
        grdsales.TextMatrix(grdsales.rows - 1, 9) = Format(RSTTRXFILE!SALES_TAX, ".00")
        grdsales.TextMatrix(grdsales.rows - 1, 10) = ""
        grdsales.TextMatrix(grdsales.rows - 1, 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST)
        grdsales.TextMatrix(grdsales.rows - 1, 12) = Format(RSTTRXFILE!TRX_TOTAL, ".000")
        
        grdsales.TextMatrix(grdsales.rows - 1, 13) = RSTTRXFILE!ITEM_CODE
        grdsales.TextMatrix(grdsales.rows - 1, 14) = RSTTRXFILE!VCH_NO
        grdsales.TextMatrix(grdsales.rows - 1, 15) = RSTTRXFILE!LINE_NO
        grdsales.TextMatrix(grdsales.rows - 1, 16) = RSTTRXFILE!TRX_TYPE
        grdsales.TextMatrix(grdsales.rows - 1, 17) = "N"
        Set RSTMFGR = New ADODB.Recordset
        RSTMFGR.Open "SELECT MANUFACTURER  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(grdsales.rows - 1, 1)) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTMFGR.EOF And RSTMFGR.BOF) Then
            grdsales.TextMatrix(grdsales.rows - 1, 18) = IIf(IsNull(RSTMFGR!MANUFACTURER), "", Trim(RSTMFGR!MANUFACTURER))
        End If
        RSTMFGR.Close
        Set RSTMFGR = Nothing
        grdsales.TextMatrix(grdsales.rows - 1, 19) = "CN"
        grdsales.TextMatrix(grdsales.rows - 1, 20) = "0"
        
        grdsales.TextMatrix(grdsales.rows - 1, 21) = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
        grdsales.TextMatrix(grdsales.rows - 1, 22) = IIf(IsNull(RSTTRXFILE!P_RETAILWOTAX), 0, RSTTRXFILE!P_RETAILWOTAX)
        grdsales.TextMatrix(grdsales.rows - 1, 23) = IIf(IsNull(RSTTRXFILE!SALE_1_FLAG), "2", RSTTRXFILE!SALE_1_FLAG)
        grdsales.TextMatrix(grdsales.rows - 1, 24) = IIf(IsNull(RSTTRXFILE!COM_AMT), "", RSTTRXFILE!COM_AMT)
        grdsales.TextMatrix(grdsales.rows - 1, 25) = IIf(IsNull(RSTTRXFILE!Category), 0, RSTTRXFILE!Category)
        grdsales.TextMatrix(grdsales.rows - 1, 26) = IIf(IsNull(RSTTRXFILE!LOOSE_FLAG), "F", RSTTRXFILE!LOOSE_FLAG)
        grdsales.TextMatrix(grdsales.rows - 1, 27) = IIf(IsNull(RSTTRXFILE!LOOSE_PACK), "1", RSTTRXFILE!LOOSE_PACK)
        
        RSTTRXFILE!check_flag = "Y"
        RSTTRXFILE!BILL_NO = Val(txtBillNo.text)
        RSTTRXFILE!BILL_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
        CMDSALERETURN.Enabled = False
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    LBLTOTAL.Caption = ""
    LBLGROSSAMT.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))), 2)
                LBLFOT.Caption = ""
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
        End Select
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
    TXTSLNO.text = grdsales.rows
    TXTPRODUCT.text = ""
    txtPrintname.text = ""
    TxtSpec.text = ""
    txtcategory.text = ""
    TxtName1.text = ""
    TXTITEMCODE.text = ""
    optnet.Value = True
    TXTVCHNO.text = ""
    TXTLINENO.text = ""
    TXTTRXTYPE.text = ""
    TrxRYear.text = ""
    TXTUNIT.text = ""
    
    TXTQTY.text = ""
    TXTEXPIRY.text = "  /  "
    TXTAPPENDQTY.text = ""
    TXTAPPENDTOTAL.text = ""
    TXTFREEAPPEND.text = ""
    txtappendcomm.text = ""
    TxtMRP.text = ""
    txtmrpbt.text = ""
    txtretaildummy.text = ""
    lblcomtype.Caption = ""
    txtcommi.text = ""
    TxtRetailmode.text = ""
    txtretail.text = ""
    txtNetrate.text = ""
    txtBatch.text = ""
    TxtWarranty.text = ""
    TxtWarranty_type.text = ""
    TXTTAX.text = ""
    TXTRETAILNOTAX.text = ""
    TXTSALETYPE.text = ""
    TXTFREE.text = ""
    TXTDISC.text = ""
    TxtDiscAmt.text = ""
    TxtCessAmt.text = ""
    TxtCessPer.text = ""
    LBLUNBILL.Caption = ""
    cmdadd.Enabled = False
    'cmddelete.Enabled = False
    CMDEXIT.Enabled = False
    CMDPre.Caption = "Save && Pre&vious"
    cmdnxt.Caption = "Save && &Next"
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    M_EDIT = False
    'FRMEHEAD.Enabled = False
    Call COSTCALCULATION
    Call Addcommission
    TxtName1.SetFocus
    'If grdsales.Rows > 1 Then grdsales.TopRow = grdsales.Rows - 1
Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdTax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub Command2_Click()
    Dim rststock As ADODB.Recordset
    'Dim RSTMINQTY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim i As Long
    
    Chkcancel.Value = 0
    On Error GoTo ERRHAND
    'If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then Exit Sub
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
   
    'grdsales.TextMatrix(I, 17) = "N"
    LBLTOTAL.Caption = ""
    LBLGROSSAMT.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.text = ""
    
    For i = 1 To grdsales.rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT P_RETAIL  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            grdsales.TextMatrix(i, 7) = IIf(IsNull(RSTTRXFILE!P_RETAIL) Or RSTTRXFILE!P_RETAIL = 0, 100, Trim(RSTTRXFILE!P_RETAIL))
            grdsales.TextMatrix(i, 6) = Round(Val(grdsales.TextMatrix(i, 7)) * 100 / (Val(grdsales.TextMatrix(i, 9)) + 100), 3)
            grdsales.TextMatrix(i, 21) = IIf(IsNull(RSTTRXFILE!P_RETAIL) Or RSTTRXFILE!P_RETAIL = 0, 100, Trim(RSTTRXFILE!P_RETAIL))
            grdsales.TextMatrix(i, 22) = Round(Val(grdsales.TextMatrix(i, 7)) * 100 / (Val(grdsales.TextMatrix(i, 9)) + 100), 3)
            grdsales.TextMatrix(i, 12) = Round(Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 3)), 3)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing

        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "GI"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
            RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(i, 32))
        End If
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 13)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(i, 11))
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 5))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
        RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(i, 24))
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
            RSTTRXFILE!COM_FLAG = "Y"
        Else
            RSTTRXFILE!COM_FLAG = "N"
        End If
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(i, 26)
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 27))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        If Trim(TxtBillName.text) = "" Then
            RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.text), 1, 30)
        Else
            RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(TxtBillName.text), 1, 30)
        End If
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(i, 17))
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
        RSTTRXFILE!TRX_TOTAL = Round(Val(grdsales.TextMatrix(i, 3)) * RSTTRXFILE!P_RETAIL, 2)
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        If grdsales.TextMatrix(i, 38) = "" Then
            'RSTTRXFILE!EXP_DATE = Null
        Else
            RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(i, 38))) & "/" & Trim(grdsales.TextMatrix(i, 38))
        End If
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(i, 28) = "", Null, grdsales.TextMatrix(i, 28))
        RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(i, 29)
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 30)
        RSTTRXFILE!ST_RATE = Val(grdsales.TextMatrix(i, 31))
        If Trim(grdsales.TextMatrix(i, 33)) = "" Then
            RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(i, 2))
        Else
            RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(i, 33))
        End If
        Val (TXTSLNO.text)
        RSTTRXFILE.Update
    
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                            
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
        Else
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
        End If
    Next i
    
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
    If Val(lblnetamount.Caption) = 0 Then
        txtcommper.text = 0
    Else
        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
    End If
    Call cmdRefresh_Click
    txtBillNo.Visible = True
    txtBillNo.Enabled = True
    txtBillNo.SetFocus
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Command3_Click()
    Chkcancel.Value = 0
    If grdsales.rows = 1 Then Exit Sub
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    thermalprn = True
    Tax_Print = False
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    'If Month(Date) >= 5 And Year(Date) >= 2021 Then Exit Sub
'    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then
'        'db.Execute "delete From USERS "
'        Exit Sub
'    End If
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMASTVAN", db, adOpenForwardOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
'        If i > 3000 Then
'            TRXMAST.Close
'            Set TRXMAST = Nothing
'            Exit Sub
'        End If
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
'    If Not IsDate(TXTINVDATE.Text) Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    Else
'        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'    End If
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
    
    If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.text <> "" Then
        MsgBox "Select Agent From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        CMBDISTI.SetFocus
        Exit Sub
    End If
            
'    If Trim(TXTAREA.Text) = "" Then
'        MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTAREA.SetFocus
'        Exit Sub
'    End If
    
    'If Val(txtcrdays.Text) = 0 Or DataList2.BoundText = "130000" Then
    Small_Print = True
    Dos_Print = False
    Set creditbill = Me
    If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
        CMDEXIT.Enabled = False
        Me.lblcredit.Caption = "0"
        Me.Generateprint
    Else
        Me.Enabled = False
        FRMDEBITRT.Show
    End If
End Sub

Public Sub DataList2_Click()
    Dim rstCustomer As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    On Error GoTo ERRHAND
    
    If CHANGE_ADDRESS = True Then
        Set rstCustomer = New ADODB.Recordset
        rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstCustomer.EOF And rstCustomer.BOF) Then
            lbladdress.Caption = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
            'If TxtBillName.Text = "" Then TxtBillName.Text = DataList2.Text
            If Len(DataList2.text) > 11 Then
                TxtBillName.text = Mid(DataList2.text, 12)
            Else
                TxtBillName.text = DataList2.text
            End If
            TxtBillName.text = DataList2.text
            'If TxtBillAddress.Text = "" Then TxtBillAddress.Text = IIf(IsNull(rstCustomer!ADDRESS), "", Trim(rstCustomer!ADDRESS))
            TxtBillAddress.text = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
            TXTTIN.text = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
            TxtPhone.text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO)) & IIf(IsNull(rstCustomer!FAXNO) Or Trim(rstCustomer!FAXNO) = "", "", ", " & Trim(rstCustomer!FAXNO))
            TXTAREA.text = IIf(IsNull(rstCustomer!Area), "", Trim(rstCustomer!Area))
            'txtphsearch.Text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO)) & IIf(IsNull(rstCustomer!FAXNO) Or Trim(rstCustomer!FAXNO) = "", "", ", " & Trim(rstCustomer!FAXNO))
            Select Case rstCustomer!Type
                Case "W"
                    cmbtype.ListIndex = 1
                    TXTTYPE.text = 2
                Case "V"
                    cmbtype.ListIndex = 2
                    TXTTYPE.text = 3
                Case "M"
                    cmbtype.ListIndex = 3
                    TXTTYPE.text = 4
                Case Else
                    cmbtype.ListIndex = 0
                    TXTTYPE.text = 1
            End Select
            'lblcusttype.Caption = IIf((IsNull(rstCustomer!Type) Or rstCustomer!Type = ""), "R", rstCustomer!Type)
            CMBBRNCH.text = ""
            If BR_FLAG = True Then
                BR_CODE.Open "Select *  from CUSTTRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "'  ORDER BY BR_NAME", db, adOpenStatic, adLockReadOnly
                BR_FLAG = False
            Else
                BR_CODE.Close
                BR_CODE.Open "Select *  from CUSTTRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "'  ORDER BY BR_NAME", db, adOpenStatic, adLockReadOnly
                BR_FLAG = False
            End If
            Set CMBBRNCH.RowSource = BR_CODE
            CMBBRNCH.ListField = "BR_NAME"
            CMBBRNCH.BoundColumn = "BR_CODE"
        Else
            txtphsearch.text = ""
            CMBBRNCH.text = ""
            Set CMBBRNCH.RowSource = Nothing
            TxtPhone.text = ""
            TXTTIN.text = ""
            lbladdress.Caption = ""
            TXTAREA.text = ""
            TxtVehicle.text = ""
            TxtOrder.text = ""
            'lblcusttype.Caption = "R"
        End If
    End If
    
    lblsubdealer = ""
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "Select * From CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lblsubdealer = IIf(IsNull(rstCustomer!CUST_TYPE), "", rstCustomer!CUST_TYPE)
        lblIGST.Caption = IIf(IsNull(rstCustomer!CUST_IGST), "N", rstCustomer!CUST_IGST)
        lbladdress.Caption = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
    End If
    rstCustomer.Close
    Set rstCustomer = Nothing
    
    If OLD_BILL = True Then GoTo SKIP
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y' AND TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        CMDDELIVERY.Enabled = True
        Call FillDelivery_Grid
    Else
        CMDDELIVERY.Enabled = False
        grdEXPIRYLIST.FixedRows = 0
        grdEXPIRYLIST.rows = 1
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM SALERETURN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        CMDSALERETURN.Enabled = True
    Else
        CMDSALERETURN.Enabled = False
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
        
    If cr_days = False Then
        Set rstCustomer = New ADODB.Recordset
        rstCustomer.Open "Select PYMT_PERIOD, ACT_CODE From CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
        If Not (rstCustomer.EOF And rstCustomer.BOF) Then
            txtcrdays.text = IIf(IsNull(rstCustomer!PYMT_PERIOD), "", rstCustomer!PYMT_PERIOD)
        End If
        rstCustomer.Close
        Set rstCustomer = Nothing
    End If

SKIP:
    If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
        txtcrdays.Enabled = False
        CMDSALERETURN.Enabled = False
        CMDDELIVERY.Enabled = False
        Frame5.Visible = False
    Else
        txtcrdays.Enabled = True
        Frame5.Visible = True
    End If
    TXTDEALER.text = DataList2.text
    lbldealer.Caption = TXTDEALER.text
    TxtCode.text = DataList2.BoundText
    
    Dim m_Rcpt_Amt As Double
    Dim Rcptamt As Double
    Dim m_OP_Bal As Double
    Dim m_Bal_Amt As Double
    
    m_Rcpt_Amt = 0
    m_OP_Bal = 0
    m_Bal_Amt = 0
    Rcptamt = 0
    lblOB.Caption = ""
    If DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select OPEN_DB from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            m_OP_Bal = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
               
        Set RSTTRXFILE = New ADODB.Recordset
        'RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' AND NOT(TRX_TYPE= 'RT' AND INV_TRX_TYPE ='WO' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & ") ", db, adOpenStatic, adLockReadOnly, adCmdText
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE <= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE= 'RT' AND INV_TRX_TYPE ='GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        If GRDRECEIPT.rows > 1 Then Rcptamt = GRDRECEIPT.TextMatrix(0, 0)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND NOT(INV_TRX_TYPE = 'GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & ") AND (TRX_TYPE = 'DR' OR TRX_TYPE = 'DB') AND INV_DATE <= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                Case Else
                    m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        lblOB.Caption = Format(Round((m_OP_Bal + m_Bal_Amt) - (m_Rcpt_Amt), 2), "0.00")
    Else
        lblOB.Caption = "0.00"
    End If
    Exit Sub
    
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
                DataList2.SetFocus
                Exit Sub
            End If
            If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
                TxtBillName.SetFocus
            Else
                If lblsubdealer.Caption = "D" Then
                    TxtBillName.text = "Cash"
                    TxtBillName.SetFocus
                Else
                    TxtName1.Enabled = True
                    TXTPRODUCT.Enabled = True
                    TXTITEMCODE.Enabled = True
                    If CMBDISTI.VisibleCount = 0 Then
                        If TXTITEMCODE.Visible = True Then
                            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
                        Else
                            If TxtName1.Enabled = True Then TxtName1.SetFocus
                        End If
                    Else
                        CMBDISTI.SetFocus
                    End If
                End If
            End If
            'FRMEHEAD.Enabled = False
            'TxtName1.Enabled = True
            'TxtName1.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMDADD_Click()
            
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If ZERO_WARN_FLAG = True And MINUS_BILL = "N" And M_EDIT = False Then
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND CLOSE_QTY < " & Val(TXTQTY.text) + Val(TXTFREE.text) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                MsgBox "That much stock not available", , "EzBiz"
                If TXTITEMCODE.Visible = True Then
                    TXTITEMCODE.Enabled = True
                    TXTITEMCODE.SetFocus
                Else
                    TxtName1.Enabled = True
                    TxtName1.SetFocus
                End If
                .Close
                Set rststock = Nothing
                Exit Sub
            End If
            .Close
        End With
        Set rststock = Nothing
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
    
    If UCase(txtcategory.text) = "PARDHA" Or UCase(txtcategory.text) = "CLOTHES" Then
        If Val(TXTRETAILNOTAX.text) < 1000 Then
            TXTTAX.text = "5"
        Else
            TXTTAX.text = "12"
        End If
    End If
    If FrmeType.Visible = True And OptNormal.Value = False And OptLoose.Value = False Then
        MsgBox "Please select Loose / Full Option", vbOKOnly, "EzBiz"
        TXTQTY.SetFocus
        Exit Sub
    End If
    If Val(TXTQTY.text) = 0 And Val(TXTFREE.text) = 0 Then
        MsgBox "Please enter the Qty", vbOKOnly, "Sales"
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
        Exit Sub
    End If
    If MDIMAIN.LBLTAXWARN.Caption = "Y" Then
        If Val(TXTTAX.text) = 0 Then
            If (MsgBox("Tax is Zero. Are you sure?", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If Val(LBLITEMCOST.Caption) <> 0 And Val(TXTRETAILNOTAX.text) <= Val(LBLITEMCOST.Caption) Then
        If MsgBox("Price less than cost. Are you sure?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then
            txtretail.Enabled = True
            txtretail.SetFocus
            Exit Sub
        End If
    End If
    
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT FREE_WARN  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
        If RSTTRXFILE!FREE_WARN = "Y" Then MsgBox "Reminder: Free Qty available on this item", , "EzBiz"
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Call TXTQTY_LostFocus
    If Val(TXTRETAILNOTAX.text) = 0 Then
        Call TXTRETAIL_LostFocus
    End If
    
    If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
        Call TXTRETAILNOTAX_LostFocus
    End If
    
    If Val(TXTQTY.text) <> 0 And MDIMAIN.StatusBar.Panels(14).text <> "Y" And Val(TXTRETAILNOTAX.text) = 0 Then
        MsgBox "Please enter the Rate", vbOKOnly, "Sales"
        TXTRETAILNOTAX.Enabled = True
        TXTRETAILNOTAX.SetFocus
        Exit Sub
    End If
    If Val(TXTQTY.text) <> 0 And MDIMAIN.StatusBar.Panels(14).text = "Y" And Val(txtretail.text) = 0 Then
        MsgBox "Please enter the Rate", vbOKOnly, "Sales"
        txtretail.Enabled = True
        txtretail.SetFocus
        Exit Sub
    End If
    If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
        Call TXTRETAILNOTAX_LostFocus
    Else
        If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
            Call TXTRETAILNOTAX_LostFocus
        Else
            Call TXTRETAIL_LostFocus
        End If
    End If
    Call TXTDISC_LostFocus
    
    'Dim rststock As ADODB.Recordset
    'Dim RSTMINQTY As ADODB.Recordset
    'Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim i As Long
    
    
    On Error GoTo ERRHAND
    'If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then Exit Sub
    'If Month(Date) >= 5 And Year(Date) >= 2021 Then Exit Sub
'    If Month(TXTINVDATE.Text) >= 6 And Year(TXTINVDATE.Text) >= 2021 Then
'        db.Execute "delete From USERS "
'        Exit Sub
'    End If
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    'If OLD_BILL = False Then Call checklastbill
    
'    cancel_bill = False
'    If OLD_BILL = False Then
'        Dim TRXMAST As ADODB.Recordset
'        Set TRXMAST = New ADODB.Recordset
'        TRXMAST.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI'", db, adOpenForwardOnly
'        If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'            If Val(txtBillNo.Text) < IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0)) Then
'                'NEW_BILL = False
'                OLD_BILL = True
'                cancel_bill = True
'            Else
'                cancel_bill = False
'            End If
'        End If
'        TRXMAST.Close
'        Set TRXMAST = Nothing
'    End If
    
    cancel_bill = False
    If Chkcancel.Value = 1 Then
        OLD_BILL = True
        cancel_bill = True
    End If
    Chkcancel.Value = 0
    
    If UCase(TxtBillName.text) = "CANCELLED BILL" Then TxtBillName.text = "Cash"
    If OLD_BILL = True And cancel_bill = False And grdsales.rows <= 1 Then OLD_BILL = False
    
    If grdsales.rows <= 1 Then db.Execute "LOCK TABLE TRXMASTVAN Write"
'    db.Execute "SET autocommit = 0"
'    db.Execute "SET SESSION TRANSACTION ISOLATION LEVEL READ COMMITTED"
    db.Execute "START TRANSACTION"
    Set RSTTRXFILE = New ADODB.Recordset
    If OLD_BILL = False And Val(txtBillNo.text) <> 1 And grdsales.rows <= 1 Then
        RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE VCH_NO= (SELECT MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'GI')", db, adOpenStatic, adLockPessimistic, adCmdText
        txtBillNo.text = RSTTRXFILE!VCH_NO + 1
        LBLBILLNO.Caption = Val(txtBillNo.text)
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "GI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = txtBillNo.text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
    Else
        RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockPessimistic, adCmdText
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "GI"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = txtBillNo.text
            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
            RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
        End If
    End If
    RSTTRXFILE!TIN = Trim(TXTTIN.text)
    RSTTRXFILE!CUST_IGST = lblIGST.Caption
    RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
    RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    RSTTRXFILE!gross_amt = Val(LBLGROSSAMT.Caption)
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.text
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.text)
    RSTTRXFILE!DISC_PERS = Val(TXTTOTALDISC.text)
    RSTTRXFILE!BILL_NAME = Trim(TxtBillName.text)
    RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
    RSTTRXFILE!Area = Trim(TXTAREA.text)
    RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
    If Val(TxtCN.text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.text)
    RSTTRXFILE!BILL_FLAG = "Y"
    RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
    RSTTRXFILE!BR_NAME = CMBBRNCH.text
    RSTTRXFILE!cr_days = Val(txtcrdays.text)
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        RSTTRXFILE!AGENT_CODE = CMBDISTI.BoundText
        RSTTRXFILE!AGENT_NAME = CMBDISTI.text
    Else
        RSTTRXFILE!AGENT_CODE = ""
        RSTTRXFILE!AGENT_NAME = ""
    End If
    RSTTRXFILE.Update
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    'OLD_BILL = True
    db.Execute "COMMIT"
    db.Execute "UNLOCK TABLES"
    
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    If grdsales.rows <= Val(TXTSLNO.text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.text), 0) = Val(TXTSLNO.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 1) = Trim(TXTITEMCODE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 2) = Trim(TXTPRODUCT.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 3) = Val(TXTQTY.text) + Val(TXTAPPENDQTY.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 4) = Val(TXTUNIT.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 5) = Format(Val(TxtMRP.text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 6) = Format(Val(TXTRETAILNOTAX.text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 7) = Format(Val(txtretail.text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 8) = Val(TXTDISC.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 9) = Val(TXTTAX.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 10) = Trim(txtBatch.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 11) = Val(LBLITEMCOST.Caption)
    
    TXTDISC.Tag = 0
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        TXTAPPENDTOTAL.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 12))
    Else
        TXTDISC.Tag = Val(TXTAPPENDQTY.text) * Val(txtretail.text) * Val(TXTDISC.text) / 100
        TXTAPPENDTOTAL.text = Format((Val(TXTAPPENDQTY.text) * Round(Val(txtretail.text), 3)) - Val(TXTDISC.Tag), ".000")
    End If
    
    TXTAPPENDTOTAL.text = ""
    
    grdsales.TextMatrix(Val(TXTSLNO.text), 12) = Format(Val(LBLSUBTOTAL.Caption) + Val(TXTAPPENDTOTAL.text), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.text), 13) = Trim(TXTITEMCODE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 14) = Trim(TXTVCHNO.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 15) = Trim(TXTLINENO.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 16) = Trim(TXTTRXTYPE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 43) = Trim(TrxRYear.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 44) = Trim(TxtSpec.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 46) = Trim(LBLUNBILL.Caption)
    If OPTVAT.Value = True And Val(TXTTAX.text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.text), 17) = "V"
    If OPTTaxMRP.Value = True And Val(TXTTAX.text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.text), 17) = "M"
    If Val(TXTTAX.text) <= 0 Or optnet.Value = True Then grdsales.TextMatrix(Val(TXTSLNO.text), 17) = "N"
    
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "N"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 18) = IIf(IsNull(RSTTRXFILE!MANUFACTURER), "", Trim(RSTTRXFILE!MANUFACTURER))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Select Case LBLDNORCN.Caption
        Case "DN"
            grdsales.TextMatrix(Val(TXTSLNO.text), 19) = "DN"
        Case "CN"
            grdsales.TextMatrix(Val(TXTSLNO.text), 19) = "CN"
        Case Else
            grdsales.TextMatrix(Val(TXTSLNO.text), 19) = "B"
    End Select
    grdsales.TextMatrix(Val(TXTSLNO.text), 20) = Val(TXTFREE.text) + Val(TXTFREEAPPEND.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 21) = Format(Val(txtretail.text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 22) = Format(Val(TXTRETAILNOTAX.text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 23) = Trim(TXTSALETYPE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 24) = Val(txtcommi.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 25) = Trim(txtcategory.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 26) = "L"
    grdsales.TextMatrix(Val(TXTSLNO.text), 27) = IIf(Val(LblPack.text) = 0, "1", Val(LblPack.text))
    grdsales.TextMatrix(Val(TXTSLNO.text), 28) = Val(TxtWarranty.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 29) = Trim(TxtWarranty_type.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 30) = Trim(lblunit.text)
    If Val(lblvan.Caption) < Val(txtretail.text) Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 31) = Val(lblvan.Caption)
    Else
        grdsales.TextMatrix(Val(TXTSLNO.text), 31) = 0
    End If
    grdsales.TextMatrix(Val(TXTSLNO.text), 38) = IIf(TXTEXPIRY.text = "  /  ", "", Trim(TXTEXPIRY.text))
    grdsales.TextMatrix(Val(TXTSLNO.text), 39) = Val(lblretail.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.text), 40) = Val(TxtCessPer.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 41) = Val(TxtCessAmt.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 42) = Trim(lblbarcode.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.text), 45) = Val(LBLNETCOST.Tag)
    
    If Trim(txtPrintname.text) = "" Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 33) = Trim(TXTPRODUCT.text)
    Else
        grdsales.TextMatrix(Val(TXTSLNO.text), 33) = Trim(txtPrintname.text)
    End If
    grdsales.TextMatrix(Val(TXTSLNO.text), 34) = Val(LblGross.Caption)
    If M_EDIT = True Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 32) = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32))
    Else
        i = 0
        Dim rstMaxNo As ADODB.Recordset
        Set rstMaxNo = New ADODB.Recordset
        rstMaxNo.Open "Select MAX(LINE_NO) From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
        If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
            grdsales.TextMatrix(Val(TXTSLNO.text), 32) = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        Else
            grdsales.TextMatrix(Val(TXTSLNO.text), 32) = Val(TXTSLNO.text)
        End If
        rstMaxNo.Close
        Set rstMaxNo = Nothing
    End If
    
    If Not (UCase(txtcategory.text) = "SERVICES" Or UCase(txtcategory.text) = "SELF") Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 13) & "' AND TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 43)) & "' AND BAL_QTY > 0", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTQTY.text) + Val(TXTFREE.text)) * Val(LblPack.text), 3)
                !BAL_QTY = !BAL_QTY - Round((Val(TXTQTY.text) + Val(TXTFREE.text)) * Val(LblPack.text), 3)
                grdsales.TextMatrix(Val(TXTSLNO.text), 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(LblPack.text))
                RSTTRXFILE.Update
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            Else
                'BALQTY = 0
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 13) & "' AND BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    If (IsNull(RSTTRXFILE!ISSUE_QTY)) Then RSTTRXFILE!ISSUE_QTY = 0
                    If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                    'BALQTY = RSTTRXFILE!BAL_QTY
                    RSTTRXFILE!ISSUE_QTY = RSTTRXFILE!ISSUE_QTY + Round((Val(TXTQTY.text) + Val(TXTFREE.text)) * Val(LblPack.text), 3)
                    RSTTRXFILE!BAL_QTY = RSTTRXFILE!BAL_QTY - Round((Val(TXTQTY.text) + Val(TXTFREE.text)) * Val(LblPack.text), 3)
                    
                    grdsales.TextMatrix(Val(TXTSLNO.text), 14) = RSTTRXFILE!VCH_NO
                    grdsales.TextMatrix(Val(TXTSLNO.text), 15) = RSTTRXFILE!LINE_NO
                    grdsales.TextMatrix(Val(TXTSLNO.text), 16) = RSTTRXFILE!TRX_TYPE
                    grdsales.TextMatrix(Val(TXTSLNO.text), 43) = RSTTRXFILE!TRX_YEAR
                    grdsales.TextMatrix(Val(TXTSLNO.text), 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(LblPack.text))
                    RSTTRXFILE.Update
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                Else
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 13) & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        grdsales.TextMatrix(Val(TXTSLNO.text), 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(LblPack.text))
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
            End If
        End With
        
'        Dim RET_PRICE, LOOS_PRICE, LOOSE_PCK, ITEM_CST As Double
'        Set RSTTRXFILE = Nothing
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "' AND  BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockReadOnly
'        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'            If Not (IsNull(RSTTRXFILE!P_RETAIL) Or RSTTRXFILE!P_RETAIL = 0) Then
'                RET_PRICE = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
'                LOOS_PRICE = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
'                LOOSE_PCK = IIf(IsNull(RSTTRXFILE!LOOSE_PACK), 0, RSTTRXFILE!LOOSE_PACK)
'                ITEM_CST = IIf(IsNull(RSTTRXFILE!ITEM_COST), 0, RSTTRXFILE!ITEM_COST)
'            End If
'        End If
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").Value = adCriteriaKey
'                If RET_PRICE > 0 Then
'                    !P_RETAIL = RET_PRICE
'                    !P_CRTN = LOOS_PRICE
'                    !LOOSE_PACK = LOOSE_PCK
'                    If ITEM_CST > 0 Then !ITEM_COST = ITEM_CST
'                End If
                '!ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTQTY.text) * Val(LblPack.text)), 3)
                !FREE_QTY = !FREE_QTY + Round((Val(TXTFREE.text) * Val(LblPack.text)), 3)
                !CLOSE_QTY = !CLOSE_QTY - Round(((Val(TXTQTY.text) + Val(TXTFREE.text)) * Val(LblPack.text)), 3)
    
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 12))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If

    db.Execute "delete From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32)) & ""
    db.Execute "START TRANSACTION"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "GI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32))
    End If
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(Val(TXTSLNO.text), 13)
    RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(Val(TXTSLNO.text), 2)
    RSTTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3))
    RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 11))
    If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 11)) = 0 Then
        RSTTRXFILE!PUR_TAX = grdsales.TextMatrix(Val(TXTSLNO.text), 9)
    Else
        RSTTRXFILE!PUR_TAX = Val(LBLNETCOST.Tag)
    End If
    RSTTRXFILE!MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
    RSTTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6))
    RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 7))
    RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 21))
    RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 22))
    RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24))
    RSTTRXFILE!Category = grdsales.TextMatrix(Val(TXTSLNO.text), 25)
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        RSTTRXFILE!COM_FLAG = "Y"
        'RSTTRXFILE!AGENT_CODE = CMBDISTI.BoundText
        'RSTTRXFILE!AGENT_NAME = CMBDISTI.Text
    Else
        RSTTRXFILE!COM_FLAG = "N"
        'RSTTRXFILE!AGENT_CODE = ""
        'RSTTRXFILE!AGENT_NAME = ""
    End If
    RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(Val(TXTSLNO.text), 26)
    RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 27))
    RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(Val(TXTSLNO.text), 9)
    RSTTRXFILE!UNIT = grdsales.TextMatrix(Val(TXTSLNO.text), 4)
    If Trim(TxtBillName.text) = "" Then
        RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.text), 1, 30)
    Else
        RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(TxtBillName.text), 1, 30)
    End If
    RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 10))
    RSTTRXFILE!ISSUE_QTY = 0
    RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 17))
    RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 18))
    Select Case grdsales.TextMatrix(Val(TXTSLNO.text), 19)
        Case "DN"
            RSTTRXFILE!CST = 1
        Case "CN"
            RSTTRXFILE!CST = 2
        Case Else
            RSTTRXFILE!CST = 0
    End Select
    RSTTRXFILE!BAL_QTY = 0
    RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(Val(TXTSLNO.text), 12)
    RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 8))
    RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(Val(TXTSLNO.text), 7)) - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6))) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3))
    RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!M_USER_ID = DataList2.BoundText
    RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 23))
    RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(Val(TXTSLNO.text), 28) = "", 0, grdsales.TextMatrix(Val(TXTSLNO.text), 28))
    RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(Val(TXTSLNO.text), 29)
    RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(Val(TXTSLNO.text), 30)
    RSTTRXFILE!ST_RATE = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 31))
    RSTTRXFILE!RETAILER_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 39))
    RSTTRXFILE!CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 40))
    RSTTRXFILE!cess_amt = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 41))
    RSTTRXFILE!BARCODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 42))
    RSTTRXFILE!ITEM_SPEC = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 44))
    RSTTRXFILE!UN_BILL = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 46))
    If Not IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 38)) Then
        'RSTTRXFILE!EXP_DATE = Null
    Else
        RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38))) & "/" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38))
    End If
    'RSTTRXFILE!EXP_DATE = Date
    If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 33)) = "" Then
        RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 2))
    Else
        RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 33))
    End If
    RSTTRXFILE!GROSS_AMOUNT = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 34))
    RSTTRXFILE!DN_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 35))
    If IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 36)) Then
        RSTTRXFILE!DN_DATE = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 36)), Format(grdsales.TextMatrix(Val(TXTSLNO.text), 36), "DD/MM/YYYY"), Null)
    End If
    RSTTRXFILE!DN_LINENO = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37))
    
'    Dim RSTUNBILL As ADODB.Recordset
'    Set RSTUNBILL = New ADODB.Recordset
'    RSTUNBILL.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'    With RSTUNBILL
'        If Not (.EOF And .BOF) Then
'            RSTTRXFILE!UN_BILL = IIf(IsNull(!UN_BILL), "N", !UN_BILL)
'        Else
'            RSTTRXFILE!UN_BILL = "N"
'        End If
'    End With
'    RSTUNBILL.Close
'    Set RSTUNBILL = Nothing
    
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.Execute "COMMIT"
    
    db.Execute "START TRANSACTION"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXSUBVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!TRX_TYPE = "GI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32))
    End If
    RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(Val(TXTSLNO.text), 14) = "", 0, grdsales.TextMatrix(Val(TXTSLNO.text), 14))
    RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(Val(TXTSLNO.text), 15) = "", 0, grdsales.TextMatrix(Val(TXTSLNO.text), 15))
    RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(Val(TXTSLNO.text), 16) = "", "MI", grdsales.TextMatrix(Val(TXTSLNO.text), 16))
    RSTTRXFILE!R_TRX_YEAR = IIf(grdsales.TextMatrix(Val(TXTSLNO.text), 43) = "", "", grdsales.TextMatrix(Val(TXTSLNO.text), 43))
    RSTTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3))
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
        
    db.Execute "COMMIT"
    
    LBLTOTAL.Caption = ""
    LBLGROSSAMT.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.text = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
        Else
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
        End If
    Next i
    
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
    If Val(lblnetamount.Caption) = 0 Then
        txtcommper.text = 0
    Else
        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
    End If
    
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
    End If
    
    
    txtretail.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTSLNO.text = grdsales.rows
    TXTPRODUCT.text = ""
    txtPrintname.text = ""
    TxtSpec.text = ""
    txtcategory.text = ""
    If MDIMAIN.lblitemrepeat.Caption = "N" Then TxtName1.text = ""
    TXTITEMCODE.text = ""
    optnet.Value = True
    TXTVCHNO.text = ""
    TXTLINENO.text = ""
    TXTTRXTYPE.text = ""
    TrxRYear.text = ""
    TXTUNIT.text = ""
    
    lblactqty.Caption = ""
    lblbarcode.Caption = ""
    lblretail.Caption = ""
    lblwsale.Caption = ""
    lblvan.Caption = ""
    LBLMRP.Caption = ""
    lblunit.text = ""
    LblPack.text = ""
    lblOr_Pack.Caption = ""
    lblcase.Caption = ""
    lblcrtnpack.Caption = ""
    LBLITEMCOST.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLNETCOST.Caption = ""
    LBLNETPROFIT.Caption = ""
    
    LBLSELPRICE.Caption = ""
    TXTQTY.text = ""
    TXTEXPIRY.text = "  /  "
    TXTAPPENDQTY.text = ""
    TXTFREEAPPEND.text = ""
    txtappendcomm.text = ""
    TXTAPPENDTOTAL.text = ""
    TxtMRP.text = ""
    txtmrpbt.text = ""
    txtretaildummy.text = ""
    lblcomtype.Caption = ""
    txtcommi.text = ""
    TxtRetailmode.text = ""
    txtretail.text = ""
    txtNetrate.text = ""
    txtBatch.text = ""
    TXTTAX.text = ""
    TXTRETAILNOTAX.text = ""
    TXTSALETYPE.text = ""
    TXTFREE.text = ""
    TxtDiscAmt.text = ""
    TXTDISC.text = ""
    TxtCessAmt.text = ""
    TxtCessPer.text = ""
    LBLUNBILL.Caption = ""
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    TxtWarranty.text = ""
    TxtWarranty_type.text = ""
    lblP_Rate.Caption = "0"
    'cmddelete.Enabled = False
    CMDEXIT.Enabled = False
    CMDPre.Caption = "Save && Pre&vious"
    cmdnxt.Caption = "Save && &Next"
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    TxtSpec.Enabled = False
    lblunit.Enabled = False
    CMDPRINT.Enabled = True
    CmdPrintA5.Enabled = True
    cmdRefresh.Enabled = True
    
    CmdDelete.Enabled = True
    CMDMODIFY.Enabled = True
    'TxtName1.Enabled = True
    M_EDIT = False
    M_ADD = True
    'OLD_BILL = True
    Call COSTCALCULATION
    Call Addcommission
    
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
    Else
        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
    End If
    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"


    If grdsales.rows >= 9 Then grdsales.TopRow = grdsales.rows - 1
    If UCase(Trim(grdsales.TextMatrix(1, 25))) = "HOME APPLIANCES" Then
        chkTerms.Value = 1
    Else
        chkTerms.Value = 0
    End If
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    
    cmdadd.Enabled = False
    txtBatch.Enabled = False
    TXTQTY.Enabled = False
    TXTFREE.Enabled = False
    TxtMRP.Enabled = False
    TXTEXPIRY.Enabled = False
    TXTTAX.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    txtretail.Enabled = False
    TXTDISC.Enabled = False
    TxtDiscAmt.Enabled = False
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    txtretaildummy.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    TxtSpec.Enabled = False
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    If TXTITEMCODE.Visible = True Then
        If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
    Else
        If TxtName1.Enabled = True Then TxtName1.SetFocus
    End If
    'grdsales.TopRow = grdsales.Rows - 1
Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = -2147168237 Then
        'db.Execute "unlock tables"
    Else
        MsgBox err.Description
        'db.Execute "unlock tables"
    End If
    On Error Resume Next
    db.Execute "ROLLBACK"
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            If MDIMAIN.StatusBar.Panels(16).text = "Y" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTDISC.Enabled = True
                TXTDISC.SetFocus
            End If
            'TxtWarranty.Enabled = True
            'TxtWarranty.SetFocus
        Case vbKeyUp
            TXTQTY.SetFocus
    End Select

End Sub

Private Sub CmdDelete_Click()
    
    If grdsales.rows <= 1 Then Exit Sub
    If M_EDIT = True Then Exit Sub
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
        MsgBox "Permission Denied", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    TXTSLNO.text = grdsales.TextMatrix(grdsales.Row, 0)
    Call TXTSLNO_KeyDown(13, 0)
    grdsales.Enabled = True
    
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    If grdsales.TextMatrix(Val(TXTSLNO.text), 19) = "DN" Then
        Select Case MsgBox("Do you want to remove the Item " & """" & grdsales.TextMatrix(Val(TXTSLNO.text), 2) & """" & " from Delivery List", vbYesNoCancel, "DELETE.....")
            Case vbCancel
                Exit Sub
            Case vbYes
                db.Execute "DELETE  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND TRX_TYPE = 'SI' AND VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 35)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37)) & " "
                db.Execute "DELETE  From TRXFILEVAN WHERE TRX_TYPE = 'DN' AND VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 35)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37)) & " "
            Case vbNo
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND TRX_TYPE = 'SI' AND VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 35)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                With RSTTRXFILE
                    If Not (.EOF And .BOF) Then
                        RSTTRXFILE!check_flag = "N"
                        RSTTRXFILE!BILL_NO = 0 ' Null
                        'RSTTRXFILE!BILL_DATE = Null
                        RSTTRXFILE.Update
                    End If
                End With
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
        End Select
    End If
    
    db.Execute "delete From TRXSUBVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32)) & ""
    db.Execute "delete From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32)) & ""
    If Not (UCase(grdsales.TextMatrix(Val(TXTSLNO.text), 25)) = "SERVICES" Or UCase(grdsales.TextMatrix(Val(TXTSLNO.text), 25)) = "SELF") Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").Value = adCriteriaKey
                '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(LblPack.text), 3)
                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20)) * Val(LblPack.text), 3)
                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))) * Val(LblPack.text), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 12))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
           
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 43)) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))) * Val(LblPack.text), 3)
                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))) * Val(LblPack.text), 3)
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
    For i = Val(TXTSLNO.text) To grdsales.rows - 2
        grdsales.TextMatrix(i, 0) = i
        grdsales.TextMatrix(i, 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(i, 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(i, 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(i, 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(i, 6) = grdsales.TextMatrix(i + 1, 6)
        grdsales.TextMatrix(i, 5) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(i, 7) = grdsales.TextMatrix(i + 1, 7)
        grdsales.TextMatrix(i, 8) = grdsales.TextMatrix(i + 1, 8)
        grdsales.TextMatrix(i, 9) = grdsales.TextMatrix(i + 1, 9)
        grdsales.TextMatrix(i, 10) = grdsales.TextMatrix(i + 1, 10)
        grdsales.TextMatrix(i, 11) = grdsales.TextMatrix(i + 1, 11)
        grdsales.TextMatrix(i, 12) = grdsales.TextMatrix(i + 1, 12)
        grdsales.TextMatrix(i, 13) = grdsales.TextMatrix(i + 1, 13)
        grdsales.TextMatrix(i, 14) = grdsales.TextMatrix(i + 1, 14)
        grdsales.TextMatrix(i, 15) = grdsales.TextMatrix(i + 1, 15)
        grdsales.TextMatrix(i, 16) = grdsales.TextMatrix(i + 1, 16)
        grdsales.TextMatrix(i, 17) = grdsales.TextMatrix(i + 1, 17)
        grdsales.TextMatrix(i, 18) = grdsales.TextMatrix(i + 1, 18)
        grdsales.TextMatrix(i, 19) = grdsales.TextMatrix(i + 1, 19)
        grdsales.TextMatrix(i, 20) = grdsales.TextMatrix(i + 1, 20)
        grdsales.TextMatrix(i, 21) = grdsales.TextMatrix(i + 1, 21)
        grdsales.TextMatrix(i, 22) = grdsales.TextMatrix(i + 1, 22)
        grdsales.TextMatrix(i, 23) = grdsales.TextMatrix(i + 1, 23)
        grdsales.TextMatrix(i, 24) = grdsales.TextMatrix(i + 1, 24)
        grdsales.TextMatrix(i, 25) = grdsales.TextMatrix(i + 1, 25)
        grdsales.TextMatrix(i, 26) = grdsales.TextMatrix(i + 1, 26)
        grdsales.TextMatrix(i, 27) = grdsales.TextMatrix(i + 1, 27)
        grdsales.TextMatrix(i, 28) = grdsales.TextMatrix(i + 1, 28)
        grdsales.TextMatrix(i, 29) = grdsales.TextMatrix(i + 1, 29)
        grdsales.TextMatrix(i, 30) = grdsales.TextMatrix(i + 1, 30)
        grdsales.TextMatrix(i, 31) = grdsales.TextMatrix(i + 1, 31)
        grdsales.TextMatrix(i, 32) = grdsales.TextMatrix(i + 1, 32)
        grdsales.TextMatrix(i, 33) = grdsales.TextMatrix(i + 1, 33)
        grdsales.TextMatrix(i, 34) = grdsales.TextMatrix(i + 1, 34)
        grdsales.TextMatrix(i, 35) = grdsales.TextMatrix(i + 1, 35)
        grdsales.TextMatrix(i, 36) = grdsales.TextMatrix(i + 1, 36)
        grdsales.TextMatrix(i, 37) = grdsales.TextMatrix(i + 1, 37)
        grdsales.TextMatrix(i, 38) = grdsales.TextMatrix(i + 1, 38)
        grdsales.TextMatrix(i, 39) = grdsales.TextMatrix(i + 1, 39)
        grdsales.TextMatrix(i, 40) = grdsales.TextMatrix(i + 1, 40)
        grdsales.TextMatrix(i, 41) = grdsales.TextMatrix(i + 1, 41)
        grdsales.TextMatrix(i, 42) = grdsales.TextMatrix(i + 1, 42)
        grdsales.TextMatrix(i, 43) = grdsales.TextMatrix(i + 1, 43)
        grdsales.TextMatrix(i, 44) = grdsales.TextMatrix(i + 1, 44)
        grdsales.TextMatrix(i, 45) = grdsales.TextMatrix(i + 1, 45)
        grdsales.TextMatrix(i, 46) = grdsales.TextMatrix(i + 1, 46)
    Next i
    grdsales.rows = grdsales.rows - 1
    
    If grdsales.rows <= 1 Then
        OLD_BILL = False
        db.Execute "delete From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
        Dim RSTITEMMAST As ADODB.Recordset
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'GI' ", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            db.Execute "delete From TRNXRCPT WHERE TRX_TYPE='RT' AND CR_NO = " & RSTITEMMAST!CR_NO & " AND CR_TRX_TYPE = 'DR'"
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='DR' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TRX_TYPE = 'GI'"
        db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_VCH_NO = " & Val(txtBillNo.text) & " AND B_TRX_TYPE = 'GI' "
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'GI' "
        db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'GI'"

    End If
    
    LBLTOTAL.Caption = ""
    LBLGROSSAMT.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.text = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
        Else
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
        End If
    Next i
    
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
    If Val(lblnetamount.Caption) = 0 Then
        txtcommper.text = 0
    Else
        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
    End If
    
    Call COSTCALCULATION
    Call Addcommission
    
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
    Else
        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
    End If
    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"


    TXTSLNO.text = Val(grdsales.rows)
    TXTPRODUCT.text = ""
    txtPrintname.text = ""
    TxtSpec.text = ""
    txtcategory.text = ""
    TxtName1.text = ""
    TXTITEMCODE.text = ""
    optnet.Value = True
    TXTVCHNO.text = ""
    TXTLINENO.text = ""
    TXTTRXTYPE.text = ""
    TrxRYear.text = ""
    TXTUNIT.text = ""
    TXTQTY.text = ""
    TXTEXPIRY.text = "  /  "
    TXTAPPENDQTY.text = ""
    TXTFREEAPPEND.text = ""
    txtappendcomm.text = ""
    TXTAPPENDTOTAL.text = ""
    txtretail.text = ""
    txtNetrate.text = ""
    txtBatch.text = ""
    TxtWarranty.text = ""
    TxtWarranty_type.text = ""
    TXTTAX.text = ""
    TXTRETAILNOTAX.text = ""
    TXTSALETYPE.text = ""
    TXTFREE.text = ""
    TxtMRP.text = ""
    lblactqty.Caption = ""
    lblbarcode.Caption = ""
    txtmrpbt.text = ""
    txtretaildummy.text = ""
    lblcomtype.Caption = ""
    txtcommi.text = ""
    TxtRetailmode.text = ""
    
    TXTDISC.text = ""
    TxtDiscAmt.text = ""
    TxtCessAmt.text = ""
    TxtCessPer.text = ""
    LBLUNBILL.Caption = ""
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    LBLDNORCN.Caption = ""
    cmdadd.Enabled = False
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    If TXTITEMCODE.Visible = True Then
        If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
    Else
        If TxtName1.Enabled = True Then TxtName1.SetFocus
    End If
    'cmddelete.Enabled = False
    'CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    CMDPre.Caption = "Save && Pre&vious"
    cmdnxt.Caption = "Save && &Next"
    M_EDIT = False
    M_ADD = True
    If grdsales.rows = 1 Then
'        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        
        CmdPrintA5.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    If grdsales.rows >= 9 Then grdsales.TopRow = grdsales.rows - 1
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            txtPrintname.text = ""
            TxtSpec.text = ""
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            txtappendcomm.text = ""
            TXTAPPENDTOTAL.text = ""
            txtretail.text = ""
            txtNetrate.text = ""
            txtBatch.text = ""
            TxtWarranty.text = ""
            TxtWarranty_type.text = ""
            TXTTAX.text = ""
            TXTRETAILNOTAX.text = ""
            TXTSALETYPE.text = ""
            TXTFREE.text = ""
            optnet.Value = True
            TxtMRP.text = ""
            txtmrpbt.text = ""
            txtretaildummy.text = ""
            lblcomtype.Caption = ""
            txtcommi.text = ""
            TxtRetailmode.text = ""
            
            TXTDISC.text = ""
            TxtDiscAmt.text = ""
            TxtCessAmt.text = ""
            TxtCessPer.text = ""
            LBLUNBILL.Caption = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TrxRYear.text = ""
            TXTUNIT.text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            TXTQTY.Enabled = False
            
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            TXTDISC.Enabled = False
            TxtDiscAmt.Enabled = False
            txtcommi.Enabled = False
            txtretaildummy.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub CmdDelivery_Click()
    Dim RSTMFGR As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i, n As Integer
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    If grdEXPIRYLIST.rows = 1 Then Exit Sub
    If grdcount.rows = 0 Then
        MsgBox "NOTHING SELECTED!!!!", vbOKOnly, "DELETE !!!!"
        Exit Sub
    End If
    
    On Error GoTo ERRHAND
    
    cancel_bill = False
    If OLD_BILL = False Then
        Dim TRXMAST As ADODB.Recordset
        Set TRXMAST = New ADODB.Recordset
        TRXMAST.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'GI'", db, adOpenForwardOnly
        If Not (TRXMAST.EOF And TRXMAST.BOF) Then
            If Val(txtBillNo.text) < IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0)) Then
                'NEW_BILL = False
                OLD_BILL = True
                cancel_bill = True
            Else
                cancel_bill = False
            End If
        End If
        TRXMAST.Close
        Set TRXMAST = Nothing
    End If
            
    If OLD_BILL = True And cancel_bill = False And grdsales.rows <= 1 Then OLD_BILL = False
    
    Dim rstMaxNo As ADODB.Recordset
    'If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    Dim RSTUNBILL As ADODB.Recordset
    
    'If OLD_BILL = False Then Call checklastbill
    If grdsales.rows <= 1 Then db.Execute "LOCK TABLE TRXMASTVAN Write"
'    db.Execute "SET autocommit = 0"
'    db.Execute "SET SESSION TRANSACTION ISOLATION LEVEL READ COMMITTED"
    db.Execute "START TRANSACTION"
    Set RSTTRXFILE = New ADODB.Recordset
    If OLD_BILL = False And Val(txtBillNo.text) <> 1 And grdsales.rows <= 1 Then
        RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE VCH_NO= (SELECT MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'GI')", db, adOpenStatic, adLockPessimistic, adCmdText
        'db.Execute "Lock tables TRXMASTVAN write"
        txtBillNo.text = RSTTRXFILE!VCH_NO + 1
        LBLBILLNO.Caption = Val(txtBillNo.text)
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "GI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = txtBillNo.text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
    Else
        RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockPessimistic, adCmdText
'        db.Execute "Lock tables TRXMASTVAN write"
'        db.BeginTrans
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "GI"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = txtBillNo.text
            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
            RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
        End If
    End If
    RSTTRXFILE!TIN = Trim(TXTTIN.text)
    RSTTRXFILE!CUST_IGST = lblIGST.Caption
    RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
    RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    RSTTRXFILE!gross_amt = Val(LBLGROSSAMT.Caption)
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.text
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.text)
    RSTTRXFILE!DISC_PERS = Val(TXTTOTALDISC.text)
    RSTTRXFILE!BILL_NAME = Trim(TxtBillName.text)
    RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
    RSTTRXFILE!Area = Trim(TXTAREA.text)
    RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
    If Val(TxtCN.text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.text)
    RSTTRXFILE!BILL_FLAG = "Y"
    RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
    RSTTRXFILE!BR_NAME = CMBBRNCH.text
    RSTTRXFILE!cr_days = Val(txtcrdays.text)
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.Execute "COMMIT"
    db.Execute "UNLOCK TABLES"
    'OLD_BILL = True
    'db.Execute "unlock tables"
    
    For n = 0 To grdcount.rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y' AND TRX_TYPE = 'SI' AND VCH_NO = " & Val(grdcount.TextMatrix(n, 0)) & " AND LINE_NO = " & Val(grdcount.TextMatrix(n, 5)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
        'db.BeginTrans
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            grdsales.rows = grdsales.rows + 1
            grdsales.FixedRows = 1
            grdsales.TextMatrix(grdsales.rows - 1, 0) = grdsales.rows
            grdsales.TextMatrix(grdsales.rows - 1, 1) = RSTTRXFILE!ITEM_CODE
            grdsales.TextMatrix(grdsales.rows - 1, 2) = RSTTRXFILE!ITEM_NAME
            grdsales.TextMatrix(grdsales.rows - 1, 3) = RSTTRXFILE!QTY
            grdsales.TextMatrix(grdsales.rows - 1, 4) = RSTTRXFILE!UNIT
            grdsales.TextMatrix(grdsales.rows - 1, 5) = Format(RSTTRXFILE!MRP, ".000")
            grdsales.TextMatrix(grdsales.rows - 1, 6) = Format(RSTTRXFILE!PTR, ".000")
            grdsales.TextMatrix(grdsales.rows - 1, 7) = Format(RSTTRXFILE!SALES_PRICE, ".000")
            grdsales.TextMatrix(grdsales.rows - 1, 8) = Format(RSTTRXFILE!LINE_DISC, ".00")
            grdsales.TextMatrix(grdsales.rows - 1, 9) = Format(RSTTRXFILE!SALES_TAX, ".00")
            grdsales.TextMatrix(grdsales.rows - 1, 10) = IIf(IsNull(RSTTRXFILE!REF_NO), "", RSTTRXFILE!REF_NO)
            grdsales.TextMatrix(grdsales.rows - 1, 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST)
            grdsales.TextMatrix(grdsales.rows - 1, 12) = Format(RSTTRXFILE!TRX_TOTAL, ".000")
        
            grdsales.TextMatrix(grdsales.rows - 1, 13) = RSTTRXFILE!ITEM_CODE
            grdsales.TextMatrix(grdsales.rows - 1, 14) = RSTTRXFILE!R_VCH_NO
            grdsales.TextMatrix(grdsales.rows - 1, 15) = RSTTRXFILE!R_LINE_NO
            grdsales.TextMatrix(grdsales.rows - 1, 16) = RSTTRXFILE!R_TRX_TYPE
            grdsales.TextMatrix(grdsales.rows - 1, 43) = RSTTRXFILE!R_TRX_YEAR
            grdsales.TextMatrix(grdsales.rows - 1, 17) = IIf(IsNull(RSTTRXFILE!FLAG), "N", RSTTRXFILE!FLAG)
            Set RSTMFGR = New ADODB.Recordset
            RSTMFGR.Open "SELECT MANUFACTURER  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(grdsales.rows - 1, 1)) & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTMFGR.EOF And RSTMFGR.BOF) Then
                grdsales.TextMatrix(grdsales.rows - 1, 18) = Trim(RSTMFGR!MANUFACTURER)
            End If
            RSTMFGR.Close
            Set RSTMFGR = Nothing
            grdsales.TextMatrix(grdsales.rows - 1, 19) = "DN"
            grdsales.TextMatrix(grdsales.rows - 1, 20) = IIf(IsNull(RSTTRXFILE!FREE_QTY), 0, RSTTRXFILE!FREE_QTY)
            grdsales.TextMatrix(grdsales.rows - 1, 21) = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
            grdsales.TextMatrix(grdsales.rows - 1, 22) = IIf(IsNull(RSTTRXFILE!P_RETAILWOTAX), 0, RSTTRXFILE!P_RETAILWOTAX)
            grdsales.TextMatrix(grdsales.rows - 1, 23) = IIf(IsNull(RSTTRXFILE!SALE_1_FLAG), "2", RSTTRXFILE!SALE_1_FLAG)
            grdsales.TextMatrix(grdsales.rows - 1, 24) = IIf(IsNull(RSTTRXFILE!COM_AMT), "", RSTTRXFILE!COM_AMT)
            grdsales.TextMatrix(grdsales.rows - 1, 25) = IIf(IsNull(RSTTRXFILE!Category), 0, RSTTRXFILE!Category)
            grdsales.TextMatrix(grdsales.rows - 1, 26) = IIf(IsNull(RSTTRXFILE!LOOSE_FLAG), "F", RSTTRXFILE!LOOSE_FLAG)
            grdsales.TextMatrix(grdsales.rows - 1, 27) = IIf(IsNull(RSTTRXFILE!LOOSE_PACK), "1", RSTTRXFILE!LOOSE_PACK)

            Set rstMaxNo = New ADODB.Recordset
            rstMaxNo.Open "Select MAX(LINE_NO) From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
            If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
                grdsales.TextMatrix(grdsales.rows - 1, 32) = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
            Else
                grdsales.TextMatrix(grdsales.rows - 1, 32) = grdsales.rows - 1
            End If
            rstMaxNo.Close
            Set rstMaxNo = Nothing
            
            grdsales.TextMatrix(grdsales.rows - 1, 35) = IIf(IsNull(RSTTRXFILE!VCH_NO), "", RSTTRXFILE!VCH_NO)
            grdsales.TextMatrix(grdsales.rows - 1, 36) = IIf(IsNull(RSTTRXFILE!VCH_DATE), "", Format(RSTTRXFILE!VCH_DATE, "DD/MM/YYYY"))
            grdsales.TextMatrix(grdsales.rows - 1, 37) = IIf(IsNull(RSTTRXFILE!LINE_NO), "", RSTTRXFILE!LINE_NO)
            'db.CommitTrans
            RSTTRXFILE!check_flag = "Y"
            RSTTRXFILE!BILL_NO = Val(txtBillNo.text)
            RSTTRXFILE!BILL_TRX_TYPE = "GI"
            RSTTRXFILE!BILL_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
            RSTTRXFILE.Properties("Update Criteria").Value = adCriteriaKey
            RSTTRXFILE.Update
            'CMDDELIVERY.Enabled = False
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From TRXSUBVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.rows - 1, 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
            RSTTRXFILE!TRX_TYPE = "GI"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(grdsales.rows - 1, 32))
        End If
        RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(grdsales.rows - 1, 14) = "", 0, grdsales.TextMatrix(grdsales.rows - 1, 14))
        RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(grdsales.rows - 1, 15) = "", 0, grdsales.TextMatrix(grdsales.rows - 1, 15))
        RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(grdsales.rows - 1, 16) = "", "MI", grdsales.TextMatrix(grdsales.rows - 1, 16))
        RSTTRXFILE!R_TRX_YEAR = IIf(grdsales.TextMatrix(grdsales.rows - 1, 43) = "", "", grdsales.TextMatrix(grdsales.rows - 1, 43))
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(grdsales.rows - 1, 3))
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.rows - 1, 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "GI"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
            RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(grdsales.rows - 1, 32))
        End If
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(grdsales.rows - 1, 13)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(grdsales.rows - 1, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(grdsales.rows - 1, 3))
        RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(grdsales.rows - 1, 11))
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(grdsales.rows - 1, 5))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(grdsales.rows - 1, 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(grdsales.rows - 1, 7))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(grdsales.rows - 1, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(grdsales.rows - 1, 22))
        RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(grdsales.rows - 1, 24))
        RSTTRXFILE!Category = grdsales.TextMatrix(grdsales.rows - 1, 25)
        If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
            RSTTRXFILE!COM_FLAG = "Y"
        Else
            RSTTRXFILE!COM_FLAG = "N"
        End If
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(grdsales.rows - 1, 26)
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(grdsales.rows - 1, 27))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(grdsales.rows - 1, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(grdsales.rows - 1, 4)
        If Trim(TxtBillName.text) = "" Then
            RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.text), 1, 30)
        Else
            RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(TxtBillName.text), 1, 30)
        End If
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(grdsales.rows - 1, 10))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(grdsales.rows - 1, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(grdsales.rows - 1, 18))
        RSTTRXFILE!CST = 1
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(grdsales.rows - 1, 12)
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(grdsales.rows - 1, 8))
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(grdsales.rows - 1, 7)) - Val(grdsales.TextMatrix(grdsales.rows - 1, 6))) * Val(grdsales.TextMatrix(grdsales.rows - 1, 3))
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(grdsales.rows - 1, 20))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(grdsales.rows - 1, 23))
        RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(grdsales.rows - 1, 28) = "", Null, grdsales.TextMatrix(grdsales.rows - 1, 28))
        RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(grdsales.rows - 1, 29)
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(grdsales.rows - 1, 30)
        RSTTRXFILE!ST_RATE = 0 'Val(grdsales.TextMatrix(grdsales.Rows - 1, 31))
        If Trim(grdsales.TextMatrix(grdsales.rows - 1, 33)) = "" Then
            RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(grdsales.rows - 1, 2))
        Else
            RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(grdsales.rows - 1, 33))
        End If
        
        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.rows - 1, 3)) * Val(grdsales.TextMatrix(grdsales.rows - 1, 7)) * Val(grdsales.TextMatrix(grdsales.rows - 1, 8)) / 100
        RSTTRXFILE!GROSS_AMOUNT = Format(Round((Val(grdsales.TextMatrix(grdsales.rows - 1, 3)) * Val(grdsales.TextMatrix(grdsales.rows - 1, 6))) - Val(TXTDISC.Tag), 2), ".000")
        RSTTRXFILE!DN_NO = Val(grdsales.TextMatrix(grdsales.rows - 1, 35))
        RSTTRXFILE!DN_DATE = IIf(IsDate(grdsales.TextMatrix(grdsales.rows - 1, 36)), Format(grdsales.TextMatrix(grdsales.rows - 1, 36), "DD/MM/YYYY"), Null)
        RSTTRXFILE!DN_LINENO = Val(grdsales.TextMatrix(grdsales.rows - 1, 37))
        
        Set RSTUNBILL = New ADODB.Recordset
        RSTUNBILL.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(grdsales.rows - 1, 1)) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        With RSTUNBILL
            If Not (.EOF And .BOF) Then
                RSTTRXFILE!UN_BILL = IIf(IsNull(!UN_BILL), "N", !UN_BILL)
            Else
                RSTTRXFILE!UN_BILL = "N"
            End If
        End With
        RSTUNBILL.Close
        Set RSTUNBILL = Nothing
    
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    Next n
    
    LBLTOTAL.Caption = ""
    LBLGROSSAMT.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))), 2)
                LBLFOT.Caption = ""
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
        End Select
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
    
    
    Call FillDelivery_Grid
    CMDPRINT.Enabled = True
    CmdPrintA5.Enabled = True
    cmdRefresh.Enabled = True
    TXTSLNO.text = grdsales.rows
    TXTPRODUCT.text = ""
    txtPrintname.text = ""
    TxtSpec.text = ""
    txtcategory.text = ""
    TxtName1.text = ""
    TXTITEMCODE.text = ""
    optnet.Value = True
    TXTVCHNO.text = ""
    TXTLINENO.text = ""
    TXTTRXTYPE.text = ""
    TrxRYear.text = ""
    TXTUNIT.text = ""
    
    TXTQTY.text = ""
    TXTEXPIRY.text = "  /  "
    TXTAPPENDQTY.text = ""
    TXTFREEAPPEND.text = ""
    txtappendcomm.text = ""
    TXTAPPENDTOTAL.text = ""
    TxtMRP.text = ""
    lblactqty.Caption = ""
    lblbarcode.Caption = ""
    txtmrpbt.text = ""
    txtretaildummy.text = ""
    lblcomtype.Caption = ""
    txtcommi.text = ""
    TxtRetailmode.text = ""
    txtretail.text = ""
    txtNetrate.text = ""
    txtBatch.text = ""
    TxtWarranty.text = ""
    TxtWarranty_type.text = ""
    TXTTAX.text = ""
    TXTRETAILNOTAX.text = ""
    TXTSALETYPE.text = ""
    TXTFREE.text = ""
    TXTDISC.text = ""
    TxtDiscAmt.text = ""
    TxtCessAmt.text = ""
    TxtCessPer.text = ""
    LBLUNBILL.Caption = ""
    cmdadd.Enabled = False
    'cmddelete.Enabled = False
    CMDEXIT.Enabled = False
    CMDPre.Caption = "Save && Pre&vious"
    cmdnxt.Caption = "Save && &Next"
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    M_EDIT = False
    'OLD_BILL = True
    'FRMEHEAD.Enabled = False
    Call COSTCALCULATION
    Call Addcommission
    If TXTITEMCODE.Visible = True Then
        If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
    Else
        If TxtName1.Enabled = True Then TxtName1.SetFocus
    End If
    'If grdsales.Rows > 1 Then grdsales.TopRow = grdsales.Rows - 1
Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDHIDE_Click()
    If frmLogin.rs!Level <> "0" Then Exit Sub
    If LBLPROFIT.Visible = True Then
        LBLPROFIT.Visible = False
        LBLTOTALCOST.Visible = False
        Label1(25).Visible = False
        Label1(26).Visible = False
        Label1(27).Visible = False
        Label1(28).Visible = False
        Label1(44).Visible = False
        Label1(45).Visible = False
        LblProfitPerc.Visible = False
        LblProfitAmt.Visible = False
        LBLITEMCOST.Visible = False
        LBLSELPRICE.Visible = False
        LBLNETCOST.Visible = False
        LBLNETPROFIT.Visible = False
    Else
        LBLPROFIT.Visible = True
        LBLTOTALCOST.Visible = True
        Label1(25).Visible = True
        Label1(26).Visible = True
        Label1(27).Visible = True
        'Label1(28).Visible = True
        Label1(44).Visible = True
        Label1(45).Visible = True
        LblProfitPerc.Visible = True
        'LblProfitAmt.Visible = True
        'LBLITEMCOST.Visible = True
        LBLNETCOST.Visible = True
        LBLNETPROFIT.Visible = True
        'LBLSELPRICE.Visible = True
    End If
End Sub

Private Sub CMDMODIFY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    
    If grdsales.rows <= 1 Then Exit Sub
    'If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    If M_EDIT = True Then Exit Sub
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
        MsgBox "Permission Denied", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    TXTSLNO.text = grdsales.TextMatrix(grdsales.Row, 0)
    If grdsales.TextMatrix(Val(TXTSLNO.text), 19) = "DN" Then
        MsgBox "Cannot modify this. The Item is being Delivered. DN# ", vbOKOnly, "Sales"
        Exit Sub
    End If
    Call TXTSLNO_KeyDown(13, 0)
    grdsales.Enabled = True
    
    If UCase(grdsales.TextMatrix(Val(TXTSLNO.text), 25)) = "SERVICE CHARGE" Then
        CMDMODIFY.Enabled = False
        CmdDelete.Enabled = False
        CMDEXIT.Enabled = False
        CMDPre.Caption = "Save && Pre&vious"
        cmdnxt.Caption = "Save && &Next"
        M_EDIT = True
        txtretail.Enabled = True
        txtretail.SetFocus
        Exit Sub
    End If
    
    M_ADD = True
    On Error GoTo ERRHAND
    If Not (UCase(grdsales.TextMatrix(Val(TXTSLNO.text), 25)) = "SERVICES" Or UCase(grdsales.TextMatrix(Val(TXTSLNO.text), 25)) = "SELF") Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                db.BeginTrans
                .Properties("Update Criteria").Value = adCriteriaKey
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(LblPack.text), 3)
                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20)) * Val(LblPack.text), 3)
                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))) * Val(LblPack.text), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 12))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 12))
                RSTTRXFILE.Update
                db.CommitTrans
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
           
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 43)) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                db.BeginTrans
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))) * Val(LblPack.text), 3)
                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))) * Val(LblPack.text), 3)
                lblactqty.Caption = !BAL_QTY
                lblbarcode.Caption = IIf(IsNull(!BARCODE), "", !BARCODE)
                RSTTRXFILE.Update
                db.CommitTrans
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    CMDPre.Caption = "Save && Pre&vious"
    cmdnxt.Caption = "Save && &Next"
    M_EDIT = True
    TXTQTY.Enabled = True
    Call txtretail_Change
    TXTQTY.SetFocus
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            txtPrintname.text = ""
            TxtSpec.text = ""
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            txtappendcomm.text = ""
            TXTAPPENDTOTAL.text = ""
            txtretail.text = ""
            txtNetrate.text = ""
            txtBatch.text = ""
            TxtWarranty.text = ""
            TxtWarranty_type.text = ""
            TXTTAX.text = ""
            TXTRETAILNOTAX.text = ""
            TXTSALETYPE.text = ""
            TXTFREE.text = ""
            
            optnet.Value = True
            TxtMRP.text = ""
            txtmrpbt.text = ""
            txtretaildummy.text = ""
            lblcomtype.Caption = ""
            txtcommi.text = ""
            TxtRetailmode.text = ""
            
            TXTDISC.text = ""
            TxtDiscAmt.text = ""
            TxtCessAmt.text = ""
            TxtCessPer.text = ""
            LBLUNBILL.Caption = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.text = ""
            
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TrxRYear.text = ""
            TXTUNIT.text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtDiscAmt.Enabled = False
            txtcommi.Enabled = False
            txtretaildummy.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub CmdPrint_Click()
        
    Chkcancel.Value = 0
    If grdsales.rows = 1 Then Exit Sub
    
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    thermalprn = False
    Tax_Print = False
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    'If Month(Date) >= 5 And Year(Date) >= 2021 Then Exit Sub
'    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then
'        'db.Execute "delete From USERS "
'        Exit Sub
'    End If
'    If Month(TXTINVDATE.Text) >= 6 And Year(TXTINVDATE.Text) >= 2021 Then
'        db.Execute "delete From USERS "
'        Exit Sub
'    End If
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMASTVAN", db, adOpenForwardOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
'        If i > 3000 Then
'            TRXMAST.Close
'            Set TRXMAST = Nothing
'            Exit Sub
'        End If
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
'    If Not IsDate(TXTINVDATE.Text) Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    Else
'        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'    End If
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
    
    If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.text <> "" Then
        MsgBox "Select Agent From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        CMBDISTI.SetFocus
        Exit Sub
    End If
            
'    If Trim(TXTAREA.Text) = "" Then
'        MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTAREA.SetFocus
'        Exit Sub
'    End If
'
    'If Val(txtcrdays.Text) = 0 Or DataList2.BoundText = "130000" Then
    Small_Print = False
    Dos_Print = False
    Chkcancel.Value = 0
    Set creditbill = Me
    If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
        CMDEXIT.Enabled = False
        Me.lblcredit.Caption = "0"
        Me.Generateprint
    Else
        Me.Enabled = False
        FRMDEBITRT.Show
    End If
    
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    txtOutstanding.text = ""
    Dim m_Rcpt_Amt As Double
    Dim Rcptamt As Double
    Dim m_OP_Bal As Double
    Dim m_Bal_Amt As Double
    
    m_Rcpt_Amt = 0
    m_OP_Bal = 0
    m_Bal_Amt = 0
    Rcptamt = 0
    'Call Print_A4
    'Exit Function
    'If thermalprn = False And (MDIMAIN.StatusBar.Panels(8).Text = "Y" Or MDIMAIN.lblprint.Caption = "Y") Then
    'If Val(MDIMAIN.StatusBar.Panels(10).Text) = 0 Then
    If thermalprn = False And ((Small_Print = False And (MDIMAIN.StatusBar.Panels(8).text = "Y" Or MDIMAIN.lblprint.Caption = "Y")) Or (Small_Print = True And MDIMAIN.lbldmpmini.Caption = "Y")) Then
        GoTo SKIP
        If DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "select OPEN_DB from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                m_OP_Bal = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
                   
            Set RSTTRXFILE = New ADODB.Recordset
            'RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' AND NOT(TRX_TYPE= 'RT' AND INV_TRX_TYPE ='GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & ") ", db, adOpenStatic, adLockReadOnly, adCmdText
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                Select Case RSTTRXFILE!TRX_TYPE
                    Case "DB"
                        m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                    Case Else
                        m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                End Select
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE= 'RT' AND INV_TRX_TYPE ='GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                Select Case RSTTRXFILE!TRX_TYPE
                    Case "DB"
                        m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                    Case Else
                        m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                End Select
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            If GRDRECEIPT.rows > 1 Then Rcptamt = GRDRECEIPT.TextMatrix(0, 0)
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND NOT(INV_TRX_TYPE = 'GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & ") AND (TRX_TYPE = 'DR' OR TRX_TYPE = 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                Select Case RSTTRXFILE!TRX_TYPE
                    Case "DB"
                        m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                    Case Else
                        m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                End Select
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            txtOutstanding.text = Round((m_OP_Bal + m_Bal_Amt) - (m_Rcpt_Amt), 0)
        End If
SKIP:
        'Call ReportGeneratION_vpestimate
        LBLFOT.Tag = ""
        If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = True Then
            If MsgBox("You do not have any permission to modify this further. Are you sure to print?", vbYesNo, "BILL..") = vbNo Then Exit Function
        Else
            Screen.MousePointer = vbHourglass
            Sleep (300)
        End If
        NEW_BILL = False
        
        If MDIMAIN.lbldmpmini = "Y" Then
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                Call ReportGeneratION_Composition
            Else
                Call ReportGeneratION
            End If
        Else
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                Call ReportGeneratION_Composition
            Else
                Call ReportGeneratION_vpestimate(0, 0)
            End If
        End If
        
        On Error GoTo CLOSEFILE
        Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
        If err.Number = 55 Then
            Close #1
            Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
        End If
        On Error GoTo ERRHAND
        
        'Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
        Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
        Print #1, "EXIT"
        Close #1
        
        '//HERE write the proper path where your command.com file exist
        Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & Rptpath & "REPO.BAT N", vbHide
        ST_PRINT = False
        'Call cmdRefresh_Click
        cmdRefresh.SetFocus
    Else
        If thermalprn = True And MDIMAIN.LBLDMPTHERMAL.Caption = "Y" Then
            LBLFOT.Tag = ""
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = True Then
                If MsgBox("You do not have any permission to modify this further. Are you sure to print?", vbYesNo, "BILL..") = vbNo Then Exit Function
            Else
                Screen.MousePointer = vbHourglass
                Sleep (300)
            End If
            NEW_BILL = False
            
            Call ReportGeneratION_WO
            
            On Error GoTo CLOSEFILE1
            Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE1:
            If err.Number = 55 Then
                Close #1
                Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
            End If
            On Error GoTo ERRHAND
            
            Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
            Print #1, "EXIT"
            Close #1
            
            '//HERE write the proper path where your command.com file exist
            Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & Rptpath & "REPO.BAT N", vbHide
            ST_PRINT = False
            'Call cmdRefresh_Click
            cmdRefresh.SetFocus
        Else
            Call Print_A4
        End If
    End If
    Screen.MousePointer = vbNormal
    Exit Function
ERRHAND:
    ST_PRINT = False
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Function
Private Sub CMDPRINT_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            'TXTPRODUCT.Text = ""
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            txtappendcomm.text = ""
            TXTAPPENDTOTAL.text = ""
            txtretail.text = ""
            txtNetrate.text = ""
            txtBatch.text = ""
            TxtWarranty.text = ""
            TxtWarranty_type.text = ""
            TXTTAX.text = ""
            TXTRETAILNOTAX.text = ""
            TXTSALETYPE.text = ""
            TXTFREE.text = ""
            
            optnet.Value = True
            TxtMRP.text = ""
            lblactqty.Caption = ""
            lblbarcode.Caption = ""
            txtmrpbt.text = ""
            txtretaildummy.text = ""
            lblcomtype.Caption = ""
            txtcommi.text = ""
            TxtRetailmode.text = ""
            
            TXTDISC.text = ""
            TxtDiscAmt.text = ""
            TxtCessAmt.text = ""
            TxtCessPer.text = ""
            LBLUNBILL.Caption = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TrxRYear.text = ""
            TXTUNIT.text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtDiscAmt.Enabled = False
            txtcommi.Enabled = False
            txtretaildummy.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub cmdRefresh_Click()
    
   ' If grdsales.Rows = 1 Then GoTo SKIP
    
'    If Not IsDate(TXTINVDATE.Text) Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
'        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTINVDATE.SetFocus
'        Exit Sub
'    Else
'        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'    End If
    
    On Error GoTo ERRHAND
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
    
    On Error GoTo ERRHAND
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenForwardOnly
    If RSTCOMPANY.RecordCount <> grdsales.rows - 1 Then
        MsgBox "There is an error. Please do verify it ", , "EzBiz"
'        RSTCOMPANY.Close
'        Set RSTCOMPANY = Nothing
'        Exit Sub
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    If txtBillNo.Tag = "Y" Then
        Call AppendSale
        If preclick = True Then Exit Sub
        lbldealer.Caption = ""
        flagchange.Caption = ""
        TXTDEALER.text = ""
        LBLRETAMT.Caption = ""
        TXTTIN.text = ""
        TxtCN.text = ""
        TXTTOTALDISC.text = ""
        txtcommper.text = ""
        LBLTOTALCOST.Caption = ""
        lblOB.Caption = ""
        Chkcancel.Value = 0
        txtcrdays.text = ""
        
        lblOB.Caption = ""
        LBLDATE.Caption = Date
        TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
        lbltime.Caption = Format(Time, "hh:mm AMPM")
        OLD_BILL = False
        LBLBILLNO.Caption = Val(txtBillNo.text)
        
        Call txtBillNo_KeyDown(13, 0)
        
        TXTDEALER.text = "CASH"
        DataList2.BoundText = "130000"
        'cmdreturn.Enabled = True
        TXTITEMCODE.Enabled = True
        TxtName1.Enabled = True
        TXTITEMCODE.Enabled = True
        'TXTITEMCODE.SetFocus
        TXTDEALER.SetFocus
    
        Exit Sub
    End If
    
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
    
    If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.text <> "" Then
        MsgBox "Select Agent From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        CMBDISTI.SetFocus
        Exit Sub
    End If
            
'    If Trim(TXTAREA.Text) = "" Then
'        MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTAREA.SetFocus
'        Exit Sub
'    End If
    
    If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then Me.lblcredit.Caption = "0"
        
    Set creditbill = Me
    If preclick = True Then
        Call AppendSale
    Else
        If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
            Call AppendSale
        Else
            Me.Enabled = False
            FRMDEBITRT.LBLTYPE.Caption = "S"
            FRMDEBITRT.Show
        End If
    End If
    
'    Call AppendSale
'    If preclick = True Then Exit Sub
'
'    LBLBILLNO.Caption = Val(txtBillNo.Text)
'    lbldealer.Caption = ""
'    flagchange.Caption = ""
'    TXTDEALER.Text = ""
'    LBLRETAMT.Caption = ""
'    TXTTIN.Text = ""
'    TxtCN.Text = ""
'    TXTTOTALDISC.Text = ""
'    txtcommper.Text = ""
'    LBLTOTALCOST.Caption = ""
'    lblOB.Caption = ""
'    Chkcancel.value = 0
'    txtcrdays.Text = ""
'    LBLDATE.Caption = Date
'    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
'    lbltime.Caption = Format(Time, "hh:mm AMPM")
'    OLD_BILL = False
'    Call TXTBILLNO_KeyDown(13, 0)
'
'    TXTDEALER.Text = "CASH"
'    DataList2.BoundText = "130000"
'    'cmdreturn.Enabled = True
'    TXTITEMCODE.Enabled = True
'    'TXTITEMCODE.SetFocus
'    TXTDEALER.SetFocus
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
    
    'Me.Enabled = False
    'FRMDEBITRT.Show
    
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            'TXTPRODUCT.Text = ""
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            txtappendcomm.text = ""
            TXTAPPENDTOTAL.text = ""
            txtretail.text = ""
            txtNetrate.text = ""
            txtBatch.text = ""
            TxtWarranty.text = ""
            TxtWarranty_type.text = ""
            TXTTAX.text = ""
            TXTRETAILNOTAX.text = ""
            TXTSALETYPE.text = ""
            TXTFREE.text = ""
            
            optnet.Value = True
            TxtMRP.text = ""
            lblactqty.Caption = ""
            lblbarcode.Caption = ""
            txtmrpbt.text = ""
            txtretaildummy.text = ""
            lblcomtype.Caption = ""
            txtcommi.text = ""
            TxtRetailmode.text = ""
            
            TXTDISC.text = ""
            TxtDiscAmt.text = ""
            TxtCessAmt.text = ""
            TxtCessPer.text = ""
            LBLUNBILL.Caption = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TrxRYear.text = ""
            TXTUNIT.text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            TXTDISC.Enabled = False
            TxtDiscAmt.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
    TxtCode.text = DataList2.BoundText
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    DataList2.BackColor = &H98F3C1
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
     If DataList2.BoundText = "" Then
        cmdreturn.Enabled = False
    Else
        cmdreturn.Enabled = True
    End If
    DataList2.BackColor = vbWhite
End Sub

Private Sub Form_Activate()
    If txtBillNo.Visible = True Then txtBillNo.SetFocus
    'If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
    'If txtretail.Enabled = True Then txtretail.SetFocus
    'If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
    'If TXTTAX.Enabled = True Then TXTTAX.SetFocus
    'If TXTDISC.Enabled = True Then TXTDISC.SetFocus
    'If txtcommi.Enabled = True Then txtcommi.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
    If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
    'If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
    'If  Then CMDDOS.SetFocus
    'If TxtName1.Enabled = True Then TxtName1.SetFocus
    
    If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
    If TxtBillName.Enabled = True Then TxtBillName.SetFocus
    If OLD_BILL = True Then
        If TXTITEMCODE.Visible = True Then
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
        Else
            If TxtName1.Enabled = True Then TxtName1.SetFocus
        End If
    Else
        If TXTDEALER.Enabled = True Then TXTDEALER.SetFocus
    End If
    'If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
    If MDIMAIN.lblub.Caption = "N" Then
        CmdReset.Visible = False
        grdsales.ColWidth(46) = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyF1
                ST_PRINT = True
                Call Generateprint
                ST_PRINT = False
            Case 37
                Call CMDPre_Click
            Case 39
                Call cmdnxt_Click
        End Select
    End If
    If Shift = vbAltMask Then
        Select Case KeyCode
            Case 97, 49
                TXTTYPE.text = 1
                cmbtype.ListIndex = 0
            Case 98, 50
                TXTTYPE.text = 2
                cmbtype.ListIndex = 1
            Case 99, 51
                TXTTYPE.text = 3
                cmbtype.ListIndex = 2
            Case 100, 52
                TXTTYPE.text = 4
                cmbtype.ListIndex = 3
        End Select
    End If
    Select Case KeyCode
        Case vbKeyF2
            On Error Resume Next
            grdsales.SetFocus
        Case vbKeyF3
            On Error Resume Next
            Txtrcvd.SetFocus
        Case vbKeyF5
            If D_PRINT = 1 Then
                If CMDPRINT.Enabled = True Then Call CmdPrint_Click
            ElseIf D_PRINT = 2 Then
                If Command3.Enabled = True Then Call Command3_Click
            Else
                If CmdPrintA5.Enabled = True Then Call CmdPrintA5_Click
            End If
        Case vbKeyF6
            On Error Resume Next
            If cmdRefresh.Enabled = True Then Call cmdRefresh_Click
    End Select
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ERRHAND
            
    If frmLogin.rs!Level = "5" Then
        TxtMRP.Locked = True
        TXTTAX.Locked = True
        txtretail.Locked = True
        TXTRETAILNOTAX.Locked = True
        txtNetrate.Locked = True
        TXTDISC.Locked = True
        TxtDiscAmt.Locked = True
        TxtCessPer.Locked = True
        TxtCessAmt.Locked = True
        TXTTOTALDISC.Locked = True
        
        TxtFrieght.Locked = True
        Txthandle.Locked = True
        CMDHIDE.Visible = False
        TXTTYPE.Locked = True
        cmbtype.Locked = True
        
        CMDSALERETURN.Visible = False
        CMDDELIVERY.Visible = False
    End If
    
    If frmLogin.rs!Level <> "0" Then
        cmdratechange.Visible = False
        'CmdAddRate.Visible = False
    End If
        
    If MDIMAIN.lblExpEnable.Caption = "Y" Then CmdExport.Visible = True
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'GI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = Val(txtBillNo.text)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        lblpre.Caption = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
        lblsuf.Caption = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        If RSTCOMPANY!mrp_plus = "Y" Then
            mrpplus = True
        Else
            mrpplus = False
        End If
        If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
            Label1(43).Visible = False
            TXTITEMCODE.Visible = False
            Label1(9).Left = 520
            Label1(9).Width = Val(Label1(9).Width) + 1530
            TxtName1.Left = 520
            TxtName1.Width = 2600
        End If
        If RSTCOMPANY!hide_deliver = "Y" Then
            grdEXPIRYLIST.Visible = False
            CHKSELECT.Visible = False
        End If
        If RSTCOMPANY!hide_terms = "Y" Then
            chkTerms.Visible = False
            Terms1.Visible = False
        End If
        If RSTCOMPANY!hide_wrnty = "Y" Then
            TxtWarranty.Visible = False
            TxtWarranty_type.Visible = False
            Label1(48).Visible = False
            If RSTCOMPANY!hide_pr_name <> "Y" Then
                txtPrintname.Left = 30
                Label1(38).Left = 30
                txtPrintname.Width = 5025
                Label1(38).Width = 5025
            Else
                If RSTCOMPANY!hide_spec <> "Y" Then
                    Label1(59).Top = 840
                    TxtSpec.Top = 1065
                End If
            End If
        End If
        If RSTCOMPANY!hide_spec = "Y" Then
            TxtSpec.Visible = False
            Label1(59).Visible = False
        End If
        If RSTCOMPANY!hide_pr_name = "Y" Then
            txtPrintname.Visible = False
            Label1(38).Visible = False
        End If
        If RSTCOMPANY!hide_serial = "Y" Then
            Label1(7).Visible = False
            txtBatch.Visible = False
            Label1(16).Left = 8325
            lblunit.Left = 8325
            Label1(37).Left = 7830
            LblPack.Left = 7830
            'Label1(9).Width = 5760
            Label1(9).Width = Val(Label1(9).Width) + 1000
            TXTPRODUCT.Width = 4680
        End If
        If RSTCOMPANY!hide_free = "Y" Then
            Label1(58).Visible = False
            TXTFREE.Visible = False
            Label1(10).Width = 1210
            TXTQTY.Width = 1210
'            Label1(24).Width = 1050
'            TxtMRP.Width = 1050
'            Label1(24).Left = 9705
'            TxtMRP.Left = 9705
        End If
        
        If RSTCOMPANY!hide_mrp = "Y" And RSTCOMPANY!hide_expiry = "Y" Then
            Label1(24).Visible = False
            TxtMRP.Visible = False
            Label1(61).Visible = False
            TXTEXPIRY.Visible = False
            Label1(12).Left = Val(Label1(12).Left) - 1720
            TXTTAX.Left = Val(TXTTAX.Left) - 1720
            Lblprice(30).Left = Val(Lblprice(30).Left) - 1720
            TXTRETAILNOTAX.Left = Val(TXTRETAILNOTAX.Left) - 1720
            Label1(11).Left = Val(Label1(11).Left) - 1720
            txtretail.Left = Val(txtretail.Left) - 1720
            Label1(13).Left = Val(Label1(13).Left) - 1720
            TXTDISC.Left = Val(TXTDISC.Left) - 1720
            Label1(65).Left = Val(Label1(65).Left) - 1720
            TxtDiscAmt.Left = Val(TxtDiscAmt.Left) - 1720
            Label1(29).Left = Val(Label1(29).Left) - 1720
            TxtCessPer.Left = Val(TxtCessPer.Left) - 1720
            Label1(62).Left = Val(Label1(62).Left) - 1720
            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 1720
            Label1(14).Left = Val(Label1(14).Left) - 1720
            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 1720
        ElseIf RSTCOMPANY!hide_mrp <> "Y" And RSTCOMPANY!hide_expiry = "Y" Then
            Label1(61).Visible = False
            TXTEXPIRY.Visible = False
'            Label1(24).Left = Val(Label1(24).Left) - 1300
'            TxtMRP.Left = Val(TXTTAX.Left) - 1300
            Label1(12).Left = Val(Label1(12).Left) - 910
            TXTTAX.Left = Val(TXTTAX.Left) - 910
            Lblprice(30).Left = Val(Lblprice(30).Left) - 910
            TXTRETAILNOTAX.Left = Val(TXTRETAILNOTAX.Left) - 910
            Label1(11).Left = Val(Label1(11).Left) - 910
            txtretail.Left = Val(txtretail.Left) - 910
            Label1(13).Left = Val(Label1(13).Left) - 910
            TXTDISC.Left = Val(TXTDISC.Left) - 910
            Label1(65).Left = Val(Label1(65).Left) - 910
            TxtDiscAmt.Left = Val(TxtDiscAmt.Left) - 910
            Label1(29).Left = Val(Label1(29).Left) - 910
            TxtCessPer.Left = Val(TxtCessPer.Left) - 910
            Label1(62).Left = Val(Label1(62).Left) - 910
            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 910
            Label1(14).Left = Val(Label1(14).Left) - 910
            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 910
        ElseIf RSTCOMPANY!hide_mrp = "Y" And RSTCOMPANY!hide_expiry <> "Y" Then
            Label1(24).Visible = False
            TxtMRP.Visible = False
            Label1(61).Left = Val(Label1(61).Left) - 810
            TXTEXPIRY.Left = Val(TXTEXPIRY.Left) - 810
            Label1(12).Left = Val(Label1(12).Left) - 810
            TXTTAX.Left = Val(TXTTAX.Left) - 810
            Lblprice(30).Left = Val(Lblprice(30).Left) - 810
            TXTRETAILNOTAX.Left = Val(TXTRETAILNOTAX.Left) - 810
            Label1(11).Left = Val(Label1(11).Left) - 810
            txtretail.Left = Val(txtretail.Left) - 810
            Label1(13).Left = Val(Label1(13).Left) - 810
            TXTDISC.Left = Val(TXTDISC.Left) - 810
            Label1(65).Left = Val(Label1(65).Left) - 810
            TxtDiscAmt.Left = Val(TxtDiscAmt.Left) - 810
            Label1(29).Left = Val(Label1(29).Left) - 810
            TxtCessPer.Left = Val(TxtCessPer.Left) - 810
            Label1(62).Left = Val(Label1(62).Left) - 810
            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 810
            Label1(14).Left = Val(Label1(14).Left) - 810
            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 810

        End If
        
        If RSTCOMPANY!hide_disc = "Y" Then
            Label1(65).Visible = False
            TxtDiscAmt.Visible = False
            Label1(29).Left = Val(Label1(29).Left) - 870
            TxtCessPer.Left = Val(TxtCessPer.Left) - 870
            Label1(62).Left = Val(Label1(62).Left) - 870
            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 870
            Label1(14).Left = Val(Label1(14).Left) - 870
            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 870
        End If
        
'        If RSTCOMPANY!hide_mrp = "Y" Then
'            Label1(24).Visible = False
'            TxtMRP.Visible = False
'            Label1(61).Left = 9960
'            TXTEXPIRY.Left = 9960
'            Label1(12).Left = Val(Label1(12).Left) - 810
'            TXTTAX.Left = Val(TXTTAX.Left) - 810
'            Lblprice(30).Left = Val(Lblprice(30).Left) - 810
'            TXTRETAILNOTAX.Left = Val(TXTRETAILNOTAX.Left) - 810
'            Label1(11).Left = Val(Label1(11).Left) - 810
'            txtretail.Left = Val(txtretail.Left) - 810
'            Label1(13).Left = Val(Label1(13).Left) - 810
'            TXTDISC.Left = Val(TXTDISC.Left) - 810
'            Label1(65).Left = Val(Label1(65).Left) - 810
'            TxtDiscAmt.Left = Val(TxtDiscAmt.Left) - 810
'            Label1(29).Left = Val(Label1(29).Left) - 810
'            TxtCessPer.Left = Val(TxtCessPer.Left) - 810
'            Label1(62).Left = Val(Label1(62).Left) - 810
'            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 810
'            Label1(14).Left = Val(Label1(14).Left) - 810
'            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 810
'
'            If RSTCOMPANY!hide_free = "Y" Then
'                Label1(10).Width = 1210
'                TXTQTY.Width = 1210
'            End If
'        End If
'
'        If RSTCOMPANY!hide_expiry = "Y" Then
'            If RSTCOMPANY!hide_mrp <> "Y" Then
''                Label1(10).Width = Val(Label1(10).Width) - 500
''                TXTQTY.Width = Val(TXTQTY.Width) - 500
'                Label1(24).Left = Val(Label1(24).Left) + 100
'                TxtMRP.Left = Val(TxtMRP.Left) + 100
'            End If
'            Label1(61).Visible = False
'            TXTEXPIRY.Visible = False
'            Label1(12).Left = Val(Label1(12).Left) - 810
'            TXTTAX.Left = Val(TXTTAX.Left) - 810
'            Lblprice(30).Left = Val(Lblprice(30).Left) - 850
'            TXTRETAILNOTAX.Left = Val(TXTRETAILNOTAX.Left) - 850
'            Label1(11).Left = Val(Label1(11).Left) - 850
'            txtretail.Left = Val(txtretail.Left) - 850
'            Label1(13).Left = Val(Label1(13).Left) - 850
'            TXTDISC.Left = Val(TXTDISC.Left) - 850
'            Label1(65).Left = Val(Label1(65).Left) - 850
'            TxtDiscAmt.Left = Val(TxtDiscAmt.Left) - 850
'            Label1(29).Left = Val(Label1(29).Left) - 850
'            TxtCessPer.Left = Val(TxtCessPer.Left) - 850
'            Label1(62).Left = Val(Label1(62).Left) - 850
'            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 850
'            Label1(14).Left = Val(Label1(14).Left) - 850
'            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 850
'
'            If RSTCOMPANY!hide_free = "Y" Then
'                If RSTCOMPANY!hide_mrp <> "Y" Then
'                    Label1(10).Width = 1060
'                    TXTQTY.Width = 1060
'                Else
'                    Label1(10).Width = 1210
'                    TXTQTY.Width = 1210
'                End If
'            Else
'                Label1(58).Width = Val(Label1(58).Width) + 100
'                TXTFREE.Width = Val(TXTFREE.Width) + 100
'            End If
'        End If
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
'    If Val(txtBillNo.Text) > 20 Then
'        Open "C:\WINDOWS\system32\mwp.lp1" For Output As #1 '//Report file Creation
'        Print #1, ""
'        Close #1
'        Exit Sub
'    End If
    
    TXTAREA.Clear
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select DISTINCT AREA From CUSTMAST ORDER BY AREA", db, adOpenForwardOnly
    Do Until rstBILL.EOF
        If Not IsNull(rstBILL!Area) Then TXTAREA.AddItem (rstBILL!Area)
        rstBILL.MoveNext
    Loop
    rstBILL.Close
    Set rstBILL = Nothing
        
    OLD_BILL = False
    BR_FLAG = True
    NEW_BILL = True
    SERIAL_FLAG = False
    lblactqty.Caption = ""
    lblbarcode.Caption = ""
    ACT_FLAG = True
    AGNT_FLAG = True
    M_ADD = False
    lblcredit.Caption = "1"
    txtcrdays.text = ""
    lblP_Rate.Caption = "0"
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    lbltime.Caption = Format(Time, "hh:mm AMPM")
    grdsales.ColWidth(0) = 550
    grdsales.ColWidth(1) = 1500
    grdsales.ColWidth(2) = 4000
    grdsales.ColWidth(3) = 900
    grdsales.ColWidth(4) = 0
    grdsales.ColWidth(5) = 1100
    grdsales.ColWidth(6) = 1300
    grdsales.ColWidth(7) = 1300
    grdsales.ColWidth(8) = 700
    grdsales.ColWidth(9) = 700
    grdsales.ColWidth(10) = 0
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 1600
    grdsales.ColWidth(13) = 0
    grdsales.ColWidth(14) = 0
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(16) = 0
    grdsales.ColWidth(17) = 0
    grdsales.ColWidth(18) = 0
    grdsales.ColWidth(19) = 0
    grdsales.ColWidth(20) = 500
    grdsales.ColWidth(21) = 0
    grdsales.ColWidth(22) = 0
    grdsales.ColWidth(23) = 0
    grdsales.ColWidth(24) = 800
    grdsales.ColWidth(25) = 0
    grdsales.ColWidth(26) = 0
    grdsales.ColWidth(27) = 0
    grdsales.ColWidth(28) = 0
    grdsales.ColWidth(29) = 0
    grdsales.ColWidth(30) = 0
    grdsales.ColWidth(31) = 0
    grdsales.ColWidth(32) = 0
    grdsales.ColWidth(33) = 0
    grdsales.ColWidth(34) = 0
    grdsales.ColWidth(35) = 0
    grdsales.ColWidth(36) = 0
    grdsales.ColWidth(37) = 0
    grdsales.ColWidth(38) = 0 '100
    grdsales.ColWidth(39) = 0
    grdsales.ColWidth(40) = 0
    grdsales.ColWidth(41) = 0
    grdsales.ColWidth(42) = 0
    grdsales.ColWidth(43) = 0
    grdsales.ColWidth(44) = 0
    grdsales.ColWidth(45) = 0
    grdsales.ColWidth(46) = 400
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "Item Code"
    grdsales.TextArray(2) = "Item Description"
    grdsales.TextArray(3) = "Qty"
    grdsales.TextArray(4) = "UOM"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "Rate"
    grdsales.TextArray(7) = "Net Rate"
    grdsales.TextArray(8) = "Disc%"
    grdsales.TextArray(9) = "Tax%"
    grdsales.TextArray(10) = "Serial No"
    grdsales.TextArray(11) = "Cost"
    grdsales.TextArray(12) = "Sub Total"
    grdsales.TextArray(13) = "Item Code"
    grdsales.TextArray(14) = "Vch No"
    grdsales.TextArray(15) = "Line No"
    grdsales.TextArray(16) = "Trx Type"
    grdsales.TextArray(17) = "Tax Mode"
    grdsales.TextArray(18) = "MFGR"
    grdsales.TextArray(19) = "" '"CN/DN"
    grdsales.TextArray(20) = "Free"
    grdsales.TextArray(21) = "PTR"
    grdsales.TextArray(22) = "PTRWOTAX"
    grdsales.TextArray(24) = "Com"
    grdsales.TextArray(31) = "" '"Code"
    grdsales.TextArray(33) = "Print Name"
    grdsales.TextArray(34) = "Gross"
    grdsales.TextArray(38) = "" '"Expiry"
    'grdsales.ColWidth(12) = 0
    'grdsales.ColWidth(13) = 0
    'grdsales.ColWidth(14) = 0
   'grdsales.ColWidth(15) = 0
    'grdsales.ColWidth(16) = 0
    
    grdsales.ColAlignment(0) = 4
    grdsales.ColAlignment(2) = 1
    grdsales.ColAlignment(3) = 4
    grdsales.ColAlignment(5) = 7
    grdsales.ColAlignment(6) = 7
    grdsales.ColAlignment(7) = 7
    grdsales.ColAlignment(8) = 4
    grdsales.ColAlignment(12) = 7
    grdsales.ColAlignment(20) = 4
    
    If frmLogin.rs!Level <> "0" Then
'        Label1(21).Visible = False
'        lblretail.Visible = False
        Label1(53).Visible = False
        txtcommper.Visible = False
        Label2.Visible = False
        lblcomamt.Visible = False
        Label1(46).Visible = False
        txtretaildummy.Visible = False
        lblcomtype.Visible = False
        'grdsales.ColWidth(31) = 0
    Else
        'grdsales.ColWidth(31) = 1100
        Label1(21).Visible = True
        lblretail.Visible = True
    End If
    
    LBLTOTAL.Caption = 0
    lblcomamt.text = 0
    LBLRETAMT.Caption = 0
    
    PHYFLAG = True
    PHYCODEFLAG = True
    TMPFLAG = True
    BATCH_FLAG = True
    ITEM_FLAG = True
    PRERATE_FLAG = True
    cr_days = False
    TxtName1.Enabled = False
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    TXTQTY.Enabled = False
    
    TxtMRP.Enabled = False
    
    txtretail.Enabled = False
    txtcommi.Enabled = False
    txtretaildummy.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    TXTDISC.Enabled = False
    TxtDiscAmt.Enabled = False
    
    'cmddelete.Enabled = False
    'CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    
    CmdPrintA5.Enabled = False
    
    TXTSLNO.text = 1
    Call fillcombo
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    CLOSEALL = 1
    TxtCN.text = ""
    M_EDIT = False
    
    TXTSLNO.text = grdsales.rows
    txtBillNo.Visible = False
    TXTDEALER.text = "CASH"
    TxtBillName.text = "Cash"
    DataList2.BoundText = "130000"
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    cmbtype.ListIndex = 0
    TXTTYPE.text = 1
'    Me.Width = 11700
'    Me.Height = 10185
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If PHYCODEFLAG = False Then PHYCODE.Close
        If TMPFLAG = False Then TMPREC.Close
        If BATCH_FLAG = False Then PHY_BATCH.Close
        If ITEM_FLAG = False Then PHY_ITEM.Close
        If PRERATE_FLAG = False Then PHY_PRERATE.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If AGNT_FLAG = False Then ACT_AGNT.Close
        If BR_FLAG = False Then BR_CODE.Close
    End If
    Cancel = CLOSEALL
    
End Sub


Private Sub GRDPOPUP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTtax As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            SERIAL_FLAG = True
            lblactqty.Caption = GRDPOPUP.Columns(1)
            lblbarcode.Caption = GRDPOPUP.Columns(24)
            txtBatch.text = GRDPOPUP.Columns(0)
            TXTVCHNO.text = GRDPOPUP.Columns(2)
            TXTLINENO.text = GRDPOPUP.Columns(3)
            TXTTRXTYPE.text = GRDPOPUP.Columns(4)
            TrxRYear.text = IIf(IsNull(GRDPOPUP.Columns(28)), "", GRDPOPUP.Columns(28))
            TxtMRP.text = IIf(IsNull(GRDPOPUP.Columns(21)), "", GRDPOPUP.Columns(21))
            TXTEXPIRY.text = IIf(IsDate(GRDPOPUP.Columns(25)), Format(GRDPOPUP.Columns(25), "MM/YY"), "  /  ")
            Select Case PHY!check_flag
                Case "M"
                    OPTTaxMRP.Value = True
                    TXTTAX.text = PHY!SALES_TAX
                Case "V"
                    OPTVAT.Value = True
                    TXTSALETYPE.text = "2"
                    TXTTAX.text = PHY!SALES_TAX
                Case Else
                    TXTSALETYPE.text = "2"
                    optnet.Value = True
                    TXTTAX.text = "0"
            End Select
            
            'TXTUNIT.Text = GRDPOPUP.Columns(5)
            
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TxtName1.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
            
            Call CONTINUE_BATCH
            TxtWarranty.text = GRDPOPUP.Columns(7)
            TxtWarranty_type.text = GRDPOPUP.Columns(8)
            Set GRDPOPUP.DataSource = Nothing
            Exit Sub
        Case vbKeyEscape
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            TXTAPPENDTOTAL.text = ""
            txtappendcomm.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TrxRYear.text = ""
            TXTUNIT.text = ""
            
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
    Dim RSTtax As ADODB.Recordset
    Dim NONSTOCKFLAG As Boolean
    Dim MINUSFLAG As Boolean
    Dim i As Long
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            SERIAL_FLAG = False
            lblactqty.Caption = ""
            lblbarcode.Caption = ""
            NONSTOCKFLAG = False
            MINUSFLAG = False
            M_STOCK = Val(GRDPOPUPITEM.Columns(2))
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            txtcommi.text = ""
            TXTPRODUCT.text = GRDPOPUPITEM.Columns(1)
            txtPrintname.text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.text = GRDPOPUPITEM.Columns(0)
            TxtMRP.text = IIf(IsNull(GRDPOPUPITEM.Columns(20)), "", GRDPOPUPITEM.Columns(20))
            txtcategory.text = IIf(IsNull(GRDPOPUPITEM.Columns(7)), "", GRDPOPUPITEM.Columns(7))
            If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TXTQTY.text = 1
                txtretail.Enabled = True
                txtretail.SetFocus
                Exit Sub
            End If
            i = 0
            If M_STOCK <= 0 Then
                MsgBox "AVAILABLE STOCK IS  " & i & " ", , "SALES"
                TXTQTY.SelStart = 0
                TXTQTY.SelLength = Len(TXTQTY.text)
                Exit Sub
                If SERIAL_FLAG = True Then
                    MsgBox "AVAILABLE STOCK IS  " & M_STOCK & " ", , "SALES"
                    Exit Sub
                End If
                    
                If (MsgBox("AVAILABLE STOCK IS  " & M_STOCK & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                    Exit Sub
                Else
                    MINUSFLAG = True
                End If
                NONSTOCKFLAG = True
            End If
            For i = 1 To grdsales.rows - 1
                If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.text) Then
                    If MsgBox("This Item Already exists.... Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then
                        Set GRDPOPUPITEM.DataSource = Nothing
                        FRMEITEM.Visible = False
                        FRMEMAIN.Enabled = True
                        TXTPRODUCT.Enabled = True
                        TXTQTY.Enabled = False
                        
                        TXTPRODUCT.SetFocus
                        Exit Sub
                    Else
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN", "DN"
                                Exit For
                        End Select
'                        If SERIAL_FLAG = False Then
'                            TXTSLNO.Text = i
'                            TXTAPPENDQTY.Text = Val(grdsales.TextMatrix(i, 3))
'                            TXTFREEAPPEND.Text = Val(grdsales.TextMatrix(i, 20))
'                            txtappendcomm.Text = Val(grdsales.TextMatrix(i, 24))
'                            Exit For
'                        End If
                    End If
                End If
            Next i
            
            Set GRDPOPUPITEM.DataSource = Nothing
            If ITEM_FLAG = True Then
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE, CATEGORY, WARRANTY, WARRANTY_TYPE, MRP, P_LWS  From ITEMMASTVAN  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE, CATEGORY, WARRANTY, WARRANTY_TYPE, MRP, P_LWS  From ITEMMASTVAN  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            End If
            Set GRDPOPUPITEM.DataSource = PHY_ITEM
            'GRDPOPUPITEM.RowHeight = 350
            If PHY_ITEM.RecordCount = 0 Then
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TxtName1.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTITEMCODE.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            End If
            
            If MDIMAIN.lblPriceSplit.Caption = "Y" Then GoTo SKIP_BATCH5
            Dim RSTBATCH As ADODB.Recordset
            Set RSTBATCH = New ADODB.Recordset
            Select Case cmbtype.ListIndex
                Case 1
                    'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND(P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                    'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                Case 2
                    'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_RETAIL >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                    'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                Case 3
                    'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND(P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                    'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    RSTBATCH.Open "Select DISTINCT ITEM_CODE, MRP, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                Case Else
                    'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_VAN >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                    'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
            End Select
            If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                If BATCH_DISPLAY = True Then
                    If RSTBATCH.RecordCount >= 1 Then
                        Call FILL_BATCHGRID
                        RSTBATCH.Close
                        Set RSTBATCH = Nothing
                        Exit Sub
                    ElseIf RSTBATCH.RecordCount = 1 Then
                        'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                        txtBatch.text = IIf(IsNull(RSTBATCH!REF_NO), "", RSTBATCH!REF_NO)
                    End If
                Else
                    If RSTBATCH.RecordCount > 1 Then
                        Call FILL_BATCHGRID
                        RSTBATCH.Close
                        Set RSTBATCH = Nothing
                        Exit Sub
                    ElseIf RSTBATCH.RecordCount = 1 Then
                        'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                        txtBatch.text = IIf(IsNull(RSTBATCH!REF_NO), "", RSTBATCH!REF_NO)
                    End If
                End If
            End If
            RSTBATCH.Close
            Set RSTBATCH = Nothing
SKIP_BATCH5:
                'TXTQTY.Text = GRDPOPUPITEM.Columns(2)
'            Select Case cmbtype.ListIndex
'                Case 0 'VP
'                    'txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", GRDPOPUPITEM.Columns(13))
'                    'kannattu
'                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", GRDPOPUPITEM.Columns(13))
'                Case 1 'RT
'                    'txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", GRDPOPUPITEM.Columns(6))
'                    'kannattu
'                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", GRDPOPUPITEM.Columns(6))
'                Case 2 'WS
'                    'txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", GRDPOPUPITEM.Columns(12))
'                    'kannattu
'                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", GRDPOPUPITEM.Columns(12))
'            End Select
            LblPack.text = IIf(IsNull(GRDPOPUPITEM.Columns(15)) Or Val(GRDPOPUPITEM.Columns(15)) = 0, "1", GRDPOPUPITEM.Columns(15))
            lblOr_Pack.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(15)) Or Val(GRDPOPUPITEM.Columns(15)) = 0, "1", GRDPOPUPITEM.Columns(15))
            'txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)) * Val(LblPack.Text))
            
            Select Case cmbtype.ListIndex
                Case 0
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                Case 1
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)))
                Case 2
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", Val(GRDPOPUPITEM.Columns(13)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", Val(GRDPOPUPITEM.Columns(13)))
                Case 3
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(20)), "", Val(GRDPOPUPITEM.Columns(20)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(20)), "", Val(GRDPOPUPITEM.Columns(20)))
                    If Val(txtretail.text) = 0 Then
                        txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                    End If
            End Select
            If Val(TxtCessPer.text) <> 0 Or Val(TxtCessAmt.text) <> 0 Then
                TXTRETAILNOTAX.text = (Val(txtretail.text) - Val(TxtCessAmt.text)) / (1 + (Val(TXTTAX.text) / 100) + (Val(TxtCessPer.text) / 100))
                txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 3)
                TXTRETAILNOTAX.text = Val(txtretail.text)
            End If
                
            lblretail.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", GRDPOPUPITEM.Columns(6))
            lblwsale.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", GRDPOPUPITEM.Columns(12))
            lblvan.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", GRDPOPUPITEM.Columns(13))
            LBLMRP.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(20)), "", GRDPOPUPITEM.Columns(20))
            lblcase.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(11)), "", GRDPOPUPITEM.Columns(11))
            lblLWPrice.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(21)), "", GRDPOPUPITEM.Columns(21))
            lblcrtnpack.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(10)), "", GRDPOPUPITEM.Columns(10))
            lblunit.text = IIf(IsNull(GRDPOPUPITEM.Columns(16)), "Nos", GRDPOPUPITEM.Columns(16))
            TxtWarranty.text = IIf(IsNull(GRDPOPUPITEM.Columns(18)), "", GRDPOPUPITEM.Columns(18))
            TxtWarranty_type.text = IIf(IsNull(GRDPOPUPITEM.Columns(19)), "", GRDPOPUPITEM.Columns(19))
        
            LblPack.text = IIf(IsNull(GRDPOPUPITEM.Columns(10)), "", GRDPOPUPITEM.Columns(10))
            If Val(LblPack.text) = 0 Then LblPack.text = "1"
            txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(11)), "", GRDPOPUPITEM.Columns(11))
            
            If GRDPOPUPITEM.Columns(7) = "A" Then
                txtretaildummy.text = IIf(IsNull(GRDPOPUPITEM.Columns(9)), "P", GRDPOPUPITEM.Columns(9))
                lblcomtype.Caption = "`"
                TxtRetailmode.text = "A"
                txtcommi.text = Format(Round(Val(txtretaildummy.text), 2), "0.00")
            Else
                txtretaildummy.text = IIf(IsNull(GRDPOPUPITEM.Columns(8)), "P", GRDPOPUPITEM.Columns(8))
                lblcomtype.Caption = "%"
                TxtRetailmode.text = "P"
                txtcommi.text = Format(Round((Val(TXTRETAILNOTAX.text) * Val(txtretaildummy.text) / 100), 2), "0.00")
            End If
            If Val(txtretaildummy.text) = 0 Then lblcomtype.Caption = "`"
            Select Case PHY_ITEM!check_flag
                Case "M"
                    OPTTaxMRP.Value = True
                    TXTTAX.text = GRDPOPUPITEM.Columns(4)
                    TXTSALETYPE.text = "2"
                Case "V"
                    OPTVAT.Value = True
                    TXTSALETYPE.text = "2"
                    TXTTAX.text = GRDPOPUPITEM.Columns(4)
                Case Else
                    TXTSALETYPE.text = "2"
                    optnet.Value = True
                    TXTTAX.text = "0"
            End Select
            
'            OPTVAT.value = True
'            TXTTAX.Text = "14.5"
'            TXTSALETYPE.Text = "2"
            
'            optnet.Value = True
            TXTUNIT.text = GRDPOPUPITEM.Columns(5)
                        
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TxtName1.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = True
            
            TXTQTY.SetFocus
        Case vbKeyEscape
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            txtappendcomm.text = ""
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TrxRYear.text = ""
            TXTUNIT.text = ""
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            
            TXTPRODUCT.SetFocus
            
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub GRDPRERATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            FRMEMAIN.Enabled = True
            TXTRETAILNOTAX.Enabled = True
            TXTRETAILNOTAX.SetFocus
    End Select
End Sub

Private Sub grdsales_DblClick()
    If MDIMAIN.lblub.Caption = "N" Then Exit Sub
    If grdsales.rows <= 1 Then Exit Sub
    If grdsales.Col <> 46 Then Exit Sub
    If grdsales.TextMatrix(grdsales.Row, 46) = "Y" Then
        If MsgBox("Are you sure you want to mark this item as Billed Item?", vbYesNo + vbDefaultButton2, "BILL..") = vbNo Then Exit Sub
        grdsales.TextMatrix(grdsales.Row, 46) = "N"
    Else
        If MsgBox("Are you sure you want to mark this item as UN-Billed Item?", vbYesNo + vbDefaultButton2, "BILL..") = vbNo Then Exit Sub
        grdsales.TextMatrix(grdsales.Row, 46) = "Y"
    End If
    On Error GoTo ERRHAND
    db.Execute "Update TRXFILEVAN set UN_BILL = '" & Trim(grdsales.TextMatrix(grdsales.Row, 46)) & "' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            On Error Resume Next
            'TXTPRODUCT.Text = grdtmp.Columns(1)
            'TXTITEMCODE.Text = grdtmp.Columns(0)
            SERIAL_FLAG = False
            lblactqty.Caption = ""
            lblbarcode.Caption = ""
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyReturn
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
                MsgBox "Permission Denied", vbOKOnly, "Sales"
                Exit Sub
            End If
'            If MDIMAIN.LBLGSTWRN.Caption = "Y" And Trim(TXTTIN.Text) = "" Then
'                MsgBox "B2C Bill Not allowed", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            On Error Resume Next
            TXTITEMCODE.text = grdtmp.Columns(0)
            Dim i As Integer
            For i = 1 To grdsales.rows - 1
                If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.text) Then
                    If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next i
            TXTPRODUCT.text = grdtmp.Columns(1)
            txtPrintname.text = grdtmp.Columns(1)
            TxtCessPer.text = IIf(IsNull(grdtmp.Columns(24)), "", grdtmp.Columns(24))
            TxtCessAmt.text = IIf(IsNull(grdtmp.Columns(25)), "", grdtmp.Columns(25))
            TXTDISC.text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
            Call TxtItemcode_KeyDown(13, 0)
            
            Set grdtmp.DataSource = Nothing
            grdtmp.Visible = False
            If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                TXTQTY.text = 1
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTQTY.Enabled = True
                
                TXTQTY.SetFocus
            End If
    End Select
End Sub

Private Sub lblcomamt_LostFocus()
    Call CHANGEBOXCOLOR(lblcomamt, False)
End Sub

Private Sub LblPack_GotFocus()
    LblPack.SelStart = 0
    LblPack.SelLength = Len(LblPack.text)
    Call CHANGEBOXCOLOR(LblPack, True)
End Sub

Private Sub LblPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(LblPack.text) = 0 Then Exit Sub
            LblPack.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.text = ""
            TXTLINENO.text = ""
            TXTTRXTYPE.text = ""
            TrxRYear.text = ""
            TXTUNIT.text = ""
            TXTPRODUCT.Enabled = True
            TxtName1.Enabled = True
            TXTITEMCODE.Enabled = True
            LblPack.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub LblPack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub LblPack_LostFocus()
    On Error Resume Next
    If Val(lblcrtnpack.Caption) = 0 Then lblcrtnpack.Caption = 1
    If Val(LblPack.text) = 0 Then LblPack.text = 1
    If Val(LblPack.text) = 1 Then
        TxtMRP.text = Val(LBLMRP.Caption) / Val(lblOr_Pack.Caption)
    Else
        TxtMRP.text = Val(LBLMRP.Caption) '/ Val(LblPack.Text)
    End If
    If Val(LblPack.text) <> Val(lblOr_Pack.Caption) Then
        'txtretail.Text = Val(lblcase.Caption) * Val(LblPack.Text)
        'TxtMRP.Text = Val(LBLMRP.Caption) / Val(LblPack.Text)
        If cmbtype.ListIndex = 1 Then
            txtretail.text = (Val(lblLWPrice.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(lblLWPrice.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        ElseIf cmbtype.ListIndex = 2 Then
            txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        ElseIf cmbtype.ListIndex = 3 Then
            txtretail.text = (Val(LBLMRP.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(LBLMRP.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        Else
            txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        End If
    Else
        'TxtMRP.Text = Val(LBLMRP.Caption)
        If cmbtype.ListIndex = 1 Then
            txtretail.text = Val(lblwsale.Caption)
            TXTRETAILNOTAX.text = Val(lblwsale.Caption)
        ElseIf cmbtype.ListIndex = 2 Then
            txtretail.text = Val(lblvan.Caption)
            TXTRETAILNOTAX.text = Val(lblvan.Caption)
        ElseIf cmbtype.ListIndex = 3 Then
            txtretail.text = Val(LBLMRP.Caption)
            TXTRETAILNOTAX.text = Val(LBLMRP.Caption)
        Else
            txtretail.text = Val(lblretail.Caption)
            TXTRETAILNOTAX.text = Val(lblretail.Caption)
        End If
    End If
    
    If Val(TxtCessPer.text) <> 0 Or Val(TxtCessAmt.text) <> 0 Then
        TXTRETAILNOTAX.text = (Val(txtretail.text) - Val(TxtCessAmt.text)) / (1 + (Val(TXTTAX.text) / 100) + (Val(TxtCessPer.text) / 100))
        txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 3)
        TXTRETAILNOTAX.text = Val(txtretail.text)
    End If
    
    If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
        Call TXTRETAILNOTAX_LostFocus
    Else
        If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
            Call TXTRETAILNOTAX_LostFocus
        Else
            Call TXTRETAIL_LostFocus
        End If
    End If
    Call CHANGEBOXCOLOR(LblPack, False)
End Sub

Private Sub lblunit_GotFocus()
    Call CHANGEBOXCOLOR(lblunit, True)
End Sub

Private Sub lblunit_LostFocus()
    Call CHANGEBOXCOLOR(lblunit, False)
End Sub

Private Sub optnet_Click()
    'TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OPTTaxMRP_Click()
    'TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OPTVAT_Click()
    'TXTRETAILNOTAX_LostFocus
End Sub

Private Sub Terms1_GotFocus()
    Call CHANGEBOXCOLOR(Terms1, True)
End Sub

Private Sub Terms1_LostFocus()
    Call CHANGEBOXCOLOR(Terms1, False)
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.text)
    Call CHANGEBOXCOLOR(txtBatch, True)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            If M_EDIT = True Then
                'If MsgBox("THIS WILL REMOVE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
                'Call REMOVE_ITEM
                Exit Sub
            End If
            LblPack.Enabled = True
            LblPack.SetFocus
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

Private Sub TxtBillAddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
                DataList2.SetFocus
                Exit Sub
            End If
            If Trim(TxtBillName.text) = "" Then TxtBillName.text = TXTDEALER.text
'                MsgBox "Enter Customer Name", vbOKOnly, "Sale Bil..."
'                TxtBillName.SetFocus
'                Exit Sub
'            End If
'            FRMEHEAD.Enabled = False
            'TxtPhone.Enabled = True
            'TxtPhone.SetFocus
            
            TXTTYPE.Enabled = True
            TXTTYPE.SetFocus
        Case vbKeyEscape
            TxtBillName.Enabled = True
            TxtBillName.SetFocus
    End Select
End Sub

Private Sub TxtBillAddress_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtBillAddress_LostFocus()
    TxtBillAddress.BackColor = vbWhite
End Sub

Private Sub TxtBillName_LostFocus()
    Call CHANGEBOXCOLOR(TxtBillName, False)
End Sub

Private Sub txtBillNo_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.text)
    cr_days = False
    TXTTOTALDISC.text = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
'    MDIMAIN.MNUENTRY.Visible = False
'    MDIMAIN.MNUREPORT.Visible = False
'    MDIMAIN.mnugud_rep.Visible = False
'    MDIMAIN.MNUTOOLS.Visible = False
End Sub

Public Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TRXMAST As ADODB.Recordset
    Dim TRXSUB As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    
    Dim i As Long
    Dim n As Integer
    Dim M As Integer

    On Error GoTo ERRHAND
    DataList2.text = TXTDEALER.text
    Call DataList2_Click

    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtBillNo.Text) = 0 Then Exit Sub
            'If Val(txtBillNo.Text) > 800 Then Exit Sub
            LBLRETAMT.Caption = ""
            TxtCN.text = ""
            OLD_BILL = False
            lblbalance.Caption = ""
            Txtrcvd.text = ""
            txtBillNo.Tag = "N"
'            Set TRXMAST = New ADODB.Recordset
'            TRXMAST.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND (ISNULL(BILL_FLAG) OR BILL_FLAG <>'Y') ", db, adOpenStatic, adLockReadOnly
'            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
'                txtBillNo.Tag = "Y"
'            Else
'                txtBillNo.Tag = "N"
'            End If
'            TRXMAST.Close
'            Set TRXMAST = Nothing
           
            grdsales.rows = 1
            i = 0
            Set TRXFILE = New ADODB.Recordset
            TRXFILE.Open "Select * From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until TRXFILE.EOF
                OLD_BILL = True
                i = i + 1
                TXTINVDATE.text = IIf(IsDate(TRXFILE!VCH_DATE), Format(TRXFILE!VCH_DATE, "DD/MM/YYYY"), "  /  /    ")
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = TRXFILE!ITEM_CODE
                grdsales.TextMatrix(i, 2) = TRXFILE!ITEM_NAME
                grdsales.TextMatrix(i, 3) = TRXFILE!QTY
                
'                Set TRXMAST = New ADODB.Recordset
'                TRXMAST.Open "SELECT UNIT From RTRXFILEVAN WHERE TRX_TYPE = '" & Trim(TRXSUB!R_TRX_TYPE) & "' AND VCH_NO = " & Val(TRXSUB!R_VCH_NO) & " AND LINE_NO = " & Val(TRXSUB!R_LINE_NO) & " AND TRX_YEAR = '" & Val(TRXSUB!R_TRX_YEAR) & "'", db, adOpenStatic, adLockReadOnly
'                If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
'                    grdsales.TextMatrix(i, 4) = Val(TRXMAST!UNIT)
'                End If
'                TRXMAST.Close
'                Set TRXMAST = Nothing
                
                Set TRXMAST = New ADODB.Recordset
                TRXMAST.Open "SELECT MANUFACTURER From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(TRXFILE!ITEM_CODE) & "'", db, adOpenStatic, adLockReadOnly
                If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                    grdsales.TextMatrix(i, 18) = IIf(IsNull(TRXMAST!MANUFACTURER), "", Trim(TRXMAST!MANUFACTURER))
                End If
                TRXMAST.Close
                Set TRXMAST = Nothing
                
                grdsales.TextMatrix(i, 4) = 1
                grdsales.TextMatrix(i, 5) = Format(TRXFILE!MRP, ".000")
                grdsales.TextMatrix(i, 6) = Format(TRXFILE!PTR, ".0000")
                grdsales.TextMatrix(i, 7) = Format(TRXFILE!SALES_PRICE, ".0000")
                grdsales.TextMatrix(i, 8) = IIf(IsNull(TRXFILE!LINE_DISC), 0, TRXFILE!LINE_DISC) 'DISC
                grdsales.TextMatrix(i, 9) = Val(TRXFILE!SALES_TAX)
        
                grdsales.TextMatrix(i, 10) = IIf(IsNull(TRXFILE!REF_NO), "", TRXFILE!REF_NO) 'SERIAL
                grdsales.TextMatrix(i, 11) = IIf(IsNull(TRXFILE!ITEM_COST), 0, TRXFILE!ITEM_COST)
                grdsales.TextMatrix(i, 12) = Format(Val(TRXFILE!TRX_TOTAL), ".000")
                
                grdsales.TextMatrix(i, 13) = TRXFILE!ITEM_CODE
                
                Set TRXSUB = New ADODB.Recordset
                TRXSUB.Open "Select * From TRXSUBVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(TRXFILE!LINE_NO) & "", db, adOpenStatic, adLockReadOnly
                If Not (TRXSUB.EOF Or TRXSUB.BOF) Then
                    grdsales.TextMatrix(i, 14) = Val(TRXSUB!R_VCH_NO)
                    grdsales.TextMatrix(i, 15) = Val(TRXSUB!R_LINE_NO)
                    grdsales.TextMatrix(i, 16) = Trim(TRXSUB!R_TRX_TYPE)
                    grdsales.TextMatrix(i, 43) = IIf(IsNull(TRXSUB!R_TRX_YEAR), "", Trim(TRXSUB!R_TRX_YEAR))
                End If
                TRXSUB.Close
                Set TRXSUB = Nothing
                
'                grdsales.TextMatrix(i, 14) = Val(TRXSUB!R_VCH_NO)
'                grdsales.TextMatrix(i, 15) = Val(TRXSUB!R_LINE_NO)
'                grdsales.TextMatrix(i, 16) = Trim(TRXSUB!R_TRX_TYPE)
'                grdsales.TextMatrix(i, 43) = IIf(IsNull(TRXSUB!R_TRX_YEAR), "", Trim(TRXSUB!R_TRX_YEAR))
                grdsales.TextMatrix(i, 17) = IIf(IsNull(TRXFILE!check_flag), "", Trim(TRXFILE!check_flag))
                TXTDEALER.text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                'TxtCode.Text = TRXFILE!ACT_CODE
                'DataList2.Text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                LBLDATE.Caption = IIf(IsNull(TRXFILE!CREATE_DATE), Date, TRXFILE!CREATE_DATE)
                Select Case TRXFILE!CST
                    Case 0
                        grdsales.TextMatrix(i, 19) = "B"
                    Case 1
                        grdsales.TextMatrix(i, 19) = "DN"
                    Case 2
                        grdsales.TextMatrix(i, 19) = "CN"
                End Select
                grdsales.TextMatrix(i, 20) = TRXFILE!FREE_QTY
                grdsales.TextMatrix(i, 21) = IIf(IsNull(TRXFILE!P_RETAIL), "0.00", Format(TRXFILE!P_RETAIL, ".0000"))
                grdsales.TextMatrix(i, 22) = IIf(IsNull(TRXFILE!P_RETAILWOTAX), "0.00", Format(TRXFILE!P_RETAILWOTAX, ".0000"))
                grdsales.TextMatrix(i, 23) = IIf(IsNull(TRXFILE!SALE_1_FLAG), "2", TRXFILE!SALE_1_FLAG)
                grdsales.TextMatrix(i, 24) = IIf(IsNull(TRXFILE!COM_AMT), "", TRXFILE!COM_AMT)
                grdsales.TextMatrix(i, 25) = IIf(IsNull(TRXFILE!Category), "", TRXFILE!Category)
                grdsales.TextMatrix(i, 26) = IIf(IsNull(TRXFILE!LOOSE_FLAG), "F", TRXFILE!LOOSE_FLAG)
                grdsales.TextMatrix(i, 27) = IIf(IsNull(TRXFILE!LOOSE_PACK), "1", TRXFILE!LOOSE_PACK)
                grdsales.TextMatrix(i, 28) = IIf(IsNull(TRXFILE!WARRANTY), "", TRXFILE!WARRANTY)
                grdsales.TextMatrix(i, 29) = IIf(IsNull(TRXFILE!WARRANTY_TYPE), "", TRXFILE!WARRANTY_TYPE)
                grdsales.TextMatrix(i, 30) = IIf(IsNull(TRXFILE!PACK_TYPE), "Nos", TRXFILE!PACK_TYPE)
                grdsales.TextMatrix(i, 31) = IIf(IsNull(TRXFILE!ST_RATE), 0, TRXFILE!ST_RATE)
                grdsales.TextMatrix(i, 32) = TRXFILE!LINE_NO
                grdsales.TextMatrix(i, 33) = IIf(IsNull(TRXFILE!PRINT_NAME), grdsales.TextMatrix(i, 2), TRXFILE!PRINT_NAME)
                grdsales.TextMatrix(i, 34) = IIf(IsNull(TRXFILE!GROSS_AMOUNT), 0, TRXFILE!GROSS_AMOUNT)
                grdsales.TextMatrix(i, 35) = IIf(IsNull(TRXFILE!DN_NO), "", TRXFILE!DN_NO)
                grdsales.TextMatrix(i, 36) = IIf(IsNull(TRXFILE!DN_DATE), "", Format(TRXFILE!DN_DATE, "DD/MM/YYYY"))
                grdsales.TextMatrix(i, 37) = IIf(IsNull(TRXFILE!DN_LINENO), "", TRXFILE!DN_LINENO)
                grdsales.TextMatrix(i, 38) = IIf(IsDate(TRXFILE!EXP_DATE), Format(TRXFILE!EXP_DATE, "MM/YY"), "")
                grdsales.TextMatrix(i, 39) = IIf(IsNull(TRXFILE!RETAILER_PRICE), 0, TRXFILE!RETAILER_PRICE)
                grdsales.TextMatrix(i, 40) = IIf(IsNull(TRXFILE!CESS_PER), 0, TRXFILE!CESS_PER)
                grdsales.TextMatrix(i, 41) = IIf(IsNull(TRXFILE!cess_amt), 0, TRXFILE!cess_amt)
                grdsales.TextMatrix(i, 42) = IIf(IsNull(TRXFILE!BARCODE), "", TRXFILE!BARCODE)
                grdsales.TextMatrix(i, 44) = IIf(IsNull(TRXFILE!ITEM_SPEC), "", TRXFILE!ITEM_SPEC)
                grdsales.TextMatrix(i, 45) = IIf(IsNull(TRXFILE!PUR_TAX), "", TRXFILE!PUR_TAX)
                grdsales.TextMatrix(i, 46) = IIf(IsNull(TRXFILE!UN_BILL), "N", TRXFILE!UN_BILL)
                cr_days = True
                'txtBillNo.Text = ""
                'LBLBILLNO.Caption = ""
                CMDPRINT.Enabled = True
                CmdPrintA5.Enabled = True
                
                cmdRefresh.Enabled = True
                TRXFILE.MoveNext
            Loop
            TRXFILE.Close
            Set TRXFILE = Nothing
                        
            TXTAMOUNT.text = ""
            TXTTOTALDISC.text = ""
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                Chkcancel.Value = 1
                TXTINVDATE.text = IIf(IsDate(TRXMAST!VCH_DATE), Format(TRXMAST!VCH_DATE, "DD/MM/YYYY"), "  /  /    ")
                If TRXMAST!SLSM_CODE = "A" Then
                    TXTTOTALDISC.text = IIf(IsNull(TRXMAST!DISCOUNT), "", TRXMAST!DISCOUNT)
                    OptDiscAmt.Value = True
                ElseIf TRXMAST!SLSM_CODE = "P" Then
                    If IsNull(TRXMAST!VCH_AMOUNT) Or TRXMAST!VCH_AMOUNT = 0 Then
                        TXTTOTALDISC.text = 0
                    Else
                        TXTTOTALDISC.text = IIf(IsNull(TRXMAST!DISCOUNT), "", Round((TRXMAST!DISCOUNT * 100 / TRXMAST!VCH_AMOUNT), 2))
                    End If
                    OPTDISCPERCENT.Value = True
                Else
                    TXTTOTALDISC.text = IIf(IsNull(TRXMAST!DISCOUNT), "", TRXMAST!DISCOUNT)
                    OptDiscAmt.Value = True
                End If
                LBLRETAMT.Caption = IIf(IsNull(TRXMAST!ADD_AMOUNT), "", Format(TRXMAST!ADD_AMOUNT, "0.00"))
                If TRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                txtcrdays.text = IIf(IsNull(TRXMAST!cr_days), "", TRXMAST!cr_days)
                TxtBillName.text = IIf(IsNull(TRXMAST!BILL_NAME), "", TRXMAST!BILL_NAME)
                TxtBillAddress.text = IIf(IsNull(TRXMAST!BILL_ADDRESS), "", TRXMAST!BILL_ADDRESS)
                TxtVehicle.text = IIf(IsNull(TRXMAST!VEHICLE), "", TRXMAST!VEHICLE)
                TxtOrder.text = IIf(IsNull(TRXMAST!D_ORDER), "", TRXMAST!D_ORDER)
                TxtFrieght.text = IIf(IsNull(TRXMAST!FRIEGHT), "", TRXMAST!FRIEGHT)
                Txthandle.text = IIf(IsNull(TRXMAST!HANDLE), "", TRXMAST!HANDLE)
                TxtPhone.text = IIf(IsNull(TRXMAST!PHONE), "", TRXMAST!PHONE)
                TXTTIN.text = IIf(IsNull(TRXMAST!TIN), "", TRXMAST!TIN)
                TXTAREA.text = IIf(IsNull(TRXMAST!Area), "", TRXMAST!Area)
                TxtCode.text = IIf(IsNull(TRXMAST!ACT_CODE), "", TRXMAST!ACT_CODE)
                DataList2.BoundText = IIf(IsNull(TRXMAST!ACT_CODE), "", TRXMAST!ACT_CODE)
                TXTDEALER.text = IIf(IsNull(TRXMAST!ACT_NAME), "", TRXMAST!ACT_NAME)
                DataList2.text = TXTDEALER.text
                DataList2.BoundText = IIf(IsNull(TRXMAST!ACT_CODE), "", TRXMAST!ACT_CODE)
                CMBDISTI.text = IIf(IsNull(TRXMAST!AGENT_NAME), "", TRXMAST!AGENT_NAME)
                CMBDISTI.BoundText = IIf(IsNull(TRXMAST!AGENT_CODE), "", TRXMAST!AGENT_CODE)
                TxtCN.text = IIf(IsNull(TRXMAST!CN_REF), "", TRXMAST!CN_REF)
                Select Case TRXMAST!BILL_TYPE
                    Case "W"
                        cmbtype.ListIndex = 1
                        TXTTYPE.text = 2
                    Case "V"
                        cmbtype.ListIndex = 2
                        TXTTYPE.text = 3
                    Case "M"
                        cmbtype.ListIndex = 3
                        TXTTYPE.text = 4
                    Case Else
                        cmbtype.ListIndex = 0
                        TXTTYPE.text = 1
                End Select
                
                GRDRECEIPT.rows = 1
                GRDRECEIPT.TextMatrix(0, 0) = IIf(IsNull(TRXMAST!RCPT_AMOUNT), 0, TRXMAST!RCPT_AMOUNT)
                GRDRECEIPT.rows = GRDRECEIPT.rows + 1
                GRDRECEIPT.TextMatrix(1, 0) = IIf(IsNull(TRXMAST!RCPT_REFNO), "", TRXMAST!RCPT_REFNO)
                If TRXMAST!BANK_FLAG = "Y" Then
                    GRDRECEIPT.rows = GRDRECEIPT.rows + 1
                    GRDRECEIPT.TextMatrix(2, 0) = "B"
                    GRDRECEIPT.rows = GRDRECEIPT.rows + 1
                    GRDRECEIPT.TextMatrix(3, 0) = IIf(IsNull(TRXMAST!CHQ_NO), "", TRXMAST!CHQ_NO)
                    GRDRECEIPT.rows = GRDRECEIPT.rows + 1
                    GRDRECEIPT.TextMatrix(4, 0) = IIf(IsNull(TRXMAST!BANK_CODE), "", TRXMAST!BANK_CODE)
                    GRDRECEIPT.rows = GRDRECEIPT.rows + 1
                    If Not IsNull(TRXMAST!CHQ_DATE) Then
                        GRDRECEIPT.TextMatrix(5, 0) = IIf(IsDate(TRXMAST!CHQ_DATE), TRXMAST!CHQ_DATE, "")
                    End If
                    If TRXMAST!CHQ_STATUS = "Y" Then
                        GRDRECEIPT.rows = GRDRECEIPT.rows + 1
                        GRDRECEIPT.TextMatrix(6, 0) = "Y"
                    Else
                        GRDRECEIPT.rows = GRDRECEIPT.rows + 1
                        GRDRECEIPT.TextMatrix(6, 0) = "N"
                    End If
                    GRDRECEIPT.rows = GRDRECEIPT.rows + 1
                    GRDRECEIPT.TextMatrix(7, 0) = IIf(IsNull(TRXMAST!BANK_NAME), "", TRXMAST!BANK_NAME)
                Else
                    GRDRECEIPT.rows = GRDRECEIPT.rows + 1
                    GRDRECEIPT.TextMatrix(2, 0) = "C"
                End If
                If IsNull(TRXMAST!TERMS) Or TRXMAST!TERMS = "" Then
                    chkTerms.Value = 0
                    Terms1.text = ""
                Else
                    chkTerms.Value = 1
                    Terms1.text = TRXMAST!TERMS
                End If
                Call CMBBRNCH_GotFocus
                CMBBRNCH.text = IIf(IsNull(TRXMAST!BR_NAME), "", TRXMAST!BR_NAME)
                CMBBRNCH.BoundText = IIf(IsNull(TRXMAST!BR_CODE), "", TRXMAST!BR_CODE)
                lblIGST.Caption = IIf(IsNull(TRXMAST!CUST_IGST), "N", TRXMAST!CUST_IGST)
                NEW_BILL = False
                'OLD_BILL = True
                lbltime.Caption = Format(TRXMAST!C_TIME, "hh:mm AMPM")
            Else
                lbltime.Caption = Format(Time, "hh:mm AMPM")
                TXTTIN.text = ""
                CMBBRNCH.text = ""
                If OLD_BILL = False Then
                    NEW_BILL = True
                    'OLD_BILL = False
                End If
                'TXTTYPE.Text = ""
                'cmbtype.ListIndex = -1
                cmbtype.ListIndex = 0
                TXTTYPE.text = 1
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
'            If OLD_BILL = False Then
'                Set TRXMAST = New ADODB.Recordset
'                TRXMAST.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI'", db, adOpenForwardOnly
'                If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'                    If Val(txtBillNo.Text) < IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0)) Then
'                        NEW_BILL = False
'                        OLD_BILL = True
'                    End If
'                End If
'                TRXMAST.Close
'                Set TRXMAST = Nothing
'            End If
'            If OLD_BILL = False Then
'                cmbtype.Tag = 0
'                db.Execute "SET autocommit = 0"
'                db.Execute "SET SESSION TRANSACTION ISOLATION LEVEL READ COMMITTED"
'                db.Execute "START TRANSACTION"
'
'                Dim rstBILL As ADODB.Recordset
'                Set rstBILL = New ADODB.Recordset
'                If Val(txtBillNo.Text) <> 1 Then
'                    rstBILL.Open "Select * From TRXMASTVAN WHERE VCH_NO= (SELECT MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI')", db, adOpenStatic, adLockPessimistic, adCmdText
'                    txtBillNo.Text = rstBILL!VCH_NO + 1
'                    LBLBILLNO.Caption = Val(txtBillNo.Text)
'                    rstBILL.AddNew
'                    rstBILL!TRX_TYPE = "GI"
'                    rstBILL!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'                    rstBILL!VCH_NO = txtBillNo.Text
'                    rstBILL!C_USER_ID = frmLogin.rs!USER_ID
'                    rstBILL!CREATE_DATE = Format(Date, "DD/MM/YYYY")
'                    rstBILL!C_TIME = Format(Time, "HH:MM:SS")
'                    rstBILL!C_USER_NAME = frmLogin.rs!USER_NAME
'                Else
'                    rstBILL.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockPessimistic, adCmdText
'                    If (rstBILL.EOF And rstBILL.BOF) Then
'                        rstBILL.AddNew
'                        rstBILL!TRX_TYPE = "GI"
'                        rstBILL!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'                        rstBILL!VCH_NO = txtBillNo.Text
'                        rstBILL!C_USER_ID = frmLogin.rs!USER_ID
'                        rstBILL!CREATE_DATE = Format(Date, "DD/MM/YYYY")
'                        rstBILL!C_TIME = Format(Time, "HH:MM:SS")
'                        rstBILL!C_USER_NAME = frmLogin.rs!USER_NAME
'                    End If
'                End If
'                rstBILL.Update
'                rstBILL.Close
'                Set rstBILL = Nothing
'                OLD_BILL = True
'                db.Execute "COMMIT"
'            End If
            
            'OLD_BILL = True
            LBLBILLNO.Caption = Val(txtBillNo.text)
            LBLTOTAL.Caption = ""
            LBLGROSSAMT.Caption = ""
            lblnetamount.Caption = ""
            LBLFOT.Caption = ""
            lblcomamt.text = ""
            For i = 1 To grdsales.rows - 1
                grdsales.TextMatrix(i, 0) = i
                Select Case grdsales.TextMatrix(i, 19)
                    Case "CN"
                        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                        LBLFOT.Caption = ""
                    Case Else
                        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                        LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                        LBLFOT.Caption = ""
                End Select
                
                If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                    lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
                Else
                    lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
                End If
            Next i
            LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
            TXTAMOUNT.text = ""
            If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
            ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
            End If
            LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
            lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
            lblnetamount.Caption = Format(lblnetamount.Caption, "0")
            If Val(lblnetamount.Caption) = 0 Then
                txtcommper.text = 0
            Else
                txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
            End If
            Call COSTCALCULATION
            Call Addcommission
            
            Dim rstCustomer As ADODB.Recordset
            Set rstCustomer = New ADODB.Recordset
            rstCustomer.Open "Select * From CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
            If Not (rstCustomer.EOF And rstCustomer.BOF) Then
                lblsubdealer = IIf(IsNull(rstCustomer!CUST_TYPE), "", rstCustomer!CUST_TYPE)
                lblIGST.Caption = IIf(IsNull(rstCustomer!CUST_IGST), "N", rstCustomer!CUST_IGST)
                lbladdress.Caption = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
            End If
            rstCustomer.Close
            Set rstCustomer = Nothing
            
            If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
                txtcrdays.Enabled = False
                CMDSALERETURN.Enabled = False
                CMDDELIVERY.Enabled = False
                Frame5.Visible = False
            Else
                txtcrdays.Enabled = True
                Frame5.Visible = True
            End If
            
            TXTSLNO.text = grdsales.rows
            txtBillNo.Visible = False
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
'            If grdsales.Rows > 1 Then
'                TXTDEALER.SetFocus
'                'TxtName1.SetFocus
'            Else
'                TXTDEALER.SetFocus
'                'TXTINVDATE.SetFocus
''                TxtName1.Enabled = False
''                TXTDEALER.Text = ""
''                TXTDEALER.SetFocus
'            End If
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            CHANGE_ADDRESS = False
            'Call Command2_Click
    End Select
    
    Exit Sub
ERRHAND:
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
    
    On Error GoTo ERRHAND
    
    txtBillNo.Tag = "N"
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND (ISNULL(BILL_FLAG) OR BILL_FLAG <>'Y') ", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
'        txtBillNo.Tag = "Y"
'    Else
'        txtBillNo.Tag = "N"
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
    TxtCN.text = ""
    lblbalance.Caption = ""
    Txtrcvd.text = ""
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI'", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
'        If Val(txtBillNo.Text) > i Then
'            MsgBox "The last bill No. is " & i, vbCritical, "BILL..."
'            txtBillNo.Visible = True
'            txtBillNo.SetFocus
'            Exit Sub
'        End If
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
      
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MIN(VCH_NO) From TRXFILEVAN WHERE TRX_TYPE = 'GI'", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
'        If Val(txtBillNo.Text) < i Then
'            MsgBox "This Year Starting Bill No. is " & i, vbCritical, "BILL..."
'            txtBillNo.Visible = True
'            txtBillNo.SetFocus
'            Exit Sub
'        End If
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    txtBillNo.Visible = False
    Call txtBillNo_KeyDown(13, 0)
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub txtcategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcategory.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            txtcategory.Enabled = False
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
    End Select
End Sub

Private Sub txtcategory_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTCODE_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE Like '" & Me.TxtCode.text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE Like '" & Me.TxtCode.text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
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
    CHANGE_ADDRESS = True
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
    CMDDELIVERY.Enabled = False
    CMDSALERETURN.Enabled = False
    Call CHANGEBOXCOLOR(TxtCode, True)
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then TXTDEALER.SetFocus
            'lbladdress.Caption = ""
            DataList2.SetFocus
        Case vbKeyEscape
            If M_ADD = True Then Exit Sub
            txtBillNo.Visible = True
            txtBillNo.SetFocus
    End Select
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCode_LostFocus()
    Call TXTDEALER_LostFocus
    Call CHANGEBOXCOLOR(TxtCode, False)
End Sub

Private Sub txtcommper_LostFocus()
    Call CHANGEBOXCOLOR(txtcommper, False)
End Sub

Private Sub txtcrdays_GotFocus()
    txtcrdays.SelStart = 0
    txtcrdays.SelLength = Len(txtcrdays.text)
End Sub

Private Sub txtcrdays_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            'If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            'If TxtName1.Enabled = True Then TxtName1.SetFocus
            'If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            'If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            'If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub txtcrdays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
    CHANGE_ADDRESS = True
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTDEALER_LostFocus()
    Dim RSTTRXFILE As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y' AND TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        CMDDELIVERY.Enabled = True
        Call FillDelivery_Grid
    Else
        CMDDELIVERY.Enabled = False
        grdEXPIRYLIST.FixedRows = 0
        grdEXPIRYLIST.rows = 1
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM SALERETURN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        CMDSALERETURN.Enabled = True
    Else
        CMDSALERETURN.Enabled = False
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    If DataList2.BoundText = "" Then
        cmdreturn.Enabled = False
    Else
        cmdreturn.Enabled = True
    End If
    Call CHANGEBOXCOLOR(TXTDEALER, False)
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTDISC_Change()
'    If Val(TXTDISC.Text) = 0 Or Val(TXTRETAILNOTAX.Text) = 0 Then
'        TxtDiscAmt.Text = ""
'    Else
'        TxtDiscAmt.Text = (Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text)) / 100
'    End If
'    If Val(TxtDiscAmt.Text) = 0 Then TxtDiscAmt.Text = ""
End Sub

Private Sub TXTDISC_GotFocus()
    If UCase(txtcategory.text) = "PARDHA" Or UCase(txtcategory.text) = "CLOTHES" Then
        If Val(TXTRETAILNOTAX.text) < 1000 Then
            TXTTAX.text = "5"
        Else
            TXTTAX.text = "12"
        End If
    End If
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.text)
    Call CHANGEBOXCOLOR(TXTDISC, True)
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            If lblsubdealer.Caption = "D" And txtretaildummy.Visible = True Then
                txtretaildummy.Enabled = True
                txtretaildummy.SetFocus
            Else
                txtcommi.text = 0
                Set GRDPRERATE.DataSource = Nothing
                fRMEPRERATE.Visible = False
                Call CMDADD_Click
            End If
'            If Val(TxtCessPer.Text) <> 0 Then
'                TxtCessPer.Enabled = True
'                TxtCessPer.SetFocus
'            Else
'            End If
'            TXTDISC.Enabled = False
'            cmdadd.Enabled = True
'            cmdadd.SetFocus
'            'TxtWarranty.Enabled = True
'            'TxtWarranty.SetFocus
        Case vbKeyEscape
            txtretail.Enabled = True
            txtretail.SetFocus
        Case vbKeyTab
            TxtDiscAmt.Enabled = True
            TxtDiscAmt.SetFocus
        Case vbKeyDown
            Call CMDADD_Click
    End Select
End Sub

Private Sub TXTDISC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDISC_LostFocus()
    
    TXTDISC.Tag = 0
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        TXTDISC.Tag = Val(txtretail.text) * Val(TXTDISC.text) / 100
        LBLSUBTOTAL.Caption = Format(Round(Val(txtretail.text) - Val(TXTDISC.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text) - Val(TXTDISC.Tag), 2), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.text) * Val(txtretail.text) * Val(TXTDISC.text) / 100
        LBLSUBTOTAL.Caption = Format(Round((Val(TXTQTY.text) * Val(txtretail.text)) - Val(TXTDISC.Tag), 2), ".000")
        LblGross.Caption = Format(Round((Val(TXTQTY.text) * Val(TXTRETAILNOTAX.text)) - Val(TXTDISC.Tag), 2), ".000")
        
        LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) + (Val(TXTRETAILNOTAX.text) - (Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100)) * Val(TXTQTY.text) * Val(TxtCessPer) / 100
        LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) + Round(Val(TxtCessAmt.text) * Val(TXTQTY.text), 3)
    End If
    Call cost_calculate
    Call CHANGEBOXCOLOR(TXTDISC, False)
    ''TXTDISC.Text = Format(TXTDISC.Text, ".000")

End Sub

Private Sub TxtDiscAmt_Change()
    If Val(TxtDiscAmt.text) = 0 Or Val(TXTRETAILNOTAX.text) = 0 Then
        TXTDISC.text = ""
    Else
        TXTDISC.text = Round((Val(TxtDiscAmt.text) * 100) / Val(TXTRETAILNOTAX.text), 2)
    End If
    If Val(TXTDISC.text) = 0 Then TXTDISC.text = ""
    'Call TXTDISC_LostFocus
    TXTDISC.Tag = 0
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        TXTDISC.Tag = Val(txtretail.text) * Val(TXTDISC.text) / 100
        LBLSUBTOTAL.Caption = Format(Round(Val(txtretail.text) - Val(TXTDISC.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text) - Val(TXTDISC.Tag), 2), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.text) * Val(txtretail.text) * Val(TXTDISC.text) / 100
        LBLSUBTOTAL.Caption = Format(Round((Val(TXTQTY.text) * Val(txtretail.text)) - Val(TXTDISC.Tag), 2), ".000")
        LblGross.Caption = Format(Round((Val(TXTQTY.text) * Val(TXTRETAILNOTAX.text)) - Val(TXTDISC.Tag), 2), ".000")
        
        LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) + (Val(TXTRETAILNOTAX.text) - (Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100)) * Val(TXTQTY.text) * Val(TxtCessPer) / 100
        LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) + Round(Val(TxtCessAmt.text) * Val(TXTQTY.text), 3)
    End If
End Sub

Private Sub TxtDiscAmt_GotFocus()
    If UCase(txtcategory.text) = "PARDHA" Or UCase(txtcategory.text) = "CLOTHES" Then
        If Val(TXTRETAILNOTAX.text) < 1000 Then
            TXTTAX.text = "5"
        Else
            TXTTAX.text = "12"
        End If
    End If
    TxtDiscAmt.SelStart = 0
    TxtDiscAmt.SelLength = Len(TxtDiscAmt.text)
    Call CHANGEBOXCOLOR(TxtDiscAmt, True)
End Sub

Private Sub TxtDiscAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            If lblsubdealer.Caption = "D" And txtretaildummy.Visible = True Then
                txtretaildummy.Enabled = True
                txtretaildummy.SetFocus
            Else
                txtcommi.text = 0
                Set GRDPRERATE.DataSource = Nothing
                fRMEPRERATE.Visible = False
                Call CMDADD_Click
            End If
'            If Val(TxtCessPer.Text) <> 0 Then
'                TxtCessPer.Enabled = True
'                TxtCessPer.SetFocus
'            Else
'            End If
'            TXTDISC.Enabled = False
'            cmdadd.Enabled = True
'            cmdadd.SetFocus
'            'TxtWarranty.Enabled = True
'            'TxtWarranty.SetFocus
        Case vbKeyEscape
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyDown
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtDiscAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtDiscAmt_LostFocus()
    Call CHANGEBOXCOLOR(TxtDiscAmt, False)
End Sub

Private Sub TXTEXPIRY_LostFocus()
    TXTEXPIRY.BackColor = vbWhite
End Sub

Private Sub TxtFrieght_GotFocus()
    TxtFrieght.SelStart = 0
    TxtFrieght.SelLength = Len(TxtFrieght.text)
    Call CHANGEBOXCOLOR(TxtFrieght, True)
End Sub

Private Sub TxtFrieght_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            'If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            'If TxtName1.Enabled = True Then TxtName1.SetFocus
            'If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            'If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            'If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            'If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
            cmdRefresh.Enabled = True
            CMDEXIT.Enabled = False
            CMDPre.Caption = "Save && Pre&vious"
            cmdnxt.Caption = "Save && &Next"
    End Select
End Sub

Private Sub TxtFrieght_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtFrieght_LostFocus()
    Call TXTTOTALDISC_LostFocus
    Call CHANGEBOXCOLOR(TxtFrieght, False)
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.text)
    TXTINVDATE.BackColor = &H98F3C1
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.text = "  /  /    " Then
                TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
                TXTDEALER.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
                TXTDEALER.SetFocus
            End If
        Case vbKeyEscape
            If M_ADD = True Then Exit Sub
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

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
    CMDDELIVERY.Enabled = False
    CMDSALERETURN.Enabled = False
    
    fRMEPRERATE.Visible = False
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    Call CHANGEBOXCOLOR(TXTDEALER, True)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList2.SetFocus
        Case vbKeyEscape
            If M_ADD = True Then Exit Sub
            txtBillNo.Visible = True
            txtBillNo.SetFocus
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

Private Sub TXTINVDATE_LostFocus()
'    If Not IsDate(TXTINVDATE.Text) Then
'        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
'        TXTINVDATE.SetFocus
'        Exit Sub
'    End If
'
'    If (DateValue(TXTINVDATE.Text) < DateValue(MDIMAIN.DTFROM.value)) Or (DateValue(TXTINVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.value))) Then
'        'db.Execute "delete from Users"
'        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
'        TXTINVDATE.SetFocus
'        Exit Sub
'    End If
    TXTINVDATE.BackColor = vbWhite
End Sub

Private Sub TXTITEMCODE_LostFocus()
    Call CHANGEBOXCOLOR(TXTITEMCODE, False)
End Sub

Private Sub TXTMRP_GotFocus()
    TxtMRP.SelStart = 0
    TxtMRP.SelLength = Len(TxtMRP.text)
    Call CHANGEBOXCOLOR(TxtMRP, True)
End Sub

Private Sub TXTMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TxtMRP.Text) = 0 Then Exit Sub
            If TXTEXPIRY.Visible = True Then
                TXTEXPIRY.Enabled = True
                TXTEXPIRY.SetFocus
            Else
                TXTTAX.SetFocus
            End If
        Case vbKeyEscape
            TXTQTY.Enabled = True
            TXTQTY.SetFocus

    End Select
End Sub

Private Sub TXTMRP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtMRP_LostFocus()
    TxtMRP.text = Format(TxtMRP.text, ".000")
    Call CHANGEBOXCOLOR(TxtMRP, False)
End Sub

Private Sub TxtName1_Change()
    If Trim(TxtName1.text) = "" Then Exit Sub
    If CHANGE_NAME = False Then Exit Sub
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset

    M_STOCK = 0
    Set grdtmp.DataSource = Nothing
    If PHYFLAG = True Then
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMASTVAN  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY CATEGORY, ITEM_SLNO", db, adOpenStatic, adLockReadOnly
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMASTVAN  WHERE ucase(CATEGORY) = 'OWN' AND ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        If MDIMAIN.lblitemrepeat.Caption = "Y" Then
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS From ITEMMASTVAN WHERE (ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '%" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        Else
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS From ITEMMASTVAN WHERE (ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        End If
        PHYFLAG = False
    Else
        PHY.Close
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMASTVAN  WHERE ucase(CATEGORY) = 'OWN' AND ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        If MDIMAIN.lblitemrepeat.Caption = "Y" Then
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS From ITEMMASTVAN WHERE (ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '%" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        Else
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS From ITEMMASTVAN WHERE (ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        End If
        PHYFLAG = False
    End If
    Set grdtmp.DataSource = PHY
    
    If PHY.RecordCount > 0 Then
        grdtmp.Visible = True
    Else
        Set grdtmp.DataSource = Nothing
        grdtmp.Visible = False
        Exit Sub
    End If
    grdtmp.Columns(0).Caption = "ITEM CODE"
    grdtmp.Columns(0).Width = 2000
    grdtmp.Columns(1).Caption = "ITEM NAME"
    grdtmp.Columns(1).Width = 6500
    grdtmp.Columns(2).Caption = "QTY"
    grdtmp.Columns(2).Width = 900
    grdtmp.Columns(6).Caption = "RT"
    grdtmp.Columns(6).Width = 900
    grdtmp.Columns(4).Width = 0
    grdtmp.Columns(4).Width = 0
    grdtmp.Columns(5).Width = 0
    grdtmp.Columns(3).Width = 0
    grdtmp.Columns(7).Width = 0
    grdtmp.Columns(8).Width = 0
    grdtmp.Columns(9).Width = 0
    grdtmp.Columns(10).Width = 0
    grdtmp.Columns(10).Caption = "L/Pack"
    grdtmp.Columns(11).Caption = "LP"
    grdtmp.Columns(11).Width = 900
    grdtmp.Columns(12).Caption = "WS"
    grdtmp.Columns(12).Width = 900
    grdtmp.Columns(13).Width = 0
    grdtmp.Columns(14).Width = 0
    grdtmp.Columns(15).Width = 0
    grdtmp.Columns(16).Width = 0
    grdtmp.Columns(17).Width = 0
    grdtmp.Columns(18).Width = 0
    grdtmp.Columns(19).Width = 0
    grdtmp.Columns(20).Caption = "MRP"
    grdtmp.Columns(20).Width = 900
    grdtmp.Columns(21).Width = 0
    grdtmp.Columns(22).Width = 2500
    grdtmp.Columns(21).Caption = "DISC"
    grdtmp.Columns(21).Width = 700
    grdtmp.Columns(22).Width = 0
    grdtmp.Columns(23).Width = 0
    grdtmp.Columns(24).Width = 0
    grdtmp.Columns(25).Width = 0
    grdtmp.Columns(26).Caption = "LW Price"
    grdtmp.Columns(26).Width = 800
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TxtName1_GotFocus()
    CHANGE_NAME = True
    TxtName1.SelStart = 0
    TxtName1.SelLength = Len(TxtName1.text)
    TxtName1_Change
    grdsales.Enabled = True
    'Set grdtmp.DataSource = Nothing
    'grdtmp.Visible = False
    
    fRMEPRERATE.Visible = False
    cmdadd.Enabled = False
    txtBatch.Enabled = False
    'TXTQTY.Enabled = False
    TXTFREE.Enabled = False
    TxtMRP.Enabled = False
    TXTEXPIRY.Enabled = False
    TXTTAX.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    txtretail.Enabled = False
    TXTDISC.Enabled = False
    TxtDiscAmt.Enabled = False
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    txtretaildummy.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    TxtSpec.Enabled = False
    Call CHANGEBOXCOLOR(TxtName1, True)
End Sub

Private Sub TxtName1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtBillNo.Text) > 100 Then Exit Sub
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
                MsgBox "Permission Denied", vbOKOnly, "Sales"
                Exit Sub
            End If
'            If MDIMAIN.LBLGSTWRN.Caption = "Y" And Trim(TXTTIN.Text) = "" Then
'                MsgBox "B2C Bill Not allowed", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            If txtBillNo.Tag = "Y" Then
                MsgBox "Cannot modify here", vbOKOnly, "Sales"
                Exit Sub
            End If
            If UCase(TxtName1.text) = "OT" Then TXTITEMCODE.text = "OT"
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            End If
            LBLDNORCN.Caption = ""
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
    End Select
End Sub

Private Sub TxtName1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtName1_LostFocus()
    Call CHANGEBOXCOLOR(TxtName1, False)
End Sub

Private Sub txtNetrate_GotFocus()
    Call CHANGEBOXCOLOR(txtNetrate, True)
End Sub

Private Sub TxtOrder_LostFocus()
    Call CHANGEBOXCOLOR(TxtOrder, False)
End Sub

Private Sub TxtPhone_LostFocus()
    Call CHANGEBOXCOLOR(TxtPhone, False)
End Sub

Private Sub txtphsearch_LostFocus()
    Call CHANGEBOXCOLOR(txtphsearch, False)
End Sub

Private Sub txtPrintname_GotFocus()
    Call CHANGEBOXCOLOR(txtPrintname, True)
End Sub

Private Sub txtPrintname_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape, vbKeyReturn
            If cmdadd.Enabled = True Then cmdadd.SetFocus
            'If CMDPRINT.Enabled = True Then CMDPRINT.SetFocus
            If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
            If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
            If txtBillNo.Visible = True Then txtBillNo.SetFocus
            'If TxtName1.Enabled = True Then TxtName1.SetFocus
            'If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            'If txtretail.Enabled = True Then txtretail.SetFocus
            'If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
            'If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            'If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            'If txtcommi.Enabled = True Then txtcommi.SetFocus
    End Select
End Sub

Private Sub txtPrintname_LostFocus()
    Call CHANGEBOXCOLOR(txtPrintname, False)
End Sub

Private Sub TXTPRODUCT_Change()
        If item_change = True Then Exit Sub
        If CHANGE_NAME = False Then Exit Sub
        Dim i As Long
        Dim RSTBATCH As ADODB.Recordset
    
        M_STOCK = 0
        Set grdtmp.DataSource = Nothing
        If PHYFLAG = True Then
            'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMASTVAN  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY CATEGORY, ITEM_SLNO", db, adOpenStatic, adLockReadOnly
            'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMASTVAN  WHERE ucase(CATEGORY) = 'OWN'AND ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "' OR MRP Like '%" & Trim(Me.TxtName1.Text) & "') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            If Trim(TxtName1.text) = "" Then
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL  From ITEMMASTVAN  WHERE (ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            Else
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL  From ITEMMASTVAN  WHERE (ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%') AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            End If
            PHYFLAG = False
        Else
            PHY.Close
            'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMASTVAN  WHERE ucase(CATEGORY) = 'OWN'AND ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "' OR MRP Like '%" & Trim(Me.TxtName1.Text) & "') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            If Trim(TxtName1.text) = "" Then
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL  From ITEMMASTVAN  WHERE (ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            Else
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL  From ITEMMASTVAN  WHERE (ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '%" & Trim(Me.TxtName1.text) & "%')  AND (ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%' ) AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            End If
            PHYFLAG = False
        End If
        Set grdtmp.DataSource = PHY
        
        If PHY.RecordCount > 0 Then
            grdtmp.Visible = True
        Else
            Set grdtmp.DataSource = Nothing
            grdtmp.Visible = False
            Exit Sub
        End If
        grdtmp.Columns(0).Caption = "ITEM CODE"
    grdtmp.Columns(0).Width = 2000
    grdtmp.Columns(1).Caption = "ITEM NAME"
    grdtmp.Columns(1).Width = 6500
    grdtmp.Columns(2).Caption = "QTY"
    grdtmp.Columns(2).Width = 900
    grdtmp.Columns(6).Caption = "RT"
    grdtmp.Columns(6).Width = 900
    grdtmp.Columns(4).Width = 0
    grdtmp.Columns(4).Width = 0
    grdtmp.Columns(5).Width = 0
    grdtmp.Columns(3).Width = 0
    grdtmp.Columns(7).Width = 0
    grdtmp.Columns(8).Width = 0
    grdtmp.Columns(9).Width = 0
    grdtmp.Columns(10).Width = 0
    grdtmp.Columns(10).Caption = "L/Pack"
    grdtmp.Columns(11).Caption = "LP"
    grdtmp.Columns(11).Width = 900
    grdtmp.Columns(12).Caption = "WS"
    grdtmp.Columns(12).Width = 900
    grdtmp.Columns(13).Width = 0
    grdtmp.Columns(14).Width = 0
    grdtmp.Columns(15).Width = 0
    grdtmp.Columns(16).Width = 0
    grdtmp.Columns(17).Width = 0
    grdtmp.Columns(18).Width = 0
    grdtmp.Columns(19).Width = 0
    grdtmp.Columns(20).Caption = "MRP"
    grdtmp.Columns(20).Width = 900
    grdtmp.Columns(21).Width = 0
    grdtmp.Columns(22).Width = 2500
    grdtmp.Columns(21).Caption = "DISC"
    grdtmp.Columns(21).Width = 700
    grdtmp.Columns(22).Width = 0
    grdtmp.Columns(23).Width = 0
    grdtmp.Columns(24).Width = 0
    grdtmp.Columns(25).Width = 0
    grdtmp.Columns(26).Caption = "LW Price"
    grdtmp.Columns(26).Width = 800
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTPRODUCT_GotFocus()
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLNETCOST.Caption = ""
    LBLNETPROFIT.Caption = ""
    
    TXTPRODUCT.Tag = TXTPRODUCT.text
    TXTPRODUCT.text = ""
    TXTPRODUCT.text = TXTPRODUCT.Tag
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.text)
    CHANGE_NAME = True
    If Trim(TXTPRODUCT.text) <> "" Or Trim(TxtName1.text) <> "" Then Call TXTPRODUCT_Change
    grdsales.Enabled = True
    'grdtmp.Visible = True
    
    fRMEPRERATE.Visible = False
    cmdadd.Enabled = False
    txtBatch.Enabled = False
    TXTQTY.Enabled = False
    TXTFREE.Enabled = False
    TxtMRP.Enabled = False
    TXTEXPIRY.Enabled = False
    TXTTAX.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    txtretail.Enabled = False
    TXTDISC.Enabled = False
    TxtDiscAmt.Enabled = False
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    txtretaildummy.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    TxtSpec.Enabled = False
    Call CHANGEBOXCOLOR(TXTPRODUCT, True)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
    
        Case vbKeyReturn
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
                MsgBox "Permission Denied", vbOKOnly, "Sales"
                Exit Sub
            End If
'            If MDIMAIN.LBLGSTWRN.Caption = "Y" And Trim(TXTTIN.Text) = "" Then
'                MsgBox "B2C Bill Not allowed", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            If Trim(TxtName1.text) = "" And Trim(TXTPRODUCT.text) = "" Then Exit Sub
            If txtBillNo.Tag = "Y" Then
                MsgBox "Cannot modify here", vbOKOnly, "Sales"
                Exit Sub
            End If
            
            M_STOCK = 0
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            txtappendcomm.text = ""
            TXTAPPENDTOTAL.text = ""
            txtretail.text = ""
            txtNetrate.text = ""
            txtBatch.text = ""
            TxtWarranty.text = ""
            TxtWarranty_type.text = ""
            TXTRETAILNOTAX.text = ""
            TXTSALETYPE.text = ""
            TXTFREE.text = ""
            optnet.Value = True
            TxtMRP.text = ""
            TXTTAX.text = ""
            TXTDISC.text = ""
            TxtDiscAmt.text = ""
            TxtCessAmt.text = ""
            TxtCessPer.text = ""
            LBLUNBILL.Caption = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            LblPack.text = "1"
            lblunit.text = "Nos"
            txtcommi.text = ""
            On Error Resume Next
            TXTITEMCODE.text = grdtmp.Columns(0)
            
            For i = 1 To grdsales.rows - 1
                If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.text) Then
                    If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next i
            
            TXTPRODUCT.text = grdtmp.Columns(1)
            txtPrintname.text = grdtmp.Columns(1)
            TxtCessPer.text = IIf(IsNull(grdtmp.Columns(24)), "", grdtmp.Columns(24))
            TxtCessAmt.text = IIf(IsNull(grdtmp.Columns(25)), "", grdtmp.Columns(25))
            TXTDISC.text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
            LBLUNBILL.Caption = IIf(IsNull(grdtmp.Columns(27)), "N", grdtmp.Columns(27))
            TxtMRP.text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
            Select Case cmbtype.ListIndex
                Case 0
                    txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                Case 1
                    txtretail.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                Case 2
                    txtretail.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                Case 3
                    txtretail.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                    If Val(TXTRETAILNOTAX.text) = 0 Then
                        txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    End If
            End Select
            LblPack.text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
            lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
            'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
'            If Trim(TXTPRODUCT.Text) = "" Then
'                TxtName1.Enabled = True
'                TxtName1.SetFocus
'                Exit Sub
'            End If
            'cmddelete.Enabled = False
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
            If UCase(TxtName1.text) = "OT" Then TXTITEMCODE.text = "OT"
            If UCase(TXTITEMCODE.text) <> "OT" Then
                Set grdtmp.DataSource = Nothing
                If PHYFLAG = True Then
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC,  CESS_PER, CESS_AMT, P_LWS, UN_BILL  From ITEMMASTVAN  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "' ", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                Else
                    PHY.Close
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC,  CESS_PER, CESS_AMT, P_LWS, UN_BILL  From ITEMMASTVAN  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "' ", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                End If
                Set grdtmp.DataSource = PHY
                
                If PHY.RecordCount = 0 Then
                    MsgBox "Item not found!!!!", , "Sales"
                    Exit Sub
                End If
                If PHY.RecordCount = 1 Then
                    SERIAL_FLAG = False
                    lblactqty.Caption = ""
                    lblbarcode.Caption = ""
                    TXTITEMCODE.text = grdtmp.Columns(0)
                    TXTPRODUCT.text = grdtmp.Columns(1)
                    txtPrintname.text = grdtmp.Columns(1)
                    TxtCessPer.text = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
                    TxtCessAmt.text = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
                    TXTDISC.text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
                    LBLUNBILL.Caption = IIf(IsNull(grdtmp.Columns(25)), "N", grdtmp.Columns(25))
                    Select Case PHY!check_flag
                        Case "M"
                            OPTTaxMRP.Value = True
                            TXTTAX.text = grdtmp.Columns(4)
                            TXTSALETYPE.text = "2"
                        Case "V"
                            OPTVAT.Value = True
                            TXTSALETYPE.text = "2"
                            TXTTAX.text = grdtmp.Columns(4)
                        Case Else
                            TXTSALETYPE.text = "2"
                            optnet.Value = True
                            TXTTAX.text = "0"
                    End Select
                    
                    If MDIMAIN.lblPriceSplit.Caption = "Y" Then GoTo SKIP_BATCH
                    Set RSTBATCH = New ADODB.Recordset
                    Select Case cmbtype.ListIndex
                        Case 1
                            'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                            'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        Case 2
                            'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_VAN >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                            'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        Case 3
                            'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                            'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            RSTBATCH.Open "Select DISTINCT ITEM_CODE, MRP, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        Case Else
                            'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_RETAIL >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                            'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    End Select
                    If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                        If BATCH_DISPLAY = True Then
                            If RSTBATCH.RecordCount >= 1 Then
                                Set grdtmp.DataSource = Nothing
                                grdtmp.Visible = False
                                Call FILL_BATCHGRID
                                RSTBATCH.Close
                                Set RSTBATCH = Nothing
                                Exit Sub
                            ElseIf RSTBATCH.RecordCount = 1 Then
                                'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                                txtBatch.text = IIf(IsNull(RSTBATCH!REF_NO), "", RSTBATCH!REF_NO)
                            End If
                        Else
                            If RSTBATCH.RecordCount > 1 Then
                                Set grdtmp.DataSource = Nothing
                                grdtmp.Visible = False
                                Call FILL_BATCHGRID
                                RSTBATCH.Close
                                Set RSTBATCH = Nothing
                                Exit Sub
                            ElseIf RSTBATCH.RecordCount = 1 Then
                                'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                                txtBatch.text = IIf(IsNull(RSTBATCH!REF_NO), "", RSTBATCH!REF_NO)
                            End If
                        End If
                    End If
                    RSTBATCH.Close
                    Set RSTBATCH = Nothing
SKIP_BATCH:
                    Call CONTINUE
                Else
                    grdtmp.Visible = True
                    grdtmp.Columns(0).Caption = "ITEM CODE"
                    grdtmp.Columns(0).Width = 1200
                    grdtmp.Columns(1).Caption = "ITEM NAME"
                    grdtmp.Columns(1).Width = 3500
                    grdtmp.Columns(2).Caption = "QTY"
                    grdtmp.Columns(2).Width = 1000
                    grdtmp.Columns(6).Caption = "RATE"
                    grdtmp.Columns(6).Width = 1000
                    grdtmp.Columns(4).Width = 0
                    grdtmp.Columns(4).Width = 0
                    grdtmp.Columns(5).Width = 0
                    grdtmp.Columns(3).Width = 0
                    grdtmp.Columns(7).Width = 0
                    grdtmp.Columns(8).Width = 0
                    grdtmp.Columns(9).Width = 0
                    grdtmp.Columns(10).Caption = "L/Pack"
                    grdtmp.Columns(11).Caption = "LP"
                    grdtmp.Columns(10).Width = 800
                    grdtmp.Columns(11).Width = 800
                    grdtmp.Columns(12).Caption = "WS"
                    grdtmp.Columns(12).Width = 800
                    grdtmp.Columns(13).Width = 0
                    grdtmp.Columns(14).Width = 0
                    grdtmp.Columns(15).Width = 0
                    grdtmp.Columns(16).Width = 0
                    grdtmp.Columns(17).Width = 0
                    grdtmp.Columns(18).Width = 0
                    grdtmp.Columns(19).Width = 0
                    grdtmp.SetFocus
                    'Call FILL_ITEMGRID
                    Exit Sub
                End If
            End If
JUMPNONSTOCK:
            TxtName1.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                TXTQTY.text = 1
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
            End If
        Case vbKeyEscape
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtDiscAmt.Enabled = False
            TxtName1.SetFocus
            'cmddelete.Enabled = False
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub
Private Function CONTINUE()
    Dim i As Long
                Dim DEL_FLAG As Boolean
                DEL_FLAG = False
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.text) And grdsales.TextMatrix(i, 19) = "DN" Then DEL_FLAG = True
                Next i
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.text) Then
                        If DEL_FLAG = False Then
                            If MsgBox("This Item Already exists in Line No. " & grdsales.TextMatrix(i, 0) & "... Do yo want to modify this item", vbYesNo + vbDefaultButton2, "BILL..") = vbYes Then
                                grdsales.Row = i
                                'grdsales.TextMatrix(grdsales.Row, 0) = grdsales.TextMatrix(i, 0)
                                TXTSLNO.text = grdsales.TextMatrix(i, 0)
                                Call CMDMODIFY_Click
                                Exit Function
                            End If
                        Else
                            Select Case grdsales.TextMatrix(i, 19)
                                Case "CN", "DN"
                                    Exit For
                            End Select
'                            If SERIAL_FLAG = False Then
'                                TXTSLNO.Text = i
'                                TXTAPPENDQTY.Text = Val(grdsales.TextMatrix(i, 3))
'                                TXTFREEAPPEND.Text = Val(grdsales.TextMatrix(i, 20))
'                                txtappendcomm.Text = Val(grdsales.TextMatrix(i, 24))
'                                Exit For
'                            End If
                        End If
                        Exit For
                    End If
                Next i
                txtcategory.text = IIf(IsNull(PHY!Category), "", PHY!Category)
                If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                    TXTQTY.text = 1
                    txtretail.Enabled = True
                    txtretail.SetFocus
                    Exit Function
                End If
            
'                Select Case cmbtype.ListIndex
'                    Case 0 'VAN
'                        'txtretail.Text = IIf(IsNull(grdtmp.Columns(13)), "", grdtmp.Columns(13))
'                        'kannattu
'                        TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(13)), "", grdtmp.Columns(13))
'                    Case 1 'RT
'                        'txtretail.Text = IIf(IsNull(grdtmp.Columns(6)), "", grdtmp.Columns(6))
'                        'kannattu
'                        TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(6)), "", grdtmp.Columns(6))
'                    Case 2 'WS
'                        'txtretail.Text = IIf(IsNull(grdtmp.Columns(12)), "", grdtmp.Columns(12))
'                        'kannattu
'                        TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(6)), "", grdtmp.Columns(6))
'                End Select
                LblPack.text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                'txtretail.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)) * Val(LblPack.Text))
                Select Case cmbtype.ListIndex
                    Case 1
                        txtretail.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    Case 2
                        txtretail.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    Case 3
                        txtretail.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        If Val(TXTRETAILNOTAX.text) = 0 Then
                            txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                            TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        End If
                    Case Else
                        txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                End Select
                If Val(TxtCessPer.text) <> 0 Or Val(TxtCessAmt.text) <> 0 Then
                    TXTRETAILNOTAX.text = (Val(txtretail.text) - Val(TxtCessAmt.text)) / (1 + (Val(TXTTAX.text) / 100) + (Val(TxtCessPer.text) / 100))
                    txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 3)
                    TXTRETAILNOTAX.text = Val(txtretail.text)
                End If
                
                lblretail.Caption = IIf(IsNull(grdtmp.Columns(6)), "", grdtmp.Columns(6))
                lblwsale.Caption = IIf(IsNull(grdtmp.Columns(12)), "", grdtmp.Columns(12))
                lblvan.Caption = IIf(IsNull(grdtmp.Columns(13)), "", grdtmp.Columns(13))
                LBLMRP.Caption = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
                lblcase.Caption = IIf(IsNull(grdtmp.Columns(11)), "", grdtmp.Columns(11))
                lblcrtnpack.Caption = IIf(IsNull(grdtmp.Columns(10)), "", grdtmp.Columns(10))
                lblLWPrice.Caption = IIf(IsNull(grdtmp.Columns(24)), "", grdtmp.Columns(24))
                
                
                lblunit.text = IIf(IsNull(grdtmp.Columns(17)), "Nos", grdtmp.Columns(17))
                TxtWarranty.text = IIf(IsNull(grdtmp.Columns(18)), "", grdtmp.Columns(18))
                TxtWarranty_type.text = IIf(IsNull(grdtmp.Columns(19)), "", grdtmp.Columns(19))
                TxtMRP.text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
                TXTDISC.text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
                'LblPack.Text = IIf(IsNull(grdtmp.Columns(10)), "", grdtmp.Columns(10))
                'If Val(LblPack.Text) = 0 Then LblPack.Text = "1"
                'txtretail.Text = IIf(IsNull(grdtmp.Columns(11)), "", grdtmp.Columns(11))
            
                If grdtmp.Columns(7) = "A" Then
                    txtretaildummy.text = IIf(IsNull(grdtmp.Columns(9)), "P", grdtmp.Columns(9))
                    lblcomtype.Caption = "`"
                    TxtRetailmode.text = "A"
                    txtcommi.text = Format(Round(Val(txtretaildummy.text), 2), "0.00")
                Else
                    txtretaildummy.text = IIf(IsNull(grdtmp.Columns(8)), "P", grdtmp.Columns(8))
                    lblcomtype.Caption = "%"
                    TxtRetailmode.text = "P"
                    txtcommi.text = Format(Round((Val(TXTRETAILNOTAX.text) * Val(txtretaildummy.text) / 100), 2), "0.00")
                End If
                If Val(txtretaildummy.text) = 0 Then lblcomtype.Caption = "`"
                
                Select Case PHY!check_flag
                    Case "M"
                        OPTTaxMRP.Value = True
                        TXTTAX.text = grdtmp.Columns(4)
                        TXTSALETYPE.text = "2"
                    Case "V"
                        OPTVAT.Value = True
                        TXTSALETYPE.text = "2"
                        TXTTAX.text = grdtmp.Columns(4)
                    Case Else
                        TXTSALETYPE.text = "2"
                        optnet.Value = True
                        TXTTAX.text = "0"
                End Select
                
'                OPTVAT.value = True
'                TXTTAX.Text = "14.5"
'                TXTSALETYPE.Text = "2"
                
                TXTUNIT.text = grdtmp.Columns(5)
                                   
                'TXTPRODUCT.Enabled = False
                'TXTQTY.Enabled = True
                '
                'TXTQTY.SetFocus
                Exit Function
End Function

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPRODUCT_LostFocus()
    CHANGE_NAME = False
    Call CHANGEBOXCOLOR(TXTPRODUCT, False)
End Sub

Private Sub TXTQTY_GotFocus()
    
    If Val(LblPack.text) = 0 Then LblPack.text = 1
'    If TxtRetailmode.Text = "A" Then
'        txtcommi.Text = Format(Round(Val(txtretaildummy.Text), 2), "0.00")
'    Else
'        txtcommi.Text = Format(Round((Val(TXTRETAILNOTAX.Text) * Val(txtretaildummy.Text) / 100), 2), "0.00")
'    End If
    If M_EDIT = False Then
        If Val(lblOr_Pack.Caption) <= 1 Then
            FrmeType.Visible = False
        Else
            FrmeType.Visible = True
        End If
        If Val(LblPack.text) = Val(lblOr_Pack.Caption) Then
            OptNormal.Value = True
        Else
            OptLoose.Value = True
        End If
    Else
        FrmeType.Visible = False
    End If
'    TxtName1.Enabled = False
'    TXTPRODUCT.Enabled = False
'    TXTITEMCODE.Enabled = False
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.text)
    TXTQTY.Tag = Trim(TXTPRODUCT.text)
    
    lblunit.Enabled = True
    cmdadd.Enabled = True
    txtBatch.Enabled = True
    TXTQTY.Enabled = True
    TXTFREE.Enabled = True
    TxtMRP.Enabled = True
    TXTEXPIRY.Enabled = True
    TXTTAX.Enabled = True
   
    txtretail.Enabled = True
    TXTRETAILNOTAX.Enabled = True
    TXTDISC.Enabled = True
    
    TxtDiscAmt.Enabled = True
    TxtCessPer.Enabled = True
    TxtCessAmt.Enabled = True
    txtcommi.Enabled = True
    txtretaildummy.Enabled = True
    TxtWarranty.Enabled = True
    TxtWarranty_type.Enabled = True
    txtPrintname.Enabled = True
    TxtSpec.Enabled = True
    
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    
    If M_EDIT = False And FrmeType.Visible = True Then
         OptNormal.Value = False
         OptLoose.Value = False
    End If
    Call CHANGEBOXCOLOR(TXTQTY, True)
    'TXTQTY.SetFocus
    
    On Error GoTo ERRHAND
    'If Val(TXTDISC.Text) <> 0 Then Exit Sub
    If M_EDIT = True Then Exit Sub
    
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    Select Case cmbtype.ListIndex
        Case 1
            RSTTRXFILE.Open "SELECT WSDISC FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenForwardOnly
        Case 2
            RSTTRXFILE.Open "SELECT VPDISC FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenForwardOnly
        Case 3
            RSTTRXFILE.Open "SELECT MRPDISC FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenForwardOnly
        Case Else
            RSTTRXFILE.Open "SELECT RTDISC FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenForwardOnly
    End Select
    If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
        TXTDISC.text = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        If Val(TXTDISC.text) = 0 Then
            Dim RSTITEMDISC As ADODB.Recordset
            Set RSTITEMDISC = New ADODB.Recordset
            RSTITEMDISC.Open "SELECT CUST_DISC From ITEMMASTVAN  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "'", db, adOpenForwardOnly
            If Not (RSTITEMDISC.EOF Or RSTITEMDISC.BOF) Then
                TXTDISC.text = IIf(IsNull(RSTITEMDISC.Fields(0)), 0, RSTITEMDISC.Fields(0))
            End If
            RSTITEMDISC.Close
            Set RSTITEMDISC = Nothing
        End If
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Double
    
    Select Case KeyCode
        Case vbKeyReturn
            If FrmeType.Visible = True And OptNormal.Value = False And OptLoose.Value = False Then
                MsgBox "Please select Loose / Full Option", vbOKOnly, "EzBiz"
                TXTQTY.SetFocus
                Exit Sub
            End If
            If Val(TXTQTY.text) = 0 Then Exit Sub
            i = 0
            If Val(LblPack.text) = 0 Then LblPack.text = 1
            If Not (UCase(txtcategory.text) = "SERVICES" Or UCase(txtcategory.text) = "SELF") Then
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT CLOSE_QTY  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                    If (IsNull(RSTTRXFILE!CLOSE_QTY)) Then RSTTRXFILE!CLOSE_QTY = 0
                    i = RSTTRXFILE!CLOSE_QTY / Val(LblPack.text)
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
                If Val(TXTQTY.text) = 0 Then Exit Sub
'                If M_EDIT = False And Val(TXTQTY.Text) > i Then
'                    MsgBox "AVAILABLE STOCK IS  " & i & " ", , "SALES"
'                    TXTQTY.SelStart = 0
'                    TXTQTY.SelLength = Len(TXTQTY.Text)
'                    Exit Sub
'                End If
                'If i <> 0 Then
                    If M_EDIT = False And SERIAL_FLAG = True And Val(TXTQTY.text) > (Val(lblactqty.Caption) / Val(LblPack.text)) Then
                        MsgBox "AVAILABLE STOCK IN THE PARTICULAR STOCK IS  " & (Val(lblactqty.Caption) / Val(LblPack.text)) & " ", , "SALES"
                        TXTQTY.SelStart = 0
                        TXTQTY.SelLength = Len(TXTQTY.text)
                        Exit Sub
                    End If
                    If M_EDIT = False And Val(TXTQTY.text) > i Then
                        If (MsgBox("AVAILABLE STOCK IS  " & i & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                            'MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                            TXTQTY.SelStart = 0
                            TXTQTY.SelLength = Len(TXTQTY.text)
                            Exit Sub
                        End If
                    End If
                'End If
SKIP:
                If UCase(TXTITEMCODE.text) = "OT" Then
                    If TxtMRP.Visible = True Then
                        TxtMRP.Enabled = True
                        TxtMRP.SetFocus
                    Else
                        If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                            txtretail.Enabled = True
                            txtretail.SetFocus
                        Else
                            If Val(TXTTAX.text) = 0 Then
                                TXTTAX.Enabled = True
                                TXTTAX.SetFocus
                            Else
                                If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                                    TXTRETAILNOTAX.Enabled = True
                                    TXTRETAILNOTAX.SetFocus
                                Else
                                    If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
                                        TXTRETAILNOTAX.Enabled = True
                                        TXTRETAILNOTAX.SetFocus
                                    Else
                                        txtretail.Enabled = True
                                        txtretail.SetFocus
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                        txtretail.Enabled = True
                        txtretail.SetFocus
                    Else
                        If Val(TXTTAX.text) = 0 Then
                            TXTTAX.Enabled = True
                            TXTTAX.SetFocus
                        Else
                            If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                                TXTRETAILNOTAX.Enabled = True
                                TXTRETAILNOTAX.SetFocus
                            Else
                                If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
                                    TXTRETAILNOTAX.Enabled = True
                                    TXTRETAILNOTAX.SetFocus
                                Else
                                    txtretail.Enabled = True
                                    txtretail.SetFocus
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                    txtretail.Enabled = True
                    txtretail.SetFocus
                Else
                    TXTTAX.Enabled = True
                    TXTTAX.SetFocus
                End If
            End If
         Case vbKeyEscape
            If txtBatch.Visible = False Then
                If M_EDIT = True Then Exit Sub
                LblPack.Enabled = True
                LblPack.SetFocus
            Else
                txtBatch.Enabled = True
                txtBatch.SetFocus
            End If
        Case vbKeyTab
            If TXTFREE.Visible = False Then
                If TxtMRP.Visible = True Then
                    TxtMRP.SetFocus
                End If
            Else
                TXTFREE.Enabled = True
                TXTFREE.SetFocus
            End If
        Case vbKeyDown
            If Val(TXTQTY.text) = 0 Then Exit Sub
'            If Val(TXTQTY.Text) <> 0 And MDIMAIN.StatusBar.Panels(14).Text <> "Y" And Val(TXTRETAILNOTAX.Text) = 0 Then
'                TXTRETAILNOTAX.Enabled = True
'                TXTRETAILNOTAX.SetFocus
'                Exit Sub
'            End If
'            If Val(TXTQTY.Text) <> 0 And MDIMAIN.StatusBar.Panels(14).Text = "Y" And Val(txtretail.Text) = 0 Then
'                txtretail.Enabled = True
'                txtretail.SetFocus
'                Exit Sub
'            End If
            CMDADD_Click
            'cmdadd.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRHAND
    Dim TRXMAST As ADODB.Recordset
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case 102, 70
            If FrmeType.Visible = False Then
                KeyAscii = 0
                Exit Sub
            End If
            If M_EDIT = False Then OptNormal.Value = True
            LblPack.text = Val(lblOr_Pack.Caption)
            Call LblPack_LostFocus
            KeyAscii = 0
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "SELECT PACK_TYPE From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                lblunit.text = IIf(IsNull(TRXMAST!PACK_TYPE) Or TRXMAST!PACK_TYPE = "", "Nos", Trim(TRXMAST!PACK_TYPE))
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
        Case 76, 108
            If FrmeType.Visible = False Then
                KeyAscii = 0
                Exit Sub
            End If
            If M_EDIT = False Then OptLoose.Value = True
            If Val(lblcrtnpack.Caption) = 0 Then lblcrtnpack.Caption = 1
            LblPack.text = Val(lblcrtnpack.Caption)
            Call LblPack_LostFocus
            KeyAscii = 0
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "SELECT FULL_PACK From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                lblunit.text = IIf(IsNull(TRXMAST!FULL_PACK) Or TRXMAST!FULL_PACK = "", "Nos", Trim(TRXMAST!FULL_PACK))
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
        Case Else
            KeyAscii = 0
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description, vbOKOnly, "EzBiz"
End Sub

Private Sub TXTQTY_LostFocus()
    
    Dim RSTITEMCOST As ADODB.Recordset
    
    TXTQTY.text = Format(TXTQTY.text, ".000")
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    If Val(TXTRETAILNOTAX.text) = 0 Then
        TXTDISC.Tag = Val(TXTDISC.text) / 100
        TXTTAX.Tag = Val(TXTTAX.text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Round(Val(txtretail.text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
        LblGross.Caption = Format((Val(TXTQTY.text) * Round(Val(TXTRETAILNOTAX.text), 3)) - Val(TXTDISC.Tag), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.text) * Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100
        TXTTAX.Tag = Val(TXTQTY.text) * Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Round(Val(TXTRETAILNOTAX.text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
        LblGross.Caption = Format((Val(TXTQTY.text) * Round(Val(TXTRETAILNOTAX.text), 3)) - Val(TXTDISC.Tag), ".000")
    End If
    
    On Error GoTo ERRHAND
    Set RSTITEMCOST = New ADODB.Recordset
    RSTITEMCOST.Open "SELECT ITEM_COST, SALES_PRICE, SALES_TAX, UN_BILL From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMCOST.EOF Or RSTITEMCOST.BOF) Then
        LBLITEMCOST.Caption = IIf(IsNull(RSTITEMCOST!ITEM_COST), "", RSTITEMCOST!ITEM_COST * Val(LblPack.text))
        LBLSELPRICE.Caption = IIf(IsNull(RSTITEMCOST!SALES_PRICE), "", RSTITEMCOST!SALES_PRICE * Val(LblPack.text))
        LBLNETCOST.Caption = Round(Val(LBLITEMCOST.Caption) + (Val(LBLITEMCOST.Caption) * IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX / 100)), 2)
        LBLNETCOST.Tag = IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX)
        LBLUNBILL.Caption = IIf(IsNull(RSTITEMCOST!UN_BILL), "N", RSTITEMCOST!UN_BILL)
    End If
    RSTITEMCOST.Close
    Set RSTITEMCOST = Nothing
    
    If Not (UCase(txtcategory.text) = "SERVICES" Or UCase(txtcategory.text) = "SELF") Then
        Set RSTITEMCOST = New ADODB.Recordset
        RSTITEMCOST.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND TRX_TYPE = '" & Trim(TXTTRXTYPE.text) & "' AND VCH_NO = " & Val(TXTVCHNO.text) & " AND LINE_NO = " & Val(TXTLINENO.text) & " AND TRX_YEAR = '" & Val(TrxRYear.text) & "' AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly, adCmdText
        With RSTITEMCOST
            If Not (.EOF And .BOF) Then
                LBLITEMCOST.Caption = IIf(IsNull(RSTITEMCOST!ITEM_COST), "", RSTITEMCOST!ITEM_COST * Val(LblPack.text))
                LBLNETCOST.Caption = Round(Val(LBLITEMCOST.Caption) + (Val(LBLITEMCOST.Caption) * IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX / 100)), 2)
                LBLNETCOST.Tag = IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX)
            Else
                RSTITEMCOST.Close
                Set RSTITEMCOST = Nothing
                Set RSTITEMCOST = New ADODB.Recordset
                RSTITEMCOST.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTITEMCOST.EOF And RSTITEMCOST.BOF) Then
                    LBLITEMCOST.Caption = IIf(IsNull(RSTITEMCOST!ITEM_COST), "", RSTITEMCOST!ITEM_COST * Val(LblPack.text))
                    LBLNETCOST.Caption = Round(Val(LBLITEMCOST.Caption) + (Val(LBLITEMCOST.Caption) * IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX / 100)), 2)
                    LBLNETCOST.Tag = IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX)
                    RSTITEMCOST.Close
                    Set RSTITEMCOST = Nothing
                Else
                    RSTITEMCOST.Close
                    Set RSTITEMCOST = Nothing
                    Set RSTITEMCOST = New ADODB.Recordset
                    RSTITEMCOST.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
                    If Not (RSTITEMCOST.EOF And RSTITEMCOST.BOF) Then
                        LBLITEMCOST.Caption = IIf(IsNull(RSTITEMCOST!ITEM_COST), "", RSTITEMCOST!ITEM_COST * Val(LblPack.text))
                        LBLNETCOST.Caption = Round(Val(LBLITEMCOST.Caption) + (Val(LBLITEMCOST.Caption) * IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX / 100)), 2)
                        LBLNETCOST.Tag = IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX)
                    End If
                    RSTITEMCOST.Close
                    Set RSTITEMCOST = Nothing
                End If
            End If
        End With
    End If
    Call CHANGEBOXCOLOR(TXTQTY, False)
    Exit Sub
ERRHAND:
    MsgBox err.Description

End Sub

Private Sub Txtrcvd_Change()
    lblbalance.Caption = Format(Round(Val(Txtrcvd.text) - Val(lblnetamount.Caption), 2))
End Sub

Private Sub Txtrcvd_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtrcvd_GotFocus()
    Txtrcvd.SelStart = 0
    Txtrcvd.SelLength = Len(Txtrcvd.text)
    Call CHANGEBOXCOLOR(Txtrcvd, True)
End Sub

Private Sub Txtrcvd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape, vbKeyReturn
            'If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If txtretail.Enabled = True Then txtretail.SetFocus
            If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
            If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
    End Select
End Sub


Private Sub Txtrcvd_LostFocus()
    Call CHANGEBOXCOLOR(Txtrcvd, False)
End Sub

Private Sub txtretaildummy_Change()
    If TxtRetailmode.text = "A" Then
        txtcommi.text = Format(Round(Val(txtretaildummy.text), 2), "0.00")
    Else
        txtcommi.text = Format(Round((Val(TXTRETAILNOTAX.text) * Val(txtretaildummy.text) / 100), 2), "0.00")
    End If
End Sub

Private Sub txtretaildummy_GotFocus()
    txtretaildummy.SelStart = 0
    txtretaildummy.SelLength = Len(txtretaildummy.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    Call CHANGEBOXCOLOR(txtretaildummy, True)
End Sub

Private Sub txtretaildummy_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If txtcommi.Text = "" Then Exit Sub
            If Val(txtcommi.text) > Val(txtretail.text) Then
                MsgBox "Commission Rate greater than actual Rate", vbOKOnly, "Sales"
                txtcommi.SetFocus
                Exit Sub
            End If
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Call TXTDISC_LostFocus
            cmdadd.Enabled = True
            cmdadd.SetFocus
'            TXTDISC.Enabled = False
'            cmdadd.Enabled = True
'            cmdadd.SetFocus
'            'TxtWarranty.Enabled = True
'            'TxtWarranty.SetFocus
        Case vbKeyEscape
            If MDIMAIN.StatusBar.Panels(16).text = "Y" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTDISC.Enabled = True
                TXTDISC.SetFocus
            End If
        Case vbKeyDown
            Call TXTDISC_LostFocus
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtretaildummy_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtretaildummy_LostFocus()
    Call CHANGEBOXCOLOR(txtretaildummy, False)
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.text)
    Chkcancel.Value = 0
    grdsales.Enabled = True
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    
    fRMEPRERATE.Visible = False
    cmdadd.Enabled = False
    txtBatch.Enabled = False
    TXTQTY.Enabled = False
    TXTFREE.Enabled = False
    TxtMRP.Enabled = False
    TXTEXPIRY.Enabled = False
    TXTTAX.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    txtretail.Enabled = False
    TXTDISC.Enabled = False
    TxtDiscAmt.Enabled = False
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    txtretaildummy.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    TxtSpec.Enabled = False
    Call CHANGEBOXCOLOR(TXTSLNO, True)
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
'            If Trim(TXTTIN.Text) = "" Then
'                MsgBox "FORM 8B Bill Not allowed", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            'If Val(TXTSLNO.Text) < grdsales.Rows Then Exit Sub
'            If MDIMAIN.LBLGSTWRN.Caption = "Y" And Trim(TXTTIN.Text) = "" Then
'                MsgBox "B2C Bill Not allowed", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            If Val(TXTSLNO.text) < grdsales.rows And frmLogin.rs!Level <> "0" And NEW_BILL = False Then
                MsgBox "Permission Denied", vbOKOnly, "Sales"
                Exit Sub
            End If
            On Error Resume Next
            grdsales.Row = Val(TXTSLNO.text)
            On Error GoTo ERRHAND
            If Val(TXTSLNO.text) = 0 Then
                SERIAL_FLAG = False
                lblactqty.Caption = ""
                lblbarcode.Caption = ""
                TXTSLNO.text = ""
                TXTPRODUCT.text = ""
                txtPrintname.text = ""
                TxtSpec.text = ""
                TXTQTY.text = ""
                TXTEXPIRY.text = "  /  "
                TXTAPPENDQTY.text = ""
                TXTFREEAPPEND.text = ""
                txtappendcomm.text = ""
                TXTAPPENDTOTAL.text = ""
                TXTFREE.text = ""
                optnet.Value = True
                TxtMRP.text = ""
                
                TXTDISC.text = ""
                TxtDiscAmt.text = ""
                TxtCessAmt.text = ""
                TxtCessPer.text = ""
                LBLUNBILL.Caption = ""
                LBLSUBTOTAL.Caption = ""
                LblGross.Caption = ""
                TXTITEMCODE.text = ""
                TXTVCHNO.text = ""
                TXTLINENO.text = ""
                TXTTRXTYPE.text = ""
                TrxRYear.text = ""
                TXTUNIT.text = ""
                TXTSLNO.text = grdsales.rows
                'cmddelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.text) >= grdsales.rows Then
                TXTSLNO.text = grdsales.rows
                'CmdDelete.Enabled = False
                'CMDMODIFY.Enabled = False
            End If
            If Val(TXTSLNO.text) < grdsales.rows Then
                lblP_Rate.Caption = "1"
                TXTSLNO.text = grdsales.TextMatrix(Val(TXTSLNO.text), 0)
                TXTPRODUCT.text = grdsales.TextMatrix(Val(TXTSLNO.text), 2)
                TXTQTY.text = grdsales.TextMatrix(Val(TXTSLNO.text), 3)
                TXTFREE.text = grdsales.TextMatrix(Val(TXTSLNO.text), 20)
                TxtMRP.text = grdsales.TextMatrix(Val(TXTSLNO.text), 5)
                TXTDISC.text = grdsales.TextMatrix(Val(TXTSLNO.text), 8)
                TXTTAX.text = grdsales.TextMatrix(Val(TXTSLNO.text), 9)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.text), 12), ".000")
                
                TXTITEMCODE.text = grdsales.TextMatrix(Val(TXTSLNO.text), 13)
                TXTVCHNO.text = grdsales.TextMatrix(Val(TXTSLNO.text), 14)
                TXTLINENO.text = grdsales.TextMatrix(Val(TXTSLNO.text), 15)
                TXTTRXTYPE.text = grdsales.TextMatrix(Val(TXTSLNO.text), 16)
                TrxRYear.text = grdsales.TextMatrix(Val(TXTSLNO.text), 43)
                TXTUNIT.text = grdsales.TextMatrix(Val(TXTSLNO.text), 4)
                TxtSpec.text = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 44))
                'TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 22)
                TxtRetailmode.text = "A"
                lblcomtype.Caption = "`"
                Select Case grdsales.TextMatrix(Val(TXTSLNO.text), 17)
                    Case "M"
                        OPTTaxMRP.Value = True
                        TXTSALETYPE.text = "2"
                    Case "V"
                        OPTVAT.Value = True
                        TXTSALETYPE.text = "2"
                    Case Else
                        TXTSALETYPE.text = "2"
                        optnet.Value = True
                        TXTTAX.text = "0"
                End Select
                txtBatch.text = grdsales.TextMatrix(Val(TXTSLNO.text), 10)
                TXTRETAILNOTAX.text = grdsales.TextMatrix(Val(TXTSLNO.text), 6)
                txtretail.text = grdsales.TextMatrix(Val(TXTSLNO.text), 7)
                TXTSALETYPE.text = grdsales.TextMatrix(Val(TXTSLNO.text), 23)
                txtcategory.text = grdsales.TextMatrix(Val(TXTSLNO.text), 25)
'                If UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICE CHARGE" Then
'                    txtretaildummy.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)), 2)
'                    'txtcommi.Text = 0 'Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) / Val(TXTQTY.Text), 2)
'                Else
'                    txtretaildummy.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)), 2)
'                    'txtcommi.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) / Val(TXTQTY.Text), 2)
'                End If
                txtretaildummy.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24))
                txtcommi.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24))
                If Not IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 38)) Then
                    TXTEXPIRY.text = "  /  "
                Else
                    TXTEXPIRY.text = grdsales.TextMatrix(Val(TXTSLNO.text), 38)
                End If
                TxtName1.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTITEMCODE.Enabled = False
                TXTQTY.Enabled = False
                
                TXTTAX.Enabled = False
                TXTFREE.Enabled = False
                txtretail.Enabled = False
                TXTRETAILNOTAX.Enabled = False
                TXTDISC.Enabled = False
                TxtDiscAmt.Enabled = False
                TxtMRP.Enabled = False
                Select Case grdsales.TextMatrix(Val(TXTSLNO.text), 19)
                    Case "CN", "DN"
                        CmdDelete.Enabled = True
                        CmdDelete.SetFocus
                        
                    Case Else
                        CMDMODIFY.Enabled = True
                        CMDMODIFY.SetFocus
                        CmdDelete.Enabled = True
                End Select
                LBLDNORCN.Caption = grdsales.TextMatrix(Val(TXTSLNO.text), 19)
                LblPack.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 27))
                lblOr_Pack.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 27))
                TxtWarranty.text = grdsales.TextMatrix(Val(TXTSLNO.text), 28)
                TxtWarranty_type.text = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 29))
                lblunit.text = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 30))
                txtPrintname.text = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 33))
                LblGross.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 34))
                lblretail.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 39))
                TxtCessPer.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 40))
                TxtCessAmt.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 41))
                lblbarcode.Caption = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 42))
                Set grdtmp.DataSource = Nothing
                grdtmp.Visible = False
                TXTSLNO.Enabled = False
                grdsales.Enabled = False
                Exit Sub
            End If
SKIP:
            lblP_Rate.Caption = "0"
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            TXTQTY.Enabled = False
            
            TXTSLNO.Enabled = False
            TXTTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtDiscAmt.Enabled = False
            Set grdtmp.DataSource = Nothing
            grdtmp.Visible = False
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            If grdsales.rows > 1 Then
                CMDPRINT.Enabled = True
                CmdPrintA5.Enabled = True
                cmdRefresh.Enabled = True
                CmdPrintA5.SetFocus
            Else
                FRMEHEAD.Enabled = True
                TxtCode.Enabled = True
                TXTDEALER.Enabled = True
                TXTDEALER.SetFocus
            End If
            LBLDNORCN.Caption = ""
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
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

Private Sub TXTSLNO_LostFocus()
    Call CHANGEBOXCOLOR(TXTSLNO, False)
End Sub

Private Sub TxtSpec_GotFocus()
    Call CHANGEBOXCOLOR(TxtSpec, True)
End Sub

Private Sub TxtSpec_LostFocus()
    Call CHANGEBOXCOLOR(TxtSpec, False)
End Sub

Private Sub TxtTax_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.text)
    If Val(TXTTAX.text) = 0 Then TXTTAX.text = ""
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    Call CHANGEBOXCOLOR(TXTTAX, True)
End Sub

Private Sub TxtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If MDIMAIN.LBLTAXWARN.Caption = "Y" Then If Trim(TXTTAX.text) = "" Then Exit Sub
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                    TXTRETAILNOTAX.Enabled = True
                    TXTRETAILNOTAX.SetFocus
                Else
                    If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
                        TXTRETAILNOTAX.Enabled = True
                        TXTRETAILNOTAX.SetFocus
                    Else
                        txtretail.Enabled = True
                        txtretail.SetFocus
                    End If
                End If
            End If
        Case vbKeyEscape
            If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                If M_EDIT = True Then Exit Sub
                TxtName1.Enabled = True
                TXTPRODUCT.Enabled = True
                TXTITEMCODE.Enabled = True
                TXTPRODUCT.SetFocus
                Exit Sub
            End If
            If UCase(txtcategory.text) <> "SERVICES" Then
                If TXTEXPIRY.Visible = True Then
                    TXTEXPIRY.Enabled = True
                    TXTEXPIRY.SetFocus
                Else
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                End If
            Else
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
            End If
        Case vbKeyDown
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtTax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
    If optnet.Value = True And Val(TXTTAX.text) > 0 Then
        OPTVAT.Value = True
        TXTRETAILNOTAX_LostFocus
    End If
    txtmrpbt.text = 100 * Val(TxtMRP.text) / (100 + Val(TXTTAX.text))
    Call CHANGEBOXCOLOR(TXTTAX, False)
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
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEGRDTMP.Visible = False
    
    
    If ITEM_FLAG = True Then
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMASTVAN  WHERE ITEM_NAME Like '%" & TXTPRODUCT.text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMASTVAN  WHERE ITEM_NAME Like '%" & TXTPRODUCT.text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    End If

    Set GRDPOPUPITEM.DataSource = PHY_ITEM
    'GRDPOPUPITEM.RowHeight = 350
    GRDPOPUPITEM.Columns(0).Visible = False
    GRDPOPUPITEM.Columns(1).Caption = "ITEM NAME"
    GRDPOPUPITEM.Columns(1).Width = 3800
    GRDPOPUPITEM.Columns(2).Caption = "QTY"
    GRDPOPUPITEM.Columns(2).Width = 1200
    GRDPOPUPITEM.Columns(3).Caption = "RT"
    GRDPOPUPITEM.Columns(3).Width = 0
    GRDPOPUPITEM.Columns(4).Caption = "WS"
    GRDPOPUPITEM.Columns(4).Width = 1220
    GRDPOPUPITEM.Columns(5).Caption = "SCHEME"
    GRDPOPUPITEM.Columns(5).Width = 0
    GRDPOPUPITEM.Columns(6).Caption = "CRTN"
    GRDPOPUPITEM.Columns(6).Width = 0
    GRDPOPUPITEM.SetFocus
End Function

Private Function STOCKADJUST()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    M_STOCK = 0
    On Error GoTo ERRHAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY From RTRXFILEVAN where ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rststock.EOF
        M_STOCK = M_STOCK + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
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
    
ERRHAND:
    MsgBox err.Description
End Function

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
    Call CHANGEBOXCOLOR(TXTFREE, False)
End Sub

Private Sub OPTDISCPERCENT_Click()
    TXTTOTALDISC.SetFocus
End Sub

Private Sub Optdiscamt_Click()
    TXTTOTALDISC.SetFocus
End Sub

Private Sub TXTTIN_GotFocus()
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    Call CHANGEBOXCOLOR(TXTTIN, True)
End Sub

Private Sub TXTTIN_LostFocus()
'    If MDIMAIN.LBLGSTWRN.Caption = "Y" Then
'        If Len(Trim(TXTTIN.Text)) <> 15 Then
'            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "SALES BILL"
'            TXTTIN.SetFocus
'            Exit Sub
'        End If
'
'        If Val(Left(Trim(TXTTIN.Text), 2)) = 0 Then
'            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "SALES BILL"
'            TXTTIN.SetFocus
'            Exit Sub
'        End If
'
'        If Val(Mid(Trim(TXTTIN.Text), 13, 1)) = 0 Then
'            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "SALES BILL"
'            TXTTIN.SetFocus
'            Exit Sub
'        End If
'
'        If Val(Mid(Trim(TXTTIN.Text), 14, 1)) <> 0 Then
'            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "SALES BILL"
'            TXTTIN.SetFocus
'            Exit Sub
'        End If
'    End If
    Call CHANGEBOXCOLOR(TXTTIN, False)
End Sub

Private Sub TXTTOTALDISC_GotFocus()
    TXTTOTALDISC.SelStart = 0
    TXTTOTALDISC.SelLength = Len(TXTTOTALDISC.text)
End Sub

Private Sub TXTTOTALDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            'If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            'If TxtName1.Enabled = True Then TxtName1.SetFocus
            'If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            'If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            'If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            'If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
            cmdRefresh.Enabled = True
            CMDEXIT.Enabled = False
            CMDPre.Caption = "Save && Pre&vious"
            cmdnxt.Caption = "Save && &Next"
        End Select
End Sub

Private Sub TXTTOTALDISC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTTOTALDISC_LostFocus()
    Dim i As Long
    lblnetamount.Caption = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                lblnetamount.Caption = Val(lblnetamount.Caption) - Val(grdsales.TextMatrix(i, 12))
            Case Else
                lblnetamount.Caption = Val(lblnetamount.Caption) + Val(grdsales.TextMatrix(i, 12))
        End Select
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.text = 0
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
    LBLPROFIT.Caption = Round(Val(LBLTOTAL.Caption) - (Val(LBLDISCAMT.Caption) + Val(lblcomamt.text)) - Val(LBLTOTALCOST.Caption), 2)
    
End Sub

Private Function ReportGeneratION()
    
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Integer
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
    '//Create Report Heading
    '//chr(27) & Chr(72) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold

    'Print #1, Chr(13)
    'Print #1, Chr(13)
    'Print #1, Chr(13)
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    'Print #1, Chr(13)
    Print #1, Chr(13)
    'If (Trim(TXTREMARKS.Text) <> "") Then Print #1, Chr(27) & Chr(72) & Chr(10) & Space(47) & AlignRight(Trim(TXTREMARKS.Text), 30)
'    If (Trim(TXTREMARKS.Text) <> "") Then Print #1, Chr(27) & Chr(67) & Chr(0) & Space(48) & AlignRight(Trim(TXTREMARKS.Text), 30)
    Dim PRNSPACE As Single
    Dim BIL_PRE, BILL_SUF, INVMSG As String
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenForwardOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        INVMSG = IIf(IsNull(RSTCOMPANY!INV_MSGS), "", RSTCOMPANY!INV_MSGS)
        
        'Print #1, Chr(27) & Chr(72) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & Chr(27) & Chr(72)
        PRNSPACE = Int(39 - Len(IIf(IsNull(RSTCOMPANY!COMP_NAME) Or RSTCOMPANY!COMP_NAME = "", 1, RSTCOMPANY!COMP_NAME)) / 2)
        If PRNSPACE < 0 Then PRNSPACE = 0
        Print #1, Chr(27) & Chr(72) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(72) & Chr(10) & _
              Space(PRNSPACE) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) & _
              Chr(27) & Chr(72)
              
'        Print #1, Space(PRNSPACE) & Chr(27) & Chr(72) & Chr(14) & _
              Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& _
              Chr(27) & Chr(72)
        
'        Print #1, Chr(27) & Chr(72) & Chr(14) & _
              Chr(15) & "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz"
        
        PRNSPACE = Int(44 - Len(IIf(IsNull(RSTCOMPANY!Address) Or RSTCOMPANY!Address = "", 1, RSTCOMPANY!Address)) / 2)
        If PRNSPACE < 0 Then PRNSPACE = 0
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & Space(PRNSPACE) & RSTCOMPANY!Address
        
        PRNSPACE = Int(44 - Len(IIf(IsNull(RSTCOMPANY!HO_NAME) Or RSTCOMPANY!HO_NAME = "", 1, RSTCOMPANY!HO_NAME)) / 2)
        If PRNSPACE < 0 Then PRNSPACE = 0
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & Space(PRNSPACE) & RSTCOMPANY!HO_NAME
        
        PRNSPACE = Int(44 - ((Len(IIf(IsNull(RSTCOMPANY!TEL_NO) Or RSTCOMPANY!TEL_NO = "", 1, RSTCOMPANY!FAX_NO)) + Len(IIf(IsNull(RSTCOMPANY!FAX_NO) Or RSTCOMPANY!FAX_NO = "", 1, RSTCOMPANY!FAX_NO)) + 9) / 2))
        If PRNSPACE < 0 Then PRNSPACE = 0
        Print #1, Space(7) & Space(PRNSPACE) & "Phone: " & RSTCOMPANY!TEL_NO & ", " & RSTCOMPANY!FAX_NO
        'Print #1,
        'Print #1, AlignRight("DL NO. " & RSTCOMPANY!DL_NO, 94)
        'Print #1, AlignRight(RSTCOMPANY!ML_NO, 94)
        'Print #1, Space(58) & AlignRight("DL NO. " & RSTCOMPANY!CST, 30)
        'Print #1, Space(58) & AlignRight(RSTCOMPANY!DL_NO, 30)
        Print #1, Space(7) & "GSTIN No. " & RSTCOMPANY!CST
        'Print #1, Chr(27) & Chr(72) & Space(7) & "State: Kerala (32) "
        Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(19) & "TAX INVOICE"
        'Print #1, Chr(27) & Chr(72) & Space(45) & "TAX INVOICE"
        'Print #1, Space(40) & Chr(27) & Chr(72) & Chr(10) & "TAX INVOICE"
        'Print #1, Chr(27) & Chr(72) & Chr(10) & Space(41) & AlignLeft("INVOICE FORM 8H", 16)
        'Print #1, Chr(27) & Chr(72) & Chr(10) & Space(37) & AlignLeft("CASH/CREDIT INVOICE", 19)
    
        'If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, LBLDATE.Caption)
        'Print #1, Space(7) & "Bill No. " & Trim(LBLBILLNO.Caption) & Chr(27) & Chr(72) & Space(2) & AlignRight("Date:" & LBLDATE.Caption, 67) '& Space(2) & LBLTIME.Caption
        'LBLDATE.Caption = Date
        
        'Print #1, Chr(27) & Chr(72) & Space(7) & "Bill No. " & "LIM " & Trim(LBLBILLNO.Caption) & AlignRight("Date:" & LBLDATE.Caption, 67)  '& Space(2) & LBLTIME.Caption
        Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Bill No. " & BIL_PRE & Trim(LBLBILLNO.Caption) & BILL_SUF, 35) & Space(2) & AlignRight("Date:" & TXTINVDATE.text, 43)  '& Space(2) & LBLTIME.Caption
        
        'Print #1, "Bill No. " & BIL_PRE & Trim(txtBillNo.Text) & BILL_SUF & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 110) '& Space(2) & LBLTIME.Caption
        Print #1, "TO: " & Trim(TxtBillName.text)
        If Trim(TxtBillAddress.text) <> "" Then Print #1, Trim(TxtBillAddress.text)
        If Trim(TxtPhone.text) <> "" Then Print #1, "Phone: " & Trim(TxtPhone.text)
        If Trim(TXTTIN.text) <> "" Then Print #1, "GSTIN: " & Trim(TXTTIN.text)
    
        'Print #1, Chr(27) & Chr(72) & Space(7) & "Bill No. " & Trim(LBLBILLNO.Caption) & AlignRight("Date:" & LBLDATE.Caption, 67)  '& Space(2) & LBLTIME.Caption
        'LBLDATE.Caption = Date
        Print #1,
        'Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Patient: " & Trim(Mid(TxtBillName.Text, 1, 35)), 35) & Space(2) & AlignRight("Doctor: " & Mid(CMBDISTI.Text, 1, 43), 50)
        'If (Trim(TXTREMARKS.Text) <> "") Then Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft(Trim(TXTREMARKS.Text), 30)
        'Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("Patient: " & Trim(Mid(TXTPATIENT.Text, 1, 35)), 35) & Space(2) & AlignRight("Doctor: " & Mid(TXTDOCTOR.Text, 1, 35), 43)
        
       ' Print #1, Chr(27) & Chr(72) & Space(7) & "Salesman: CS"
    
        Print #1, Space(3) & RepeatString("-", 89)
        Print #1, Space(3) & AlignLeft("HSN/SAC", 8) & Space(0) & _
                AlignLeft("Description", 11) & Space(23) & _
                AlignLeft("GST%", 5) & Space(4) & _
                AlignLeft("Rate", 5) & Space(3) & _
                AlignLeft("Qty", 5) & Space(2) & _
                AlignLeft("Disc", 5) & Space(3) & _
                AlignLeft("Amount", 12) & _
                Chr(27) & Chr(72)  '//Bold Ends
        
'        Print #1, Space(3) & RepeatString("-", 89)
'        Print #1, Space(7) & AlignLeft("Description", 11) & Space(15) & _
'                AlignLeft("MFR", 3) & Space(3) & _
'                AlignLeft("Batch", 6) & Space(4) & _
'                AlignLeft("Ex.Dt", 6) & Space(4) & _
'                AlignLeft("Rate", 4) & Space(3) & _
'                AlignLeft("Qty", 4) & Space(4) & _
'                AlignLeft("Disc", 4) & Space(4) & _
'                AlignLeft("Amount", 6) & _
'                Chr(27) & Chr(72)  '//Bold Ends
        
        Print #1, Space(3) & RepeatString("-", 89)
        Dim HSNCODE As String
        Dim RSTHSNCODE As ADODB.Recordset
        Dim MRPVALUE, TaxAmt, DISCAMT, KFCVAL As Double
        TaxAmt = 0
        KFCVAL = 0
        'TotalTax = 0
        DISCAMT = 0
        MRPVALUE = 0
        For i = 1 To grdsales.rows - 1
'            If Val(grdsales.TextMatrix(i, 8)) > 0 Then
'                TaxAmt = Round((Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 8)) / 100) * Val(grdsales.TextMatrix(i, 3)), 3)
'
'            End If
'            TotalTax = TotalTax + TaxAmt
            MRPVALUE = MRPVALUE + Val(Val(grdsales.TextMatrix(i, 7))) * Val(grdsales.TextMatrix(i, 3))
            DISCAMT = (Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100) * Val(grdsales.TextMatrix(i, 3))
            'KFCVAL = KFCVAL + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 47)) / 100) * Val(grdsales.TextMatrix(i, 3))
            Set RSTHSNCODE = New ADODB.Recordset
            RSTHSNCODE.Open "SELECT * From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTHSNCODE.EOF And RSTHSNCODE.BOF) Then
                HSNCODE = IIf(IsNull(RSTHSNCODE!REMARKS), "", RSTHSNCODE!REMARKS)
            End If
            RSTHSNCODE.Close
            Set RSTHSNCODE = Nothing
            
            If Trim(grdsales.TextMatrix(i, 33)) = "" Then
                grdsales.Tag = Trim(grdsales.TextMatrix(i, 2))
            Else
                grdsales.Tag = Trim(grdsales.TextMatrix(i, 33))
            End If
            Print #1, Space(3) & AlignLeft(Left(HSNCODE, 6), 8) & _
                AlignLeft(Mid(grdsales.Tag, 1, 34), 34) & _
                AlignRight(Format(Round(grdsales.TextMatrix(i, 9), 2), "0.00"), 7) & _
                AlignRight(Format(Round(grdsales.TextMatrix(i, 7), 2), "0.00"), 9) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 6) & _
                AlignRight(Format(Round(DISCAMT, 2), "0.00"), 9) & Space(0) & _
                AlignRight(Format(grdsales.TextMatrix(i, 12), "0.00"), 10) & _
                Chr(27) & Chr(72)  '//Bold Ends
        Next i
        
'        Print #1, Space(3) & RepeatString("-", 89)
'        For i = 1 To grdsales.Rows - 1
'            Print #1, Space(7) & AlignLeft(grdsales.TextMatrix(i, 2), 26) & _
'                AlignLeft(grdsales.TextMatrix(i, 17), 5) & Space(2) & _
'                AlignLeft(grdsales.TextMatrix(i, 9), 9) & _
'                AlignLeft(Format(grdsales.TextMatrix(i, 10), "mm/yy"), 5) & Space(0) & _
'                AlignRight(Format(Round(grdsales.TextMatrix(i, 6), 2), "0.00"), 10) & _
'                AlignRight(grdsales.TextMatrix(i, 3), 4) & Space(1) & _
'                AlignRight(Format(Round((Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 7)) / 100) * Val(grdsales.TextMatrix(i, 3)), 2), "0.00"), 8) & _
'                AlignRight(Format(grdsales.TextMatrix(i, 11), "0.00"), 10) & _
'                Chr(27) & Chr(72)  '//Bold Ends
'        Next i
'
        'Print #1, Space(7) & AlignRight("-------------", 89)
        'Print #1, Space(7) & AlignRight("-------------", 89)
'        LBLUSER.Tag = ""
'        Set RSTHSNCODE = New ADODB.Recordset
'        RSTHSNCODE.Open "SELECT * from USERS WHERE USER_ID = '" & LBLUSER.Caption & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (RSTHSNCODE.EOF And RSTHSNCODE.BOF) Then
'            LBLUSER.Tag = IIf(IsNull(RSTHSNCODE!USER_NAME), "", RSTHSNCODE!USER_NAME)
'        End If
'        RSTHSNCODE.Close
'        Set RSTHSNCODE = Nothing
'        LBLUSER.Tag = "-----Salesman: " & Trim(LBLUSER.Tag)
'        Print #1,
'        Print #1,
'        Print #1, Space(7) & Trim(LBLUSER.Tag) & RepeatString("-", 89 - Len(LBLUSER.Tag))
    
'        If TotalTax > 0 Then
'            Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(3) & AlignLeft("CGST Tax Amt: " & Format(Round(TotalTax / 2, 2), "0.00"), 48)
'            Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(3) & AlignLeft("SGST Tax Amt: " & Format(Round(TotalTax / 2, 2), "0.00"), 48)
'            Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(3) & AlignLeft("IGST Tax Amt: " & "0.00", 48)
'        End If
        If Val(LBLDISCAMT.Caption) > 0 And Val(LBLRETAMT.Caption) > 0 Then
            Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(20) & AlignRight("RET AMOUNT: ", 13) & AlignRight((Format(Round(LBLRETAMT.Caption, 2), "####.00")), 10)
            'Print #1, Chr(27) & Chr(72) & Space(20) & AlignRight("CRDT AMOUNT ", 66) & AlignRight((Format(LBLRETAMT.Caption, "####.00")), 10)
        ElseIf Val(LBLDISCAMT.Caption) > 0 And Val(LBLRETAMT.Caption) = 0 Then
        ElseIf Val(LBLDISCAMT.Caption) = 0 And Val(LBLRETAMT.Caption) > 0 Then
            Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(20) & AlignRight("RET AMOUNT: ", 13) & AlignRight((Format(Round(LBLRETAMT.Caption, 2), "####.00")), 10)
        End If
        Print #1,
        Print #1, Space(3) & RepeatString("-", 89)
        
        Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(20) & AlignRight("BILL AMOUNT: ", 13) & AlignRight((Format(Round(MRPVALUE, 2), "####.00")), 10)
        If Val(LBLDISCAMT.Caption) > 0 Then Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(24) & AlignRight("DISC AMOUNT: ", 13) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 10)
        Num = Round((MRPVALUE - Val(LBLDISCAMT.Caption) - Val(LBLRETAMT.Caption)), 0) - (MRPVALUE - Val(LBLDISCAMT.Caption) - Val(LBLRETAMT.Caption))
        If Num < 0 Then
            Print #1, Chr(27) & Chr(72) & Space(14) & AlignRight("Round off:", 55) & "" & AlignRight(Format(Round(MRPVALUE - Val(LBLDISCAMT.Caption) - Val(LBLRETAMT.Caption), 0) - (MRPVALUE - Val(LBLDISCAMT.Caption) - Val(LBLRETAMT.Caption)), "0.00"), 22)
        Else
            Print #1, Chr(27) & Chr(72) & Space(14) & AlignRight("Round off:", 55) & "" & AlignRight(Format(Round(MRPVALUE - Val(LBLDISCAMT.Caption) - Val(LBLRETAMT.Caption), 0) - (MRPVALUE - Val(LBLDISCAMT.Caption) - Val(LBLRETAMT.Caption)), "0.00"), 22)
        End If
        Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(20) & AlignRight("NET AMOUNT:  ", 13) & AlignRight((Format(Round(Val(lblnetamount.Caption), 0), "####.00")), 9)
        Num = CCur(Round(MRPVALUE - Val(LBLDISCAMT.Caption) - Val(LBLRETAMT.Caption), 0))
        Print #1, Chr(27) & Chr(72) & Space(3) & AlignLeft("Rupees " & Words_1_all(Num) & " only", 45)
        Print #1, Space(3) & RepeatString("-", 89)
        PRNSPACE = Int(44 - Len("YOU SAVED RS. " & Format(Round(Val(LBLDISCAMT.Caption), 0), "####.00")) / 2)
        If PRNSPACE < 0 Then PRNSPACE = 0
        If Val(LBLDISCAMT.Caption) > 0 Then
'            Print #1, Space(3) & RepeatString("-", 89)
'            PRNSPACE = Int(44 - Len("YOU SAVED RS. " & Format(Round(Val(lbltotaldisc.Caption), 0), "####.00")) / 2)
'            If PRNSPACE < 0 Then PRNSPACE = 0
'            Print #1, Space(PRNSPACE) & Chr(27) & Chr(72) & Chr(14) & Chr(15) & "YOU SAVED RS. " & Format(Round(Val(lbltotaldisc.Caption), 0), "####.00")
'            Print #1, Space(7) & RepeatString("-", 89
            Print #1, Space(PRNSPACE) & Chr(27) & Chr(72) & Chr(14) & Chr(15) & "YOU SAVED RS. " & Format(Round(Val(LBLDISCAMT.Caption), 0), "####.00")
        End If
        
        
        Print #1, Space(3) & RepeatString("-", 89)
        
         'Dim TaxAmt As Double
    Dim taxableamt As Double
    Dim Taxsplit As String
    
    TaxAmt = 0
    taxableamt = 0
    Taxsplit = ""
    Dim RSTtax As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT DISTINCT SALES_TAX From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND SALES_TAX > 0 order by SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        TaxAmt = 0
        taxableamt = 0
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select * From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND SALES_TAX = " & RSTTRXFILE!SALES_TAX & " AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            If OPTDISCPERCENT.Value = True Then
                grdtmp.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * Val(TXTTOTALDISC.text) / 100)
            Else
                grdtmp.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
            End If
            
            taxableamt = Round(taxableamt + Val(grdtmp.Tag) * Val(RSTtax!QTY), 2)
            TaxAmt = Round(TaxAmt + (Val(grdtmp.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY, 2)
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
        'Print #1, Space(7) & "GST " & RSTTRXFILE!SALES_TAX & "%: " & "  TAXABLE VALUE: " & Format(Round(TaxableAmt, 2), "0.00") & "   TAX AMT: " & Format(Round(TaxAmt, 2), "0.00")
        Print #1, Space(3) & "TAXABLE VALUE: " & Format(Round(taxableamt, 2), "0.00") & "   " & "SGST " & RSTTRXFILE!SALES_TAX / 2 & "%: " & " TAX AMT: " & Format(Round(TaxAmt / 2, 2), "0.00") & "   " & "CGST " & RSTTRXFILE!SALES_TAX / 2 & "%: " & " TAX AMT: " & Format(Round(TaxAmt / 2, 2), "0.00")
        Taxsplit = Taxsplit & "GST " & RSTTRXFILE!SALES_TAX & "%: " & "Taxable: " & Format(Round(taxableamt, 2), "0.00") & " Tax: " & Format(Round(TaxAmt, 2), "0.00") & " "
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
    If Val(KFCVAL) > 0 Then Print #1, Space(7) & "Kerala Flood Cess: " & Format(Round(KFCVAL, 2), "0.00")
    'Print #1, Space(3) & "MRP VALUE: " & Format(Round(MRPVALUE, 2), "0.00")
    'Print #1, Space(3) & RepeatString("-", 89)
    
        Print #1, Chr(27) & Chr(67) & Chr(0)
        'Print #1, Space(7) & Chr(27) & Chr(72) & "Pharmacist: " & IIf(IsNull(RSTCOMPANY!VEHICLE) Or RSTCOMPANY!VEHICLE = "", "", RSTCOMPANY!VEHICLE)
        If Not (IsNull(RSTCOMPANY!INV_MSGS) Or RSTCOMPANY!INV_MSGS = "") Then
            PRNSPACE = Int(44 - Len(IIf(IsNull(RSTCOMPANY!INV_MSGS) Or RSTCOMPANY!INV_MSGS = "", "", RSTCOMPANY!INV_MSGS)) / 2)
            Print #1, Space(PRNSPACE) & Chr(27) & Chr(72) & IIf(IsNull(RSTCOMPANY!INV_MSGS) Or RSTCOMPANY!INV_MSGS = "", "", RSTCOMPANY!INV_MSGS)
        End If
        Print #1, Chr(27) & Chr(72) & Space(3) & AlignRight("Time:" & Format(Time, "hh:mm AM/PM"), 88)  '& Space(2) & formAT(LBLTIME.Caption,
        'MessageBox.Show(Date.Now.ToString("h:mm:ss tt"))
        
        PRNSPACE = 40 'Int(44 - Len("**** GET WELL SOON ****") / 2)
        'Print #1, Space(PRNSPACE) & Chr(27) & Chr(72) & "**** GET WELL SOON ****"
        'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** Get Well Soon ****", 40)
'        Print #1, Space(3) & RepeatString("-", 89)
'        Print #1, Chr(27) & Chr(72) & Space(7) & "Certified that the medicines sold as per this bill have been purchased locally from "
'        Print #1, Chr(27) & Chr(72) & Space(7) & "registered dealers who have certified in the related sale bills that such medicines "
'        Print #1, Chr(27) & Chr(72) & Space(7) & "and duly suffered compound tax"
        Print #1, Chr(27) & Chr(72) & Space(3) & "Thank You... E.&.O.E" '   SUBJECT TO ALAPPUZHA JURISDICTION"
        Print #1, Chr(27) & Chr(72) & Space(3) & "Certified that the above particulars are true and correct"
        'Print #1, Chr(27) & Chr(72) & Space(21) & AlignRight("*** WE ACCEPT ALL CREDIT/ DEBIT CARDS ***", 41)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing

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
    
    Close #1 '//Closing the file
    Exit Function

ERRHAND:
    MsgBox err.Description
End Function

Private Sub TXTRETAIL_GotFocus()
'    If M_EDIT = False Then
'        If Val(LBLITEMCOST.Caption) <> 0 Then txtretail.Text = Round(Val(LBLITEMCOST.Caption) + (Val(LBLITEMCOST.Caption) * 10 / 100), 3)
'    End If
    If UCase(txtcategory.text) = "PARDHA" Or UCase(txtcategory.text) = "CLOTHES" Then
        If Val(TXTRETAILNOTAX.text) < 1000 Then
            TXTTAX.text = "5"
        Else
            TXTTAX.text = "12"
        End If
    End If
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    txtretail.SelStart = 0
    txtretail.SelLength = Len(txtretail.text)
    If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
    TxtName1.Enabled = False
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    TxtWarranty.Enabled = True
    TxtWarranty_type.Enabled = True
    txtPrintname.Enabled = True
    TxtSpec.Enabled = True
    Call CHANGEBOXCOLOR(txtretail, True)
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.text) <> 0 And Val(txtretail.text) = 0 Then Exit Sub
            If Val(TXTQTY.text) = 0 And Val(TXTFREE.text) <> 0 And Val(txtretail.text) <> 0 Then
                MsgBox "The Item is issued as free", vbOKOnly, "Sales"
                txtretail.SetFocus
                Exit Sub
            End If
'            If Val(TXTTAX.Text) = 0 Then
'                MsgBox "Please enter the Tax", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            TXTDISC.Enabled = True
            If MDIMAIN.StatusBar.Panels(16).text = "Y" Then
                cmdadd.Enabled = True
                cmdadd.SetFocus
            Else
                TXTDISC.Enabled = True
                TxtDiscAmt.Enabled = True
                If TxtDiscAmt.Visible = True Then
                    TxtDiscAmt.SetFocus
                Else
                    TXTDISC.SetFocus
                End If
            End If
        Case vbKeyEscape
            TXTRETAILNOTAX.Enabled = True
            TXTRETAILNOTAX.SetFocus
        Case 116
            Call FILL_PREVIIOUSRATE
        Case 117
            If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
        Case vbKeyDown
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

Private Sub TXTRETAILNOTAX_LostFocus()
    TXTRETAILNOTAX.text = Format(Val(TXTRETAILNOTAX.text), "0.0000")
    If UCase(txtcategory.text) = "PARDHA" Or UCase(txtcategory.text) = "CLOTHES" Then
        If Val(TXTRETAILNOTAX.text) < 1000 Then
            TXTTAX.text = "5"
        Else
            TXTTAX.text = "12"
        End If
    End If
    ''If lblP_Rate.Caption = "0" Then
    If Val(TXTRETAILNOTAX.text) <> 0 Then
        If OPTTaxMRP.Value = True Then
            txtretail.text = Round(Val(TXTRETAILNOTAX.text) + Val(txtmrpbt.text) * Val(TXTTAX.text) / 100, 4)
        End If
        If OPTVAT.Value = True Then
            txtretail.text = Round(Val(TXTRETAILNOTAX.text) + Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100, 4)
        End If
        If optnet.Value = True Then
            txtretail.text = TXTRETAILNOTAX.text
        End If
        TXTRETAILNOTAX.text = Format(Val(TXTRETAILNOTAX.text), "0.0000")
        If TxtRetailmode.text = "A" Then
            txtcommi.text = Format(Round(Val(txtretaildummy.text), 2), "0.00")
        Else
            txtcommi.text = Format(Round((Val(TXTRETAILNOTAX.text) * Val(txtretaildummy.text) / 100), 2), "0.00")
        End If
    End If
    Call cost_calculate
    Call CHANGEBOXCOLOR(TXTRETAILNOTAX, False)
End Sub

Private Sub TXTRETAILNOTAX_GotFocus()
    TXTRETAILNOTAX.SelStart = 0
    TXTRETAILNOTAX.SelLength = Len(TXTRETAILNOTAX.text)
    If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
    TxtName1.Enabled = False
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    
    cmdadd.Enabled = True
    txtBatch.Enabled = True
    TXTQTY.Enabled = True
    TXTFREE.Enabled = True
    TxtMRP.Enabled = True
    TXTEXPIRY.Enabled = True
    TXTTAX.Enabled = True
    TXTRETAILNOTAX.Enabled = True
    txtretail.Enabled = True
    TXTDISC.Enabled = True
    TxtDiscAmt.Enabled = True
    TxtCessPer.Enabled = True
    TxtCessAmt.Enabled = True
    txtcommi.Enabled = True
    txtretaildummy.Enabled = True
    TxtWarranty.Enabled = True
    TxtWarranty_type.Enabled = True
    txtPrintname.Enabled = True
    TxtSpec.Enabled = True
    Call CHANGEBOXCOLOR(TXTRETAILNOTAX, True)
End Sub

Private Sub TXTRETAILNOTAX_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTRETAILNOTAX.Text) = 0 Then Exit Sub
            txtretail.Enabled = True
            txtretail.SetFocus
        Case vbKeyEscape
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
        Case 116
            Call FILL_PREVIIOUSRATE
        Case 117
            If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
        Case vbKeyDown
            Call CMDADD_Click
    End Select
End Sub

Private Sub TXTRETAILNOTAX_KeyPress(KeyAscii As Integer)
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
'    If Val(txtretail.Text) = 0 Then
'        optnet.value = True
'        TXTTAX.Text = 0
'    End If
    If OPTVAT.Value = False Then TXTTAX.text = 0
    TXTRETAILNOTAX.text = Round(Val(txtretail.text) * 100 / (Val(TXTTAX.text) + 100), 4)
    TXTRETAILNOTAX.text = Format(Val(TXTRETAILNOTAX.text), "0.0000")
    txtretail.text = Format(Val(txtretail.text), "0.0000")
    
    If Val(LBLITEMCOST.Caption) <> 0 Then
        LblProfitPerc.Caption = Round(((Val(TXTRETAILNOTAX.text) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption), 2)
        LblProfitPerc.Caption = Format(Val(LblProfitPerc.Caption), "0.00")
    End If
    
    LblProfitAmt.Caption = Round((Val(TXTRETAILNOTAX.text) - Val(LBLITEMCOST.Caption)) * Val(TXTQTY.text), 2)
    LblProfitAmt.Caption = Format(Val(LblProfitAmt.Caption), "0.00")
    
    LBLNETPROFIT.Caption = Round((Val(txtretail.text) - Val(LBLNETCOST.Caption)) * Val(TXTQTY.text), 2)
    LBLNETPROFIT.Caption = Format(Val(LBLNETPROFIT.Caption), "0.00")
    
    If TxtRetailmode.text = "A" Then
        txtcommi.text = Format(Round(Val(txtretaildummy.text), 2), "0.00")
    Else
        txtcommi.text = Format(Round((Val(TXTRETAILNOTAX.text) * Val(txtretaildummy.text) / 100), 2), "0.00")
    End If
    Call cost_calculate
    Call CHANGEBOXCOLOR(txtretail, False)
    'TXTDISC.Tag = 0
    'TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
    'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag), ".000")
End Sub

Private Sub TxtBillName_GotFocus()
    TxtBillName.SelStart = 0
    TxtBillName.SelLength = Len(TxtBillName.text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    Call CHANGEBOXCOLOR(TxtBillName, True)
End Sub

Private Sub TxtBillName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Trim(TxtBillName.text) = "" Then TxtBillName.text = TXTDEALER.text
            If Trim(TxtBillName.text) = "WS" Then TxtBillName.text = "Cash"
            TxtBillAddress.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select

End Sub

Private Sub TxtBillName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(Chr(KeyAscii))
    End Select
End Sub

Private Sub TxtBillAddress_GotFocus()
    TxtBillAddress.SelStart = 0
    TxtBillAddress.SelLength = Len(TxtBillAddress.text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    TxtBillAddress.BackColor = &H98F3C1
End Sub

Private Function fillcombo()
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If AGNT_FLAG = True Then
        ACT_AGNT.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='911')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    Else
        ACT_AGNT.Close
        ACT_AGNT.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='911')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    End If
    
    Set Me.CMBDISTI.RowSource = ACT_AGNT
    CMBDISTI.ListField = "ACT_NAME"
    CMBDISTI.BoundColumn = "ACT_CODE"
    Screen.MousePointer = vbNormal
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub CMBDISTI_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If CMBDISTI.Text = "" Then Exit Sub
            If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.text <> "" Then
                MsgBox "Select Agent From List", vbOKOnly, "Sale Bill..."
                CMBDISTI.SetFocus
                Exit Sub
            End If
            
'            If Trim(TXTAREA.Text) = "" Then
'                MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'                TXTAREA.SetFocus
'                Exit Sub
'            End If
            
'            If Not IsDate(TXTINVDATE.Text) Then
'                MsgBox "Enter Proper date for Invoice", vbOKOnly, "Sale Bill..."
'                TXTINVDATE.SetFocus
'                Exit Sub
'            End If
'
            'FRMEHEAD.Enabled = False
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
        Case vbKeyEscape
            cmbtype.Enabled = True
            cmbtype.SetFocus
    End Select
End Sub

Private Sub CMBDISTI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyD, Asc("d")
                CMBDISTI.Tag = KeyAscii
            Case vbKeyP, Asc("p")
                If Val(CMBDISTI.Tag) = 68 Or Val(CMBDISTI.Tag) = 100 Or Val(CMBDISTI.Tag) = 85 Or Val(CMBDISTI.Tag) = 117 Then
                    'CMDDUPPURCHASE_Click
                End If
                CMBDISTI.Tag = ""
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcommi_GotFocus()
    If Val(txtcommi.text) = 0 Then txtcommi.text = ""
    txtcommi.SelStart = 0
    txtcommi.SelLength = Len(txtcommi.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub txtcommi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If txtcommi.Text = "" Then Exit Sub
            If Val(txtcommi.text) > Val(txtretail.text) Then
                MsgBox "Commission Rate greater than actual Rate", vbOKOnly, "Sales"
                txtcommi.SetFocus
                Exit Sub
            End If
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Call TXTDISC_LostFocus
            cmdadd.SetFocus
'            TXTDISC.Enabled = False
'            cmdadd.Enabled = True
'            cmdadd.SetFocus
'            'TxtWarranty.Enabled = True
'            'TxtWarranty.SetFocus
        Case vbKeyEscape
            If MDIMAIN.StatusBar.Panels(16).text = "Y" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTDISC.Enabled = True
                TXTDISC.SetFocus
            End If
        Case vbKeyDown
            Call TXTDISC_LostFocus
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtcommi_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcommi_LostFocus()
    txtcommi.text = Format(txtcommi.text, ".000")
End Sub

Private Sub TXTAREA_GotFocus()
    TXTAREA.SelStart = 0
    TXTAREA.SelLength = Len(TXTAREA.text)
End Sub

Private Sub TXTAREA_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
'            If Trim(TXTAREA.Text) = "" Then
'                MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'                'TXTAREA.SetFocus
'                Exit Sub
'            End If
            TxtBillName.SetFocus
            'FRMEHEAD.Enabled = False
            'TxtName1.Enabled = True
            'TxtName1.SetFocus
    End Select
End Sub

Private Sub TXTAREA_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Function FILL_PREVIIOUSRATE()
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILEVAN  WHERE TRX_TYPE ='GI' AND ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILEVAN  WHERE TRX_TYPE ='GI' AND ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    End If
    
    If PHY_PRERATE.RecordCount > 0 Then
        FRMEMAIN.Enabled = False
        fRMEPRERATE.Visible = True
        Set GRDPRERATE.DataSource = PHY_PRERATE
        GRDPRERATE.Columns(0).Caption = "ITEM CODE"
        GRDPRERATE.Columns(1).Caption = "OUTWARD"
        GRDPRERATE.Columns(2).Caption = "DATE"
        GRDPRERATE.Columns(3).Caption = "SOLD QTY"
        GRDPRERATE.Columns(4).Caption = "RATE"
        GRDPRERATE.Columns(5).Caption = "CUSTOMER"
        GRDPRERATE.Columns(6).Caption = "INV NO"
    
        GRDPRERATE.Columns(0).Visible = False
        GRDPRERATE.Columns(1).Width = 3500
        GRDPRERATE.Columns(2).Width = 1300
        GRDPRERATE.Columns(3).Width = 1200
        GRDPRERATE.Columns(4).Width = 1500
        GRDPRERATE.Columns(5).Visible = False
        GRDPRERATE.Columns(6).Width = 1400
        
        
        GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(7).text
    End If
End Function

Private Sub TxtItemcode_GotFocus()
    LBLITEMCOST.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLNETCOST.Caption = ""
    LBLNETPROFIT.Caption = ""
    
    LBLSELPRICE.Caption = ""
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.text)
    grdsales.Enabled = True
    
    fRMEPRERATE.Visible = False
    cmdadd.Enabled = False
    txtBatch.Enabled = False
    TXTQTY.Enabled = False
    TXTFREE.Enabled = False
    TxtMRP.Enabled = False
    TXTEXPIRY.Enabled = False
    TXTTAX.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    txtretail.Enabled = False
    TXTDISC.Enabled = False
    TxtDiscAmt.Enabled = False
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    txtretaildummy.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    TxtSpec.Enabled = False
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    Call CHANGEBOXCOLOR(TXTITEMCODE, True)
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
                MsgBox "Permission Denied", vbOKOnly, "Sales"
                Exit Sub
            End If
'            If MDIMAIN.LBLGSTWRN.Caption = "Y" And Trim(TXTTIN.Text) = "" Then
'                MsgBox "B2C Bill Not allowed", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            If txtBillNo.Tag = "Y" Then
                MsgBox "Cannot modify here", vbOKOnly, "Sales"
                Exit Sub
            End If
            M_STOCK = 0
            'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If Trim(TXTITEMCODE.text) = "" Then
                TxtName1.SetFocus
                Exit Sub
            End If
            'cmddelete.Enabled = False
            TXTQTY.text = ""
            TXTEXPIRY.text = "  /  "
            TXTAPPENDQTY.text = ""
            TXTFREEAPPEND.text = ""
            txtappendcomm.text = ""
            TXTAPPENDTOTAL.text = ""
            txtretail.text = ""
            txtNetrate.text = ""
            txtBatch.text = ""
            TxtWarranty.text = ""
            TxtWarranty_type.text = ""
            TXTRETAILNOTAX.text = ""
            TXTSALETYPE.text = ""
            TXTFREE.text = ""
            optnet.Value = True
            TxtMRP.text = ""
            TXTTAX.text = ""
            TXTDISC.text = ""
            TxtDiscAmt.text = ""
            TxtCessAmt.text = ""
            TxtCessPer.text = ""
            LBLUNBILL.Caption = ""
            txtcommi.text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
            
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL From ITEMMASTVAN  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y')", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL From ITEMMASTVAN  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y')", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount > 0 Then
                TxtCessPer.text = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
                TxtCessAmt.text = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
                TXTDISC.text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
                LBLUNBILL.Caption = IIf(IsNull(grdtmp.Columns(25)), "N", grdtmp.Columns(25))
            End If
            
            If PHY.RecordCount = 0 Then
                Set grdtmp.DataSource = Nothing
                If PHYFLAG = True Then
                    'PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_TAX, LINE_DISC, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, ITEM_SIZE, ITEM_COLOR, BARCODE, REF_NO, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, SALES_PRICE, P_WS, CRTN_PACK, P_CRTN, MRP, TAX_MODE, EXP_DATE  From RTRXFILEVAN  WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_TAX, LINE_DISC, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, BARCODE, REF_NO, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, SALES_PRICE, P_WS, P_VAN, CRTN_PACK, P_CRTN, MRP, TAX_MODE, EXP_DATE, TRX_YEAR, P_LWS, CUST_DISC  From RTRXFILEVAN  WHERE BARCODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                Else
                    PHY.Close
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_TAX, LINE_DISC, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, BARCODE, REF_NO, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, SALES_PRICE, P_WS, P_VAN, CRTN_PACK, P_CRTN, MRP, TAX_MODE, EXP_DATE, TRX_YEAR, P_LWS, CUST_DISC  From RTRXFILEVAN  WHERE BARCODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                End If
                Set grdtmp.DataSource = PHY
                If PHY.RecordCount = 0 Then
                    If PHYFLAG = True Then
                        PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL From ITEMMASTVAN  WHERE BARCODE = '" & Me.TXTITEMCODE.text & "' ", db, adOpenStatic, adLockReadOnly
                        PHYFLAG = False
                    Else
                        PHY.Close
                        PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL From ITEMMASTVAN  WHERE BARCODE = '" & Me.TXTITEMCODE.text & "' ", db, adOpenStatic, adLockReadOnly
                        PHYFLAG = False
                    End If
                    If PHY.RecordCount = 0 Then
                        MsgBox "Item not exists", vbOKOnly, "Sales"
                        Exit Sub
                    Else
                        Set grdtmp.DataSource = PHY
                        TxtMRP.text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
                        Select Case cmbtype.ListIndex
                            Case 1
                                txtretail.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                            Case 2
                                txtretail.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                            Case 3
                                txtretail.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                                If Val(txtretail.text) = 0 Then
                                    txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                                End If
                            Case Else
                                txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        End Select
                        LblPack.text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                        lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                        TXTDISC.text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
                        'TXTEXPIRY.Text = IIf(isdate(grdtmp.Columns(22)),Format(grdtmp.Columns(22), "MM/YY"),"  /  ")
                        TXTITEMCODE.text = grdtmp.Columns(0)
        
'                        For i = 1 To grdsales.Rows - 1
'                            If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
'                                If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
'                                    Exit Sub
'                                Else
'                                    Exit For
'                                End If
'                            End If
'                        Next i
                        
                        item_change = True
                        TXTPRODUCT.text = grdtmp.Columns(1)
                        item_change = False
                        txtPrintname.text = grdtmp.Columns(1)
                        Select Case PHY!check_flag
                            Case "M"
                                OPTTaxMRP.Value = True
                                TXTTAX.text = grdtmp.Columns(4)
                                TXTSALETYPE.text = "2"
                            Case "V"
                                OPTVAT.Value = True
                                TXTSALETYPE.text = "2"
                                TXTTAX.text = grdtmp.Columns(4)
                            Case Else
                                TXTSALETYPE.text = "2"
                                optnet.Value = True
                                TXTTAX.text = "0"
                        End Select
                        
                        If MDIMAIN.lblPriceSplit.Caption = "Y" Then GoTo SKIP_BATCH2
                        Set RSTBATCH = New ADODB.Recordset
                        Select Case cmbtype.ListIndex
                            Case 1
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            Case 2
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_VAN >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            Case 3
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, MRP, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            Case Else
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_RETAIL >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                
                        End Select
                        If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                            If BATCH_DISPLAY = True Then
                                If RSTBATCH.RecordCount >= 1 Then
                                    Call FILL_BATCHGRID
                                    RSTBATCH.Close
                                    Set RSTBATCH = Nothing
                                    Exit Sub
                                ElseIf RSTBATCH.RecordCount = 1 Then
                                    'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                                    txtBatch.text = IIf(IsNull(RSTBATCH!REF_NO), "", RSTBATCH!REF_NO)
                                End If
                            Else
                                If RSTBATCH.RecordCount > 1 Then
                                    Call FILL_BATCHGRID
                                    RSTBATCH.Close
                                    Set RSTBATCH = Nothing
                                    Exit Sub
                                ElseIf RSTBATCH.RecordCount = 1 Then
                                    'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                                    txtBatch.text = IIf(IsNull(RSTBATCH!REF_NO), "", RSTBATCH!REF_NO)
                                End If
                            End If
                        End If
                        RSTBATCH.Close
                        Set RSTBATCH = Nothing
SKIP_BATCH2:
                        'Call CONTINUE
                        TxtName1.Enabled = False
                        TXTPRODUCT.Enabled = False
                        TXTITEMCODE.Enabled = False
                        If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                            TXTQTY.text = 1
                            txtretail.Enabled = True
                            txtretail.SetFocus
                        Else
                            TXTQTY.Enabled = True
                            TXTQTY.SetFocus
                        End If
                    End If
                    Exit Sub
                End If
                If IsDate(grdtmp.Columns(28)) Then
                    If DateDiff("d", Date, grdtmp.Columns(28)) < 0 Then
                        MsgBox "Item Expired....", vbOKOnly, "BILL.."
                        Exit Sub
                    End If
                    If DateDiff("d", Date, grdtmp.Columns(28)) < 60 Then
                        If (MsgBox("Expiry < " & Val(DateDiff("d", Date, grdtmp.Columns(28))) & "Days", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then Exit Sub
                    End If
                End If
                
                SERIAL_FLAG = True
                lblactqty.Caption = grdtmp.Columns(2)
                lblbarcode.Caption = grdtmp.Columns(15)
                TXTITEMCODE.text = grdtmp.Columns(0)
                TXTEXPIRY.text = IIf(IsDate(grdtmp.Columns(28)), Format(grdtmp.Columns(28), "MM/YY"), "  /  ")
                
                If MDIMAIN.lblPriceSplit.Caption = "Y" Then GoTo SKIP_BATCH3
                Set RSTBATCH = New ADODB.Recordset
                Select Case cmbtype.ListIndex
                    Case 1
                        RSTBATCH.Open "Select DISTINCT BARCODE, P_WS From RTRXFILEVAN WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    Case 2
                        RSTBATCH.Open "Select DISTINCT BARCODE, P_VAN From RTRXFILEVAN WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    Case 3
                        RSTBATCH.Open "Select DISTINCT BARCODE, MRP From RTRXFILEVAN WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    Case Else
                        RSTBATCH.Open "Select DISTINCT BARCODE, P_RETAIL From RTRXFILEVAN WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                End Select
                If BATCH_DISPLAY = True Then
                    If RSTBATCH.RecordCount >= 1 Then
                        Call FILL_BATCHGRID2
                        RSTBATCH.Close
                        Set RSTBATCH = Nothing
                        Exit Sub
                    End If
                Else
                    If RSTBATCH.RecordCount > 1 Then
                        Call FILL_BATCHGRID2
                        RSTBATCH.Close
                        Set RSTBATCH = Nothing
                        Exit Sub
                    End If
                End If
                RSTBATCH.Close
                Set RSTBATCH = Nothing
SKIP_BATCH3:
                
                TXTITEMCODE.text = grdtmp.Columns(0)
                item_change = True
                TXTPRODUCT.text = grdtmp.Columns(1)
                item_change = False
                TXTUNIT.text = "1" 'grdtmp.Columns(4)
                TxtMRP.text = IIf(IsNull(grdtmp.Columns(26)), "", grdtmp.Columns(26))
                If grdtmp.Columns(6) = "A" Then
                    txtretaildummy.text = IIf(IsNull(grdtmp.Columns(8)), "", grdtmp.Columns(8))
                    TxtRetailmode.text = "A"
                    lblcomtype.Caption = "`"
                Else
                    txtretaildummy.text = IIf(IsNull(grdtmp.Columns(7)), "", grdtmp.Columns(7))
                    TxtRetailmode.text = "P"
                    lblcomtype.Caption = "%"
                End If
                If Val(txtretaildummy.text) = 0 Then lblcomtype.Caption = "`"
                TXTEXPIRY.text = IIf(IsDate(grdtmp.Columns(22)), Format(grdtmp.Columns(22), "MM/YY"), "  /  ")
                lblunit.text = grdtmp.Columns(12)
                TxtWarranty.text = grdtmp.Columns(13)
                TxtWarranty_type.text = grdtmp.Columns(14)
                'txtbarcode.Text = grdtmp.Columns(15)
                txtBatch.text = grdtmp.Columns(16)
                TXTVCHNO.text = grdtmp.Columns(17)
                TXTLINENO.text = grdtmp.Columns(18)
                TXTTRXTYPE.text = grdtmp.Columns(19)
                TrxRYear.text = grdtmp.Columns(29)
                TXTDISC.text = IIf(IsNull(grdtmp.Columns(31)), "", grdtmp.Columns(31))
                LBLITEMCOST.Caption = grdtmp.Columns(20)
                LblPack.text = IIf(IsNull(grdtmp.Columns(11)) Or Val(grdtmp.Columns(11)) = 0, "1", grdtmp.Columns(11))
                lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(11)) Or Val(grdtmp.Columns(11)) = 0, "1", grdtmp.Columns(11))
                Select Case cmbtype.ListIndex
                    Case 1
                        txtretail.text = IIf(IsNull(grdtmp.Columns(22)), "", Val(grdtmp.Columns(22)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(22)), "", Val(grdtmp.Columns(22)))
                    Case 2
                        txtretail.text = IIf(IsNull(grdtmp.Columns(23)), "", Val(grdtmp.Columns(23)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(23)), "", Val(grdtmp.Columns(23)))
                    Case 3
                        txtretail.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                        If Val(txtretail.text) = 0 Then
                            txtretail.text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                            'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                        End If
                    Case Else
                        txtretail.text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                End Select
                LBLSELPRICE.Caption = Val(txtretail.text)
                lblretail.Caption = IIf(IsNull(grdtmp.Columns(5)), "", grdtmp.Columns(5))
                lblwsale.Caption = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
                lblvan.Caption = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
                LBLMRP.Caption = IIf(IsNull(grdtmp.Columns(26)), "", grdtmp.Columns(26))
                lblcase.Caption = IIf(IsNull(grdtmp.Columns(25)), "", grdtmp.Columns(25))
                lblLWPrice.Caption = IIf(IsNull(grdtmp.Columns(30)), "", grdtmp.Columns(30))
                lblcrtnpack.Caption = IIf(IsNull(grdtmp.Columns(24)), "", grdtmp.Columns(24))
                
                If TxtRetailmode.text = "A" Then
                    txtcommi.text = Format(Round(Val(txtretaildummy.text), 2), "0.00")
                Else
                    txtcommi.text = Format(Round((Val(TXTRETAILNOTAX.text) * Val(txtretaildummy.text) / 100), 2), "0.00")
                End If
                
                Dim RSTtax As ADODB.Recordset
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                With RSTtax
                    If Not (.EOF And .BOF) Then
                        Select Case grdtmp.Columns(9)
                            Case "M"
                                OPTTaxMRP.Value = True
                                TXTTAX.text = grdtmp.Columns(3)
                                TXTSALETYPE.text = "2"
                            Case "V"
                                If (!Category = "GENERAL" And !REMARKS = "1") Then
                                    OPTTaxMRP.Value = True
                                    TXTSALETYPE.text = "1"
                                Else
                                    OPTVAT.Value = True
                                    TXTSALETYPE.text = "2"
                                End If
                                TXTTAX.text = IIf(IsNull(RSTtax!SALES_TAX), "", RSTtax!SALES_TAX)
                                If Val(TXTDISC.text) = 0 Then TXTDISC.text = IIf(IsNull(RSTtax!CUST_DISC), "", RSTtax!CUST_DISC)
                                txtcategory.text = IIf(IsNull(RSTtax!Category), "", RSTtax!Category)
                                If UCase(txtcategory.text) = "PARDHA" Or UCase(txtcategory.text) = "CLOTHES" Then
                                    If Val(TXTRETAILNOTAX.text) < 1000 Then
                                        TXTTAX.text = "5"
                                    Else
                                        TXTTAX.text = "12"
                                    End If
                                End If
                            Case Else
                                TXTSALETYPE.text = "2"
                                optnet.Value = True
                                TXTTAX.text = "0"
                        End Select
                    Else
                        optnet.Value = True
                        TXTTAX.text = "0"
                    End If
                End With
                RSTtax.Close
                Set RSTtax = Nothing
'                TXTITEMCODE.Enabled = False
'                TXTPRODUCT.Enabled = False
'                TXTQTY.Enabled = True
'                TXTQTY.SetFocus
                If Val(LblPack.text) > 1 Or cmbtype.ListIndex <> 0 Then
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                    TxtName1.Enabled = False
                    TXTITEMCODE.Enabled = False
                    TXTPRODUCT.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                    Exit Sub
                Else
                    TXTQTY.text = "1.00"
                    Call TXTQTY_LostFocus
                    If MDIMAIN.LblKFCNet.Caption <> "N" Then
                        Call txtNetrate_LostFocus
                        Call TXTDISC_LostFocus
                        Call CMDADD_Click
                    Else
                        If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                            Call TXTRETAILNOTAX_LostFocus
                        Else
                            If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
                                Call TXTRETAILNOTAX_LostFocus
                            ElseIf Val(TxtMRP.text) <> 0 And Val(Round(Val(TxtMRP.text), 2)) = Val(Round(Val(txtretail.text), 2)) And mrpplus = False Then
                                Call txtNetrate_LostFocus
                            Else
                                Call TXTRETAIL_LostFocus
                            End If
                        End If
                            Call TXTDISC_LostFocus
                            Call CMDADD_Click
                        End If
                    End If
                    Exit Sub
            End If
            
            SERIAL_FLAG = False
            lblactqty.Caption = ""
            lblbarcode.Caption = ""
            If PHY.RecordCount = 1 Then
                TxtMRP.text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
                Select Case cmbtype.ListIndex
                    Case 1
                        txtretail.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    Case 2
                        txtretail.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    Case 3
                        txtretail.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        If Val(txtretail.text) = 0 Then
                            txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                            TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        End If
                    Case Else
                        txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                End Select
                LblPack.text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                TXTDISC.text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
                'TXTEXPIRY.Text = IIf(isdate(grdtmp.Columns(22)),Format(grdtmp.Columns(22), "MM/YY"),"  /  ")
                TXTITEMCODE.text = grdtmp.Columns(0)

                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.text) Then
                        If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                Next i
                
                item_change = True
                TXTPRODUCT.text = grdtmp.Columns(1)
                item_change = False
                txtPrintname.text = grdtmp.Columns(1)
                Select Case PHY!check_flag
                    Case "M"
                        OPTTaxMRP.Value = True
                        TXTTAX.text = grdtmp.Columns(4)
                        TXTSALETYPE.text = "2"
                    Case "V"
                        OPTVAT.Value = True
                        TXTSALETYPE.text = "2"
                        TXTTAX.text = grdtmp.Columns(4)
                    Case Else
                        TXTSALETYPE.text = "2"
                        optnet.Value = True
                        TXTTAX.text = "0"
                End Select
                
                If MDIMAIN.lblPriceSplit.Caption = "Y" Then GoTo SKIP_BATCH4
                Set RSTBATCH = New ADODB.Recordset
                Select Case cmbtype.ListIndex
                    Case 1
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 2
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_VAN >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 3
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, MRP, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case Else
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, EXP_DATE From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_RETAIL >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, REF_NO From RTRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        
                End Select
                If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                    If BATCH_DISPLAY = True Then
                        If RSTBATCH.RecordCount >= 1 Then
                            Call FILL_BATCHGRID
                            RSTBATCH.Close
                            Set RSTBATCH = Nothing
                            Exit Sub
                        ElseIf RSTBATCH.RecordCount = 1 Then
                            'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                            txtBatch.text = IIf(IsNull(RSTBATCH!REF_NO), "", RSTBATCH!REF_NO)
                        End If
                    Else
                        If RSTBATCH.RecordCount > 1 Then
                            Call FILL_BATCHGRID
                            RSTBATCH.Close
                            Set RSTBATCH = Nothing
                            Exit Sub
                        ElseIf RSTBATCH.RecordCount = 1 Then
                            'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                            txtBatch.text = IIf(IsNull(RSTBATCH!REF_NO), "", RSTBATCH!REF_NO)
                        End If
                    End If
                End If
                RSTBATCH.Close
                Set RSTBATCH = Nothing
SKIP_BATCH4:
                Call CONTINUE
            Else
                Call FILL_ITEMGRID
                Exit Sub
            End If
JUMPNONSTOCK:
            TxtName1.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                TXTQTY.text = 1
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
            End If
        Case vbKeyEscape
            TXTITEMCODE.Enabled = False
            TxtName1.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            Exit Sub
'            TxtName1.Enabled = False
'            TXTSLNO.Enabled = True
'            TXTSLNO.SetFocus
'            Exit Sub
            If grdsales.rows > 1 Then
                CMDPRINT.Enabled = True
                CmdPrintA5.Enabled = True
                cmdRefresh.Enabled = True
                CmdPrintA5.SetFocus
            Else
                FRMEHEAD.Enabled = True
                TxtCode.Enabled = True
                TXTDEALER.Enabled = True
                TXTDEALER.SetFocus
            End If
    End Select
    Exit Sub
ERRHAND:
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

Function FILL_BATCHGRID()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME From RTRXFILEVAN  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME From RTRXFILEVAN  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "Serial No."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "VCH No"
    GRDPOPUP.Columns(3).Caption = "Line No"
    GRDPOPUP.Columns(4).Caption = "Trx Type"
    GRDPOPUP.Columns(7).Caption = "" '"Warranty"
    GRDPOPUP.Columns(8).Caption = ""
    GRDPOPUP.Columns(9).Caption = "RT"
    GRDPOPUP.Columns(10).Caption = "WS"
    GRDPOPUP.Columns(11).Caption = "VN"
    GRDPOPUP.Columns(25).Caption = "Expiry"
    'GRDPOPUP.Columns(32).Caption = "Godown"
    
    GRDPOPUP.Columns(0).Width = 3000
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 0
    GRDPOPUP.Columns(3).Width = 0
    GRDPOPUP.Columns(4).Width = 0
    GRDPOPUP.Columns(5).Width = 0
    GRDPOPUP.Columns(6).Width = 0
    GRDPOPUP.Columns(7).Width = 0
    GRDPOPUP.Columns(8).Width = 0
    GRDPOPUP.Columns(9).Width = 1000
    GRDPOPUP.Columns(10).Width = 1000
    GRDPOPUP.Columns(11).Width = 1000
    GRDPOPUP.Columns(12).Width = 0
    GRDPOPUP.Columns(13).Width = 0
    GRDPOPUP.Columns(14).Width = 0
    GRDPOPUP.Columns(15).Width = 0
    GRDPOPUP.Columns(16).Width = 0
    GRDPOPUP.Columns(17).Width = 0
    GRDPOPUP.Columns(18).Width = 0
    GRDPOPUP.Columns(19).Width = 0
    GRDPOPUP.Columns(20).Width = 0
    GRDPOPUP.Columns(21).Width = 0
    GRDPOPUP.Columns(22).Width = 0
    GRDPOPUP.Columns(23).Width = 0
    GRDPOPUP.Columns(24).Width = 0
    GRDPOPUP.Columns(25).Width = 1200
    GRDPOPUP.Columns(26).Width = 0
    GRDPOPUP.Columns(27).Width = 0
    GRDPOPUP.Columns(28).Width = 0
    GRDPOPUP.Columns(29).Width = 0
    GRDPOPUP.Columns(30).Width = 0
    GRDPOPUP.Columns(31).Width = 0
    
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
    
    
End Function

Function FILL_PREVIIOUSRATE2()
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        'PHY_PRERATE.Open "Select TOP 10 ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILEVAN  WHERE TRX_TYPE ='GI' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, LINE_DISC, VCH_NO, ITEM_NAME, M_USER_ID  From TRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC LIMIT 10", db, adOpenForwardOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        'PHY_PRERATE.Open "Select TOP 10 ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILEVAN  WHERE TRX_TYPE ='GI' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, LINE_DISC, VCH_NO, ITEM_NAME, M_USER_ID  From TRXFILEVAN WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC LIMIT 10", db, adOpenForwardOnly
        PRERATE_FLAG = False
    End If
    
    If PHY_PRERATE.RecordCount > 0 Then
        'FRMEMAIN.Enabled = False
        fRMEPRERATE.Visible = True
        Set GRDPRERATE.DataSource = PHY_PRERATE
        GRDPRERATE.Columns(0).Caption = "ITEM CODE"
        GRDPRERATE.Columns(1).Caption = "OUTWARD"
        GRDPRERATE.Columns(2).Caption = "DATE"
        GRDPRERATE.Columns(3).Caption = "SOLD QTY"
        GRDPRERATE.Columns(4).Caption = "RATE"
        GRDPRERATE.Columns(5).Caption = "NET RATE"
        GRDPRERATE.Columns(6).Caption = "Disc%"
        GRDPRERATE.Columns(7).Caption = "INV NO"
    
        GRDPRERATE.Columns(0).Visible = False
        GRDPRERATE.Columns(1).Width = 2500
        GRDPRERATE.Columns(2).Width = 1100
        GRDPRERATE.Columns(3).Width = 1100
        GRDPRERATE.Columns(4).Width = 1100
        GRDPRERATE.Columns(5).Width = 1100
        GRDPRERATE.Columns(6).Width = 1100
        GRDPRERATE.Columns(7).Width = 1200
        GRDPRERATE.Columns(8).Visible = False
        GRDPRERATE.Columns(9).Visible = False
        
        'GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(8).text
    End If
End Function

Private Sub TxtPhone_GotFocus()
    TxtPhone.SelStart = 0
    TxtPhone.SelLength = Len(TxtPhone.text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    Call CHANGEBOXCOLOR(TxtPhone, True)
End Sub

Private Sub TxtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
'            CMBDISTI.SetFocus
            'FRMEHEAD.Enabled = False
            TXTITEMCODE.Enabled = True
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            If CMBDISTI.VisibleCount = 0 Then
                If TXTITEMCODE.Visible = True Then
                    If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
                Else
                    If TxtName1.Enabled = True Then TxtName1.SetFocus
                End If
            Else
                CMBDISTI.SetFocus
            End If
        Case vbKeyEscape
            TxtVehicle.SetFocus
    End Select

End Sub

Private Sub TxtPhone_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTTYPE_Change()
    If Val(TXTTYPE.text) = 1 Then
        cmbtype.ListIndex = 0
    ElseIf Val(TXTTYPE.text) = 2 Then
        cmbtype.ListIndex = 1
    ElseIf Val(TXTTYPE.text) = 3 Then
        cmbtype.ListIndex = 2
    ElseIf Val(TXTTYPE.text) = 4 Then
        cmbtype.ListIndex = 3
    Else
        cmbtype.ListIndex = -1
    End If
End Sub

Private Sub TXTTYPE_LostFocus()
    If cmbtype.ListIndex = -1 Then
        'MsgBox "Select Bill Type from the List", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
     If Val(TXTTYPE.text) = 0 Or Val(TXTTYPE.text) > 4 Then
        MsgBox "Enter Bill Type", vbOKOnly, "Sales"
        TXTTYPE.Enabled = True
        TXTTYPE.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 0 And Val(TXTTYPE.text) <> 1 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 1 And Val(TXTTYPE.text) <> 2 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 2 And Val(TXTTYPE.text) <> 3 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 3 And Val(TXTTYPE.text) <> 4 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
    Call CHANGEBOXCOLOR(TXTTYPE, False)
End Sub

Private Sub TxtVehicle_GotFocus()
    'If Trim(TxtVehicle.Text) = "" Then TxtVehicle.Text = "KL-04-N-8931"
    TxtVehicle.SelStart = 0
    TxtVehicle.SelLength = Len(TxtVehicle.text)
    Call CHANGEBOXCOLOR(TxtVehicle, True)
End Sub

Private Sub TxtVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
                TxtPhone.SetFocus
            Else
                TXTITEMCODE.Enabled = True
                TxtName1.Enabled = True
                TXTPRODUCT.Enabled = True
                If TXTITEMCODE.Visible = True Then
                    If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
                Else
                    If TxtName1.Enabled = True Then TxtName1.SetFocus
                End If
'                FRMEHEAD.Enabled = False
'                TxtName1.Enabled = True
'                TxtName1.SetFocus
            End If
        Case vbKeyEscape
            cmbtype.SetFocus
    End Select

End Sub

Private Sub TxtVehicle_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtVehicle_LostFocus()
    Call CHANGEBOXCOLOR(TxtVehicle, False)
End Sub

Private Sub TxtWarranty_GotFocus()
    Call CHANGEBOXCOLOR(TxtWarranty, True)
End Sub

Private Sub TxtWarranty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.text) = 0 Then
                cmdadd.Enabled = True
                cmdadd.SetFocus
            Else
                TxtWarranty_type.SetFocus
            End If
        Case vbKeyEscape
            If MDIMAIN.StatusBar.Panels(16).text = "Y" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTDISC.Enabled = True
                TXTDISC.SetFocus
            End If
        Case vbKeyDown
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

Private Sub TxtWarranty_LostFocus()
    Call CHANGEBOXCOLOR(TxtWarranty, False)
End Sub

Private Sub TxtWarranty_type_GotFocus()
    Call CHANGEBOXCOLOR(TxtWarranty_type, True)
End Sub

Private Sub TxtWarranty_type_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.text) <> 0 And Trim(TxtWarranty_type.text) = "" Then
                MsgBox "Please enter Period for Warranty", , "Sales"
                TxtWarranty_type.SetFocus
                Exit Sub
            End If
            If Val(TxtWarranty.text) = 0 Then TxtWarranty_type.text = ""
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            TxtWarranty.SetFocus
        Case vbKeyDown
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtWarranty_type_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z")
            KeyAscii = Asc(Chr(KeyAscii))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Function checklastbill()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ERRHAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'GI'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    LBLBILLNO.Caption = Val(txtBillNo.text)
    
Exit Function
ERRHAND:
    MsgBox err.Description
End Function

'Private Function ReportGeneratION_estimate()
'
'    Dim RSTCOMPANY As ADODB.Recordset
'    Dim RSTTRXFILE As ADODB.Recordset
'    Dim Num As Currency
'    Dim SN As Integer
'    Dim i As Long
'    SN = 0
'
'    On Error GoTo CLOSEFILE
'    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
'
'CLOSEFILE:
'    If Err.Number = 55 Then
'        Close #1
'        Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
'    End If
'    On Error GoTo errhand
'    '//Create Report Heading
'    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
'    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
'
'
'    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
'            Chr (106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
'    'Print #1, Chr(13)
'        Print #1, AlignLeft("ESTIMATE", 25)
'        Print #1, RepeatString("-", 80)
'        Print #1, AlignLeft("Sl", 2) & Space(1) & _
'                AlignLeft("Comm Code", 14) & Space(1) & _
'                AlignLeft("Description", 35) & _
'                AlignLeft("Qty", 4) & Space(3) & _
'                AlignLeft("Rate", 10) & Space(3) & _
'                AlignLeft("Amount", 12) '& _
'                Chr (27) & Chr(72) '//Bold Ends
'
'        Print #1, RepeatString("-", 80)
'
'        For i = 1 To grdsales.Rows - 1
'            If Trim(grdsales.TextMatrix(i, 33)) = "" Then
'                grdsales.Tag = Trim(grdsales.TextMatrix(i, 2))
'            Else
'                grdsales.Tag = Trim(grdsales.TextMatrix(i, 33))
'            End If
'
'            Print #1, AlignLeft(Val(i), 3) & _
'                Space(15) & AlignLeft(grdsales.Tag, 34) & _
'                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 4) & _
'                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 9) & _
'                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 13) '& _
'                Chr (27) & Chr(72) '//Bold Ends
'        Next i
'
'        Print #1, AlignRight("-------------", 80)
'        If Val(LBLDISCAMT.Caption) <> 0 Then
'            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
'            Print #1, AlignRight("DISC AMOUNT ", 65) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
'        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
'            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
'        End If
'        If Val(Txthandle.Text) <> 0 Then
'            Print #1, AlignRight("HANDLING CHARGE ", 65) & AlignRight((Format(Txthandle.Text, "####.00")), 12)
'        End If
'        If Val(TxtFrieght.Text) <> 0 Then
'            Print #1, AlignRight("FRIEGHT CHARGE  ", 65) & AlignRight((Format(Txthandle.Text, "####.00")), 12)
'        End If
'
'        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 57) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
'        Print #1, AlignRight("Round off ", 65) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
'        Print #1, Chr(13)
'        Print #1, AlignRight("NET AMOUNT ", 65) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 12)
'        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
'        Num = CCur(Round(lblnetamount.Caption, 0))
'        Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & ")", 80)
'        Print #1, RepeatString("-", 80)
'        'Print #1, Chr(27) & Chr(71) & Chr(0)
''        If Trim(TXTTIN.Text) <> "" Then
''            Print #1, "Certified that all the particulars shown in the above Tax Invoice are true and correct"
''            Print #1, "and that my/our Registration under KVAT ACT 2003 is valid as on the date of this bill"
''            Print #1, RepeatString("-", 80)
''        End If
'        'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** THANK YOU ****", 40)
'
'
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'
'    Close #1 '//Closing the file
'    Exit Function
'
'errhand:
'    MsgBox Err.Description
'End Function

Private Function ReportGeneratION_vpestimate(Op_Bal As Double, RCPT_AMT As Double)
    
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
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
    '//chr(27) & chr(42) & chr(1) - for Enlarge letter and bold


'    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(55) & Chr(27) & Chr(55)
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    Print #1,
    Print #1,
    
    Dim BIL_PRE, BILL_SUF, INVMSG As String
    Dim PRNSPACE As Single
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        INVMSG = IIf(IsNull(RSTCOMPANY!INV_MSGS), "", RSTCOMPANY!INV_MSGS)
        
        PRNSPACE = Int(54 - Len(IIf(IsNull(RSTCOMPANY!COMP_NAME) Or RSTCOMPANY!COMP_NAME = "", 1, RSTCOMPANY!COMP_NAME)) / 2)
        If PRNSPACE < 0 Then PRNSPACE = 0
        Print #1, Chr(27) & Chr(72) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(72) & Chr(10) & _
              Space(PRNSPACE) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 50) & _
              Chr(27) & Chr(72)
              
        PRNSPACE = Int(62 - Len(IIf(IsNull(RSTCOMPANY!Address) Or RSTCOMPANY!Address = "", 1, RSTCOMPANY!Address)) / 2)
        If PRNSPACE < 0 Then PRNSPACE = 0
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & Space(PRNSPACE) & RSTCOMPANY!Address
        
        If Not IsNull(RSTCOMPANY!HO_NAME) Or RSTCOMPANY!HO_NAME = "" Then
            PRNSPACE = Int(62 - Len(IIf(IsNull(RSTCOMPANY!HO_NAME) Or RSTCOMPANY!HO_NAME = "", 1, RSTCOMPANY!HO_NAME)) / 2)
            If PRNSPACE < 0 Then PRNSPACE = 0
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & Space(PRNSPACE) & RSTCOMPANY!HO_NAME
        End If
        
        PRNSPACE = Int(62 - (7 + Len(IIf(IsNull(RSTCOMPANY!TEL_NO) Or RSTCOMPANY!TEL_NO = "", 0, RSTCOMPANY!TEL_NO)) + Len(IIf(IsNull(RSTCOMPANY!FAX_NO) Or RSTCOMPANY!FAX_NO = "", 0, RSTCOMPANY!FAX_NO))) / 2)
        If PRNSPACE < 0 Then PRNSPACE = 0
        Print #1, Space(7) & Space(PRNSPACE) & "Phone: " & RSTCOMPANY!TEL_NO & ", " & RSTCOMPANY!FAX_NO
        
        PRNSPACE = Int(62 - (Len(IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", 1, RSTCOMPANY!CST)) + 6) / 2)
        'PRNSPACE = PRNSPACE - 6
        If PRNSPACE < 0 Then PRNSPACE = 0
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & Space(PRNSPACE) & "GSTIN:" & RSTCOMPANY!CST
                
        If MDIMAIN.LBLSTATE.Caption = "32" Then
            Print #1, "State: Kerala (32 - KL)"
        Else
            Print #1, "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption)
        End If
        Print #1, RepeatString("-", 125)
        'Print #1,
        '''Print #1,  "TIN No. " & RSTCOMPANY!KGST
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    'Print #1, Space(31) & "The KVAT Rules 2005"
    PRNSPACE = Int(27)
    Print #1, Chr(27) & Chr(72) & Chr(10) & _
               Chr(14) & Chr(15) & Space(PRNSPACE) & "TAX INVOICE" & _
              Chr(27) & Chr(72)
    'Print #1, Space(57) & "TAX INVOICE"
    'Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7)
    Print #1, RepeatString("-", 125)
    'Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & RepeatString("-", 125)
    'Print #1, RepeatString("-", 125)
'    If lblcredit.Caption = 0 Then
'        Print #1, Space(32) & AlignLeft("CASH SALE", 30)
'    Else
'        Print #1, Space(32) & AlignLeft("CREDIT SALE", 30)
'    End If
    'Print #1, RepeatString("-", 125)
    Print #1, "D.N. NO & Date" & Space(10) & "P.O. NO. & Date" & Space(15) & "D.Doc.NO & Date" & Space(15) & "Del Terms" & Space(15) & "Veh. No"
    Print #1,
    Print #1, RepeatString("-", 125)
    'Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & AlignLeft("INVOICE FORM 8H", 16)

    'If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, LBLDATE.Caption)
    Print #1, "Bill No. " & BIL_PRE & Trim(txtBillNo.text) & BILL_SUF & Space(2) & AlignRight("Date:" & TXTINVDATE.text, 110) '& Space(2) & LBLTIME.Caption
    Print #1, "TO: " & TxtBillName.text
    If Trim(TxtBillAddress.text) <> "" Then Print #1, TxtBillAddress.text
    If Trim(TxtPhone.text) <> "" Then Print #1, "Phone: " & TxtPhone.text
    If Trim(TXTTIN.text) <> "" Then Print #1, "GSTIN: " & TXTTIN.text
    'LBLDATE.Caption = Date

   ' Print #1, Chr(27) & Chr(72) &  "Salesman: CS"

    Print #1, RepeatString("-", 125)
    Print #1, AlignLeft("Sl", 3) & _
            AlignLeft("Description", 50) & _
            AlignLeft("HSN", 8) & Space(0) & _
            AlignRight("Qty", 8) & Space(0) & _
            AlignRight("Rate", 9) & Space(0) & _
            AlignRight("Disc", 8) & Space(0) & _
            AlignRight("CGST%", 8) & Space(0) & _
            AlignRight("SGST%", 8) & Space(0) & _
            AlignRight("GST Amt", 9) & Space(0) & _
            AlignRight("Amount", 13) '& _
            Chr (27) & Chr(72) '//Bold Ends

    Print #1, RepeatString("-", 125)
    Dim TotalTax, TaxAmt, DISCAMT, Totaldiscamt, KFCVAL As Double
    Dim HSNCODE As String
    Dim RSTHSNCODE As ADODB.Recordset
    TaxAmt = 0
    TotalTax = 0
    DISCAMT = 0
    KFCVAL = 0
    Totaldiscamt = 0
    For i = 1 To grdsales.rows - 1
        DISCAMT = (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 8)) / 100) * Val(grdsales.TextMatrix(i, 3))
        'KFCVAL = KFCVAL + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 47)) / 100) * Val(grdsales.TextMatrix(i, 3))
        Totaldiscamt = Totaldiscamt + DISCAMT
        If Val(creditbill.grdsales.TextMatrix(i, 9)) > 0 Then
            TaxAmt = Round((Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3)) - DISCAMT) * Val(grdsales.TextMatrix(i, 9)) / 100, 2)
        End If
        TotalTax = TotalTax + TaxAmt
        
        Set RSTHSNCODE = New ADODB.Recordset
        RSTHSNCODE.Open "SELECT * From ITEMMASTVAN WHERE ITEM_CODE = '" & creditbill.grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTHSNCODE.EOF And RSTHSNCODE.BOF) Then
            HSNCODE = IIf(IsNull(RSTHSNCODE!REMARKS), "", RSTHSNCODE!REMARKS)
        Else
            HSNCODE = ""
        End If
        RSTHSNCODE.Close
        Set RSTHSNCODE = Nothing
        
        If Trim(grdsales.TextMatrix(i, 33)) = "" Then
            grdsales.Tag = Trim(grdsales.TextMatrix(i, 2))
        Else
            grdsales.Tag = Trim(grdsales.TextMatrix(i, 33))
        End If
        Print #1, AlignLeft(grdsales.TextMatrix(i, 0), 3) & Space(0) & _
            AlignLeft(Mid(grdsales.Tag, 1, 50), 50) & Space(0) & _
            AlignLeft(HSNCODE, 8) & _
            AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 8) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)), 2), "0.00"), 9) & _
            AlignRight(Format(Round(DISCAMT, 2), "0.00"), 8) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 9)) / 2, 2), "0.00"), 8) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 9)) / 2, 2), "0.00"), 8) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 9)) / 100, 2), "0.00"), 9) & _
            AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 13) '& _
            Chr (27) & Chr(72) '//Bold Ends
        If MDIMAIN.LBLSPACE.Caption = "Y" Then Print #1, Chr(13)
        
    Next i
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, RepeatString("-", 125)
    Print #1, "TOTAL:" & AlignRight(Format(Round(Totaldiscamt, 2), "0.00"), 80) & AlignRight(Format(Round(TotalTax / 2, 2), "0.00"), 8) & AlignRight(Format(Round(TotalTax / 2, 2), "0.00"), 8) & AlignRight(Format(Round(TotalTax, 2), "0.00"), 9) & AlignRight(Format(Round(Val(LBLTOTAL.Caption), 2), "0.00"), 13)
    Print #1, RepeatString("-", 125)
    If Val(KFCVAL) > 0 Then Print #1, Space(7) & "Kerala Flood Cess: " & Format(Round(KFCVAL, 2), "0.00")
    
    
'    If TotalTax > 0 Then
'        Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(2) & AlignLeft("CGST Tax Amt: " & Format(Round(TotalTax / 2, 2), "0.00"), 48)
'        Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(2) & AlignLeft("SGST Tax Amt: " & Format(Round(TotalTax / 2, 2), "0.00"), 48)
'        Print #1, Chr(27) & Chr(72) & Chr(14) & Chr(15) & Space(2) & AlignLeft("IGST Tax Amt: " & "0.00", 48)
'    End If
        
    If Val(LBLDISCAMT.Caption) <> 0 Then
        Print #1, AlignRight("BILL AMOUNT ", 108) & AlignRight((Format(Val(LBLTOTAL.Caption), "####.00")), 12)
        Print #1, AlignRight("DISC AMOUNT ", 108) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
    ElseIf Val(LBLDISCAMT.Caption) = 0 Then
        Print #1, AlignRight("BILL AMOUNT ", 108) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
    End If
    If Val(Txthandle.text) <> 0 Then
        Print #1, AlignRight("HANDLING CHARGE ", 108) & AlignRight((Format(Txthandle.text, "####.00")), 12)
    End If
    If Val(TxtFrieght.text) <> 0 Then
        Print #1, AlignRight("FRIEGHT CHARGE  ", 108) & AlignRight((Format(Txthandle.text, "####.00")), 12)
    End If
    'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 57) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
    Print #1, AlignRight("Round off ", 108) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
    'Print #1, Chr(13)
    'Print #1, AlignRight("NET AMOUNT ", 108) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 12)
    Print #1, Chr(27) & Chr(72) & Chr(10) & _
               Chr(14) & Chr(15) & Space(40) & "NET AMOUNT:" & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 10) & _
              Chr(27) & Chr(72)
    'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
    Num = CCur(Round(lblnetamount.Caption, 0))
    Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & ")", 125)
    Print #1, RepeatString("-", 125)
    'Print #1, Chr(27) & Chr(71) & Chr(0)
    Print #1, "Certified that all the above particulars are true and correct"
    Print #1, RepeatString("-", 125)
    Dim taxableamt As Double
    Dim Taxsplit As String
    
    TaxAmt = 0
    taxableamt = 0
    Taxsplit = ""
    Dim RSTtax As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT DISTINCT SALES_TAX From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND SALES_TAX > 0 order by SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        TaxAmt = 0
        taxableamt = 0
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select * From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND SALES_TAX = " & RSTTRXFILE!SALES_TAX & " AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            If OPTDISCPERCENT.Value = True Then
                grdtmp.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * Val(TXTTOTALDISC.text) / 100)
            Else
                grdtmp.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
            End If
            
            taxableamt = Round(taxableamt + Val(grdtmp.Tag) * Val(RSTtax!QTY), 2)
            TaxAmt = Round(TaxAmt + (Val(grdtmp.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY, 2)
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
        'Print #1, Space(7) & "GST " & RSTTRXFILE!SALES_TAX & "%: " & "  TAXABLE VALUE: " & Format(Round(TaxableAmt, 2), "0.00") & "   TAX AMT: " & Format(Round(TaxAmt, 2), "0.00")
        Print #1, Space(0) & "TAXABLE VALUE: " & Format(Round(taxableamt, 2), "0.00") & "   " & "SGST " & RSTTRXFILE!SALES_TAX / 2 & "%: " & " TAX AMT: " & Format(Round(TaxAmt / 2, 2), "0.00") & "   " & "CGST " & RSTTRXFILE!SALES_TAX / 2 & "%: " & " TAX AMT: " & Format(Round(TaxAmt / 2, 2), "0.00")
        Taxsplit = Taxsplit & "GST " & RSTTRXFILE!SALES_TAX & "%: " & "Taxable: " & Format(Round(taxableamt, 2), "0.00") & " Tax: " & Format(Round(TaxAmt, 2), "0.00") & " "
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Print #1, RepeatString("-", 125)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Space(75) & AlignRight("For " & MDIMAIN.StatusBar.Panels(5).text, 50)
    If Not INVMSG = "" Then
        Print #1, Chr(13)
        PRNSPACE = Int(62 - Len(INVMSG) / 2)
        Print #1, Space(PRNSPACE) & Chr(27) & Chr(72) & INVMSG
    Else
        Print #1, Chr(13)
        Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** THANK YOU VISIT AGAIN ****", 40)
    End If
    
    

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
    'Print #1, Chr(13)
    
    Close #1 '//Closing the file
    Exit Function

ERRHAND:
    MsgBox err.Description
End Function

Private Function CONTINUE_BATCH()
    
    If Not (IsNull(GRDPOPUP.Columns(30)) Or GRDPOPUP.Columns(30) = "") Then
        item_change = True
        TXTITEMCODE.text = GRDPOPUP.Columns(30)
        TXTPRODUCT.text = GRDPOPUP.Columns(6)
        item_change = False
    End If
    M_STOCK = Val(GRDPOPUP.Columns(1))
    If M_STOCK <= 0 Then
        MsgBox "AVAILABLE STOCK IS  " & M_STOCK & " ", , "SALES"
        Exit Function
    End If
            
    Dim i As Double
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.text) Then
                        If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                            Exit Function
                        Else
                            Select Case grdsales.TextMatrix(i, 19)
                                Case "CN", "DN"
                                    Exit For
                            End Select
'                            If SERIAL_FLAG = False Then
'                                TXTSLNO.Text = i
'                                TXTAPPENDQTY.Text = Val(grdsales.TextMatrix(i, 3))
'                                TXTFREEAPPEND.Text = Val(grdsales.TextMatrix(i, 20))
'                                txtappendcomm.Text = Val(grdsales.TextMatrix(i, 24))
'                                Exit For
'                            End If
                        End If
                        Exit For
                    End If
                Next i
                txtcategory.text = IIf(IsNull(PHY!Category), "", PHY!Category)
                If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                    TXTQTY.text = 1
                    txtretail.Enabled = True
                    txtretail.SetFocus
                    Exit Function
                End If
                Select Case cmbtype.ListIndex
                    Case 1
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(13)), "", GRDPOPUP.Columns(13))
                        'Kannattu
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                    Case 2
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
                        'Kannattu
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
                    Case 3
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
                        'Kannattu
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(21)), "", GRDPOPUP.Columns(21))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(21)), "", GRDPOPUP.Columns(21))
                        If Val(txtretail.text) = 0 Then
                            txtretail.text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                            TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                        End If
                    Case Else
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(20)), "", GRDPOPUP.Columns(20))
                        'Kannattu
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                End Select
                If Val(TxtCessPer.text) <> 0 Or Val(TxtCessAmt.text) <> 0 Then
                    TXTRETAILNOTAX.text = (Val(txtretail.text) - Val(TxtCessAmt.text)) / (1 + (Val(TXTTAX.text) / 100) + (Val(TxtCessPer.text) / 100))
                    txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 3)
                    TXTRETAILNOTAX.text = Val(txtretail.text)
                End If

                'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                lblretail.Caption = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                lblwsale.Caption = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                lblvan.Caption = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
                LBLMRP.Caption = IIf(IsNull(GRDPOPUP.Columns(21)), "", GRDPOPUP.Columns(21))
                lblcase.Caption = IIf(IsNull(GRDPOPUP.Columns(12)), "", GRDPOPUP.Columns(12))
                lblcrtnpack.Caption = IIf(IsNull(GRDPOPUP.Columns(22)), "", GRDPOPUP.Columns(22))
                lblLWPrice.Caption = IIf(IsNull(GRDPOPUP.Columns(29)), "", GRDPOPUP.Columns(29))
                LblPack.text = IIf(IsNull(GRDPOPUP.Columns(14)) Or GRDPOPUP.Columns(14) = "", "1", GRDPOPUP.Columns(14))
                lblOr_Pack.Caption = IIf(IsNull(GRDPOPUP.Columns(14)) Or GRDPOPUP.Columns(14) = "", "1", GRDPOPUP.Columns(14))
                lblunit.text = IIf(IsNull(GRDPOPUP.Columns(15)), "Nos", GRDPOPUP.Columns(15))
                TxtWarranty.text = IIf(IsNull(GRDPOPUP.Columns(7)), "", GRDPOPUP.Columns(7))
                TxtWarranty_type.text = IIf(IsNull(GRDPOPUP.Columns(8)), "", GRDPOPUP.Columns(8))
                If GRDPOPUP.Columns(16) = "A" Then
                    txtretaildummy.text = IIf(IsNull(GRDPOPUP.Columns(18)), "P", GRDPOPUP.Columns(18))
                    TxtRetailmode.text = "A"
                    lblcomtype.Caption = "`"
                    txtcommi.text = Format(Round(Val(txtretaildummy.text), 2), "0.00")
                Else
                    txtretaildummy.text = IIf(IsNull(GRDPOPUP.Columns(17)), "P", GRDPOPUP.Columns(17))
                    TxtRetailmode.text = "P"
                    lblcomtype.Caption = "%"
                    txtcommi.text = Format(Round((Val(TXTRETAILNOTAX.text) * Val(txtretaildummy.text) / 100), 2), "0.00")
                End If
                If Val(txtretaildummy.text) = 0 Then lblcomtype.Caption = "`"
'                If GRDPOPUP.Columns(19) >= 5 Then
'                    Select Case PHY!CHECK_FLAG
'                        Case "M", "I"
'                            OPTTaxMRP.value = True
'                            TXTTAX.Text = GRDPOPUP.Columns(19)
'                            TXTSALETYPE.Text = "2"
'                        Case "V"
'                            OPTVAT.value = True
'                            TXTSALETYPE.Text = "2"
'                            TXTTAX.Text = GRDPOPUP.Columns(19)
'                        Case Else
'                            TXTSALETYPE.Text = "2"
'                            optnet.value = True
'                            TXTTAX.Text = "0"
'                    End Select
'                End If
                TXTUNIT.text = GRDPOPUP.Columns(20)
                                   
                'TXTPRODUCT.Enabled = False
                'TXTQTY.Enabled = True
                '
                'OptLoose.value = True
                'TXTQTY.SetFocus
                Exit Function
End Function

Private Sub TXTTYPE_GotFocus()
    TXTTYPE.SelStart = 0
    TXTTYPE.SelLength = Len(TXTTYPE.text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    Call CHANGEBOXCOLOR(TXTTYPE, True)
End Sub

Private Sub TXTTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Val(TXTTYPE.text) = 0 Or Val(TXTTYPE.text) > 4 Then
                MsgBox "Enter Bill Type", vbOKOnly, "Sales"
                TXTTYPE.Enabled = True
                TXTTYPE.SetFocus
                Exit Sub
            End If
            cmbtype.Enabled = True
            cmbtype.SetFocus
        Case vbKeyEscape
            TxtBillAddress.Enabled = True
            TxtBillAddress.SetFocus
    End Select
End Sub

Private Sub TXTTYPE_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cmbtype_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtVehicle.Enabled = True
            TxtVehicle.SetFocus
        Case vbKeyEscape
            TXTTYPE.Enabled = True
            TXTTYPE.SetFocus
    End Select
End Sub

Private Sub cmbtype_LostFocus()
'    If cmbtype.ListIndex = -1 Then
'        MsgBox "Select Bill Type from the List", vbOKOnly, "Sales"
'        cmbtype.SetFocus
'        Exit Sub
'    End If
'    If cmbtype.ListIndex = 0 And Val(TXTTYPE.Text) <> 1 Then
'        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
'        TXTTYPE.SetFocus
'        Exit Sub
'    End If
'    If cmbtype.ListIndex = 1 And Val(TXTTYPE.Text) <> 2 Then
'        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
'        TXTTYPE.SetFocus
'        Exit Sub
'    End If
'    If cmbtype.ListIndex = 2 And Val(TXTTYPE.Text) <> 3 Then
'        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
'        TXTTYPE.SetFocus
'        Exit Sub
'    End If
    
End Sub

Private Function REMOVE_ITEM()
    Dim i As Long
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Function
      
    For i = Val(TXTSLNO.text) To grdsales.rows - 2
        grdsales.TextMatrix(i, 0) = i
        grdsales.TextMatrix(i, 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(i, 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(i, 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(i, 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(i, 6) = grdsales.TextMatrix(i + 1, 6)
        grdsales.TextMatrix(i, 5) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(i, 7) = grdsales.TextMatrix(i + 1, 7)
        grdsales.TextMatrix(i, 8) = grdsales.TextMatrix(i + 1, 8)
        grdsales.TextMatrix(i, 9) = grdsales.TextMatrix(i + 1, 9)
        grdsales.TextMatrix(i, 10) = grdsales.TextMatrix(i + 1, 10)
        grdsales.TextMatrix(i, 11) = grdsales.TextMatrix(i + 1, 11)
        grdsales.TextMatrix(i, 12) = grdsales.TextMatrix(i + 1, 12)
        grdsales.TextMatrix(i, 13) = grdsales.TextMatrix(i + 1, 13)
        grdsales.TextMatrix(i, 14) = grdsales.TextMatrix(i + 1, 14)
        grdsales.TextMatrix(i, 15) = grdsales.TextMatrix(i + 1, 15)
        grdsales.TextMatrix(i, 16) = grdsales.TextMatrix(i + 1, 16)
        grdsales.TextMatrix(i, 17) = grdsales.TextMatrix(i + 1, 17)
        grdsales.TextMatrix(i, 18) = grdsales.TextMatrix(i + 1, 18)
        grdsales.TextMatrix(i, 19) = grdsales.TextMatrix(i + 1, 19)
        grdsales.TextMatrix(i, 20) = grdsales.TextMatrix(i + 1, 20)
        grdsales.TextMatrix(i, 21) = grdsales.TextMatrix(i + 1, 21)
        grdsales.TextMatrix(i, 22) = grdsales.TextMatrix(i + 1, 22)
        grdsales.TextMatrix(i, 23) = grdsales.TextMatrix(i + 1, 23)
        grdsales.TextMatrix(i, 24) = grdsales.TextMatrix(i + 1, 24)
        grdsales.TextMatrix(i, 25) = grdsales.TextMatrix(i + 1, 25)
        grdsales.TextMatrix(i, 26) = grdsales.TextMatrix(i + 1, 26)
        grdsales.TextMatrix(i, 27) = grdsales.TextMatrix(i + 1, 27)
        grdsales.TextMatrix(i, 28) = grdsales.TextMatrix(i + 1, 28)
        grdsales.TextMatrix(i, 29) = grdsales.TextMatrix(i + 1, 29)
        grdsales.TextMatrix(i, 30) = grdsales.TextMatrix(i + 1, 30)
        grdsales.TextMatrix(i, 31) = grdsales.TextMatrix(i + 1, 31)
        grdsales.TextMatrix(i, 32) = grdsales.TextMatrix(i + 1, 32)
        grdsales.TextMatrix(i, 33) = grdsales.TextMatrix(i + 1, 33)
        grdsales.TextMatrix(i, 34) = grdsales.TextMatrix(i + 1, 34)
        grdsales.TextMatrix(i, 35) = grdsales.TextMatrix(i + 1, 35)
        grdsales.TextMatrix(i, 36) = grdsales.TextMatrix(i + 1, 36)
        grdsales.TextMatrix(i, 37) = grdsales.TextMatrix(i + 1, 37)
        grdsales.TextMatrix(i, 38) = grdsales.TextMatrix(i + 1, 38)
        grdsales.TextMatrix(i, 39) = grdsales.TextMatrix(i + 1, 39)
        grdsales.TextMatrix(i, 40) = grdsales.TextMatrix(i + 1, 40)
        grdsales.TextMatrix(i, 41) = grdsales.TextMatrix(i + 1, 41)
        grdsales.TextMatrix(i, 42) = grdsales.TextMatrix(i + 1, 42)
        grdsales.TextMatrix(i, 43) = grdsales.TextMatrix(i + 1, 43)
        grdsales.TextMatrix(i, 44) = grdsales.TextMatrix(i + 1, 44)
        grdsales.TextMatrix(i, 45) = grdsales.TextMatrix(i + 1, 45)
        grdsales.TextMatrix(i, 46) = grdsales.TextMatrix(i + 1, 46)
    Next i
    grdsales.rows = grdsales.rows - 1
    
    LBLTOTAL.Caption = ""
    LBLGROSSAMT.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.text = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
        Else
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
        End If
    Next i
    
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
    If Val(lblnetamount.Caption) = 0 Then
        txtcommper.text = 0
    Else
        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
    End If
    Call COSTCALCULATION
    Call Addcommission
    
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
    Else
        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
    End If
    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"

    
    TXTSLNO.text = Val(grdsales.rows)
    TXTPRODUCT.text = ""
    txtPrintname.text = ""
    TxtSpec.text = ""
    txtcategory.text = ""
    TxtName1.text = ""
    TXTITEMCODE.text = ""
    optnet.Value = True
    TXTVCHNO.text = ""
    TXTLINENO.text = ""
    TXTTRXTYPE.text = ""
    TrxRYear.text = ""
    TXTUNIT.text = ""
    TXTQTY.text = ""
    TXTEXPIRY.text = "  /  "
    TXTAPPENDQTY.text = ""
    TXTFREEAPPEND.text = ""
    txtappendcomm.text = ""
    TXTAPPENDTOTAL.text = ""
    txtretail.text = ""
    txtNetrate.text = ""
    txtBatch.text = ""
    TxtWarranty.text = ""
    TxtWarranty_type.text = ""
    TXTTAX.text = ""
    TXTRETAILNOTAX.text = ""
    TXTSALETYPE.text = ""
    TXTFREE.text = ""
    TxtMRP.text = ""
    txtmrpbt.text = ""
    txtretaildummy.text = ""
    lblcomtype.Caption = ""
    txtcommi.text = ""
    TxtRetailmode.text = ""
    
    TXTDISC.text = ""
    TxtDiscAmt.text = ""
    TxtCessAmt.text = ""
    TxtCessPer.text = ""
    LBLUNBILL.Caption = ""
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    LBLDNORCN.Caption = ""
    cmdadd.Enabled = False
    'cmddelete.Enabled = False
    'CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    CMDPre.Caption = "Save && Pre&vious"
    cmdnxt.Caption = "Save && &Next"
            
    M_EDIT = False
    M_ADD = True
    TXTQTY.Enabled = False
    TXTITEMCODE.Enabled = True
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    If TXTITEMCODE.Visible = True Then
        If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
    Else
        If TxtName1.Enabled = True Then TxtName1.SetFocus
    End If
    If grdsales.rows >= 9 Then grdsales.TopRow = grdsales.rows - 1

End Function

Private Function Addcommission()
    Dim i As Long
    On Error GoTo ERRHAND
    lblActAmt.Caption = ""
    For i = 1 To grdsales.rows - 1
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblActAmt.Caption = Val(lblActAmt.Caption) + Val(grdsales.TextMatrix(i, 24))
        Else
            lblActAmt.Caption = Val(lblActAmt.Caption) + (Val(grdsales.TextMatrix(i, 24)) * Val(grdsales.TextMatrix(i, 3)))
        End If
    Next i
    
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Dim RSTTRXFILE As ADODB.Recordset
                Dim i As Integer
                Case 31  'ST_RATE
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.00")
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                Case 3 'Qty
                    db.Execute "START TRANSACTION"
                    If Not (UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICES" Or UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SELF") Then
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        With RSTTRXFILE
                            If Not (.EOF And .BOF) Then
                                .Properties("Update Criteria").Value = adCriteriaKey
                                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                                
                                !ISSUE_QTY = !ISSUE_QTY - Round(Val(TXTsample.text) * Val(grdsales.TextMatrix(grdsales.Row, 27)), 3)
                                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(grdsales.Row, 20)) * Val(grdsales.TextMatrix(grdsales.Row, 27)), 3)
                                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) + Val(grdsales.TextMatrix(grdsales.Row, 20))) * Val(grdsales.TextMatrix(grdsales.Row, 27)), 3)
                                
                                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(grdsales.Row, 12))
                                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(grdsales.Row, 12))
                                RSTTRXFILE.Update
                                      
                                TXTDISC.Tag = 0
                                If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                                    TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                                    grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(TXTDISC.Tag), 4), ".0000")
                                    grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) - Val(TXTDISC.Tag), 4), ".0000")
                                Else
                                    TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                                    grdsales.TextMatrix(grdsales.Row, 12) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7))) - Val(TXTDISC.Tag), 4), ".0000")
                                    grdsales.TextMatrix(grdsales.Row, 34) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 6))) - Val(TXTDISC.Tag), 4), ".0000")
                                End If
                    
                                'If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                                !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTsample.text) * Val(grdsales.TextMatrix(grdsales.Row, 27))), 3)
                                !FREE_QTY = !FREE_QTY + Round((Val(TXTFREE.text) * Val(grdsales.TextMatrix(grdsales.Row, 27))), 3)
                                !CLOSE_QTY = !CLOSE_QTY - Round(((Val(TXTsample.text) + Val(grdsales.TextMatrix(grdsales.Row, 20))) * Val(grdsales.TextMatrix(grdsales.Row, 27))), 3)
                    
                                'If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                                !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(grdsales.Row, 12))
                                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                                !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(grdsales.Row, 12))
                                RSTTRXFILE.Update
                                
                            End If
                        End With
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                           
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE TRX_TYPE = '" & Trim(grdsales.TextMatrix(grdsales.Row, 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(grdsales.Row, 43)) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
                        With RSTTRXFILE
                            If Not (.EOF And .BOF) Then
                                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) + Val(grdsales.TextMatrix(grdsales.Row, 20))) * Val(grdsales.TextMatrix(grdsales.Row, 27)), 3)
                                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) + Val(grdsales.TextMatrix(grdsales.Row, 20))) * Val(grdsales.TextMatrix(grdsales.Row, 27)), 3)
                                lblactqty.Caption = !BAL_QTY
                                lblbarcode.Caption = IIf(IsNull(!BARCODE), "", !BARCODE)
                                RSTTRXFILE.Update
                                
                                !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTsample.text) + Val(grdsales.TextMatrix(grdsales.Row, 20))) * Val(grdsales.TextMatrix(grdsales.Row, 27)), 3)
                                !BAL_QTY = !BAL_QTY - Round((Val(TXTsample.text) + Val(grdsales.TextMatrix(grdsales.Row, 20))) * Val(grdsales.TextMatrix(grdsales.Row, 27)), 3)
                                grdsales.TextMatrix(grdsales.Row, 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(grdsales.TextMatrix(grdsales.Row, 27)))
                                
                                RSTTRXFILE.Update
                                
                                RSTTRXFILE.Close
                                Set RSTTRXFILE = Nothing
                            Else
                                '===========
                                RSTTRXFILE.Close
                                Set RSTTRXFILE = Nothing
                                
                                Set RSTTRXFILE = New ADODB.Recordset
                                RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "' AND BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
                                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                                    If (IsNull(RSTTRXFILE!ISSUE_QTY)) Then RSTTRXFILE!ISSUE_QTY = 0
                                    If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                                    'BALQTY = RSTTRXFILE!BAL_QTY
                                    RSTTRXFILE!ISSUE_QTY = RSTTRXFILE!ISSUE_QTY + Round((Val(TXTsample.text) + Val(grdsales.TextMatrix(grdsales.Row, 20))) * Val(grdsales.TextMatrix(grdsales.Row, 27)), 3)
                                    RSTTRXFILE!BAL_QTY = RSTTRXFILE!BAL_QTY - Round((Val(TXTsample.text) + Val(grdsales.TextMatrix(grdsales.Row, 20))) * Val(grdsales.TextMatrix(grdsales.Row, 27)), 3)
                                    
                                    grdsales.TextMatrix(grdsales.Row, 14) = RSTTRXFILE!VCH_NO
                                    grdsales.TextMatrix(grdsales.Row, 15) = RSTTRXFILE!LINE_NO
                                    grdsales.TextMatrix(grdsales.Row, 16) = RSTTRXFILE!TRX_TYPE
                                    grdsales.TextMatrix(grdsales.Row, 43) = RSTTRXFILE!TRX_YEAR
                                    grdsales.TextMatrix(grdsales.Row, 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(grdsales.TextMatrix(grdsales.Row, 27)))
                                    RSTTRXFILE.Update
                                    RSTTRXFILE.Close
                                    Set RSTTRXFILE = Nothing
                                Else
                                    RSTTRXFILE.Close
                                    Set RSTTRXFILE = Nothing
                                    Set RSTTRXFILE = New ADODB.Recordset
                                    RSTTRXFILE.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
                                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                                        grdsales.TextMatrix(grdsales.Row, 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(grdsales.TextMatrix(grdsales.Row, 27)))
                                    End If
                                    RSTTRXFILE.Close
                                    Set RSTTRXFILE = Nothing
                                End If
                                '===========
                            End If
                        End With
                        
                    End If
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.000")
                    
                    TXTDISC.Tag = 0
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) - Val(TXTDISC.Tag), 4), ".0000")
                    Else
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7))) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 6))) - Val(TXTDISC.Tag), 4), ".0000")
                    End If
                    
                    LBLTOTAL.Caption = ""
                    LBLGROSSAMT.Caption = ""
                    lblnetamount.Caption = ""
                    LBLFOT.Caption = ""
                    lblcomamt.text = ""
        
                    For i = 1 To grdsales.rows - 1
                        grdsales.TextMatrix(i, 0) = i
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN"
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                        End Select
                        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
                        Else
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
                        End If
                    Next i
                    
                    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
                    TXTAMOUNT.text = ""
                    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
                    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
                    End If
                    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
                    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
                    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
                    If Val(lblnetamount.Caption) = 0 Then
                        txtcommper.text = 0
                    Else
                        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
                    End If
                    Call COSTCALCULATION
                    
                    TXTDISC.Tag = (Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(grdsales.TextMatrix(grdsales.Row, 6))) * Val(grdsales.TextMatrix(grdsales.Row, 3))
                    db.Execute "Update TRXFILEVAN set QTY = " & Val(grdsales.TextMatrix(grdsales.Row, 3)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"
                    db.Execute "COMMIT"
                    
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                    
                Case 5  'MRP
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.000")
                    db.Execute "START TRANSACTION"
                    db.Execute "Update TRXFILEVAN set MRP = " & Val(grdsales.TextMatrix(grdsales.Row, 5)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    db.Execute "COMMIT"
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
            
                Case 6  'RATE
                    TXTDISC.Tag = 0
                    grdsales.TextMatrix(grdsales.Row, 7) = Format(Round(Val(TXTsample.text) + Val(TXTsample.text) * Val(grdsales.TextMatrix(grdsales.Row, 9)) / 100, 4), "0.0000")
                    grdsales.TextMatrix(grdsales.Row, 21) = Format(Round(Val(TXTsample.text) + Val(TXTsample.text) * Val(grdsales.TextMatrix(grdsales.Row, 9)) / 100, 4), "0.0000")
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(TXTsample.text) - Val(TXTDISC.Tag), 4), ".0000")
                    Else
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7))) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(TXTsample.text)) - Val(TXTDISC.Tag), 4), ".0000")
                    End If
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.000")
                    LBLTOTAL.Caption = ""
                    LBLGROSSAMT.Caption = ""
                    lblnetamount.Caption = ""
                    LBLFOT.Caption = ""
                    lblcomamt.text = ""
                    For i = 1 To grdsales.rows - 1
                        grdsales.TextMatrix(i, 0) = i
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN"
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                        End Select
                        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
                        Else
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
                        End If
                    Next i
                    
                    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
                    TXTAMOUNT.text = ""
                    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
                    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
                    End If
                    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
                    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
                    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
                    If Val(lblnetamount.Caption) = 0 Then
                        txtcommper.text = 0
                    Else
                        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
                    End If
                    Call COSTCALCULATION
                    
                    db.Execute "START TRANSACTION"
                    TXTDISC.Tag = (Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(grdsales.TextMatrix(grdsales.Row, 6))) * Val(grdsales.TextMatrix(grdsales.Row, 3))
                    db.Execute "Update TRXFILEVAN set SALES_PRICE = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", PTR = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"
                    db.Execute "COMMIT"
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                    
                Case 7  'NET RATE
                    TXTDISC.Tag = 0
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "PARDHA" Or UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "CLOTHES" Then
                        If Val(grdsales.TextMatrix(grdsales.Row, 6)) < 1000 Then
                           grdsales.TextMatrix(grdsales.Row, 9) = "5"
                        Else
                            grdsales.TextMatrix(grdsales.Row, 9) = "12"
                        End If
                    End If
                    'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                    grdsales.TextMatrix(grdsales.Row, 6) = Format(Round(Val(TXTsample.text) * 100 / (Val(grdsales.TextMatrix(grdsales.Row, 9)) + 100), 4), "0.0000")
                    grdsales.TextMatrix(grdsales.Row, 22) = Format(Round(Val(TXTsample.text) * 100 / (Val(grdsales.TextMatrix(grdsales.Row, 9)) + 100), 4), "0.0000")
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        TXTDISC.Tag = Val(TXTsample.text) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTsample.text) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) - Val(TXTDISC.Tag), 4), ".0000")
                    Else
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(TXTsample.text) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(TXTsample.text)) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 6))) - Val(TXTDISC.Tag), 4), ".0000")
                    End If
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.000")
                    LBLTOTAL.Caption = ""
                    LBLGROSSAMT.Caption = ""
                    lblnetamount.Caption = ""
                    LBLFOT.Caption = ""
                    lblcomamt.text = ""
                    For i = 1 To grdsales.rows - 1
                        grdsales.TextMatrix(i, 0) = i
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN"
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                        End Select
                        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
                        Else
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
                        End If
                    Next i
                    
                    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
                    TXTAMOUNT.text = ""
                    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
                    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
                    End If
                    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
                    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
                    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
                    If Val(lblnetamount.Caption) = 0 Then
                        txtcommper.text = 0
                    Else
                        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
                    End If
                    Call COSTCALCULATION
                    
                    TXTDISC.Tag = (Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(grdsales.TextMatrix(grdsales.Row, 6))) * Val(grdsales.TextMatrix(grdsales.Row, 3))
                    db.Execute "START TRANSACTION"
                    db.Execute "Update TRXFILEVAN set SALES_PRICE = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", PTR = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"
                    db.Execute "COMMIT"
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                
                Case 8  'Disc
                    TXTDISC.Tag = 0
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(TXTsample.text) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) - Val(TXTDISC.Tag), 4), ".0000")
                    Else
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(TXTsample.text) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7))) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 6))) - Val(TXTDISC.Tag), 4), ".0000")
                    End If
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.00")
                    LBLTOTAL.Caption = ""
                    LBLGROSSAMT.Caption = ""
                    lblnetamount.Caption = ""
                    LBLFOT.Caption = ""
                    lblcomamt.text = ""
                    For i = 1 To grdsales.rows - 1
                        grdsales.TextMatrix(i, 0) = i
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN"
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                        End Select
                        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
                        Else
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
                        End If
                    Next i
                    
                    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
                    TXTAMOUNT.text = ""
                    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
                    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
                    End If
                    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
                    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
                    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
                    If Val(lblnetamount.Caption) = 0 Then
                        txtcommper.text = 0
                    Else
                        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
                    End If
                    Call COSTCALCULATION
                    
                    db.Execute "START TRANSACTION"
                    db.Execute "Update TRXFILEVAN set LINE_DISC = " & Val(grdsales.TextMatrix(grdsales.Row, 8)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"
                    db.Execute "COMMIT"
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                    
                Case 9  'TAX
                    TXTDISC.Tag = 0
                    If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                        grdsales.TextMatrix(grdsales.Row, 7) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) + Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(TXTsample.text) / 100, 3), "0.000")
                        grdsales.TextMatrix(grdsales.Row, 21) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) + Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(TXTsample.text) / 100, 3), "0.000")
                    End If
                    grdsales.TextMatrix(grdsales.Row, 6) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) * 100 / (Val(TXTsample.text) + 100), 3), "0.000")
                    grdsales.TextMatrix(grdsales.Row, 22) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) * 100 / (Val(TXTsample.text) + 100), 3), "0.000")
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(TXTsample.text) - Val(TXTDISC.Tag), 4), ".0000")
                    Else
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7))) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 6))) - Val(TXTDISC.Tag), 4), ".0000")
                    End If
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.000")
                    grdsales.TextMatrix(grdsales.Row, 17) = "V"
                    
                    LBLTOTAL.Caption = ""
                    LBLGROSSAMT.Caption = ""
                    lblnetamount.Caption = ""
                    LBLFOT.Caption = ""
                    lblcomamt.text = ""
                    For i = 1 To grdsales.rows - 1
                        grdsales.TextMatrix(i, 0) = i
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN"
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                        End Select
                        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
                        Else
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
                        End If
                    Next i
                    
                    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
                    TXTAMOUNT.text = ""
                    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
                    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
                    End If
                    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
                    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
                    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
                    If Val(lblnetamount.Caption) = 0 Then
                        txtcommper.text = 0
                    Else
                        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
                    End If
                    Call COSTCALCULATION
                    
                    TXTDISC.Tag = (Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(grdsales.TextMatrix(grdsales.Row, 6))) * Val(grdsales.TextMatrix(grdsales.Row, 3))
                    db.Execute "START TRANSACTION"
                    db.Execute "Update TRXFILEVAN set SALES_TAX = " & Val(grdsales.TextMatrix(grdsales.Row, 9)) & ", SALES_PRICE = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", PTR = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"
                    db.Execute "COMMIT"
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                Case 24  'Commission
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.00")
                    LBLTOTAL.Caption = ""
                    LBLGROSSAMT.Caption = ""
                    lblnetamount.Caption = ""
                    LBLFOT.Caption = ""
                    lblcomamt.text = ""
                    For i = 1 To grdsales.rows - 1
                        grdsales.TextMatrix(i, 0) = i
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN"
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                        End Select
                        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
                        Else
                            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
                        End If
                    Next i
                    
                    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
                    TXTAMOUNT.text = ""
                    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
                    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
                        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
                    End If
                    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
                    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
                    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
                    If Val(lblnetamount.Caption) = 0 Then
                        txtcommper.text = 0
                    Else
                        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
                    End If
                    Call COSTCALCULATION
                    
                    db.Execute "START TRANSACTION"
                    db.Execute "Update TRXFILEVAN set COM_AMT = " & Val(grdsales.TextMatrix(grdsales.Row, 24)) & "  WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'"
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update TRXMASTVAN set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update TRXMASTVAN set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    'db.Execute "Update TRXMASTVAN set PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.Text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.Text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
                    db.Execute "COMMIT"
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
            End Select
            Call COSTCALCULATE
        Case vbKeyEscape
            TXTsample.Visible = False
            grdsales.SetFocus
    End Select
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.Execute "ROLLBACK"
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case grdsales.Col
        Case 31
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

Private Sub grdsales_Click()
    TXTsample.Visible = False
    grdsales.SetFocus
    If grdsales.rows <= 1 Then Exit Sub
    Call COSTCALCULATE
    If M_EDIT = True Then Exit Sub
    If TXTQTY.Enabled = True Then Exit Sub
    lblretail.Caption = ""
    lblwsale.Caption = ""
    lblvan.Caption = ""
    LBLMRP.Caption = ""
    lblcase.Caption = ""
    lblLWPrice.Caption = ""
    lblcrtnpack.Caption = 1
        
    Dim RSTITEMMAST As ADODB.Recordset
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  From RTRXFILEVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "' AND TRX_TYPE = '" & Trim(grdsales.TextMatrix(grdsales.Row, 16)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 14)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 15)) & " AND TRX_YEAR = '" & Val(grdsales.TextMatrix(grdsales.Row, 43)) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        lblretail.Caption = IIf(IsNull(RSTITEMMAST!P_RETAIL) Or RSTITEMMAST!P_RETAIL = "", "", RSTITEMMAST!P_RETAIL)
        lblwsale.Caption = IIf(IsNull(RSTITEMMAST!P_WS) Or RSTITEMMAST!P_WS = 0, "", RSTITEMMAST!P_WS)
        lblvan.Caption = IIf(IsNull(RSTITEMMAST!P_VAN) Or RSTITEMMAST!P_VAN = 0, "", RSTITEMMAST!P_VAN)
        LBLMRP.Caption = IIf(IsNull(RSTITEMMAST!MRP) Or RSTITEMMAST!MRP = 0, "", RSTITEMMAST!MRP)
        lblcase.Caption = IIf(IsNull(RSTITEMMAST!P_CRTN) Or RSTITEMMAST!P_CRTN = 0, "", RSTITEMMAST!P_CRTN)
        lblLWPrice.Caption = IIf(IsNull(RSTITEMMAST!P_LWS) Or RSTITEMMAST!P_LWS = 0, "", RSTITEMMAST!P_LWS)
        lblcrtnpack.Caption = IIf(IsNull(RSTITEMMAST!CRTN_PACK) Or RSTITEMMAST!CRTN_PACK = 0, 1, RSTITEMMAST!CRTN_PACK)
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
        
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        If Val(lblretail.Caption) = 0 Then lblretail.Caption = IIf(IsNull(RSTITEMMAST!P_RETAIL) Or RSTITEMMAST!P_RETAIL = "", "", RSTITEMMAST!P_RETAIL)
        If Val(lblwsale.Caption) = 0 Then lblwsale.Caption = IIf(IsNull(RSTITEMMAST!P_WS) Or RSTITEMMAST!P_WS = 0, "", RSTITEMMAST!P_WS)
        If Val(lblvan.Caption) = 0 Then lblvan.Caption = IIf(IsNull(RSTITEMMAST!P_VAN) Or RSTITEMMAST!P_VAN = 0, "", RSTITEMMAST!P_VAN)
        If Val(LBLMRP.Caption) = 0 Then LBLMRP.Caption = IIf(IsNull(RSTITEMMAST!MRP) Or RSTITEMMAST!MRP = 0, "", RSTITEMMAST!MRP)
        lblcase.Caption = IIf(IsNull(RSTITEMMAST!P_CRTN) Or RSTITEMMAST!P_CRTN = 0, "", RSTITEMMAST!P_CRTN)
        lblLWPrice.Caption = IIf(IsNull(RSTITEMMAST!P_LWS) Or RSTITEMMAST!P_LWS = 0, "", RSTITEMMAST!P_LWS)
        lblcrtnpack.Caption = IIf(IsNull(RSTITEMMAST!CRTN_PACK) Or RSTITEMMAST!CRTN_PACK = 0, 1, RSTITEMMAST!CRTN_PACK)
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If grdsales.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then Exit Sub
            Select Case grdsales.Col
                Case 3
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = grdsales.CellTop + 100
                    TXTsample.Left = grdsales.CellLeft + 0
                    TXTsample.Width = grdsales.CellWidth
                    TXTsample.Height = grdsales.CellHeight
                    TXTsample.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TXTsample.SetFocus
                Case 31, 5, 6, 7, 9, 24
                    If frmLogin.rs!Level = "5" Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = grdsales.CellTop + 100
                    TXTsample.Left = grdsales.CellLeft + 0
                    TXTsample.Width = grdsales.CellWidth
                    TXTsample.Height = grdsales.CellHeight
                    TXTsample.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TXTsample.SetFocus
                Case 8
                    If frmLogin.rs!Level = "5" Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = grdsales.CellTop + 100
                    TXTsample.Left = grdsales.CellLeft + 0
                    TXTsample.Width = grdsales.CellWidth
                    TXTsample.Height = grdsales.CellHeight
                    TXTsample.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TXTsample.SetFocus
            End Select
    End Select
End Sub

Private Sub grdsales_Scroll()
    TXTsample.Visible = False
    grdsales.SetFocus
End Sub

Private Sub Txthandle_GotFocus()
    Txthandle.SelStart = 0
    Txthandle.SelLength = Len(Txthandle.text)
    Call CHANGEBOXCOLOR(Txthandle, True)
End Sub

Private Sub Txthandle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            'If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            'If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            'If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            'If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            'If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            'If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
            cmdRefresh.Enabled = True
            CMDEXIT.Enabled = False
            CMDPre.Caption = "Save && Pre&vious"
            cmdnxt.Caption = "Save && &Next"
    End Select
End Sub

Private Sub Txthandle_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txthandle_LostFocus()
    Call TXTTOTALDISC_LostFocus
    Call CHANGEBOXCOLOR(Txthandle, False)
End Sub

Public Function AppendSale()
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
    On Error GoTo ERRHAND
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Any changes made will not be saved", vbOKOnly, "Sales"
        GoTo SKIP
    End If
    
    LBLGROSSAMT.Caption = ""
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.text = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                'If Val(grdsales.TextMatrix(I, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(I, 7)) - Val(grdsales.TextMatrix(I, 6))) * Val(grdsales.TextMatrix(I, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                'If Val(grdsales.TextMatrix(I, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(I, 7)) - Val(grdsales.TextMatrix(I, 6))) * Val(grdsales.TextMatrix(I, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 24)), 2)
        Else
            lblcomamt.text = Round(Val(lblcomamt.text) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)), 2)
        End If
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(Val(TXTTOTALDISC.text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        TXTAMOUNT.text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text)
    lblnetamount.Caption = Format(lblnetamount.Caption, "0")
    If Val(lblnetamount.Caption) = 0 Then
        txtcommper.text = 0
    Else
        txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
    End If
    Call COSTCALCULATION
    Call Addcommission
    
    If grdsales.rows <= 1 Then db.Execute "LOCK TABLE TRXMASTVAN Write"
    'If OLD_BILL = False Then Call checklastbill
'    db.Execute "SET autocommit = 0"
'    db.Execute "SET SESSION TRANSACTION ISOLATION LEVEL READ COMMITTED"
    db.Execute "START TRANSACTION"
    Set RSTTRXFILE = New ADODB.Recordset
'    If OLD_BILL = False And Val(txtBillNo.Text) <> 1 Then
'        RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE VCH_NO= (SELECT MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI')", db, adOpenStatic, adLockPessimistic, adCmdText
'        txtBillNo.Text = RSTTRXFILE!VCH_NO + 1
'        LBLBILLNO.Caption = Val(txtBillNo.Text)
'        RSTTRXFILE.AddNew
'        RSTTRXFILE!TRX_TYPE = "GI"
'        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'        RSTTRXFILE!VCH_NO = txtBillNo.Text
'        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
'        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
'        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
'        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
'    Else
'        RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockPessimistic, adCmdText
''        db.Execute "Lock tables TRXMASTVAN write"
''        db.BeginTrans
'        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'            RSTTRXFILE.AddNew
'            RSTTRXFILE!TRX_TYPE = "GI"
'            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'            RSTTRXFILE!VCH_NO = txtBillNo.Text
'            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
'            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
'            RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
'            RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
'        End If
'    End If
    RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockPessimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "GI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = txtBillNo.text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
    End If
    RSTTRXFILE!TIN = Trim(TXTTIN.text)
    RSTTRXFILE!CUST_IGST = lblIGST.Caption
    RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
    RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    RSTTRXFILE!gross_amt = Val(LBLGROSSAMT.Caption)
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.text
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.text)
    RSTTRXFILE!DISC_PERS = Val(TXTTOTALDISC.text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
    RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
    If Val(TxtCN.text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.text)
    RSTTRXFILE!BILL_FLAG = "Y"
    If chkTerms.Value = 1 And Trim(Terms1.text) <> "" Then
        RSTTRXFILE!TERMS = Trim(Terms1.text)
    Else
        RSTTRXFILE!TERMS = ""
    End If
    RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
    RSTTRXFILE!BR_NAME = CMBBRNCH.text
    RSTTRXFILE!cr_days = Val(txtcrdays.text)
    RSTTRXFILE.Update
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.Execute "COMMIT"
    db.Execute "UNLOCK TABLES"
    'OLD_BILL = True
    'db.Execute "unlock tables"
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'GI' ", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        db.Execute "delete From TRNXRCPT WHERE TRX_TYPE='RT' AND CR_NO = " & RSTITEMMAST!CR_NO & " AND CR_TRX_TYPE = 'DR'"
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    'db.Execute "delete From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    'db.Execute "delete From TRXSUBVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    'db.Execute "delete From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='DR' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TRX_TYPE = 'GI'"
    db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_VCH_NO = " & Val(txtBillNo.text) & " AND B_TRX_TYPE = 'GI' "
    db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'GI' "
    db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'GI'"
                            
    'DB.Execute "delete From P_Rate WHERE TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        If Trim(TxtBillName.text) = "" Then
            RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.text), 1, 30)
        Else
            RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(TxtBillName.text), 1, 30)
        End If
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    If grdsales.rows = 1 Then
        Dim TRXMAST As ADODB.Recordset
        Set TRXMAST = New ADODB.Recordset
        TRXMAST.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'GI'", db, adOpenForwardOnly
        If Not (TRXMAST.EOF And TRXMAST.BOF) Then
            If Val(txtBillNo.text) < IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0)) Then
                db.Execute "Update TRXMASTVAN set BILL_NAME = 'Cancelled Bill' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
            Else
                db.Execute "delete From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & ""
            End If
        End If
        TRXMAST.Close
        Set TRXMAST = Nothing
        GoTo SKIP
    End If
    
    
    Dim Max_No As Long
    Max_No = 0
    Set rstMaxRec = New ADODB.Recordset
    rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
        Max_No = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
    End If
    rstMaxRec.Close
    Set rstMaxRec = Nothing
    
    Dim Cash_Flag As Boolean
    Dim RECNO, INVNO As Long
    Dim TRXTYPE, INVTRXTYPE, INVTYPE As String
    Cash_Flag = False
    If grdsales.rows = 1 Then
        db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'GI'"
    Else
        db.Execute "START TRANSACTION"
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'GI'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!REC_NO = Max_No + 1
            RSTITEMMAST!INV_TYPE = "RT"
            RSTITEMMAST!INV_TRX_TYPE = "GI"
            RSTITEMMAST!INV_NO = Val(txtBillNo.text)
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        End If
        'If lblcredit.Caption <> "0" Then
        If DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
            If GRDRECEIPT.rows <= 1 Then
                RSTITEMMAST!TRX_TYPE = "CR"
                RSTITEMMAST!AMOUNT = Val(lblnetamount.Caption)
                RSTITEMMAST!check_flag = "C"
            Else
                RSTITEMMAST!AMOUNT = Val(GRDRECEIPT.TextMatrix(0, 0))
                RSTITEMMAST!TRX_TYPE = "CR"
                RSTITEMMAST!check_flag = "S"
            End If
        Else
            RSTITEMMAST!AMOUNT = Val(lblnetamount.Caption)
            RSTITEMMAST!TRX_TYPE = "CR"
            RSTITEMMAST!check_flag = "S"
        End If
        If RSTITEMMAST!check_flag = "C" Then
            Cash_Flag = False
        Else
            Cash_Flag = True
        End If
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = Trim(DataList2.text)
        RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
        RECNO = RSTITEMMAST!REC_NO
        INVNO = RSTITEMMAST!INV_NO
        TRXTYPE = RSTITEMMAST!TRX_TYPE
        INVTRXTYPE = RSTITEMMAST!INV_TRX_TYPE
        INVTYPE = RSTITEMMAST!INV_TYPE
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        db.Execute "COMMIT"
    End If
    If Cash_Flag = False Then db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'GI'"
    
    i = 0
    If DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
        'If lblcredit.Caption <> "0" Then
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI' "
        Dim CRNO2 As Double
        CRNO2 = 1
        Set rstMaxRec = New ADODB.Recordset
        rstMaxRec.Open "Select MAX(CR_NO) From DBTPYMT", db, adOpenForwardOnly
        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
            i = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
            CRNO2 = i
        End If
        rstMaxRec.Close
        Set rstMaxRec = Nothing
        
        db.Execute "START TRANSACTION"
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!TRX_TYPE = "DR"
            RSTITEMMAST!INV_TRX_TYPE = "GI"
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTITEMMAST!CR_NO = i
            RSTITEMMAST!INV_NO = Val(txtBillNo.text)
            'RSTITEMMAST!RCPT_AMT = 0
        End If
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = DataList2.text
        RSTITEMMAST!INV_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
'        If lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0 Then
'            RSTITEMMAST!INV_AMT = Val(lblActAmt.Caption)
'        Else
'            RSTITEMMAST!INV_AMT = Val(lblnetamount.Caption)
'        End If
        RSTITEMMAST!INV_AMT = Val(lblnetamount.Caption)
        If Trim(CMBBRNCH.text) <> "" Then
            RSTITEMMAST!BR_ADDRESS = Trim(CMBBRNCH.text)
        Else
            RSTITEMMAST!BR_ADDRESS = Trim(TXTAREA.text)
        End If
        If Trim(DataList2.text) <> Trim(TxtBillName.text) Then
            RSTITEMMAST!BR_ADDRESS = IIf(RSTITEMMAST!BR_ADDRESS = "", Trim(TxtBillName.text), RSTITEMMAST!BR_ADDRESS & " " & Trim(TxtBillName.text))
        End If
        RSTITEMMAST!COMM_AMT = Val(lblcomamt.text)
        RSTITEMMAST!PYMT_PERIOD = Val(txtcrdays.text)
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        db.Execute "COMMIT"
        
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'GI' "
        
        Dim MAXRCPTNO As Long
        MAXRCPTNO = 1
        Set rstBILL = New ADODB.Recordset
        rstBILL.Open "Select MAX(REC_NO) From DBTPYMT WHERE TRX_TYPE = 'RT' AND '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenForwardOnly
        If Not (rstBILL.EOF And rstBILL.BOF) Then
            MAXRCPTNO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        End If
        rstBILL.Close
        Set rstBILL = Nothing
        
        Set rstMaxRec = New ADODB.Recordset
        rstMaxRec.Open "Select MAX(CR_NO) From DBTPYMT", db, adOpenForwardOnly
        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
            i = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
            CRNO2 = i
        End If
        rstMaxRec.Close
        Set rstMaxRec = Nothing
        
        If lblcredit.Caption = "0" Then
            db.Execute "START TRANSACTION"
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'GI' ", db, adOpenStatic, adLockOptimistic, adCmdText
            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE.AddNew
                RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                RSTTRXFILE!TRX_TYPE = "RT"
                RSTTRXFILE!INV_TRX_TYPE = "GI"
                RSTTRXFILE!INV_NO = Val(txtBillNo.text)
                RSTTRXFILE!CR_NO = i
                RSTTRXFILE!REC_NO = MAXRCPTNO
            End If
            RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
            If GRDRECEIPT.rows <= 1 Then
                RSTTRXFILE!RCPT_AMT = 0
                RSTTRXFILE!REF_NO = ""
            Else
                RSTTRXFILE!RCPT_AMT = Val(GRDRECEIPT.TextMatrix(0, 0))
                RSTTRXFILE!REF_NO = Trim(GRDRECEIPT.TextMatrix(1, 0))
            End If
            If GRDRECEIPT.rows > 1 And Trim(GRDRECEIPT.TextMatrix(2, 0)) = "B" Then
                i = 1
                db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_VCH_NO = " & Val(txtBillNo.text) & " AND B_TRX_TYPE = 'GI' "
                db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'GI'"
                Set rstMaxRec = New ADODB.Recordset
                rstMaxRec.Open "Select MAX(TRX_NO) From BANK_TRX WHERE TRX_TYPE = 'CR' AND BILL_TRX_TYPE = 'RT' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' ", db, adOpenForwardOnly
                If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
                    i = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
                End If
                rstMaxRec.Close
                Set rstMaxRec = Nothing
            
                Dim BANKTRX As ADODB.Recordset
                Set BANKTRX = New ADODB.Recordset
                BANKTRX.Open "Select * From BANK_TRX", db, adOpenStatic, adLockOptimistic, adCmdText
                BANKTRX.AddNew
                BANKTRX!TRX_TYPE = "CR"
                BANKTRX!TRX_NO = i
                BANKTRX!BILL_TRX_TYPE = "RT"
                BANKTRX!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                BANKTRX!B_TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                BANKTRX!B_VCH_NO = Val(txtBillNo.text)
                BANKTRX!B_TRX_TYPE = "GI"
                BANKTRX!BANK_CODE = GRDRECEIPT.TextMatrix(4, 0)
                BANKTRX!BANK_NAME = GRDRECEIPT.TextMatrix(7, 0)
                'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
                BANKTRX!TRX_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
                BANKTRX!TRX_AMOUNT = Val(GRDRECEIPT.TextMatrix(0, 0))
                BANKTRX!ACT_CODE = DataList2.BoundText
                BANKTRX!ACT_NAME = DataList2.text
                'RSTTRXFILE!INV_DATE = LBLDATE.Caption
                BANKTRX!REF_NO = "BRANCH SALES - " & txtBillNo.text 'DataList2.Text
                BANKTRX!ENTRY_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
                BANKTRX!CHQ_DATE = Format(GRDRECEIPT.TextMatrix(5, 0), "DD/MM/YYYY")
                BANKTRX!BANK_FLAG = "Y"
                If GRDRECEIPT.TextMatrix(6, 0) = "N" Then
                    BANKTRX!check_flag = "N"
                Else
                    BANKTRX!check_flag = "Y"
                End If
                BANKTRX!CHQ_NO = Trim(GRDRECEIPT.TextMatrix(3, 0))
                
                'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
                BANKTRX.Update
                BANKTRX.Close
                Set BANKTRX = Nothing
                RSTTRXFILE!BANK_FLAG = "Y"
                RSTTRXFILE!B_TRX_TYPE = "CR"
                RSTTRXFILE!B_BILL_TRX_TYPE = "RT"
                RSTTRXFILE!B_TRX_NO = i
                RSTTRXFILE!B_TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                RSTTRXFILE!BANK_CODE = GRDRECEIPT.TextMatrix(4, 0)
                
                'RSTTRXFILE!C_TRX_TYPE = Null
                'RSTTRXFILE!C_REC_NO = Null
                'RSTTRXFILE!C_INV_TRX_TYPE = Null
                'RSTTRXFILE!C_INV_TYPE = Null
                'RSTTRXFILE!C_INV_NO = Null
            Else
                RSTTRXFILE!BANK_FLAG = "N"
                'RSTTRXFILE!B_TRX_TYPE = Null
                'RSTTRXFILE!B_TRX_NO = Null
                'RSTTRXFILE!B_BILL_TRX_TYPE = Null
                'RSTTRXFILE!B_TRX_YEAR = Null
                'RSTTRXFILE!BANK_CODE = Null

                RSTTRXFILE!C_TRX_TYPE = TRXTYPE
                RSTTRXFILE!C_REC_NO = RECNO
                RSTTRXFILE!C_INV_TRX_TYPE = INVTRXTYPE
                RSTTRXFILE!C_INV_TYPE = INVTYPE
                RSTTRXFILE!C_INV_NO = INVNO
            End If
            RSTTRXFILE!ACT_CODE = DataList2.BoundText
            RSTTRXFILE!ACT_NAME = DataList2.text
            RSTTRXFILE!INV_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
            RSTTRXFILE!INV_AMT = 0
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            db.Execute "COMMIT"
            
            Dim BillNO As Long
            Set rstBILL = New ADODB.Recordset
            rstBILL.Open "Select MAX(RCPT_NO) From TRNXRCPT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
            If Not (rstBILL.EOF And rstBILL.BOF) Then
                BillNO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
            End If
            rstBILL.Close
            Set rstBILL = Nothing
            
            Dim SEL_AMOUNT As Double
            
            If GRDRECEIPT.rows <= 1 Then
                SEL_AMOUNT = 0
            Else
                SEL_AMOUNT = Val(GRDRECEIPT.TextMatrix(0, 0))
            End If
            
            If SEL_AMOUNT <= 0 Then GoTo SKIP2
            BillNO = BillNO + 1
            db.Execute "START TRANSACTION"
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRNXRCPT ", db, adOpenStatic, adLockOptimistic, adCmdText
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "RT"
            RSTTRXFILE!RCPT_NO = BillNO
            RSTTRXFILE!INV_NO = Val(txtBillNo.text)
            RSTTRXFILE!INV_TRX_TYPE = "GI"
            RSTTRXFILE!INV_TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
            If SEL_AMOUNT > Val(lblnetamount.Caption) Then
                RSTTRXFILE!RCPT_AMOUNT = Val(lblnetamount.Caption)
            Else
                RSTTRXFILE!RCPT_AMOUNT = SEL_AMOUNT
            End If
            RSTTRXFILE!ACT_CODE = DataList2.BoundText
            RSTTRXFILE!ACT_NAME = DataList2.text
            RSTTRXFILE!RCPT_ENTRY_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE!REF_NO = ""
            RSTTRXFILE!INV_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
            RSTTRXFILE!CR_NO = CRNO2
            RSTTRXFILE!CR_TRX_TYPE = "DR"
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            db.Execute "COMMIT"
SKIP2:
            
            Dim RCVDAMOUNT As Double
            RCVDAMOUNT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRNXRCPT WHERE INV_NO = " & Val(txtBillNo.text) & " AND INV_TRX_TYPE = 'GI' AND INV_TRX_YEAR = '" & Year(MDIMAIN.DTFROM.Value) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                RCVDAMOUNT = RCVDAMOUNT + IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT)
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            db.Execute "START TRANSACTION"
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE TRX_TYPE = 'DR' AND CR_NO = " & CRNO2 & " AND INV_TRX_TYPE = 'GI' AND TRX_YEAR = '" & Year(MDIMAIN.DTFROM.Value) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE!RCVD_AMOUNT = RCVDAMOUNT
                RSTTRXFILE.Update
            End If
            db.Execute "COMMIT"
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        End If
    Else
    '    db.Execute "delete From DBTPYMT WHERE TRX_TYPE='DR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TRX_TYPE = 'GI'"
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'GI' "
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'GI' "
        db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_VCH_NO = " & Val(txtBillNo.text) & " AND B_TRX_TYPE = 'GI' "
    End If
    
'    E_DATE = Format(TXTINVDATE.Text, "MM/DD/YYYY")
'    If Day(E_DATE) <= 12 Then
'        DAY_DATE = Format(Month(E_DATE), "00")
'        MONTH_DATE = Format(Day(E_DATE), "00")
'        YEAR_DATE = Format(Year(E_DATE), "0000")
'        E_DATE = DAY_DATE & "/" & MONTH_DATE & "/" & YEAR_DATE
'    End If
'    E_DATE = Format(E_DATE, "MM/DD/YYYY")
'
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        RSTITEMMAST!Area = Trim(TXTAREA.Text)
'        RSTITEMMAST!KGST = Trim(TXTTIN.Text)
'        RSTITEMMAST!ADDRESS = Trim(TxtBillAddress.Text)
'        RSTITEMMAST.Update
'    End If
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    
    db.Execute "START TRANSACTION"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!TRX_TYPE = "GI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!gross_amt = Val(LBLGROSSAMT.Caption)
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
    Else
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!gross_amt = Val(LBLGROSSAMT.Caption)
    End If
    
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        RSTTRXFILE!Area = RSTITEMMAST!Area
'    End If
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    RSTTRXFILE!TIN = Trim(TXTTIN.text)
    RSTTRXFILE!CUST_IGST = lblIGST.Caption
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.text
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
    RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
    RSTTRXFILE!REF_NO = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "A"
        RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.text)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "P"
        RSTTRXFILE!DISCOUNT = Round(RSTTRXFILE!VCH_AMOUNT * Val(TXTTOTALDISC.text) / 100, 2)
    End If
    RSTTRXFILE!check_flag = "I"
    If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
    RSTTRXFILE!CFORM_NO = Time
    RSTTRXFILE!REMARKS = Left(DataList2.text, 50)
    RSTTRXFILE!DISC_PERS = Val(TXTTOTALDISC.text)
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!VEHICLE = Trim(TxtVehicle.text)
    RSTTRXFILE!D_ORDER = Trim(TxtOrder.text)
    RSTTRXFILE!PHONE = Trim(TxtPhone.text)
    RSTTRXFILE!FRIEGHT = Val(TxtFrieght.text)
    RSTTRXFILE!HANDLE = Val(Txthandle.text)
    RSTTRXFILE!Area = Trim(TXTAREA.text)
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!cr_days = Val(txtcrdays.text)
    RSTTRXFILE!BILL_NAME = Trim(TxtBillName.text)
    RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.text)
    txtcommi.Tag = ""
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        RSTTRXFILE!AGENT_CODE = CMBDISTI.BoundText
        RSTTRXFILE!AGENT_NAME = CMBDISTI.text
        For i = 1 To grdsales.rows - 1
            txtcommi.Tag = Val(txtcommi.Tag) + (Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)))
        Next i
        RSTTRXFILE!COMM_AMT = Val(txtcommi.Tag)
    Else
        RSTTRXFILE!AGENT_CODE = ""
        RSTTRXFILE!AGENT_NAME = ""
    End If
    
    Select Case cmbtype.ListIndex
        Case 1
            RSTTRXFILE!BILL_TYPE = "W"
        Case 2
            RSTTRXFILE!BILL_TYPE = "V"
        Case 3
            RSTTRXFILE!BILL_TYPE = "M"
        Case Else
            RSTTRXFILE!BILL_TYPE = "R"
    End Select
    If Val(TxtCN.text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.text)
    If GRDRECEIPT.rows <= 1 Or DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
        RSTTRXFILE!RCPT_AMOUNT = 0
        RSTTRXFILE!RCPT_REFNO = ""
        RSTTRXFILE!BANK_FLAG = "N"
        'RSTTRXFILE!CHQ_NO = Null
        'RSTTRXFILE!BANK_CODE = Null
        'RSTTRXFILE!BANK_NAME = Null
        'RSTTRXFILE!CHQ_DATE = Null
        RSTTRXFILE!CHQ_STATUS = "N"
    Else
        RSTTRXFILE!RCPT_AMOUNT = Val(GRDRECEIPT.TextMatrix(0, 0))
        RSTTRXFILE!RCPT_REFNO = Trim(GRDRECEIPT.TextMatrix(1, 0))
        If Trim(GRDRECEIPT.TextMatrix(2, 0)) = "B" Then
            RSTTRXFILE!BANK_FLAG = "Y"
            RSTTRXFILE!CHQ_NO = Trim(GRDRECEIPT.TextMatrix(3, 0))
            RSTTRXFILE!BANK_CODE = Trim(GRDRECEIPT.TextMatrix(4, 0))
            RSTTRXFILE!BANK_NAME = Trim(GRDRECEIPT.TextMatrix(7, 0))
            RSTTRXFILE!CHQ_DATE = Format(GRDRECEIPT.TextMatrix(5, 0), "DD/MM/YYYY")
            If GRDRECEIPT.TextMatrix(6, 0) = "Y" Then
                RSTTRXFILE!CHQ_STATUS = "Y"
            Else
                RSTTRXFILE!CHQ_STATUS = "N"
            End If
        Else
            RSTTRXFILE!BANK_FLAG = "N"
        End If
    End If
    RSTTRXFILE!BILL_FLAG = "Y"
    If chkTerms.Value = 1 And Trim(Terms1.text) <> "" Then
        RSTTRXFILE!TERMS = Trim(Terms1.text)
    Else
        RSTTRXFILE!TERMS = ""
    End If
    RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
    RSTTRXFILE!BR_NAME = CMBBRNCH.text
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.Execute "COMMIT"
    
''''    For i = 1 To grdsales.Rows - 1
''''        If Val(grdsales.TextMatrix(i, 6)) <> 0 Then
''''            Set RSTP_RATE = New ADODB.Recordset
''''            RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & grdsales.TextMatrix(i, 13) & "'", DB, adOpenStatic, adLockOptimistic, adCmdText
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
        
    Dim unbillflag As Boolean
    unbillflag = True
    If MDIMAIN.lblRemoveUbill.Caption <> "Y" Or Val(LBLRETAMT.Caption) <> 0 Then
        unbillflag = False
    Else
        For i = 1 To grdsales.rows - 1
            If grdsales.TextMatrix(i, 46) <> "Y" Then
                unbillflag = False
                Exit For
            End If
    '        Set RSTTRXFILE = New ADODB.Recordset
    '        RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y')", db, adOpenStatic, adLockReadOnly, adCmdText
    '        With RSTTRXFILE
    '            If Not (.EOF And .BOF) Then
    '                unbillflag = False
    '                RSTTRXFILE.Close
    '                Set RSTTRXFILE = Nothing
    '                Exit For
    '            End If
    '        End With
    '        RSTTRXFILE.Close
    '        Set RSTTRXFILE = Nothing
            
        Next i
    End If
'    If unbillflag = True Then
'        Dim LASTWOBILL As Long
'        Set rstBILL = New ADODB.Recordset
'        rstBILL.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'WO'", db, adOpenStatic, adLockReadOnly
'        If Not (rstBILL.EOF And rstBILL.BOF) Then
'            LASTWOBILL = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
'        End If
'        rstBILL.Close
'        Set rstBILL = Nothing
'        db.Execute "START TRANSACTION"
'        db.Execute "Update TRXMASTVAN set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO', REF_NO = '" & Val(txtBillNo.Text) & "' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
'        db.Execute "Update TRXSUBVAN set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
'        db.Execute "Update TRXFILEVAN set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
'
'        db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TRX_TYPE = 'GI'"
'        db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTWOBILL & ", B_TRX_TYPE = 'WO' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_TRX_TYPE = 'GI' AND B_VCH_NO = " & Val(txtBillNo.Text) & ""
'        db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TRX_TYPE = 'GI'"
'        db.Execute "Update CASHATRXFILE set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TRX_TYPE = 'GI'"
'        db.Execute "Update RTRXFILEVAN set TRX_TYPE= 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= 'GI' AND VCH_NO = " & Val(TxtCN.Text) & ""
'        'db.Execute "delete From RTRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= '" & Trim(LBLTYPE.Caption) & "' AND VCH_NO = " & Val(TxtCN.Text) & " "
'        db.Execute "COMMIT"
'    End If
    
'    For i = 1 To grdsales.Rows - 1
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'            RSTTRXFILE!CHECK_FLAG = "Y"
'            RSTTRXFILE.Update
'        End If
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
'    Next i
       
SKIP:
    i = 0
'    Set rstMaxRec = New ADODB.Recordset
'    rstMaxRec.Open "Select MAX(CR_NO) From CRDTPYMT", db, adOpenStatic, adLockReadOnly
'    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
'        i = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
'    End If
'    rstMaxRec.Close
'    Set rstMaxRec = Nothing
'
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM CRDTPYMT WHERE INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'DR'", db, adOpenStatic, adLockOptimistic, adCmdText
'    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        RSTITEMMAST.AddNew
'        RSTITEMMAST!TRX_TYPE = "DR"
'        RSTITEMMAST!CR_NO = i
'        RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
'        RSTITEMMAST!RCPT_AMOUNT = 0
'    End If
'    RSTITEMMAST!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'    RSTITEMMAST!INV_AMT = Val(lblnetamount.Caption)
'    If lblcredit.Caption = "0" Then
'        RSTITEMMAST!BAL_AMT = 0
'        RSTITEMMAST!CHECK_FLAG = "Y"
'    Else
'        RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption) - RSTITEMMAST!RCPT_AMOUNT
'        RSTITEMMAST!CHECK_FLAG = "N"
'    End If
'    RSTITEMMAST!PINV = ""
'    RSTITEMMAST!ACT_CODE = DataList2.BoundText
'    RSTITEMMAST.Update
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    
'    Set rstBILL = New ADODB.Recordset
'    rstBILL.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI'", db, adOpenStatic, adLockReadOnly
'    If Not (rstBILL.EOF And rstBILL.BOF) Then
'        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
'        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
'    End If
'    rstBILL.Close
'    Set rstBILL = Nothing
'    OLD_BILL = False
    
    TXTAREA.Clear
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select DISTINCT AREA From CUSTMAST ORDER BY AREA", db, adOpenForwardOnly
    Do Until rstBILL.EOF
        If Not IsNull(rstBILL!Area) Then TXTAREA.AddItem (rstBILL!Area)
        rstBILL.MoveNext
    Loop
    rstBILL.Close
    Set rstBILL = Nothing
    
    txtBillNo.text = 1
    LBLBILLNO.Caption = 1
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'GI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = Val(txtBillNo.text)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    OLD_BILL = False
    
    
    TXTAREA.text = ""
    TxtPhone.text = ""
    txtphsearch.text = ""
    TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
    lbladdress.Caption = ""
    LBLDNORCN.Caption = ""
    lblnetamount.Caption = ""
    lblOB.Caption = ""
    LBLFOT.Caption = ""
    LBLRETAMT.Caption = ""
    LBLPROFIT.Caption = ""
    LBLDATE.Caption = Date
    LBLTOTAL.Caption = ""
    LBLGROSSAMT.Caption = ""
    lblcomamt.text = ""
    TXTTOTALDISC.text = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
    TXTAMOUNT.text = ""
    LBLDISCAMT.Caption = ""
    lblbalance.Caption = ""
    Txtrcvd.text = ""
    TxtCN.text = ""
    grdsales.rows = 1
    TXTSLNO.text = 1
    M_EDIT = False
    NEW_BILL = True
    lbltime.Caption = Format(Time, "hh:mm AMPM")
    cmdRefresh.Enabled = False
    CMDEXIT.Enabled = True
    CMDPre.Caption = "<<Pre&vious (Ctrl + Left)"
    cmdnxt.Caption = "&Next>>      (Ctrl + Right)"
    CMDPRINT.Enabled = False
    
    CmdPrintA5.Enabled = False
    FRMEHEAD.Enabled = True
    TXTDEALER.Enabled = True
    TxtCode.Enabled = True
    'TXTTYPE.Text = 1
    'TXTDEALER.SetFocus
    LBLITEMCOST.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLNETCOST.Caption = ""
    LBLNETPROFIT.Caption = ""
    
    LBLSELPRICE.Caption = ""
    TXTQTY.Tag = ""
    TxtCode.text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblcredit.Caption = "1"
    txtcrdays.text = ""
    CMBDISTI.text = ""
    TxtBillAddress.text = ""
    TxtVehicle.text = ""
    TxtOrder.text = ""
    TxtFrieght.text = ""
    Txthandle.text = ""
    TxtBillName.text = ""
    txtOutstanding.text = ""
    TXTTIN.text = ""
    lblsubdealer.Caption = ""
    lblActAmt.Caption = ""
    cr_days = False
    CHANGE_ADDRESS = False
    cmdreturn.Enabled = False
    M_ADD = False
    TXTDEALER.text = ""
    txtphsearch.text = ""
    Terms1.text = Terms1.Tag
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    cmbtype.ListIndex = 0
    TXTTYPE.text = 1
    'TXTDEALER.SetFocus
    GRDRECEIPT.TextMatrix(0, 0) = 0
    GRDRECEIPT.rows = 1
    TxtBillName.text = "Cash"
    CMBBRNCH.text = ""
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTDEALER.Enabled = True
    'TXTDEALER.SetFocus
    'TxtBillName.SetFocus
    TXTSLNO.Enabled = False
    TXTDEALER.text = "CASH"
    DataList2.BoundText = "130000"
    'cmdreturn.Enabled = True
    TXTITEMCODE.Enabled = True
    'TXTITEMCODE.SetFocus
    TXTDEALER.SetFocus
    
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = -2147168237 Then
        'db.Execute "unlock tables"
    Else
        MsgBox err.Description
        'db.Execute "unlock tables"
    End If
    On Error Resume Next
    db.Execute "ROLLBACK"
    'db.Execute "COMMIT"
End Function

Private Function Print_A4()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim DN_CN_FLag As Boolean
    Dim i As Long
    Dim CN As Integer
    Dim DN As Integer
    Dim b As Integer
    Dim Num, Figre As Currency
    
    On Error GoTo ERRHAND
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenForwardOnly
    If RSTCOMPANY.RecordCount <> grdsales.rows - 1 Then
        MsgBox "There is an error. Please do verify it ", , "EzBiz"
'        RSTCOMPANY.Close
'        Set RSTCOMPANY = Nothing
'        Exit Function
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
'    If CMDDELIVERY.Enabled = True Then
'        If (MsgBox("Delivered Items Available... Do you want to add these Items too...", vbYesNo + vbDefaultButton2, "SALES") = vbYes) Then CmdDelivery_Click
'    End If
    
'    If CMDSALERETURN.Enabled = True Then
'        If (MsgBox("Returned Items Available... Do you want to add these Items too...", vbYesNo, "SALES") = vbYes) Then CMDSALERETURN_Click
'    End If
    
    DN = 0
    CN = 0
    b = 0
    DN_CN_FLag = False
    
    txtOutstanding.text = ""
    Dim m_Rcpt_Amt As Double
    Dim Rcptamt As Double
    Dim m_OP_Bal As Double
    Dim m_Bal_Amt As Double
    
    m_Rcpt_Amt = 0
    m_OP_Bal = 0
    m_Bal_Amt = 0
    Rcptamt = 0
    
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompAddress4, CompAddress5, CompTin, CompCST, BIL_PRE, BILL_SUF, DL, ML, DL1, DL2, INV_TERMS, INV_MSG, BANK_DET, PAN_NO, OS_FLAG As String
    Dim QtnTerms, QtnTerms1, QtnTerms2, QtnTerms3, QtnTerms4 As String
    
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
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
        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        If Trim(TxtVehicle.text) = "" Then TxtVehicle.text = IIf(IsNull(RSTCOMPANY!VEHICLE), "", RSTCOMPANY!VEHICLE)
        INV_TERMS = IIf(IsNull(RSTCOMPANY!INV_TERMS) Or RSTCOMPANY!INV_TERMS = "", "", RSTCOMPANY!INV_TERMS)
        INV_MSG = IIf(IsNull(RSTCOMPANY!INV_MSGS) Or RSTCOMPANY!INV_MSGS = "", "", RSTCOMPANY!INV_MSGS)
        BANK_DET = IIf(IsNull(RSTCOMPANY!bank_details) Or RSTCOMPANY!bank_details = "", "", RSTCOMPANY!bank_details)
        PAN_NO = IIf(IsNull(RSTCOMPANY!PAN_NO) Or RSTCOMPANY!PAN_NO = "", "", RSTCOMPANY!PAN_NO)
        If thermalprn = True Then
            OS_FLAG = IIf(IsNull(RSTCOMPANY!OSPTY_FLAG) Or RSTCOMPANY!OSPTY_FLAG = "", "", RSTCOMPANY!OSPTY_FLAG)
        Else
            OS_FLAG = IIf(IsNull(RSTCOMPANY!OSB2B_FLAG) Or RSTCOMPANY!OSB2B_FLAG = "", "", RSTCOMPANY!OSB2B_FLAG)
        End If
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
    
    If Not OS_FLAG = "Y" Then GoTo SKIP_OUT
    If DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select OPEN_DB from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            m_OP_Bal = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
               
        Set RSTTRXFILE = New ADODB.Recordset
        'RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' AND NOT(TRX_TYPE= 'RT' AND INV_TRX_TYPE ='GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & ") ", db, adOpenStatic, adLockReadOnly, adCmdText
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE <= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE= 'RT' AND INV_TRX_TYPE ='GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        If GRDRECEIPT.rows > 1 Then Rcptamt = GRDRECEIPT.TextMatrix(0, 0)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND NOT(INV_TRX_TYPE = 'GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & ") AND (TRX_TYPE = 'DR' OR TRX_TYPE = 'DB') AND INV_DATE <= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                Case Else
                    m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        txtOutstanding.text = Round((m_OP_Bal + m_Bal_Amt) - (m_Rcpt_Amt), 0)
    End If
SKIP_OUT:
'    Set RSTTRXFILE = New ADODB.Recordset
'    If OLD_BILL = False And Val(txtBillNo.Text) <> 1 Then
'        RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE VCH_NO= (SELECT MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'GI')", db, adOpenStatic, adLockOptimistic, adCmdText
'        db.Execute "Lock tables TRXMASTVAN write"
'        txtBillNo.Text = RSTTRXFILE!VCH_NO + 1
'        LBLBILLNO.Caption = Val(txtBillNo.Text)
'        db.BeginTrans
'        RSTTRXFILE.AddNew
'        RSTTRXFILE!TRX_TYPE = "GI"
'        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'        RSTTRXFILE!VCH_NO = txtBillNo.Text
'        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
'        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
'        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
'        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
'    Else
'        RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
'        db.Execute "Lock tables TRXMASTVAN write"
'        db.BeginTrans
'        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'            RSTTRXFILE.AddNew
'            RSTTRXFILE!TRX_TYPE = "GI"
'            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'            RSTTRXFILE!VCH_NO = txtBillNo.Text
'            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
'            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
'            RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
'            RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
'        End If
'    End If
'    RSTTRXFILE!TIN = Trim(TXTTIN.Text)
'    RSTTRXFILE!CUST_IGST = lblIGST.Caption
'    RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
'    RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
'    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'    RSTTRXFILE!ACT_CODE = DataList2.BoundText
'    RSTTRXFILE!ACT_NAME = DataList2.Text
'    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
'    RSTTRXFILE!DISC_PERS = Val(TXTTOTALDISC.Text)
'    RSTTRXFILE!ADD_AMOUNT = 0
'    RSTTRXFILE!ROUNDED_OFF = 0
'    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
'    RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
'    If Val(TxtCN.Text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.Text)
'    RSTTRXFILE!BILL_FLAG = "Y"
'    If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
'        RSTTRXFILE!TERMS = Trim(Terms1.Text)
'    Else
'        RSTTRXFILE!TERMS = ""
'    End If
'    RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
'    RSTTRXFILE!BR_NAME = CMBBRNCH.Text
'    RSTTRXFILE!cr_days = Val(txtcrdays.Text)
'    Select Case cmbtype.ListIndex
'        Case 3
'            RSTTRXFILE!BILL_TYPE = "M"
'        Case 1
'            RSTTRXFILE!BILL_TYPE = "W"
'        Case 2
'            RSTTRXFILE!BILL_TYPE = "V"
'        Case 4
'            RSTTRXFILE!BILL_TYPE = "5"
'        Case 5
'            RSTTRXFILE!BILL_TYPE = "6"
'        Case 6
'            RSTTRXFILE!BILL_TYPE = "7"
'        Case Else
'            RSTTRXFILE!BILL_TYPE = "R"
'    End Select
'    RSTTRXFILE.Update
'    db.CommitTrans
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
'    OLD_BILL = True
'    db.Execute "unlock tables"
    
    'db.Execute "delete from TEMPTRXFILE WHERE VCH_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'GI' AND TRX_YEAR = '" & Year(MDIMAIN.DTFROM.value) & "' "
    db.Execute "delete from TEMPTRXFILE WHERE VCH_NO = " & Val(txtBillNo.text) & " "
    'Dim RSTUNBILL As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        
        If thermalprn = False And grdsales.TextMatrix(i, 46) = "Y" Then GoTo SKIP_UNBILL
'            Set RSTUNBILL = New ADODB.Recordset
'            RSTUNBILL.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "' AND UN_BILL = 'Y'", db, adOpenStatic, adLockReadOnly, adCmdText
'            With RSTUNBILL
'                If Not (.EOF And .BOF) Then
'                    RSTUNBILL.Close
'                    Set RSTUNBILL = Nothing
'                    GoTo SKIP_UNBILL
'                End If
'            End With
'            RSTUNBILL.Close
'            Set RSTUNBILL = Nothing
'            GoTo SKIP_UNBILL
'        End If
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "GI"
        'RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 13)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 5))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!kfc_tax = 0
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        Select Case cmbtype.ListIndex
            Case 3
                RSTTRXFILE!VCH_DESC = "M"
            Case 1
                RSTTRXFILE!VCH_DESC = "W"
            Case 2
                RSTTRXFILE!VCH_DESC = "V"
            Case 4
                RSTTRXFILE!VCH_DESC = "5"
            Case 5
                RSTTRXFILE!VCH_DESC = "6"
            Case 6
                RSTTRXFILE!VCH_DESC = "7"
            Case Else
                RSTTRXFILE!VCH_DESC = "R"
        End Select
        'RSTTRXFILE!VCH_DESC = "" '"Issued to     " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(i, 17))
        'RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!C_USER_ID = IIf(IsNull(RSTITEMMAST!REMARKS), "", Left(RSTITEMMAST!REMARKS, 8))
            RSTTRXFILE!MFGR = IIf(IsNull(RSTITEMMAST!ITEM_MAL), "", RSTITEMMAST!ITEM_MAL)
            RSTTRXFILE!M_USER_ID = IIf(IsNull(RSTITEMMAST!FULL_PACK), RSTTRXFILE!PACK_TYPE, RSTITEMMAST!FULL_PACK)
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        
        Select Case grdsales.TextMatrix(i, 19)
            Case "DN"
                DN_CN_FLag = True
                RSTTRXFILE!CST = 1
            Case "CN"
                DN_CN_FLag = True
                RSTTRXFILE!CST = 2
            Case Else
                RSTTRXFILE!CST = 0
        End Select
        
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        If grdsales.TextMatrix(i, 38) = "" Then
            'RSTTRXFILE!EXP_DATE = Null
        Else
            RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(i, 38))) & "/" & Trim(grdsales.TextMatrix(i, 38))
        End If
        RSTTRXFILE!RETAILER_PRICE = Val(grdsales.TextMatrix(i, 39))
        RSTTRXFILE!CESS_PER = Val(grdsales.TextMatrix(i, 40))
        RSTTRXFILE!cess_amt = Val(grdsales.TextMatrix(i, 41))
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 7)) 'Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 6)) 'Val(grdsales.TextMatrix(i, 22))
        If Tax_Print = False Then
            RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        Else
            RSTTRXFILE!SALES_TAX = Val(TxtTaxPrint.text)
        End If
        If Trim(grdsales.TextMatrix(i, 33)) = "" Then
            RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        Else
            RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 33)
        End If
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(i, 24))
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 27))
        RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(i, 28) = "", Null, grdsales.TextMatrix(i, 28))
        RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(i, 29)
        If Trim(grdsales.TextMatrix(i, 30)) = "" Then
            RSTTRXFILE!PACK_TYPE = "Nos"
        Else
            RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 30)
        End If
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(i, 26)
        RSTTRXFILE!ITEM_SPEC = Trim(grdsales.TextMatrix(i, 44))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
'        RSTTRXFILE!C_USER_ID = "SM"
'        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
'        Dim RSTITEMMAST As ADODB.Recordset
'        Set RSTITEMMAST = New ADODB.Recordset
'        RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
'        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'            RSTTRXFILE!Area = RSTITEMMAST!Area
'        End If
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.Update
SKIP_UNBILL:
    Next i
    
    Dim rstTRXMAST As ADODB.Recordset
    Set rstTRXMAST = New ADODB.Recordset
    rstTRXMAST.Open "SELECT * From RTRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(TxtCN.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstTRXMAST.EOF
        i = i + 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "XC"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = ""
        RSTTRXFILE!ITEM_CODE = rstTRXMAST!ITEM_CODE
        RSTTRXFILE!ITEM_NAME = rstTRXMAST!ITEM_NAME
        RSTTRXFILE!QTY = rstTRXMAST!QTY
        RSTTRXFILE!MRP = rstTRXMAST!MRP
        RSTTRXFILE!PTR = rstTRXMAST!PTR
        RSTTRXFILE!SALES_PRICE = -rstTRXMAST!SALES_PRICE
        RSTTRXFILE!SALES_TAX = rstTRXMAST!SALES_TAX
        RSTTRXFILE!UNIT = rstTRXMAST!UNIT
        RSTTRXFILE!VCH_DESC = "" '"Returned From  " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = rstTRXMAST!REF_NO
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = "V"
        RSTTRXFILE!MFGR = rstTRXMAST!MFGR
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = -rstTRXMAST!TRX_TOTAL
        RSTTRXFILE!LINE_DISC = 0 'rsttrxmast!LINE_DISC
        RSTTRXFILE!SCHEME = 0
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!P_RETAIL = -(rstTRXMAST!SALES_PRICE + (rstTRXMAST!SALES_PRICE * rstTRXMAST!SALES_TAX / 100))
        RSTTRXFILE!P_RETAILWOTAX = -rstTRXMAST!SALES_PRICE
        RSTTRXFILE!SALE_1_FLAG = ""
        RSTTRXFILE!COM_AMT = 0
        RSTTRXFILE!LOOSE_PACK = 1
        RSTTRXFILE!WARRANTY = 0
        RSTTRXFILE!WARRANTY_TYPE = ""
        RSTTRXFILE!PACK_TYPE = rstTRXMAST!PACK_TYPE
        'RSTTRXFILE!LOOSE_FLAG = rstTRXMAST!LOOSE_FLAG
        RSTTRXFILE!COM_FLAG = "N"
        RSTTRXFILE!ITEM_COST = rstTRXMAST!ITEM_COST
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "" '"SM"
                
        RSTTRXFILE.Update
        
        rstTRXMAST.MoveNext
    Loop
    rstTRXMAST.Close
    Set rstTRXMAST = Nothing
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim RSTtax As ADODB.Recordset
    Dim TaxAmt As Double
    Dim taxableamt As Double
    Dim Taxsplit As String
    TaxAmt = 0
    taxableamt = 0
    Taxsplit = ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT DISTINCT SALES_TAX From TEMPTRXFILE WHERE TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " order by SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        TaxAmt = 0
        taxableamt = 0
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select * From TEMPTRXFILE WHERE SALES_TAX = " & RSTTRXFILE!SALES_TAX & " AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            If OPTDISCPERCENT.Value = True Then
                grdtmp.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * Val(TXTTOTALDISC.text) / 100)
            Else
                grdtmp.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
            End If
            
            taxableamt = taxableamt + Val(grdtmp.Tag) * Val(RSTtax!QTY)
            TaxAmt = TaxAmt + (Val(grdtmp.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
        'Taxsplit = Taxsplit & "GST " & RSTTRXFILE!SALES_TAX & "%: " & "Taxable: " & Format(Round(TaxableAmt, 2), "0.00") & " Tax: " & Format(Round(TaxAmt, 2), "0.00") & "|"
        Taxsplit = Taxsplit & "Taxable: " & Format(Round(taxableamt, 2), "0.00") & " SGST " & RSTTRXFILE!SALES_TAX / 2 & "%: Tax: " & Format(Round(TaxAmt / 2, 2), "0.00") & " CGST " & RSTTRXFILE!SALES_TAX / 2 & "%: Tax: " & Format(Round(TaxAmt / 2, 2), "0.00") & "|"
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    'Call ReportGeneratION_vpestimate
    LBLFOT.Tag = ""
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = True Then
        If MsgBox("You do not have any permission to modify this further. Are you sure to print?", vbYesNo, "BILL..") = vbNo Then Exit Function
    Else
        Screen.MousePointer = vbHourglass
        Sleep (150)
    End If
              
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        DL1 = IIf(IsNull(RSTCOMPANY!DL_NO), "", Trim(RSTCOMPANY!DL_NO))
        DL2 = IIf(IsNull(RSTCOMPANY!REMARKS), "", Trim(RSTCOMPANY!REMARKS))
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
                
    NEW_BILL = False
    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0)) - Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 2)), 2)
'    Figre = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0))
'    Num = Abs(Figre)
'    If Figre < 0 Then
'        LBLFOT.Tag = "(-)Rupees " & Words_1_all(Num) & " Only"
'    ElseIf Figre > 0 Then
'        LBLFOT.Tag = "(Rupees " & Words_1_all(Num) & " Only)"
'    End If
    If thermalprn = True Then
        If MDIMAIN.LBLDMPTHERMAL.Caption = "Y" Then
            Call ReportGeneratION_WO
            On Error GoTo CLOSEFILE
            Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
            If err.Number = 55 Then
                Close #1
                Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
            End If
            On Error GoTo ERRHAND
        
            Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
            Print #1, "EXIT"
            Close #1
        
            '//HERE write the proper path where your command.com file exist
            Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & Rptpath & "REPO.BAT N", vbHide
            Exit Function
        End If
        LBLFOT.Tag = ""
        Screen.MousePointer = vbHourglass
        Sleep (300)
                                        
        NEW_BILL = False
        lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0)) - Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 2)), 2)
        Figre = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.text) + Val(Txthandle.text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0))
        Num = Abs(Figre)
        If Figre < 0 Then
            LBLFOT.Tag = "(-)Rupees " & Words_1_all(Num) & " Only"
        ElseIf Figre > 0 Then
            LBLFOT.Tag = "(Rupees " & Words_1_all(Num) & " Only)"
        End If
        If Small_Print = True Then
            'ReportNameVar = Rptpath & "rptbillretail"
            ReportNameVar = Rptpath & "RPTOUTPASS"
        Else
            ReportNameVar = Rptpath & "RPTOUTPASS"
        End If
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        For i = 1 To Report.Database.Tables.COUNT
            Report.Database.Tables.Item(i).SetLogOnInfo strConnection
            If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                If Report.Database.Tables(i).Name = "TRXFILE" Then
                    Set oRs = New ADODB.Recordset
                    Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
                    Report.Database.Tables(i).SetDataSource oRs, 3
                    Set oRs = Nothing
                Else
                    Set oRs = New ADODB.Recordset
                    Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
                    Report.Database.Tables(i).SetDataSource oRs, 3
                    Set oRs = Nothing
                End If
            End If
        Next i
        Report.RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & " and ({TRXFILE.TRX_TYPE}= 'GI' or {TRXFILE.TRX_TYPE}= 'XC'))"
        
        Report.DiscardSavedData
        Report.VerifyOnEveryPrint = True
        Set CRXFormulaFields = Report.FormulaFields
        For Each CRXFormulaField In CRXFormulaFields
            If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
            If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.text = "'" & CompName & "'"
            If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.text = "'" & CompAddress1 & "'"
            If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.text = "'" & CompAddress2 & "'"
            If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.text = "'" & CompAddress3 & "'"
            If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.text = "'" & CompAddress4 & "'"
            If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
            If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
            If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.text = "'" & TxtBillName.text & "'"
            If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.text = "'" & Trim(lbladdress.Caption) & "'"
            If TxtPhone.text = "" Then
                If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
            Else
                If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
                'If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "' & chr(13) & 'Ph: ' & '" & Trim(TxtPhone.Text) & "'"
            End If
            'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
            If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
            If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
    '        If Tax_Print = False Then
    '            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
    '        End If
            If chkTerms.Value = 1 And Trim(Terms1.text) <> "" Then
                If CRXFormulaField.Name = "{@condition}" Then CRXFormulaField.text = "'" & Trim(Terms1.text) & "'"
            End If
            If CRXFormulaField.Name = "{@salesman}" Then CRXFormulaField.text = "'" & CMBDISTI.text & "'"
            If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.text = "'" & INV_TERMS & "'"
            If CRXFormulaField.Name = "{@inv_msg}" Then CRXFormulaField.text = "'" & INV_MSG & "'"
            If CRXFormulaField.Name = "{@Terms}" Then CRXFormulaField.text = "'" & QtnTerms & "'"
            If CRXFormulaField.Name = "{@Terms1}" Then CRXFormulaField.text = "'" & QtnTerms1 & "'"
            If CRXFormulaField.Name = "{@Terms2}" Then CRXFormulaField.text = "'" & QtnTerms2 & "'"
            If CRXFormulaField.Name = "{@Terms3}" Then CRXFormulaField.text = "'" & QtnTerms3 & "'"
            If CRXFormulaField.Name = "{@Terms4}" Then CRXFormulaField.text = "'" & QtnTerms4 & "'"
            
            If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.text = "'GSTIN: ' & '" & Trim(TXTTIN.text) & "'"
            If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.text = "'" & TxtPhone.text & "'"
            If CRXFormulaField.Name = "{@VCH_NO}" Then
                Me.Tag = BIL_PRE & Format(Trim(txtBillNo.text), bill_for) & BILL_SUF
                CRXFormulaField.text = "'" & Me.Tag & "' "
            End If
            If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.text = "'" & Trim(TxtVehicle.text) & "'"
            If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.text = "'" & Trim(TxtOrder.text) & "'"
            If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.text = " " & Val(LBLDISCAMT.Caption) & " "
            If CRXFormulaField.Name = "{@TaxSplit}" Then CRXFormulaField.text = "'" & Taxsplit & "'"
    '            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
            If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.text = "'" & Trim(TxtCode.text) & "'"
            If CRXFormulaField.Name = "{@RcptAmt}" Then CRXFormulaField.text = " " & Rcptamt & " "
            If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.text = " " & Val(txtOutstanding.text) & " "
            If OptDiscAmt.Value = True Then
                If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'N'"
            Else
                If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'Y'"
            End If
            If CRXFormulaField.Name = "{@Frieght}" Then CRXFormulaField.text = "'" & Trim(lblFrieght.text) & "'"
            If CRXFormulaField.Name = "{@FC}" Then CRXFormulaField.text = " " & Val(TxtFrieght.text) & " "
            If CRXFormulaField.Name = "{@HANDLE}" Then CRXFormulaField.text = " '" & Trim(lblhandle.text) & "' "
            If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.text = " " & Val(Txthandle.text) & " "
            If CRXFormulaField.Name = "{@DISCPER}" Then CRXFormulaField.text = " " & Val(TXTTOTALDISC.text) & " "
            
            If Val(LBLRETAMT.Caption) = 0 Then
                If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'N' "
            Else
                If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'Y' "
            End If
            If CRXFormulaField.Name = "{@EXCHANGE}" Then CRXFormulaField.text = " " & Val(LBLRETAMT.Caption) & " "
            If lblcredit.Caption = "0" Then
                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Cash'"
            Else
                If Val(txtcrdays.text) > 0 Then
                    If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'" & txtcrdays.text & "'" & "' Days Credit'"
                Else
                    If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Credit'"
                End If
            End If
        Next
        
        Set Printer = Printers(thermalprinter)
        Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
        If MDIMAIN.LBLTHPREVIEW.Caption = "Y" Then
            'Preview
            frmreport.Caption = "BILL"
            Call GENERATEREPORT
            Screen.MousePointer = vbNormal
        Else
            '    '''No Preview
            Report.PrintOut (False)
            Set CRXFormulaFields = Nothing
            Set CRXFormulaField = Nothing
            Set crxApplication = Nothing
            Set Report = Nothing
            'Call cmdRefresh_Click
            Call AppendSale
        End If
        Screen.MousePointer = vbNormal
        
        Exit Function
    End If
    If Val(MDIMAIN.StatusBar.Panels(11).text) = 1 Then
        If Trim(lblIGST.Caption) <> "Y" Then
            If Small_Print = True Then
                'ReportNameVar = Rptpath & "rptbillretail"
                ReportNameVar = Rptpath & "RPTGSTBILLA51"
            Else
                ReportNameVar = Rptpath & "RPTGSTBILL1"
            End If
        Else
            ReportNameVar = Rptpath & "RPTGSTWBILL1"
        End If
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & " and ({TRXFILE.TRX_TYPE}= 'GI' or {TRXFILE.TRX_TYPE}= 'XC'))"
        If Val(lblnetamount.Caption) >= Val(MDIMAIN.LBLHSNSUM.Caption) Or Trim(lblIGST.Caption) = "Y" Then
            For i = 1 To Report.OpenSubreport("RPTHSN.rpt").Database.Tables.COUNT
                Report.OpenSubreport("RPTHSN.rpt").Database.Tables(i).SetLogOnInfo strConnection
                If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                    If Report.OpenSubreport("RPTHSN.rpt").Database.Tables(i).Name = "TRXFILE" Then
                        Set oRs = New ADODB.Recordset
                        Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
                        Report.OpenSubreport("RPTHSN.rpt").Database.SetDataSource oRs, 3, i
                        Set oRs = Nothing
                    Else
                        Set oRs = New ADODB.Recordset
                        Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTHSN.rpt").Database.Tables(i).Name & " ")
                        Report.OpenSubreport("RPTHSN.rpt").Database.SetDataSource oRs, 3, i
                        Set oRs = Nothing
                    End If
                End If
            Next i
            Report.OpenSubreport("RPTHSN.rpt").RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & " and ({TRXFILE.TRX_TYPE}= 'GI' or {TRXFILE.TRX_TYPE}= 'XC'))"
            Report.OpenSubreport("RPTHSN.rpt").DiscardSavedData
            Report.OpenSubreport("RPTHSN.rpt").VerifyOnEveryPrint = True
        
            Set CRXFormulaFields = Report.OpenSubreport("RPTHSN.rpt").FormulaFields
            Report.OpenSubreport("RPTHSN.rpt").DiscardSavedData
            For Each CRXFormulaField In CRXFormulaFields
                If OptDiscAmt.Value = True Then
                    If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'N'"
                Else
                    If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'Y'"
                End If
                If CRXFormulaField.Name = "{@DISCPER}" Then CRXFormulaField.text = " " & Val(TXTTOTALDISC.text) & " "
            Next
        End If
        
        
        For i = 1 To Report.Database.Tables.COUNT
            Report.Database.Tables.Item(i).SetLogOnInfo strConnection
            If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                If Report.Database.Tables(i).Name = "TRXFILE" Then
                    Set oRs = New ADODB.Recordset
                    Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
                    Report.Database.Tables(i).SetDataSource oRs, 3
                    Set oRs = Nothing
                Else
                    Set oRs = New ADODB.Recordset
                    Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
                    Report.Database.Tables(i).SetDataSource oRs, 3
                    Set oRs = Nothing
                End If
            End If
        Next i
        Report.DiscardSavedData
        Report.VerifyOnEveryPrint = True
        
        Set CRXFormulaFields = Report.FormulaFields
        For Each CRXFormulaField In CRXFormulaFields
            If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
            If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.text = "'" & CompName & "'"
            If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.text = "'" & CompAddress1 & "'"
            If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.text = "'" & CompAddress2 & "'"
            If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.text = "'" & CompAddress3 & "'"
            If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.text = "'" & CompAddress4 & "'"
            If CRXFormulaField.Name = "{@Comp_Address5}" Then CRXFormulaField.text = "'" & CompAddress5 & "'"
            If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
            If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
            If CRXFormulaField.Name = "{@DL}" Then CRXFormulaField.text = "'" & DL & "'"
            If CRXFormulaField.Name = "{@HSNSUM_FLAG}" Then
                If Val(lblnetamount.Caption) >= Val(MDIMAIN.LBLHSNSUM.Caption) Or Trim(lblIGST.Caption) = "Y" Then
                    CRXFormulaField.text = "'Y'"
                Else
                    CRXFormulaField.text = "'N'"
                End If
            End If
            If CRXFormulaField.Name = "{@ML}" Then CRXFormulaField.text = "'" & ML & "'"
            If CRXFormulaField.Name = "{@salesman}" Then CRXFormulaField.text = "'" & CMBDISTI.text & "'"
            If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.text = "'" & INV_TERMS & "'"
            If CRXFormulaField.Name = "{@inv_msg}" Then CRXFormulaField.text = "'" & INV_MSG & "'"
            If CRXFormulaField.Name = "{@Terms}" Then CRXFormulaField.text = "'" & QtnTerms & "'"
            If CRXFormulaField.Name = "{@Terms1}" Then CRXFormulaField.text = "'" & QtnTerms1 & "'"
            If CRXFormulaField.Name = "{@Terms2}" Then CRXFormulaField.text = "'" & QtnTerms2 & "'"
            If CRXFormulaField.Name = "{@Terms3}" Then CRXFormulaField.text = "'" & QtnTerms3 & "'"
            If CRXFormulaField.Name = "{@Terms4}" Then CRXFormulaField.text = "'" & QtnTerms4 & "'"
            If CRXFormulaField.Name = "{@TaxSplit}" Then CRXFormulaField.text = "'" & Taxsplit & "'"
            If CRXFormulaField.Name = "{@Rcvd}" Then CRXFormulaField.text = " " & Val(Txtrcvd.text) & " "
            If CRXFormulaField.Name = "{@BalAmt}" Then CRXFormulaField.text = " " & Abs(Val(lblbalance.Caption)) & " "
            If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.text = "'" & BANK_DET & "'"
            If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.text = "'" & PAN_NO & "'"
            If CRXFormulaField.Name = "{@DL1}" Then CRXFormulaField.text = "'" & DL1 & "'"
            If CRXFormulaField.Name = "{@DL2}" Then CRXFormulaField.text = "'" & DL2 & "'"
            If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.text = "'" & TxtBillName.text & "'"
            If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.text = "'" & Trim(TXTDEALER.text) & "'"
            If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.text = "'" & lbladdress.Caption & "'"
            If TxtPhone.text = "" Then
                If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
            Else
                If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
                'If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "' & chr(13) & 'Ph: ' & '" & Trim(TxtPhone.Text) & "'"
            End If
            If lblIGST.Caption = "Y" Then
                If CRXFormulaField.Name = "{@IGSTFLAG}" Then CRXFormulaField.text = "'Y'"
            Else
                If CRXFormulaField.Name = "{@IGSTFLAG}" Then CRXFormulaField.text = "'N'"
            End If
            'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
            If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
            If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
    '        If Tax_Print = False Then
    '            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
    '        End If
            If chkTerms.Value = 1 And Trim(Terms1.text) <> "" Then
                If CRXFormulaField.Name = "{@condition}" Then CRXFormulaField.text = "'" & Trim(Terms1.text) & "'"
            End If
            If CRXFormulaField.Name = "{@Area}" Then CRXFormulaField.text = "'" & Trim(TXTAREA.text) & "'"
            If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.text = "'GSTIN: ' & '" & Trim(TXTTIN.text) & "'"
            If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.text = "'" & TxtPhone.text & "'"
            If CRXFormulaField.Name = "{@VCH_NO}" Then
                Me.Tag = BIL_PRE & Format(Trim(txtBillNo.text), bill_for) & BILL_SUF
                CRXFormulaField.text = "'" & Me.Tag & "' "
            End If
            If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.text = "'" & Trim(TxtVehicle.text) & "'"
            If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.text = "'" & Trim(TxtOrder.text) & "'"
            If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.text = " " & Val(LBLDISCAMT.Caption) & " "
    '            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
            If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.text = "'" & Trim(TxtCode.text) & "'"
            If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.text = " " & Val(txtOutstanding.text) & " "
            If OptDiscAmt.Value = True Then
                If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'N'"
            Else
                If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'Y'"
            End If
            If CRXFormulaField.Name = "{@RcptAmt}" Then CRXFormulaField.text = " " & Rcptamt & " "
            If CRXFormulaField.Name = "{@Frieght}" Then CRXFormulaField.text = "'" & Trim(lblFrieght.text) & "'"
            If CRXFormulaField.Name = "{@FC}" Then CRXFormulaField.text = " " & Val(TxtFrieght.text) & " "
            If CRXFormulaField.Name = "{@HANDLE}" Then CRXFormulaField.text = " '" & Trim(lblhandle.text) & "' "
            If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.text = " " & Val(Txthandle.text) & " "
            If CRXFormulaField.Name = "{@DISCPER}" Then CRXFormulaField.text = " " & Val(TXTTOTALDISC.text) & " "
            
            If Val(LBLRETAMT.Caption) = 0 Then
                If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'N' "
            Else
                If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'Y' "
            End If
            If CRXFormulaField.Name = "{@EXCHANGE}" Then CRXFormulaField.text = " " & Val(LBLRETAMT.Caption) & " "
            If lblcredit.Caption = "0" Then
                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Cash'"
            Else
                If Val(txtcrdays.text) > 0 Then
                    If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'" & txtcrdays.text & "'" & "' Days Credit'"
                Else
                    If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Credit'"
                End If
            End If
        Next
    Else
        If Trim(lblIGST.Caption) <> "Y" Then
            'GENERAL
            If Small_Print = True Then
                'ReportNameVar = Rptpath & "rptbillretail"
                ReportNameVar = Rptpath & "RPTGSTBILLA5"
            Else
                ReportNameVar = Rptpath & "rptGSTBILL"
            End If
            
            'Pharma
'            If cmbtype.ListIndex = 1 Then
'                'ReportNameVar = Rptpath & "rptbillretail"
'                ReportNameVar = Rptpath & "RPTGSTBILL"
'            Else
'                ReportNameVar = Rptpath & "rptGSTBILL"
'            End If
        Else
            If Small_Print = True Then
                ReportNameVar = Rptpath & "RPTGSTWBILLA5"
            Else
                ReportNameVar = Rptpath & "RPTGSTWBILL"
            End If
        End If
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        If Val(MDIMAIN.StatusBar.Panels(11).text) > 4 Then MDIMAIN.StatusBar.Panels(11).text = 4
        Set CRXFormulaFields = Report.FormulaFields
        For Each CRXFormulaField In CRXFormulaFields
            If CRXFormulaField.Name = "{@Copies}" Then CRXFormulaField.text = " " & Val(MDIMAIN.StatusBar.Panels(11).text) & " "
        Next
        For i = 1 To Report.OpenSubreport("RPTBILL1.rpt").Database.Tables.COUNT
            Report.OpenSubreport("RPTBILL1.rpt").Database.Tables(i).SetLogOnInfo strConnection
            If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                If Report.OpenSubreport("RPTBILL1.rpt").Database.Tables(i).Name = "TRXFILE" Then
                    Set oRs = New ADODB.Recordset
                    Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
                    Report.OpenSubreport("RPTBILL1.rpt").Database.SetDataSource oRs, 3, i
                    Set oRs = Nothing
                Else
                    Set oRs = New ADODB.Recordset
                    Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTBILL1.rpt").Database.Tables(i).Name & " ")
                    Report.OpenSubreport("RPTBILL1.rpt").Database.SetDataSource oRs, 3, i
                    Set oRs = Nothing
                End If
            End If
        Next i
        Report.OpenSubreport("RPTBILL1.rpt").RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & " and ({TRXFILE.TRX_TYPE}= 'GI' or {TRXFILE.TRX_TYPE}= 'XC'))"
        Report.OpenSubreport("RPTBILL1.rpt").DiscardSavedData
        Report.OpenSubreport("RPTBILL1.rpt").VerifyOnEveryPrint = True
        
        If Val(MDIMAIN.StatusBar.Panels(11).text) > 1 Then
            For i = 1 To Report.OpenSubreport("RPTBILL2.rpt").Database.Tables.COUNT
                Report.OpenSubreport("RPTBILL2.rpt").Database.Tables(i).SetLogOnInfo strConnection
                If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                    If Report.OpenSubreport("RPTBILL2.rpt").Database.Tables(i).Name = "TRXFILE" Then
                        Set oRs = New ADODB.Recordset
                        Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
                        Report.OpenSubreport("RPTBILL2.rpt").Database.SetDataSource oRs, 3, i
                        Set oRs = Nothing
                    Else
                        Set oRs = New ADODB.Recordset
                        Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTBILL2.rpt").Database.Tables(i).Name & " ")
                        Report.OpenSubreport("RPTBILL2.rpt").Database.SetDataSource oRs, 3, i
                        Set oRs = Nothing
                    End If
                End If
            Next i
            Report.OpenSubreport("RPTBILL2.rpt").RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & " and ({TRXFILE.TRX_TYPE}= 'GI' or {TRXFILE.TRX_TYPE}= 'XC'))"
            Report.OpenSubreport("RPTBILL2.rpt").DiscardSavedData
            Report.OpenSubreport("RPTBILL2.rpt").VerifyOnEveryPrint = True
        End If
        
        If Val(MDIMAIN.StatusBar.Panels(11).text) > 2 Then
            For i = 1 To Report.OpenSubreport("RPTBILL3.rpt").Database.Tables.COUNT
                Report.OpenSubreport("RPTBILL3.rpt").Database.Tables(i).SetLogOnInfo strConnection
                If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                    If Report.OpenSubreport("RPTBILL3.rpt").Database.Tables(i).Name = "TRXFILE" Then
                        Set oRs = New ADODB.Recordset
                        Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
                        Report.OpenSubreport("RPTBILL3.rpt").Database.SetDataSource oRs, 3, i
                        Set oRs = Nothing
                    Else
                        Set oRs = New ADODB.Recordset
                        Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTBILL3.rpt").Database.Tables(i).Name & " ")
                        Report.OpenSubreport("RPTBILL3.rpt").Database.SetDataSource oRs, 3, i
                        Set oRs = Nothing
                    End If
                End If
            Next i
            Report.OpenSubreport("RPTBILL3.rpt").RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & " and ({TRXFILE.TRX_TYPE}= 'GI' or {TRXFILE.TRX_TYPE}= 'XC'))"
            Report.OpenSubreport("RPTBILL3.rpt").DiscardSavedData
            Report.OpenSubreport("RPTBILL3.rpt").VerifyOnEveryPrint = True
        End If
        If Val(MDIMAIN.StatusBar.Panels(11).text) > 3 Then
            For i = 1 To Report.OpenSubreport("RPTBILL4.rpt").Database.Tables.COUNT
                Report.OpenSubreport("RPTBILL4.rpt").Database.Tables(i).SetLogOnInfo strConnection
                If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                    If Report.OpenSubreport("RPTBILL4.rpt").Database.Tables(i).Name = "TRXFILE" Then
                        Set oRs = New ADODB.Recordset
                        Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
                        Report.OpenSubreport("RPTBILL4.rpt").Database.SetDataSource oRs, 3, i
                        Set oRs = Nothing
                    Else
                        Set oRs = New ADODB.Recordset
                        Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTBILL4.rpt").Database.Tables(i).Name & " ")
                        Report.OpenSubreport("RPTBILL4.rpt").Database.SetDataSource oRs, 3, i
                        Set oRs = Nothing
                    End If
                End If
            Next i
            Report.OpenSubreport("RPTBILL4.rpt").RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & " and ({TRXFILE.TRX_TYPE}= 'GI' or {TRXFILE.TRX_TYPE}= 'XC'))"
            Report.OpenSubreport("RPTBILL4.rpt").DiscardSavedData
            Report.OpenSubreport("RPTBILL4.rpt").VerifyOnEveryPrint = True
        End If
        For i = 1 To Val(MDIMAIN.StatusBar.Panels(11).text)
            Set CRXFormulaFields = Report.OpenSubreport("RPTBILL" & i & ".rpt").FormulaFields
            For Each CRXFormulaField In CRXFormulaFields
                If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
                If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.text = "'" & CompName & "'"
                If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.text = "'" & CompAddress1 & "'"
                If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.text = "'" & CompAddress2 & "'"
                If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.text = "'" & CompAddress3 & "'"
                If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.text = "'" & CompAddress4 & "'"
                If CRXFormulaField.Name = "{@Comp_Address5}" Then CRXFormulaField.text = "'" & CompAddress5 & "'"
                If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
                If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
                If CRXFormulaField.Name = "{@DL}" Then CRXFormulaField.text = "'" & DL & "'"
                If CRXFormulaField.Name = "{@ML}" Then CRXFormulaField.text = "'" & ML & "'"
                If CRXFormulaField.Name = "{@DL1}" Then CRXFormulaField.text = "'" & DL1 & "'"
                If CRXFormulaField.Name = "{@DL2}" Then CRXFormulaField.text = "'" & DL2 & "'"
                If CRXFormulaField.Name = "{@salesman}" Then CRXFormulaField.text = "'" & CMBDISTI.text & "'"
                If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.text = "'" & INV_TERMS & "'"
                If CRXFormulaField.Name = "{@inv_msg}" Then CRXFormulaField.text = "'" & INV_MSG & "'"
                If CRXFormulaField.Name = "{@Terms}" Then CRXFormulaField.text = "'" & QtnTerms & "'"
                If CRXFormulaField.Name = "{@Terms1}" Then CRXFormulaField.text = "'" & QtnTerms1 & "'"
                If CRXFormulaField.Name = "{@Terms2}" Then CRXFormulaField.text = "'" & QtnTerms2 & "'"
                If CRXFormulaField.Name = "{@Terms3}" Then CRXFormulaField.text = "'" & QtnTerms3 & "'"
                If CRXFormulaField.Name = "{@Terms4}" Then CRXFormulaField.text = "'" & QtnTerms4 & "'"
                If CRXFormulaField.Name = "{@TaxSplit}" Then CRXFormulaField.text = "'" & Taxsplit & "'"
                If CRXFormulaField.Name = "{@Rcvd}" Then CRXFormulaField.text = " " & Val(Txtrcvd.text) & " "
                If CRXFormulaField.Name = "{@BalAmt}" Then CRXFormulaField.text = " " & Abs(Val(lblbalance.Caption)) & " "
                If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.text = "'" & BANK_DET & "'"
                If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.text = "'" & PAN_NO & "'"
                If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.text = "'" & Trim(TxtBillName.text) & "'"
                If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.text = "'" & Trim(TXTDEALER.text) & "'"
                If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.text = "'" & Trim(lbladdress.Caption) & "'"
                If CRXFormulaField.Name = "{DLNO2}" Then CRXFormulaField.text = "'" & DL1 & "'"
                If CRXFormulaField.Name = "{DLNO}" Then CRXFormulaField.text = "'" & DL2 & "'"
                If TxtPhone.text = "" Then
                    If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
                Else
                    If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
                    'If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "' & chr(13) & 'Ph: ' & '" & Trim(TxtPhone.Text) & "'"
                End If
                If chkTerms.Value = 1 And Trim(Terms1.text) <> "" Then
                    If CRXFormulaField.Name = "{@condition}" Then CRXFormulaField.text = "'" & Trim(Terms1.text) & "'"
                End If
                If CRXFormulaField.Name = "{@Area}" Then CRXFormulaField.text = "'" & Trim(TXTAREA.text) & "'"
                'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
                If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
        '            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
        '            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
                If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
        '        If Tax_Print = False Then
        '            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
        '        End If
                If chkTerms.Value = 1 And Trim(Terms1.text) <> "" Then
                    If CRXFormulaField.Name = "{@condition}" Then CRXFormulaField.text = "'" & Trim(Terms1.text) & "'"
                End If
                If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.text = "'GSTIN: ' & '" & Trim(TXTTIN.text) & "'"
                If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.text = "'" & TxtPhone.text & "'"
                If CRXFormulaField.Name = "{@VCH_NO}" Then
                    Me.Tag = BIL_PRE & Format(Trim(txtBillNo.text), bill_for) & BILL_SUF
                    CRXFormulaField.text = "'" & Me.Tag & "' "
                End If
                If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.text = "'" & Trim(TxtVehicle.text) & "'"
                If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.text = "'" & Trim(TxtOrder.text) & "'"
                If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.text = " " & Val(LBLDISCAMT.Caption) & " "
        '            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
                If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.text = "'" & Trim(TxtCode.text) & "'"
                If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.text = " " & Val(txtOutstanding.text) & " "
                If OptDiscAmt.Value = True Then
                    If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'N'"
                Else
                    If CRXFormulaField.Name = "{@discflag}" Then CRXFormulaField.text = " 'Y'"
                End If
                If CRXFormulaField.Name = "{@RcptAmt}" Then CRXFormulaField.text = " " & Rcptamt & " "
                If CRXFormulaField.Name = "{@Frieght}" Then CRXFormulaField.text = "'" & Trim(lblFrieght.text) & "'"
                If CRXFormulaField.Name = "{@FC}" Then CRXFormulaField.text = " " & Val(TxtFrieght.text) & " "
                If CRXFormulaField.Name = "{@HANDLE}" Then CRXFormulaField.text = " '" & Trim(lblhandle.text) & "' "
                If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.text = " " & Val(Txthandle.text) & " "
                If CRXFormulaField.Name = "{@DISCPER}" Then CRXFormulaField.text = " " & Val(TXTTOTALDISC.text) & " "
                
                If Val(LBLRETAMT.Caption) = 0 Then
                    If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'N' "
                Else
                    If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'Y' "
                End If
                If CRXFormulaField.Name = "{@EXCHANGE}" Then CRXFormulaField.text = " " & Val(LBLRETAMT.Caption) & " "
                If lblcredit.Caption = "0" Then
                    If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Cash'"
                Else
                    If Val(txtcrdays.text) > 0 Then
                        If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'" & txtcrdays.text & "'" & "' Days Credit'"
                    Else
                        If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Credit'"
                    End If
                End If
            Next
        Next i
    End If
    
    If Small_Print = True Then
        Set Printer = Printers(billprinterA5)
    Else
        Set Printer = Printers(billprinter)
    End If
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
    If MDIMAIN.StatusBar.Panels(13).text = "Y" Then
        'Preview
        frmreport.Caption = "BILL"
        Call GENERATEREPORT
        Screen.MousePointer = vbNormal
    Else
        '    '''No Preview
        Report.PrintOut (False)
        Set CRXFormulaFields = Nothing
        Set CRXFormulaField = Nothing
        Set crxApplication = Nothing
        Set Report = Nothing
        'Call cmdRefresh_Click
        Call AppendSale
        Exit Function
    End If
    
SKIP:
    CMDEXIT.Enabled = False
    CMDPre.Caption = "Save && Pre&vious"
    cmdnxt.Caption = "Save && &Next"
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    TXTQTY.Enabled = False
    
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    txtretail.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTDISC.Enabled = False
    TxtDiscAmt.Enabled = False
    txtcommi.Enabled = False
    'OLD_BILL = True
    ''rptPRINT.Action = 1
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Function

Private Sub CmdTax_Click()
    If grdsales.rows <= 1 Then Exit Sub
    If Trim(TxtTaxPrint.text) = "" Then Exit Sub
    Tax_Print = True
    Call Generateprint
    TxtTaxPrint.text = ""
End Sub

Private Sub TxtFree_GotFocus()
    TXTFREE.SelStart = 0
    TXTFREE.SelLength = Len(TXTFREE.text)
    TXTFREE.Tag = Trim(TXTPRODUCT.text)
    Call CHANGEBOXCOLOR(TXTFREE, True)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.text) = 0 And Val(TXTFREE.text) = 0 Then
                TXTFREE.Enabled = True
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            End If
            If Val(TXTFREE.text) = 0 Then GoTo SKIP
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT CLOSE_QTY  From ITEMMASTVAN WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!CLOSE_QTY)) Then RSTTRXFILE!CLOSE_QTY = 0
                i = RSTTRXFILE!CLOSE_QTY / Val(LblPack.text)
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
'            If M_EDIT = False And (Val(TXTQTY.Text) + Val(TXTFREE.Text) > i) Then
'                MsgBox "AVAILABLE STOCK IS  " & i & " ", , "SALES"
'                TXTQTY.SelStart = 0
'                TXTQTY.SelLength = Len(TXTQTY.Text)
'                Exit Sub
'            End If
'            If i <> 0 And Val(TXTFREE.Text) <> 0 Then
                If M_EDIT = False And SERIAL_FLAG = True And (Val(TXTFREE.text) + Val(TXTQTY.text)) > (Val(lblactqty.Caption) / Val(LblPack.text)) Then
                    MsgBox "AVAILABLE STOCK IN THE PARTICULAR STOCK IS  " & (Val(lblactqty.Caption) / Val(LblPack.text)) & " ", , "SALES"
                    TXTQTY.SelStart = 0
                    TXTQTY.SelLength = Len(TXTQTY.text)
                    Exit Sub
                End If
                If M_EDIT = False And (Val(TXTFREE.text) + Val(TXTQTY.text)) > i Then
                    If SERIAL_FLAG = True Then
                        MsgBox "AVAILABLE STOCK IS  " & i & " ", , "SALES"
                        TXTFREE.SelStart = 0
                        TXTFREE.SelLength = Len(TXTFREE.text)
                        Exit Sub
                    End If
                    If (MsgBox("AVAILABLE STOCK IS  " & i & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                        'MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                        TXTQTY.SelStart = 0
                        TXTQTY.SelLength = Len(TXTQTY.text)
                        Exit Sub
                    End If
                End If
'            End If
            
SKIP:
            If UCase(TXTITEMCODE.text) = "OT" Then
                If TxtMRP.Visible = True Then
                    TxtMRP.Enabled = True
                    TxtMRP.SetFocus
                Else
                    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                        txtretail.Enabled = True
                        txtretail.SetFocus
                    Else
                        If Val(TXTTAX.text) = 0 Then
                            TXTTAX.Enabled = True
                            TXTTAX.SetFocus
                        Else
                            If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                                TXTRETAILNOTAX.Enabled = True
                                TXTRETAILNOTAX.SetFocus
                            Else
                                If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
                                    TXTRETAILNOTAX.Enabled = True
                                    TXTRETAILNOTAX.SetFocus
                                Else
                                    txtretail.Enabled = True
                                    txtretail.SetFocus
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                    txtretail.Enabled = True
                    txtretail.SetFocus
                Else
                    If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                        TXTRETAILNOTAX.Enabled = True
                        TXTRETAILNOTAX.SetFocus
                    Else
                        If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
                            TXTRETAILNOTAX.Enabled = True
                            TXTRETAILNOTAX.SetFocus
                        Else
                            txtretail.Enabled = True
                            txtretail.SetFocus
                        End If
                    End If
                End If
            End If
'            TXTFREE.Enabled = False
'            TXTTAX.Enabled = True
'            TXTTAX.SetFocus
         Case vbKeyEscape
            
            TXTFREE.Enabled = True
            TXTQTY.Enabled = True
            
            TXTQTY.SetFocus
        Case vbKeyDown
            If Val(TXTQTY.text) = 0 And Val(TXTFREE.text) = 0 Then Exit Sub
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

Private Sub TxtOrder_GotFocus()
    'If Trim(TxtOrder.Text) = "" Then TxtOrder.Text = "KL-04-N-8931"
    TxtOrder.SelStart = 0
    TxtOrder.SelLength = Len(TxtOrder.text)
    Call CHANGEBOXCOLOR(TxtOrder, True)
End Sub

Private Sub TxtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
                TxtPhone.SetFocus
            Else
               TxtVehicle.SetFocus
            End If
        Case vbKeyEscape
            cmbtype.SetFocus
    End Select

End Sub

Private Sub TxtOrder_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub CMBBRNCH_Change()
    
    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTITEMMAST = New ADODB.Recordset
    'RSTITEMMAST.Open "SELECT * FROM CUSTTRXFILE WHERE ACT_CODE = '" & Txtsuplcode.Text & "' and ACT_CODE <> '130000'", db, adOpenStatic, adLockReadOnly
    RSTITEMMAST.Open "SELECT * FROM CUSTTRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "' AND BR_CODE = '" & CMBBRNCH.BoundText & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        lbladdress.Caption = CMBBRNCH.text & Chr(13) & Trim(RSTITEMMAST!Address)
        'TxtBillName.Text = CMBBRNCH.Text
        TxtBillAddress.text = IIf(IsNull(RSTITEMMAST!Address), "", Trim(RSTITEMMAST!Address))
        TXTTIN.text = IIf(IsNull(RSTITEMMAST!KGST), "", Trim(RSTITEMMAST!KGST))
        TxtPhone.text = IIf(IsNull(RSTITEMMAST!TELNO), "", Trim(RSTITEMMAST!TELNO))
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CMBBRNCH_GotFocus()
    CMBBRNCH.text = ""
    If BR_FLAG = True Then
        BR_CODE.Open "Select *  from CUSTTRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "'  ORDER BY BR_NAME", db, adOpenStatic, adLockReadOnly
        BR_FLAG = False
    Else
        BR_CODE.Close
        BR_CODE.Open "Select *  from CUSTTRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "'  ORDER BY BR_NAME", db, adOpenStatic, adLockReadOnly
        BR_FLAG = False
    End If
    Set CMBBRNCH.RowSource = BR_CODE
    CMBBRNCH.ListField = "BR_NAME"
    CMBBRNCH.BoundColumn = "BR_CODE"
End Sub

Private Sub CMBBRNCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtBillName.SetFocus
            'FRMEHEAD.Enabled = False
            'TxtName1.Enabled = True
            'TxtName1.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Function FillDelivery_Grid()
    
    Dim i As Long
    Screen.MousePointer = vbHourglass
    i = 0
    grdEXPIRYLIST.TextMatrix(0, 0) = ""
    grdEXPIRYLIST.TextMatrix(0, 1) = "SL"
    grdEXPIRYLIST.TextMatrix(0, 2) = "DN #"
    grdEXPIRYLIST.TextMatrix(0, 3) = "DN DATE"
    grdEXPIRYLIST.TextMatrix(0, 4) = "ITEM_CODE"
    grdEXPIRYLIST.TextMatrix(0, 5) = "ITEM_NAME"
    grdEXPIRYLIST.TextMatrix(0, 6) = "QTY"
    grdEXPIRYLIST.TextMatrix(0, 7) = "LINE NO"
    grdEXPIRYLIST.TextMatrix(0, 8) = "FLAG"
    
    grdEXPIRYLIST.ColWidth(0) = 300
    grdEXPIRYLIST.ColWidth(1) = 400
    grdEXPIRYLIST.ColWidth(2) = 800
    grdEXPIRYLIST.ColWidth(3) = 1100
    grdEXPIRYLIST.ColWidth(4) = 0
    grdEXPIRYLIST.ColWidth(5) = 3500
    grdEXPIRYLIST.ColWidth(6) = 700
    grdEXPIRYLIST.ColWidth(7) = 0
    grdEXPIRYLIST.ColWidth(8) = 0
    
    grdEXPIRYLIST.ColAlignment(0) = 4
    grdEXPIRYLIST.ColAlignment(1) = 4
    grdEXPIRYLIST.ColAlignment(2) = 4
    grdEXPIRYLIST.ColAlignment(3) = 4
    'grdEXPIRYLIST.ColAlignment(4) = 4
    grdEXPIRYLIST.ColAlignment(5) = 1
    grdEXPIRYLIST.ColAlignment(6) = 4
    'grdEXPIRYLIST.ColAlignment(7) = 4
    'grdEXPIRYLIST.ColAlignment(8) = 4
    
    
    Dim RSTTRXFILE As ADODB.Recordset
    On Error GoTo ERRHAND
    i = 0
    grdEXPIRYLIST.FixedRows = 0
    grdEXPIRYLIST.rows = 1
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y' AND TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        i = i + 1
        grdEXPIRYLIST.rows = grdEXPIRYLIST.rows + 1
        grdEXPIRYLIST.FixedRows = 1
        'grdEXPIRYLIST.TextMatrix(i, 0) = i
        grdEXPIRYLIST.TextMatrix(i, 1) = i
        grdEXPIRYLIST.TextMatrix(i, 2) = RSTTRXFILE!VCH_NO
        grdEXPIRYLIST.TextMatrix(i, 3) = IIf(IsNull(RSTTRXFILE!VCH_DATE), "", Format(RSTTRXFILE!VCH_DATE, "DD/MM/YYYY"))
        grdEXPIRYLIST.TextMatrix(i, 4) = IIf(IsNull(RSTTRXFILE!ITEM_CODE), "", RSTTRXFILE!ITEM_CODE)
        grdEXPIRYLIST.TextMatrix(i, 5) = IIf(IsNull(RSTTRXFILE!ITEM_NAME), "", RSTTRXFILE!ITEM_NAME)
        grdEXPIRYLIST.TextMatrix(i, 6) = IIf(IsNull(RSTTRXFILE!QTY), "", RSTTRXFILE!QTY)
        grdEXPIRYLIST.TextMatrix(i, 7) = IIf(IsNull(RSTTRXFILE!LINE_NO), "", RSTTRXFILE!LINE_NO)
        grdEXPIRYLIST.TextMatrix(i, 8) = "N"
        With grdEXPIRYLIST
          .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
          Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
          .TextMatrix(i, 1) = i
        End With
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    grdEXPIRYLIST.Visible = True
    CHKSELECT.Value = 0
    grdcount.rows = 0
    Screen.MousePointer = vbNormal
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub CHKSELECT_Click()
    Dim i As Long
    If grdEXPIRYLIST.rows = 1 Then Exit Sub
    For i = 1 To grdEXPIRYLIST.rows - 1
        If CHKSELECT.Value = 1 Then
            With grdEXPIRYLIST
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picUnchecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
              .TextMatrix(.Row, 8) = "Y"
            End With
        Else
            With grdEXPIRYLIST
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
              .TextMatrix(.Row, 8) = "N"
            End With
        End If
    Next i
    Call fillcount
End Sub

Private Function fillcount()
    Dim i, n As Long
    
    grdcount.rows = 0
    i = 0
    On Error GoTo ERRHAND
    For n = 1 To grdEXPIRYLIST.rows - 1
        If grdEXPIRYLIST.TextMatrix(n, 8) = "Y" Then
            grdcount.rows = grdcount.rows + 1
            grdcount.TextMatrix(i, 0) = grdEXPIRYLIST.TextMatrix(n, 2)
            grdcount.TextMatrix(i, 1) = grdEXPIRYLIST.TextMatrix(n, 3)
            grdcount.TextMatrix(i, 2) = grdEXPIRYLIST.TextMatrix(n, 4)
            grdcount.TextMatrix(i, 3) = grdEXPIRYLIST.TextMatrix(n, 5)
            grdcount.TextMatrix(i, 4) = grdEXPIRYLIST.TextMatrix(n, 6)
            grdcount.TextMatrix(i, 5) = grdEXPIRYLIST.TextMatrix(n, 7)
            grdcount.TextMatrix(i, 6) = grdEXPIRYLIST.TextMatrix(n, 8)
            i = i + 1
        End If
    Next n
    Exit Function
ERRHAND:
    MsgBox err.Description
    
End Function

Private Function markitems()
    Dim i, n As Long
    
    i = 0
    On Error GoTo ERRHAND
    For n = 1 To grdEXPIRYLIST.rows - 1
        If grdEXPIRYLIST.TextMatrix(n, 8) = "Y" Then
            grdcount.rows = grdcount.rows + 1
            grdcount.TextMatrix(i, 0) = grdEXPIRYLIST.TextMatrix(n, 2)
            grdcount.TextMatrix(i, 1) = grdEXPIRYLIST.TextMatrix(n, 3)
            grdcount.TextMatrix(i, 2) = grdEXPIRYLIST.TextMatrix(n, 4)
            grdcount.TextMatrix(i, 3) = grdEXPIRYLIST.TextMatrix(n, 5)
            grdcount.TextMatrix(i, 4) = grdEXPIRYLIST.TextMatrix(n, 6)
            grdcount.TextMatrix(i, 5) = grdEXPIRYLIST.TextMatrix(n, 7)
            grdcount.TextMatrix(i, 6) = grdEXPIRYLIST.TextMatrix(n, 8)
            grdcount.TextMatrix(i, 7) = grdEXPIRYLIST.TextMatrix(n, 9)
            i = i + 1
        End If
    Next n
    Exit Function
ERRHAND:
    MsgBox err.Description
    
End Function

Private Sub grdEXPIRYLIST_Click()
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    If grdEXPIRYLIST.rows = 1 Then Exit Sub
    With grdEXPIRYLIST
        oldx = .Col
        oldy = .Row
        .Row = oldy: .Col = 0: .CellPictureAlignment = 4
            'If grdEXPIRYLIST.Col = 0 Then
                If grdEXPIRYLIST.CellPicture = picChecked Then
                    Set grdEXPIRYLIST.CellPicture = picUnchecked
                    '.Col = .Col + 2  ' I use data that is in column #1, usually an Index or ID #
                    'strTextCheck = .Text
                    ' When you de-select a CheckBox, we need to strip out the #
                    'strChecked = strChecked & strTextCheck & ","
                    ' Don't forget to strip off the trailing , before passing the string
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 8) = "Y"
                    Call fillcount
                Else
                    Set grdEXPIRYLIST.CellPicture = picChecked
                    '.Col = .Col + 2
                    'strTextCheck = .Text
                    'strChecked = Replace(strChecked, strTextCheck & ",", "")
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 8) = "N"
                    Call fillcount
                End If
            'End If
        .Col = oldx
        .Row = oldy
    End With
End Sub


Function FILL_BATCHGRID2()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME From RTRXFILEVAN  WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME From RTRXFILEVAN  WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
        BATCH_FLAG = False
    End If
    Set GRDPOPUP.DataSource = PHY_BATCH
    'TXTITEMCODE.Text = GRDPOPUP.Columns(26)
    item_change = True
    TXTPRODUCT.text = GRDPOPUP.Columns(6)
    item_change = False
    GRDPOPUP.Columns(0).Caption = "Serial No."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "VCH No"
    GRDPOPUP.Columns(3).Caption = "Line No"
    GRDPOPUP.Columns(4).Caption = "Trx Type"
    GRDPOPUP.Columns(7).Caption = "" '"Warranty"
    GRDPOPUP.Columns(8).Caption = ""
    GRDPOPUP.Columns(25).Caption = "Expiry"
    
    GRDPOPUP.Columns(0).Width = 1500
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 0
    GRDPOPUP.Columns(3).Width = 0
    GRDPOPUP.Columns(4).Width = 0
    GRDPOPUP.Columns(5).Width = 0
    GRDPOPUP.Columns(6).Width = 0
    GRDPOPUP.Columns(7).Width = 0
    GRDPOPUP.Columns(8).Width = 0
    GRDPOPUP.Columns(9).Width = 1000
    GRDPOPUP.Columns(10).Width = 1000
    GRDPOPUP.Columns(11).Width = 1000
    GRDPOPUP.Columns(12).Width = 0
    GRDPOPUP.Columns(13).Width = 0
    GRDPOPUP.Columns(14).Width = 0
    GRDPOPUP.Columns(15).Width = 0
    GRDPOPUP.Columns(16).Width = 0
    GRDPOPUP.Columns(17).Width = 0
    GRDPOPUP.Columns(18).Width = 0
    GRDPOPUP.Columns(19).Width = 0
    GRDPOPUP.Columns(20).Width = 0
    GRDPOPUP.Columns(21).Width = 0
    GRDPOPUP.Columns(22).Width = 0
    GRDPOPUP.Columns(23).Width = 0
    GRDPOPUP.Columns(24).Width = 0
    GRDPOPUP.Columns(25).Width = 1200
    GRDPOPUP.Columns(26).Width = 0
    GRDPOPUP.Columns(27).Width = 0
    GRDPOPUP.Columns(28).Width = 0
    GRDPOPUP.Columns(29).Width = 0
    GRDPOPUP.Columns(30).Width = 0
    GRDPOPUP.Columns(31).Width = 3000
    
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True

End Function

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
    TXTEXPIRY.BackColor = &H98F3C1
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
        
            If Len(Trim(TXTEXPIRY.text)) = 1 Then GoTo SKIP
            If Len(Trim(TXTEXPIRY.text)) < 5 Then Exit Sub
            If Val(Mid(TXTEXPIRY.text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.text, 4, 5)) = 0 Then Exit Sub
SKIP:
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
            End If
        Case vbKeyEscape
             If Len(Trim(TXTEXPIRY.text)) = 1 Then GoTo Nextstep
            If Len(Trim(TXTEXPIRY.text)) < 5 Then Exit Sub
            If Val(Mid(TXTEXPIRY.text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.text, 4, 5)) = 0 Then Exit Sub
Nextstep:
            If TxtMRP.Visible = True Then
                TxtMRP.Enabled = True
                TxtMRP.SetFocus
            Else
                TXTQTY.SetFocus
            End If
    End Select
End Sub

Private Sub TxtCessPer_GotFocus()
    TxtCessPer.SelStart = 0
    TxtCessPer.SelLength = Len(TxtCessPer.text)
    Call CHANGEBOXCOLOR(TxtCessPer, True)
End Sub

Private Sub TxtCessPer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtCessAmt.text) <> 0 Then
                TxtCessAmt.Enabled = True
                TxtCessAmt.SetFocus
            Else
                If lblsubdealer.Caption = "D" And txtretaildummy.Visible = True Then
                    txtretaildummy.Enabled = True
                    txtretaildummy.SetFocus
                Else
                    txtcommi.text = 0
                    Set GRDPRERATE.DataSource = Nothing
                    fRMEPRERATE.Visible = False
                    Call CMDADD_Click
                End If
            End If
        Case vbKeyEscape
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyDown
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
    
    TxtCessPer.Tag = 0
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        TxtCessPer.Tag = Val(txtretail.text) * Val(TxtCessPer.text) / 100
        LBLSUBTOTAL.Caption = Format(Round(Val(txtretail.text) - Val(TxtCessPer.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text) - Val(TxtCessPer.Tag), 2), ".000")
    Else
        TxtCessPer.Tag = Val(TXTQTY.text) * Val(txtretail.text) * Val(TxtCessPer.text) / 100
        LBLSUBTOTAL.Caption = Format(Round((Val(TXTQTY.text) * Val(txtretail.text)) - Val(TxtCessPer.Tag), 2), ".000")
        LblGross.Caption = Format(Round((Val(TXTQTY.text) * Val(TXTRETAILNOTAX.text)) - Val(TxtCessPer.Tag), 2), ".000")
    End If
    Call CHANGEBOXCOLOR(TxtCessPer, False)
    ''TxtCessPer.Text = Format(TxtCessPer.Text, ".000")

End Sub

Private Sub TxtCessAmt_GotFocus()
    TxtCessAmt.SelStart = 0
    TxtCessAmt.SelLength = Len(TxtCessAmt.text)
    Call CHANGEBOXCOLOR(TxtCessAmt, True)
End Sub

Private Sub TxtCessAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If lblsubdealer.Caption = "D" And txtretaildummy.Visible = True Then
                txtretaildummy.Enabled = True
                txtretaildummy.SetFocus
            Else
                txtcommi.text = 0
                Set GRDPRERATE.DataSource = Nothing
                fRMEPRERATE.Visible = False
                Call CMDADD_Click
            End If
        Case vbKeyEscape
            TxtCessPer.Enabled = True
            TxtCessPer.SetFocus
        Case vbKeyDown
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtCessAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCessAmt_LostFocus()
    
    TxtCessAmt.Tag = 0
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        TxtCessAmt.Tag = Val(txtretail.text) * Val(TxtCessAmt.text) / 100
        LBLSUBTOTAL.Caption = Format(Round(Val(txtretail.text) - Val(TxtCessAmt.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text) - Val(TxtCessAmt.Tag), 2), ".000")
    Else
        TxtCessAmt.Tag = Val(TXTQTY.text) * Val(txtretail.text) * Val(TxtCessAmt.text) / 100
        LBLSUBTOTAL.Caption = Format(Round((Val(TXTQTY.text) * Val(txtretail.text)) - Val(TxtCessAmt.Tag), 2), ".000")
        LblGross.Caption = Format(Round((Val(TXTQTY.text) * Val(TXTRETAILNOTAX.text)) - Val(TxtCessAmt.Tag), 2), ".000")
    End If
    Call CHANGEBOXCOLOR(TxtCessAmt, False)
    ''TxtCessAmt.Text = Format(TxtCessAmt.Text, ".000")

End Sub

Private Function ReportGeneratION_WO()
    
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
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
    '//chr(27) & chr(42) & chr(1) - for Enlarge letter and bold


    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(55) & Chr(27) & Chr(55)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
        Print #1, AlignLeft("ESTIMATE", 25)
'        If CHKName.value = 0 Then
'
'        Else
'            Print #1, AlignLeft("V.P. STORES", 25)
'            Print #1, AlignLeft("AREEPARAMBU, CHERTHALA", 25)
'            Print #1,
'            Print #1, "TO: " & TxtBillName.Text
'            If Trim(TxtBillAddress.Text) <> "" Then Print #1, TxtBillAddress.Text
'            If Trim(TxtPhone.Text) <> "" Then Print #1, "Phone: " & TxtPhone.Text
'        End If
        Print #1, "No. " & Trim(txtBillNo.text) & Space(2) & AlignRight("Date:" & TXTINVDATE.text, 45) '& Space(2) & LBLTIME.Caption
        'Print #1, AlignRight("Date:" & TXTINVDATE.Text, 51)
        Print #1, "TO: " & TxtBillName.text
        Print #1, "    " & Left(Trim(TxtBillAddress.text), 50)
        Print #1, "    " & Mid(Trim(TxtBillAddress.text), 51)
        Print #1, RepeatString("-", 53)
        Print #1, AlignLeft("Sl", 2) & Space(2) & _
                AlignLeft("Description", 23) & _
                AlignLeft("UOM", 5) & Space(1) & _
                AlignLeft("Qty", 8) & Space(0) & _
                AlignLeft("Rate", 5) & Space(0) & _
                AlignLeft("Amount", 7) '& _
                Chr (27) & Chr(72) '//Bold Ends
    
        Print #1, RepeatString("-", 53)
        'Print #1, RepeatString("-", 53)
    
        For i = 1 To grdsales.rows - 1
            If Trim(grdsales.TextMatrix(i, 33)) = "" Then
                grdsales.Tag = Trim(grdsales.TextMatrix(i, 2))
            Else
                grdsales.Tag = Trim(grdsales.TextMatrix(i, 33))
            End If
            Print #1, AlignLeft(Val(i), 4) & _
                Space(0) & AlignLeft(Mid(grdsales.Tag, 1, 22), 22) & _
                AlignRight(Left(grdsales.TextMatrix(i, 30), 4), 4) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 6) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 8) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 8) '& _
                Chr (27) & Chr(72) '//Bold Ends
            Print #1,
        Next i
        Print #1, RepeatString("-", 53)
        
        'Print #1, AlignRight("-------------", 47)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 38) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 11)
            Print #1, AlignRight("DISC AMOUNT ", 38) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 11)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 38) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 11)
        End If
        If Val(Txthandle.text) <> 0 Then
            Print #1, AlignRight("HANDLING CHARGE ", 38) & AlignRight((Format(Txthandle.text, "####.00")), 11)
        End If
        If Val(TxtFrieght.text) <> 0 Then
            Print #1, AlignRight("FRIEGHT CHARGE  ", 38) & AlignRight((Format(Txthandle.text, "####.00")), 11)
        End If
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 47) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, AlignRight("Round off ", 38) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 11)
        Print #1, Chr(13)
        If Val(LBLRETAMT.Caption) <> 0 Then Print #1, AlignRight("RETURN AMOUNT ", 38) & AlignRight((Format(Round(LBLRETAMT.Caption, 0), "####.00")), 11)
'        If Val(Txthandle.Text) <> 0 Then
'            Print #1, AlignRight("HANDLING CHARGE ", 38) & AlignRight((Format(Txthandle.Text, "####.00")), 11)
'        End If
        If Val(TxtFrieght.text) <> 0 Then
            Print #1, AlignRight("FRIEGHT CHARGE  ", 38) & AlignRight((Format(Txthandle.text, "####.00")), 11)
        End If
        'Print #1, AlignRight("NET AMOUNT ", 38) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 11)
        Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & AlignRight("NET AMOUNT ", 30) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 11)
        Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(18) & " "
        'Print #1, Chr(27) & Chr(30) & Chr(15) & AlignRight("NET AMOUNT ", 30) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 11)
        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(18) & AlignRight("NET AMOUNT ", 38) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 11)
        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
        Num = CCur(Round(lblnetamount.Caption, 0))
        Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & ")", 55)
        Print #1, RepeatString("-", 53)
        'Print #1, RepeatString("-", 53)
'        If OP_BAL > 0 Then
'            Print #1, AlignRight("Old Outstanding", 38) & AlignRight((Format(OP_BAL, "####.00")), 12)
'        End If
'        If RCPT_AMT > 0 Then
'            Print #1, AlignRight("Received Amt", 38) & AlignRight((Format(RCPT_AMT, "####.00")), 12)
'        End If
'        If Not (RCPT_AMT = 0 And OP_BAL = 0) Then
'            Print #1, AlignRight("Total Bal Amt", 38) & AlignRight((Format((Val(lblnetamount.Caption) + OP_BAL) - RCPT_AMT, "####.00")), 12)
'        End If
        'Print #1, Chr(27) & Chr(71) & Chr(0)
    
        'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** THANK YOU ****", 42)
    

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
    Exit Function

ERRHAND:
    MsgBox err.Description
End Function

Private Function ReportGeneratION_Composition()
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
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
    'Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        'Print #1, Chr(27) & Chr(71) & Chr(10) & AlignRight("786", 32)
        Print #1, Chr(27) & Chr(71) & Chr(10) & AlignLeft(RSTCOMPANY!COMP_NAME, 50)
        Print #1, AlignLeft(RSTCOMPANY!Address, 50)
        Print #1, AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, "Phone: " & RSTCOMPANY!TEL_NO & ", " & RSTCOMPANY!FAX_NO
        Print #1, Space(0) & "GSTIN No. " & RSTCOMPANY!DL_NO
        If MDIMAIN.LBLSTATE.Caption = "32" Then
            Print #1, Space(0) & "State: Kerala(32)"
        Else
            Print #1, Space(0) & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption)
        End If
        Print #1, RepeatString("-", 60)
        Print #1,
        
        Print #1, "Bill No. " & Trim(txtBillNo.text) & Space(2) & AlignRight("Date:" & TXTINVDATE.text, 40) '& Space(2) & LBLTIME.Caption
        Print #1, "TO: " & TxtBillName.text '& " (" & Trim(TxtCode.Text) & ")"
        If Trim(TxtBillAddress.text) <> "" Then Print #1, TxtBillAddress.text
        If Trim(TxtPhone.text) <> "" Then Print #1, "Phone: " & TxtPhone.text
        If Trim(TXTTIN.text) <> "" Then Print #1, "GSTIN: " & Trim(TXTTIN.text)
        Print #1,
        
        Print #1, Chr(27) & Chr(72) & Chr(10) & Space(19) & AlignLeft("BILL OF SUPPLY", 20)
        Print #1, Space(12) & "(To be prepared in Triplicate)"
        Print #1, Space(9) & "(Not Eligible for Input Tax Credit)"
    
        Print #1, RepeatString("-", 60)
        Print #1, AlignLeft("Sl", 2) & Space(1) & _
                AlignLeft("HSNCODE", 8) & Space(0) & _
                AlignLeft("Description", 22) & _
                AlignLeft("Qty", 4) & Space(3) & _
                AlignLeft("MRP", 4) & Space(3) & _
                AlignLeft("Rate", 7) & Space(0) & _
                AlignLeft("Amount", 9) '& _
                Chr (27) & Chr(72) '//Bold Ends
    
        Print #1, RepeatString("-", 60)
        
        Dim HSNCODE As String
        Dim RSTHSNCODE As ADODB.Recordset
        For i = 1 To grdsales.rows - 1
            Set RSTHSNCODE = New ADODB.Recordset
            RSTHSNCODE.Open "SELECT * From ITEMMASTVAN WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTHSNCODE.EOF And RSTHSNCODE.BOF) Then
                HSNCODE = IIf(IsNull(RSTHSNCODE!REMARKS), "", RSTHSNCODE!REMARKS)
            End If
            RSTHSNCODE.Close
            Set RSTHSNCODE = Nothing
            If Trim(grdsales.TextMatrix(i, 33)) = "" Then
                grdsales.Tag = Trim(grdsales.TextMatrix(i, 2))
            Else
                grdsales.Tag = Trim(grdsales.TextMatrix(i, 33))
            End If
            Print #1, AlignLeft(Val(i), 3) & _
                Space(0) & AlignLeft(HSNCODE, 8) & _
                Space(0) & AlignLeft(Mid(grdsales.Tag, 1, 18), 18) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2) & Mid(grdsales.TextMatrix(i, 30), 1, 3), 7) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 5)), 2), "0.00"), 7) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 8) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 9) '& _
                Chr (27) & Chr(72) '//Bold Ends
        Next i
    
        Print #1, AlignRight("-------------", 60)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 47) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
            Print #1, AlignRight("DISC AMOUNT ", 47) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 47) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
        End If
        If Val(Txthandle.text) <> 0 Then
            Print #1, AlignRight("HANDLING CHARGE ", 47) & AlignRight((Format(Txthandle.text, "####.00")), 12)
        End If
        If Val(TxtFrieght.text) <> 0 Then
            Print #1, AlignRight("FRIEGHT CHARGE  ", 47) & AlignRight((Format(Txthandle.text, "####.00")), 12)
        End If
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 53) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, AlignRight("Round off ", 47) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
        Print #1, Chr(13)
        Print #1, AlignRight("NET AMOUNT ", 47) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 12)
        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
        Num = CCur(Round(lblnetamount.Caption, 0))
        Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & " only)", 60)
        Print #1, RepeatString("-", 60)
        Print #1, "****Thank You... Visit Again****"
        Print #1, "E.&.O.E" '          SUBJECT TO ALAPPUZHA JURISDICTION"
        'Print #1, Chr(13)
        'Print #1, Chr(13)
        'Print #1, Space(34) & "For ECO SHOP"
        'Print #1, Chr(27) & Chr(71) & Chr(0)
    
        'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** THANK YOU ****", 32)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing

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
    Exit Function

ERRHAND:
    MsgBox err.Description
End Function

Private Sub TxtSpec_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape, vbKeyReturn
            If cmdadd.Enabled = True Then cmdadd.SetFocus
            'If CMDPRINT.Enabled = True Then CMDPRINT.SetFocus
            If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
            If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
            If txtBillNo.Visible = True Then txtBillNo.SetFocus
            'If TxtName1.Enabled = True Then TxtName1.SetFocus
            'If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Visible = True Then
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            Else
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            'If txtretail.Enabled = True Then txtretail.SetFocus
            'If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
            'If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            'If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            'If txtcommi.Enabled = True Then txtcommi.SetFocus
    End Select
End Sub

Private Sub TxtSpec_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPrintname_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub CMDPre_Click()
    Dim currbillno As Long
    currbillno = Val(txtBillNo.text)
    On Error GoTo ERRHAND
    preclick = True
    If CMDEXIT.Enabled = False Then cmdRefresh_Click
    preclick = False
    If currbillno <= 1 Then
        txtBillNo.text = 1
    Else
        txtBillNo.text = currbillno - 1
    End If
    LBLBILLNO.Caption = Val(txtBillNo.text)
            
    TXTDEALER.text = "CASH"
    DataList2.BoundText = "130000"
    CMBDISTI.text = ""
    LBLRETAMT.Caption = ""
    TXTTIN.text = ""
    TxtCN.text = ""
    TXTTOTALDISC.text = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
    lblOB.Caption = ""
    Chkcancel.Value = 0
    
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    lbltime.Caption = Format(Time, "hh:mm AMPM")
    OLD_BILL = False
    
    Call txtBillNo_KeyDown(13, 0)
    
    Dim m_Rcpt_Amt As Double
    Dim Rcptamt As Double
    Dim m_OP_Bal As Double
    Dim m_Bal_Amt As Double
    
    m_Rcpt_Amt = 0
    m_OP_Bal = 0
    m_Bal_Amt = 0
    Rcptamt = 0
    lblOB.Caption = ""
    If DataList2.BoundText <> "" And DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
        Dim RSTTRXFILE As ADODB.Recordset
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select OPEN_DB from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            m_OP_Bal = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
               
        Set RSTTRXFILE = New ADODB.Recordset
        'RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' AND NOT(TRX_TYPE= 'RT' AND INV_TRX_TYPE ='WO' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & ") ", db, adOpenStatic, adLockReadOnly, adCmdText
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE <= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE= 'RT' AND INV_TRX_TYPE ='GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        If GRDRECEIPT.rows > 1 Then Rcptamt = GRDRECEIPT.TextMatrix(0, 0)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND NOT(INV_TRX_TYPE = 'GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & ") AND (TRX_TYPE = 'DR' OR TRX_TYPE = 'DB') AND INV_DATE <= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                Case Else
                    m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        lblOB.Caption = Format(Round((m_OP_Bal + m_Bal_Amt) - (m_Rcpt_Amt), 2), "0.00")
    Else
        lblOB.Caption = "0.00"
    End If
    
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdnxt_Click()
        
    Dim currbillno As Long
    currbillno = Val(txtBillNo.text)
    preclick = True
    If CMDEXIT.Enabled = False Then cmdRefresh_Click
    preclick = False
    Dim rstBILL As ADODB.Recordset
    Dim lastbillno As Long
    
    On Error GoTo ERRHAND
    
    TXTDEALER.text = "CASH"
    DataList2.BoundText = "130000"
    CMBDISTI.text = ""
    lblOB.Caption = ""
    txtphsearch.text = ""
    txtcrdays.text = ""
    'lblcrlimit.Caption = ""
    LBLRETAMT.Caption = ""
    TXTTIN.text = ""
    'TxtUID.Text = ""
    TxtCN.text = ""
    TXTTOTALDISC.text = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
    lblOB.Caption = ""
    'lblCBFLAG.Caption = ""
    Chkcancel.Value = 0
    OLD_BILL = False
    
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    lbltime.Caption = Format(Time, "hh:mm AMPM")
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'GI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        lastbillno = IIf(IsNull(rstBILL.Fields(0)), 0, rstBILL.Fields(0))
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    If currbillno > lastbillno Then Exit Sub
    txtBillNo.text = currbillno + 1
    LBLBILLNO.Caption = Val(txtBillNo.text)
    
    
    Call txtBillNo_KeyDown(13, 0)
    
    Dim m_Rcpt_Amt As Double
    Dim Rcptamt As Double
    Dim m_OP_Bal As Double
    Dim m_Bal_Amt As Double
    
    m_Rcpt_Amt = 0
    m_OP_Bal = 0
    m_Bal_Amt = 0
    Rcptamt = 0
    lblOB.Caption = ""
    If DataList2.BoundText <> "" And DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
        Dim RSTTRXFILE As ADODB.Recordset
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select OPEN_DB from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            m_OP_Bal = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
               
        Set RSTTRXFILE = New ADODB.Recordset
        'RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' AND NOT(TRX_TYPE= 'RT' AND INV_TRX_TYPE ='WO' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & ") ", db, adOpenStatic, adLockReadOnly, adCmdText
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE <= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    m_Rcpt_Amt = m_Rcpt_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE= 'RT' AND INV_TRX_TYPE ='GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    m_Rcpt_Amt = m_Rcpt_Amt - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        If GRDRECEIPT.rows > 1 Then Rcptamt = GRDRECEIPT.TextMatrix(0, 0)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND NOT(INV_TRX_TYPE = 'GI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & ") AND (TRX_TYPE = 'DR' OR TRX_TYPE = 'DB') AND INV_DATE <= '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                Case Else
                    m_Bal_Amt = m_Bal_Amt + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        lblOB.Caption = Format(Round((m_OP_Bal + m_Bal_Amt) - (m_Rcpt_Amt), 2), "0.00")
    Else
        lblOB.Caption = "0.00"
    End If
    
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub txtNetrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape, vbKeyTab
            If txtretail.Enabled = True Then txtretail.SetFocus
    End Select
End Sub

Private Sub txtNetrate_LostFocus()
    Call CHANGEBOXCOLOR(txtNetrate, False)
    If Val(txtNetrate.text) = 0 Then Exit Sub
    If Val(TxtCessPer.text) <> 0 Or Val(TxtCessAmt.text) <> 0 Then
        TXTRETAILNOTAX.text = (Val(txtNetrate.text) - Val(TxtCessAmt.text)) / (1 + (Val(TXTTAX.text) / 100) + (Val(TxtCessPer.text) / 100))
        txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 3)
        TXTRETAILNOTAX.text = Val(txtretail.text)
    End If
    Call CHANGEBOXCOLOR(txtNetrate, False)
End Sub

Private Sub txtretail_Change()
    Dim discamnt As Double
    discamnt = 0
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        txtNetrate.text = Format(Val(txtretail.text), "0.00")
    Else
        discamnt = Val(txtretail.text) * Val(TXTDISC.text) / 100
        txtNetrate.text = Format(Round(Val(txtretail.text) - discamnt, 2), ".00")
        
        txtNetrate.text = Val(txtNetrate.text) + (Val(TXTRETAILNOTAX.text) - (Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100)) * Val(TxtCessPer) / 100
        txtNetrate.text = Val(txtNetrate.text) + Round(Val(TxtCessAmt.text), 3)
        txtNetrate.text = Format(Val(txtNetrate.text), "0.00")
    End If
End Sub

Private Sub lblcomamt_GotFocus()
    lblcomamt.SelStart = 0
    lblcomamt.SelLength = Len(lblcomamt.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    Call CHANGEBOXCOLOR(lblcomamt, True)
End Sub

Private Sub lblcomamt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If MsgBox("This will reset the individual commission. Are you sure?", vbYesNo + vbDefaultButton2, "Ez Biz") = vbNo Then Exit Sub
            txtcommper.text = Format(Round((Val(lblcomamt.text) * 100) / Val(lblnetamount.Caption), 2), "0.00")
            Dim i As Integer
            Dim sum_ary As Double
            Dim GROSSAMT As Double
            
            On Error GoTo ERRHAND
            sum_ary = 0
            GROSSAMT = 0
            For i = 1 To grdsales.rows - 1
                'If Aray(i) < sn Then sn = Aray(i)
                If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                    sum_ary = sum_ary + Val(grdsales.TextMatrix(i, 3))
                Else
                    sum_ary = sum_ary + Val(grdsales.TextMatrix(i, 3)) '* Val(grdsales.TextMatrix(i, 24))
                End If
            Next i
            
            lblcomamt.Tag = ""
            For i = 1 To grdsales.rows - 1
                grdsales.TextMatrix(i, 24) = Round((Val(lblcomamt.text) / sum_ary), 3)
                db.Execute "Update TRXFILEVAN set COM_AMT = " & Val(grdsales.TextMatrix(i, 24)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & " "
'                grossamt = Round((Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 14))) * (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 5))), 3)
'                grdsales.TextMatrix(i, 8) = Format(Round(((grossamt / (Val(grdsales.TextMatrix(i, 5)) * (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))))) + ((Val(grdsales.TextMatrix(i, 32)) / Val(grdsales.TextMatrix(i, 5))))), 4), ".0000")
'                LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(i, 32)), ".00")
'                db.Execute "Update RTRXFILEVAN set ITEM_COST = " & Val(grdsales.TextMatrix(i, 8)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
'                db.Execute "Update RTRXFILEVAN set EXPENSE = " & Val(grdsales.TextMatrix(i, 32)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
               
            Next i
            
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            CMDEXIT.Enabled = False
            cmdRefresh.Enabled = True
            CMDPre.Caption = "Save && Pre&vious"
            cmdnxt.Caption = "Save && &Next"
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub lblcomamt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcommper_GotFocus()
    txtcommper.SelStart = 0
    txtcommper.SelLength = Len(txtcommper.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    Call CHANGEBOXCOLOR(txtcommper, True)
End Sub

Private Sub txtcommper_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If MsgBox("This will reset the individual commission. Are you sure?", vbYesNo + vbDefaultButton2, "Ez Biz") = vbNo Then Exit Sub
            lblcomamt.text = Format(Round(Val(lblnetamount.Caption) * Val(txtcommper.text) / 100, 2), "0.00")
            Dim i As Integer
            Dim sum_ary As Double
            Dim GROSSAMT As Double
            
            On Error GoTo ERRHAND
            sum_ary = 0
            GROSSAMT = 0
            For i = 1 To grdsales.rows - 1
                'If Aray(i) < sn Then sn = Aray(i)
                If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                    sum_ary = sum_ary + Val(grdsales.TextMatrix(i, 3))
                Else
                    sum_ary = sum_ary + Val(grdsales.TextMatrix(i, 3)) '* Val(grdsales.TextMatrix(i, 24))
                End If
            Next i
            
            lblcomamt.Tag = ""
            For i = 1 To grdsales.rows - 1
                grdsales.TextMatrix(i, 24) = Round((Val(lblcomamt.text) / sum_ary), 3)
                db.Execute "Update TRXFILEVAN set COM_AMT = " & Val(grdsales.TextMatrix(i, 24)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='GI' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & " "
'                grossamt = Round((Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 14))) * (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 5))), 3)
'                grdsales.TextMatrix(i, 8) = Format(Round(((grossamt / (Val(grdsales.TextMatrix(i, 5)) * (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))))) + ((Val(grdsales.TextMatrix(i, 32)) / Val(grdsales.TextMatrix(i, 5))))), 4), ".0000")
'                LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(i, 32)), ".00")
'                db.Execute "Update RTRXFILEVAN set ITEM_COST = " & Val(grdsales.TextMatrix(i, 8)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
'                db.Execute "Update RTRXFILEVAN set EXPENSE = " & Val(grdsales.TextMatrix(i, 32)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
               
            Next i
            
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            CMDEXIT.Enabled = False
            cmdRefresh.Enabled = True
            CMDPre.Caption = "Save && Pre&vious"
            cmdnxt.Caption = "Save && &Next"
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub txtcommper_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtphsearch_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_REC.State <> 1 Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME, TELNO, FAXNO from CUSTMAST WHERE TELNO Like '" & Me.txtphsearch.text & "%' OR FAXNO Like '" & Me.txtphsearch.text & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME, TELNO, FAXNO from CUSTMAST WHERE TELNO Like '" & Me.txtphsearch.text & "%' OR FAXNO Like '" & Me.txtphsearch.text & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
    CHANGE_ADDRESS = True
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub txtphsearch_GotFocus()
    txtphsearch.SelStart = 0
    txtphsearch.SelLength = Len(txtphsearch.text)
    CMDDELIVERY.Enabled = False
    CMDSALERETURN.Enabled = False
    
    fRMEPRERATE.Visible = False
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    Call CHANGEBOXCOLOR(txtphsearch, True)
End Sub

Private Sub txtphsearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList2.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub txtphsearch_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
''        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
''            KeyAscii = Asc(UCase(Chr(KeyAscii)))
''        Case Else
''            KeyAscii = 0
    End Select
End Sub

Private Sub CHANGEBOXCOLOR(BOX As TextBox, texton As Boolean)
    If texton Then
        BOX.BackColor = &H98F3C1
    Else
        BOX.BackColor = vbWhite
    End If
End Sub

Private Sub TxtWarranty_type_LostFocus()
    Call CHANGEBOXCOLOR(TxtWarranty_type, False)
End Sub

Private Sub CmdExport_Click()
    
    If grdsales.rows <= 1 Then Exit Sub
    
    On Error GoTo ERRHAND
    Dim i As Integer
    Dim allunbill As Boolean
    allunbill = True
    
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 46) = "N" Then
            allunbill = False
            Exit For
        End If
    Next i
    
    If allunbill = True Then
        MsgBox "All are unbilled Items", , "Export Bill"
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to export the bill?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
    
    On Error GoTo ERRHAND
    Dim Strconnct As String
    Dim db2 As New ADODB.Connection
    Strconnct = "Driver={MySQL ODBC 5.1 Driver};Server=" & DBPath & ";Port=3306;Database=" & dbase2 & ";User=root; Password=###%%database%%###ret; Option=2;"
    db2.Open Strconnct
    db2.CursorLocation = adUseClient
    
    Dim BillType As String
    Dim rstBILL As ADODB.Recordset
    Dim BILL_NUM As Long
    BillType = "GI"
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = '" & BillType & "'", db2, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        BILL_NUM = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    If MsgBox("Are you sure you want to export the bill as B2C Bill No. " & BILL_NUM, vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then
        db2.Close
        Set db2 = Nothing
        Exit Sub
    End If
    
    db2.Execute "delete From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & ""
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim PR_CODE, PR_NAME As String
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 13) = "" Then GoTo SKIP_2
        If grdsales.TextMatrix(i, 46) = "Y" Then GoTo SKIP_2
        
        Set rstTRXMAST = New ADODB.Recordset
        rstTRXMAST.Open "Select * From ITEMMASTVAN WHERE ITEM_NAME = '" & Trim(grdsales.TextMatrix(i, 2)) & "' ", db2, adOpenStatic, adLockReadOnly, adCmdText
        If (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMASTVAN ", db2, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                If IsNull(RSTITEMMAST.Fields(0)) Then
                    PR_CODE = 1
                Else
                    PR_CODE = Val(RSTITEMMAST.Fields(0)) + 1
                End If
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * From ITEMMASTVAN", db2, adOpenStatic, adLockOptimistic, adCmdText
            RSTITEMMAST.AddNew
            RSTITEMMAST!ITEM_CODE = PR_CODE
            RSTITEMMAST!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
            PR_NAME = Trim(grdsales.TextMatrix(i, 2))
            RSTITEMMAST!Category = "GENERAL"
            RSTITEMMAST!UNIT = 1
            RSTITEMMAST!MANUFACTURER = "GENERAL"
            RSTITEMMAST!DEAD_STOCK = "N"
            RSTITEMMAST!REMARKS = ""
            RSTITEMMAST!REORDER_QTY = 1
            RSTITEMMAST!PACK_TYPE = "Nos"
            RSTITEMMAST!FULL_PACK = "Nos"
            RSTITEMMAST!BIN_LOCATION = ""
            RSTITEMMAST!ITEM_COST = 0
            RSTITEMMAST!MRP = 0
            RSTITEMMAST!SALES_TAX = 0
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
            RSTITEMMAST!P_RETAIL = 0
            RSTITEMMAST!P_WS = 0
            RSTITEMMAST!CRTN_PACK = 1
            RSTITEMMAST!P_CRTN = 0
            RSTITEMMAST!LOOSE_PACK = 1
            RSTITEMMAST!UN_BILL = "N"
            RSTITEMMAST.Update
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        Else
            PR_CODE = rstTRXMAST!ITEM_CODE
            PR_NAME = rstTRXMAST!ITEM_NAME
        End If
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  From ITEMMASTVAN WHERE ITEM_CODE = '" & PR_CODE & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").Value = adCriteriaKey
                '!ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(I, 3)) + Val(grdsales.TextMatrix(I, 20))
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 27))), 3)
                !FREE_QTY = !FREE_QTY + Round((Val(grdsales.TextMatrix(i, 20)) * Val(grdsales.TextMatrix(i, 27))), 3)
                !CLOSE_QTY = !CLOSE_QTY - Round(((Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 20))) * Val(grdsales.TextMatrix(i, 27))), 3)
    
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(i, 12))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(i, 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From TRXFILEVAN", db2, adOpenStatic, adLockOptimistic, adCmdText
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = BillType
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = BILL_NUM
        RSTTRXFILE!VCH_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = Trim(grdsales.TextMatrix(i, 25))
        RSTTRXFILE!ITEM_CODE = PR_CODE
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(i, 11))
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 5))
        
'        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
'        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
'        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
'        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
'
        
        
        
        'TXTRETAILNOTAX.Text = Round(Val(txtretail.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                        
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6)) 'Round(Val(grdsales.TextMatrix(i, 7)) * 100 / ((Val(TXTAREA.Tag) + Val(grdsales.TextMatrix(i, 9))) + 100), 4)
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 6)) 'Round(Val(grdsales.TextMatrix(i, 7)) * 100 / ((Val(TXTAREA.Tag) + Val(grdsales.TextMatrix(i, 9))) + 100), 4)
                 
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7)) 'Round(RSTTRXFILE!PTR + (RSTTRXFILE!PTR * Val(grdsales.TextMatrix(i, 9)) / 100), 3)
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 7)) 'Round(RSTTRXFILE!PTR + (RSTTRXFILE!PTR * Val(grdsales.TextMatrix(i, 9)) / 100), 3)
        RSTTRXFILE!kfc_tax = 0
        RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(i, 24))
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
            RSTTRXFILE!COM_FLAG = "Y"
        Else
            RSTTRXFILE!COM_FLAG = "N"
        End If
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(i, 26)
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 27))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        If Trim(TxtBillName.text) = "" Then
            RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.text), 1, 30)
        Else
            RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(TxtBillName.text), 1, 30)
        End If
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(i, 17))
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
        
        TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(i, 6)) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 8)) / 100)
        TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(i, 9)) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(i, 40)) / 100) + Val(grdsales.TextMatrix(i, 41))
        If UCase(grdsales.TextMatrix(i, 25)) = "SERVICE CHARGE" Then
            'TXTDISC.Tag = Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
            RSTTRXFILE!TRX_TOTAL = Round(Val(TXTRETAILNOTAX.Tag), 4)
        Else
            'TXTDISC.Tag = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
            RSTTRXFILE!TRX_TOTAL = Round(Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(i, 3)), 4)
        End If
                    
        'RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(i, 28) = "", Null, grdsales.TextMatrix(i, 28))
        RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(i, 29)
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 30)
        RSTTRXFILE!PUR_TAX = Val(grdsales.TextMatrix(i, 45))
        
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db2, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!Area = RSTITEMMAST!Area
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
SKIP_2:
    
    Next i
    
    db2.Execute "delete From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & ""
    db2.Execute "delete From TRXSUBVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & ""
    'db2.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='DR' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TRX_TYPE = '" & BillType & "'"
    'db2.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_VCH_NO = " & Val(LBLBILLNO.Caption) & " AND B_TRX_TYPE = '" & BillType & "' "
    'db2.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = '" & BillType & "'"
    
    'db2.Execute "delete From P_Rate WHERE TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & ""
    
    i = 0
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        RSTITEMMAST!Area = Trim(TXTAREA.Text)
'        RSTITEMMAST!KGST = Trim(TXTTIN.Text)
'        RSTITEMMAST!ADDRESS = Trim(TxtBillAddress.Text)
'        RSTITEMMAST.Update
'    End If
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXMASTVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & "", db2, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = BILL_NUM
        RSTTRXFILE!TRX_TYPE = BillType
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!gross_amt = Val(LBLGROSSAMT.Caption)
        RSTTRXFILE!VCH_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.text
        RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.text)
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
    End If
    
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db2, adOpenStatic, adLockReadOnly
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        RSTTRXFILE!Area = RSTITEMMAST!Area
'    End If
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    
    RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
    RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    RSTTRXFILE!gross_amt = Val(LBLGROSSAMT.Caption)
    RSTTRXFILE!VCH_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.text
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
    RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
    RSTTRXFILE!REF_NO = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "A"
        RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.text)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "P"
        RSTTRXFILE!DISCOUNT = Round(RSTTRXFILE!VCH_AMOUNT * Val(TXTTOTALDISC.text) / 100, 2)
    End If
    RSTTRXFILE!check_flag = "I"
    RSTTRXFILE!POST_FLAG = "Y"
    RSTTRXFILE!CFORM_NO = Time
    RSTTRXFILE!REMARKS = DataList2.text
    RSTTRXFILE!DISC_PERS = 0
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!VEHICLE = ""
    RSTTRXFILE!PHONE = Trim(TxtPhone.text)
    RSTTRXFILE!TIN = ""
    RSTTRXFILE!FRIEGHT = Val(TxtFrieght.text)
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!cr_days = 0
    RSTTRXFILE!BILL_NAME = Trim(TxtBillName.text)
    RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.text)
    RSTTRXFILE!cr_days = Val(txtcrdays.text)
    txtcommi.Tag = ""
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        RSTTRXFILE!AGENT_CODE = CMBDISTI.BoundText
        RSTTRXFILE!AGENT_NAME = CMBDISTI.text
        For i = 1 To grdsales.rows - 1
            txtcommi.Tag = Val(txtcommi.Tag) + Val(grdsales.TextMatrix(i, 24))
        Next i
        RSTTRXFILE!COMM_AMT = Val(txtcommi.Tag)
    Else
        RSTTRXFILE!AGENT_CODE = ""
        RSTTRXFILE!AGENT_NAME = ""
    End If
   
    Select Case cmbtype.ListIndex
        Case 3
            RSTTRXFILE!BILL_TYPE = "M"
        Case 1
            RSTTRXFILE!BILL_TYPE = "W"
        Case 2
            RSTTRXFILE!BILL_TYPE = "V"
        Case Else
            RSTTRXFILE!BILL_TYPE = "R"
    End Select
    RSTTRXFILE!CN_REF = Null
    
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXSUBVAN ", db2, adOpenStatic, adLockOptimistic, adCmdText
    
    'grdsales.TextMatrix(I, 15) = Trim(TXTTRXTYPE.Text)
    
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 13) = "" Then GoTo SKIP_3
        If grdsales.TextMatrix(i, 46) = "Y" Then GoTo SKIP_3
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = BILL_NUM
        RSTTRXFILE!TRX_TYPE = BillType
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!R_VCH_NO = 0
        RSTTRXFILE!R_LINE_NO = 0
        RSTTRXFILE!R_TRX_TYPE = "MI"
        RSTTRXFILE!R_TRX_YEAR = ""
        RSTTRXFILE!QTY = grdsales.TextMatrix(i, 3)
        RSTTRXFILE.Update
SKIP_3:
    Next i
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim TOTAL_AMT As Double
    TOTAL_AMT = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT SUM(TRX_TOTAL) From TRXFILEVAN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & "", db2, adOpenForwardOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        TOTAL_AMT = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        db2.Execute "Update TRXMASTVAN SET NET_AMOUNT = " & TOTAL_AMT & ", VCH_AMOUNT = " & TOTAL_AMT & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & " "
        db2.Execute "Update DBTPYMT set INV_AMT = " & TOTAL_AMT & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & BILL_NUM & " AND INV_TRX_TYPE = '" & BillType & "'"
        db2.Execute "Update BANK_TRX set TRX_AMOUNT = " & TOTAL_AMT & " WHERE B_TRX_YEAR= '" & Year(MDIMAIN.DTFROM.Value) & "' AND B_TRX_TYPE = '" & BillType & "' AND B_VCH_NO = " & BILL_NUM & ""
        db2.Execute "Update CASHATRXFILE set AMOUNT = " & TOTAL_AMT & " WHERE TRX_YEAR= '" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & BILL_NUM & " AND INV_TRX_TYPE = '" & BillType & "'"
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
            
    MsgBox "Export done successfully", , "EzBiz"
    db2.Close
    Set db2 = Nothing
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Function cost_calculate()
    If Val(LBLITEMCOST.Caption) <> 0 Then
        TXTRETAILNOTAX.Tag = Val(TXTRETAILNOTAX.text) - (Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100)
        LblProfitPerc.Caption = Round(((Val(TXTRETAILNOTAX.Tag) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption), 2)
        LblProfitPerc.Caption = Format(Val(LblProfitPerc.Caption), "0.00")
        
        LBLNETPROFIT.Caption = (Val(TXTRETAILNOTAX.Tag) - Val(LBLITEMCOST.Caption)) * Val(TXTQTY.text) '        LBLNETPROFIT.Caption = Round((((Val(TXTRETAILNOTAX.Tag) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption)) * Val(TXTQTY.Text), 2)
        LBLNETPROFIT.Caption = Format(Val(LBLNETPROFIT.Caption), "0.00")
        
    End If
    
    LblProfitAmt.Caption = Round(Val(LBLNETPROFIT.Caption), 2)  'Round((Val(TXTRETAILNOTAX.Text) - Val(LBLITEMCOST.Caption)) * Val(TXTQTY.Text), 2)
    LblProfitAmt.Caption = Format(Val(LblProfitAmt.Caption), "0.00")
End Function


Private Function COSTCALCULATE()
    
    LBLNETCOST.Caption = Round(((Val(grdsales.TextMatrix(grdsales.Row, 11)) + (Val(grdsales.TextMatrix(grdsales.Row, 11)) * Val(grdsales.TextMatrix(grdsales.Row, 45)) / 100))), 2)
    LBLITEMCOST.Caption = Val(grdsales.TextMatrix(grdsales.Row, 11))
'    If Val(LBLITEMCOST.Caption) = 0 Then
'        LblProfitPerc.Caption = 0
'    Else
'        If Val(grdsales.TextMatrix(grdsales.Row, 3)) = 0 Then
'            LblProfitPerc.Caption = 0
'        Else
'            LblProfitPerc.Caption = Round((((Val(grdsales.TextMatrix(grdsales.Row, 12)) / Val(grdsales.TextMatrix(grdsales.Row, 3))) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption), 2)
'        End If
'    End If
'    If Val(grdsales.TextMatrix(grdsales.Row, 3)) = 0 Then
'        LBLNETPROFIT.Caption = 0
'    Else
'        LBLNETPROFIT.Caption = Round(((Val(grdsales.TextMatrix(grdsales.Row, 12)) / Val(grdsales.TextMatrix(grdsales.Row, 3))) - Val(LBLITEMCOST.Caption)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 2)
'    End If
'    LBLNETCOST.Caption = Format(Val(LBLNETCOST.Caption), "0.00")
'    LblProfitPerc.Caption = Format(Val(LblProfitPerc.Caption), "0.00")
'    LBLNETPROFIT.Caption = Format(Val(LBLNETPROFIT.Caption), "0.00")
                
    If Val(LBLITEMCOST.Caption) <> 0 Then
        TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(grdsales.Row, 6)) - (Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100)
        LblProfitPerc.Caption = Round(((Val(TXTRETAILNOTAX.Tag) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption), 2)
        LblProfitPerc.Caption = Format(Val(LblProfitPerc.Caption), "0.00")
        LBLNETPROFIT.Caption = (Val(TXTRETAILNOTAX.Tag) - Val(LBLITEMCOST.Caption)) * Val(grdsales.TextMatrix(grdsales.Row, 3))
        'LBLNETPROFIT.Caption = Round((((Val(TXTRETAILNOTAX.Tag) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 2)
        LBLNETPROFIT.Caption = Format(Val(LBLNETPROFIT.Caption), "0.00")
    Else
        LblProfitPerc.Caption = 0
        LBLNETPROFIT.Caption = 0
    End If
    LblProfitAmt.Caption = Round(Val(LBLNETPROFIT.Caption), 2)  'Round((Val(TXTRETAILNOTAX.Text) - Val(LBLITEMCOST.Caption)) * Val(TXTQTY.Text), 2)
    LblProfitAmt.Caption = Format(Val(LblProfitAmt.Caption), "0.00")
    
    Exit Function
    
ERRHAND:
    MsgBox err.Description
End Function

Private Function COSTCALCULATION()
    Dim COST As Double
    Dim n As Integer
    'Dim RSTITEMMAST As ADODB.Recordset
    
    LBLTOTALCOST.Caption = ""
    LBLPROFIT.Caption = ""
    COST = 0
    On Error GoTo ERRHAND
    For n = 1 To grdsales.rows - 1
        'COST = COST + (Val(grdsales.TextMatrix(N, 11)) * Val(grdsales.TextMatrix(N, 3)))
        'COST = COST + ((Val(grdsales.TextMatrix(N, 11)) + (Val(grdsales.TextMatrix(N, 11)) * Val(grdsales.TextMatrix(N, 45)) / 100)) * (Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20))))
        
        LBLTOTALCOST.Caption = Val(LBLTOTALCOST.Caption) + (Val(grdsales.TextMatrix(n, 11)) * (Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20))))
        
        COST = (Val(grdsales.TextMatrix(n, 11))) '* (Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20))))
        'If COST <> 0 Then
        TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(n, 6)) - (Val(grdsales.TextMatrix(n, 6)) * Val(grdsales.TextMatrix(n, 8)) / 100)
        'LblProfitPerc.Caption = Round(((Val(TXTRETAILNOTAX.Tag) - Val(LBLITEMCOST.Caption)) * 100) / COST, 2)
        'LblProfitPerc.Caption = Format(Val(LblProfitPerc.Caption), "0.00")
        'LBLPROFIT.Caption = Val(LBLPROFIT.Caption) + (((Val(TXTRETAILNOTAX.Tag) - COST) * 100) / COST) * (Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20)))
        LBLPROFIT.Caption = Val(LBLPROFIT.Caption) + (Val(TXTRETAILNOTAX.Tag) - COST) * (Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20)))
'    Else
'        'LblProfitPerc.Caption = 0
'        LBLPROFIT.Caption = 0
'    End If
    Next n
    
    LBLTOTALCOST.Caption = Round(Val(LBLTOTALCOST.Caption), 2)
    LBLPROFIT.Caption = Round(Val(LBLPROFIT.Caption) - (Val(LBLDISCAMT.Caption) + Val(lblcomamt.text)), 2)
    LBLTOTALCOST.Caption = Format(Val(LBLTOTALCOST.Caption), "0.00")
    LBLPROFIT.Caption = Format(Val(LBLPROFIT.Caption), "0.00")
    
    'LBLPROFIT.Caption = Round(Val(LBLTOTAL.Caption) - (Val(LBLDISCAMT.Caption) + Val(lblcomamt.Text)) - COST, 2)
    
    
'    LBLNETCOST.Caption = Round(((Val(grdsales.TextMatrix(grdsales.Row, 11)) + (Val(grdsales.TextMatrix(grdsales.Row, 11)) * Val(grdsales.TextMatrix(grdsales.Row, 45)) / 100))), 2)
'    LBLITEMCOST.Caption = Val(grdsales.TextMatrix(grdsales.Row, 11))
'    If Val(LBLITEMCOST.Caption) <> 0 Then
'        TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(grdsales.Row, 6)) - (Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100)
'        LblProfitPerc.Caption = Round(((Val(TXTRETAILNOTAX.Tag) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption), 2)
'        LblProfitPerc.Caption = Format(Val(LblProfitPerc.Caption), "0.00")
'        LBLNETPROFIT.Caption = Round((((Val(TXTRETAILNOTAX.Tag) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 2)
'        LBLNETPROFIT.Caption = Format(Val(LBLNETPROFIT.Caption), "0.00")
'    Else
'        LblProfitPerc.Caption = 0
'        LBLNETPROFIT.Caption = 0
'    End If
'    LblProfitAmt.Caption = Round(Val(LBLNETPROFIT.Caption), 2)  'Round((Val(TXTRETAILNOTAX.Text) - Val(LBLITEMCOST.Caption)) * Val(TXTQTY.Text), 2)
'    LblProfitAmt.Caption = Format(Val(LblProfitAmt.Caption), "0.00")
    
    Exit Function
    
ERRHAND:
    MsgBox err.Description
End Function





