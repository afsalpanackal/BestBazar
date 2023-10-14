VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMservice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GST SERVICE BILLS"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18495
   Icon            =   "frmservice.frx":0000
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
      Left            =   3735
      TabIndex        =   84
      Top             =   3900
      Visible         =   0   'False
      Width           =   9555
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   2655
         Left            =   15
         TabIndex        =   85
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
         Left            =   3630
         TabIndex        =   87
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
         TabIndex        =   86
         Top             =   15
         Width           =   3615
      End
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   1860
      TabIndex        =   55
      Top             =   3765
      Visible         =   0   'False
      Width           =   8340
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2835
         Left            =   30
         TabIndex        =   58
         Top             =   360
         Width           =   8250
         _ExtentX        =   14552
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
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   30
         Visible         =   0   'False
         Width           =   5205
      End
   End
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3270
      Left            =   1875
      TabIndex        =   59
      Top             =   3690
      Visible         =   0   'False
      Width           =   10965
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   3165
         Left            =   45
         TabIndex        =   60
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
      Height          =   4140
      Left            =   90
      TabIndex        =   92
      Top             =   2760
      Visible         =   0   'False
      Width           =   13200
      _ExtentX        =   23283
      _ExtentY        =   7303
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
      Left            =   975
      TabIndex        =   0
      Top             =   30
      Width           =   885
   End
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   11130
      Left            =   -150
      TabIndex        =   48
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
         TabIndex        =   61
         Top             =   10635
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00F5E7EC&
         ForeColor       =   &H008080FF&
         Height          =   2415
         Left            =   210
         TabIndex        =   49
         Top             =   -90
         Width           =   18435
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
            ForeColor       =   &H000000FF&
            Height          =   585
            Left            =   10650
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   1110
            Width           =   6750
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   6780
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
            ForeColor       =   &H00004080&
            Height          =   315
            ItemData        =   "frmservice.frx":030A
            Left            =   7455
            List            =   "frmservice.frx":0317
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1845
            Width           =   2010
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   10215
            MaxLength       =   35
            TabIndex        =   15
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   10650
            MaxLength       =   35
            TabIndex        =   14
            Top             =   1710
            Width           =   2490
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00F5E7EC&
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
            Height          =   1005
            Left            =   9480
            TabIndex        =   88
            Top             =   90
            Width           =   3705
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
               Height          =   330
               Left            =   915
               MaxLength       =   35
               TabIndex        =   12
               Top             =   630
               Width           =   2745
            End
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
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   45
               TabIndex        =   11
               Top             =   180
               Width           =   3615
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
               Left            =   75
               TabIndex        =   89
               Top             =   660
               Width           =   870
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00F5E7EC&
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
            Left            =   5610
            TabIndex        =   73
            Top             =   90
            Width           =   3840
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
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   90
               MaxLength       =   100
               TabIndex        =   7
               Top             =   225
               Width           =   3645
            End
            Begin MSForms.TextBox TxtBillAddress 
               Height          =   1095
               Left            =   90
               TabIndex        =   8
               Top             =   570
               Width           =   3645
               VariousPropertyBits=   -1400879077
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
            Left            =   4800
            TabIndex        =   2
            Top             =   150
            Width           =   780
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
            Width           =   3990
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   300
            Left            =   2445
            TabIndex        =   1
            Top             =   150
            Width           =   1395
            _ExtentX        =   2461
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
            Width           =   3990
            _ExtentX        =   7038
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
            Left            =   13230
            TabIndex        =   16
            Top             =   420
            Visible         =   0   'False
            Width           =   3240
            _ExtentX        =   5715
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
            TabIndex        =   165
            Top             =   2040
            Width           =   3990
            _ExtentX        =   7038
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            TabIndex        =   210
            Top             =   1875
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblIGST 
            BackColor       =   &H00F5E7EC&
            Height          =   285
            Left            =   5715
            TabIndex        =   209
            Top             =   2100
            Width           =   690
         End
         Begin VB.Label lblsubdealer 
            Height          =   405
            Left            =   60
            TabIndex        =   204
            Top             =   1260
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
            Left            =   165
            TabIndex        =   166
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
            Left            =   9570
            TabIndex        =   164
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
            TabIndex        =   101
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
            TabIndex        =   96
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
            Left            =   6780
            TabIndex        =   95
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
            Left            =   5655
            TabIndex        =   94
            Top             =   1860
            Width           =   1110
         End
         Begin MSForms.ComboBox TXTAREA 
            Height          =   315
            Left            =   1590
            TabIndex        =   6
            Top             =   1725
            Width           =   3990
            VariousPropertyBits=   746604571
            ForeColor       =   0
            MaxLength       =   20
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "7038;556"
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
            Left            =   9540
            TabIndex        =   93
            Top             =   1665
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
            Left            =   9555
            TabIndex        =   91
            Top             =   2040
            Width           =   660
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "AREA"
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
            Left            =   735
            TabIndex        =   83
            Top             =   1695
            Width           =   825
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
            Left            =   13260
            TabIndex        =   74
            Top             =   165
            Visible         =   0   'False
            Width           =   720
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
            Left            =   3960
            TabIndex        =   72
            Top             =   165
            Width           =   855
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
            Left            =   480
            TabIndex        =   63
            Top             =   870
            Width           =   1230
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
            Left            =   1890
            TabIndex        =   62
            Top             =   165
            Width           =   630
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            Left            =   1305
            TabIndex        =   51
            Top             =   930
            Visible         =   0   'False
            Width           =   1215
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
            Height          =   315
            Left            =   915
            TabIndex        =   50
            Top             =   135
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F5E7EC&
         ForeColor       =   &H008080FF&
         Height          =   4725
         Left            =   210
         TabIndex        =   54
         Top             =   2220
         Width           =   18435
         Begin VB.Frame Frame3 
            BackColor       =   &H00F5E7EC&
            Height          =   4740
            Left            =   14190
            TabIndex        =   167
            Top             =   30
            Width           =   3285
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
               TabIndex        =   174
               Top             =   3810
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
               TabIndex        =   173
               Top             =   3825
               Width           =   510
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
               TabIndex        =   172
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
               TabIndex        =   171
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
               TabIndex        =   170
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
               TabIndex        =   169
               Top             =   3465
               Width           =   1230
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
               TabIndex        =   168
               Top             =   5175
               Width           =   1530
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "COMM AMT"
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
               Left            =   105
               TabIndex        =   200
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
               Left            =   465
               TabIndex        =   199
               Top             =   3810
               Width           =   1440
            End
            Begin VB.Label LBLRETAMT 
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
               Height          =   405
               Left            =   120
               TabIndex        =   198
               Top             =   2085
               Width           =   1545
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
               Left            =   105
               TabIndex        =   197
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
               ForeColor       =   &H00008000&
               Height          =   450
               Index           =   45
               Left            =   1755
               TabIndex        =   196
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
               Left            =   1770
               TabIndex        =   195
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   44
               Left            =   1755
               TabIndex        =   194
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
               Left            =   1770
               TabIndex        =   193
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
               Left            =   120
               TabIndex        =   192
               Top             =   300
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
               Left            =   105
               TabIndex        =   191
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
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   405
               Left            =   120
               TabIndex        =   190
               Top             =   885
               Width           =   1545
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
               Left            =   105
               TabIndex        =   189
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
               TabIndex        =   188
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
               TabIndex        =   187
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
               Left            =   1785
               TabIndex        =   186
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
               TabIndex        =   185
               Top             =   4605
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   27
               Left            =   1755
               TabIndex        =   184
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
               TabIndex        =   183
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
               TabIndex        =   182
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
               TabIndex        =   181
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
               Left            =   105
               TabIndex        =   180
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
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   405
               Left            =   120
               TabIndex        =   179
               Top             =   1485
               Width           =   1485
            End
            Begin VB.Label lblcomamt 
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
               Height          =   405
               Left            =   135
               TabIndex        =   178
               Top             =   2700
               Width           =   1545
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
               TabIndex        =   177
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
               TabIndex        =   176
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
               TabIndex        =   175
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
            TabIndex        =   102
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
            TabIndex        =   71
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
            Left            =   12210
            TabIndex        =   70
            Top             =   4260
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
            Left            =   10125
            TabIndex        =   69
            Top             =   4260
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
            Left            =   11070
            TabIndex        =   68
            Top             =   4260
            Width           =   1125
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4140
            Left            =   30
            TabIndex        =   17
            Top             =   120
            Width           =   14145
            _ExtentX        =   24950
            _ExtentY        =   7303
            _Version        =   393216
            Rows            =   1
            Cols            =   42
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   450
            BackColorFixed  =   12320767
            ForeColorFixed  =   255
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
            Left            =   6750
            TabIndex        =   100
            Top             =   4305
            Width           =   405
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
            TabIndex        =   99
            Top             =   4335
            Visible         =   0   'False
            Width           =   855
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
            Left            =   7155
            TabIndex        =   98
            Top             =   4305
            Width           =   840
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
            Left            =   6120
            TabIndex        =   97
            Top             =   4305
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L.Price"
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
            Left            =   5280
            TabIndex        =   90
            Top             =   4305
            Width           =   840
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
            TabIndex        =   82
            Top             =   4920
            Visible         =   0   'False
            Width           =   780
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
            Left            =   4050
            TabIndex        =   81
            Top             =   4305
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
            Left            =   2325
            TabIndex        =   80
            Top             =   4305
            Width           =   855
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
            Left            =   720
            TabIndex        =   79
            Top             =   4305
            Width           =   855
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
            Left            =   3240
            TabIndex        =   77
            Top             =   4305
            Width           =   795
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
            Left            =   1635
            TabIndex        =   76
            Top             =   4305
            Width           =   690
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
            Left            =   90
            TabIndex        =   75
            Top             =   4305
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F5E7EC&
         ForeColor       =   &H008080FF&
         Height          =   4365
         Left            =   210
         TabIndex        =   103
         Top             =   6840
         Width           =   18450
         Begin VB.TextBox TxtCessAmt 
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
            Left            =   13200
            MaxLength       =   5
            TabIndex        =   213
            Top             =   450
            Width           =   855
         End
         Begin VB.TextBox TxtCessPer 
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
            Left            =   12570
            MaxLength       =   5
            TabIndex        =   211
            Top             =   375
            Width           =   615
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   450
            Left            =   10770
            TabIndex        =   207
            Top             =   2085
            Visible         =   0   'False
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
            Picture         =   "frmservice.frx":0327
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   202
            Top             =   2040
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.PictureBox picChecked 
            Height          =   285
            Left            =   15435
            Picture         =   "frmservice.frx":0669
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   201
            Top             =   2010
            Visible         =   0   'False
            Width           =   285
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
            TabIndex        =   37
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
            TabIndex        =   27
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
            TabIndex        =   46
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
            TabIndex        =   47
            Top             =   1515
            Width           =   7065
         End
         Begin VB.CommandButton Command2 
            Height          =   435
            Left            =   15690
            TabIndex        =   132
            Top             =   3330
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox Txtrcvd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   12150
            MaxLength       =   7
            TabIndex        =   44
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtcommi 
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
            Left            =   14070
            MaxLength       =   6
            TabIndex        =   33
            Top             =   375
            Width           =   615
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
            Left            =   15225
            TabIndex        =   131
            Top             =   1560
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
            TabIndex        =   20
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
            TabIndex        =   130
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
            Left            =   8265
            TabIndex        =   38
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
            TabIndex        =   39
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
            TabIndex        =   129
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
            TabIndex        =   128
            Top             =   1080
            Width           =   645
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
            Left            =   11385
            TabIndex        =   40
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
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   125
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
            TabIndex        =   124
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
            TabIndex        =   24
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
            TabIndex        =   19
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
            TabIndex        =   35
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
            Left            =   6615
            TabIndex        =   36
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
            TabIndex        =   41
            Top             =   855
            Width           =   885
         End
         Begin VB.TextBox TXTDISC 
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
            Left            =   11955
            MaxLength       =   5
            TabIndex        =   32
            Top             =   375
            Width           =   600
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
            Left            =   9420
            MaxLength       =   4
            TabIndex        =   29
            Top             =   375
            Width           =   570
         End
         Begin VB.TextBox TXTQTY 
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
            Left            =   8730
            MaxLength       =   8
            TabIndex        =   25
            Top             =   375
            Width           =   675
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
            TabIndex        =   21
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
            TabIndex        =   18
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
            TabIndex        =   34
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
            TabIndex        =   28
            Top             =   2085
            Visible         =   0   'False
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
            TabIndex        =   123
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
            TabIndex        =   122
            Top             =   2250
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
            TabIndex        =   121
            Top             =   2220
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
            Left            =   9420
            MaxLength       =   7
            TabIndex        =   26
            Top             =   2085
            Visible         =   0   'False
            Width           =   525
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
            TabIndex        =   120
            Top             =   4290
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00F5E7EC&
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
            Left            =   13515
            TabIndex        =   119
            Top             =   1110
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00F5E7EC&
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
            Left            =   13500
            TabIndex        =   118
            Top             =   855
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
            TabIndex        =   116
            Top             =   3180
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.OptionButton optnet 
            BackColor       =   &H00F5E7EC&
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
            Left            =   14460
            TabIndex        =   117
            Top             =   870
            Width           =   720
         End
         Begin VB.TextBox TXTRETAILNOTAX 
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
            Left            =   10005
            MaxLength       =   9
            TabIndex        =   30
            Top             =   375
            Width           =   960
         End
         Begin VB.TextBox txtretail 
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
            Left            =   10980
            MaxLength       =   9
            TabIndex        =   31
            Top             =   375
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
            Left            =   3120
            MaxLength       =   6
            TabIndex        =   115
            Top             =   4035
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtretaildummy 
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
            MaxLength       =   6
            TabIndex        =   114
            Top             =   3660
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
            TabIndex        =   113
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
            TabIndex        =   112
            Top             =   2085
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            TabIndex        =   109
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
            TabIndex        =   108
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
            Left            =   15225
            TabIndex        =   43
            Top             =   1095
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
            TabIndex        =   107
            Top             =   2685
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.CheckBox Chkcancel 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5E7EC&
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   15240
            TabIndex        =   42
            Top             =   855
            Width           =   1230
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
            TabIndex        =   22
            Top             =   375
            Width           =   480
         End
         Begin VB.Frame FrmeType 
            BackColor       =   &H0080C0FF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   720
            Left            =   10260
            TabIndex        =   104
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
               TabIndex        =   106
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
               TabIndex        =   105
               Top             =   405
               Width           =   1020
            End
         End
         Begin MSFlexGridLib.MSFlexGrid GRDRECEIPT 
            Height          =   1125
            Left            =   0
            TabIndex        =   133
            Top             =   1920
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
            Left            =   13200
            TabIndex        =   214
            Top             =   135
            Width           =   855
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
            Left            =   12585
            TabIndex        =   212
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
            TabIndex        =   208
            Top             =   1860
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lblbarcode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   30
            TabIndex        =   206
            Top             =   1530
            Width           =   2865
         End
         Begin VB.Label lblactqty 
            Height          =   375
            Left            =   4200
            TabIndex        =   205
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
            TabIndex        =   163
            Top             =   2745
            Visible         =   0   'False
            Width           =   1560
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
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   59
            Left            =   2505
            TabIndex        =   162
            Top             =   2760
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
            Left            =   9420
            TabIndex        =   161
            Top             =   1860
            Visible         =   0   'False
            Width           =   525
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
            TabIndex        =   160
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
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   480
            Left            =   13740
            TabIndex        =   45
            Top             =   1320
            Width           =   1455
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
            Left            =   14070
            TabIndex        =   159
            Top             =   1770
            Width           =   825
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
            Left            =   12495
            TabIndex        =   158
            Top             =   1770
            Width           =   1065
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
            Left            =   14070
            TabIndex        =   157
            Top             =   150
            Width           =   615
         End
         Begin VB.Label LBLTYPE 
            Caption         =   "SV"
            Height          =   330
            Left            =   2805
            TabIndex        =   156
            Top             =   2340
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
            TabIndex        =   155
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
            TabIndex        =   154
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
            TabIndex        =   153
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
            TabIndex        =   152
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
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   14700
            TabIndex        =   151
            Top             =   375
            Width           =   1425
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Size"
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
            TabIndex        =   150
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
            TabIndex        =   149
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
            Left            =   14700
            TabIndex        =   148
            Top             =   150
            Width           =   1425
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
            Left            =   11970
            TabIndex        =   147
            Top             =   150
            Width           =   585
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
            Left            =   9420
            TabIndex        =   146
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
            Left            =   10980
            TabIndex        =   145
            Top             =   150
            Width           =   960
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
            TabIndex        =   144
            Top             =   150
            Width           =   675
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
            TabIndex        =   143
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
            TabIndex        =   142
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
            TabIndex        =   141
            Top             =   1860
            Visible         =   0   'False
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
            TabIndex        =   140
            Top             =   2325
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
            Left            =   10005
            TabIndex        =   139
            Top             =   150
            Width           =   960
         End
         Begin VB.Label lblP_Rate 
            Caption         =   "0"
            Height          =   390
            Left            =   13200
            TabIndex        =   138
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
            TabIndex        =   137
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
            TabIndex        =   136
            Top             =   840
            Width           =   1110
         End
         Begin VB.Label lblunit 
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
            Left            =   7320
            TabIndex        =   23
            Top             =   375
            Width           =   390
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
            TabIndex        =   135
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
            TabIndex        =   134
            Top             =   150
            Width           =   480
         End
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   13155
      TabIndex        =   64
      Top             =   3090
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1482
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grdcount 
      Height          =   5145
      Left            =   0
      TabIndex        =   203
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
      TabIndex        =   67
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11355
      TabIndex        =   66
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   65
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMservice"
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
Dim ITEM_CHANGE As Boolean

Private Sub cmbtype_GotFocus()
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub cmdadd_GotFocus()
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    If Val(txtretail.Text) = 0 Then
        Call TXTRETAILNOTAX_LostFocus
    End If
    If Val(TXTRETAILNOTAX.Text) = 0 Then
        Call TXTRETAIL_LostFocus
    End If
    If MDIMAIN.StatusBar.Panels(14).Text <> "Y" Then
        Call TXTRETAILNOTAX_LostFocus
    Else
        Call TXTRETAIL_LostFocus
    End If
    Call TXTDISC_LostFocus
End Sub

Private Sub CmdDeleteAll_Click()
    Dim i As Long
    Dim N As Long
    If Chkcancel.value = 0 Then Exit Sub
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rststock As ADODB.Recordset
'    If grdsales.Rows = 1 Then Exit Sub
'    If MsgBox("ARE YOU SURE YOU WANT TO CANCEL THE BILL!!!!!", vbYesNo, "DELETE!!!") = vbNo Then
'        Chkcancel.value = 0
'        Exit Sub
'    End If
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    If grdsales.Rows > 1 Then
        If frmLogin.rs!Level <> "0" And NEW_BILL = False Then
            MsgBox "Permission Denied", vbOKOnly, "Sales"
            Exit Sub
        End If
        If MsgBox("ARE YOU SURE YOU WANT TO CANCEL THE BILL!!!!!", vbYesNo, "DELETE!!!") = vbNo Then
            Chkcancel.value = 0
            Exit Sub
        End If
    End If
    
    'db.Execute "delete From RTRXFILE WHERE TRX_TYPE='CN' AND VCH_NO = " & Val(TxtCN.Text) & ""
        
    db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " "
    db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " "

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE TRX_TYPE='SV' AND VCH_NO = " & Val(TxtCN.Text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
    With RSTTRXFILE
        Do Until .EOF
            If Not (UCase(RSTTRXFILE!Category) = "SERVICES" Or UCase(RSTTRXFILE!Category) = "SELF") Then
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                With rststock
                    If Not (.EOF And .BOF) Then
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
    
    db.Execute "delete From RTRXFILE WHERE TRX_TYPE='SV' AND VCH_NO = " & Val(TxtCN.Text) & ""
    For N = 1 To grdsales.Rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(N, 13) & "' AND CATEGORY <> 'SERVICES' AND CATEGORY <> 'SERVICE CHARGE' AND CATEGORY <> 'SELF'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20)))
                !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(N, 3)) * Val(grdsales.TextMatrix(N, 27)), 3)
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(N, 20)) * Val(grdsales.TextMatrix(N, 27)), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(N, 12))
                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20))) * Val(grdsales.TextMatrix(N, 27)), 3)
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(N, 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
           
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(N, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(N, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(N, 15)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20))) * Val(grdsales.TextMatrix(N, 27)), 3)
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20))) * Val(grdsales.TextMatrix(N, 27)), 3)
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    Next N
    grdsales.FixedRows = 0
    grdsales.Rows = 1
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    Call AppendSale
    Chkcancel.value = 0
End Sub

Private Sub CMDDOS_Click()
    Chkcancel.value = 0
    If grdsales.Rows = 1 Then Exit Sub
    
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST", db, adOpenForwardOnly
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
            TXTSLNO.Text = grdsales.Rows
            'TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            TxtCessAmt.Text = ""
            TxtCessPer.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
            End If
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub CmdPrintA5_Click()
    
    Chkcancel.value = 0
    If grdsales.Rows = 1 Then Exit Sub
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    Tax_Print = False
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    If Month(Date) >= 5 And Year(Date) >= 2020 Then Exit Sub
    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2020 Then
        'db.Execute "delete From USERS "
        Exit Sub
    End If
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST", db, adOpenForwardOnly
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
    
    If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.Text <> "" Then
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
            TXTSLNO.Text = grdsales.Rows
            'TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            TxtCessAmt.Text = ""
            TxtCessPer.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
            End If
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
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
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If txtretail.Enabled = True Then txtretail.SetFocus
            If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            ''If txtcommi.Enabled = True Then txtcommi.SetFocus
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
    
    On Error GoTo eRRhAND
    'If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM SALERETURN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND CHECK_FLAG <> 'Y'", db, adOpenStatic, adLockOptimistic, adCmdText
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
        grdsales.TextMatrix(grdsales.Rows - 1, 10) = ""
        grdsales.TextMatrix(grdsales.Rows - 1, 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST)
        grdsales.TextMatrix(grdsales.Rows - 1, 12) = Format(RSTTRXFILE!TRX_TOTAL, ".000")
        
        grdsales.TextMatrix(grdsales.Rows - 1, 13) = RSTTRXFILE!ITEM_CODE
        grdsales.TextMatrix(grdsales.Rows - 1, 14) = RSTTRXFILE!VCH_NO
        grdsales.TextMatrix(grdsales.Rows - 1, 15) = RSTTRXFILE!line_no
        grdsales.TextMatrix(grdsales.Rows - 1, 16) = RSTTRXFILE!TRX_TYPE
        grdsales.TextMatrix(grdsales.Rows - 1, 17) = "N"
        Set RSTMFGR = New ADODB.Recordset
        RSTMFGR.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(grdsales.Rows - 1, 1)) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTMFGR.EOF And RSTMFGR.BOF) Then
            grdsales.TextMatrix(grdsales.Rows - 1, 18) = IIf(IsNull(RSTMFGR!MANUFACTURER), "", Trim(RSTMFGR!MANUFACTURER))
        End If
        RSTMFGR.Close
        Set RSTMFGR = Nothing
        grdsales.TextMatrix(grdsales.Rows - 1, 19) = "CN"
        grdsales.TextMatrix(grdsales.Rows - 1, 20) = "0"
        
        grdsales.TextMatrix(grdsales.Rows - 1, 21) = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
        grdsales.TextMatrix(grdsales.Rows - 1, 22) = IIf(IsNull(RSTTRXFILE!P_RETAILWOTAX), 0, RSTTRXFILE!P_RETAILWOTAX)
        grdsales.TextMatrix(grdsales.Rows - 1, 23) = IIf(IsNull(RSTTRXFILE!SALE_1_FLAG), "2", RSTTRXFILE!SALE_1_FLAG)
        grdsales.TextMatrix(grdsales.Rows - 1, 24) = IIf(IsNull(RSTTRXFILE!COM_AMT), "", RSTTRXFILE!COM_AMT)
        grdsales.TextMatrix(grdsales.Rows - 1, 25) = IIf(IsNull(RSTTRXFILE!Category), 0, RSTTRXFILE!Category)
        grdsales.TextMatrix(grdsales.Rows - 1, 26) = IIf(IsNull(RSTTRXFILE!LOOSE_FLAG), "F", RSTTRXFILE!LOOSE_FLAG)
        grdsales.TextMatrix(grdsales.Rows - 1, 27) = IIf(IsNull(RSTTRXFILE!LOOSE_PACK), "1", RSTTRXFILE!LOOSE_PACK)
        
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
                LBLFOT.Caption = ""
            Case Else
                LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))), 2)
                LBLFOT.Caption = ""
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
        End Select
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    lblnetamount.Caption = Round(lblnetamount.Caption, 0)
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    txtPrintname.Text = ""
    txtcategory.Text = ""
    TxtName1.Text = ""
    TXTITEMCODE.Text = ""
    optnet.value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    
    TXTQTY.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTAPPENDQTY.Text = ""
    TXTAPPENDTOTAL.Text = ""
    TXTFREEAPPEND.Text = ""
    txtappendcomm.Text = ""
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    txtcommi.Text = ""
    TxtRetailmode.Text = ""
    txtretail.Text = ""
    txtBatch.Text = ""
    TxtWarranty.Text = ""
    TxtWarranty_type.Text = ""
    TXTTAX.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    TXTDISC.Text = ""
    TxtCessAmt.Text = ""
    TxtCessPer.Text = ""
    cmdadd.Enabled = False
    'cmddelete.Enabled = False
    CMDEXIT.Enabled = False
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
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub CmdTax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub Command1_Click()

    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim DN_CN_FLag As Boolean
    Dim i As Long
    Dim CN As Integer
    Dim DN As Integer
    Dim b As Integer
    Dim Num, Figre As Currency
    
    If grdsales.Rows <= 1 Then Exit Sub
    On Error GoTo eRRhAND
    db.Execute "delete From TEMPTRXFILE "
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "SV"
        'RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!line_no = i
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
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
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
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
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(i, 24))
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 27))
        RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(i, 28) = "", Null, grdsales.TextMatrix(i, 28))
        RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(i, 29)
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 30)
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(i, 26)
        
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
        Dim RSTITEMMAST As ADODB.Recordset
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!Area = RSTITEMMAST!Area
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.Update
    Next i
    
    Dim rstTRXMAST As ADODB.Recordset
    Set rstTRXMAST = New ADODB.Recordset
    rstTRXMAST.Open "SELECT * from RTRXFILE WHERE TRX_TYPE='SV' AND VCH_NO = " & Val(TxtCN.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstTRXMAST.EOF
        i = i + 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "XC"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!line_no = i
        RSTTRXFILE!Category = ""
        RSTTRXFILE!ITEM_CODE = rstTRXMAST!ITEM_CODE
        RSTTRXFILE!ITEM_NAME = rstTRXMAST!ITEM_NAME
        RSTTRXFILE!QTY = rstTRXMAST!QTY
        RSTTRXFILE!MRP = rstTRXMAST!MRP
        RSTTRXFILE!PTR = rstTRXMAST!PTR
        RSTTRXFILE!SALES_PRICE = -rstTRXMAST!SALES_PRICE
        RSTTRXFILE!SALES_TAX = rstTRXMAST!SALES_TAX
        RSTTRXFILE!UNIT = rstTRXMAST!UNIT
        RSTTRXFILE!VCH_DESC = "Returned From  " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = rstTRXMAST!REF_NO
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = "V"
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
        RSTTRXFILE!C_USER_ID = "SM"
                
        RSTTRXFILE.Update
        
        rstTRXMAST.MoveNext
    Loop
    rstTRXMAST.Close
    Set rstTRXMAST = Nothing
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
    'Call ReportGeneratION_vpestimate
    LBLFOT.Tag = ""
    Screen.MousePointer = vbHourglass
    Sleep (300)
    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0)) - Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 2)), 2)
    Figre = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0))
    Num = Abs(Figre)
    If Figre < 0 Then
        LBLFOT.Tag = "(-)Rupees " & Words_1_all(Num) & " Only"
    ElseIf Figre > 0 Then
        LBLFOT.Tag = "(Rupees " & Words_1_all(Num) & " Only)"
    End If
    ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTcommission"
    
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.RecordSelectionFormula = "( {TRXFILE.TRX_TYPE}='SV' AND {TRXFILE.VCH_NO}= " & Val(txtBillNo.Text) & ")"
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    Report.DiscardSavedData
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & TxtBillName.Text & "'"
        If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.Text = "'" & TXTDEALER.Text & "'"
        If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.Text = "'" & lbladdress.Caption & "'"
        If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.Text = "'" & TXTDEALER.Text & "'"
        If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.Text = "'" & lbladdress.Caption & "'"
        If TxtPhone.Text = "" Then
            If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
        Else
            If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
            'If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "' & chr(13) & 'Ph: ' & '" & Trim(TxtPhone.Text) & "'"
        End If
        'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
'            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
'            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
        If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
        If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
        If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
        If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & TxtPhone.Text & "'"
        If CRXFormulaField.Name = "{@VCH_NO}" Then CRXFormulaField.Text = "'" & Trim(txtBillNo.Text) & "' & '-A' "
        If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
        If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = " " & Val(LBLDISCAMT.Caption) & " "
'            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
        If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.Text = "'" & Trim(TxtCode.Text) & "'"
        If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.Text = " " & Val(txtOutstanding.Text) & " "
        If CRXFormulaField.Name = "{@Frieght}" Then CRXFormulaField.Text = "'" & Trim(lblFrieght.Text) & "'"
        If CRXFormulaField.Name = "{@FC}" Then CRXFormulaField.Text = " " & Val(TxtFrieght.Text) & " "
        If CRXFormulaField.Name = "{@HANDLE}" Then CRXFormulaField.Text = " '" & Trim(lblhandle.Text) & "' "
            If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.Text = " " & Val(Txthandle.Text) & " "
        If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.Text = "'" & Val(Txthandle.Text) & "'"
        If Val(LBLRETAMT.Caption) = 0 Then
            If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'N' "
        Else
            If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'Y' "
        End If
        If lblcredit.Caption = "0" Then
            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Cash'"
        Else
            If Val(txtcrdays.Text) > 0 Then
                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'" & txtcrdays.Text & "'" & "' Days Credit'"
            Else
                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Credit'"
            End If
        End If
    Next
    frmreport.Caption = "BILL"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description

End Sub

Private Sub Command2_Click()
    Dim rststock As ADODB.Recordset
    'Dim RSTMINQTY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim i As Long
    
    Chkcancel.value = 0
    On Error GoTo eRRhAND
    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2020 Then Exit Sub
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
   
    'grdsales.TextMatrix(I, 17) = "N"
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.Caption = ""
    
    For i = 1 To grdsales.Rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT P_RETAIL  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenStatic, adLockReadOnly
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
        RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "SV"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
            RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
            RSTTRXFILE!line_no = Val(grdsales.TextMatrix(i, 32))
        End If
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
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
        If CMBDISTI.BoundText <> "" Then
            RSTTRXFILE!COM_FLAG = "Y"
        Else
            RSTTRXFILE!COM_FLAG = "N"
        End If
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(i, 26)
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 27))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.Text), 1, 30)
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 10))
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
        Val (TXTSLNO.Text)
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
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
        Else
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
        End If
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    lblnetamount.Caption = Round(lblnetamount.Caption, 0)
    Call cmdRefresh_Click
    txtBillNo.Visible = True
    txtBillNo.Enabled = True
    txtBillNo.SetFocus
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Command3_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim DN_CN_FLag As Boolean
    Dim i As Long
    Dim CN As Integer
    Dim DN As Integer
    Dim b As Integer
    Dim Num, Figre As Currency
    
    On Error GoTo eRRhAND
    
    If grdsales.Rows <= 1 Then Exit Sub
    
    If OLD_BILL = False Then Call checklastbill
    OLD_BILL = True
'    db.Execute "delete From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & ""
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "Select * FROM TRXMAST ", db, adOpenStatic, adLockOptimistic, adCmdText
'    RSTTRXFILE.AddNew
'    RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
'    RSTTRXFILE!TRX_TYPE = "SV"
'    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
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
'    RSTTRXFILE!BILL_NAME = Trim(TxtBillName.Text)
'    RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.Text)
'    RSTTRXFILE!Area = Trim(TXTAREA.Text)
'    'RSTTRXFILE!BILL_FLAG = "Y"
'    RSTTRXFILE.Update
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    
    db.Execute "delete from TEMPTRXFILE"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "SV"
        'RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!line_no = i
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 13)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 5))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "" '"Issued to     " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
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
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
        If Tax_Print = False Then
            RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        Else
            RSTTRXFILE!SALES_TAX = Val(TxtTaxPrint.Text)
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
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 30)
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(i, 26)
        
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
'        Dim RSTITEMMAST As ADODB.Recordset
'        Set RSTITEMMAST = New ADODB.Recordset
'        RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
'        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'            RSTTRXFILE!Area = RSTITEMMAST!Area
'        End If
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.Update
    Next i
    
    Dim rstTRXMAST As ADODB.Recordset
    Set rstTRXMAST = New ADODB.Recordset
    rstTRXMAST.Open "SELECT * from RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(TxtCN.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstTRXMAST.EOF
        i = i + 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "XC"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!line_no = i
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
        RSTTRXFILE!CHECK_FLAG = "V"
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
        RSTTRXFILE!C_USER_ID = "SM"
                
        RSTTRXFILE.Update
        
        rstTRXMAST.MoveNext
    Loop
    rstTRXMAST.Close
    Set rstTRXMAST = Nothing
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
'    Call ReportGeneratION_estimate
'    On Error GoTo CLOSEFILE
'    Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\repo.bat" For Output As #1 '//Creating Batch file
'CLOSEFILE:
'    If Err.Number = 55 Then
'        Close #1
'        Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\repo.bat" For Output As #1 '//Creating Batch file
'    End If
'    On Error GoTo eRRhAND
'
'    Print #1, "TYPE " & MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\Report.txt > PRN"
'    Print #1, "EXIT"
'    Close #1
'
'    '//HERE write the proper path where your command.com file exist
'    Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\REPO.BAT N", vbHide
'    Exit Sub
    
    LBLFOT.Tag = ""
    Screen.MousePointer = vbHourglass
    Sleep (300)
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompAddress4, CompTin, CompCST, BIL_PRE, BILL_SUF As String
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        CompName = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
        CompAddress1 = IIf(IsNull(RSTCOMPANY!Address), "", RSTCOMPANY!Address)
        CompAddress2 = IIf(IsNull(RSTCOMPANY!HO_NAME), "", RSTCOMPANY!HO_NAME)
        CompAddress3 = "Ph: " & IIf(IsNull(RSTCOMPANY!TEL_NO), "", RSTCOMPANY!TEL_NO) & IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", ", " & RSTCOMPANY!FAX_NO)
        CompAddress4 = IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!EMAIL_ADD)
        CompTin = IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", "", "GSTIN No. " & RSTCOMPANY!CST)
        CompCST = IIf(IsNull(RSTCOMPANY!DL_NO) Or RSTCOMPANY!DL_NO = "", "", "CST No. " & RSTCOMPANY!DL_NO)
'        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
'        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        If Trim(TxtVehicle.Text) = "" Then TxtVehicle.Text = IIf(IsNull(RSTCOMPANY!VEHICLE), "", RSTCOMPANY!VEHICLE)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
                                    
    NEW_BILL = False
    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0)) - Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 2)), 2)
    Figre = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0))
    Num = Abs(Figre)
    If Figre < 0 Then
        LBLFOT.Tag = "(-)Rupees " & Words_1_all(Num) & " Only"
    ElseIf Figre > 0 Then
        LBLFOT.Tag = "(Rupees " & Words_1_all(Num) & " Only)"
    End If
    If Small_Print = True Then
        'ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\rptbillretail"
        ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTOUTPASS"
    Else
        ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTOUTPASS"
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    Report.DiscardSavedData
    Set CRXFormulaFields = Report.FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.Text = "'" & CompName & "'"
        If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.Text = "'" & CompAddress1 & "'"
        If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.Text = "'" & CompAddress2 & "'"
        If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.Text = "'" & CompAddress3 & "'"
        If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.Text = "'" & CompAddress4 & "'"
        If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.Text = "'" & CompTin & "'"
        If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.Text = "'" & CompCST & "'"
        If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & TxtBillName.Text & "'"
        If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.Text = "'" & Trim(lbladdress.Caption) & "'"
        If TxtPhone.Text = "" Then
            If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
        Else
            If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
            'If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "' & chr(13) & 'Ph: ' & '" & Trim(TxtPhone.Text) & "'"
        End If
        'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
'            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
'            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
        If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
'        If Tax_Print = False Then
'            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
'        End If
        If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
            If CRXFormulaField.Name = "{@condition}" Then CRXFormulaField.Text = "'" & Trim(Terms1.Text) & "'"
        End If
        If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
        If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & TxtPhone.Text & "'"
        If CRXFormulaField.Name = "{@VCH_NO}" Then CRXFormulaField.Text = "'" & BIL_PRE & "' &'" & Format(Trim(txtBillNo.Text), "0000") & "' & '" & BILL_SUF & "' "
        If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
        If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.Text = "'" & Trim(TxtOrder.Text) & "'"
        If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = " " & Val(LBLDISCAMT.Caption) & " "
'            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
        If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.Text = "'" & Trim(TxtCode.Text) & "'"
        If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.Text = " " & Val(txtOutstanding.Text) & " "
        If CRXFormulaField.Name = "{@Frieght}" Then CRXFormulaField.Text = "'" & Trim(lblFrieght.Text) & "'"
        If CRXFormulaField.Name = "{@FC}" Then CRXFormulaField.Text = " " & Val(TxtFrieght.Text) & " "
        If CRXFormulaField.Name = "{@HANDLE}" Then CRXFormulaField.Text = " '" & Trim(lblhandle.Text) & "' "
        If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.Text = " " & Val(Txthandle.Text) & " "
        If CRXFormulaField.Name = "{@DISCPER}" Then CRXFormulaField.Text = " " & Val(TXTTOTALDISC.Text) & " "
        
        If Val(LBLRETAMT.Caption) = 0 Then
            If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'N' "
        Else
            If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'Y' "
        End If
        If lblcredit.Caption = "0" Then
            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Cash'"
        Else
            If Val(txtcrdays.Text) > 0 Then
                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'" & txtcrdays.Text & "'" & "' Days Credit'"
            Else
                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Credit'"
            End If
        End If
    Next
    
    
    If MDIMAIN.StatusBar.Panels(13).Text = "Y" Then
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
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Public Sub DataList2_Click()
    Dim rstCustomer As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    On Error GoTo eRRhAND
    
    If CHANGE_ADDRESS = True Then
        Set rstCustomer = New ADODB.Recordset
        rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstCustomer.EOF And rstCustomer.BOF) Then
            lbladdress.Caption = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
            'If TxtBillName.Text = "" Then TxtBillName.Text = DataList2.Text
            If Len(DataList2.Text) > 11 Then
                TxtBillName.Text = Mid(DataList2.Text, 12)
            Else
                TxtBillName.Text = DataList2.Text
            End If
            TxtBillName.Text = DataList2.Text
            'If TxtBillAddress.Text = "" Then TxtBillAddress.Text = IIf(IsNull(rstCustomer!ADDRESS), "", Trim(rstCustomer!ADDRESS))
            TxtBillAddress.Text = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
            TXTTIN.Text = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
            TxtPhone.Text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO))
            TXTAREA.Text = IIf(IsNull(rstCustomer!Area), "", Trim(rstCustomer!Area))
            Select Case rstCustomer!Type
                Case "W"
                    cmbtype.ListIndex = 1
                    TXTTYPE.Text = 2
                Case "V"
                    cmbtype.ListIndex = 2
                    TXTTYPE.Text = 3
                Case Else
                    cmbtype.ListIndex = 0
                    TXTTYPE.Text = 1
            End Select
            'lblcusttype.Caption = IIf((IsNull(rstCustomer!Type) Or rstCustomer!Type = "R"), "R", "W")
            CMBBRNCH.Text = ""
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
            CMBBRNCH.Text = ""
            Set CMBBRNCH.RowSource = Nothing
            TxtPhone.Text = ""
            TXTTIN.Text = ""
            lbladdress.Caption = ""
            TXTAREA.Text = ""
            TxtVehicle.Text = ""
            TxtOrder.Text = ""
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
    If cr_days = False Then
        Set rstCustomer = New ADODB.Recordset
        rstCustomer.Open "Select PYMT_PERIOD, ACT_CODE From CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
        If Not (rstCustomer.EOF And rstCustomer.BOF) Then
            txtcrdays.Text = IIf(IsNull(rstCustomer!PYMT_PERIOD), "", rstCustomer!PYMT_PERIOD)
        End If
        rstCustomer.Close
        Set rstCustomer = Nothing
    End If

SKIP:
    If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
        txtcrdays.Enabled = False
        Frame5.Visible = False
    Else
        txtcrdays.Enabled = True
        Frame5.Visible = True
    End If
    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text
    TxtCode.Text = DataList2.BoundText
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
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
            If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
                TxtBillName.SetFocus
            Else
                TxtName1.Enabled = True
                TXTPRODUCT.Enabled = True
                TXTITEMCODE.Enabled = True
                If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                    TXTPRODUCT.SetFocus
                Else
                    TXTITEMCODE.SetFocus
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
    
    If Val(TXTQTY.Text) = 0 And Val(TXTFREE.Text) = 0 Then
        MsgBox "Please enter the Qty", vbOKOnly, "Sales"
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
        Exit Sub
    End If
    If MDIMAIN.LBLTAXWARN.Caption = "Y" Then
        If Val(TXTTAX.Text) = 0 Then
            If (MsgBox("Tax is Zero. Are you sure?", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
                Exit Sub
            End If
        End If
    End If
    If Val(TXTRETAILNOTAX.Text) = 0 Then
        Call TXTRETAIL_LostFocus
    End If
    If MDIMAIN.StatusBar.Panels(14).Text <> "Y" Then
        Call TXTRETAILNOTAX_LostFocus
    End If
    If Val(TXTQTY.Text) <> 0 And MDIMAIN.StatusBar.Panels(14).Text <> "Y" And Val(TXTRETAILNOTAX.Text) = 0 Then
        MsgBox "Please enter the Rate", vbOKOnly, "Sales"
        TXTRETAILNOTAX.Enabled = True
        TXTRETAILNOTAX.SetFocus
        Exit Sub
    End If
    If Val(TXTQTY.Text) <> 0 And MDIMAIN.StatusBar.Panels(14).Text = "Y" And Val(txtretail.Text) = 0 Then
        MsgBox "Please enter the Rate", vbOKOnly, "Sales"
        txtretail.Enabled = True
        txtretail.SetFocus
        Exit Sub
    End If
    If MDIMAIN.StatusBar.Panels(14).Text <> "Y" Then
        Call TXTRETAILNOTAX_LostFocus
    Else
        Call TXTRETAIL_LostFocus
    End If
    Call TXTDISC_LostFocus
'    Call TxtCessPer_LostFocus
    
    Dim rststock As ADODB.Recordset
    'Dim RSTMINQTY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim i As Long
    
    Chkcancel.value = 0
    On Error GoTo eRRhAND
    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2020 Then Exit Sub
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    If OLD_BILL = False Then Call checklastbill
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "SV"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!CUST_IGST = lblIGST.Caption
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.Text
        RSTTRXFILE!BILL_NAME = Trim(TxtBillName.Text)
        RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.Text)
        RSTTRXFILE!Area = Trim(TXTAREA.Text)
        RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
        RSTTRXFILE!DISC_PERS = Val(TXTTOTALDISC.Text)
        RSTTRXFILE!ADD_AMOUNT = 0
        RSTTRXFILE!ROUNDED_OFF = 0
        RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
        RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
        If Val(TxtCN.Text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.Text)
        RSTTRXFILE!BILL_FLAG = "Y"
        RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
        RSTTRXFILE!BR_NAME = CMBBRNCH.Text
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text) + Val(TXTAPPENDQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Format(Val(TxtMRP.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTRETAILNOTAX.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Format(Val(txtretail.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(TXTDISC.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Val(TXTTAX.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Val(LBLITEMCOST.Caption)
    
    TXTDISC.Tag = 0
    If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
        TXTAPPENDTOTAL.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
    Else
        TXTDISC.Tag = Val(TXTAPPENDQTY.Text) * Val(txtretail.Text) * Val(TXTDISC.Text) / 100
        TXTAPPENDTOTAL.Text = Format((Val(TXTAPPENDQTY.Text) * Round(Val(txtretail.Text), 3)) - Val(TXTDISC.Tag), ".000")
    End If
    
    TXTAPPENDTOTAL.Text = ""
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Format(Val(LBLSUBTOTAL.Caption) + Val(TXTAPPENDTOTAL.Text), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Trim(TXTTRXTYPE.Text)
    
    If OPTVAT.value = True And Val(TXTTAX.Text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "V"
    If OPTTaxMRP.value = True And Val(TXTTAX.Text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "M"
    If Val(TXTTAX.Text) <= 0 Or optnet.value = True Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "N"
    
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "N"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = IIf(IsNull(RSTTRXFILE!MANUFACTURER), "", Trim(RSTTRXFILE!MANUFACTURER))
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
    grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = Val(TXTFREE.Text) + Val(TXTFREEAPPEND.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 21) = Format(Val(txtretail.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 22) = Format(Val(TXTRETAILNOTAX.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 23) = Trim(TXTSALETYPE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = Val(txtcommi.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = Trim(txtcategory.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 26) = "L"
    grdsales.TextMatrix(Val(TXTSLNO.Text), 27) = IIf(Val(LblPack.Text) = 0, "1", Val(LblPack.Text))
    grdsales.TextMatrix(Val(TXTSLNO.Text), 28) = Val(TxtWarranty.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 29) = Trim(TxtWarranty_type.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 30) = Trim(lblunit.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 38) = IIf(TXTEXPIRY.Text = "  /  ", "", Trim(TXTEXPIRY.Text))
    grdsales.TextMatrix(Val(TXTSLNO.Text), 39) = Val(lblretail.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 40) = Val(TxtCessPer.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 41) = Val(TxtCessAmt.Text)
    If Trim(txtPrintname.Text) = "" Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 33) = Trim(TXTPRODUCT.Text)
    Else
        grdsales.TextMatrix(Val(TXTSLNO.Text), 33) = Trim(txtPrintname.Text)
    End If
    grdsales.TextMatrix(Val(TXTSLNO.Text), 34) = Val(LblGross.Caption)
    If M_EDIT = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32))
    Else
        i = 0
        Dim rstMaxNo As ADODB.Recordset
        Set rstMaxNo = New ADODB.Recordset
        rstMaxNo.Open "Select MAX(LINE_NO) From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
        If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
            grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        Else
            grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = Val(TXTSLNO.Text)
        End If
        rstMaxNo.Close
        Set rstMaxNo = Nothing
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "SV"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!line_no = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32))
    End If
    RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = "", 0, grdsales.TextMatrix(Val(TXTSLNO.Text), 14))
    RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = "", 0, grdsales.TextMatrix(Val(TXTSLNO.Text), 15))
    RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = "", "MI", grdsales.TextMatrix(Val(TXTSLNO.Text), 16))
    RSTTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "SV"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!line_no = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32))
    End If
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
    RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
    RSTTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
    RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
    RSTTRXFILE!MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
    RSTTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
    RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
    RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21))
    RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22))
    RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
    RSTTRXFILE!Category = grdsales.TextMatrix(Val(TXTSLNO.Text), 25)
    If CMBDISTI.BoundText <> "" Then
        RSTTRXFILE!COM_FLAG = "Y"
    Else
        RSTTRXFILE!COM_FLAG = "N"
    End If
    RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(Val(TXTSLNO.Text), 26)
    RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 27))
    RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
    RSTTRXFILE!UNIT = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
    RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.Text), 1, 30)
    RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
    RSTTRXFILE!ISSUE_QTY = 0
    RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
    RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
    Select Case grdsales.TextMatrix(Val(TXTSLNO.Text), 19)
        Case "DN"
            RSTTRXFILE!CST = 1
        Case "CN"
            RSTTRXFILE!CST = 2
        Case Else
            RSTTRXFILE!CST = 0
    End Select
    RSTTRXFILE!BAL_QTY = 0
    RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(Val(TXTSLNO.Text), 12)
    RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
    RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
    RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!M_USER_ID = DataList2.BoundText
    RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23))
    RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 28) = "", Null, grdsales.TextMatrix(Val(TXTSLNO.Text), 28))
    RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(Val(TXTSLNO.Text), 29)
    RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(Val(TXTSLNO.Text), 30)
    RSTTRXFILE!ST_RATE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 31))
    RSTTRXFILE!RETAILER_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39))
    RSTTRXFILE!CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40))
    RSTTRXFILE!CESS_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 41))
    If Not IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 38)) Then
        'RSTTRXFILE!EXP_DATE = Null
    Else
        RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 38))) & "/" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 38))
    End If
    
    If Trim(grdsales.TextMatrix(i, 33)) = "" Then
        RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2))
    Else
        RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 33))
    End If
    RSTTRXFILE!GROSS_AMOUNT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 34))
    RSTTRXFILE!DN_NO = Val(grdsales.TextMatrix(i, 35))
    If IsDate(grdsales.TextMatrix(i, 36)) Then
        RSTTRXFILE!DN_DATE = IIf(IsDate(grdsales.TextMatrix(i, 36)), Format(grdsales.TextMatrix(i, 36), "DD/MM/YYYY"), Null)
    End If
    RSTTRXFILE!DN_LINENO = Val(grdsales.TextMatrix(i, 37))
        
    RSTTRXFILE.Update

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    If Not (UCase(txtcategory.Text) = "SERVICES" Or UCase(txtcategory.Text) = "SELF") Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & " AND BAL_QTY > 0", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text), 3)
                !BAL_QTY = !BAL_QTY - Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text), 3)
                grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(LblPack.Text))
                RSTTRXFILE.Update
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            Else
                'BALQTY = 0
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "' AND BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    If (IsNull(RSTTRXFILE!ISSUE_QTY)) Then RSTTRXFILE!ISSUE_QTY = 0
                    If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                    'BALQTY = RSTTRXFILE!BAL_QTY
                    RSTTRXFILE!ISSUE_QTY = RSTTRXFILE!ISSUE_QTY + Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text), 3)
                    RSTTRXFILE!BAL_QTY = RSTTRXFILE!BAL_QTY - Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text), 3)
                    
                    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = RSTTRXFILE!VCH_NO
                    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = RSTTRXFILE!line_no
                    grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = RSTTRXFILE!TRX_TYPE
                    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(LblPack.Text))
                    RSTTRXFILE.Update
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                Else
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(LblPack.Text))
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
            End If
        End With
        
'        Dim RET_PRICE, LOOS_PRICE, LOOSE_PCK, ITEM_CST As Double
'        Set RSTTRXFILE = Nothing
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "' AND  BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockReadOnly
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
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
'                If RET_PRICE > 0 Then
'                    !P_RETAIL = RET_PRICE
'                    !P_CRTN = LOOS_PRICE
'                    !LOOSE_PACK = LOOSE_PCK
'                    If ITEM_CST > 0 Then !ITEM_COST = ITEM_CST
'                End If
                '!ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTQTY.Text) * Val(LblPack.Text)), 3)
                !FREE_QTY = !FREE_QTY + Round((Val(TXTFREE.Text) * Val(LblPack.Text)), 3)
                !CLOSE_QTY = !CLOSE_QTY - Round(((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text)), 3)
    
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
        Else
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
        End If
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    lblnetamount.Caption = Round(lblnetamount.Caption, 0)
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    txtPrintname.Text = ""
    txtcategory.Text = ""
    'TxtName1.Text = ""
    TXTITEMCODE.Text = ""
    optnet.value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    
    lblactqty.Caption = ""
    lblbarcode.Caption = ""
    lblretail.Caption = ""
    lblwsale.Caption = ""
    lblvan.Caption = ""
    lblunit.Caption = ""
    LblPack.Text = ""
    lblOr_Pack.Caption = ""
    lblcase.Caption = ""
    lblcrtnpack.Caption = ""
    LBLITEMCOST.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTAPPENDQTY.Text = ""
    TXTFREEAPPEND.Text = ""
    txtappendcomm.Text = ""
    TXTAPPENDTOTAL.Text = ""
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    txtcommi.Text = ""
    TxtRetailmode.Text = ""
    txtretail.Text = ""
    txtBatch.Text = ""
    TXTTAX.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    TXTDISC.Text = ""
    TxtCessAmt.Text = ""
    TxtCessPer.Text = ""
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    TxtWarranty.Text = ""
    TxtWarranty_type.Text = ""
    lblP_Rate.Caption = "0"
    'cmddelete.Enabled = False
    CMDEXIT.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    CMDPRINT.Enabled = True
    CmdPrintA5.Enabled = True
    cmdRefresh.Enabled = True
    
    CmdDelete.Enabled = True
    CMDMODIFY.Enabled = True
    'TxtName1.Enabled = True
    M_EDIT = False
    M_ADD = True
    OLD_BILL = True
    Call COSTCALCULATION
    Call Addcommission
    If grdsales.Rows >= 9 Then grdsales.TopRow = grdsales.Rows - 1
    If UCase(Trim(grdsales.TextMatrix(1, 25))) = "HOME APPLIANCES" Then
        chkTerms.value = 1
    Else
        chkTerms.value = 0
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
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    
    If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
        TXTPRODUCT.SetFocus
    Else
        TXTITEMCODE.SetFocus
    End If
    'grdsales.TopRow = grdsales.Rows - 1
Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            If MDIMAIN.StatusBar.Panels(16).Text = "Y" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTDISC.Enabled = True
                TXTDISC.SetFocus
            End If
            'TxtWarranty.Enabled = True
            'TxtWarranty.SetFocus
    End Select

End Sub

Private Sub CmdDelete_Click()
    
    If grdsales.Rows <= 1 Then Exit Sub
    If M_EDIT = True Then Exit Sub
    If frmLogin.rs!Level <> "0" And NEW_BILL = False Then
        MsgBox "Permission Denied", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    TXTSLNO.Text = grdsales.TextMatrix(grdsales.Row, 0)
    Call TXTSLNO_KeyDown(13, 0)
    grdsales.Enabled = True

    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & ""
    db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & ""
    If Not (UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICES" Or UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SELF") Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(LblPack.Text), 3)
                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) * Val(LblPack.Text), 3)
                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
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
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
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
        grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = grdsales.TextMatrix(i + 1, 24)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = grdsales.TextMatrix(i + 1, 25)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 26) = grdsales.TextMatrix(i + 1, 26)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 27) = grdsales.TextMatrix(i + 1, 27)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 28) = grdsales.TextMatrix(i + 1, 28)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 29) = grdsales.TextMatrix(i + 1, 29)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 30) = grdsales.TextMatrix(i + 1, 30)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 31) = grdsales.TextMatrix(i + 1, 31)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = grdsales.TextMatrix(i + 1, 32)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 33) = grdsales.TextMatrix(i + 1, 33)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 34) = grdsales.TextMatrix(i + 1, 34)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 35) = grdsales.TextMatrix(i + 1, 35)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 36) = grdsales.TextMatrix(i + 1, 36)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 37) = grdsales.TextMatrix(i + 1, 37)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 38) = grdsales.TextMatrix(i + 1, 38)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 39) = grdsales.TextMatrix(i + 1, 39)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 40) = grdsales.TextMatrix(i + 1, 40)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 41) = grdsales.TextMatrix(i + 1, 41)
    Next i
    grdsales.Rows = grdsales.Rows - 1
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
        Else
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
        End If
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    lblnetamount.Caption = Round(lblnetamount.Caption, 0)
    
    Call COSTCALCULATION
    Call Addcommission
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    txtPrintname.Text = ""
    txtcategory.Text = ""
    TxtName1.Text = ""
    TXTITEMCODE.Text = ""
    optnet.value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTAPPENDQTY.Text = ""
    TXTFREEAPPEND.Text = ""
    txtappendcomm.Text = ""
    TXTAPPENDTOTAL.Text = ""
    txtretail.Text = ""
    txtBatch.Text = ""
    TxtWarranty.Text = ""
    TxtWarranty_type.Text = ""
    TXTTAX.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    TxtMRP.Text = ""
    lblactqty.Caption = ""
    lblbarcode.Caption = ""
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    txtcommi.Text = ""
    TxtRetailmode.Text = ""
    
    TXTDISC.Text = ""
    TxtCessAmt.Text = ""
    TxtCessPer.Text = ""
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    LBLDNORCN.Caption = ""
    cmdadd.Enabled = False
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
        TXTPRODUCT.SetFocus
    Else
        TXTITEMCODE.SetFocus
    End If
    'cmddelete.Enabled = False
    'CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    M_ADD = True
    If grdsales.Rows = 1 Then
'        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        
        CmdPrintA5.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    If grdsales.Rows >= 9 Then grdsales.TopRow = grdsales.Rows - 1
End Sub

Private Sub CmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            txtPrintname.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            TxtCessAmt.Text = ""
            TxtCessPer.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
            End If
            TXTQTY.Enabled = False
            
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
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
        LblProfitAmt.Visible = True
        LBLITEMCOST.Visible = True
        'LBLSELPRICE.Visible = True
    End If
End Sub

Private Sub CMDMODIFY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    
    If grdsales.Rows <= 1 Then Exit Sub
    'If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    If M_EDIT = True Then Exit Sub
    If frmLogin.rs!Level <> "0" And NEW_BILL = False Then
        MsgBox "Permission Denied", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    TXTSLNO.Text = grdsales.TextMatrix(grdsales.Row, 0)
    If grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = "DN" Then
        MsgBox "Cannot modify this. The Item is being Delivered. DN# ", vbOKOnly, "Sales"
        Exit Sub
    End If
    Call TXTSLNO_KeyDown(13, 0)
    grdsales.Enabled = True
    
    If UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICE CHARGE" Then
        CMDMODIFY.Enabled = False
        CmdDelete.Enabled = False
        CMDEXIT.Enabled = False
        M_EDIT = True
        txtretail.Enabled = True
        txtretail.SetFocus
        Exit Sub
    End If
    
    M_ADD = True
    On Error GoTo eRRhAND
    If Not (UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICES" Or UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SELF") Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(LblPack.Text), 3)
                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) * Val(LblPack.Text), 3)
                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
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
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
                lblactqty.Caption = !BAL_QTY
                lblbarcode.Caption = IIf(IsNull(!BARCODE), "", !BARCODE)
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
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
            txtPrintname.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            
            optnet.value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            TxtCessAmt.Text = ""
            TxtCessPer.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.Text = ""
            
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
            End If
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub CmdPrint_Click()
        
    Chkcancel.value = 0
    If grdsales.Rows = 1 Then Exit Sub
    
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    Tax_Print = False
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    If Month(Date) >= 5 And Year(Date) >= 2020 Then Exit Sub
    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2020 Then
        db.Execute "delete From USERS "
        Exit Sub
    End If
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST", db, adOpenForwardOnly
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
    
    If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.Text <> "" Then
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
    Chkcancel.value = 0
    Set creditbill = Me
    If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
        Me.lblcredit.Caption = "0"
        Me.Generateprint
    Else
        Me.Enabled = False
        FRMDEBITRT.Show
    End If
    
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    txtOutstanding.Text = ""
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
    If MDIMAIN.StatusBar.Panels(8).Text = "Y" Then
        If DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "select OPEN_DB from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                m_OP_Bal = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
                   
            Set RSTTRXFILE = New ADODB.Recordset
            'RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' AND NOT(TRX_TYPE= 'RT' AND INV_TRX_TYPE ='SV' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & ") ", db, adOpenStatic, adLockReadOnly, adCmdText
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
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
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE= 'RT' AND INV_TRX_TYPE ='SV' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
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
            
            If GRDRECEIPT.Rows > 1 Then Rcptamt = GRDRECEIPT.TextMatrix(0, 0)
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND NOT(INV_TRX_TYPE = 'SV' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & ") AND (TRX_TYPE = 'DR' OR TRX_TYPE = 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
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
            
            txtOutstanding.Text = (m_OP_Bal + m_Bal_Amt) - (m_Rcpt_Amt)
        End If
    
        'Call ReportGeneratION_vpestimate
        LBLFOT.Tag = ""
        If frmLogin.rs!Level <> "0" And NEW_BILL = True Then
            If MsgBox("You do not have any permission to modify this further. Are you sure to print?", vbYesNo, "BILL..") = vbNo Then Exit Function
        Else
            Screen.MousePointer = vbHourglass
            Sleep (300)
        End If
        NEW_BILL = False
    
        Call ReportGeneratION_vpestimate(Val(txtOutstanding.Text), Rcptamt)
    
        On Error GoTo CLOSEFILE
        Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
        If Err.Number = 55 Then
            Close #1
            Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\repo.bat" For Output As #1 '//Creating Batch file
        End If
        On Error GoTo eRRhAND
        
        Print #1, "TYPE " & MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\Report.txt > PRN"
        Print #1, "EXIT"
        Close #1
        
        '//HERE write the proper path where your command.com file exist
        Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\REPO.BAT N", vbHide
        ST_PRINT = False
        'Call cmdRefresh_Click
        cmdRefresh.SetFocus
    Else
        Call Print_A4
    End If
    Screen.MousePointer = vbNormal
    Exit Function
eRRhAND:
    ST_PRINT = False
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Function

Private Sub CMDPRINT_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            'TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            
            optnet.value = True
            TxtMRP.Text = ""
            lblactqty.Caption = ""
            lblbarcode.Caption = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            TxtCessAmt.Text = ""
            TxtCessPer.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
            End If
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
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
    
    If txtBillNo.Tag = "Y" Then
        Call AppendSale
        Exit Sub
    End If
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
    
    If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.Text <> "" Then
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
    Call AppendSale
    TxtCN.Text = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    Chkcancel.value = 0
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    cmbtype.ListIndex = 0
    TXTTYPE.Text = 1
    'Me.Enabled = False
    'FRMDEBITRT.Show
    
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            'TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            
            optnet.value = True
            TxtMRP.Text = ""
            lblactqty.Caption = ""
            lblbarcode.Caption = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            TxtCessAmt.Text = ""
            TxtCessPer.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
            End If
            
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            TXTDISC.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    TxtCode.Text = DataList2.BoundText
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub Form_Activate()
    If txtBillNo.Visible = True Then txtBillNo.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
    If txtretail.Enabled = True Then txtretail.SetFocus
    If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
    If TXTTAX.Enabled = True Then TXTTAX.SetFocus
    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
    ''If txtcommi.Enabled = True Then txtcommi.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
    If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
    'If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
    'If  Then CMDDOS.SetFocus
    
    'If TXTDEALER.Enabled = True Then TXTDEALER.SetFocus
    If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
    If TxtBillName.Enabled = True Then TxtBillName.SetFocus
    If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
    If TXTDEALER.Enabled = True Then TXTDEALER.SetFocus
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyF1
                ST_PRINT = True
                Call Generateprint
                ST_PRINT = False
        End Select
    End If
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo eRRhAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'SV'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        If RSTCOMPANY!TERMS_FLAG = "Y" Then
            chkTerms.value = 1
            Terms1.Text = IIf(IsNull(RSTCOMPANY!Terms1), "", RSTCOMPANY!Terms1)
        Else
            chkTerms.value = 0
            Terms1.Text = ""
        End If
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Terms1.Tag = Terms1.Text
    
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
    
    BR_FLAG = True
    NEW_BILL = True
    SERIAL_FLAG = False
    lblactqty.Caption = ""
    lblbarcode.Caption = ""
    ACT_FLAG = True
    AGNT_FLAG = True
    M_ADD = False
    lblcredit.Caption = "1"
    txtcrdays.Text = ""
    lblP_Rate.Caption = "0"
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 600
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 4000
    grdsales.ColWidth(3) = 900
    grdsales.ColWidth(4) = 0
    grdsales.ColWidth(5) = 1200
    grdsales.ColWidth(6) = 1300
    grdsales.ColWidth(7) = 1300
    grdsales.ColWidth(8) = 900
    grdsales.ColWidth(9) = 900
    grdsales.ColWidth(10) = 0
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 1900
    grdsales.ColWidth(13) = 0
    grdsales.ColWidth(14) = 0
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(16) = 0
    grdsales.ColWidth(17) = 0
    grdsales.ColWidth(18) = 0
    grdsales.ColWidth(19) = 0
    grdsales.ColWidth(20) = 1100
    grdsales.ColWidth(21) = 0
    grdsales.ColWidth(22) = 0
    grdsales.ColWidth(23) = 0
    grdsales.ColWidth(24) = 1000
    grdsales.ColWidth(25) = 0
    grdsales.ColWidth(26) = 0
    grdsales.ColWidth(27) = 0
    grdsales.ColWidth(28) = 0
    grdsales.ColWidth(29) = 0
    grdsales.ColWidth(30) = 0
    grdsales.ColWidth(32) = 0
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "UNIT"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "RATE"
    grdsales.TextArray(7) = "NET RATE"
    grdsales.TextArray(8) = "DISC %"
    grdsales.TextArray(9) = "TAX %"
    grdsales.TextArray(10) = "Serial No"
    grdsales.TextArray(11) = "COST"
    grdsales.TextArray(12) = "SUB TOTAL"
    grdsales.TextArray(13) = "ITEM CODE"
    grdsales.TextArray(14) = "Vch No"
    grdsales.TextArray(15) = "Line No"
    grdsales.TextArray(16) = "Trx Type"
    grdsales.TextArray(17) = "Tax Mode"
    grdsales.TextArray(18) = "MFGR"
    grdsales.TextArray(19) = "CN/DN"
    grdsales.TextArray(20) = "FREE"
    grdsales.TextArray(21) = "PTR"
    grdsales.TextArray(22) = "PTRWOTAX"
    grdsales.TextArray(24) = "Commi"
    grdsales.TextArray(31) = "Code"
    grdsales.TextArray(33) = "Print Name"
    grdsales.TextArray(34) = "Gross"
    grdsales.TextArray(38) = "Expiry"
    'grdsales.ColWidth(12) = 0
    'grdsales.ColWidth(13) = 0
    'grdsales.ColWidth(14) = 0
   'grdsales.ColWidth(15) = 0
    'grdsales.ColWidth(16) = 0
    
    grdsales.ColAlignment(0) = 4
    grdsales.ColAlignment(2) = 1
    grdsales.ColAlignment(3) = 4
    grdsales.ColAlignment(5) = 7
    grdsales.ColAlignment(7) = 7
    grdsales.ColAlignment(8) = 4
    grdsales.ColAlignment(12) = 7
    grdsales.ColAlignment(20) = 4
    
    If frmLogin.rs!Level <> "0" Then
        Label1(21).Visible = False
        lblretail.Visible = False
        grdsales.ColWidth(31) = 0
    Else
        grdsales.ColWidth(31) = 1100
        Label1(21).Visible = True
        lblretail.Visible = True
    End If
    
    LBLTOTAL.Caption = 0
    lblcomamt.Caption = 0
    LBLRETAMT.Caption = 0
    
    PHYFLAG = True
    PHYCODEFLAG = True
    TMPFLAG = True
    BATCH_FLAG = True
    ITEM_FLAG = True
    PRERATE_FLAG = True
    cr_days = False
    
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    TXTQTY.Enabled = False
    
    TxtMRP.Enabled = False
    
    txtretail.Enabled = False
    txtcommi.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    TXTDISC.Enabled = False
    txtcommi.Enabled = False
    'cmddelete.Enabled = False
    'CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    
    CmdPrintA5.Enabled = False
    
    TXTSLNO.Text = 1
    Call FILLCOMBO
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    CLOSEALL = 1
    TxtCN.Text = ""
    M_EDIT = False
    
    TXTSLNO.Text = grdsales.Rows
    txtBillNo.Visible = False
    TXTDEALER.Text = "CASH"
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    cmbtype.ListIndex = 0
    TXTTYPE.Text = 1
'    Me.Width = 11700
'    Me.Height = 10185
    Me.Left = 500
    Me.Top = 0
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
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
            txtBatch.Text = GRDPOPUP.Columns(0)
            TXTVCHNO.Text = GRDPOPUP.Columns(2)
            TXTLINENO.Text = GRDPOPUP.Columns(3)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(4)
            TxtMRP.Text = IIf(IsNull(GRDPOPUP.Columns(21)), "", GRDPOPUP.Columns(21))
            TXTEXPIRY.Text = IIf(IsDate(GRDPOPUP.Columns(25)), Format(GRDPOPUP.Columns(25), "MM/YY"), "  /  ")
            'TXTUNIT.Text = GRDPOPUP.Columns(5)
            
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = True
            
            TXTQTY.SetFocus
            
            Call CONTINUE_BATCH
            TxtWarranty.Text = GRDPOPUP.Columns(7)
            TxtWarranty_type.Text = GRDPOPUP.Columns(8)
            Set GRDPOPUP.DataSource = Nothing
            Exit Sub
        
            'TXTQTY.Text = GRDPOPUP.Columns(1)
            TxtMRP.Text = GRDPOPUP.Columns(3)
'            Select Case cmbtype.ListIndex
'                Case 0
'                    'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(20)), "", GRDPOPUP.Columns(20))
'                    'Kannattu
'                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(20)), "", GRDPOPUP.Columns(20))
'                Case 1
'                    'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(13)), "", GRDPOPUP.Columns(13))
'                    'Kannattu
'                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(13)), "", GRDPOPUP.Columns(13))
'                Case 2
'                    'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
'                    'Kannattu
'                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
'            End Select
            LblPack.Text = "1"
            'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", Val(GRDPOPUP.Columns(19)) * Val(LblPack.Text))
            Select Case cmbtype.ListIndex
                Case 0
                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(13)), "", Val(GRDPOPUP.Columns(13)))
                    txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(13)), "", Val(GRDPOPUP.Columns(13)))
                Case 1
                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", Val(GRDPOPUP.Columns(19)))
                    txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", Val(GRDPOPUP.Columns(19)))
                Case 2
                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(20)), "", Val(GRDPOPUP.Columns(20)))
                    txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(20)), "", Val(GRDPOPUP.Columns(20)))
            End Select
            If Val(TxtCessPer.Text) <> 0 Or Val(TxtCessAmt.Text) <> 0 Then
                TXTRETAILNOTAX.Text = (Val(txtretail.Text) - Val(TxtCessAmt.Text)) / (1 + (Val(TXTTAX.Text) / 100) + (Val(TxtCessPer.Text) / 100))
                txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + (Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100), 3)
                TXTRETAILNOTAX.Text = Val(txtretail.Text)
            End If
                
            lblretail.Caption = IIf(IsNull(GRDPOPUP.Columns(13)), "", GRDPOPUP.Columns(13))
            lblwsale.Caption = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
            lblvan.Caption = IIf(IsNull(GRDPOPUP.Columns(20)), "", GRDPOPUP.Columns(20))
            lblcase.Caption = IIf(IsNull(GRDPOPUP.Columns(18)), "", GRDPOPUP.Columns(18))
            lblcrtnpack.Caption = IIf(IsNull(GRDPOPUP.Columns(17)), "", GRDPOPUP.Columns(17))
            
            LblPack.Text = IIf(IsNull(GRDPOPUP.Columns(17)), "1", GRDPOPUP.Columns(17))
            If Val(LblPack.Text) = 0 Then LblPack.Text = "1"
            txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(18)), "", GRDPOPUP.Columns(18))
            
            If GRDPOPUP.Columns(14) = "A" Then
                txtretaildummy.Text = IIf(IsNull(GRDPOPUP.Columns(16)), "P", GRDPOPUP.Columns(16))
                TxtRetailmode.Text = "A"
            Else
                txtretaildummy.Text = IIf(IsNull(GRDPOPUP.Columns(15)), "P", GRDPOPUP.Columns(15))
                TxtRetailmode.Text = "P"
            End If
            
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & GRDPOPUP.Columns(6) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            With RSTtax
                If Not (.EOF And .BOF) Then
                    Select Case GRDPOPUP.Columns(12)
                        Case "M"
                            OPTTaxMRP.value = True
                            TXTTAX.Text = IIf(IsNull(RSTtax!SALES_TAX), "", RSTtax!SALES_TAX)
                            TXTSALETYPE.Text = "2"
                        Case "V"
                            If (!Category = "GENERAL" And !REMARKS = "1") Then
                                OPTTaxMRP.value = True
                                TXTSALETYPE.Text = "1"
                            Else
                                OPTVAT.value = True
                                TXTSALETYPE.Text = "2"
                            End If
                            TXTTAX.Text = IIf(IsNull(RSTtax!SALES_TAX), "", RSTtax!SALES_TAX)
                        Case Else
                            TXTSALETYPE.Text = "2"
                            optnet.value = True
                            TXTTAX.Text = "0"
                    End Select
                Else
                    optnet.value = True
                    TXTTAX.Text = "0"
                End If
            End With
            
'            OPTVAT.value = True
'            TXTTAX.Text = "14.5"
'            TXTSALETYPE.Text = "2"
'
            RSTtax.Close
            Set RSTtax = Nothing
            
            TXTVCHNO.Text = GRDPOPUP.Columns(8)
            TXTLINENO.Text = GRDPOPUP.Columns(9)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(10)
            TXTUNIT.Text = GRDPOPUP.Columns(11)
                        
            Set GRDPOPUP.DataSource = Nothing
            
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            txtBatch.Enabled = True
            
            txtBatch.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtappendcomm.Text = ""
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

Private Sub GRDPOPUPITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim NONSTOCKFLAG As Boolean
    Dim MINUSFLAG As Boolean
    Dim i As Long
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            SERIAL_FLAG = False
            lblactqty.Caption = ""
            lblbarcode.Caption = ""
            NONSTOCKFLAG = False
            MINUSFLAG = False
            M_STOCK = Val(GRDPOPUPITEM.Columns(2))
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            txtcommi.Text = ""
            TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            txtPrintname.Text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            TxtMRP.Text = IIf(IsNull(GRDPOPUPITEM.Columns(20)), "", GRDPOPUPITEM.Columns(20))
            txtcategory.Text = IIf(IsNull(GRDPOPUPITEM.Columns(7)), "", GRDPOPUPITEM.Columns(7))
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                txtretail.Enabled = True
                txtretail.SetFocus
                Exit Sub
            End If
            i = 0
            If M_STOCK <= 0 Then
                MsgBox "AVAILABLE STOCK IS  " & i & " ", , "SALES"
                TXTQTY.SelStart = 0
                TXTQTY.SelLength = Len(TXTQTY.Text)
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
            For i = 1 To grdsales.Rows - 1
                If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
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
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE, CATEGORY, WARRANTY, WARRANTY_TYPE, MRP  From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE, CATEGORY, WARRANTY, WARRANTY_TYPE, MRP  From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            End If
            Set GRDPOPUPITEM.DataSource = PHY_ITEM
            'GRDPOPUPITEM.RowHeight = 350
            If PHY_ITEM.RecordCount = 0 Then
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TXTPRODUCT.Enabled = False
                TXTITEMCODE.Enabled = False
                TXTQTY.Enabled = True
                
                TXTQTY.SetFocus
                Exit Sub
            End If
            
            Dim RSTBATCH As ADODB.Recordset
            Set RSTBATCH = New ADODB.Recordset
            Select Case cmbtype.ListIndex
                Case 0
                    'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_RETAIL >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                    RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND P_RETAIL >0 ", db, adOpenStatic, adLockReadOnly
                Case 1
                    'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND(P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                    RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ANDP_WS >0 ", db, adOpenStatic, adLockReadOnly
                Case Else
                    'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_VAN >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                    RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND P_VAN >0 ", db, adOpenStatic, adLockReadOnly
            End Select
            If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                If RSTBATCH.RecordCount > 1 Then
                    Call FILL_BATCHGRID
                    RSTBATCH.Close
                    Set RSTBATCH = Nothing
                    Exit Sub
                ElseIf RSTBATCH.RecordCount = 1 Then
                    'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                End If
            End If
            RSTBATCH.Close
            Set RSTBATCH = Nothing

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
            LblPack.Text = IIf(IsNull(GRDPOPUPITEM.Columns(15)) Or Val(GRDPOPUPITEM.Columns(15)) = 0, "1", GRDPOPUPITEM.Columns(15))
            lblOr_Pack.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(15)) Or Val(GRDPOPUPITEM.Columns(15)) = 0, "1", GRDPOPUPITEM.Columns(15))
            'txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)) * Val(LblPack.Text))
            
            Select Case cmbtype.ListIndex
                Case 0
                    txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                Case 1
                    txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)))
                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)))
                Case 2
                    txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", Val(GRDPOPUPITEM.Columns(13)))
                    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", Val(GRDPOPUPITEM.Columns(13)))
            End Select
            If Val(TxtCessPer.Text) <> 0 Or Val(TxtCessAmt.Text) <> 0 Then
                TXTRETAILNOTAX.Text = (Val(txtretail.Text) - Val(TxtCessAmt.Text)) / (1 + (Val(TXTTAX.Text) / 100) + (Val(TxtCessPer.Text) / 100))
                txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + (Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100), 3)
                TXTRETAILNOTAX.Text = Val(txtretail.Text)
            End If
                
            lblretail.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", GRDPOPUPITEM.Columns(6))
            lblwsale.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", GRDPOPUPITEM.Columns(12))
            lblvan.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", GRDPOPUPITEM.Columns(13))
            lblcase.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(11)), "", GRDPOPUPITEM.Columns(11))
            lblcrtnpack.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(10)), "", GRDPOPUPITEM.Columns(10))
            lblunit.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(16)), "Nos", GRDPOPUPITEM.Columns(16))
            TxtWarranty.Text = IIf(IsNull(GRDPOPUPITEM.Columns(18)), "", GRDPOPUPITEM.Columns(18))
            TxtWarranty_type.Text = IIf(IsNull(GRDPOPUPITEM.Columns(19)), "", GRDPOPUPITEM.Columns(19))
        
            LblPack.Text = IIf(IsNull(GRDPOPUPITEM.Columns(10)), "", GRDPOPUPITEM.Columns(10))
            If Val(LblPack.Text) = 0 Then LblPack.Text = "1"
            txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(11)), "", GRDPOPUPITEM.Columns(11))
            
            If GRDPOPUPITEM.Columns(7) = "A" Then
                txtretaildummy.Text = IIf(IsNull(GRDPOPUPITEM.Columns(9)), "P", GRDPOPUPITEM.Columns(9))
                TxtRetailmode.Text = "A"
            Else
                txtretaildummy.Text = IIf(IsNull(GRDPOPUPITEM.Columns(8)), "P", GRDPOPUPITEM.Columns(8))
                TxtRetailmode.Text = "P"
            End If
            Select Case PHY_ITEM!CHECK_FLAG
                Case "M"
                    OPTTaxMRP.value = True
                    TXTTAX.Text = GRDPOPUPITEM.Columns(4)
                    TXTSALETYPE.Text = "2"
                Case "V"
                    OPTVAT.value = True
                    TXTSALETYPE.Text = "2"
                    TXTTAX.Text = GRDPOPUPITEM.Columns(4)
                Case Else
                    TXTSALETYPE.Text = "2"
                    optnet.value = True
                    TXTTAX.Text = "0"
            End Select
            
'            OPTVAT.value = True
'            TXTTAX.Text = "14.5"
'            TXTSALETYPE.Text = "2"
            
'            optnet.Value = True
            TXTUNIT.Text = GRDPOPUPITEM.Columns(5)
                        
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            txtBatch.Enabled = True
            
            txtBatch.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
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
            On Error Resume Next
            TXTITEMCODE.Text = grdtmp.Columns(0)
            TXTPRODUCT.Text = grdtmp.Columns(1)
            txtPrintname.Text = grdtmp.Columns(1)
            TxtCessPer.Text = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
            TxtCessAmt.Text = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
            Call TxtItemcode_KeyDown(13, 0)
            
            Set grdtmp.DataSource = Nothing
            grdtmp.Visible = False
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                txtBatch.Enabled = True
                
                txtBatch.SetFocus
            End If
    End Select
End Sub

Private Sub LblPack_GotFocus()
    LblPack.SelStart = 0
    LblPack.SelLength = Len(LblPack.Text)
End Sub

Private Sub LblPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(LblPack.Text) = 0 Then Exit Sub
            LblPack.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
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
    If Val(LblPack.Text) <> Val(lblOr_Pack.Caption) Then
        'txtretail.Text = Val(lblcase.Caption) * Val(LblPack.Text)
        txtretail.Text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.Text)
    Else
        If cmbtype.ListIndex = 0 Then
            txtretail.Text = Val(lblretail.Caption)
        ElseIf cmbtype.ListIndex = 1 Then
            txtretail.Text = Val(lblwsale.Caption)
        ElseIf cmbtype.ListIndex = 0 Then
            txtretail.Text = Val(lblvan.Caption)
        End If
    End If
    Call TXTRETAIL_LostFocus
End Sub

Private Sub optnet_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OPTTaxMRP_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OPTVAT_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub TXTBATCH_GotFocus()
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
    TxtCessPer.Enabled = True
    TxtCessAmt.Enabled = True
    txtcommi.Enabled = True
    TxtWarranty.Enabled = True
    TxtWarranty_type.Enabled = True
    txtPrintname.Enabled = True
    
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
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
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/")
            KeyAscii = Asc(Chr(KeyAscii))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtBillAddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            If Trim(TxtBillName.Text) = "" Then TxtBillName.Text = TXTDEALER.Text
'                MsgBox "Enter Customer Name", vbOKOnly, "EzBiz"
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

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
    cr_days = False
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
'    MDIMAIN.MNUENTRY.Visible = False
'    MDIMAIN.MNUREPORT.Visible = False
'    MDIMAIN.mnugud_rep.Visible = False
'    MDIMAIN.MNUTOOLS.Visible = False
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TRXMAST As ADODB.Recordset
    Dim TRXSUB As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    
    Dim i As Long
    Dim N As Integer
    Dim M As Integer

    On Error GoTo eRRhAND
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click

    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtBillNo.Text) = 0 Then Exit Sub
            'If Val(txtBillNo.Text) > 800 Then Exit Sub
            lblbalance.Caption = ""
            Txtrcvd.Text = ""
            txtBillNo.Tag = "N"
'            Set TRXMAST = New ADODB.Recordset
'            TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " AND (ISNULL(BILL_FLAG) OR BILL_FLAG <>'Y') ", db, adOpenStatic, adLockReadOnly
'            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
'                txtBillNo.Tag = "Y"
'            Else
'                txtBillNo.Tag = "N"
'            End If
'            TRXMAST.Close
'            Set TRXMAST = Nothing
           
            grdsales.Rows = 1
            i = 0
            Set TRXSUB = New ADODB.Recordset
            TRXSUB.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until TRXSUB.EOF
                Set TRXFILE = New ADODB.Recordset
                TRXFILE.Open "Select * From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(TRXSUB!line_no) & "", db, adOpenStatic, adLockReadOnly
                If Not (TRXFILE.EOF And TRXFILE.BOF) Then
                    i = i + 1
                    TXTINVDATE.Text = Format(TRXFILE!VCH_DATE, "DD/MM/YYYY")
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
                        grdsales.TextMatrix(i, 18) = IIf(IsNull(TRXMAST!MANUFACTURER), "", Trim(TRXMAST!MANUFACTURER))
                    End If
                    TRXMAST.Close
                    Set TRXMAST = Nothing
                    
                    grdsales.TextMatrix(i, 5) = Format(TRXFILE!MRP, ".000")
                    grdsales.TextMatrix(i, 6) = Format(TRXFILE!PTR, ".0000")
                    grdsales.TextMatrix(i, 7) = Format(TRXFILE!SALES_PRICE, ".0000")
                    grdsales.TextMatrix(i, 8) = IIf(IsNull(TRXFILE!LINE_DISC), 0, TRXFILE!LINE_DISC) 'DISC
                    grdsales.TextMatrix(i, 9) = Val(TRXFILE!SALES_TAX)
            
                    grdsales.TextMatrix(i, 10) = IIf(IsNull(TRXFILE!REF_NO), "", TRXFILE!REF_NO) 'SERIAL
                    grdsales.TextMatrix(i, 11) = IIf(IsNull(TRXFILE!ITEM_COST), 0, TRXFILE!ITEM_COST)
                    grdsales.TextMatrix(i, 12) = Format(Val(TRXFILE!TRX_TOTAL), ".000")
                    
                    grdsales.TextMatrix(i, 13) = TRXFILE!ITEM_CODE
                    grdsales.TextMatrix(i, 14) = Val(TRXSUB!R_VCH_NO)
                    grdsales.TextMatrix(i, 15) = Val(TRXSUB!R_LINE_NO)
                    grdsales.TextMatrix(i, 16) = Trim(TRXSUB!R_TRX_TYPE)
                    grdsales.TextMatrix(i, 17) = IIf(IsNull(TRXFILE!CHECK_FLAG), "", Trim(TRXFILE!CHECK_FLAG))
                    TXTDEALER.Text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
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
                    grdsales.TextMatrix(i, 32) = TRXFILE!line_no
                    grdsales.TextMatrix(i, 33) = IIf(IsNull(TRXFILE!PRINT_NAME), grdsales.TextMatrix(i, 2), TRXFILE!PRINT_NAME)
                    grdsales.TextMatrix(i, 34) = IIf(IsNull(TRXFILE!GROSS_AMOUNT), 0, TRXFILE!GROSS_AMOUNT)
                    grdsales.TextMatrix(i, 35) = IIf(IsNull(TRXFILE!DN_NO), "", TRXFILE!DN_NO)
                    grdsales.TextMatrix(i, 36) = IIf(IsNull(TRXFILE!DN_DATE), "", Format(TRXFILE!DN_DATE, "DD/MM/YYYY"))
                    grdsales.TextMatrix(i, 37) = IIf(IsNull(TRXFILE!DN_LINENO), "", TRXFILE!DN_LINENO)
                    grdsales.TextMatrix(i, 38) = IIf(IsDate(TRXFILE!EXP_DATE), Format(TRXFILE!EXP_DATE, "MM/YY"), "")
                    grdsales.TextMatrix(i, 39) = IIf(IsNull(TRXFILE!RETAILER_PRICE), 0, TRXFILE!RETAILER_PRICE)
                    grdsales.TextMatrix(i, 40) = IIf(IsNull(TRXFILE!CESS_PER), 0, TRXFILE!CESS_PER)
                    grdsales.TextMatrix(i, 41) = IIf(IsNull(TRXFILE!CESS_AMT), 0, TRXFILE!CESS_AMT)
                    cr_days = True
                    'txtBillNo.Text = ""
                    'LBLBILLNO.Caption = ""
                    CMDPRINT.Enabled = True
                    CmdPrintA5.Enabled = True
                    
                    cmdRefresh.Enabled = True

                End If
                TRXFILE.Close
                Set TRXFILE = Nothing
                TRXSUB.MoveNext
            Loop
            TRXSUB.Close
            Set TRXSUB = Nothing
            
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                If TRXMAST!SLSM_CODE = "A" Then
                    TXTTOTALDISC.Text = IIf(IsNull(TRXMAST!DISCOUNT), "", TRXMAST!DISCOUNT)
                    OptDiscAmt.value = True
                ElseIf TRXMAST!SLSM_CODE = "P" Then
                    If IsNull(TRXMAST!VCH_AMOUNT) Or TRXMAST!VCH_AMOUNT = 0 Then
                        TXTTOTALDISC.Text = 0
                    Else
                        TXTTOTALDISC.Text = IIf(IsNull(TRXMAST!DISCOUNT), "", Round((TRXMAST!DISCOUNT * 100 / TRXMAST!VCH_AMOUNT), 2))
                    End If
                    OPTDISCPERCENT.value = True
                End If
                LBLRETAMT.Caption = IIf(IsNull(TRXMAST!ADD_AMOUNT), "", Format(TRXMAST!ADD_AMOUNT, "0.00"))
                If TRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                txtcrdays.Text = IIf(IsNull(TRXMAST!cr_days), "", TRXMAST!cr_days)
                TxtBillName.Text = IIf(IsNull(TRXMAST!BILL_NAME), "", TRXMAST!BILL_NAME)
                TxtBillAddress.Text = IIf(IsNull(TRXMAST!BILL_ADDRESS), "", TRXMAST!BILL_ADDRESS)
                TxtVehicle.Text = IIf(IsNull(TRXMAST!VEHICLE), "", TRXMAST!VEHICLE)
                TxtOrder.Text = IIf(IsNull(TRXMAST!D_ORDER), "", TRXMAST!D_ORDER)
                TxtFrieght.Text = IIf(IsNull(TRXMAST!FRIEGHT), "", TRXMAST!FRIEGHT)
                Txthandle.Text = IIf(IsNull(TRXMAST!Handle), "", TRXMAST!Handle)
                TxtPhone.Text = IIf(IsNull(TRXMAST!PHONE), "", TRXMAST!PHONE)
                TXTTIN.Text = IIf(IsNull(TRXMAST!TIN), "", TRXMAST!TIN)
                TXTAREA.Text = IIf(IsNull(TRXMAST!Area), "", TRXMAST!Area)
                TXTDEALER.Text = IIf(IsNull(TRXMAST!ACT_NAME), "", TRXMAST!ACT_NAME)
                DataList2.BoundText = IIf(IsNull(TRXMAST!ACT_CODE), "", TRXMAST!ACT_CODE)
                CMBDISTI.Text = IIf(IsNull(TRXMAST!AGENT_NAME), "", TRXMAST!AGENT_NAME)
                CMBDISTI.BoundText = IIf(IsNull(TRXMAST!AGENT_CODE), "", TRXMAST!AGENT_CODE)
                TxtCN.Text = IIf(IsNull(TRXMAST!CN_REF), "", TRXMAST!CN_REF)
                Select Case TRXMAST!BILL_TYPE
                    Case "R"
                        cmbtype.ListIndex = 0
                        TXTTYPE.Text = 1
                    Case "W"
                        cmbtype.ListIndex = 1
                        TXTTYPE.Text = 2
                    Case "V"
                        cmbtype.ListIndex = 2
                        TXTTYPE.Text = 3
                End Select
                
                GRDRECEIPT.Rows = 1
                GRDRECEIPT.TextMatrix(0, 0) = IIf(IsNull(TRXMAST!RCPT_AMOUNT), 0, TRXMAST!RCPT_AMOUNT)
                GRDRECEIPT.Rows = GRDRECEIPT.Rows + 1
                GRDRECEIPT.TextMatrix(1, 0) = IIf(IsNull(TRXMAST!RCPT_REFNO), "", TRXMAST!RCPT_REFNO)
                If TRXMAST!BANK_FLAG = "Y" Then
                    GRDRECEIPT.Rows = GRDRECEIPT.Rows + 1
                    GRDRECEIPT.TextMatrix(2, 0) = "B"
                    GRDRECEIPT.Rows = GRDRECEIPT.Rows + 1
                    GRDRECEIPT.TextMatrix(3, 0) = IIf(IsNull(TRXMAST!CHQ_NO), "", TRXMAST!CHQ_NO)
                    GRDRECEIPT.Rows = GRDRECEIPT.Rows + 1
                    GRDRECEIPT.TextMatrix(4, 0) = IIf(IsNull(TRXMAST!BANK_CODE), "", TRXMAST!BANK_CODE)
                    GRDRECEIPT.Rows = GRDRECEIPT.Rows + 1
                    If Not IsNull(TRXMAST!CHQ_DATE) Then
                        GRDRECEIPT.TextMatrix(5, 0) = IIf(IsDate(TRXMAST!CHQ_DATE), TRXMAST!CHQ_DATE, "")
                    End If
                    If TRXMAST!CHQ_STATUS = "Y" Then
                        GRDRECEIPT.Rows = GRDRECEIPT.Rows + 1
                        GRDRECEIPT.TextMatrix(6, 0) = "Y"
                    Else
                        GRDRECEIPT.Rows = GRDRECEIPT.Rows + 1
                        GRDRECEIPT.TextMatrix(6, 0) = "N"
                    End If
                    GRDRECEIPT.Rows = GRDRECEIPT.Rows + 1
                    GRDRECEIPT.TextMatrix(7, 0) = IIf(IsNull(TRXMAST!BANK_NAME), "", TRXMAST!BANK_NAME)
                Else
                    GRDRECEIPT.Rows = GRDRECEIPT.Rows + 1
                    GRDRECEIPT.TextMatrix(2, 0) = "C"
                End If
                If IsNull(TRXMAST!TERMS) Or TRXMAST!TERMS = "" Then
                    chkTerms.value = 0
                    Terms1.Text = ""
                Else
                    chkTerms.value = 1
                    Terms1.Text = TRXMAST!TERMS
                End If
                Call CMBBRNCH_GotFocus
                CMBBRNCH.Text = IIf(IsNull(TRXMAST!BR_NAME), "", TRXMAST!BR_NAME)
                CMBBRNCH.BoundText = IIf(IsNull(TRXMAST!BR_CODE), "", TRXMAST!BR_CODE)
                NEW_BILL = False
                OLD_BILL = True
            Else
                CMBBRNCH.Text = ""
                NEW_BILL = True
                OLD_BILL = False
                'TXTTYPE.Text = ""
                'cmbtype.ListIndex = -1
                cmbtype.ListIndex = 0
                TXTTYPE.Text = 1
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            If OLD_BILL = False Then
                Dim rstBILL As ADODB.Recordset
                Set rstBILL = New ADODB.Recordset
                rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'SV'", db, adOpenForwardOnly
                If Not (rstBILL.EOF And rstBILL.BOF) Then
                    cmbtype.Tag = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0))
                End If
                rstBILL.Close
                Set rstBILL = Nothing
                If Val(txtBillNo.Text) < Val(cmbtype.Tag) And Val(cmbtype.Tag) <> 0 Then
                    Set rstBILL = New ADODB.Recordset
                    rstBILL.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                    If (rstBILL.EOF And rstBILL.BOF) Then
                        rstBILL.AddNew
                        rstBILL!VCH_NO = Val(txtBillNo.Text)
                        rstBILL!TRX_TYPE = "SV"
                        rstBILL!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
                    End If
                    rstBILL.Update
                    rstBILL.Close
                    Set rstBILL = Nothing
                    OLD_BILL = True
                End If
            End If
            
            LBLBILLNO.Caption = Val(txtBillNo.Text)
            LBLTOTAL.Caption = ""
            lblnetamount.Caption = ""
            LBLFOT.Caption = ""
            lblcomamt.Caption = ""
            For i = 1 To grdsales.Rows - 1
                grdsales.TextMatrix(i, 0) = i
                Select Case grdsales.TextMatrix(i, 19)
                    Case "CN"
                        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                        LBLFOT.Caption = ""
                    Case Else
                        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                        LBLFOT.Caption = ""
                End Select
                
                If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                    lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
                Else
                    lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
                End If
            Next i
            LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
            TXTAMOUNT.Text = ""
            If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
                TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
            ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
                TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
            End If
            LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
            lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
            lblnetamount.Caption = Round(lblnetamount.Caption, 0)
            Call COSTCALCULATION
            Call Addcommission
            
            TXTSLNO.Text = grdsales.Rows
            txtBillNo.Visible = False
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If grdsales.Rows > 1 Then
                TXTDEALER.SetFocus
                'TxtName1.SetFocus
            Else
                TXTDEALER.SetFocus
                'TXTINVDATE.SetFocus
'                TxtName1.Enabled = False
'                TXTDEALER.Text = ""
'                TXTDEALER.SetFocus
            End If
            CHANGE_ADDRESS = False
            'Call Command2_Click
    End Select
    
    Exit Sub
eRRhAND:
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
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo eRRhAND
    
    txtBillNo.Tag = "N"
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " AND (ISNULL(BILL_FLAG) OR BILL_FLAG <>'Y') ", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
'        txtBillNo.Tag = "Y"
'    Else
'        txtBillNo.Tag = "N"
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
    lblbalance.Caption = ""
    Txtrcvd.Text = ""
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'SV'", db, adOpenStatic, adLockReadOnly
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
'    TRXMAST.Open "Select MIN(VCH_NO) FROM TRXFILE WHERE TRX_TYPE = 'SV'", db, adOpenStatic, adLockReadOnly
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
    Call TXTBILLNO_KeyDown(13, 0)
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.Text)
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
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
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
    On Error GoTo eRRhAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE Like '" & Me.TxtCode.Text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE Like '" & Me.TxtCode.Text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
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
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.Text)
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

Private Sub txtcrdays_GotFocus()
    txtcrdays.SelStart = 0
    txtcrdays.SelLength = Len(txtcrdays.Text)
End Sub

Private Sub txtcrdays_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            'If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            ''If txtcommi.Enabled = True Then txtcommi.SetFocus
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
    On Error GoTo eRRhAND
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
    CHANGE_ADDRESS = True
    Exit Sub
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub TXTDISC_GotFocus()
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.Text)
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtCessPer.Text) <> 0 Then
                TxtCessPer.Enabled = True
                TxtCessPer.SetFocus
            Else
                If lblsubdealer.Caption = "D" Then
                    txtcommi.Enabled = True
                    txtcommi.SetFocus
                Else
                    txtcommi.Text = 0
                    Set GRDPRERATE.DataSource = Nothing
                    fRMEPRERATE.Visible = False
                    Call CMDADD_Click
                End If
            End If
'            TXTDISC.Enabled = False
'            cmdadd.Enabled = True
'            cmdadd.SetFocus
'            'TxtWarranty.Enabled = True
'            'TxtWarranty.SetFocus
        Case vbKeyEscape
            txtretail.Enabled = True
            txtretail.SetFocus
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
    If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
        TXTDISC.Tag = Val(txtretail.Text) * Val(TXTDISC.Text) / 100
        LBLSUBTOTAL.Caption = Format(Round(Val(txtretail.Text) - Val(TXTDISC.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.Text) - Val(TXTDISC.Tag), 2), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.Text) * Val(txtretail.Text) * Val(TXTDISC.Text) / 100
        LBLSUBTOTAL.Caption = Format(Round((Val(TXTQTY.Text) * Val(txtretail.Text)) - Val(TXTDISC.Tag), 2), ".000")
        LblGross.Caption = Format(Round((Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text)) - Val(TXTDISC.Tag), 2), ".000")
        LBLSUBTOTAL.Caption = Round(Val(LBLSUBTOTAL.Caption) + (Val(TXTRETAILNOTAX.Text) - (Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100)) * Val(TXTQTY.Text) * Val(TxtCessPer) / 100 + (Val(TxtCessAmt.Text) * Val(TXTQTY.Text)), 3)
        'LBLSUBTOTAL.Caption = Round(Val(LBLSUBTOTAL.Caption) + (Round(Val(TxtCessAmt.Text) * Val(TXTQTY.Text), 3)), 3)
    End If
    
    ''TXTDISC.Text = Format(TXTDISC.Text, ".000")

End Sub

Private Sub TxtFrieght_GotFocus()
    TxtFrieght.SelStart = 0
    TxtFrieght.SelLength = Len(TxtFrieght.Text)
End Sub

Private Sub TxtFrieght_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            'If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            ''If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
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
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
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

Private Sub TXTMRP_GotFocus()
    TxtMRP.SelStart = 0
    TxtMRP.SelLength = Len(TxtMRP.Text)
End Sub

Private Sub TXTMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TxtMRP.Text) = 0 Then Exit Sub
            TXTEXPIRY.Enabled = True
            TXTEXPIRY.SetFocus
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
    TxtMRP.Text = Format(TxtMRP.Text, ".000")
End Sub

Private Sub TxtName1_Change()
    If CHANGE_NAME = False Then Exit Sub
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset

    M_STOCK = 0
    Set grdtmp.DataSource = Nothing
    If PHYFLAG = True Then
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY CATEGORY, ITEM_SLNO", db, adOpenStatic, adLockReadOnly
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE ucase(CATEGORY) = 'OWN' AND ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "') AND (CATEGORY = 'SERVICES' OR CATEGORY = 'SERVICE CHARGE' OR CATEGORY = 'SELF') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    Else
        PHY.Close
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE ucase(CATEGORY) = 'OWN' AND ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "') AND (CATEGORY = 'SERVICES' OR CATEGORY = 'SERVICE CHARGE' OR CATEGORY = 'SELF') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
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
    grdtmp.Columns(1).Width = 4600
    grdtmp.Columns(2).Caption = "QTY"
    grdtmp.Columns(2).Width = 900
    grdtmp.Columns(6).Caption = "RT"
    grdtmp.Columns(6).Width = 800
    grdtmp.Columns(4).Width = 0
    grdtmp.Columns(4).Width = 0
    grdtmp.Columns(5).Width = 0
    grdtmp.Columns(3).Width = 0
    grdtmp.Columns(7).Width = 0
    grdtmp.Columns(8).Width = 0
    grdtmp.Columns(9).Width = 0
    grdtmp.Columns(10).Width = 800
    grdtmp.Columns(10).Caption = "L/Pack"
    grdtmp.Columns(11).Caption = "LP"
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
    grdtmp.Columns(20).Caption = "MRP"
    grdtmp.Columns(20).Width = 900
    grdtmp.Columns(21).Width = 0
    grdtmp.Columns(22).Width = 2500
    grdtmp.Columns(21).Caption = "DISC"
    grdtmp.Columns(21).Width = 800
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TxtName1_GotFocus()
    CHANGE_NAME = True
    TxtName1.SelStart = 0
    TxtName1.SelLength = Len(TxtName1.Text)
    TxtName1_Change
    grdsales.Enabled = True
    'Set grdtmp.DataSource = Nothing
    'grdtmp.Visible = False
    
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
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    
End Sub

Private Sub TxtName1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtBillNo.Text) > 100 Then Exit Sub
            If txtBillNo.Tag = "Y" Then
                MsgBox "Cannot modify here", vbOKOnly, "Sales"
                Exit Sub
            End If
            If UCase(TxtName1.Text) = "OT" Then TXTITEMCODE.Text = "OT"
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            TXTITEMCODE.SetFocus
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

Private Sub txtPrintname_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape, vbKeyReturn
            If cmdadd.Enabled = True Then cmdadd.SetFocus
            'If CMDPRINT.Enabled = True Then CMDPRINT.SetFocus
            If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
            If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
            If txtBillNo.Visible = True Then txtBillNo.SetFocus
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If txtretail.Enabled = True Then txtretail.SetFocus
            If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            ''If txtcommi.Enabled = True Then txtcommi.SetFocus
    End Select
End Sub

Private Sub TXTPRODUCT_Change()
        If ITEM_CHANGE = True Then Exit Sub
        If CHANGE_NAME = False Then Exit Sub
        Dim i As Long
        Dim RSTBATCH As ADODB.Recordset
    
        M_STOCK = 0
        Set grdtmp.DataSource = Nothing
        If PHYFLAG = True Then
            'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY CATEGORY, ITEM_SLNO", db, adOpenStatic, adLockReadOnly
            'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE ucase(CATEGORY) = 'OWN'AND ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "' OR MRP Like '%" & Trim(Me.TxtName1.Text) & "') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT  From ITEMMAST  WHERE ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "') AND (CATEGORY = 'SERVICES' OR CATEGORY = 'SERVICE CHARGE' OR CATEGORY = 'SELF') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            PHYFLAG = False
        Else
            PHY.Close
            'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE ucase(CATEGORY) = 'OWN'AND ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "' OR MRP Like '%" & Trim(Me.TxtName1.Text) & "') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT  From ITEMMAST  WHERE ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "') AND (CATEGORY = 'SERVICES' OR CATEGORY = 'SERVICE CHARGE' OR CATEGORY = 'SELF') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
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
        grdtmp.Columns(1).Width = 4600
        grdtmp.Columns(2).Caption = "QTY"
        grdtmp.Columns(2).Width = 900
        grdtmp.Columns(6).Caption = "RT"
        grdtmp.Columns(6).Width = 800
        grdtmp.Columns(4).Width = 0
        grdtmp.Columns(4).Width = 0
        grdtmp.Columns(5).Width = 0
        grdtmp.Columns(3).Width = 0
        grdtmp.Columns(7).Width = 0
        grdtmp.Columns(8).Width = 0
        grdtmp.Columns(9).Width = 0
        grdtmp.Columns(10).Width = 800
        grdtmp.Columns(10).Caption = "L/Pack"
        grdtmp.Columns(11).Caption = "LP"
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
        grdtmp.Columns(20).Caption = "MRP"
        grdtmp.Columns(20).Width = 900
        grdtmp.Columns(21).Width = 0
        grdtmp.Columns(22).Width = 2500
        grdtmp.Columns(21).Caption = "DISC"
        grdtmp.Columns(21).Width = 800
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_GotFocus()
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    TXTPRODUCT.Tag = TXTPRODUCT.Text
    TXTPRODUCT.Text = ""
    TXTPRODUCT.Text = TXTPRODUCT.Tag
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
    CHANGE_NAME = True
    If Trim(TXTPRODUCT.Text) <> "" Or Trim(TxtName1.Text) <> "" Then Call TXTPRODUCT_Change
    grdsales.Enabled = True
    'grdtmp.Visible = True
    
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
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Select Case KeyCode
    
        Case vbKeyReturn
            If Trim(TxtName1.Text) = "" And Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If txtBillNo.Tag = "Y" Then
                MsgBox "Cannot modify here", vbOKOnly, "Sales"
                Exit Sub
            End If
            M_STOCK = 0
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.value = True
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            TxtCessAmt.Text = ""
            TxtCessPer.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            LblPack.Text = "1"
            lblunit.Caption = "Nos"
            txtcommi.Text = ""
            On Error Resume Next
            TXTPRODUCT.Text = grdtmp.Columns(1)
            txtPrintname.Text = grdtmp.Columns(1)
            TxtCessPer.Text = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
            TxtCessAmt.Text = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
            TxtMRP.Text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
            Select Case cmbtype.ListIndex
                Case 0
                    txtretail.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                Case 1
                    txtretail.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                Case 2
                    txtretail.Text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
            End Select
            LblPack.Text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
            lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
            'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
'            If Trim(TXTPRODUCT.Text) = "" Then
'                TxtName1.Enabled = True
'                TxtName1.SetFocus
'                Exit Sub
'            End If
            'cmddelete.Enabled = False
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
            If UCase(TxtName1.Text) = "OT" Then TXTITEMCODE.Text = "OT"
            If UCase(TXTITEMCODE.Text) <> "OT" Then
                Set grdtmp.DataSource = Nothing
                If PHYFLAG = True Then
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC,  CESS_PER, CESS_AMT  From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.Text & "' AND (CATEGORY = 'SERVICES' OR CATEGORY = 'SERVICE CHARGE' OR CATEGORY = 'SELF')", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                Else
                    PHY.Close
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC,  CESS_PER, CESS_AMT  From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.Text & "' AND (CATEGORY = 'SERVICES' OR CATEGORY = 'SERVICE CHARGE' OR CATEGORY = 'SELF')", db, adOpenStatic, adLockReadOnly
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
                    TXTITEMCODE.Text = grdtmp.Columns(0)
                    TXTPRODUCT.Text = grdtmp.Columns(1)
                    txtPrintname.Text = grdtmp.Columns(1)
                    TxtCessPer.Text = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
                    TxtCessAmt.Text = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
                    Select Case PHY!CHECK_FLAG
                        Case "M"
                            OPTTaxMRP.value = True
                            TXTTAX.Text = grdtmp.Columns(4)
                            TXTSALETYPE.Text = "2"
                        Case "V"
                            OPTVAT.value = True
                            TXTSALETYPE.Text = "2"
                            TXTTAX.Text = grdtmp.Columns(4)
                        Case Else
                            TXTSALETYPE.Text = "2"
                            optnet.value = True
                            TXTTAX.Text = "0"
                    End Select
                    Set RSTBATCH = New ADODB.Recordset
                    Select Case cmbtype.ListIndex
                        Case 0
                            'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_RETAIL >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                            RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND P_RETAIL >0 ", db, adOpenStatic, adLockReadOnly
                        Case 1
                            'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                            RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND P_WS >0 ", db, adOpenStatic, adLockReadOnly
                        Case Else
                            'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_VAN >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                            RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND P_VAN >0 ", db, adOpenStatic, adLockReadOnly
                    End Select
                    If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                        If RSTBATCH.RecordCount > 1 Then
                            Set grdtmp.DataSource = Nothing
                            grdtmp.Visible = False
                            Call FILL_BATCHGRID
                            RSTBATCH.Close
                            Set RSTBATCH = Nothing
                            Exit Sub
                        ElseIf RSTBATCH.RecordCount = 1 Then
                            'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                        End If
                    End If
                    RSTBATCH.Close
                    Set RSTBATCH = Nothing
                    
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
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                txtBatch.Enabled = True
                
                txtBatch.SetFocus
            End If
        Case vbKeyEscape
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtName1.SetFocus
            'cmddelete.Enabled = False
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
    End Select
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub
Private Function CONTINUE()
    Dim i As Long
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
                        If grdsales.TextMatrix(i, 19) <> "DN" And MsgBox("This Item Already exists in Line No. " & grdsales.TextMatrix(i, 0) & "... Do yo want to modify this item", vbYesNo + vbDefaultButton2, "BILL..") = vbYes Then
                            grdsales.Row = i
                            TXTSLNO.Text = grdsales.TextMatrix(i, 0)
                            Call CMDMODIFY_Click
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
                txtcategory.Text = IIf(IsNull(PHY!Category), "", PHY!Category)
                If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
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
                LblPack.Text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                'txtretail.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)) * Val(LblPack.Text))
                Select Case cmbtype.ListIndex
                    Case 0
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    Case 1
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                        TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    Case 2
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                        TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                End Select
                If Val(TxtCessPer.Text) <> 0 Or Val(TxtCessAmt.Text) <> 0 Then
                    TXTRETAILNOTAX.Text = (Val(txtretail.Text) - Val(TxtCessAmt.Text)) / (1 + (Val(TXTTAX.Text) / 100) + (Val(TxtCessPer.Text) / 100))
                    txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + (Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100), 3)
                    TXTRETAILNOTAX.Text = Val(txtretail.Text)
                End If
                lblretail.Caption = IIf(IsNull(grdtmp.Columns(6)), "", grdtmp.Columns(6))
                lblwsale.Caption = IIf(IsNull(grdtmp.Columns(12)), "", grdtmp.Columns(12))
                lblvan.Caption = IIf(IsNull(grdtmp.Columns(13)), "", grdtmp.Columns(13))
                lblcase.Caption = IIf(IsNull(grdtmp.Columns(11)), "", grdtmp.Columns(11))
                lblcrtnpack.Caption = IIf(IsNull(grdtmp.Columns(10)), "", grdtmp.Columns(10))
                
                
                lblunit.Caption = IIf(IsNull(grdtmp.Columns(17)), "Nos", grdtmp.Columns(17))
                TxtWarranty.Text = IIf(IsNull(grdtmp.Columns(18)), "", grdtmp.Columns(18))
                TxtWarranty_type.Text = IIf(IsNull(grdtmp.Columns(19)), "", grdtmp.Columns(19))
                TxtMRP.Text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
                TXTDISC.Text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
                'LblPack.Text = IIf(IsNull(grdtmp.Columns(10)), "", grdtmp.Columns(10))
                'If Val(LblPack.Text) = 0 Then LblPack.Text = "1"
                'txtretail.Text = IIf(IsNull(grdtmp.Columns(11)), "", grdtmp.Columns(11))
            
                If grdtmp.Columns(7) = "A" Then
                    txtretaildummy.Text = IIf(IsNull(grdtmp.Columns(9)), "P", grdtmp.Columns(9))
                    TxtRetailmode.Text = "A"
                Else
                    txtretaildummy.Text = IIf(IsNull(grdtmp.Columns(8)), "P", grdtmp.Columns(8))
                    TxtRetailmode.Text = "P"
                End If
                Select Case PHY!CHECK_FLAG
                    Case "M"
                        OPTTaxMRP.value = True
                        TXTTAX.Text = grdtmp.Columns(4)
                        TXTSALETYPE.Text = "2"
                    Case "V"
                        OPTVAT.value = True
                        TXTSALETYPE.Text = "2"
                        TXTTAX.Text = grdtmp.Columns(4)
                    Case Else
                        TXTSALETYPE.Text = "2"
                        optnet.value = True
                        TXTTAX.Text = "0"
                End Select
                
'                OPTVAT.value = True
'                TXTTAX.Text = "14.5"
'                TXTSALETYPE.Text = "2"
                
                TXTUNIT.Text = grdtmp.Columns(5)
                                   
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
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTPRODUCT_LostFocus()
    CHANGE_NAME = False
End Sub

Private Sub TXTQTY_GotFocus()
    If Val(LblPack.Text) = 0 Then LblPack.Text = 1
    If M_EDIT = False Then
        If LblPack.Text = 1 Then
            FrmeType.Visible = False
        Else
            FrmeType.Visible = True
        End If
        If Val(LblPack.Text) = Val(lblOr_Pack.Caption) Then
            OptNormal.value = True
        Else
            OptLoose.value = True
        End If
    Else
        FrmeType.Visible = False
    End If
    TxtName1.Enabled = False
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
    TXTQTY.Tag = Trim(TXTPRODUCT.Text)
    
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
    TxtCessPer.Enabled = True
    TxtCessAmt.Enabled = True
    txtcommi.Enabled = True
    TxtWarranty.Enabled = True
    TxtWarranty_type.Enabled = True
    txtPrintname.Enabled = True
    
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    TXTQTY.SetFocus
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Double
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            i = 0
            If Val(LblPack.Text) = 0 Then LblPack.Text = 1
            If Not (UCase(txtcategory.Text) = "SERVICES" Or UCase(txtcategory.Text) = "SELF") Then
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT CLOSE_QTY  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                    If (IsNull(RSTTRXFILE!CLOSE_QTY)) Then RSTTRXFILE!CLOSE_QTY = 0
                    i = RSTTRXFILE!CLOSE_QTY / Val(LblPack.Text)
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
                If Val(TXTQTY.Text) = 0 Then Exit Sub
'                If M_EDIT = False And Val(TXTQTY.Text) > i Then
'                    MsgBox "AVAILABLE STOCK IS  " & i & " ", , "SALES"
'                    TXTQTY.SelStart = 0
'                    TXTQTY.SelLength = Len(TXTQTY.Text)
'                    Exit Sub
'                End If
                'If i <> 0 Then
                    If M_EDIT = False And SERIAL_FLAG = True And Val(TXTQTY.Text) > Val(lblactqty.Caption) Then
                        MsgBox "AVAILABLE STOCK IS  " & Val(lblactqty.Caption) & " ", , "SALES"
                        TXTQTY.SelStart = 0
                        TXTQTY.SelLength = Len(TXTQTY.Text)
                        Exit Sub
                    End If
                    If M_EDIT = False And Val(TXTQTY.Text) > i Then
                        If (MsgBox("AVAILABLE STOCK IS  " & i & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                            'MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                            TXTQTY.SelStart = 0
                            TXTQTY.SelLength = Len(TXTQTY.Text)
                            Exit Sub
                        End If
                    End If
                'End If
SKIP:
                If Val(TXTTAX.Text) = 0 Then
                    TXTTAX.Enabled = True
                    TXTTAX.SetFocus
                Else
                    If MDIMAIN.StatusBar.Panels(14).Text <> "Y" Then
                        TXTRETAILNOTAX.Enabled = True
                        TXTRETAILNOTAX.SetFocus
                    Else
                        txtretail.Enabled = True
                        txtretail.SetFocus
                    End If
                End If
            Else
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
            End If
         Case vbKeyEscape
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyDown
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            cmdadd.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
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
            If M_EDIT = False Then OptNormal.value = True
            LblPack.Text = Val(lblOr_Pack.Caption)
            Call LblPack_LostFocus
            KeyAscii = 0
        Case 76, 108
            If FrmeType.Visible = False Then
                KeyAscii = 0
                Exit Sub
            End If
            If M_EDIT = False Then OptLoose.value = True
            If Val(lblcrtnpack.Caption) = 0 Then lblcrtnpack.Caption = 1
            LblPack.Text = Val(lblcrtnpack.Caption)
            Call LblPack_LostFocus
            KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    
    Dim RSTITEMCOST As ADODB.Recordset
    
    TXTQTY.Text = Format(TXTQTY.Text, ".000")
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    If Val(TXTRETAILNOTAX.Text) = 0 Then
        TXTDISC.Tag = Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(txtretail.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
        LblGross.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
        LblGross.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag), ".000")
    End If
    
    On Error GoTo eRRhAND
    Set RSTITEMCOST = New ADODB.Recordset
    RSTITEMCOST.Open "SELECT ITEM_COST, SALES_PRICE FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
        
    If Not (RSTITEMCOST.EOF Or RSTITEMCOST.BOF) Then
        LBLITEMCOST.Caption = IIf(IsNull(RSTITEMCOST!ITEM_COST), "", RSTITEMCOST!ITEM_COST * Val(LblPack.Text))
        LBLSELPRICE.Caption = IIf(IsNull(RSTITEMCOST!SALES_PRICE), "", RSTITEMCOST!SALES_PRICE * Val(LblPack.Text))
    End If
    RSTITEMCOST.Close
    Set RSTITEMCOST = Nothing
    
    Exit Sub
eRRhAND:
    MsgBox Err.Description

End Sub

Private Sub Txtrcvd_Change()
    lblbalance.Caption = Format(Round(Val(Txtrcvd.Text) - Val(lblnetamount.Caption), 2))
End Sub

Private Sub Txtrcvd_GotFocus()
    Txtrcvd.SelStart = 0
    Txtrcvd.SelLength = Len(Txtrcvd.Text)
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


Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
    Chkcancel.value = 0
    grdsales.Enabled = True
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    
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
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
'            If Trim(TXTTIN.Text) = "" Then
'                MsgBox "FORM 8B Bill Not allowed", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            'If Val(TXTSLNO.Text) < grdsales.Rows Then Exit Sub
            On Error Resume Next
            grdsales.Row = Val(TXTSLNO.Text)
            On Error GoTo eRRhAND
            If Val(TXTSLNO.Text) = 0 Then
                SERIAL_FLAG = False
                lblactqty.Caption = ""
                lblbarcode.Caption = ""
                TXTSLNO.Text = ""
                TXTPRODUCT.Text = ""
                txtPrintname.Text = ""
                TXTQTY.Text = ""
                TXTEXPIRY.Text = "  /  "
                TXTAPPENDQTY.Text = ""
                TXTFREEAPPEND.Text = ""
                txtappendcomm.Text = ""
                TXTAPPENDTOTAL.Text = ""
                TXTFREE.Text = ""
                optnet.value = True
                TxtMRP.Text = ""
                
                TXTDISC.Text = ""
                TxtCessAmt.Text = ""
                TxtCessPer.Text = ""
                LBLSUBTOTAL.Caption = ""
                LblGross.Caption = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTUNIT.Text = ""
                TXTSLNO.Text = grdsales.Rows
                'cmddelete.Enabled = False
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
                TXTFREE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 20)
                TxtMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), ".000")
                
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 15)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 16)
                TXTUNIT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                'TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 22)
                TxtRetailmode.Text = "A"
                
                Select Case grdsales.TextMatrix(Val(TXTSLNO.Text), 17)
                    Case "M"
                        OPTTaxMRP.value = True
                        TXTSALETYPE.Text = "2"
                    Case "V"
                        OPTVAT.value = True
                        TXTSALETYPE.Text = "2"
                    Case Else
                        TXTSALETYPE.Text = "2"
                        optnet.value = True
                        TXTTAX.Text = "0"
                End Select
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                txtretail.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                TXTSALETYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 23)
                txtcategory.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 25)
'                If UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICE CHARGE" Then
'                    txtretaildummy.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)), 2)
'                    'txtcommi.Text = 0 'Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) / Val(TXTQTY.Text), 2)
'                Else
'                    txtretaildummy.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)), 2)
'                    'txtcommi.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) / Val(TXTQTY.Text), 2)
'                End If
                txtretaildummy.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
                txtcommi.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
                If Not IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 38)) Then
                    TXTEXPIRY.Text = "  /  "
                Else
                    TXTEXPIRY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 38)
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
                LblPack.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 27))
                lblOr_Pack.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 27))
                TxtWarranty.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 28)
                TxtWarranty_type.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 29))
                lblunit.Caption = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 30))
                txtPrintname.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 33))
                LblGross.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 34))
                lblretail.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39))
                TxtCessPer.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40))
                TxtCessAmt.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 41))
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
            Set grdtmp.DataSource = Nothing
            grdtmp.Visible = False
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
            End If
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            If grdsales.Rows > 1 Then
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
eRRhAND:
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

Private Sub TxtTax_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
    If Val(TXTTAX.Text) = 0 Then TXTTAX.Text = ""
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TxtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If MDIMAIN.LBLTAXWARN.Caption = "Y" Then If Trim(TXTTAX.Text) = "" Then Exit Sub
            If MDIMAIN.StatusBar.Panels(14).Text <> "Y" Then
                TXTRETAILNOTAX.Enabled = True
                TXTRETAILNOTAX.SetFocus
            Else
                txtretail.Enabled = True
                txtretail.SetFocus
            End If
        Case vbKeyEscape
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
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
    If optnet.value = True And Val(TXTTAX.Text) > 0 Then
        OPTVAT.value = True
        TXTRETAILNOTAX_LostFocus
    End If
    txtmrpbt.Text = 100 * Val(TxtMRP.Text) / (100 + Val(TXTTAX.Text))
    
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
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
    On Error GoTo eRRhAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from RTRXFILE where RTRXFILE.ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
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
    
eRRhAND:
    MsgBox Err.Description
End Function


Private Function COSTCALCULATION()
    Dim COST As Double
    Dim N As Integer
    'Dim RSTITEMMAST As ADODB.Recordset
    
    LBLTOTALCOST.Caption = ""
    LBLPROFIT.Caption = ""
    COST = 0
    On Error GoTo eRRhAND
    For N = 1 To grdsales.Rows - 1
        'COST = COST + (Val(grdsales.TextMatrix(N, 11)) * Val(grdsales.TextMatrix(N, 3)))
        COST = COST + ((Val(grdsales.TextMatrix(N, 11)) + (Val(grdsales.TextMatrix(N, 11)) * Val(grdsales.TextMatrix(N, 9)) / 100)) * Val(grdsales.TextMatrix(N, 3)))
    Next N
    
    LBLTOTALCOST.Caption = Round(COST, 2)
    LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - COST, 2)

    Exit Function
    
eRRhAND:
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
    On Error GoTo eRRhAND
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Any changes made will not be saved", vbOKOnly, "Sales"
        GoTo SKIP
    End If
    If OLD_BILL = False Then Call checklastbill
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'SV' ", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        db.Execute "delete From TRNXRCPT WHERE TRX_TYPE='RT' AND CR_NO = " & RSTITEMMAST!CR_NO & " "
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    db.Execute "delete FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    'db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    'db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='DR' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TRX_TYPE = 'SV'"
    db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_VCH_NO = " & Val(LBLBILLNO.Caption) & " AND B_TRX_TYPE = 'SV' "
    db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'SV' "
    db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'SV'"
                            
    'DB.Execute "delete From P_Rate WHERE TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    If grdsales.Rows = 1 Then
        If OLD_BILL = True Then
            Dim LASTBILL As Long
            LASTBILL = 1
            Set rstBILL = New ADODB.Recordset
            rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'SV'", db, adOpenForwardOnly
            If Not (rstBILL.EOF And rstBILL.BOF) Then
                LASTBILL = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0))
            End If
            rstBILL.Close
            Set rstBILL = Nothing
            
            If LASTBILL = 1 Then GoTo SKIP
            If LASTBILL <> Val(txtBillNo.Text) - 1 Then
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND  TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    RSTTRXFILE.AddNew
                    RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
                    RSTTRXFILE!TRX_TYPE = "SV"
                    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
                End If
                RSTTRXFILE!CUST_IGST = lblIGST.Caption
                RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
                RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
                RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                RSTTRXFILE!ACT_CODE = DataList2.BoundText
                RSTTRXFILE!ACT_NAME = DataList2.Text
                RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
                RSTTRXFILE!DISC_PERS = Val(TXTTOTALDISC.Text)
                RSTTRXFILE!ADD_AMOUNT = 0
                RSTTRXFILE!ROUNDED_OFF = 0
                RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
                RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
                If Val(TxtCN.Text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.Text)
                RSTTRXFILE!BILL_FLAG = "Y"
                If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
                    RSTTRXFILE!TERMS = Trim(Terms1.Text)
                Else
                    RSTTRXFILE!TERMS = ""
                End If
                RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
                RSTTRXFILE!BR_NAME = CMBBRNCH.Text
                RSTTRXFILE.Update
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            End If
        End If
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT *  FROM QTNMAST WHERE BILL_NO = " & Val(txtBillNo.Text) & " AND BILLTYPE = 'SV' ", db, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'            RSTTRXFILE!BILL_NO = Null
'            RSTTRXFILE!billtype = Null
'            RSTTRXFILE.Update
'        End If
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
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
    If grdsales.Rows = 1 Then
        db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'SV'"
    Else
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'SV'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!rec_no = Max_No + 1
            RSTITEMMAST!INV_TYPE = "RT"
            RSTITEMMAST!INV_TRX_TYPE = "SV"
            RSTITEMMAST!INV_NO = Val(LBLBILLNO.Caption)
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        End If
        'If lblcredit.Caption <> "0" Then
        If DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
            If GRDRECEIPT.Rows <= 1 Then
                RSTITEMMAST!TRX_TYPE = "CR"
                RSTITEMMAST!AMOUNT = Val(lblnetamount.Caption)
                RSTITEMMAST!CHECK_FLAG = "C"
            Else
                RSTITEMMAST!AMOUNT = Val(GRDRECEIPT.TextMatrix(0, 0))
                RSTITEMMAST!TRX_TYPE = "CR"
                RSTITEMMAST!CHECK_FLAG = "S"
            End If
        Else
            RSTITEMMAST!AMOUNT = Val(lblnetamount.Caption)
            RSTITEMMAST!TRX_TYPE = "CR"
            RSTITEMMAST!CHECK_FLAG = "S"
        End If
        If RSTITEMMAST!CHECK_FLAG = "C" Then
            Cash_Flag = False
        Else
            Cash_Flag = True
        End If
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = Trim(DataList2.Text)
        RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
        RECNO = RSTITEMMAST!rec_no
        INVNO = RSTITEMMAST!INV_NO
        TRXTYPE = RSTITEMMAST!TRX_TYPE
        INVTRXTYPE = RSTITEMMAST!INV_TRX_TYPE
        INVTYPE = RSTITEMMAST!INV_TYPE
        
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    End If
    If Cash_Flag = False Then db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'SV'"
    
    i = 0
    If DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
        'If lblcredit.Caption <> "0" Then
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'SV' "
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
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'SV'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!TRX_TYPE = "DR"
            RSTITEMMAST!INV_TRX_TYPE = "SV"
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
            RSTITEMMAST!CR_NO = i
            RSTITEMMAST!INV_NO = Val(LBLBILLNO.Caption)
            'RSTITEMMAST!RCPT_AMT = 0
        End If
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = DataList2.Text
        RSTITEMMAST!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'        If lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0 Then
'            RSTITEMMAST!INV_AMT = Val(lblActAmt.Caption)
'        Else
'            RSTITEMMAST!INV_AMT = Val(lblnetamount.Caption)
'        End If
        RSTITEMMAST!INV_AMT = Val(lblnetamount.Caption)
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'SV' "
        Set rstMaxRec = New ADODB.Recordset
        rstMaxRec.Open "Select MAX(CR_NO) From DBTPYMT", db, adOpenForwardOnly
        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
            i = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
        End If
        rstMaxRec.Close
        Set rstMaxRec = Nothing
        
        If lblcredit.Caption = "0" Then
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'SV' ", db, adOpenStatic, adLockOptimistic, adCmdText
            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE.AddNew
                RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
                RSTTRXFILE!TRX_TYPE = "RT"
                RSTTRXFILE!INV_TRX_TYPE = "SV"
                RSTTRXFILE!INV_NO = Val(LBLBILLNO.Caption)
                RSTTRXFILE!CR_NO = i
            End If
            RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
            If GRDRECEIPT.Rows <= 1 Then
                RSTTRXFILE!RCPT_AMT = 0
                RSTTRXFILE!REF_NO = ""
            Else
                RSTTRXFILE!RCPT_AMT = Val(GRDRECEIPT.TextMatrix(0, 0))
                RSTTRXFILE!REF_NO = Trim(GRDRECEIPT.TextMatrix(1, 0))
            End If
            If GRDRECEIPT.Rows > 1 And Trim(GRDRECEIPT.TextMatrix(2, 0)) = "B" Then
                i = 1
                db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_VCH_NO = " & Val(LBLBILLNO.Caption) & " AND B_TRX_TYPE = 'SV' "
                db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'SV'"
                Set rstMaxRec = New ADODB.Recordset
                rstMaxRec.Open "Select MAX(TRX_NO) From BANK_TRX WHERE TRX_TYPE = 'CR' AND BILL_TRX_TYPE = 'RT' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenForwardOnly
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
                BANKTRX!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
                BANKTRX!BANK_CODE = GRDRECEIPT.TextMatrix(4, 0)
                BANKTRX!BANK_NAME = GRDRECEIPT.TextMatrix(7, 0)
                'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
                BANKTRX!TRX_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                BANKTRX!TRX_AMOUNT = Val(GRDRECEIPT.TextMatrix(0, 0))
                BANKTRX!ACT_CODE = DataList2.BoundText
                BANKTRX!ACT_NAME = DataList2.Text
                'RSTTRXFILE!INV_DATE = LBLDATE.Caption
                BANKTRX!REF_NO = "From " & Trim(DataList2.Text)
                BANKTRX!ENTRY_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                BANKTRX!CHQ_DATE = Format(GRDRECEIPT.TextMatrix(5, 0), "DD/MM/YYYY")
                BANKTRX!BANK_FLAG = "Y"
                If GRDRECEIPT.TextMatrix(6, 0) = "N" Then
                    BANKTRX!CHECK_FLAG = "N"
                Else
                    BANKTRX!CHECK_FLAG = "Y"
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
                RSTTRXFILE!B_TRX_YEAR = Year(MDIMAIN.DTFROM.value)
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
            RSTTRXFILE!ACT_NAME = DataList2.Text
            RSTTRXFILE!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
            RSTTRXFILE!INV_AMT = 0
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            Dim BillNO As Long
            Set rstBILL = New ADODB.Recordset
            rstBILL.Open "Select MAX(RCPT_NO) From TRNXRCPT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
            If Not (rstBILL.EOF And rstBILL.BOF) Then
                BillNO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
            End If
            rstBILL.Close
            Set rstBILL = Nothing
            
            Dim SEL_AMOUNT As Double
            
            If GRDRECEIPT.Rows <= 1 Then
                SEL_AMOUNT = 0
            Else
                SEL_AMOUNT = Val(GRDRECEIPT.TextMatrix(0, 0))
            End If
            
            If SEL_AMOUNT <= 0 Then GoTo SKIP2
            BillNO = BillNO + 1
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRNXRCPT ", db, adOpenStatic, adLockOptimistic, adCmdText
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "RT"
            RSTTRXFILE!RCPT_NO = BillNO
            RSTTRXFILE!INV_NO = Val(LBLBILLNO.Caption)
            RSTTRXFILE!INV_TRX_TYPE = "SV"
            RSTTRXFILE!INV_TRX_YEAR = Year(MDIMAIN.DTFROM.value)
            RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
            If SEL_AMOUNT > Val(lblnetamount.Caption) Then
                RSTTRXFILE!RCPT_AMOUNT = Val(lblnetamount.Caption)
            Else
                RSTTRXFILE!RCPT_AMOUNT = SEL_AMOUNT
            End If
            RSTTRXFILE!ACT_CODE = DataList2.BoundText
            RSTTRXFILE!ACT_NAME = DataList2.Text
            RSTTRXFILE!RCPT_ENTRY_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE!REF_NO = ""
            RSTTRXFILE!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
            RSTTRXFILE!CR_NO = CRNO2
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
SKIP2:
            
            Dim RCVDAMOUNT As Double
            RCVDAMOUNT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRNXRCPT WHERE INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TRX_TYPE = 'SV' AND INV_TRX_YEAR = '" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                RCVDAMOUNT = RCVDAMOUNT + IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT)
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From DBTPYMT WHERE TRX_TYPE = 'DR' AND CR_NO = " & CRNO2 & " AND INV_TRX_TYPE = 'SV' AND TRX_YEAR = '" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE!RCVD_AMOUNT = RCVDAMOUNT
                RSTTRXFILE.Update
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        End If
    Else
    '    db.Execute "delete From DBTPYMT WHERE TRX_TYPE='DR' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND INV_TRX_TYPE = 'SV'"
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'SV' "
        db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'SV' "
        db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_VCH_NO = " & Val(LBLBILLNO.Caption) & " AND B_TRX_TYPE = 'SV' "
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
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "SV"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    Else
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    End If
    
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        RSTTRXFILE!Area = RSTITEMMAST!Area
'    End If
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    RSTTRXFILE!CUST_IGST = lblIGST.Caption
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.Text
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
    RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
    RSTTRXFILE!REF_NO = ""
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "A"
        RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "P"
        RSTTRXFILE!DISCOUNT = Round(RSTTRXFILE!VCH_AMOUNT * Val(TXTTOTALDISC.Text) / 100, 2)
    End If
    RSTTRXFILE!CHECK_FLAG = "I"
    If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
    RSTTRXFILE!CFORM_NO = Time
    RSTTRXFILE!REMARKS = Left(DataList2.Text, 50)
    RSTTRXFILE!DISC_PERS = Val(TXTTOTALDISC.Text)
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!VEHICLE = Trim(TxtVehicle.Text)
    RSTTRXFILE!D_ORDER = Trim(TxtOrder.Text)
    RSTTRXFILE!PHONE = Trim(TxtPhone.Text)
    RSTTRXFILE!TIN = Trim(TXTTIN.Text)
    RSTTRXFILE!FRIEGHT = Val(TxtFrieght.Text)
    RSTTRXFILE!Handle = Val(Txthandle.Text)
    RSTTRXFILE!Area = Trim(TXTAREA.Text)
    RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!cr_days = Val(txtcrdays.Text)
    RSTTRXFILE!BILL_NAME = Trim(TxtBillName.Text)
    RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.Text)
    txtcommi.Tag = ""
    If CMBDISTI.BoundText <> "" Then
        RSTTRXFILE!AGENT_CODE = CMBDISTI.BoundText
        RSTTRXFILE!AGENT_NAME = CMBDISTI.Text
        For i = 1 To grdsales.Rows - 1
            txtcommi.Tag = Val(txtcommi.Tag) + (Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24)))
        Next i
        RSTTRXFILE!COMM_AMT = Val(txtcommi.Tag)
    Else
        RSTTRXFILE!AGENT_CODE = ""
        RSTTRXFILE!AGENT_NAME = ""
    End If
    
    Select Case cmbtype.ListIndex
        Case 0
            RSTTRXFILE!BILL_TYPE = "R"
        Case 1
            RSTTRXFILE!BILL_TYPE = "W"
        Case Else
            RSTTRXFILE!BILL_TYPE = "V"
    End Select
    If Val(TxtCN.Text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.Text)
    If GRDRECEIPT.Rows <= 1 Or DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
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
    If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
        RSTTRXFILE!TERMS = Trim(Terms1.Text)
    Else
        RSTTRXFILE!TERMS = ""
    End If
    RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
    RSTTRXFILE!BR_NAME = CMBBRNCH.Text
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
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
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'SV'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    OLD_BILL = False
    
    TXTAREA.Clear
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select DISTINCT AREA From CUSTMAST ORDER BY AREA", db, adOpenForwardOnly
    Do Until rstBILL.EOF
        If Not IsNull(rstBILL!Area) Then TXTAREA.AddItem (rstBILL!Area)
        rstBILL.MoveNext
    Loop
    rstBILL.Close
    Set rstBILL = Nothing
    
    TXTAREA.Text = ""
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    lbladdress.Caption = ""
    LBLDNORCN.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    LBLRETAMT.Caption = ""
    LBLPROFIT.Caption = ""
    LBLDATE.Caption = Date
    LBLTOTAL.Caption = ""
    lblcomamt.Caption = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    TXTAMOUNT.Text = ""
    LBLDISCAMT.Caption = ""
    lblbalance.Caption = ""
    Txtrcvd.Text = ""
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    M_EDIT = False
    NEW_BILL = True
    cmdRefresh.Enabled = False
    CMDEXIT.Enabled = True
    CMDPRINT.Enabled = False
    
    CmdPrintA5.Enabled = False
    CMDEXIT.Enabled = True
    FRMEHEAD.Enabled = True
    TXTDEALER.Enabled = True
    TxtCode.Enabled = True
    'TXTTYPE.Text = 1
    'TXTDEALER.SetFocus
    LBLITEMCOST.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Tag = ""
    TxtCode.Text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblcredit.Caption = "1"
    txtcrdays.Text = ""
    CMBDISTI.Text = ""
    TxtBillAddress.Text = ""
    TxtVehicle.Text = ""
    TxtOrder.Text = ""
    TxtFrieght.Text = ""
    Txthandle.Text = ""
    TxtBillName.Text = ""
    txtOutstanding.Text = ""
    TXTTIN.Text = ""
    lblsubdealer.Caption = ""
    lblActAmt.Caption = ""
    cr_days = False
    CHANGE_ADDRESS = False
    M_ADD = False
    TXTDEALER.Text = ""
    Terms1.Text = Terms1.Tag
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    cmbtype.ListIndex = 0
    TXTTYPE.Text = 1
    'TXTDEALER.SetFocus
    GRDRECEIPT.TextMatrix(0, 0) = 0
    GRDRECEIPT.Rows = 1
    TxtBillName.Text = ""
    CMBBRNCH.Text = ""
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    TXTDEALER.Text = "CASH"
    TXTDEALER.SetFocus
    Exit Function
eRRhAND:
    MsgBox Err.Description
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
End Sub

Private Sub TXTTOTALDISC_GotFocus()
    TXTTOTALDISC.SelStart = 0
    TXTTOTALDISC.SelLength = Len(TXTTOTALDISC.Text)
End Sub

Private Sub TXTTOTALDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            'If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            ''If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
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
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    lblnetamount.Caption = Round(lblnetamount.Caption, 0)
    LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - Val(LBLTOTALCOST.Caption), 2)
    
End Sub

Private Function ReportGeneratION()
    
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\Report.txt" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\Report.txt" For Output As #1 '//Report file Creation
    End If
    On Error GoTo eRRhAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold


    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr (106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Chr(10) & AlignLeft(RSTCOMPANY!COMP_NAME, 30)
        Print #1, AlignLeft(RSTCOMPANY!Address, 50)
        Print #1, AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, "Phone: " & RSTCOMPANY!TEL_NO & ", " & RSTCOMPANY!FAX_NO
        Print #1, "Tin: " & RSTCOMPANY!KGST
        Print #1, RepeatString("-", 80)
        'Print #1,
        '''Print #1,  "TIN No. " & RSTCOMPANY!KGST
    
        Print #1, Space(31) & "The KVAT Rules 2005"
        If Trim(TXTTIN.Text) <> "" Then
            Print #1, Space(20) & "FORM NO. 8 See rule 58(10), TAX INVOICE"
        Else
            Print #1, Space(20) & "FORM NO. 8B See rule 58(10), RETAIL INVOICE"
        End If
        Print #1, Space(32) & AlignLeft("CASH / CREDIT SALE", 25)
        Print #1, RepeatString("-", 80)
        Print #1, "D.N. NO & Date" & Space(5) & "P.O. NO. & Date" & Space(5) & "D.Doc.NO & Date" & Space(5) & "Del Terms" & Space(5) & "Veh. No"
        Print #1,
        Print #1, RepeatString("-", 80)
        'Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & AlignLeft("INVOICE FORM 8H", 16)
    
        'If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, LBLDATE.Caption)
        Print #1, "Bill No. " & Trim(LBLBILLNO.Caption) & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 67) '& Space(2) & LBLTIME.Caption
        Print #1, "TO: " & TxtBillName.Text
        If Trim(TxtBillAddress.Text) <> "" Then Print #1, TxtBillAddress.Text
        If Trim(TxtPhone.Text) <> "" Then Print #1, "Phone: " & TxtPhone.Text
        If Trim(TXTTIN.Text) <> "" Then Print #1, "TIN: " & TXTTIN.Text
        'LBLDATE.Caption = Date
    
       ' Print #1, Chr(27) & Chr(72) &  "Salesman: CS"
    
        Print #1, RepeatString("-", 80)
        Print #1, AlignLeft("Description", 22) & _
                AlignLeft("Comm Code", 9) & Space(1) & _
                AlignLeft("Qty", 4) & Space(1) & _
                AlignLeft("Rate", 7) & Space(1) & _
                AlignLeft("Tax", 5) & Space(1) & _
                AlignLeft("Tax Amt", 7) & Space(1) & _
                AlignLeft("Net Rate", 10) & Space(3) & _
                AlignLeft("Amount", 12) '& _
                Chr (27) & Chr(72) '//Bold Ends
    
        Print #1, RepeatString("-", 80)
    
        For i = 1 To grdsales.Rows - 1
            Print #1, AlignLeft(grdsales.TextMatrix(i, 2), 22) & Space(9) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 4) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)), 2), "0.00"), 7) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 9)), 2), "0.00"), 7) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 9)) / 100, 2), "0.00"), 7) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 10) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 12) '& _
                Chr (27) & Chr(72) '//Bold Ends
        Next i
    
        Print #1, AlignRight("-------------", 80)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
            Print #1, AlignRight("DISC AMOUNT ", 65) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
        End If
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 57) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, AlignRight("Round off ", 65) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
        Print #1, Chr(13)
        Print #1, AlignRight("NET AMOUNT ", 65) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 12)
        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
        Num = CCur(Round(LBLTOTAL.Caption, 0))
        Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & ")", 80)
        Print #1, RepeatString("-", 80)
        'Print #1, Chr(27) & Chr(71) & Chr(0)
        If Trim(TXTTIN.Text) <> "" Then
            Print #1, "Certified that all the particulars shown in the above Tax Invoice are true and correct"
            Print #1, "and that my/our Registration under KVAT ACT 2003 is valid as on the date of this bill"
            Print #1, RepeatString("-", 80)
        End If
        Print #1, "Thank You... E.&.O.E SUBJECT TO CHERHALA JURISDICTION"
        Print #1, "For ELECTROCRAFTS"
        'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** THANK YOU ****", 40)
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

    Close #1 '//Closing the file
    Exit Function

eRRhAND:
    MsgBox Err.Description
End Function

Private Sub TXTRETAIL_GotFocus()
'    If M_EDIT = False Then
'        If Val(LBLITEMCOST.Caption) <> 0 Then txtretail.Text = Round(Val(LBLITEMCOST.Caption) + (Val(LBLITEMCOST.Caption) * 10 / 100), 3)
'    End If
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    txtretail.SelStart = 0
    txtretail.SelLength = Len(txtretail.Text)
    If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
    TxtName1.Enabled = False
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) <> 0 And Val(txtretail.Text) = 0 Then Exit Sub
            If Val(TXTQTY.Text) = 0 And Val(TXTFREE.Text) <> 0 And Val(txtretail.Text) <> 0 Then
                MsgBox "The Item is issued as free", vbOKOnly, "Sales"
                txtretail.SetFocus
                Exit Sub
            End If
'            If Val(TXTTAX.Text) = 0 Then
'                MsgBox "Please enter the Tax", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            TXTDISC.Enabled = True
            If MDIMAIN.StatusBar.Panels(16).Text = "Y" Then
                cmdadd.Enabled = True
                cmdadd.SetFocus
            Else
                TXTDISC.SetFocus
            End If
        Case vbKeyEscape
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                If M_EDIT = True Then Exit Sub
                TxtName1.Enabled = True
                TXTPRODUCT.Enabled = True
                TXTITEMCODE.Enabled = True
                TXTPRODUCT.SetFocus
            Else
                TXTRETAILNOTAX.Enabled = True
                TXTRETAILNOTAX.SetFocus
            End If
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
    TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.0000")
    ''If lblP_Rate.Caption = "0" Then
    If Val(TXTRETAILNOTAX.Text) <> 0 Then
        If OPTTaxMRP.value = True Then
            txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + Val(txtmrpbt.Text) * Val(TXTTAX.Text) / 100, 4)
        End If
        If OPTVAT.value = True Then
            txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100, 4)
        End If
        If optnet.value = True Then
            txtretail.Text = TXTRETAILNOTAX.Text
        End If
        TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.0000")
        If TxtRetailmode.Text = "A" Then
            txtcommi.Text = Format(Round(Val(txtretaildummy.Text), 2), "0.00")
        Else
            txtcommi.Text = Format(Round((Val(TXTRETAILNOTAX.Text) * Val(txtretaildummy.Text) / 100), 2), "0.00")
        End If
    End If
End Sub

Private Sub TXTRETAILNOTAX_GotFocus()
    TXTRETAILNOTAX.SelStart = 0
    TXTRETAILNOTAX.SelLength = Len(TXTRETAILNOTAX.Text)
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
    TxtCessPer.Enabled = True
    TxtCessAmt.Enabled = True
    txtcommi.Enabled = True
    TxtWarranty.Enabled = True
    TxtWarranty_type.Enabled = True
    txtPrintname.Enabled = True
    
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
    If OPTVAT.value = False Then TXTTAX.Text = 0
    TXTRETAILNOTAX.Text = Round(Val(txtretail.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
    TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.0000")
    txtretail.Text = Format(Val(txtretail.Text), "0.0000")
    
    If Val(LBLITEMCOST.Caption) <> 0 Then
        LblProfitPerc.Caption = Round(((Val(TXTRETAILNOTAX.Text) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption), 2)
        LblProfitPerc.Caption = Format(Val(LblProfitPerc.Caption), "0.00")
    End If
    
    LblProfitAmt.Caption = Round((Val(TXTRETAILNOTAX.Text) - Val(LBLITEMCOST.Caption)) * Val(TXTQTY.Text), 2)
    LblProfitAmt.Caption = Format(Val(LblProfitAmt.Caption), "0.00")
    
    If TxtRetailmode.Text = "A" Then
        txtcommi.Text = Format(Round(Val(txtretaildummy.Text), 2), "0.00")
    Else
        txtcommi.Text = Format(Round((Val(TXTRETAILNOTAX.Text) * Val(txtretaildummy.Text) / 100), 2), "0.00")
    End If
    'TXTDISC.Tag = 0
    'TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
    'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag), ".000")
End Sub

Private Sub TxtBillName_GotFocus()
    TxtBillName.SelStart = 0
    TxtBillName.SelLength = Len(TxtBillName.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TxtBillName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Trim(TxtBillName.Text) = "" Then TxtBillName.Text = TXTDEALER.Text
            If Trim(TxtBillName.Text) = "WS" Then TxtBillName.Text = "Cash"
            TxtBillAddress.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select

End Sub

Private Sub TxtBillName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtBillAddress_GotFocus()
    TxtBillAddress.SelStart = 0
    TxtBillAddress.SelLength = Len(TxtBillAddress.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Function FILLCOMBO()
    On Error GoTo eRRhAND
    
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

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Sub CMBDISTI_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If CMBDISTI.Text = "" Then Exit Sub
            If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.Text <> "" Then
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
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
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
    If Val(txtcommi.Text) = 0 Then txtcommi.Text = ""
    txtcommi.SelStart = 0
    txtcommi.SelLength = Len(txtcommi.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub txtcommi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If txtcommi.Text = "" Then Exit Sub
            If Val(txtcommi.Text) > Val(txtretail.Text) Then
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
            If MDIMAIN.StatusBar.Panels(16).Text = "Y" Then
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
    txtcommi.Text = Format(txtcommi.Text, ".000")
End Sub

Private Sub TXTAREA_GotFocus()
    TXTAREA.SelStart = 0
    TXTAREA.SelLength = Len(TXTAREA.Text)
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
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Function FILL_PREVIIOUSRATE()
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILE  WHERE TRX_TYPE ='SV' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILE  WHERE TRX_TYPE ='SV' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
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
        LBLHEAD(2).Caption = GRDPRERATE.Columns(7).Text
    End If
End Function

Private Sub TxtItemcode_GotFocus()
    LBLITEMCOST.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    grdsales.Enabled = True
    
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
    TxtCessPer.Enabled = False
    TxtCessAmt.Enabled = False
    txtcommi.Enabled = False
    TxtWarranty.Enabled = False
    TxtWarranty_type.Enabled = False
    txtPrintname.Enabled = False
    
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            
            If txtBillNo.Tag = "Y" Then
                MsgBox "Cannot modify here", vbOKOnly, "Sales"
                Exit Sub
            End If
            M_STOCK = 0
            'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If Trim(TXTITEMCODE.Text) = "" Then
                TxtName1.SetFocus
                Exit Sub
            End If
            'cmddelete.Enabled = False
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.value = True
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            TxtCessAmt.Text = ""
            TxtCessPer.Text = ""
            txtcommi.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
            
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT  From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.Text & "' AND (CATEGORY = 'SERVICES' OR CATEGORY = 'SERVICE CHARGE' OR CATEGORY = 'SELF')", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT  From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.Text & "' AND (CATEGORY = 'SERVICES' OR CATEGORY = 'SERVICE CHARGE' OR CATEGORY = 'SELF')", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount > 0 Then
                TxtCessPer.Text = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
                TxtCessAmt.Text = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
            End If
            
            If PHY.RecordCount = 0 Then
                Set grdtmp.DataSource = Nothing
                If PHYFLAG = True Then
                    'PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_TAX, LINE_DISC, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, ITEM_SIZE, ITEM_COLOR, BARCODE, REF_NO, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, SALES_PRICE, P_WS, CRTN_PACK, P_CRTN, MRP, TAX_MODE, EXP_DATE  From RTRXFILE  WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_TAX, LINE_DISC, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, BARCODE, REF_NO, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, SALES_PRICE, P_WS, P_VAN, CRTN_PACK, P_CRTN, MRP, TAX_MODE, EXP_DATE  From RTRXFILE  WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                Else
                    PHY.Close
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_TAX, LINE_DISC, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, BARCODE, REF_NO, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, SALES_PRICE, P_WS, P_VAN, CRTN_PACK, P_CRTN, MRP, TAX_MODE, EXP_DATE  From RTRXFILE  WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                End If
                Set grdtmp.DataSource = PHY
                If PHY.RecordCount = 0 Then
                    MsgBox "Item not exists", vbOKOnly, "Sales"
                    Exit Sub
                End If
                If Trim(grdtmp.Columns(27)) <> "" Then
                    If DateDiff("d", Date, grdtmp.Columns(28)) < 0 Then
                        MsgBox "Item Expired....", vbOKOnly, "BILL.."
                        Exit Sub
                    End If
                    If DateDiff("d", Date, grdtmp.Columns(27)) < 60 Then
                        If (MsgBox("Expiry < " & Val(DateDiff("d", Date, grdtmp.Columns(29))) & "Days", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then Exit Sub
                    End If
                End If
                
                SERIAL_FLAG = True
                lblactqty.Caption = grdtmp.Columns(2)
                lblbarcode.Caption = grdtmp.Columns(15)
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTEXPIRY.Text = IIf(IsDate(grdtmp.Columns(28)), Format(grdtmp.Columns(28), "MM/YY"), "  /  ")
                Set RSTBATCH = New ADODB.Recordset
                Select Case cmbtype.ListIndex
                    Case 1
                        RSTBATCH.Open "Select DISTINCT BARCODE, P_WS From RTRXFILE WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    Case 2
                        RSTBATCH.Open "Select DISTINCT BARCODE, P_VAN From RTRXFILE WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    Case Else
                        RSTBATCH.Open "Select DISTINCT BARCODE, P_RETAIL From RTRXFILE WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                End Select
                If RSTBATCH.RecordCount > 1 Then
                    Call FILL_BATCHGRID2
                    RSTBATCH.Close
                    Set RSTBATCH = Nothing
                    Exit Sub
                End If
                RSTBATCH.Close
                Set RSTBATCH = Nothing
                TXTITEMCODE.Text = grdtmp.Columns(0)
                ITEM_CHANGE = True
                TXTPRODUCT.Text = grdtmp.Columns(1)
                ITEM_CHANGE = False
                TXTUNIT.Text = "1" 'grdtmp.Columns(4)
                TxtMRP.Text = IIf(IsNull(grdtmp.Columns(26)), "", grdtmp.Columns(26))
                If grdtmp.Columns(6) = "A" Then
                    txtretaildummy.Text = IIf(IsNull(grdtmp.Columns(8)), "", grdtmp.Columns(8))
                    TxtRetailmode.Text = "A"
                Else
                    txtretaildummy.Text = IIf(IsNull(grdtmp.Columns(7)), "", grdtmp.Columns(7))
                    TxtRetailmode.Text = "P"
                End If
                TXTEXPIRY.Text = IIf(IsDate(grdtmp.Columns(22)), Format(grdtmp.Columns(22), "MM/YY"), "  /  ")
                lblunit.Caption = grdtmp.Columns(12)
                TxtWarranty.Text = grdtmp.Columns(13)
                TxtWarranty_type.Text = grdtmp.Columns(14)
                'txtbarcode.Text = grdtmp.Columns(15)
                txtBatch.Text = grdtmp.Columns(16)
                TXTVCHNO.Text = grdtmp.Columns(17)
                TXTLINENO.Text = grdtmp.Columns(18)
                TXTTRXTYPE.Text = grdtmp.Columns(19)
                LBLITEMCOST.Caption = grdtmp.Columns(20)
                LblPack.Text = IIf(IsNull(grdtmp.Columns(11)) Or Val(grdtmp.Columns(11)) = 0, "1", grdtmp.Columns(11))
                lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(11)) Or Val(grdtmp.Columns(11)) = 0, "1", grdtmp.Columns(11))
                Select Case cmbtype.ListIndex
                    Case 1
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(22)), "", Val(grdtmp.Columns(22)))
                    Case 2
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(23)), "", Val(grdtmp.Columns(23)))
                    Case Else
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                End Select
                LBLSELPRICE.Caption = Val(txtretail.Text)
                lblretail.Caption = IIf(IsNull(grdtmp.Columns(5)), "", grdtmp.Columns(5))
                lblwsale.Caption = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
                lblvan.Caption = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
                lblcase.Caption = IIf(IsNull(grdtmp.Columns(25)), "", grdtmp.Columns(25))
                lblcrtnpack.Caption = IIf(IsNull(grdtmp.Columns(24)), "", grdtmp.Columns(24))
                
                Dim RSTtax As ADODB.Recordset
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                With RSTtax
                    If Not (.EOF And .BOF) Then
                        Select Case grdtmp.Columns(9)
                            Case "M"
                                OPTTaxMRP.value = True
                                TXTTAX.Text = grdtmp.Columns(3)
                                TXTSALETYPE.Text = "2"
                            Case "V"
                                If (!Category = "GENERAL" And !REMARKS = "1") Then
                                    OPTTaxMRP.value = True
                                    TXTSALETYPE.Text = "1"
                                Else
                                    OPTVAT.value = True
                                    TXTSALETYPE.Text = "2"
                                End If
                                TXTTAX.Text = grdtmp.Columns(3)
                            Case Else
                                TXTSALETYPE.Text = "2"
                                optnet.value = True
                                TXTTAX.Text = "0"
                        End Select
                    Else
                        optnet.value = True
                        TXTTAX.Text = "0"
                    End If
                End With
                RSTtax.Close
                Set RSTtax = Nothing
                TXTITEMCODE.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            End If
            SERIAL_FLAG = False
            lblactqty.Caption = ""
            lblbarcode.Caption = ""
            If PHY.RecordCount = 1 Then
                TxtMRP.Text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
                Select Case cmbtype.ListIndex
                    Case 0
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    Case 1
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                        TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    Case 2
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                        TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                End Select
                LblPack.Text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                TXTDISC.Text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
                'TXTEXPIRY.Text = IIf(isdate(grdtmp.Columns(22)),Format(grdtmp.Columns(22), "MM/YY"),"  /  ")
                TXTITEMCODE.Text = grdtmp.Columns(0)
                ITEM_CHANGE = True
                TXTPRODUCT.Text = grdtmp.Columns(1)
                ITEM_CHANGE = False
                txtPrintname.Text = grdtmp.Columns(1)
                Select Case PHY!CHECK_FLAG
                    Case "M"
                        OPTTaxMRP.value = True
                        TXTTAX.Text = grdtmp.Columns(4)
                        TXTSALETYPE.Text = "2"
                    Case "V"
                        OPTVAT.value = True
                        TXTSALETYPE.Text = "2"
                        TXTTAX.Text = grdtmp.Columns(4)
                    Case Else
                        TXTSALETYPE.Text = "2"
                        optnet.value = True
                        TXTTAX.Text = "0"
                End Select
                Set RSTBATCH = New ADODB.Recordset
                Select Case cmbtype.ListIndex
                    Case 0
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_RETAIL >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND P_RETAIL >0 ", db, adOpenStatic, adLockReadOnly
                    Case 1
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND P_WS >0 ", db, adOpenStatic, adLockReadOnly
                    Case Else
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_VAN >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND P_VAN >0 ", db, adOpenStatic, adLockReadOnly
                End Select
                If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                    If RSTBATCH.RecordCount > 1 Then
                        Call FILL_BATCHGRID
                        RSTBATCH.Close
                        Set RSTBATCH = Nothing
                        Exit Sub
                    ElseIf RSTBATCH.RecordCount = 1 Then
                        'TXTEXPIRY.Text = IIf(IsDate(RSTBATCH!EXP_DATE), Format(RSTBATCH!EXP_DATE, "MM/YY"), "  /  ")
                    End If
                End If
                RSTBATCH.Close
                Set RSTBATCH = Nothing
                Call CONTINUE
            Else
                Call FILL_ITEMGRID
                Exit Sub
            End If
JUMPNONSTOCK:
            TxtName1.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                txtBatch.Enabled = True
                txtBatch.SetFocus
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
            If grdsales.Rows > 1 Then
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
eRRhAND:
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

Function FILL_BATCHGRID()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
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
    GRDPOPUP.Columns(25).Caption = "Expiry"
    
    GRDPOPUP.Columns(0).Width = 4100
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
    
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).Text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
    
    
End Function

Function FILL_PREVIIOUSRATE2()
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        'PHY_PRERATE.Open "Select TOP 10 ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILE  WHERE TRX_TYPE ='HI' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, LINE_DISC, VCH_NO, ITEM_NAME, M_USER_ID  From TRXFILE WHERE TRX_TYPE ='HI' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC LIMIT 10", db, adOpenForwardOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        'PHY_PRERATE.Open "Select TOP 10 ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILE  WHERE TRX_TYPE ='HI' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, LINE_DISC, VCH_NO, ITEM_NAME, M_USER_ID  From TRXFILE WHERE TRX_TYPE ='HI' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC LIMIT 10", db, adOpenForwardOnly
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
        LBLHEAD(2).Caption = GRDPRERATE.Columns(7).Text
    End If
End Function

Private Sub TxtPhone_GotFocus()
    TxtPhone.SelStart = 0
    TxtPhone.SelLength = Len(TxtPhone.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TxtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
'            CMBDISTI.SetFocus
            'FRMEHEAD.Enabled = False
            TXTITEMCODE.Enabled = True
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                TXTPRODUCT.SetFocus
            Else
                TXTITEMCODE.SetFocus
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

Private Sub TXTTYPE_LostFocus()
    If cmbtype.ListIndex = -1 Then
        'MsgBox "Select Bill Type from the List", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
     If Val(TXTTYPE.Text) = 0 Or Val(TXTTYPE.Text) > 3 Then
        MsgBox "Enter Bill Type", vbOKOnly, "Sales"
        TXTTYPE.Enabled = True
        TXTTYPE.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 0 And Val(TXTTYPE.Text) <> 1 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 1 And Val(TXTTYPE.Text) <> 2 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 2 And Val(TXTTYPE.Text) <> 3 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
End Sub

Private Sub TxtVehicle_GotFocus()
    'If Trim(TxtVehicle.Text) = "" Then TxtVehicle.Text = "KL-04-N-8931"
    TxtVehicle.SelStart = 0
    TxtVehicle.SelLength = Len(TxtVehicle.Text)
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
                If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
                    TXTPRODUCT.SetFocus
                Else
                    TXTITEMCODE.SetFocus
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

Private Sub TxtWarranty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.Text) = 0 Then
                cmdadd.Enabled = True
                cmdadd.SetFocus
            Else
                TxtWarranty_type.SetFocus
            End If
        Case vbKeyEscape
            If MDIMAIN.StatusBar.Panels(16).Text = "Y" Then
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

Private Sub TxtWarranty_type_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.Text) <> 0 And Trim(TxtWarranty_type.Text) = "" Then
                MsgBox "Please enter Period for Warranty", , "Sales"
                TxtWarranty_type.SetFocus
                Exit Sub
            End If
            If Val(TxtWarranty.Text) = 0 Then TxtWarranty_type.Text = ""
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
    On Error GoTo eRRhAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'SV'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    LBLBILLNO.Caption = Val(txtBillNo.Text)
    
Exit Function
eRRhAND:
    MsgBox Err.Description
End Function

Private Function ReportGeneratION_estimate()
    
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\Report.txt" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\Report.txt" For Output As #1 '//Report file Creation
    End If
    On Error GoTo eRRhAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold


    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr (106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
        Print #1, AlignLeft("ESTIMATE", 25)
        Print #1, RepeatString("-", 80)
        Print #1, AlignLeft("Sl", 2) & Space(1) & _
                AlignLeft("Comm Code", 14) & Space(1) & _
                AlignLeft("Description", 35) & _
                AlignLeft("Qty", 4) & Space(3) & _
                AlignLeft("Rate", 10) & Space(3) & _
                AlignLeft("Amount", 12) '& _
                Chr (27) & Chr(72) '//Bold Ends
    
        Print #1, RepeatString("-", 80)
    
        For i = 1 To grdsales.Rows - 1
            Print #1, AlignLeft(Val(i), 3) & _
                Space(15) & AlignLeft(grdsales.TextMatrix(i, 2), 34) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 4) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 9) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 13) '& _
                Chr (27) & Chr(72) '//Bold Ends
        Next i
    
        Print #1, AlignRight("-------------", 80)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
            Print #1, AlignRight("DISC AMOUNT ", 65) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
        End If
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 57) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, AlignRight("Round off ", 65) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
        Print #1, Chr(13)
        Print #1, AlignRight("NET AMOUNT ", 65) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 12)
        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
        Num = CCur(Round(LBLTOTAL.Caption, 0))
        Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & ")", 80)
        Print #1, RepeatString("-", 80)
        'Print #1, Chr(27) & Chr(71) & Chr(0)
'        If Trim(TXTTIN.Text) <> "" Then
'            Print #1, "Certified that all the particulars shown in the above Tax Invoice are true and correct"
'            Print #1, "and that my/our Registration under KVAT ACT 2003 is valid as on the date of this bill"
'            Print #1, RepeatString("-", 80)
'        End If
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

eRRhAND:
    MsgBox Err.Description
End Function

Private Function ReportGeneratION_vpestimate(Op_Bal As Double, RCPT_AMT As Double)
    
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\Report.txt" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\Report.txt" For Output As #1 '//Report file Creation
    End If
    On Error GoTo eRRhAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(42) & chr(1) - for Enlarge letter and bold


    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(55) & Chr(27) & Chr(55)
    Print #1,
    Print #1,
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Chr(10) & AlignLeft(RSTCOMPANY!COMP_NAME, 30)
        Print #1, AlignLeft(RSTCOMPANY!Address, 50)
        Print #1, AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, "Phone: " & RSTCOMPANY!TEL_NO & ", " & RSTCOMPANY!FAX_NO
        Print #1, "Tin: " & RSTCOMPANY!KGST
        Print #1, RepeatString("-", 80)
        'Print #1,
        '''Print #1,  "TIN No. " & RSTCOMPANY!KGST
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Print #1, Space(31) & "The KVAT Rules 2005"
    If Trim(TXTTIN.Text) <> "" Then
        Print #1, Space(20) & "FORM NO. 8 See rule 58(10), TAX INVOICE"
    Else
        Print #1, Space(20) & "FORM NO. 8B See rule 58(10), RETAIL INVOICE"
    End If
    If lblcredit.Caption = 0 Then
        Print #1, Space(32) & AlignLeft("CASH SALE", 30)
    Else
        Print #1, Space(32) & AlignLeft("CREDIT SALE", 30)
    End If
    'Print #1, RepeatString("-", 80)
    Print #1, "D.N. NO & Date" & Space(5) & "P.O. NO. & Date" & Space(5) & "D.Doc.NO & Date" & Space(5) & "Del Terms" & Space(5) & "Veh. No"
    Print #1,
    Print #1, RepeatString("-", 80)
    'Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & AlignLeft("INVOICE FORM 8H", 16)

    'If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, LBLDATE.Caption)
    Print #1, "Bill No. " & Trim(LBLBILLNO.Caption) & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 67) '& Space(2) & LBLTIME.Caption
    Print #1, "TO: " & TxtBillName.Text
    If Trim(TxtBillAddress.Text) <> "" Then Print #1, TxtBillAddress.Text
    If Trim(TxtPhone.Text) <> "" Then Print #1, "Phone: " & TxtPhone.Text
    If Trim(TXTTIN.Text) <> "" Then Print #1, "TIN: " & TXTTIN.Text
    'LBLDATE.Caption = Date

   ' Print #1, Chr(27) & Chr(72) &  "Salesman: CS"

    Print #1, RepeatString("-", 80)
    Print #1, AlignLeft("Description", 22) & _
            AlignLeft("Code", 7) & Space(1) & _
            AlignLeft("Qty", 6) & Space(1) & _
            AlignLeft("Rate", 7) & Space(1) & _
            AlignLeft("Tax", 5) & Space(1) & _
            AlignLeft("Tax Amt", 7) & Space(1) & _
            AlignLeft("Net Rate", 10) & Space(3) & _
            AlignLeft("Amount", 12) '& _
            Chr (27) & Chr(72) '//Bold Ends

    Print #1, RepeatString("-", 80)

    For i = 1 To grdsales.Rows - 1
        Print #1, AlignLeft(grdsales.TextMatrix(i, 2), 22) & Space(7) & _
            AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 6) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)), 2), "0.00"), 7) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 9)), 2), "0.00"), 7) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 9)) / 100, 2), "0.00"), 7) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 10) & _
            AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 12) '& _
            Chr (27) & Chr(72) '//Bold Ends
    Next i

    Print #1, AlignRight("-------------", 80)
    If Val(LBLDISCAMT.Caption) <> 0 Then
        Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
        Print #1, AlignRight("DISC AMOUNT ", 65) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
    ElseIf Val(LBLDISCAMT.Caption) = 0 Then
        Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
    End If
    'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 57) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
    Print #1, AlignRight("Round off ", 65) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
    Print #1, Chr(13)
    Print #1, AlignRight("NET AMOUNT ", 65) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 12)
    'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
    Num = CCur(Round(LBLTOTAL.Caption, 0))
    Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & ")", 80)
    Print #1, RepeatString("-", 80)
    'Print #1, Chr(27) & Chr(71) & Chr(0)
    If Trim(TXTTIN.Text) <> "" Then
        Print #1, "Certified that all the particulars shown in the above Tax Invoice are true and correct"
        Print #1, "and that my/our Registration under KVAT ACT 2003 is valid as on the date of this bill"
        Print #1, RepeatString("-", 80)
    End If
    Print #1, "Thank You... E.&.O.E SUBJECT TO ALAPPUZHA JURISDICTION"
    Print #1, "For Alleppey Marine Metals"
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
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    
    Close #1 '//Closing the file
    Exit Function

eRRhAND:
    MsgBox Err.Description
End Function

Private Function CONTINUE_BATCH()

    M_STOCK = Val(GRDPOPUP.Columns(1))
    
    If M_STOCK <= 0 Then
        MsgBox "AVAILABLE STOCK IS  " & M_STOCK & " ", , "SALES"
        Exit Function
    End If
            
    Dim i As Double
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
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
                txtcategory.Text = IIf(IsNull(PHY!Category), "", PHY!Category)
                If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                    txtretail.Enabled = True
                    txtretail.SetFocus
                    Exit Function
                End If
                Select Case cmbtype.ListIndex
                    Case 0
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(20)), "", GRDPOPUP.Columns(20))
                        'Kannattu
                        txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                        TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                    Case 1
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(13)), "", GRDPOPUP.Columns(13))
                        'Kannattu
                        txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                        TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                    Case 2
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
                        'Kannattu
                        txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
                        TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
                End Select
                If Val(TxtCessPer.Text) <> 0 Or Val(TxtCessAmt.Text) <> 0 Then
                    TXTRETAILNOTAX.Text = (Val(txtretail.Text) - Val(TxtCessAmt.Text)) / (1 + (Val(TXTTAX.Text) / 100) + (Val(TxtCessPer.Text) / 100))
                    txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + (Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100), 3)
                    TXTRETAILNOTAX.Text = Val(txtretail.Text)
                End If
            
                'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                lblretail.Caption = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                lblwsale.Caption = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                lblvan.Caption = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
                lblcase.Caption = IIf(IsNull(GRDPOPUP.Columns(12)), "", GRDPOPUP.Columns(12))
                lblcrtnpack.Caption = IIf(IsNull(GRDPOPUP.Columns(22)), "", GRDPOPUP.Columns(22))
                'LblPack.Text = IIf(IsNull(GRDPOPUP.Columns(14)) Or GRDPOPUP.Columns(14) = "", "1", GRDPOPUP.Columns(14))
                'lblOr_Pack.Caption = IIf(IsNull(GRDPOPUP.Columns(14)) Or GRDPOPUP.Columns(14) = "", "1", GRDPOPUP.Columns(14))
                lblunit.Caption = IIf(IsNull(GRDPOPUP.Columns(15)), "Nos", GRDPOPUP.Columns(15))
                TxtWarranty.Text = IIf(IsNull(GRDPOPUP.Columns(7)), "", GRDPOPUP.Columns(7))
                TxtWarranty_type.Text = IIf(IsNull(GRDPOPUP.Columns(8)), "", GRDPOPUP.Columns(8))
                If GRDPOPUP.Columns(16) = "A" Then
                    txtretaildummy.Text = IIf(IsNull(GRDPOPUP.Columns(18)), "P", GRDPOPUP.Columns(18))
                    TxtRetailmode.Text = "A"
                Else
                    txtretaildummy.Text = IIf(IsNull(GRDPOPUP.Columns(17)), "P", GRDPOPUP.Columns(17))
                    TxtRetailmode.Text = "P"
                End If
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
                TXTUNIT.Text = GRDPOPUP.Columns(20)
                                   
                'TXTPRODUCT.Enabled = False
                'TXTQTY.Enabled = True
                '
                'OptLoose.value = True
                'TXTQTY.SetFocus
                Exit Function
End Function

Private Sub TXTTYPE_GotFocus()
    TXTTYPE.SelStart = 0
    TXTTYPE.SelLength = Len(TXTTYPE.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TXTTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Val(TXTTYPE.Text) = 0 Or Val(TXTTYPE.Text) > 3 Then
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
    If cmbtype.ListIndex = -1 Then
        MsgBox "Select Bill Type from the List", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 0 And Val(TXTTYPE.Text) <> 1 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        TXTTYPE.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 1 And Val(TXTTYPE.Text) <> 2 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        TXTTYPE.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 2 And Val(TXTTYPE.Text) <> 3 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        TXTTYPE.SetFocus
        Exit Sub
    End If
End Sub

Private Function REMOVE_ITEM()
    Dim i As Long
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Function
      
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
        grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = grdsales.TextMatrix(i + 1, 24)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = grdsales.TextMatrix(i + 1, 25)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 26) = grdsales.TextMatrix(i + 1, 26)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 27) = grdsales.TextMatrix(i + 1, 27)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 28) = grdsales.TextMatrix(i + 1, 28)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 29) = grdsales.TextMatrix(i + 1, 29)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 30) = grdsales.TextMatrix(i + 1, 30)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 31) = grdsales.TextMatrix(i + 1, 31)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = grdsales.TextMatrix(i + 1, 32)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 33) = grdsales.TextMatrix(i + 1, 33)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 34) = grdsales.TextMatrix(i + 1, 34)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 35) = grdsales.TextMatrix(i + 1, 35)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 36) = grdsales.TextMatrix(i + 1, 36)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 37) = grdsales.TextMatrix(i + 1, 37)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 38) = grdsales.TextMatrix(i + 1, 38)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 39) = grdsales.TextMatrix(i + 1, 39)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 40) = grdsales.TextMatrix(i + 1, 40)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 41) = grdsales.TextMatrix(i + 1, 41)
    Next i
    grdsales.Rows = grdsales.Rows - 1
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
        Else
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
        End If
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    lblnetamount.Caption = Round(lblnetamount.Caption, 0)
    Call COSTCALCULATION
    Call Addcommission
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    txtPrintname.Text = ""
    txtcategory.Text = ""
    TxtName1.Text = ""
    TXTITEMCODE.Text = ""
    optnet.value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTAPPENDQTY.Text = ""
    TXTFREEAPPEND.Text = ""
    txtappendcomm.Text = ""
    TXTAPPENDTOTAL.Text = ""
    txtretail.Text = ""
    txtBatch.Text = ""
    TxtWarranty.Text = ""
    TxtWarranty_type.Text = ""
    TXTTAX.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    txtcommi.Text = ""
    TxtRetailmode.Text = ""
    
    TXTDISC.Text = ""
    TxtCessAmt.Text = ""
    TxtCessPer.Text = ""
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    LBLDNORCN.Caption = ""
    cmdadd.Enabled = False
    'cmddelete.Enabled = False
    'CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    M_ADD = True
    TXTQTY.Enabled = False
    TXTITEMCODE.Enabled = True
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
        TXTPRODUCT.SetFocus
    Else
        TXTITEMCODE.SetFocus
    End If
    If grdsales.Rows >= 9 Then grdsales.TopRow = grdsales.Rows - 1

End Function

Private Function Addcommission()
    Dim i As Long
    On Error GoTo eRRhAND
    lblActAmt.Caption = ""
    For i = 1 To grdsales.Rows - 1
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblActAmt.Caption = Val(lblActAmt.Caption) + Val(grdsales.TextMatrix(i, 24))
        Else
            lblActAmt.Caption = Val(lblActAmt.Caption) + (Val(grdsales.TextMatrix(i, 24)) * Val(grdsales.TextMatrix(i, 3)))
        End If
    Next i
    
    Exit Function
eRRhAND:
    MsgBox Err.Description
End Function

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 31  'ST_RATE
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.Text), "0.00")
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                Case 8  'Disc
                    TXTDISC.Tag = 0
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(TXTsample.Text) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) - Val(TXTDISC.Tag), 4), ".0000")
                    Else
                        TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(TXTsample.Text) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7))) - Val(TXTDISC.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 6))) - Val(TXTDISC.Tag), 4), ".0000")
                    End If
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.Text), "0.00")
                    LBLTOTAL.Caption = ""
                    lblnetamount.Caption = ""
                    LBLFOT.Caption = ""
                    lblcomamt.Caption = ""
                    Dim i As Long
                    For i = 1 To grdsales.Rows - 1
                        grdsales.TextMatrix(i, 0) = i
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN"
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                        End Select
                        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
                        Else
                            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
                        End If
                    Next i
                    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
                    TXTAMOUNT.Text = ""
                    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
                        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
                    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
                        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
                    End If
                    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
                    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
                    lblnetamount.Caption = Round(lblnetamount.Caption, 0)
                    
                    Dim RSTTRXFILE As ADODB.Recordset
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(grdsales.Row, 8))
                        RSTTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(grdsales.Row, 12))
                        RSTTRXFILE.Update
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
                        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
                        RSTTRXFILE.Update
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            grdsales.SetFocus
    End Select
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
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If grdsales.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            If frmLogin.rs!Level = "0" Then
                Select Case grdsales.Col
                    Case 31
                        TXTsample.Visible = True
                        TXTsample.Top = grdsales.CellTop + 100
                        TXTsample.Left = grdsales.CellLeft + 0
                        TXTsample.Width = grdsales.CellWidth
                        TXTsample.Height = grdsales.CellHeight
                        TXTsample.Text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                        TXTsample.SetFocus
                    Case 8
                        TXTsample.Visible = True
                        TXTsample.Top = grdsales.CellTop + 100
                        TXTsample.Left = grdsales.CellLeft + 0
                        TXTsample.Width = grdsales.CellWidth
                        TXTsample.Height = grdsales.CellHeight
                        TXTsample.Text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                        TXTsample.SetFocus
                End Select
            End If
    End Select
End Sub

Private Sub grdsales_Scroll()
    TXTsample.Visible = False
    grdsales.SetFocus
End Sub

Private Sub Txthandle_GotFocus()
    Txthandle.SelStart = 0
    Txthandle.SelLength = Len(Txthandle.Text)
End Sub

Private Sub Txthandle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            'If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            ''If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
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
End Sub

Private Function Print_A4()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim DN_CN_FLag As Boolean
    Dim i As Long
    Dim CN As Integer
    Dim DN As Integer
    Dim b As Integer
    Dim Num, Figre As Currency
    
    On Error GoTo eRRhAND
    DN = 0
    CN = 0
    b = 0
    DN_CN_FLag = False
    
    txtOutstanding.Text = ""
    Dim m_Rcpt_Amt As Double
    Dim Rcptamt As Double
    Dim m_OP_Bal As Double
    Dim m_Bal_Amt As Double
    
    m_Rcpt_Amt = 0
    m_OP_Bal = 0
    m_Bal_Amt = 0
    Rcptamt = 0
    
    If DataList2.BoundText <> "130000" And DataList2.BoundText <> "130001" Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select OPEN_DB from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            m_OP_Bal = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
               
        Set RSTTRXFILE = New ADODB.Recordset
        'RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' AND NOT(TRX_TYPE= 'RT' AND INV_TRX_TYPE ='SV' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & ") ", db, adOpenStatic, adLockReadOnly, adCmdText
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE <> 'DR' OR TRX_TYPE <> 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
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
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE= 'RT' AND INV_TRX_TYPE ='SV' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
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
        
        If GRDRECEIPT.Rows > 1 Then Rcptamt = GRDRECEIPT.TextMatrix(0, 0)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND NOT(INV_TRX_TYPE = 'SV' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(LBLBILLNO.Caption) & ") AND (TRX_TYPE = 'DR' OR TRX_TYPE = 'DB') AND INV_DATE >= '" & Format(TXTINVDATE.Text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
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
        
        txtOutstanding.Text = (m_OP_Bal + m_Bal_Amt) - (m_Rcpt_Amt)
    End If
    
    If OLD_BILL = False Then Call checklastbill
    OLD_BILL = True
'    db.Execute "delete From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(txtBillNo.Text) & ""
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "Select * FROM TRXMAST ", db, adOpenStatic, adLockOptimistic, adCmdText
'    RSTTRXFILE.AddNew
'    RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
'    RSTTRXFILE!TRX_TYPE = "SV"
'    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
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
'    RSTTRXFILE!BILL_NAME = Trim(TxtBillName.Text)
'    RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.Text)
'    RSTTRXFILE!Area = Trim(TXTAREA.Text)
'    'RSTTRXFILE!BILL_FLAG = "Y"
'    RSTTRXFILE.Update
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    
    db.Execute "delete from TEMPTRXFILE"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "SV"
        'RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!line_no = i
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 13)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 5))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "" '"Issued to     " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
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
        RSTTRXFILE!CESS_AMT = Val(grdsales.TextMatrix(i, 41))
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
        If Tax_Print = False Then
            RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        Else
            RSTTRXFILE!SALES_TAX = Val(TxtTaxPrint.Text)
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
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 30)
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(i, 26)
        
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
'        Dim RSTITEMMAST As ADODB.Recordset
'        Set RSTITEMMAST = New ADODB.Recordset
'        RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
'        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'            RSTTRXFILE!Area = RSTITEMMAST!Area
'        End If
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.Update
    Next i
    
    Dim rstTRXMAST As ADODB.Recordset
    Set rstTRXMAST = New ADODB.Recordset
    rstTRXMAST.Open "SELECT * from RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='SV' AND VCH_NO = " & Val(TxtCN.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstTRXMAST.EOF
        i = i + 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "XC"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!line_no = i
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
        RSTTRXFILE!CHECK_FLAG = "V"
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
        RSTTRXFILE!C_USER_ID = "SM"
                
        RSTTRXFILE.Update
        
        rstTRXMAST.MoveNext
    Loop
    rstTRXMAST.Close
    Set rstTRXMAST = Nothing
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    'Call ReportGeneratION_vpestimate
    LBLFOT.Tag = ""
    If frmLogin.rs!Level <> "0" And NEW_BILL = True Then
        If MsgBox("You do not have any permission to modify this further. Are you sure to print?", vbYesNo, "BILL..") = vbNo Then Exit Function
    Else
        Screen.MousePointer = vbHourglass
        Sleep (300)
    End If
    
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompAddress4, CompAddress5, CompTin, CompCST, BIL_PRE, BILL_SUF, DL, ML, DL1, DL2, INV_TERMS, BANK_DET, PAN_NO As String
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenStatic, adLockReadOnly, adCmdText
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
        'BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
        'BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        If Trim(TxtVehicle.Text) = "" Then TxtVehicle.Text = IIf(IsNull(RSTCOMPANY!VEHICLE), "", RSTCOMPANY!VEHICLE)
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
        DL2 = IIf(IsNull(RSTCOMPANY!REMARKS), "", Trim(RSTCOMPANY!REMARKS))
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
                
    NEW_BILL = False
    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0)) - Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 2)), 2)
    Figre = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0))
    Num = Abs(Figre)
    If Figre < 0 Then
        LBLFOT.Tag = "(-)Rupees " & Words_1_all(Num) & " Only"
    ElseIf Figre > 0 Then
        LBLFOT.Tag = "(Rupees " & Words_1_all(Num) & " Only)"
    End If
    If Val(MDIMAIN.StatusBar.Panels(10).Text) = 1 Then
        If Trim(lblIGST.Caption) <> "Y" Then
            If Small_Print = True Then
                'ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\rptbillretail"
                ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTGSTBILLA51"
            Else
                ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTGSTBILL1"
            End If
        Else
            ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTGSTWBILL"
        End If
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Set CRXFormulaFields = Report.FormulaFields
        
        For i = 1 To Report.Database.Tables.COUNT
            Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        Next i
        Report.DiscardSavedData
        Set CRXFormulaFields = Report.FormulaFields
        For Each CRXFormulaField In CRXFormulaFields
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
            If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.Text = "'" & INV_TERMS & "'"
            If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.Text = "'" & BANK_DET & "'"
            If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.Text = "'" & PAN_NO & "'"
            If CRXFormulaField.Name = "{@DL1}" Then CRXFormulaField.Text = "'" & DL1 & "'"
            If CRXFormulaField.Name = "{@DL2}" Then CRXFormulaField.Text = "'" & DL2 & "'"
            If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & TxtBillName.Text & "'"
            If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.Text = "'" & Trim(TXTDEALER.Text) & "'"
            If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.Text = "'" & lbladdress.Caption & "'"
            If TxtPhone.Text = "" Then
                If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
            Else
                If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
                'If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "' & chr(13) & 'Ph: ' & '" & Trim(TxtPhone.Text) & "'"
            End If
            If lblIGST.Caption = "Y" Then
                If CRXFormulaField.Name = "{@IGSTFLAG}" Then CRXFormulaField.Text = "'Y'"
            Else
                If CRXFormulaField.Name = "{@IGSTFLAG}" Then CRXFormulaField.Text = "'N'"
            End If
            'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
            If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
            If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
    '        If Tax_Print = False Then
    '            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
    '        End If
            If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
                If CRXFormulaField.Name = "{@condition}" Then CRXFormulaField.Text = "'" & Trim(Terms1.Text) & "'"
            End If
            If CRXFormulaField.Name = "{@Area}" Then CRXFormulaField.Text = "'" & Trim(TXTAREA.Text) & "'"
            If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
            If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & TxtPhone.Text & "'"
            If CRXFormulaField.Name = "{@VCH_NO}" Then CRXFormulaField.Text = "'" & BIL_PRE & "' &'" & Format(Trim(txtBillNo.Text), "0000") & "' & '" & BILL_SUF & "' "
            If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
            If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.Text = "'" & Trim(TxtOrder.Text) & "'"
            If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = " " & Val(LBLDISCAMT.Caption) & " "
    '            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
            If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.Text = "'" & Trim(TxtCode.Text) & "'"
            If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.Text = " " & Val(txtOutstanding.Text) & " "
            If CRXFormulaField.Name = "{@RcptAmt}" Then CRXFormulaField.Text = " " & Rcptamt & " "
            If CRXFormulaField.Name = "{@Frieght}" Then CRXFormulaField.Text = "'" & Trim(lblFrieght.Text) & "'"
            If CRXFormulaField.Name = "{@FC}" Then CRXFormulaField.Text = " " & Val(TxtFrieght.Text) & " "
            If CRXFormulaField.Name = "{@HANDLE}" Then CRXFormulaField.Text = " '" & Trim(lblhandle.Text) & "' "
            If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.Text = " " & Val(Txthandle.Text) & " "
            If CRXFormulaField.Name = "{@DISCPER}" Then CRXFormulaField.Text = " " & Val(TXTTOTALDISC.Text) & " "
            
            If Val(LBLRETAMT.Caption) = 0 Then
                If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'N' "
            Else
                If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'Y' "
            End If
            If lblcredit.Caption = "0" Then
                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Cash'"
            Else
                If Val(txtcrdays.Text) > 0 Then
                    If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'" & txtcrdays.Text & "'" & "' Days Credit'"
                Else
                    If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Credit'"
                End If
            End If
        Next
    Else
        If Trim(lblIGST.Caption) <> "Y" Then
            If Small_Print = True Then
                'ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\rptbillretail"
                ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTGSTSERVICEA5"
            Else
                ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTGSTSERVICE"
            End If
        Else
            ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RPTGSTWSERVICE"
        End If
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Set CRXFormulaFields = Report.FormulaFields
        
        For i = 1 To Report.OpenSubreport("RPTBILL1.rpt").Database.Tables.COUNT
            Report.OpenSubreport("RPTBILL1.rpt").Database.Tables(i).SetLogOnInfo strConnection
        Next i
        For i = 1 To Report.OpenSubreport("RPTBILL2.rpt").Database.Tables.COUNT
            Report.OpenSubreport("RPTBILL2.rpt").Database.Tables(i).SetLogOnInfo strConnection
        Next i
        For i = 1 To Report.OpenSubreport("RPTBILL3.rpt").Database.Tables.COUNT
            Report.OpenSubreport("RPTBILL3.rpt").Database.Tables(i).SetLogOnInfo strConnection
        Next i
        For i = 1 To 3
            'Set CRXFormulaFields = Report.FormulaFields
            Set CRXFormulaFields = Report.OpenSubreport("RPTBILL" & i & ".rpt").FormulaFields
            For Each CRXFormulaField In CRXFormulaFields
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
                If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & Trim(TxtBillName.Text) & "'"
                If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.Text = "'" & Trim(TXTDEALER.Text) & "'"
                If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.Text = "'" & Trim(lbladdress.Caption) & "'"
                If CRXFormulaField.Name = "{DLNO2}" Then CRXFormulaField.Text = "'" & DL1 & "'"
                If CRXFormulaField.Name = "{DLNO}" Then CRXFormulaField.Text = "'" & DL2 & "'"
                If TxtPhone.Text = "" Then
                    If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
                Else
                    If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
                    'If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "' & chr(13) & 'Ph: ' & '" & Trim(TxtPhone.Text) & "'"
                End If
                If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
                    If CRXFormulaField.Name = "{@condition}" Then CRXFormulaField.Text = "'" & Trim(Terms1.Text) & "'"
                End If
                If CRXFormulaField.Name = "{@Area}" Then CRXFormulaField.Text = "'" & Trim(TXTAREA.Text) & "'"
                'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
                If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
        '            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
        '            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
                If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
        '        If Tax_Print = False Then
        '            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
        '        End If
                If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
                    If CRXFormulaField.Name = "{@condition}" Then CRXFormulaField.Text = "'" & Trim(Terms1.Text) & "'"
                End If
                If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
                If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & TxtPhone.Text & "'"
                If CRXFormulaField.Name = "{@VCH_NO}" Then CRXFormulaField.Text = "'" & BIL_PRE & "' &'" & Format(Trim(txtBillNo.Text), "0000") & "' & '" & BILL_SUF & "' "
                If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
                If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.Text = "'" & Trim(TxtOrder.Text) & "'"
                If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = " " & Val(LBLDISCAMT.Caption) & " "
        '            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
                If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.Text = "'" & Trim(TxtCode.Text) & "'"
                If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.Text = " " & Val(txtOutstanding.Text) & " "
                If CRXFormulaField.Name = "{@RcptAmt}" Then CRXFormulaField.Text = " " & Rcptamt & " "
                If CRXFormulaField.Name = "{@Frieght}" Then CRXFormulaField.Text = "'" & Trim(lblFrieght.Text) & "'"
                If CRXFormulaField.Name = "{@FC}" Then CRXFormulaField.Text = " " & Val(TxtFrieght.Text) & " "
                If CRXFormulaField.Name = "{@HANDLE}" Then CRXFormulaField.Text = " '" & Trim(lblhandle.Text) & "' "
                If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.Text = " " & Val(Txthandle.Text) & " "
                If CRXFormulaField.Name = "{@DISCPER}" Then CRXFormulaField.Text = " " & Val(TXTTOTALDISC.Text) & " "
                
                If Val(LBLRETAMT.Caption) = 0 Then
                    If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'N' "
                Else
                    If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'Y' "
                End If
                If lblcredit.Caption = "0" Then
                    If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Cash'"
                Else
                    If Val(txtcrdays.Text) > 0 Then
                        If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'" & txtcrdays.Text & "'" & "' Days Credit'"
                    Else
                        If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Credit'"
                    End If
                End If
            Next
        Next i
    End If
    
    
    If MDIMAIN.StatusBar.Panels(13).Text = "Y" Then
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
        Call cmdRefresh_Click
        Exit Function
    End If
    
SKIP:
    CMDEXIT.Enabled = False
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = True
    TXTITEMCODE.Enabled = True
    TXTQTY.Enabled = False
    
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    txtretail.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTDISC.Enabled = False
    txtcommi.Enabled = False
    OLD_BILL = True
    ''rptPRINT.Action = 1
    Exit Function
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Function

Private Sub CmdTax_Click()
    If grdsales.Rows <= 1 Then Exit Sub
    If Trim(TxtTaxPrint.Text) = "" Then Exit Sub
    Tax_Print = True
    Call Generateprint
    TxtTaxPrint.Text = ""
End Sub

Private Sub TxtFree_GotFocus()
    TXTFREE.SelStart = 0
    TXTFREE.SelLength = Len(TXTFREE.Text)
    TXTFREE.Tag = Trim(TXTPRODUCT.Text)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 And Val(TXTFREE.Text) = 0 Then
                TXTFREE.Enabled = True
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            End If
            If Val(TXTFREE.Text) = 0 Then GoTo SKIP
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT CLOSE_QTY  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!CLOSE_QTY)) Then RSTTRXFILE!CLOSE_QTY = 0
                i = RSTTRXFILE!CLOSE_QTY / Val(LblPack.Text)
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
                If M_EDIT = False And SERIAL_FLAG = True And (Val(TXTFREE.Text) + Val(TXTQTY.Text)) > Val(lblactqty.Caption) Then
                    MsgBox "AVAILABLE STOCK IS  " & Val(lblactqty.Caption) & " ", , "SALES"
                    TXTQTY.SelStart = 0
                    TXTQTY.SelLength = Len(TXTQTY.Text)
                    Exit Sub
                End If
                If M_EDIT = False And (Val(TXTFREE.Text) + Val(TXTQTY.Text)) > i Then
                    If SERIAL_FLAG = True Then
                        MsgBox "AVAILABLE STOCK IS  " & i & " ", , "SALES"
                        TXTFREE.SelStart = 0
                        TXTFREE.SelLength = Len(TXTFREE.Text)
                        Exit Sub
                    End If
                    If (MsgBox("AVAILABLE STOCK IS  " & i & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                        'MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                        TXTQTY.SelStart = 0
                        TXTQTY.SelLength = Len(TXTQTY.Text)
                        Exit Sub
                    End If
                End If
'            End If
            
SKIP:
            txtretail.Enabled = True
            txtretail.SetFocus
'            TXTFREE.Enabled = False
'            TXTTAX.Enabled = True
'            TXTTAX.SetFocus
         Case vbKeyEscape
            TXTFREE.Enabled = True
            TXTQTY.Enabled = True
            
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

Private Sub TxtOrder_GotFocus()
    'If Trim(TxtOrder.Text) = "" Then TxtOrder.Text = "KL-04-N-8931"
    TxtOrder.SelStart = 0
    TxtOrder.SelLength = Len(TxtOrder.Text)
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
    On Error GoTo eRRhAND
    Set RSTITEMMAST = New ADODB.Recordset
    'RSTITEMMAST.Open "SELECT * FROM CUSTTRXFILE WHERE ACT_CODE = '" & Txtsuplcode.Text & "' and ACT_CODE <> '130000'", db, adOpenStatic, adLockReadOnly
    RSTITEMMAST.Open "SELECT * FROM CUSTTRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "' AND BR_CODE = '" & CMBBRNCH.BoundText & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        lbladdress.Caption = CMBBRNCH.Text & Chr(13) & Trim(RSTITEMMAST!Address)
        'TxtBillName.Text = CMBBRNCH.Text
        TxtBillAddress.Text = IIf(IsNull(RSTITEMMAST!Address), "", Trim(RSTITEMMAST!Address))
        TXTTIN.Text = IIf(IsNull(RSTITEMMAST!KGST), "", Trim(RSTITEMMAST!KGST))
        TxtPhone.Text = IIf(IsNull(RSTITEMMAST!TELNO), "", Trim(RSTITEMMAST!TELNO))
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub CMBBRNCH_GotFocus()
    CMBBRNCH.Text = ""
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

Function FILL_BATCHGRID2()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, ITEM_CODE From RTRXFILE  WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, ITEM_CODE From RTRXFILE  WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
        BATCH_FLAG = False
    End If
    Set GRDPOPUP.DataSource = PHY_BATCH
    TXTITEMCODE.Text = GRDPOPUP.Columns(26)
    ITEM_CHANGE = True
    TXTPRODUCT.Text = GRDPOPUP.Columns(6)
    ITEM_CHANGE = False
    GRDPOPUP.Columns(0).Caption = "Serial No."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "VCH No"
    GRDPOPUP.Columns(3).Caption = "Line No"
    GRDPOPUP.Columns(4).Caption = "Trx Type"
    GRDPOPUP.Columns(7).Caption = "" '"Warranty"
    GRDPOPUP.Columns(8).Caption = ""
    GRDPOPUP.Columns(25).Caption = "Expiry"
    
    GRDPOPUP.Columns(0).Width = 4100
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
    
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).Text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True

End Function

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
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
        Case vbKeyEscape
             If Len(Trim(TXTEXPIRY.Text)) = 1 Then GoTo Nextstep
            If Len(Trim(TXTEXPIRY.Text)) < 5 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
Nextstep:
            TxtMRP.Enabled = True
            TxtMRP.SetFocus
    End Select
End Sub

Private Sub TxtCessPer_GotFocus()
    TxtCessPer.SelStart = 0
    TxtCessPer.SelLength = Len(TxtCessPer.Text)
End Sub

Private Sub TxtCessPer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtCessAmt.Text) <> 0 Then
                TxtCessAmt.Enabled = True
                TxtCessAmt.SetFocus
            Else
                If lblsubdealer.Caption = "D" Then
                    txtcommi.Enabled = True
                    txtcommi.SetFocus
                Else
                    txtcommi.Text = 0
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
    
'    TxtCessPer.Tag = 0
'    If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
'        TxtCessPer.Tag = Val(txtretail.Text) * Val(TxtCessPer.Text) / 100
'        LBLSUBTOTAL.Caption = Format(Round(Val(txtretail.Text) - Val(TxtCessPer.Tag), 2), ".000")
'        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.Text) - Val(TxtCessPer.Tag), 2), ".000")
'    Else
'        TxtCessPer.Tag = Val(TXTQTY.Text) * Val(txtretail.Text) * Val(TxtCessPer.Text) / 100
'        LBLSUBTOTAL.Caption = Format(Round((Val(TXTQTY.Text) * Val(txtretail.Text)) - Val(TxtCessPer.Tag), 2), ".000")
'        LblGross.Caption = Format(Round((Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text)) - Val(TxtCessPer.Tag), 2), ".000")
'    End If
    
    ''TxtCessPer.Text = Format(TxtCessPer.Text, ".000")

End Sub

Private Sub TxtCessAmt_GotFocus()
    TxtCessAmt.SelStart = 0
    TxtCessAmt.SelLength = Len(TxtCessAmt.Text)
End Sub

Private Sub TxtCessAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If lblsubdealer.Caption = "D" Then
                txtcommi.Enabled = True
                txtcommi.SetFocus
            Else
                txtcommi.Text = 0
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
    
'    TxtCessAmt.Tag = 0
'    If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
'        TxtCessAmt.Tag = Val(txtretail.Text) * Val(TxtCessAmt.Text) / 100
'        LBLSUBTOTAL.Caption = Format(Round(Val(txtretail.Text) - Val(TxtCessAmt.Tag), 2), ".000")
'        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.Text) - Val(TxtCessAmt.Tag), 2), ".000")
'    Else
'        TxtCessAmt.Tag = Val(TXTQTY.Text) * Val(txtretail.Text) * Val(TxtCessAmt.Text) / 100
'        LBLSUBTOTAL.Caption = Format(Round((Val(TXTQTY.Text) * Val(txtretail.Text)) - Val(TxtCessAmt.Tag), 2), ".000")
'        LblGross.Caption = Format(Round((Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text)) - Val(TxtCessAmt.Tag), 2), ".000")
'    End If
'
'    ''TxtCessAmt.Text = Format(TxtCessAmt.Text, ".000")

End Sub



