VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMDELIVERY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Delivery"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16665
   Icon            =   "FrmDelivery.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   16665
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   7650
      TabIndex        =   74
      Top             =   4020
      Visible         =   0   'False
      Width           =   8955
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   2610
         Left            =   30
         TabIndex        =   75
         Top             =   420
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4604
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
         Left            =   3810
         TabIndex        =   77
         Top             =   45
         Width           =   5115
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
         Left            =   30
         TabIndex        =   76
         Top             =   45
         Width           =   3780
      End
   End
   Begin MSDataGridLib.DataGrid grdtmp 
      Height          =   4125
      Left            =   2130
      TabIndex        =   143
      Top             =   2970
      Visible         =   0   'False
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   7276
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
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3270
      Left            =   900
      TabIndex        =   43
      Top             =   3135
      Visible         =   0   'False
      Width           =   10965
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   3165
         Left            =   45
         TabIndex        =   44
         Top             =   60
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
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   2745
      TabIndex        =   39
      Top             =   3270
      Visible         =   0   'False
      Width           =   5835
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2565
         Left            =   90
         TabIndex        =   42
         Top             =   360
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   4524
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
         Index           =   9
         Left            =   90
         TabIndex        =   41
         Top             =   105
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
         Height          =   240
         Index           =   0
         Left            =   3135
         TabIndex        =   40
         Top             =   105
         Visible         =   0   'False
         Width           =   2610
      End
   End
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   -150
      TabIndex        =   11
      Top             =   -15
      Width           =   16770
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
         Left            =   13950
         MaxLength       =   15
         TabIndex        =   45
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00FFC0C0&
         Height          =   2460
         Left            =   210
         TabIndex        =   33
         Top             =   -75
         Width           =   16560
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H000000FF&
            Height          =   2310
            Left            =   9045
            TabIndex        =   146
            Top             =   90
            Width           =   3705
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
               Height          =   330
               Left            =   840
               MaxLength       =   35
               TabIndex        =   149
               Top             =   1530
               Width           =   2805
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
               Height          =   330
               Left            =   810
               MaxLength       =   35
               TabIndex        =   148
               Top             =   1155
               Width           =   2835
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
               Height          =   330
               Left            =   735
               MaxLength       =   35
               TabIndex        =   147
               Top             =   1905
               Width           =   2925
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "LPO No."
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
               Left            =   75
               TabIndex        =   153
               Top             =   1530
               Width           =   765
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
               Height          =   930
               Left            =   45
               TabIndex        =   152
               Top             =   210
               Width           =   3615
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "TRN No."
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
               TabIndex        =   151
               Top             =   1170
               Width           =   735
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
               Left            =   75
               TabIndex        =   150
               Top             =   1920
               Width           =   660
            End
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
            Left            =   6360
            TabIndex        =   8
            Top             =   2175
            Visible         =   0   'False
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
            ItemData        =   "FrmDelivery.frx":030A
            Left            =   6390
            List            =   "FrmDelivery.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1755
            Width           =   2580
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H000000FF&
            Height          =   1650
            Left            =   5205
            TabIndex        =   57
            Top             =   90
            Width           =   3810
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
               MaxLength       =   35
               TabIndex        =   6
               Top             =   225
               Width           =   3645
            End
            Begin MSForms.TextBox TxtBillAddress 
               Height          =   1005
               Left            =   90
               TabIndex        =   7
               Top             =   585
               Width           =   3645
               VariousPropertyBits=   746604571
               MaxLength       =   100
               BorderStyle     =   1
               Size            =   "6429;1773"
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
            Height          =   315
            Left            =   4200
            TabIndex        =   2
            Top             =   510
            Width           =   960
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
            Height          =   330
            Left            =   1290
            TabIndex        =   3
            Top             =   855
            Width           =   3855
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
            Left            =   1305
            TabIndex        =   0
            Top             =   150
            Width           =   885
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   300
            Left            =   3735
            TabIndex        =   1
            Top             =   150
            Width           =   1425
            _ExtentX        =   2514
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
            Height          =   780
            Left            =   1290
            TabIndex        =   4
            Top             =   1215
            Width           =   3855
            _ExtentX        =   6800
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
         Begin MSDataListLib.DataCombo CMBDISTI 
            Height          =   2055
            Left            =   13485
            TabIndex        =   10
            Top             =   300
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   3625
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
         Begin MSForms.ComboBox TXTAREA 
            Height          =   330
            Left            =   1290
            TabIndex        =   5
            Top             =   2085
            Width           =   3870
            VariousPropertyBits=   746604571
            ForeColor       =   255
            MaxLength       =   20
            DisplayStyle    =   3
            Size            =   "6826;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            DropButtonStyle =   0
            BorderColor     =   255
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
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
            Left            =   105
            TabIndex        =   73
            Top             =   2085
            Width           =   825
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "3.   WS"
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
            Index           =   39
            Left            =   7020
            TabIndex        =   72
            Top             =   2085
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "2.   RT"
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
            Index           =   38
            Left            =   6165
            TabIndex        =   71
            Top             =   2070
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "1.   VP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   165
            Index           =   37
            Left            =   5310
            TabIndex        =   70
            Top             =   2055
            Visible         =   0   'False
            Width           =   675
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
            Index           =   16
            Left            =   5235
            TabIndex        =   59
            Top             =   1770
            Width           =   1110
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
            Left            =   12810
            TabIndex        =   58
            Top             =   390
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Days"
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
            Left            =   2760
            TabIndex        =   56
            Top             =   540
            Width           =   1140
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
            Left            =   105
            TabIndex        =   47
            Top             =   885
            Width           =   1230
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "INV DATE"
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
            Left            =   2760
            TabIndex        =   46
            Top             =   165
            Width           =   885
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
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   37
            Top             =   195
            Width           =   1230
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
            Left            =   135
            TabIndex        =   36
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
            Left            =   1290
            TabIndex        =   35
            Top             =   510
            Visible         =   0   'False
            Width           =   1305
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
            Left            =   1305
            TabIndex        =   34
            Top             =   150
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   4785
         Left            =   210
         TabIndex        =   38
         Top             =   2280
         Width           =   16560
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            Height          =   4620
            Left            =   13230
            TabIndex        =   113
            Top             =   105
            Width           =   3285
            Begin VB.Label LBLTOTAL 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label2"
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
               Height          =   500
               Left            =   195
               TabIndex        =   131
               Top             =   330
               Width           =   1440
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
               Left            =   180
               TabIndex        =   130
               Top             =   105
               Width           =   1485
            End
            Begin VB.Label lblnetamount 
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
               TabIndex        =   129
               Top             =   1245
               Width           =   1440
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   23
               Left            =   195
               TabIndex        =   128
               Top             =   1005
               Width           =   1440
            End
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   25
               Left            =   1785
               TabIndex        =   127
               Top             =   105
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
               Left            =   1785
               TabIndex        =   126
               Top             =   1245
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
               Height          =   495
               Left            =   1785
               TabIndex        =   125
               Top             =   2085
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
               Left            =   1785
               TabIndex        =   124
               Top             =   3015
               Visible         =   0   'False
               Width           =   1440
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   27
               Left            =   1785
               TabIndex        =   123
               Top             =   1860
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
               Left            =   1785
               TabIndex        =   122
               Top             =   2790
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
               TabIndex        =   121
               Top             =   4050
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
               TabIndex        =   120
               Top             =   3780
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
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
               Left            =   195
               TabIndex        =   119
               Top             =   1845
               Width           =   1440
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
               Height          =   495
               Left            =   195
               TabIndex        =   118
               Top             =   2085
               Width           =   1440
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
               Height          =   495
               Left            =   195
               TabIndex        =   117
               Top             =   3015
               Visible         =   0   'False
               Width           =   1440
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
               Left            =   180
               TabIndex        =   116
               Top             =   2760
               Visible         =   0   'False
               Width           =   1455
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
               Height          =   495
               Left            =   1785
               TabIndex        =   115
               Top             =   330
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   26
               Left            =   1785
               TabIndex        =   114
               Top             =   975
               Visible         =   0   'False
               Width           =   1425
            End
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
            TabIndex        =   55
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
            TabIndex        =   54
            Top             =   4380
            Width           =   930
         End
         Begin VB.OptionButton OPTDISCPERCENT 
            BackColor       =   &H00FFC0C0&
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
            Height          =   330
            Left            =   10050
            TabIndex        =   53
            Top             =   4380
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton OptDiscAmt 
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   52
            Top             =   4380
            Width           =   1125
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4215
            Left            =   90
            TabIndex        =   12
            Top             =   150
            Width           =   13110
            _ExtentX        =   23125
            _ExtentY        =   7435
            _Version        =   393216
            Rows            =   1
            Cols            =   33
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
         Begin VB.Label lblunit 
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
            Left            =   6225
            TabIndex        =   141
            Top             =   4425
            Width           =   765
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
            Height          =   285
            Index           =   42
            Left            =   4740
            TabIndex        =   140
            Top             =   4425
            Width           =   600
         End
         Begin VB.Label lblpack 
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
            Left            =   5340
            TabIndex        =   139
            Top             =   4425
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
            Left            =   8490
            TabIndex        =   69
            Top             =   4425
            Width           =   780
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
            Height          =   285
            Left            =   9270
            TabIndex        =   68
            Top             =   4425
            Width           =   615
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
            Height          =   285
            Left            =   7635
            TabIndex        =   67
            Top             =   4425
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
            Height          =   285
            Left            =   3840
            TabIndex        =   66
            Top             =   4425
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
            Height          =   285
            Left            =   2325
            TabIndex        =   65
            Top             =   4425
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
            Height          =   285
            Left            =   720
            TabIndex        =   64
            Top             =   4425
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
            Left            =   7050
            TabIndex        =   63
            Top             =   4425
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
            Height          =   285
            Index           =   30
            Left            =   3240
            TabIndex        =   62
            Top             =   4425
            Width           =   600
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
            Height          =   285
            Index           =   22
            Left            =   1635
            TabIndex        =   61
            Top             =   4425
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
            Height          =   285
            Index           =   21
            Left            =   90
            TabIndex        =   60
            Top             =   4425
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Height          =   1755
         Left            =   210
         TabIndex        =   78
         Top             =   6975
         Width           =   16560
         Begin VB.TextBox TXTRYEAR 
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
            Left            =   14985
            MaxLength       =   8
            TabIndex        =   145
            Top             =   270
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
            TabIndex        =   135
            Top             =   900
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
            TabIndex        =   134
            Top             =   840
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
            TabIndex        =   133
            Top             =   1200
            Visible         =   0   'False
            Width           =   1200
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
            Left            =   13755
            MaxLength       =   8
            TabIndex        =   132
            Top             =   1215
            Visible         =   0   'False
            Width           =   1200
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
            Left            =   555
            MaxLength       =   15
            TabIndex        =   112
            Top             =   390
            Width           =   1500
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
            Left            =   4995
            MaxLength       =   6
            TabIndex        =   91
            Top             =   1620
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
            TabIndex        =   90
            Top             =   1605
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtcommi 
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
            Left            =   11295
            MaxLength       =   6
            TabIndex        =   89
            Top             =   375
            Width           =   855
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
            TabIndex        =   88
            Top             =   1620
            Visible         =   0   'False
            Width           =   795
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
            Left            =   8460
            MaxLength       =   9
            TabIndex        =   20
            Top             =   375
            Width           =   915
         End
         Begin VB.TextBox TXTRETAILNOTAX 
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
            Left            =   7605
            MaxLength       =   9
            TabIndex        =   19
            Top             =   375
            Width           =   840
         End
         Begin VB.OptionButton optnet 
            BackColor       =   &H00FF8080&
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
            Height          =   240
            Left            =   9600
            TabIndex        =   25
            Top             =   855
            Width           =   1005
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
            Left            =   2115
            MaxLength       =   6
            TabIndex        =   87
            Top             =   1635
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00FF8080&
            Caption         =   "VAT %"
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
            Height          =   240
            Left            =   8505
            TabIndex        =   24
            Top             =   855
            Width           =   1140
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00FF8080&
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
            Height          =   240
            Left            =   6630
            TabIndex        =   23
            Top             =   855
            Visible         =   0   'False
            Width           =   1875
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
            TabIndex        =   86
            Top             =   1230
            Visible         =   0   'False
            Width           =   630
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
            Left            =   6450
            MaxLength       =   7
            TabIndex        =   16
            Top             =   375
            Width           =   540
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
            TabIndex        =   26
            Top             =   1170
            Width           =   1380
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
            TabIndex        =   85
            Top             =   1170
            Width           =   420
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
            Left            =   10080
            TabIndex        =   84
            Top             =   1170
            Width           =   420
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
            Height          =   300
            Left            =   10185
            MaxLength       =   6
            TabIndex        =   17
            Top             =   1170
            Visible         =   0   'False
            Width           =   705
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
            Left            =   2940
            TabIndex        =   27
            Top             =   1170
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
            Height          =   450
            Left            =   60
            TabIndex        =   13
            Top             =   375
            Width           =   480
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
            Height          =   450
            Left            =   2070
            TabIndex        =   14
            Top             =   375
            Width           =   3450
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
            Left            =   5535
            MaxLength       =   8
            TabIndex        =   15
            Top             =   375
            Width           =   900
         End
         Begin VB.TextBox TXTTAX 
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
            Left            =   7005
            MaxLength       =   4
            TabIndex        =   18
            Top             =   375
            Width           =   585
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
            Left            =   9390
            MaxLength       =   4
            TabIndex        =   22
            Top             =   375
            Width           =   660
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
            Left            =   6480
            TabIndex        =   30
            Top             =   1170
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
            Left            =   8895
            TabIndex        =   32
            Top             =   1170
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
            Left            =   5295
            TabIndex        =   29
            Top             =   1170
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
            Height          =   405
            Left            =   4095
            TabIndex        =   28
            Top             =   1170
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
            Left            =   555
            TabIndex        =   83
            Top             =   1650
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VB.TextBox txtBatch 
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
            Height          =   450
            Left            =   10065
            MaxLength       =   15
            TabIndex        =   21
            Top             =   375
            Width           =   1215
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
            TabIndex        =   82
            Top             =   1395
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
            TabIndex        =   81
            Top             =   1395
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
            TabIndex        =   80
            Top             =   1425
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
            Left            =   12480
            TabIndex        =   79
            Top             =   990
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
            Height          =   405
            Left            =   7695
            TabIndex        =   31
            Top             =   1170
            Width           =   1125
         End
         Begin VB.Frame FrmeType 
            BackColor       =   &H00FF8080&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4470
            TabIndex        =   136
            Top             =   750
            Width           =   1965
            Begin VB.OptionButton OptNormal 
               BackColor       =   &H00FF8080&
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
               Height          =   210
               Left            =   45
               TabIndex        =   138
               Top             =   120
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.OptionButton OptLoose 
               BackColor       =   &H00FF8080&
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
               Height          =   210
               Left            =   1140
               TabIndex        =   137
               Top             =   120
               Width           =   810
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Category"
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
            Index           =   44
            Left            =   555
            TabIndex        =   144
            Top             =   150
            Width           =   1485
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
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   43
            Left            =   555
            TabIndex        =   142
            Top             =   1425
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Commi"
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
            Index           =   5
            Left            =   11295
            TabIndex        =   111
            Top             =   150
            Width           =   855
         End
         Begin VB.Label lblP_Rate 
            Caption         =   "0"
            Height          =   390
            Left            =   10935
            TabIndex        =   110
            Top             =   1095
            Visible         =   0   'False
            Width           =   375
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
            Left            =   7605
            TabIndex        =   109
            Top             =   150
            Width           =   840
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
            Index           =   29
            Left            =   6450
            TabIndex        =   108
            Top             =   150
            Width           =   540
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
            TabIndex        =   107
            Top             =   885
            Visible         =   0   'False
            Width           =   510
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
            Left            =   10185
            TabIndex        =   106
            Top             =   945
            Visible         =   0   'False
            Width           =   705
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
            Left            =   60
            TabIndex        =   105
            Top             =   150
            Width           =   480
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
            Left            =   2070
            TabIndex        =   104
            Top             =   150
            Width           =   3450
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
            Left            =   5535
            TabIndex        =   103
            Top             =   150
            Width           =   900
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
            Left            =   8460
            TabIndex        =   102
            Top             =   150
            Width           =   915
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
            Left            =   7005
            TabIndex        =   101
            Top             =   150
            Width           =   585
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
            Left            =   9390
            TabIndex        =   100
            Top             =   150
            Width           =   660
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
            Left            =   12165
            TabIndex        =   99
            Top             =   150
            Width           =   1395
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
            TabIndex        =   98
            Top             =   1395
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Part No."
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
            Left            =   10065
            TabIndex        =   97
            Top             =   150
            Width           =   1215
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
            Left            =   12165
            TabIndex        =   96
            Top             =   375
            Width           =   1395
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
            TabIndex        =   95
            Top             =   1410
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
            TabIndex        =   94
            Top             =   1440
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
            Left            =   11340
            TabIndex        =   93
            Top             =   1005
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
            TabIndex        =   92
            Top             =   1410
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   13155
      TabIndex        =   48
      Top             =   3090
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1482
      _Version        =   393216
   End
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   -15
      TabIndex        =   51
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11355
      TabIndex        =   50
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   49
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMDELIVERY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Dim M_EDIT, OLD_BILL As Boolean
Dim Delivered_flag As Boolean
Dim BillNO As String

Private Sub cmbtype_GotFocus()
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    If cmbtype.ListIndex = -1 Then cmbtype.ListIndex = 0
End Sub

Private Sub cmbtype_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If cmbtype.ListIndex = -1 Then
                MsgBox "Select Bill Type from the List", vbOKOnly, "Sales"
                cmbtype.SetFocus
                Exit Sub
            End If
'            If cmbtype.ListIndex = 0 And Val(TXTTYPE.Text) <> 1 Then
'                MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
'                cmbtype.SetFocus
'                Exit Sub
'            End If
'            If cmbtype.ListIndex = 1 And Val(TXTTYPE.Text) <> 2 Then
'                MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
'                cmbtype.SetFocus
'                Exit Sub
'            End If
'            If cmbtype.ListIndex = 2 And Val(TXTTYPE.Text) <> 3 Then
'                MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
'                cmbtype.SetFocus
'                Exit Sub
'            End If
            TxtVehicle.Enabled = True
            TxtVehicle.SetFocus
        Case vbKeyEscape
            TxtBillName.Enabled = True
            TxtBillName.SetFocus
    End Select
End Sub

Private Sub cmbtype_LostFocus()
    If cmbtype.ListIndex = -1 Then
        MsgBox "Select Bill Type from the List", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
'    If cmbtype.ListIndex = 0 And Val(TXTTYPE.Text) <> 1 Then
'        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
'        cmbtype.SetFocus
'        Exit Sub
'    End If
'    If cmbtype.ListIndex = 1 And Val(TXTTYPE.Text) <> 2 Then
'        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
'        cmbtype.SetFocus
'        Exit Sub
'    End If
'    If cmbtype.ListIndex = 2 And Val(TXTTYPE.Text) <> 3 Then
'        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
'        cmbtype.SetFocus
'        Exit Sub
'    End If
    'CMBDISTI.Enabled = True
    'CMBDISTI.SetFocus
End Sub

Private Sub DataList2_Click()
    Dim rstCustomer As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    On Error GoTo Errhand
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = DataList2.Text & Chr(13) & Trim(rstCustomer!Address)
        If TxtBillName.Text = "" Then TxtBillName.Text = DataList2.Text
        'TxtBillName.Text = DataList2.Text
        'If TxtBillAddress.Text = "" Then TxtBillAddress.Text = IIf(IsNull(rstCustomer!ADDRESS), "", Trim(rstCustomer!ADDRESS))
        TxtBillAddress.Text = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
        TXTTIN.Text = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
        TxtPhone.Text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO))
        TXTAREA.Text = IIf(IsNull(rstCustomer!Area), "", Trim(rstCustomer!Area))
        'lblcusttype.Caption = IIf((IsNull(rstCustomer!Type) Or rstCustomer!Type = "R"), "R", "W")
        
    Else
        TxtPhone.Text = ""
        TXTTIN.Text = ""
        lbladdress.Caption = ""
        TXTAREA.Text = ""
        'lblcusttype.Caption = "R"
    End If
        
    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text
    Exit Sub
    
Errhand:
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
            TxtBillName.SetFocus
            'FRMEHEAD.Enabled = False
            'TXTSLNO.Enabled = True
            'TXTSLNO.SetFocus
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
    'Dim RSTMINQTY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo Errhand
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
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Format(Val(LBLSUBTOTAL.Caption) + Val(TXTAPPENDTOTAL.Text), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Trim(TXTTRXTYPE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = Trim(TXTRYEAR.Text)
    
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
    grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = Val(txtcommi.Text) + Val(txtappendcomm.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = Trim(txtcategory.Text)
    
    Select Case OptLoose.value
        Case True
            grdsales.TextMatrix(Val(TXTSLNO.Text), 26) = "L"
        Case Else
            grdsales.TextMatrix(Val(TXTSLNO.Text), 26) = "F"
    End Select
    grdsales.TextMatrix(Val(TXTSLNO.Text), 27) = IIf(Val(LblPack.Caption) = 0, "1", Val(LblPack.Caption))
    
    If Not (UCase(txtcategory.Text) = "SERVICES" Or UCase(txtcategory.Text) = "SELF") Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTQTY.Text) * Val(LblPack.Caption)), 3)
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !FREE_QTY = !FREE_QTY + Round((Val(TXTFREE.Text) * Val(LblPack.Caption)), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                !CLOSE_QTY = !CLOSE_QTY - Round(((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Caption)), 3)
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & "' AND BAL_QTY > 0", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Caption), 3)
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY - Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Caption), 3)
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    LBLCOMAMT.Caption = ""
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
        LBLCOMAMT.Caption = Val(LBLCOMAMT.Caption) + Val(grdsales.TextMatrix(i, 24))
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    optnet.value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTRYEAR.Text = ""
    TXTUNIT.Text = ""
    
    lblretail.Caption = ""
    lblwsale.Caption = ""
    lblvan.Caption = ""
    lblunit.Caption = ""
    LblPack.Caption = ""
    lblcase.Caption = ""
    lblcrtnpack.Caption = ""
    lblunit.Caption = ""
    LblPack.Caption = ""
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Text = ""
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
    OptNormal.value = True
    OptNormal.value = True
    TXTDISC.Text = ""
    LBLSUBTOTAL.Caption = ""
    lblP_Rate.Caption = "0"
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    txtcategory.Enabled = True
    txtcategory.SetFocus
    'TXTSLNO.Enabled = True
    M_EDIT = False
    Call COSTCALCULATION
    
    'grdsales.TopRow = grdsales.Rows - 1
Exit Sub
Errhand:
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
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    If Not (UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICES" Or UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SELF") Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
                !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(LblPack.Caption), 3)
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) * Val(LblPack.Caption), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Caption), 3)
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
           
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Caption), 3)
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Caption), 3)
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
    Next i
    grdsales.Rows = grdsales.Rows - 1
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    LBLCOMAMT.Caption = ""
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
        LBLCOMAMT.Caption = Val(LBLCOMAMT.Caption) + Val(grdsales.TextMatrix(i, 24))
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
    
    Call COSTCALCULATION
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    optnet.value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTRYEAR.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTAPPENDQTY.Text = ""
    TXTFREEAPPEND.Text = ""
    txtappendcomm.Text = ""
    TXTAPPENDTOTAL.Text = ""
    txtretail.Text = ""
    txtBatch.Text = ""
    TXTTAX.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    OptNormal.value = True
    TxtMRP.Text = ""
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    txtcommi.Text = ""
    TxtRetailmode.Text = ""
    
    TXTDISC.Text = ""
    LBLSUBTOTAL.Caption = ""
    LBLDNORCN.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    If grdsales.Rows = 1 Then
'        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    
End Sub

Private Sub CmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            OptNormal.value = True
            optnet.value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTRYEAR.Text = ""
            TXTUNIT.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            TXTDISC.Enabled = False
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
        LBLITEMCOST.Visible = False
        LBLSELPRICE.Visible = False
    Else
        LBLPROFIT.Visible = True
        LBLTOTALCOST.Visible = True
        Label1(25).Visible = True
        Label1(26).Visible = True
        Label1(27).Visible = True
        Label1(28).Visible = True
        LBLITEMCOST.Visible = True
        LBLSELPRICE.Visible = True
    End If
End Sub

Private Sub CMDMODIFY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    If UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICE CHARGE" Then
        CMDMODIFY.Enabled = False
        CmdDelete.Enabled = False
        CMDEXIT.Enabled = False
        M_EDIT = True
        TXTRETAILNOTAX.Enabled = True
        TXTRETAILNOTAX.SetFocus
        Exit Sub
    End If
    
    On Error GoTo Errhand
    
    If Not (UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICES" Or UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SELF") Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                '!ISSUE_QTY = !ISSUE_QTY - ((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))))
                !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(LblPack.Caption), 3)
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) * Val(LblPack.Caption), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Caption), 3)
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
           
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Caption), 3)
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Caption), 3)
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
    FrmeType.Enabled = True
    TXTQTY.SetFocus
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            OptNormal.value = True
            optnet.value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTRYEAR.Text = ""
            TXTUNIT.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub CmdPrint_Click()
        
    'Exit Sub
    If grdsales.Rows = 1 Then Exit Sub
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
    Me.lblcredit.Caption = "0"
    Me.Generateprint
End Sub

Public Function Generateprint()
    'Exit Function
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    Dim CN As Integer
    Dim DN As Integer
    Dim b As Integer
    Dim Num As Currency
    
    On Error GoTo Errhand
    
    DN = 0
    CN = 0
    b = 0
        
    If OLD_BILL = False Then Call checklastbill
    OLD_BILL = True
    'db.Execute "delete From TEMPTRXFILE WHERE TRX_TYPE='DN' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From TEMPTRXFILE"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "DN"
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
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 10)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "Issued to     " & Trim(DataList2.Text)
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
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(i, 24))
        
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Call ReportGeneratION
    
    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)) - Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 2)), 2)
    Num = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0))
    LBLFOT.Tag = "(DHS " & Words_1_all(Num) & " Only)"

    ReportNameVar = Rptpath & "rptDelivery"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "( {TRXFILE.TRX_TYPE}='DN' AND {TRXFILE.VCH_NO}= " & Val(txtBillNo.Text) & " )"
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    Set oRs = New ADODB.Recordset
    Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" Then Report.Database.SetDataSource oRs, 3, 1
    Set oRs = Nothing
    
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        'If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & DataList2.Text & "'"
        If Trim(TxtBillName.Text) <> "" Then
            If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & TxtBillName.Text & "'"
        Else
            If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & DataList2.Text & "'"
        End If
        If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
        If CRXFormulaField.Name = "{@LPO}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
        'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
        If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
        If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
        If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
        If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
        If Trim(TXTTIN.Text) = "" Then
            If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'TAX INVOICE FORM 8B'"
        Else
            If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'TAX INVOICE FORM 8'"
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
    
    GoTo SKIP

    
    On Error GoTo CLOSEFILE
    Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
    End If
    On Error GoTo Errhand
    
    Print #1, "TYPE " & Rptpath & "Report.txt > PRN"
    Print #1, "EXIT"
    Close #1
    
    '//HERE write the proper path where your command.com file exist
    'Shell "C:\WINDOW\COMMAND.COM /C " & Rptpath & "REPO.BAT N", vbHide
    Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & Rptpath & "REPO.BAT N", vbHide
    
    cmdRefresh.SetFocus

'    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption), 2)) - Val(Round(Val(LBLTOTAL.Caption), 0)), 2)
SKIP:
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    TXTQTY.Enabled = False
    FrmeType.Enabled = False
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    txtretail.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTDISC.Enabled = False
    
    ''rptPRINT.Action = 1
    Exit Function
Errhand:
    MsgBox Err.Description
End Function

Private Sub CMDPRINT_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            OptNormal.value = True
            optnet.value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTRYEAR.Text = ""
            TXTUNIT.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub cmdRefresh_Click()
    If Delivered_flag = True Then
        MsgBox "Cannot Modify Since Already Appended with Bill No. " & BillNO, , "Delivery"
        CMDEXIT.Enabled = True
        CMDEXIT.SetFocus
        Exit Sub
    End If
            
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    End If
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
        DataList2.SetFocus
        Exit Sub
    End If
    Call AppendSale
    
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            OptNormal.value = True
            optnet.value = True
            TxtMRP.Text = ""
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTRYEAR.Text = ""
            TXTUNIT.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub Form_Activate()
    If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If txtcategory.Enabled = True Then txtcategory.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
    If txtretail.Enabled = True Then txtretail.SetFocus
    If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
    If TXTTAX.Enabled = True Then TXTTAX.SetFocus
    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
    If CMDPRINT.Enabled = True Then CMDPRINT.SetFocus
    If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
    If txtBillNo.Visible = True Then txtBillNo.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo Errhand
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TEMPCN WHERE TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
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
    Delivered_flag = False
    SERIAL_FLAG = False
    ACT_FLAG = True
    AGNT_FLAG = True
    lblcredit.Caption = "1"
    txtcrdays.Text = ""
    lblP_Rate.Caption = "0"
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 600
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 4800
    grdsales.ColWidth(3) = 1400
    grdsales.ColWidth(5) = 1400
    grdsales.ColWidth(7) = 1700
    grdsales.ColWidth(6) = 1700
    grdsales.ColWidth(8) = 1200
    grdsales.ColWidth(9) = 1200
    grdsales.ColWidth(12) = 1900
    grdsales.ColWidth(20) = 1100
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "UNIT"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "PTR"
    grdsales.TextArray(7) = "RATE"
    grdsales.TextArray(8) = "DISC %"
    grdsales.TextArray(9) = "TAX %"
    grdsales.TextArray(10) = "Serial No"
    grdsales.TextArray(11) = "COST"
    grdsales.TextArray(12) = "SUB TOTAL"
    grdsales.TextArray(13) = "ITEM CODE"
    grdsales.TextArray(14) = "Vch No"
    grdsales.TextArray(15) = "" '"Line No"
    grdsales.TextArray(16) = "" '"Trx Type"
    grdsales.TextArray(17) = "Tax Mode"
    grdsales.TextArray(18) = "MFGR"
    grdsales.TextArray(19) = "CN/DN"
    grdsales.TextArray(20) = "FREE"
    grdsales.TextArray(21) = "PTR"
    grdsales.TextArray(22) = "PTRWOTAX"
    grdsales.TextArray(24) = "Comm"
    grdsales.TextArray(32) = "" '"TrxYear"
    
    grdsales.ColWidth(4) = 0
    grdsales.ColWidth(10) = 0
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(13) = 0
    grdsales.ColWidth(14) = 0
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(16) = 0
    grdsales.ColWidth(32) = 0
    grdsales.ColWidth(17) = 1500 '0
    grdsales.ColWidth(18) = 0
    grdsales.ColWidth(19) = 0
    grdsales.ColWidth(21) = 0
    grdsales.ColWidth(22) = 0
    grdsales.ColWidth(23) = 0
    grdsales.ColWidth(24) = 1700
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
    
    LBLTOTAL.Caption = 0
    LBLCOMAMT.Caption = 0
    
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
    FrmeType.Enabled = False
    TxtMRP.Enabled = False
    
    txtretail.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    TXTDISC.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    TXTSLNO.Text = 1
    Call FILLCOMBO
    TXTSLNO.Enabled = False
    CLOSEALL = 1
    M_EDIT = False
'    Me.Width = 11700
'    Me.Height = 10185
    Me.Left = 0
    Me.Top = 0
    Exit Sub
Errhand:
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
    
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDPOPUP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTtax As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            SERIAL_FLAG = True
            txtBatch.Text = GRDPOPUP.Columns(0)
            TXTVCHNO.Text = GRDPOPUP.Columns(2)
            TXTLINENO.Text = GRDPOPUP.Columns(3)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(4)
            TXTRYEAR.Text = GRDPOPUP.Columns(7)
            'TXTUNIT.Text = GRDPOPUP.Columns(5)
            Set GRDPOPUP.DataSource = Nothing
            
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = True
            FrmeType.Enabled = True
            TXTQTY.SetFocus
            
            Call CONTINUE
            Exit Sub
            'TXTQTY.Text = GRDPOPUP.Columns(1)
            TxtMRP.Text = GRDPOPUP.Columns(3)
            Select Case cmbtype.ListIndex
                Case 0
                    txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(13)), "", GRDPOPUP.Columns(13))
                Case 1
                    txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
                Case 2
                    txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(20)), "", GRDPOPUP.Columns(20))
            End Select
            lblretail.Caption = IIf(IsNull(GRDPOPUP.Columns(13)), "", GRDPOPUP.Columns(13))
            lblwsale.Caption = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
            lblvan.Caption = IIf(IsNull(GRDPOPUP.Columns(20)), "", GRDPOPUP.Columns(20))
            lblcase.Caption = IIf(IsNull(GRDPOPUP.Columns(18)), "", GRDPOPUP.Columns(18))
            lblcrtnpack.Caption = IIf(IsNull(GRDPOPUP.Columns(17)), "", GRDPOPUP.Columns(17))
            
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
                            If (!Category = "GENERAL" And !Remarks = "1") Then
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
            RSTtax.Close
            Set RSTtax = Nothing
            
            TXTVCHNO.Text = GRDPOPUP.Columns(8)
            TXTLINENO.Text = GRDPOPUP.Columns(9)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(10)
            TXTRYEAR = GRDPOPUP.Columns(4)
            TXTUNIT.Text = GRDPOPUP.Columns(11)
                        
            Set GRDPOPUP.DataSource = Nothing
            
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = True
            FrmeType.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtappendcomm.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTRYEAR.Text = ""
            TXTUNIT.Text = ""
            
            Set GRDPOPUP.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
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
    
    On Error GoTo Errhand
    Select Case KeyCode
        Case vbKeyReturn
            NONSTOCKFLAG = False
            MINUSFLAG = False
            M_STOCK = Val(GRDPOPUPITEM.Columns(2))
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            txtcategory.Text = IIf(IsNull(PHY!Category), "", PHY!Category)
            If UCase(PHY!Category) = "SERVICE CHARGE" Then
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TXTPRODUCT.Enabled = False
                TXTITEMCODE.Enabled = False
                TXTRETAILNOTAX.Enabled = True
                TXTRETAILNOTAX.SetFocus
                Exit Sub
            End If
            i = 0
            If M_STOCK <= 0 Then
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
                        FrmeType.Enabled = False
                        TXTPRODUCT.SetFocus
                        Exit Sub
                    Else
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN", "DN"
                                Exit For
                        End Select
                        If SERIAL_FLAG = False Then
                            TXTSLNO.Text = i
                            TXTAPPENDQTY.Text = Val(grdsales.TextMatrix(i, 3))
                            TXTFREEAPPEND.Text = Val(grdsales.TextMatrix(i, 20))
                            txtappendcomm.Text = Val(grdsales.TextMatrix(i, 24))
                            Exit For
                        End If
                    End If
                End If
            Next i
            
            Set GRDPOPUPITEM.DataSource = Nothing
            If ITEM_FLAG = True Then
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE  From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE  From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
                FrmeType.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            End If
            
            Dim RSTBATCH As ADODB.Recordset
            Set RSTBATCH = New ADODB.Recordset
            RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND LENGTH(REF_NO)>0 AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
            If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                Call FILL_BATCHGRID
                RSTBATCH.Close
                Set RSTBATCH = Nothing
                Exit Sub
            End If
            Set RSTBATCH = Nothing
                
                'TXTQTY.Text = GRDPOPUPITEM.Columns(2)
            Select Case cmbtype.ListIndex
                Case 0 'RT
                    txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", GRDPOPUPITEM.Columns(6))
                Case 1 'WS
                    txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", GRDPOPUPITEM.Columns(12))
                Case 2 'VP
                    txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", GRDPOPUPITEM.Columns(13))
            End Select
            lblretail.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", GRDPOPUPITEM.Columns(6))
            lblwsale.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", GRDPOPUPITEM.Columns(12))
            lblvan.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", GRDPOPUPITEM.Columns(13))
            lblcase.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(11)), "", GRDPOPUPITEM.Columns(11))
            lblcrtnpack.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(10)), "", GRDPOPUPITEM.Columns(10))
            If GRDPOPUPITEM.Columns(15) = "" Or IsNull(GRDPOPUPITEM.Columns(15)) Or GRDPOPUPITEM.Columns(15) = "0" Then
                LblPack.Caption = "1"
            Else
                LblPack.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(15)), "1", GRDPOPUPITEM.Columns(15))
            End If
            'lblpack.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(15)), "1", GRDPOPUPITEM.Columns(15))
            lblunit.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(16)), "Nos", GRDPOPUPITEM.Columns(16))
            
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
            
            TXTSALETYPE.Text = "2"
'            optnet.Value = True
            TXTUNIT.Text = GRDPOPUPITEM.Columns(5)
                        
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = True
            FrmeType.Enabled = True
            TXTQTY.SetFocus
            
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTRYEAR.Text = ""
            TXTUNIT.Text = ""
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
            TXTPRODUCT.SetFocus
            
    End Select
    Exit Sub
Errhand:
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
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyReturn
            On Error Resume Next
            TXTPRODUCT.Text = grdtmp.Columns(1)
            TXTITEMCODE.Text = grdtmp.Columns(0)
            Call TxtItemcode_KeyDown(13, 0)
            
            Set grdtmp.DataSource = Nothing
            grdtmp.Visible = False
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                TXTQTY.Enabled = True
                FrmeType.Enabled = True
                TXTQTY.SetFocus
            End If
    End Select
End Sub

Private Sub OptLoose_Click()
    If Val(LblPack.Caption) > 1 Then
        If cmbtype.ListIndex = 2 Then
            txtretail.Text = Val(lblvan.Caption) / Val(LblPack.Caption)
        ElseIf cmbtype.ListIndex = 0 Then
            txtretail.Text = Val(lblretail.Caption) / Val(LblPack.Caption)
        Else
            txtretail.Text = Val(lblwsale.Caption) / Val(LblPack.Caption)
        End If
        Call TXTRETAIL_LostFocus
    End If
    On Error Resume Next
    TXTQTY.SetFocus
End Sub

Private Sub OptLoose_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub optnet_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OptNormal_Click()
    If Val(LblPack.Caption) > 1 Then
        If cmbtype.ListIndex = 2 Then
            txtretail.Text = Val(lblvan.Caption)
        ElseIf cmbtype.ListIndex = 0 Then
            txtretail.Text = Val(lblretail.Caption)
        Else
            txtretail.Text = Val(lblwsale.Caption)
        End If
        Call TXTRETAIL_LostFocus
    End If
    On Error Resume Next
    TXTQTY.SetFocus
End Sub

Private Sub OptNormal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub OPTTaxMRP_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OPTVAT_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtBatch.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            txtBatch.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
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
'            TXTSLNO.Enabled = True
'            TXTSLNO.SetFocus
            cmbtype.Enabled = True
            cmbtype.SetFocus
            
            'TXTTYPE.Enabled = True
            'TXTTYPE.SetFocus
        Case vbKeyEscape
            TxtBillName.Enabled = True
            TxtBillName.SetFocus
    End Select
End Sub

Private Sub TxtBillAddress_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
    cr_days = False
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TRXMAST As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    
    Dim i As Long
    Dim N As Integer
    Dim M As Integer

    On Error GoTo Errhand
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click

    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtBillNo.Text) = 0 Then Exit Sub
            grdsales.Rows = 1
            i = 0
            Set TRXFILE = New ADODB.Recordset
            TRXFILE.Open "Select * From TEMPCN WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            If TRXFILE.RecordCount = 0 Then
                OLD_BILL = False
            Else
                OLD_BILL = True
            End If
            Do Until TRXFILE.EOF
                i = i + 1
                TXTINVDATE.Text = Format(TRXFILE!VCH_DATE, "DD/MM/YYYY")
                grdsales.Rows = grdsales.Rows + 1
                grdsales.FixedRows = 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = TRXFILE!ITEM_CODE
                grdsales.TextMatrix(i, 2) = TRXFILE!ITEM_NAME
                grdsales.TextMatrix(i, 3) = TRXFILE!QTY
                grdsales.TextMatrix(i, 4) = Val(TRXFILE!UNIT)
                
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
                grdsales.TextMatrix(i, 8) = 0 'DISC
                grdsales.TextMatrix(i, 9) = Val(TRXFILE!SALES_TAX)
        
                grdsales.TextMatrix(i, 10) = TRXFILE!REF_NO
                grdsales.TextMatrix(i, 11) = IIf(IsNull(TRXFILE!ITEM_COST), 0, TRXFILE!ITEM_COST)
                grdsales.TextMatrix(i, 12) = Format(Val(TRXFILE!TRX_TOTAL), ".000")
                
                grdsales.TextMatrix(i, 13) = TRXFILE!ITEM_CODE
                grdsales.TextMatrix(i, 14) = Val(TRXFILE!R_VCH_NO)
                grdsales.TextMatrix(i, 15) = Val(TRXFILE!R_LINE_NO)
                grdsales.TextMatrix(i, 16) = Trim(TRXFILE!R_TRX_TYPE)
                grdsales.TextMatrix(i, 32) = Trim(TRXFILE!R_TRX_YEAR)
                grdsales.TextMatrix(i, 17) = IIf(IsNull(TRXFILE!CHECK_FLAG), "", Trim(TRXFILE!CHECK_FLAG))
                TXTDEALER.Text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                'DataList2.Text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                TXTINVDATE.Text = IIf(IsNull(TRXFILE!VCH_DATE), Date, Format(TRXFILE!VCH_DATE, "DD/MM/YYYY"))
                grdsales.TextMatrix(i, 19) = "DN"
                grdsales.TextMatrix(i, 20) = TRXFILE!FREE_QTY
                grdsales.TextMatrix(i, 21) = IIf(IsNull(TRXFILE!P_RETAIL), "0.00", Format(TRXFILE!P_RETAIL, ".0000"))
                grdsales.TextMatrix(i, 22) = IIf(IsNull(TRXFILE!P_RETAILWOTAX), "0.00", Format(TRXFILE!P_RETAILWOTAX, ".0000"))
                grdsales.TextMatrix(i, 23) = IIf(IsNull(TRXFILE!SALE_1_FLAG), "2", TRXFILE!SALE_1_FLAG)
                grdsales.TextMatrix(i, 24) = IIf(IsNull(TRXFILE!COM_AMT), "2", TRXFILE!COM_AMT)
                grdsales.TextMatrix(i, 25) = IIf(IsNull(TRXFILE!Category), 0, TRXFILE!Category)
                grdsales.TextMatrix(i, 26) = IIf(IsNull(TRXFILE!LOOSE_FLAG), "F", TRXFILE!LOOSE_FLAG)
                grdsales.TextMatrix(i, 27) = IIf(IsNull(TRXFILE!LOOSE_PACK), "1", TRXFILE!LOOSE_PACK)
                CMBDISTI.BoundText = IIf(IsNull(TRXFILE!AGENT_CODE), "", TRXFILE!AGENT_CODE)
                grdsales.TextMatrix(i, 28) = IIf(IsNull(TRXFILE!PACK_TYPE), "Pkt", TRXFILE!PACK_TYPE)
                grdsales.TextMatrix(i, 29) = IIf(IsNull(TRXFILE!CHECK_FLAG), "N", TRXFILE!CHECK_FLAG)
                TxtVehicle.Text = IIf(IsNull(TRXFILE!VEHICLE), "", TRXFILE!VEHICLE)
                If TRXFILE!CHECK_FLAG = "Y" Then
                    Delivered_flag = True
                    BillNO = IIf(IsNull(TRXFILE!BILL_NO), "", TRXFILE!BILL_NO)
                    grdsales.TextMatrix(i, 30) = IIf(IsNull(TRXFILE!BILL_NO), "", TRXFILE!BILL_NO)
                    grdsales.TextMatrix(i, 31) = IIf(IsNull(TRXFILE!BILL_TRX_TYPE), "", TRXFILE!BILL_TRX_TYPE)
                Else
                    Delivered_flag = False
                    BillNO = ""
                    grdsales.TextMatrix(i, 30) = ""
                    grdsales.TextMatrix(i, 31) = ""
                End If
                                
                If TRXFILE!BILL_TYPE = "R" Then
                    cmbtype.ListIndex = 0
                ElseIf TRXFILE!BILL_TYPE = "W" Then
                    cmbtype.ListIndex = 1
                End If
                TxtBillName.Text = IIf(IsNull(TRXFILE!BILL_NAME), "", TRXFILE!BILL_NAME)
                TRXFILE.MoveNext
        Loop
        TRXFILE.Close
        Set TRXFILE = Nothing
            
        LBLBILLNO.Caption = Val(txtBillNo.Text)
            
            
        LBLTOTAL.Caption = ""
        lblnetamount.Caption = ""
        LBLFOT.Caption = ""
        LBLCOMAMT.Caption = ""
        For i = 1 To grdsales.Rows - 1
            grdsales.TextMatrix(i, 0) = i
            LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
            If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
            LBLFOT.Caption = ""
            LBLCOMAMT.Caption = Val(LBLCOMAMT.Caption) + Val(grdsales.TextMatrix(i, 24))
        Next i
            'lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
            
            Call COSTCALCULATION
            
            
            TXTSLNO.Text = grdsales.Rows
            txtBillNo.Visible = False
            TXTSLNO.Enabled = True
            
            If grdsales.Rows > 1 Then
                'TXTSLNO.SetFocus
                TXTDEALER.SetFocus
            Else
                TXTDEALER.SetFocus
'                TXTSLNO.Enabled = False
'                TXTDEALER.Text = ""
'                TXTDEALER.SetFocus
            End If
    
    End Select

    Exit Sub
Errhand:
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

    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From TEMPCN WHERE TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        If Val(txtBillNo.Text) > i Then
            MsgBox "The last bill No. is " & i, vbCritical, "BILL..."
            txtBillNo.Visible = True
            txtBillNo.SetFocus
            Exit Sub
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
      
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MIN(VCH_NO) FROM TRXFILE WHERE TRX_TYPE = 'DN'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
        If Val(txtBillNo.Text) < i Then
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

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.Text)
    SERIAL_FLAG = False
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
Private Sub txtcrdays_GotFocus()
    txtcrdays.SelStart = 0
    txtcrdays.SelLength = Len(txtcrdays.Text)
End Sub

Private Sub txtcrdays_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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
    On Error GoTo Errhand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_CODE <>'130000' AND ACT_CODE <>'130001' And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_CODE <>'130000' AND ACT_CODE <>'130001' And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
Errhand:
    MsgBox Err.Description
    
End Sub

Private Sub TXTDISC_GotFocus()
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtBatch.Enabled = True
            TXTDISC.Enabled = False
            txtBatch.SetFocus
        Case vbKeyEscape
            txtretail.Enabled = True
            TXTDISC.Enabled = False
            txtretail.SetFocus
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
        LBLSUBTOTAL.Caption = Format(Round(Val(txtretail.Text), 3) - Val(TXTDISC.Tag), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.Text) * Val(txtretail.Text) * Val(TXTDISC.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(txtretail.Text), 3)) - Val(TXTDISC.Tag), ".000")
    End If
    
    ''TXTDISC.Text = Format(TXTDISC.Text, ".000")

End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
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
            TxtMRP.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
        Case vbKeyEscape
            TxtMRP.Enabled = False
            TXTFREE.Enabled = True
            TXTFREE.SetFocus
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

Private Sub TxtPhone_GotFocus()
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TxtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            CMBDISTI.Enabled = True
            CMBDISTI.SetFocus
            'TXTTYPE.Enabled = True
            'TXTTYPE.SetFocus
        Case vbKeyEscape
            TxtVehicle.Enabled = True
            TxtVehicle.SetFocus
    End Select
End Sub

Private Sub TxtPhone_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
    End Select
End Sub


Private Sub TXTPRODUCT_Change()
    On Error GoTo Errhand
    
    Set grdtmp.DataSource = Nothing
    If PHYFLAG = True Then
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE  From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '%" & Me.txtcategory.Text & "%' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    Else
        PHY.Close
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE  From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '%" & Me.txtcategory.Text & "%' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    End If
    Set grdtmp.DataSource = PHY
    
    If PHY.RecordCount > 0 Then
        grdtmp.Visible = True
    Else
        grdtmp.Visible = False
        Exit Sub
    End If
    grdtmp.Columns(0).Visible = False
    grdtmp.Columns(1).Caption = "ITEM NAME"
    grdtmp.Columns(1).Width = 3800
    grdtmp.Columns(2).Caption = "QTY"
    grdtmp.Columns(2).Width = 1200
    grdtmp.Columns(3).Width = 0
    grdtmp.Columns(4).Width = 0
    grdtmp.Columns(5).Width = 0
    grdtmp.Columns(6).Caption = "RT"
    grdtmp.Columns(6).Width = 1220
    grdtmp.Columns(7).Width = 0
    grdtmp.Columns(8).Width = 0
    grdtmp.Columns(9).Width = 0
    grdtmp.Columns(10).Width = 0
    grdtmp.Columns(11).Width = 0
    grdtmp.Columns(12).Width = 1200
    grdtmp.Columns(12).Caption = "WS"
    grdtmp.Columns(13).Width = 0
    grdtmp.Columns(14).Width = 0
    grdtmp.Columns(15).Width = 0
    grdtmp.Columns(16).Width = 0
    grdtmp.Columns(17).Width = 0
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_GotFocus()
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
    Call TXTPRODUCT_Change
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset
    
    On Error GoTo Errhand
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
        Case vbKeyReturn
            On Error Resume Next
            TXTITEMCODE.Text = ""
            TXTITEMCODE.Text = grdtmp.Columns(0)
            If Trim(TXTITEMCODE.Text) = "" Then
                MsgBox "Item not found!!!!", , "EzBiz"
                Exit Sub
            Else
                Call TxtItemcode_KeyDown(13, 0)
            End If
            Exit Sub
            
            M_STOCK = 0
            'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If Trim(TXTPRODUCT.Text) = "" Then
                TXTPRODUCT.Enabled = False
                txtcategory.Enabled = True
                txtcategory.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False
            TXTQTY.Text = ""
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            OptNormal.value = True
            optnet.value = True
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
            
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE  From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE  From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            Set grdtmp.DataSource = PHY
            
            If PHY.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "Sales"
                Exit Sub
            End If
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                
                Set RSTBATCH = New ADODB.Recordset
                RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND LENGTH(REF_NO)>0 AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                    Call FILL_BATCHGRID
                    RSTBATCH.Close
                    Set RSTBATCH = Nothing
                    Exit Sub
                End If
                Set RSTBATCH = Nothing
                Call CONTINUE
            Else
                Call FILL_ITEMGRID
                Exit Sub
            End If
JUMPNONSTOCK:
                TXTPRODUCT.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                FrmeType.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
        Case vbKeyEscape
            txtcategory.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
            TXTTAX.Enabled = False
            TXTDISC.Enabled = False
            txtcategory.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub
Private Function CONTINUE()
    Dim i As Long
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                            Exit Function
                        Else
                            Select Case grdsales.TextMatrix(i, 19)
                                Case "CN", "DN"
                                    Exit For
                            End Select
                            If SERIAL_FLAG = False Then
                                TXTSLNO.Text = i
                                TXTAPPENDQTY.Text = Val(grdsales.TextMatrix(i, 3))
                                TXTFREEAPPEND.Text = Val(grdsales.TextMatrix(i, 20))
                                txtappendcomm.Text = Val(grdsales.TextMatrix(i, 24))
                                Exit For
                            End If
                        End If
                    End If
                Next i
                txtcategory.Text = IIf(IsNull(PHY!Category), "", PHY!Category)
                If UCase(PHY!Category) = "SERVICE CHARGE" Then
                    TXTTAX.Enabled = True
                    TXTTAX.SetFocus
                    Exit Function
                End If
                Select Case cmbtype.ListIndex
                    Case 0 'RT
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(6)), "", grdtmp.Columns(6))
                    Case 1 'WS
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(12)), "", grdtmp.Columns(12))
                    Case 2 'VAN
                        txtretail.Text = IIf(IsNull(grdtmp.Columns(13)), "", grdtmp.Columns(13))
                End Select
                lblretail.Caption = IIf(IsNull(grdtmp.Columns(6)), "", grdtmp.Columns(6))
                lblwsale.Caption = IIf(IsNull(grdtmp.Columns(12)), "", grdtmp.Columns(12))
                lblvan.Caption = IIf(IsNull(grdtmp.Columns(13)), "", grdtmp.Columns(13))
                lblcase.Caption = IIf(IsNull(grdtmp.Columns(11)), "", grdtmp.Columns(11))
                lblcrtnpack.Caption = IIf(IsNull(grdtmp.Columns(10)), "", grdtmp.Columns(10))
                If grdtmp.Columns(16) = "" Or IsNull(grdtmp.Columns(16)) Or grdtmp.Columns(16) = "0" Then
                    LblPack.Caption = "1"
                Else
                    LblPack.Caption = IIf(IsNull(grdtmp.Columns(16)), "1", grdtmp.Columns(16))
                End If
                lblunit.Caption = IIf(IsNull(grdtmp.Columns(17)), "Nos", grdtmp.Columns(17))
                
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
                TXTUNIT.Text = grdtmp.Columns(5)
                                   
                'TXTPRODUCT.Enabled = False
                'TXTQTY.Enabled = True
                'FrmeType.Enabled = True
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

Private Sub TXTQTY_GotFocus()
    TXTFREE.Enabled = True
    Dim RSTITEMCOST As ADODB.Recordset
    
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
    TXTQTY.Tag = Trim(TXTPRODUCT.Text)
    On Error GoTo Errhand
    
    Set RSTITEMCOST = New ADODB.Recordset
    RSTITEMCOST.Open "SELECT ITEM_COST, SALES_PRICE FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMCOST.EOF Or RSTITEMCOST.BOF) Then
        LBLITEMCOST.Caption = IIf(IsNull(RSTITEMCOST!ITEM_COST), "", RSTITEMCOST!ITEM_COST)
        LBLSELPRICE.Caption = IIf(IsNull(RSTITEMCOST!SALES_PRICE), "", RSTITEMCOST!SALES_PRICE)
    End If
    RSTITEMCOST.Close
    Set RSTITEMCOST = Nothing
    
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Double
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If Not (UCase(txtcategory.Text) = "SERVICES" Or UCase(txtcategory.Text) = "SELF") Then
                i = 0
                
                If Val(LblPack.Caption) = 0 Then LblPack.Caption = 1
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT CLOSE_QTY  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                    If (IsNull(RSTTRXFILE!CLOSE_QTY)) Then RSTTRXFILE!CLOSE_QTY = 0
                    i = RSTTRXFILE!CLOSE_QTY / Val(LblPack.Caption)
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
                'If Val(TXTQTY.Text) = 0 Then Exit Sub
                If i <> 0 Then
                    If Val(TXTQTY.Text) > i Then
                        If SERIAL_FLAG = True Then
                            MsgBox "AVAILABLE STOCK IS  " & i & " ", , "SALES"
                            TXTQTY.SelStart = 0
                            TXTQTY.SelLength = Len(TXTQTY.Text)
                            Exit Sub
                        End If
                        If (MsgBox("AVAILABLE STOCK IS  " & i & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                            'MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                            TXTQTY.SelStart = 0
                            TXTQTY.SelLength = Len(TXTQTY.Text)
                            Exit Sub
                        End If
                    End If
                End If
            End If
SKIP:
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTRYEAR.Text = ""
            TXTUNIT.Text = ""
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
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
    TXTQTY.Text = Format(TXTQTY.Text, ".000")
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    If Val(TXTRETAILNOTAX.Text) = 0 Then
        TXTDISC.Tag = Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(txtretail.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    End If
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
'            If Delivered_flag = True Then
'                MsgBox "Cannot Modify Since Already Appended with Bill No. " & BillNO, , "Delivery"
'                TXTSLNO.SetFocus
'                Exit Sub
'            End If
            
            If Val(TXTSLNO.Text) = 0 Then
                SERIAL_FLAG = False
                TXTSLNO.Text = ""
                TXTPRODUCT.Text = ""
                TXTQTY.Text = ""
                TXTAPPENDQTY.Text = ""
                TXTFREEAPPEND.Text = ""
                txtappendcomm.Text = ""
                TXTAPPENDTOTAL.Text = ""
                TXTFREE.Text = ""
                OptNormal.value = True
                optnet.value = True
                TxtMRP.Text = ""
                
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTRYEAR.Text = ""
                TXTUNIT.Text = ""
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
                If grdsales.TextMatrix(Val(TXTSLNO.Text), 29) = "Y" Then
                    MsgBox "Cannot Modify Since Already Appended with Bill No. " & grdsales.TextMatrix(Val(TXTSLNO.Text), 30), , "Delivery"
                    TXTSLNO.SetFocus
                    Exit Sub
                End If
                lblP_Rate.Caption = "1"
                TXTSLNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 0)
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 3)
                TXTFREE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 20)
                TxtMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), ".000")
                
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 15)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 16)
                TXTRYEAR.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 32)
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
                
                TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                txtretail.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                TXTSALETYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 23)
                txtcategory.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 25)
                If UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICE CHARGE" Then
                    txtretaildummy.Text = 0
                    txtcommi.Text = 0 'Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) / Val(TXTQTY.Text), 2)
                Else
                    txtretaildummy.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) / Val(TXTQTY.Text), 2)
                    txtcommi.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) / Val(TXTQTY.Text), 2)
                End If
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                FrmeType.Enabled = False
                TXTTAX.Enabled = False
                TXTFREE.Enabled = False
                txtretail.Enabled = False
                TXTRETAILNOTAX.Enabled = False
                TXTDISC.Enabled = False
                TxtMRP.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Select Case grdsales.TextMatrix(Val(TXTSLNO.Text), 26)
                    Case "L"
                        OptLoose.value = True
                    Case Else
                        OptNormal.value = True
                End Select
                LBLDNORCN.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 19)
                LblPack.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 27))
                grdtmp.Visible = False
                Exit Sub
            End If
SKIP:
            'grdtmp.Visible = False
            lblP_Rate.Caption = "0"
            TXTSLNO.Enabled = False
            txtcategory.Enabled = True
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
            TXTTAX.Enabled = False
            TXTDISC.Enabled = False
            txtcategory.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                optnet.value = True
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTRYEAR.Text = ""
                TXTUNIT.Text = ""
                TXTQTY.Text = ""
                TXTAPPENDQTY.Text = ""
                TXTAPPENDTOTAL.Text = ""
                TXTFREEAPPEND.Text = ""
                txtappendcomm.Text = ""
                txtretail.Text = ""
                txtBatch.Text = ""
                TXTTAX.Text = ""
                TXTRETAILNOTAX.Text = ""
                TXTSALETYPE.Text = ""
                TXTFREE.Text = ""
                OptNormal.value = True
                TxtMRP.Text = ""
                
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                lblP_Rate.Caption = "0"
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.Rows > 1 Then
                TXTSLNO.Enabled = False
                CMDPRINT.Enabled = True
                cmdRefresh.Enabled = True
                CMDPRINT.SetFocus
            Else
                TXTSLNO.Enabled = False
                FRMEHEAD.Enabled = True
                TXTDEALER.Enabled = True
                TXTDEALER.SetFocus
            End If
            LBLDNORCN.Caption = ""
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

Private Sub TxtTax_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TxtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRETAILNOTAX.Enabled = True
            TXTTAX.Enabled = False
            TXTRETAILNOTAX.SetFocus
        Case vbKeyEscape
            TXTQTY.Enabled = True
            TXTTAX.Enabled = False
            TXTQTY.SetFocus
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
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
    GRDPOPUPITEM.Columns(3).Width = 1220
    GRDPOPUPITEM.Columns(4).Caption = "WS"
    GRDPOPUPITEM.Columns(4).Width = 1220
    GRDPOPUPITEM.Columns(5).Caption = "VAN"
    GRDPOPUPITEM.Columns(5).Width = 1220
    GRDPOPUPITEM.Columns(6).Caption = "CRTN"
    GRDPOPUPITEM.Columns(6).Width = 1220
    GRDPOPUPITEM.SetFocus
End Function

Private Function STOCKADJUST()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    M_STOCK = 0
    On Error GoTo Errhand
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
    
Errhand:
    MsgBox Err.Description
End Function


Private Function COSTCALCULATION()
    Dim RSTCOST As ADODB.Recordset
    Dim COST As Double
    Dim N As Integer
    'Dim RSTITEMMAST As ADODB.Recordset
    
     LBLTOTALCOST.Caption = ""
     LBLPROFIT.Caption = ""
        COST = 0
    On Error GoTo Errhand
    For N = 1 To grdsales.Rows - 1
        Set RSTCOST = New ADODB.Recordset
        RSTCOST.Open "SELECT ITEM_COST FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(N, 1)) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTCOST.EOF
            If OptLoose.value = True And Val(LblPack.Caption) > 1 Then
                If Not IsNull(RSTCOST!ITEM_COST) Then COST = COST + (RSTCOST!ITEM_COST) * (Val(grdsales.TextMatrix(N, 3) * Val(LblPack.Caption)))
            Else
                If Not IsNull(RSTCOST!ITEM_COST) Then COST = COST + (RSTCOST!ITEM_COST) * Val(grdsales.TextMatrix(N, 3))
            End If
            RSTCOST.MoveNext
        Loop
        RSTCOST.Close
        Set RSTCOST = Nothing
    Next N
    
    LBLTOTALCOST.Caption = Round(COST, 2)
    LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - COST, 2)

    Exit Function
    
Errhand:
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
    On Error GoTo Errhand
    
    If OLD_BILL = False Then Call checklastbill
    'db.Execute "delete From TRXMAST WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    'db.Execute "delete From TRXSUB WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From TEMPCN WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From TRXFILE WHERE TRX_TYPE='DN' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    'DB.Execute "delete From P_Rate WHERE TRX_TYPE='SI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    
    If grdsales.Rows = 1 Then GoTo SKIP
    
    E_DATE = Format(TXTINVDATE.Text, "MM/DD/YYYY")
    If Day(E_DATE) <= 12 Then
        DAY_DATE = Format(Month(E_DATE), "00")
        MONTH_DATE = Format(Day(E_DATE), "00")
        YEAR_DATE = Format(Year(E_DATE), "0000")
        E_DATE = DAY_DATE & "/" & MONTH_DATE & "/" & YEAR_DATE
    End If
    E_DATE = Format(E_DATE, "MM/DD/YYYY")
    
    TRXVALUE = 0
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPCN", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "SI"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!line_no = i
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
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
        If CMBDISTI.BoundText <> "" Then
            RSTTRXFILE!COM_FLAG = "Y"
        Else
            RSTTRXFILE!COM_FLAG = "N"
        End If
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "Delivered to  " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 10)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
        RSTTRXFILE!CST = 1
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!VEHICLE = Trim(TxtVehicle.Text)
        If Trim(TxtBillName.Text) = "" Then
            RSTTRXFILE!BILL_NAME = DataList2.Text
        Else
            RSTTRXFILE!BILL_NAME = Trim(TxtBillName.Text)
        End If
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!Area = RSTITEMMAST!Area
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
            
        RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(i, 14) = "", 0, grdsales.TextMatrix(i, 14))
        RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(i, 15) = "", 0, grdsales.TextMatrix(i, 15))
        RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(i, 16) = "", "MI", grdsales.TextMatrix(i, 16))
        RSTTRXFILE!R_TRX_YEAR = IIf(grdsales.TextMatrix(i, 32) = "", "", grdsales.TextMatrix(i, 32))
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.Text
            
        If CMBDISTI.BoundText <> "" Then
            RSTTRXFILE!AGENT_CODE = CMBDISTI.BoundText
            RSTTRXFILE!AGENT_NAME = CMBDISTI.Text
            RSTTRXFILE!COMM_AMT = Val(LBLCOMAMT.Caption)
        Else
            RSTTRXFILE!AGENT_CODE = ""
            RSTTRXFILE!AGENT_NAME = ""
        End If
        If cmbtype.ListIndex = 0 Then
            RSTTRXFILE!BILL_TYPE = "R"
        ElseIf cmbtype.ListIndex = 1 Then
            RSTTRXFILE!BILL_TYPE = "W"
        End If
                    
        RSTTRXFILE!CHECK_FLAG = "N"
        RSTTRXFILE!BILL_NO = 0 'Null
        RSTTRXFILE!BILL_DATE = Null
            
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "DN"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!line_no = i
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
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
        If CMBDISTI.BoundText <> "" Then
            RSTTRXFILE!COM_FLAG = "Y"
        Else
            RSTTRXFILE!COM_FLAG = "N"
        End If
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "Delivered to  " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 10)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
        RSTTRXFILE!CST = 1
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        If CMBDISTI.BoundText <> "" Then
            RSTTRXFILE!COM_FLAG = "Y"
        Else
            RSTTRXFILE!COM_FLAG = "N"
        End If
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(i, 26)
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 27))
        'RSTTRXFILE!VEHICLE = Trim(TxtVehicle.Text)
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!Area = RSTITEMMAST!Area
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
SKIP:
    i = 0
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TEMPCN WHERE TRX_TYPE = 'SI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
        
    OLD_BILL = False
    Delivered_flag = False
    BillNO = ""
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    lbladdress.Caption = ""
    LBLDNORCN.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    LBLPROFIT.Caption = ""
    LBLDATE.Caption = Date
    LBLTOTAL.Caption = ""
    LBLCOMAMT.Caption = ""
    LBLTOTALCOST.Caption = ""
    LBLDISCAMT.Caption = ""
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    M_EDIT = False
    cmdRefresh.Enabled = False
    CMDEXIT.Enabled = True
    CMDPRINT.Enabled = False
    CMDEXIT.Enabled = True
    TXTSLNO.Enabled = False
    FRMEHEAD.Enabled = True
    TXTDEALER.Enabled = True
    TXTDEALER.SetFocus
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Tag = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblcredit.Caption = "0"
    txtcrdays.Text = ""
    TxtVehicle.Text = ""
    TxtBillName.Text = ""
    CMBDISTI.Text = ""
    cmbtype.ListIndex = -1
    Exit Function
Errhand:
    MsgBox Err.Description
End Function

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
            
            If Val(TXTFREE.Text) = 0 Then GoTo SKIP
            If Not (UCase(txtcategory.Text) = "SERVICES" Or UCase(txtcategory.Text) = "SELF") Then
                i = 0
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT CLOSE_QTY  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                    If (IsNull(RSTTRXFILE!CLOSE_QTY)) Then RSTTRXFILE!CLOSE_QTY = 0
                    i = RSTTRXFILE!CLOSE_QTY / Val(LblPack.Caption)
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
                If i > 0 Then
                    If Val(TXTFREE.Text) + Val(TXTQTY.Text) > i Then
                        MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                        TXTFREE.SelStart = 0
                        TXTFREE.SelLength = Len(TXTFREE.Text)
                        Exit Sub
                    End If
                End If
            End If
SKIP:
            TXTFREE.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
         Case vbKeyEscape
            TXTFREE.Enabled = False
            TXTQTY.Enabled = True
            FrmeType.Enabled = True
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
            If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2) + Val(LBLFOT.Caption)
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
    Open Rptpath & "Report.txt" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If Err.Number = 55 Then
        Close #1
        Open Rptpath & "Report.txt" For Output As #1 '//Report file Creation
    End If
    On Error GoTo Errhand
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold


    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) & _
              Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & AlignLeft(RSTCOMPANY!Address, 50)
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(8) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(7) & "Phone: " & RSTCOMPANY!TEL_NO & ", " & RSTCOMPANY!FAX_NO
        Print #1, Space(7) & "Trn No. " & RSTCOMPANY!KGST
        Print #1, Space(7) & RepeatString("-", 84)
        'Print #1,
        '''Print #1, Space(7) & "TIN No. " & RSTCOMPANY!KGST
        
        Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & "The KVAT Rules 2005"
        If Trim(TXTTIN.Text) <> "" Then
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(31) & "FORM NO. 8 See rule 58(10), TAX INVOICE"
        Else
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(29) & "FORM NO. 8B See rule 58(10), RETAIL INVOICE"
        End If
        Print #1, Chr(27) & Chr(71) & Chr(10) & Space(43) & AlignLeft("CASH / CREDIT SALE", 25)
        Print #1, Space(7) & RepeatString("-", 84)
        Print #1, Chr(27) & Chr(71) & Chr(10) & Space(7) & "D.N. NO & Date" & Space(6) & "P.O. NO. & Date" & Space(6) & "D.Doc.NO & Date" & Space(6) & "Del Terms" & Space(6) & "Veh. No"
        Print #1,
        Print #1, Space(7) & RepeatString("-", 84)
        'Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & AlignLeft("INVOICE FORM 8H", 16)
    
        'If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, LBLDATE.Caption)
        Print #1, Space(7) & "Bill No. " & Trim(LBLBILLNO.Caption) & Chr(27) & Chr(72) & Space(16) & AlignRight("Date:" & TXTINVDATE.Text, 57) '& Space(2) & LBLTIME.Caption
        Print #1, Chr(27) & Chr(71) & Chr(10) & Space(7) & "TO: " & TxtBillName.Text
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(12) & TxtBillAddress.Text
        If Trim(TxtPhone.Text) <> "" Then Print #1, Chr(27) & Chr(67) & Chr(0) & Space(12) & "Phone: " & TxtPhone.Text
        If Trim(TXTTIN.Text) <> "" Then Print #1, Chr(27) & Chr(67) & Chr(0) & Space(12) & "Trn No. " & TXTTIN.Text
        'LBLDATE.Caption = Date
    
       ' Print #1, Chr(27) & Chr(72) & Space(7) & "Salesman: CS"
    
        Print #1, Space(7) & RepeatString("-", 84)
        Print #1, Space(7) & AlignLeft("Description", 22) & _
                AlignLeft("Comm Code", 9) & Space(3) & _
                AlignLeft("Qty", 4) & Space(2) & _
                AlignLeft("Rate", 8) & Space(2) & _
                AlignLeft("Tax", 5) & Space(2) & _
                AlignLeft("Tax Amt", 7) & Space(2) & _
                AlignLeft("Net Rate", 10) & Space(2) & _
                AlignLeft("Amount", 12) & _
                Chr(27) & Chr(72)  '//Bold Ends
    
        Print #1, Space(7) & RepeatString("-", 84)
    
        For i = 1 To grdsales.Rows - 1
            Print #1, Space(7) & AlignLeft(grdsales.TextMatrix(i, 2), 22) & Space(9) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 5) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)), 2), "0.00"), 10) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 9)), 2), "0.00"), 8) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 9)) / 100, 2), "0.00"), 8) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 10) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 12) & _
                Chr(27) & Chr(72)  '//Bold Ends
        Next i
    
        Print #1, Space(7) & AlignRight("-------------", 84)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("BILL AMOUNT ", 68) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
            Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("DISC AMOUNT ", 68) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("BILL AMOUNT ", 68) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
        End If
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 57) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Round off ", 68) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
        Print #1, Chr(13)
        Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(25) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Round(Val(lblnetamount.Caption), 0), "####.00")), 9)
        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
        Num = CCur(Round(LBLTOTAL.Caption, 0))
        Print #1, Chr(27) & Chr(72) & Space(7) & AlignLeft("(DHS " & Words_1_all(Num) & ")", 80)
        Print #1, Space(7) & RepeatString("-", 84)
        Print #1, Chr(27) & Chr(67) & Chr(0)
        If Trim(TXTTIN.Text) <> "" Then
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & "Certified that all the particulars shown in the above Tax Invoice are true and correct"
            Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & "and that my/our Registration under KVAT ACT 2003 is valid as on the date of this bill"
            Print #1, Space(7) & RepeatString("-", 84)
        End If
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & "Thank You... E.&.O.E SUBJECT TO CHERHALA JURISDICTION"
        Print #1, Chr(27) & Chr(67) & Chr(0) & Space(67) & "For ELECTROCRAFTS"
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

Errhand:
    MsgBox Err.Description
End Function

Private Sub TXTRETAIL_GotFocus()
    txtretail.SelStart = 0
    txtretail.SelLength = Len(txtretail.Text)
    Call FILL_PREVIIOUSRATE
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtretail.Text) = 0 Then Exit Sub
            txtretail.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyEscape
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = True
            TXTRETAILNOTAX.SetFocus
        Case 116
            Call FILL_PREVIIOUSRATE
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
            txtcommi.Text = Format(Round(Val(txtretaildummy.Text) * Val(TXTQTY.Text), 2), "0.00")
        Else
            txtcommi.Text = Format(Round((Val(TXTRETAILNOTAX.Text) * Val(txtretaildummy.Text) / 100) * Val(TXTQTY.Text), 2), "0.00")
        End If
    End If
End Sub

Private Sub TXTRETAILNOTAX_GotFocus()
    TXTRETAILNOTAX.SelStart = 0
    TXTRETAILNOTAX.SelLength = Len(TXTRETAILNOTAX.Text)
End Sub

Private Sub TXTRETAILNOTAX_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTRETAILNOTAX.Text) = 0 Then Exit Sub
            TXTRETAILNOTAX.Enabled = False
            txtretail.Enabled = True
            txtretail.SetFocus
        Case vbKeyEscape
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
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

    If OPTVAT.value = False Then TXTTAX.Text = 0
    TXTRETAILNOTAX.Text = Round(Val(txtretail.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
    TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.0000")
    txtretail.Text = Format(Val(txtretail.Text), "0.0000")
    
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
    On Error GoTo Errhand
    
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

Errhand:
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
            FRMEHEAD.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TxtPhone.Enabled = True
            TxtPhone.SetFocus
    End Select
End Sub

Private Sub CMBDISTI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcommi_GotFocus()
    txtcommi.SelStart = 0
    txtcommi.SelLength = Len(txtcommi.Text)
End Sub

Private Sub txtcommi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.Enabled = True
            txtcommi.Enabled = False
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTDISC.Enabled = True
            txtcommi.Enabled = False
            TXTDISC.SetFocus
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

Private Sub TXTTYPE_GotFocus()
    TXTTYPE.SelStart = 0
    TXTTYPE.SelLength = Len(TXTTYPE.Text)
End Sub

Private Sub TXTTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Val(TXTTYPE.Text) = 0 Or Val(TXTTYPE.Text) > 3 Then
                MsgBox "Enter the Type Code", vbOKOnly, "Sale Bill..."
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

Private Sub TXTTYPE_LostFocus()
     If Val(TXTTYPE.Text) = 0 Or Val(TXTTYPE.Text) > 3 Then
        MsgBox "Enter Bill Type", vbOKOnly, "Sales"
        TXTTYPE.SetFocus
        Exit Sub
    End If
    cmbtype.SetFocus
End Sub

Private Sub TXTAREA_GotFocus()
    TXTAREA.SelStart = 0
    TXTAREA.SelLength = Len(TXTAREA.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
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
            'TXTSLNO.Enabled = True
            'TXTSLNO.SetFocus
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
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILE  WHERE (TRX_TYPE='GI' OR TRX_TYPE ='SI' OR TRX_TYPE ='RI' OR TRX_TYPE ='DN') AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY VCH_DATE DESC LIMIT 10", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILE  WHERE (TRX_TYPE='GI' OR TRX_TYPE ='SI' OR TRX_TYPE ='RI' OR TRX_TYPE ='DN') AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY VCH_DATE DESC LIMIT 10", db, adOpenStatic, adLockReadOnly
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
        GRDPRERATE.Columns(5).Caption = "CUSTOMER"
        GRDPRERATE.Columns(6).Caption = "INV NO"
    
        GRDPRERATE.Columns(0).Visible = False
        GRDPRERATE.Columns(1).Width = 3500
        GRDPRERATE.Columns(2).Width = 1300
        GRDPRERATE.Columns(3).Width = 1200
        GRDPRERATE.Columns(4).Width = 1500
        GRDPRERATE.Columns(5).Visible = False
        GRDPRERATE.Columns(6).Width = 1400
        
        
        'GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(7).Text
    End If
    
End Function

Private Sub TxtItemcode_GotFocus()
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
    SERIAL_FLAG = False
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset
    
    On Error GoTo Errhand
    Select Case KeyCode
        Case vbKeyReturn
            M_STOCK = 0
            'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If Trim(TXTITEMCODE.Text) = "" Then
                TXTPRODUCT.Enabled = True
                TXTITEMCODE.Enabled = False
                TXTPRODUCT.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False
            TXTQTY.Text = ""
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            OptNormal.value = True
            optnet.value = True
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
            
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE  From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.Text & "' ", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE  From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.Text & "' ", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            Set grdtmp.DataSource = PHY
            
            If PHY.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "Sales"
                Exit Sub
            End If
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                
                Set RSTBATCH = New ADODB.Recordset
                RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND LENGTH(REF_NO)>0 AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
                    Call FILL_BATCHGRID
                    RSTBATCH.Close
                    Set RSTBATCH = Nothing
                    Exit Sub
                End If
                Set RSTBATCH = Nothing
                Call CONTINUE
            Else
                Call FILL_ITEMGRID
                Exit Sub
            End If
JUMPNONSTOCK:
                TXTITEMCODE.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                FrmeType.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            FrmeType.Enabled = False
            TXTTAX.Enabled = False
            TXTDISC.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
Errhand:
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
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, R_TRX_YEAR From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, R_TRX_YEAR From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "Serial No."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "VCH No"
    GRDPOPUP.Columns(3).Caption = "Line No"
    GRDPOPUP.Columns(4).Caption = "Trx Type"
    
    GRDPOPUP.Columns(0).Width = 4100
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 0
    GRDPOPUP.Columns(3).Width = 0
    GRDPOPUP.Columns(4).Width = 0
    GRDPOPUP.Columns(5).Width = 0
    GRDPOPUP.Columns(6).Width = 0
    
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).Text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
       
End Function

Private Function checklastbill()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo Errhand
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TEMPCN WHERE TRX_TYPE = 'SI'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
Exit Function
Errhand:
    MsgBox Err.Description
End Function

Private Sub TxtVehicle_GotFocus()
    TxtVehicle.SelStart = 0
    TxtVehicle.SelLength = Len(TxtVehicle.Text)
    cr_days = False
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TxtVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            TxtPhone.Enabled = True
            TxtPhone.SetFocus
            'TXTTYPE.Enabled = True
            'TXTTYPE.SetFocus
        Case vbKeyEscape
            TxtBillName.Enabled = True
            TxtBillName.SetFocus
    End Select
End Sub

Private Sub TxtVehicle_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
    End Select
End Sub



