VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMQUOTATION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QUOTATION"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18495
   Icon            =   "Frmquotn.frx":0000
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
      Height          =   2760
      Left            =   4365
      TabIndex        =   21
      Top             =   3885
      Visible         =   0   'False
      Width           =   9555
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   2370
         Left            =   15
         TabIndex        =   22
         Top             =   390
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   4180
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   15
         Width           =   3615
      End
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   1875
      TabIndex        =   4
      Top             =   3450
      Visible         =   0   'False
      Width           =   10950
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2835
         Left            =   30
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   7860
      End
   End
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   1860
      TabIndex        =   8
      Top             =   3450
      Visible         =   0   'False
      Width           =   10935
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   3165
         Left            =   45
         TabIndex        =   9
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
      Height          =   3330
      Left            =   90
      TabIndex        =   25
      Top             =   3315
      Visible         =   0   'False
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   5874
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
      TabIndex        =   2
      Top             =   -15
      Width           =   18660
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00F8E4FA&
         Height          =   2415
         Left            =   210
         TabIndex        =   67
         Top             =   -75
         Width           =   18435
         Begin VB.Frame Frame6 
            BackColor       =   &H00F8E4FA&
            ForeColor       =   &H00000000&
            Height          =   840
            Left            =   16185
            TabIndex        =   236
            Top             =   1455
            Width           =   2160
            Begin VB.OptionButton OptPI 
               BackColor       =   &H00F8E4FA&
               Caption         =   "Proforma Invoice"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   90
               TabIndex        =   238
               Top             =   510
               Width           =   1995
            End
            Begin VB.OptionButton OptQtn 
               BackColor       =   &H00F8E4FA&
               Caption         =   "Quotation"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   90
               TabIndex        =   237
               Top             =   180
               Value           =   -1  'True
               Width           =   1695
            End
         End
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
            Left            =   4890
            TabIndex        =   113
            Top             =   465
            Width           =   1560
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
            TabIndex        =   83
            Top             =   465
            Width           =   3285
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
            TabIndex        =   82
            Top             =   1275
            Width           =   645
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00F8E4FA&
            Caption         =   "Address"
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
            Left            =   6480
            TabIndex        =   79
            Top             =   90
            Width           =   3750
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
               Left            =   45
               MaxLength       =   100
               TabIndex        =   80
               Top             =   225
               Width           =   3645
            End
            Begin MSForms.TextBox TxtBillAddress 
               Height          =   1095
               Left            =   45
               TabIndex        =   81
               Top             =   570
               Width           =   3645
               VariousPropertyBits=   -1400879077
               ForeColor       =   16711680
               MaxLength       =   200
               BorderStyle     =   1
               Size            =   "6429;1931"
               SpecialEffect   =   0
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00F8E4FA&
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
            Left            =   10230
            TabIndex        =   77
            Top             =   90
            Width           =   3600
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
               TabIndex        =   78
               Top             =   180
               Width           =   3510
            End
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
            Left            =   11325
            MaxLength       =   35
            TabIndex        =   76
            Top             =   1710
            Width           =   2490
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
            TabIndex        =   75
            Top             =   465
            Width           =   1470
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
            Left            =   10890
            MaxLength       =   35
            TabIndex        =   74
            Top             =   2055
            Width           =   2925
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
            ItemData        =   "Frmquotn.frx":030A
            Left            =   8355
            List            =   "Frmquotn.frx":0323
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   1845
            Width           =   1860
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
            Left            =   7695
            MaxLength       =   1
            TabIndex        =   72
            Top             =   1845
            Width           =   630
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
            Left            =   11325
            MaxLength       =   300
            TabIndex        =   71
            Top             =   1110
            Width           =   2490
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
            Left            =   11070
            MaxLength       =   35
            TabIndex        =   70
            Top             =   765
            Width           =   2745
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
            Left            =   16170
            TabIndex        =   69
            Top             =   930
            Width           =   1410
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
            Left            =   16170
            TabIndex        =   68
            Top             =   390
            Width           =   1410
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   300
            Left            =   3255
            TabIndex        =   84
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
            TabIndex        =   85
            Top             =   840
            Width           =   4860
            _ExtentX        =   8573
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
            Left            =   13860
            TabIndex        =   86
            Top             =   420
            Width           =   2280
            _ExtentX        =   4022
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
            TabIndex        =   87
            Top             =   2040
            Width           =   4860
            _ExtentX        =   8573
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
         Begin VB.Label flagchange2 
            Height          =   315
            Left            =   270
            TabIndex        =   263
            Top             =   690
            Visible         =   0   'False
            Width           =   495
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
            ForeColor       =   &H00404000&
            Height          =   165
            Index           =   66
            Left            =   9120
            TabIndex        =   114
            Top             =   2190
            Width           =   675
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
            TabIndex        =   112
            Top             =   135
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
            Height          =   315
            Left            =   1590
            TabIndex        =   111
            Top             =   930
            Visible         =   0   'False
            Width           =   1215
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
            TabIndex        =   110
            Top             =   495
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "QTN NO."
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
            TabIndex        =   109
            Top             =   150
            Width           =   1140
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
            TabIndex        =   108
            Top             =   150
            Width           =   570
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
            TabIndex        =   107
            Top             =   870
            Width           =   1260
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
            Left            =   13860
            TabIndex        =   106
            Top             =   165
            Width           =   720
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
            TabIndex        =   105
            Top             =   1725
            Width           =   1545
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
            Left            =   10245
            TabIndex        =   104
            Top             =   2040
            Width           =   660
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
            Left            =   10230
            TabIndex        =   103
            Top             =   1665
            Width           =   1110
         End
         Begin MSForms.ComboBox TXTAREA 
            Height          =   315
            Left            =   1590
            TabIndex        =   102
            Top             =   1725
            Width           =   4860
            VariousPropertyBits=   746604571
            ForeColor       =   16711680
            MaxLength       =   200
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "8572;556"
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
            Left            =   6480
            TabIndex        =   101
            Top             =   1860
            Width           =   1110
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
            TabIndex        =   100
            Top             =   2175
            Width           =   585
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
            TabIndex        =   99
            Top             =   2190
            Width           =   615
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
            TabIndex        =   98
            Top             =   2190
            Width           =   615
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
            Left            =   10260
            TabIndex        =   97
            Top             =   1350
            Width           =   1110
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
            TabIndex        =   96
            Top             =   2085
            Width           =   1440
         End
         Begin VB.Label lblsubdealer 
            Height          =   405
            Left            =   60
            TabIndex        =   95
            Top             =   2475
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblIGST 
            BackColor       =   &H00FEF1FC&
            Height          =   285
            Left            =   5715
            TabIndex        =   94
            Top             =   2100
            Width           =   690
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
            TabIndex        =   93
            Top             =   1875
            Visible         =   0   'False
            Width           =   720
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
            Left            =   10245
            TabIndex        =   92
            Top             =   795
            Width           =   870
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
            TabIndex        =   91
            Top             =   1290
            Width           =   855
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
            Height          =   195
            Left            =   945
            TabIndex        =   90
            Top             =   150
            Width           =   450
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
            Height          =   180
            Left            =   2145
            TabIndex        =   89
            Top             =   165
            Width           =   510
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
            TabIndex        =   88
            Top             =   150
            Width           =   930
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
         Left            =   11430
         MaxLength       =   15
         TabIndex        =   10
         Top             =   10635
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F8E4FA&
         ForeColor       =   &H008080FF&
         Height          =   4380
         Left            =   210
         TabIndex        =   3
         Top             =   2250
         Width           =   18435
         Begin VB.TextBox TxtItemName 
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
            Left            =   1560
            TabIndex        =   261
            Top             =   645
            Visible         =   0   'False
            Width           =   3735
         End
         Begin MSDataListLib.DataList DataList3 
            Height          =   1425
            Left            =   1560
            TabIndex        =   262
            Top             =   990
            Visible         =   0   'False
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   2514
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
         Begin VB.TextBox TxtAddRate 
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
            Left            =   17610
            MaxLength       =   3
            TabIndex        =   260
            Top             =   1860
            Width           =   705
         End
         Begin VB.CommandButton CmdAddRate 
            Appearance      =   0  'Flat
            BackColor       =   &H00C4D6E6&
            Caption         =   "Add Rate%"
            Height          =   450
            Left            =   17535
            MaskColor       =   &H008080FF&
            Style           =   1  'Graphical
            TabIndex        =   259
            Top             =   2250
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.CommandButton cmdratechange 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            Caption         =   "Rate Change"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   17520
            MaskColor       =   &H008080FF&
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   3810
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00F4F0DB&
            Height          =   4335
            Left            =   14220
            TabIndex        =   29
            Top             =   0
            Width           =   3285
            Begin VB.CommandButton Cmdbillconvert 
               BackColor       =   &H00FF80FF&
               Caption         =   "Make this as Invoice"
               CausesValidation=   0   'False
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
               TabIndex        =   115
               Top             =   3840
               Width           =   1965
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
               Left            =   60
               TabIndex        =   64
               Top             =   2715
               Width           =   615
            End
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
               Left            =   855
               TabIndex        =   63
               Top             =   2715
               Width           =   870
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
               TabIndex        =   35
               Top             =   4770
               Visible         =   0   'False
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
               TabIndex        =   34
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
               TabIndex        =   33
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
               TabIndex        =   32
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
               TabIndex        =   31
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
               TabIndex        =   30
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
               Left            =   45
               TabIndex        =   265
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
               Left            =   690
               TabIndex        =   65
               Top             =   2790
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
               TabIndex        =   62
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
               TabIndex        =   61
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
               Height          =   330
               Index           =   53
               Left            =   45
               TabIndex        =   60
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
               TabIndex        =   59
               Top             =   4860
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
               Left            =   45
               TabIndex        =   58
               Top             =   2085
               Width           =   1680
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
               TabIndex        =   57
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
               TabIndex        =   56
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
               TabIndex        =   55
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
               TabIndex        =   54
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
               Left            =   1785
               TabIndex        =   53
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
               Left            =   45
               TabIndex        =   52
               Top             =   300
               Width           =   1680
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
               TabIndex        =   51
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
               Left            =   45
               TabIndex        =   50
               Top             =   885
               Width           =   1680
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
               TabIndex        =   49
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
               TabIndex        =   48
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
               TabIndex        =   47
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
               TabIndex        =   46
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
               TabIndex        =   45
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
               TabIndex        =   44
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
               TabIndex        =   43
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
               TabIndex        =   42
               Top             =   4455
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
               TabIndex        =   41
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
               TabIndex        =   40
               Top             =   1260
               Width           =   1635
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
               Left            =   45
               TabIndex        =   39
               Top             =   1485
               Width           =   1680
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
               TabIndex        =   38
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
               TabIndex        =   37
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
               TabIndex        =   36
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
            TabIndex        =   27
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
            TabIndex        =   18
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
            Height          =   375
            Left            =   13260
            TabIndex        =   17
            Top             =   3855
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
            Left            =   11145
            TabIndex        =   16
            Top             =   3870
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
            Left            =   12105
            TabIndex        =   15
            Top             =   3870
            Width           =   1125
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   3675
            Left            =   30
            TabIndex        =   1
            Top             =   120
            Width           =   14145
            _ExtentX        =   24950
            _ExtentY        =   6482
            _Version        =   393216
            Rows            =   1
            Cols            =   48
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   420
            BackColor       =   16050128
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
            Height          =   240
            Index           =   21
            Left            =   45
            TabIndex        =   258
            Top             =   3810
            Width           =   855
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
            Height          =   225
            Index           =   22
            Left            =   1050
            TabIndex        =   257
            Top             =   3810
            Width           =   855
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
            Height          =   225
            Index           =   30
            Left            =   2085
            TabIndex        =   256
            Top             =   3810
            Width           =   855
         End
         Begin VB.Label lblretail 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   45
            TabIndex        =   255
            Top             =   4035
            Width           =   855
         End
         Begin VB.Label lblwsale 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1050
            TabIndex        =   254
            Top             =   4035
            Width           =   855
         End
         Begin VB.Label lblvan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2085
            TabIndex        =   253
            Top             =   4035
            Width           =   855
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
            Height          =   225
            Index           =   42
            Left            =   8100
            TabIndex        =   252
            Top             =   3810
            Width           =   885
         End
         Begin VB.Label lblcrtnpack 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   9945
            TabIndex        =   251
            Top             =   4020
            Width           =   705
         End
         Begin VB.Label lblcase 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   9015
            TabIndex        =   250
            Top             =   4035
            Width           =   900
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
            Height          =   225
            Index           =   52
            Left            =   9945
            TabIndex        =   249
            Top             =   3810
            Width           =   705
         End
         Begin VB.Label lblLWPrice 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   8100
            TabIndex        =   248
            Top             =   4035
            Width           =   900
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
            Height          =   225
            Index           =   63
            Left            =   9015
            TabIndex        =   247
            Top             =   3810
            Width           =   900
         End
         Begin VB.Label LBLMRP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3120
            TabIndex        =   246
            Top             =   4035
            Width           =   855
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
            Index           =   67
            Left            =   3120
            TabIndex        =   245
            Top             =   3810
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Price 7"
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
            Index           =   72
            Left            =   6210
            TabIndex        =   244
            Top             =   3810
            Width           =   855
         End
         Begin VB.Label lblprice7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6210
            TabIndex        =   243
            Top             =   4035
            Width           =   855
         End
         Begin VB.Label lblprice6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5175
            TabIndex        =   242
            Top             =   4035
            Width           =   855
         End
         Begin VB.Label lblprice5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4140
            TabIndex        =   241
            Top             =   4035
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Price 6"
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
            Index           =   73
            Left            =   5175
            TabIndex        =   240
            Top             =   3810
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Price 5"
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
            Index           =   74
            Left            =   4140
            TabIndex        =   239
            Top             =   3810
            Width           =   855
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
            TabIndex        =   26
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
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   4920
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F8E4FA&
         ForeColor       =   &H008080FF&
         Height          =   4365
         Left            =   210
         TabIndex        =   116
         Top             =   6525
         Width           =   18450
         Begin VB.TextBox TXTCNTYPE 
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
            Left            =   2295
            MaxLength       =   6
            TabIndex        =   232
            Top             =   2985
            Visible         =   0   'False
            Width           =   795
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
            Left            =   17265
            MaxLength       =   6
            TabIndex        =   185
            Top             =   1095
            Width           =   630
         End
         Begin VB.TextBox TxtDiscAmt 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   15480
            MaxLength       =   5
            TabIndex        =   184
            Top             =   375
            Width           =   750
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   14085
            TabIndex        =   183
            Top             =   375
            Width           =   870
         End
         Begin VB.TextBox TxtSpec 
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
            ForeColor       =   &H00000000&
            Height          =   1050
            Left            =   30
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   182
            Top             =   1770
            Width           =   4965
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
            TabIndex        =   181
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   7050
            MaxLength       =   5
            TabIndex        =   180
            Top             =   375
            Width           =   375
         End
         Begin VB.CommandButton Command4 
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
            Left            =   13530
            TabIndex        =   179
            Top             =   1545
            Width           =   300
         End
         Begin VB.TextBox TxtCessAmt 
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   16785
            MaxLength       =   5
            TabIndex        =   178
            Top             =   450
            Width           =   465
         End
         Begin VB.TextBox TxtCessPer 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   16245
            MaxLength       =   5
            TabIndex        =   177
            Top             =   375
            Width           =   525
         End
         Begin VB.PictureBox picUnchecked 
            Height          =   285
            Left            =   16080
            Picture         =   "Frmquotn.frx":0353
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   176
            Top             =   2925
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.PictureBox picChecked 
            Height          =   285
            Left            =   15435
            Picture         =   "Frmquotn.frx":0695
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   175
            Top             =   2895
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
            TabIndex        =   174
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
            Left            =   7170
            TabIndex        =   173
            Top             =   855
            Width           =   855
         End
         Begin VB.TextBox txtPrintname 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   1155
            TabIndex        =   172
            Top             =   1065
            Width           =   3840
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
            Left            =   5010
            TabIndex        =   171
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
            Left            =   5010
            MaxLength       =   255
            TabIndex        =   170
            Top             =   1515
            Width           =   7065
         End
         Begin VB.CommandButton Command2 
            Height          =   435
            Left            =   15690
            TabIndex        =   169
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
            TabIndex        =   168
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
            TabIndex        =   167
            Top             =   2655
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
            TabIndex        =   166
            Top             =   2370
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtName1 
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
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   2010
            MaxLength       =   15
            TabIndex        =   165
            Top             =   390
            Width           =   1005
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
            TabIndex        =   164
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
            Left            =   8040
            TabIndex        =   163
            Top             =   855
            Width           =   915
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
            Left            =   8985
            TabIndex        =   162
            Top             =   855
            Width           =   915
         End
         Begin VB.TextBox TxtWarranty 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   30
            MaxLength       =   30
            TabIndex        =   161
            Top             =   1080
            Width           =   450
         End
         Begin VB.TextBox TxtWarranty_type 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   495
            MaxLength       =   30
            TabIndex        =   160
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
            TabIndex        =   159
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
            TabIndex        =   158
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
            TabIndex        =   157
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
            TabIndex        =   156
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
            TabIndex        =   155
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   7440
            MaxLength       =   15
            TabIndex        =   154
            Top             =   375
            Width           =   840
         End
         Begin VB.TextBox TXTITEMCODE 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   480
            TabIndex        =   153
            Top             =   375
            Width           =   1515
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
            Left            =   5745
            TabIndex        =   152
            Top             =   855
            Width           =   720
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
            Left            =   6480
            TabIndex        =   151
            Top             =   855
            Width           =   675
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
            TabIndex        =   150
            Top             =   855
            Width           =   885
         End
         Begin VB.TextBox TXTDISC 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   14970
            MaxLength       =   5
            TabIndex        =   149
            Top             =   375
            Width           =   495
         End
         Begin VB.TextBox TXTTAX 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   11205
            MaxLength       =   4
            TabIndex        =   148
            Top             =   375
            Width           =   390
         End
         Begin VB.TextBox TXTQTY 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   8295
            MaxLength       =   8
            TabIndex        =   147
            Top             =   375
            Width           =   480
         End
         Begin VB.TextBox TXTPRODUCT 
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
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   3030
            TabIndex        =   146
            Top             =   390
            Width           =   3510
         End
         Begin VB.TextBox TXTSLNO 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   30
            TabIndex        =   145
            Top             =   375
            Width           =   435
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
            Left            =   5010
            TabIndex        =   144
            Top             =   855
            Width           =   720
         End
         Begin VB.TextBox TxtMRP 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   9825
            MaxLength       =   6
            TabIndex        =   143
            Top             =   375
            Width           =   675
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
            TabIndex        =   142
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
            TabIndex        =   141
            Top             =   4740
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
            TabIndex        =   140
            Top             =   4710
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.TextBox TXTFREE 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   8790
            MaxLength       =   7
            TabIndex        =   139
            Top             =   375
            Width           =   405
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
            TabIndex        =   138
            Top             =   4290
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00DFF9F5&
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
            TabIndex        =   137
            Top             =   2325
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00DFF9F5&
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
            TabIndex        =   136
            Top             =   2070
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
            TabIndex        =   135
            Top             =   3180
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.OptionButton optnet 
            BackColor       =   &H00DFF9F5&
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
            TabIndex        =   134
            Top             =   2085
            Width           =   720
         End
         Begin VB.TextBox TXTRETAILNOTAX 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   12120
            MaxLength       =   9
            TabIndex        =   133
            Top             =   375
            Width           =   990
         End
         Begin VB.TextBox txtretail 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   13125
            MaxLength       =   9
            TabIndex        =   132
            Top             =   375
            Width           =   945
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
            TabIndex        =   131
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
            TabIndex        =   130
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
            TabIndex        =   129
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
            TabIndex        =   128
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
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   125
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
            TabIndex        =   124
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
            TabIndex        =   123
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
            TabIndex        =   122
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   6555
            MaxLength       =   8
            TabIndex        =   121
            Top             =   375
            Width           =   480
         End
         Begin VB.Frame FrmeType 
            BackColor       =   &H0080C0FF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   690
            Left            =   9930
            TabIndex        =   118
            Top             =   825
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
               TabIndex        =   120
               Top             =   120
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
               TabIndex        =   119
               Top             =   390
               Width           =   1020
            End
         End
         Begin VB.ComboBox CMBBillType 
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
            ForeColor       =   &H00004080&
            Height          =   360
            ItemData        =   "Frmquotn.frx":09D7
            Left            =   9210
            List            =   "Frmquotn.frx":09E7
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   465
            Width           =   615
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   450
            Left            =   10515
            TabIndex        =   186
            Top             =   375
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   794
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
            Enabled         =   0   'False
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GRDRECEIPT 
            Height          =   1125
            Left            =   0
            TabIndex        =   187
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
            Left            =   5025
            TabIndex        =   188
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
            Left            =   13965
            TabIndex        =   235
            Top             =   1605
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
            Left            =   15030
            TabIndex        =   234
            Top             =   1605
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
            Left            =   16080
            TabIndex        =   233
            Top             =   1605
            Width           =   1650
         End
         Begin VB.Label TxtKFC 
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   11610
            TabIndex        =   231
            Top             =   375
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "KFC%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   69
            Left            =   11610
            TabIndex        =   230
            Top             =   150
            Width           =   495
         End
         Begin VB.Label LBLUNBILL 
            Height          =   225
            Left            =   15015
            TabIndex        =   229
            Top             =   1545
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
            Left            =   17955
            TabIndex        =   228
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   65
            Left            =   15480
            TabIndex        =   227
            Top             =   150
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Net Rate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   225
            Index           =   64
            Left            =   14085
            TabIndex        =   226
            Top             =   150
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Product Specifications"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   59
            Left            =   30
            TabIndex        =   225
            Top             =   1530
            Width           =   4965
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Add Cess"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Index           =   62
            Left            =   16785
            TabIndex        =   224
            Top             =   150
            Width           =   465
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Cess%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   29
            Left            =   16245
            TabIndex        =   223
            Top             =   150
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Expiry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   61
            Left            =   10515
            TabIndex        =   222
            Top             =   150
            Width           =   675
         End
         Begin VB.Label lblbarcode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   30
            TabIndex        =   221
            Top             =   3645
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.Label lblactqty 
            Height          =   375
            Left            =   4200
            TabIndex        =   220
            Top             =   1560
            Visible         =   0   'False
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
            TabIndex        =   219
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   58
            Left            =   8790
            TabIndex        =   218
            Top             =   150
            Width           =   405
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Name to be Printed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   38
            Left            =   1155
            TabIndex        =   217
            Top             =   840
            Width           =   3840
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
            TabIndex        =   216
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
            TabIndex        =   215
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
            TabIndex        =   214
            Top             =   795
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Comm"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   46
            Left            =   17265
            TabIndex        =   213
            Top             =   855
            Width           =   630
         End
         Begin VB.Label LBLTYPE 
            Caption         =   "HI"
            Height          =   330
            Left            =   4230
            TabIndex        =   212
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
            TabIndex        =   211
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
            TabIndex        =   210
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
            TabIndex        =   209
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
            TabIndex        =   208
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   17265
            TabIndex        =   207
            Top             =   375
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Serial No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   7
            Left            =   7440
            TabIndex        =   206
            Top             =   150
            Width           =   840
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
            TabIndex        =   205
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   14
            Left            =   17265
            TabIndex        =   204
            Top             =   150
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Disc%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   13
            Left            =   14970
            TabIndex        =   203
            Top             =   150
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Tax"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   12
            Left            =   11205
            TabIndex        =   202
            Top             =   150
            Width           =   390
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Rate+GST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   11
            Left            =   13125
            TabIndex        =   201
            Top             =   150
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Qty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   10
            Left            =   8295
            TabIndex        =   200
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   9
            Left            =   2010
            TabIndex        =   199
            Top             =   150
            Width           =   4530
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "SL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   8
            Left            =   30
            TabIndex        =   198
            Top             =   150
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "MRP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   24
            Left            =   9825
            TabIndex        =   197
            Top             =   150
            Width           =   675
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
            TabIndex        =   196
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   30
            Left            =   12120
            TabIndex        =   195
            Top             =   150
            Width           =   990
         End
         Begin VB.Label lblP_Rate 
            Caption         =   "0"
            Height          =   390
            Left            =   13200
            TabIndex        =   194
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   43
            Left            =   480
            TabIndex        =   193
            Top             =   150
            Width           =   1515
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Warranty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   48
            Left            =   30
            TabIndex        =   192
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   16
            Left            =   7050
            TabIndex        =   191
            Top             =   150
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Pack"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   37
            Left            =   6555
            TabIndex        =   190
            Top             =   150
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Index           =   71
            Left            =   9210
            TabIndex        =   189
            Top             =   150
            Width           =   600
         End
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   13155
      TabIndex        =   11
      Top             =   3090
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1217
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grdcount 
      Height          =   5145
      Left            =   0
      TabIndex        =   28
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
   Begin VB.Label LBLitem 
      Height          =   315
      Left            =   0
      TabIndex        =   264
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   -15
      TabIndex        =   14
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11355
      TabIndex        =   13
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   12
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMQUOTATION"
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
Dim ITEM_REC As New ADODB.Recordset
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
Dim Small_Print, Dos_Print, ST_PRINT As Boolean
Dim CHANGE_ADDRESS, CHANGE_NAME As Boolean
Dim item_change As Boolean
Dim thermalprn As Boolean
Dim mrpplus As Boolean

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
    If cmbtype.ListIndex = 4 Then TXTTYPE.text = 5
    If cmbtype.ListIndex = 5 Then TXTTYPE.text = 6
    If cmbtype.ListIndex = 6 Then TXTTYPE.text = 7
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

Private Sub Cmdbillconvert_Click()
    Dim BillType As String
    If grdsales.rows = 1 Then Exit Sub
        
    Select Case Trim(TXTCNTYPE.text)
        Case "GI"
            BillType = "-GST B2B Sales"
        Case "HI"
            BillType = "-GST B2C Sales"
        Case "WO"
            BillType = "-Petty Sales"
    End Select
    If Val(TxtCN.text) > 0 Then
        If (MsgBox("Already added to " & BillType & " Bill No: " & Val(TxtCN.text) & ". Do you want to make the invoice again?", vbYesNo + vbDefaultButton2, "QUOTATION") = vbNo) Then Exit Sub
        'MsgBox "Already added to " & BillType & " Bill No: " & Val(TxtCN.Text)
        'Exit Sub
    End If
    
    'If (MsgBox("Are you sure you want to make this Quotation as Bill?", vbYesNo, "QUOTATION") = vbNo) Then Exit Sub


    Me.Enabled = False
    M_ADD = True
    Set creditbill = Me
    frmINVTYPE.Show
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        frmINVTYPE.Opt8B.Visible = False
        frmINVTYPE.Opt8.Visible = True
        frmINVTYPE.Opt8.Caption = "SALES BILL"
        frmINVTYPE.OptPetty.Visible = True
        frmINVTYPE.Opt8.Value = True
    Else
        frmINVTYPE.Opt8B.Visible = True
        frmINVTYPE.Opt8.Visible = True
        frmINVTYPE.OptPetty.Visible = True
        If Trim(TXTTIN.text) = "" Then
            frmINVTYPE.Opt8.Value = True
        Else
            frmINVTYPE.Opt8B.Value = True
        End If
    End If
    
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
    
    If grdsales.rows > 1 Then
        If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
            MsgBox "Permission Denied", vbOKOnly, "QUOTATION"
            Exit Sub
        End If
        If MsgBox("ARE YOU SURE YOU WANT TO CANCEL THE QUOTATION!!!!!", vbYesNo, "DELETE!!!") = vbNo Then
            Chkcancel.Value = 0
            Exit Sub
        End If
    End If
    
        
    db.Execute "delete FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " "
    
    
    
    grdsales.FixedRows = 0
    grdsales.rows = 1
    LBLGROSSAMT.Caption = ""
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    TXTTOTALDISC.text = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
    flagchange2.Caption = ""
    Call AppendSale
    Chkcancel.Value = 0
End Sub

Private Sub CMDDOS_Click()
    Chkcancel.Value = 0
    If grdsales.rows = 1 Then Exit Sub
    
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) FROM QTNMAST", db, adOpenForwardOnly
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
    Me.lblcredit.Caption = "0"
    Me.Generateprint
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
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                TxtName1.SetFocus
            Else
                TXTITEMCODE.SetFocus
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
    'Tax_Print = False
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    'If Month(Date) >= 5 And Year(Date) >= 2021 Then Exit Sub
'    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then
'        'db.Execute "delete From USERS "
'        Exit Sub
'    End If
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) FROM QTNMAST", db, adOpenForwardOnly
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
    Me.lblcredit.Caption = "0"
    Me.Generateprint
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
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                TxtName1.SetFocus
            Else
                TXTITEMCODE.SetFocus
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
        Case 4
            If (MsgBox("Are you sure you want to change the whole rate to Price 5?", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then Exit Sub
        Case 5
            If (MsgBox("Are you sure you want to change the whole rate to Price 6?", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then Exit Sub
        Case 6
            If (MsgBox("Are you sure you want to change the whole rate to Price 7?", vbYesNo + vbDefaultButton2, "SALES") = vbNo) Then Exit Sub
        Case Else
            MsgBox "Please select a valid Rate type from the list", , "SALES"
            On Error Resume Next
            cmbtype.SetFocus
            Exit Sub
    End Select
    On Error GoTo ERRHAND
     
    Dim i As Integer
    Dim pricechange As Boolean
    Dim Newprice As Double
    Dim RSTTRXFILE As ADODB.Recordset
    For i = 1 To grdsales.rows - 1
        Newprice = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND PRICE_CHANGE ='Y'", db, adOpenForwardOnly
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                pricechange = True
            Else
                pricechange = False
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Select Case cmbtype.ListIndex
            Case 1
                If Not (MDIMAIN.lblPriceSplit.Caption = "Y" Or PC_FLAG = "Y" Or pricechange = True) Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
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
                If Newprice = 0 Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
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
                If Not (MDIMAIN.lblPriceSplit.Caption = "Y" Or PC_FLAG = "Y" Or pricechange = True) Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
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
                    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
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
                    If Not (MDIMAIN.lblPriceSplit.Caption = "Y" Or PC_FLAG = "Y" Or pricechange = True) Then
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
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
                    If Newprice = 0 Then
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
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
                If Not (MDIMAIN.lblPriceSplit.Caption = "Y" Or PC_FLAG = "Y" Or pricechange = True) Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
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
                If Newprice = 0 Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
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
            Case 4
                If Not (MDIMAIN.lblPriceSplit.Caption = "Y" Or PC_FLAG = "Y" Or pricechange = True) Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!PRICE5), 0, RSTTRXFILE!PRICE5)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!PRICE5), 0, RSTTRXFILE!PRICE5)
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
                    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!PRICE5), 0, RSTTRXFILE!PRICE5)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!PRICE5), 0, RSTTRXFILE!PRICE5)
                                Else
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                                End If
                            End If
                        End If
                    End With
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
            Case 5
                If Not (MDIMAIN.lblPriceSplit.Caption = "Y" Or PC_FLAG = "Y" Or pricechange = True) Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!PRICE6), 0, RSTTRXFILE!PRICE6)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!PRICE6), 0, RSTTRXFILE!PRICE6)
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
                    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!PRICE6), 0, RSTTRXFILE!PRICE6)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!PRICE6), 0, RSTTRXFILE!PRICE6)
                                Else
                                    Newprice = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
                                End If
                            End If
                        End If
                    End With
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
            Case 6
                If Not (MDIMAIN.lblPriceSplit.Caption = "Y" Or PC_FLAG = "Y" Or pricechange = True) Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!PRICE7), 0, RSTTRXFILE!PRICE7)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!PRICE7), 0, RSTTRXFILE!PRICE7)
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
                    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
                    With RSTTRXFILE
                        If Not (.EOF And .BOF) Then
                            If IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0 Or RSTTRXFILE!LOOSE_PACK = 1 Then
                                Newprice = IIf(IsNull(RSTTRXFILE!PRICE7), 0, RSTTRXFILE!PRICE7)
                            Else
                                If Val(grdsales.TextMatrix(i, 27)) <> 1 Then
                                    Newprice = IIf(IsNull(RSTTRXFILE!PRICE7), 0, RSTTRXFILE!PRICE7)
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
                If Not (MDIMAIN.lblPriceSplit.Caption = "Y" Or PC_FLAG = "Y" Or pricechange = True) Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' ", db, adOpenForwardOnly
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
                If Newprice = 0 Then
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 13)) & "'", db, adOpenForwardOnly
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
        
        TXTDISC.Tag = 0
        If UCase(grdsales.TextMatrix(i, 25)) = "PARDHA" Or UCase(grdsales.TextMatrix(i, 25)) = "CLOTHES" Then
            If Val(grdsales.TextMatrix(i, 6)) < 1000 Then
               grdsales.TextMatrix(i, 9) = "5"
            Else
                grdsales.TextMatrix(i, 9) = "12"
            End If
        End If
        
        If Trim(TXTTIN.text) = "" And Trim(lblIGST.Caption) <> "Y" And MDIMAIN.lblkfc.Caption = "Y" And IsDate(MDIMAIN.DTKFCSTART.Value) And IsDate(MDIMAIN.DTKFCEND.Value) Then
            If DateValue(TXTINVDATE.text) >= DateValue(MDIMAIN.DTKFCSTART.Value) And DateValue(TXTINVDATE.text) <= DateValue(MDIMAIN.DTKFCEND.Value) Then
                If Val(grdsales.TextMatrix(i, 9)) = 12 Or Val(grdsales.TextMatrix(i, 9)) = 18 Or Val(grdsales.TextMatrix(i, 9)) = 28 Then
                    grdsales.TextMatrix(i, 47) = 1
                Else
                    grdsales.TextMatrix(i, 47) = 0
                End If
            End If
        Else
            TxtKFC.Caption = 0
        End If
        
        TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(i, 6)) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 8)) / 100)
        TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(grdsales.TextMatrix(i, 9)) + Val(grdsales.TextMatrix(i, 47))) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(i, 40)) / 100) + Val(grdsales.TextMatrix(i, 41))
        If UCase(grdsales.TextMatrix(i, 25)) = "SERVICE CHARGE" Then
            'TXTDISC.Tag = Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
            grdsales.TextMatrix(i, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag), 4), ".0000")
            grdsales.TextMatrix(i, 34) = Format(Round(Val(grdsales.TextMatrix(i, 6)), 4), ".0000")
        Else
            'TXTDISC.Tag = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
            grdsales.TextMatrix(i, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(i, 3)), 4), ".0000")
            grdsales.TextMatrix(i, 34) = Format(Round(Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3)), 4), ".0000")
        End If
                                
'        If UCase(grdsales.TextMatrix(i, 25)) = "SERVICE CHARGE" Then
'            TXTDISC.Tag = Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
'            grdsales.TextMatrix(i, 12) = Format(Round(Val(grdsales.TextMatrix(i, 7)) - Val(TXTDISC.Tag), 4), ".0000")
'            grdsales.TextMatrix(i, 34) = Format(Round(Val(grdsales.TextMatrix(i, 6)) - Val(TXTDISC.Tag), 4), ".0000")
'        Else
'            TXTDISC.Tag = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
'            grdsales.TextMatrix(i, 12) = Format(Round((Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 7))) - Val(TXTDISC.Tag), 4), ".0000")
'            grdsales.TextMatrix(i, 34) = Format(Round((Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 6))) - Val(TXTDISC.Tag), 4), ".0000")
'        End If
        
        
        TXTDISC.Tag = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        db.BeginTrans
        db.Execute "Update QTNSUB set SALES_PRICE = " & Val(grdsales.TextMatrix(i, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(i, 7)) & ", PTR = " & Val(grdsales.TextMatrix(i, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(i, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(i, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & ""
        db.CommitTrans
SKIP:
    Next i
    
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
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                If MDIMAIN.lblgst.Caption = "R" Then
                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                End If
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
    Else
        db.Execute "Update QTNMAST set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
    End If
    'db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.Text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'QT'"
    
                    
                    
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CmdTax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub Command3_Click()
    Chkcancel.Value = 0
    If grdsales.rows = 1 Then Exit Sub
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    thermalprn = False
    'Tax_Print = False
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    'If Month(Date) >= 5 And Year(Date) >= 2021 Then Exit Sub
'    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then
'        'db.Execute "delete From USERS "
'        Exit Sub
'    End If
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) FROM QTNMAST", db, adOpenForwardOnly
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
    thermalprn = True
    Small_Print = False
    Dos_Print = False
    Set creditbill = Me
    Me.lblcredit.Caption = "0"
    Me.Generateprint
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
                Case "5"
                    cmbtype.ListIndex = 4
                    TXTTYPE.text = 5
                Case "6"
                    cmbtype.ListIndex = 5
                    TXTTYPE.text = 6
                Case "7"
                    cmbtype.ListIndex = 6
                    TXTTYPE.text = 7
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
    
    
SKIP:
    If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then
        Frame5.Visible = False
    Else
        Frame5.Visible = True
    End If
    TXTDEALER.text = DataList2.text
    lbldealer.Caption = TXTDEALER.text
    TxtCode.text = DataList2.BoundText
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
                        If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                            TxtName1.SetFocus
                        Else
                            TXTITEMCODE.SetFocus
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
        MsgBox "Please enter the Qty", vbOKOnly, "QUOTATION"
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
        Exit Sub
    End If
    If MDIMAIN.LBLTAXWARN.Caption = "Y" Then
        If Val(TXTTAX.text) = 0 Then
            If (MsgBox("Tax is Zero. Are you sure?", vbYesNo + vbDefaultButton2, "QUOTATION") = vbNo) Then
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
                Exit Sub
            End If
        End If
    End If
    Call TXTQTY_LostFocus
    If Val(TXTRETAILNOTAX.text) = 0 Then
        Call TXTRETAIL_LostFocus
    End If
    
    If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
        Call TXTRETAILNOTAX_LostFocus
    End If
    
    If Val(TXTQTY.text) <> 0 And MDIMAIN.StatusBar.Panels(14).text <> "Y" And Val(TXTRETAILNOTAX.text) = 0 Then
        MsgBox "Please enter the Rate", vbOKOnly, "QUOTATION"
        TXTRETAILNOTAX.Enabled = True
        TXTRETAILNOTAX.SetFocus
        Exit Sub
    End If
    If Val(TXTQTY.text) <> 0 And MDIMAIN.StatusBar.Panels(14).text = "Y" And Val(txtretail.text) = 0 Then
        MsgBox "Please enter the Rate", vbOKOnly, "QUOTATION"
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
    
    Dim rststock As ADODB.Recordset
    'Dim RSTMINQTY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim i As Long
    
    Chkcancel.Value = 0
    On Error GoTo ERRHAND
    'If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then Exit Sub
    'If Month(Date) >= 5 And Year(Date) >= 2021 Then Exit Sub
'    If Month(TXTINVDATE.Text) >= 6 And Year(TXTINVDATE.Text) >= 2021 Then
'        db.Execute "delete From USERS "
'        Exit Sub
'    End If
    
    'If OLD_BILL = False Then Call checklastbill
    Set RSTTRXFILE = New ADODB.Recordset
    If OLD_BILL = False And Val(txtBillNo.text) <> 1 Then
        RSTTRXFILE.Open "Select * FROM QTNMAST WHERE VCH_NO= (SELECT MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'QT')", db, adOpenStatic, adLockOptimistic, adCmdText
        txtBillNo.text = RSTTRXFILE!VCH_NO + 1
        LBLBILLNO.Caption = Val(txtBillNo.text)
        db.BeginTrans
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "QT"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = txtBillNo.text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
    Else
        RSTTRXFILE.Open "Select * FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "QT"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = txtBillNo.text
            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
            RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
        End If
    End If
    'RSTTRXFILE!CUST_IGST = lblIGST.Caption
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
    'RSTTRXFILE!BILL_FLAG = "Y"
    'RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
    'RSTTRXFILE!BR_NAME = CMBBRNCH.Text
    'RSTTRXFILE!cr_days = Val(txtcrdays.Text)
    RSTTRXFILE.Update
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    OLD_BILL = True
    
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
    grdsales.TextMatrix(Val(TXTSLNO.text), 47) = Val(TxtKFC.Caption)
    If OPTVAT.Value = True And Val(TXTTAX.text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.text), 17) = "V"
    If OPTTaxMRP.Value = True And Val(TXTTAX.text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.text), 17) = "M"
    If Val(TXTTAX.text) <= 0 Or optnet.Value = True Then grdsales.TextMatrix(Val(TXTSLNO.text), 17) = "N"
    
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "N"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
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
        rstMaxNo.Open "Select MAX(LINE_NO) FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
        If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
            grdsales.TextMatrix(Val(TXTSLNO.text), 32) = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        Else
            grdsales.TextMatrix(Val(TXTSLNO.text), 32) = Val(TXTSLNO.text)
        End If
        rstMaxNo.Close
        Set rstMaxNo = Nothing
    End If

    db.Execute "delete FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32)) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "QT"
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
    Else
        RSTTRXFILE!COM_FLAG = "N"
    End If
    RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(Val(TXTSLNO.text), 26)
    RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 27))
    RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(Val(TXTSLNO.text), 9)
    RSTTRXFILE!UNIT = grdsales.TextMatrix(Val(TXTSLNO.text), 4)
    RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.text), 1, 30)
    RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 10))
    RSTTRXFILE!ISSUE_QTY = 0
    RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 17))
    RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 18))
    RSTTRXFILE!CST = 0
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
    RSTTRXFILE!kfc_tax = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 47))
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
    'RSTTRXFILE!DN_LINENO = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37))
    
'    Dim RSTUNBILL As ADODB.Recordset
'    Set RSTUNBILL = New ADODB.Recordset
'    RSTUNBILL.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
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
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
        
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
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                If MDIMAIN.lblgst.Caption = "R" Then
                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                End If
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
    End If
    
    
    txtretail.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTSLNO.text = grdsales.rows
    TXTPRODUCT.text = ""
    txtPrintname.text = ""
    TxtSpec.text = ""
    txtcategory.text = ""
    'TxtName1.Text = ""
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
    lblprice5.Caption = ""
    lblprice6.Caption = ""
    lblprice7.Caption = ""
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
    OLD_BILL = True
    Call COSTCALCULATION
    Call Addcommission
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
    If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
        TxtName1.SetFocus
    Else
        TXTITEMCODE.SetFocus
    End If
    'grdsales.TopRow = grdsales.Rows - 1
Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
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
        MsgBox "Permission Denied", vbOKOnly, "QUOTATION"
        Exit Sub
    End If
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "QUOTATION"
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
        
    db.Execute "delete FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32)) & ""
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
        grdsales.TextMatrix(i, 47) = grdsales.TextMatrix(i + 1, 47)
    Next i
    grdsales.rows = grdsales.rows - 1
    
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
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                If MDIMAIN.lblgst.Caption = "R" Then
                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                End If
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
    If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
        TxtName1.SetFocus
    Else
        TXTITEMCODE.SetFocus
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
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                TxtName1.SetFocus
            Else
                TXTITEMCODE.SetFocus
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
        MsgBox "Permission Denied", vbOKOnly, "QUOTATION"
        Exit Sub
    End If
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Cannot modify here", vbOKOnly, "QUOTATION"
        Exit Sub
    End If
    
    TXTSLNO.text = grdsales.TextMatrix(grdsales.Row, 0)
    If grdsales.TextMatrix(Val(TXTSLNO.text), 19) = "DN" Then
        MsgBox "Cannot modify this. The Item is being Delivered. DN# ", vbOKOnly, "QUOTATION"
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
    
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    CMDPre.Caption = "Save && Pre&vious"
    cmdnxt.Caption = "Save && &Next"
    M_EDIT = True
    TXTQTY.Enabled = True
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
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                TxtName1.SetFocus
            Else
                TXTITEMCODE.SetFocus
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
    'Tax_Print = False
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
'    TRXMAST.Open "Select MAX(VCH_NO) FROM QTNMAST", db, adOpenForwardOnly
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
    Me.lblcredit.Caption = "0"
    Me.Generateprint
    
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
    If thermalprn = False And (MDIMAIN.StatusBar.Panels(8).text = "Y" Or MDIMAIN.lblprint.Caption = "Y") Then
    'If Val(MDIMAIN.StatusBar.Panels(10).Text) = 0 Then
        'Call ReportGeneratION_vpestimate
        LBLFOT.Tag = ""
        If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = True Then
            If MsgBox("You do not have any permission to modify this further. Are you sure to print?", vbYesNo, "BILL..") = vbNo Then Exit Function
        Else
            Screen.MousePointer = vbHourglass
            Sleep (300)
        End If
        NEW_BILL = False
        
        If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
            Call ReportGeneratION_Composition
        Else
            Call ReportGeneratION_vpestimate(0, 0)
        End If
        
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
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                TxtName1.SetFocus
            Else
                TXTITEMCODE.SetFocus
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
    
    If txtBillNo.Tag = "Y" Then
        flagchange2.Caption = ""
        Call AppendSale
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
    
    lblcredit.Caption = "0"
    Call AppendSale
    TxtCN.text = ""
    TXTCNTYPE.text = ""
    TXTTOTALDISC.text = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
    Chkcancel.Value = 0
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    cmbtype.ListIndex = 0
    TXTTYPE.text = 1
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
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                TxtName1.SetFocus
            Else
                TXTITEMCODE.SetFocus
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
    DataList2.BackColor = vbWhite
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
    'If txtcommi.Enabled = True Then txtcommi.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
    If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
    'If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
    'If  Then CMDDOS.SetFocus
    If TxtName1.Enabled = True Then TxtName1.SetFocus
    
    If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
    If TxtBillName.Enabled = True Then TxtBillName.SetFocus
'    If OLD_BILL = True Then
'        If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
'    Else
'        If TXTDEALER.Enabled = True Then TXTDEALER.SetFocus
'    End If
    If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
        If TxtName1.Enabled = True Then TxtName1.SetFocus
    Else
        If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
    End If
    
    grdsales.ColWidth(46) = 0
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
        
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'QT'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = Val(txtBillNo.text)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    If MDIMAIN.lblgst.Caption <> "R" Then LBLGROSSAMT.Visible = False
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        lblpre.Caption = "" 'IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
        lblsuf.Caption = "" 'IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        If RSTCOMPANY!mrp_plus = "Y" Then
            mrpplus = True
        Else
            mrpplus = False
        End If
        If MDIMAIN.StatusBar.Panels(15).text = "Y" Then   ' barcode
            Label1(43).Visible = False
            TXTITEMCODE.Visible = False
            Label1(9).Left = 500
            Label1(9).Width = Val(Label1(9).Width) + 1500
            TxtName1.Left = 500
            TxtName1.Width = 2500
        End If
        If RSTCOMPANY!hide_deliver = "Y" Then    'delivery list
            grdEXPIRYLIST.Visible = False
            CHKSELECT.Visible = False
        End If
        If RSTCOMPANY!hide_terms = "Y" Then      'terms
            chkTerms.Visible = False
            Terms1.Visible = False
        End If
        If RSTCOMPANY!hide_wrnty = "Y" Then     'warranty
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
        If RSTCOMPANY!hide_serial = "Y" Then            'batch
            Label1(7).Visible = False
            txtBatch.Visible = False
            Label1(16).Left = 7760
            lblunit.Left = 7760
            Label1(37).Left = 7260
            LblPack.Left = 7260
            Label1(10).Left = 8150
            TXTQTY.Left = 8150
            Label1(10).Width = 630
            TXTQTY.Width = 630
            'Label1(9).Width = 5760
            Label1(9).Width = Val(Label1(9).Width) + 700
            TXTPRODUCT.Width = Val(TXTPRODUCT.Width) + 700
        End If
        If RSTCOMPANY!hide_free = "Y" Then
            Label1(58).Visible = False
            TXTFREE.Visible = False
            Label1(10).Width = 1050
            TXTQTY.Width = 1050
'            Label1(24).Width = 1050
'            TxtMRP.Width = 1050
'            Label1(24).Left = 9705
'            TxtMRP.Left = 9705
        End If
        
        If RSTCOMPANY!billtype_flag = "Y" And RSTCOMPANY!hide_mrp = "N" Then
            Label1(24).Left = 9210
            Label1(24).Width = 1290
            TxtMRP.Left = 9210
            TxtMRP.Width = 1290
            Label1(71).Visible = False
            CMBBillType.Visible = False
        ElseIf RSTCOMPANY!billtype_flag = "Y" And RSTCOMPANY!hide_mrp = "Y" And RSTCOMPANY!hide_serial = "Y" Then
            Label1(10).Left = Label1(10).Left + 630
            TXTQTY.Left = TXTQTY.Left + 630
            Label1(10).Width = Label1(10).Width '+ 100
            TXTQTY.Width = TXTQTY.Width '+ 100
            Label1(16).Left = Label1(16).Left + 630
            lblunit.Left = lblunit.Left + 630
            Label1(37).Left = Label1(37).Left + 630
            LblPack.Left = LblPack.Left + 630
            Label1(9).Width = Val(Label1(9).Width) + 630
            TXTPRODUCT.Width = Val(TXTPRODUCT.Width) + 630
            Label1(71).Visible = False
            CMBBillType.Visible = False
        ElseIf RSTCOMPANY!billtype_flag = "Y" And RSTCOMPANY!hide_mrp = "Y" And RSTCOMPANY!hide_serial = "N" Then

'            Label1(10).Width = Label1(10).Width '+ 100
'            TXTQTY.Width = TXTQTY.Width '+ 100
'            Label1(16).Left = Label1(16).Left + 630
'            lblunit.Left = lblunit.Left + 630
'            Label1(37).Left = Label1(37).Left + 630
'            LblPack.Left = LblPack.Left + 630
'            Label1(9).Width = Val(Label1(9).Width) + 630
'            TXTPRODUCT.Width = Val(TXTPRODUCT.Width) + 630
            Label1(10).Left = Label1(10).Left + 480
            TXTQTY.Left = TXTQTY.Left + 480
            Label1(7).Width = Label1(7).Width + 480
            txtBatch.Width = txtBatch.Width + 480
            Label1(71).Visible = False
            CMBBillType.Visible = False
        End If
        
        If RSTCOMPANY!hide_mrp = "Y" And RSTCOMPANY!hide_expiry = "Y" Then
            Label1(24).Visible = False
            TxtMRP.Visible = False
            Label1(61).Visible = False
            TXTEXPIRY.Visible = False
            Label1(69).Left = Val(Label1(69).Left) - 1370
            TxtKFC.Left = Val(TxtKFC.Left) - 1370
            Label1(12).Left = Val(Label1(12).Left) - 1370
            TXTTAX.Left = Val(TXTTAX.Left) - 1370
            Lblprice(30).Left = Val(Lblprice(30).Left) - 1370
            TXTRETAILNOTAX.Left = Val(TXTRETAILNOTAX.Left) - 1370
            Label1(11).Left = Val(Label1(11).Left) - 1370
            txtretail.Left = Val(txtretail.Left) - 1370
            Label1(64).Left = Val(Label1(64).Left) - 1370
            txtNetrate.Left = Val(txtNetrate.Left) - 1370
            Label1(13).Left = Val(Label1(13).Left) - 1370
            TXTDISC.Left = Val(TXTDISC.Left) - 1370
            Label1(65).Left = Val(Label1(65).Left) - 1370
            TxtDiscAmt.Left = Val(TxtDiscAmt.Left) - 1370
            Label1(29).Left = Val(Label1(29).Left) - 1370
            TxtCessPer.Left = Val(TxtCessPer.Left) - 1370
            Label1(62).Left = Val(Label1(62).Left) - 1370
            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 1370
            Label1(14).Left = Val(Label1(14).Left) - 1370
            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 1370
        ElseIf RSTCOMPANY!hide_mrp <> "Y" And RSTCOMPANY!hide_expiry = "Y" Then
            Label1(61).Visible = False
            TXTEXPIRY.Visible = False
'            Label1(24).Left = Val(Label1(24).Left) - 1300
'            TxtMRP.Left = Val(TXTTAX.Left) - 1300
            Label1(12).Left = Val(Label1(12).Left) - 690
            TXTTAX.Left = Val(TXTTAX.Left) - 690
            Label1(69).Left = Val(Label1(69).Left) - 690
            TxtKFC.Left = Val(TxtKFC.Left) - 690
            Lblprice(30).Left = Val(Lblprice(30).Left) - 690
            TXTRETAILNOTAX.Left = Val(TXTRETAILNOTAX.Left) - 690
            Label1(11).Left = Val(Label1(11).Left) - 690
            txtretail.Left = Val(txtretail.Left) - 690
            Label1(64).Left = Val(Label1(64).Left) - 690
            txtNetrate.Left = Val(txtNetrate.Left) - 690
            Label1(13).Left = Val(Label1(13).Left) - 690
            TXTDISC.Left = Val(TXTDISC.Left) - 690
            Label1(65).Left = Val(Label1(65).Left) - 690
            TxtDiscAmt.Left = Val(TxtDiscAmt.Left) - 690
            Label1(29).Left = Val(Label1(29).Left) - 690
            TxtCessPer.Left = Val(TxtCessPer.Left) - 690
            Label1(62).Left = Val(Label1(62).Left) - 690
            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 690
            Label1(14).Left = Val(Label1(14).Left) - 690
            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 690
        ElseIf RSTCOMPANY!hide_mrp = "Y" And RSTCOMPANY!hide_expiry <> "Y" Then
            Label1(24).Visible = False
            TxtMRP.Visible = False
            Label1(61).Left = Val(Label1(61).Left) - 1300
            TXTEXPIRY.Left = Val(TXTEXPIRY.Left) - 1300
            Label1(12).Left = Val(Label1(12).Left) - 1300
            TXTTAX.Left = Val(TXTTAX.Left) - 1300
             Label1(69).Left = Val(Label1(69).Left) - 1300
            TxtKFC.Left = Val(TxtKFC.Left) - 1300
            Lblprice(30).Left = Val(Lblprice(30).Left) - 1300
            TXTRETAILNOTAX.Left = Val(TXTRETAILNOTAX.Left) - 1300
            Label1(11).Left = Val(Label1(11).Left) - 1300
            txtretail.Left = Val(txtretail.Left) - 1300
            Label1(64).Left = Val(Label1(64).Left) - 1300
            txtNetrate.Left = Val(txtNetrate.Left) - 1300
            Label1(13).Left = Val(Label1(13).Left) - 1300
            TXTDISC.Left = Val(TXTDISC.Left) - 1300
            Label1(65).Left = Val(Label1(65).Left) - 1300
            TxtDiscAmt.Left = Val(TxtDiscAmt.Left) - 1300
            Label1(29).Left = Val(Label1(29).Left) - 1300
            TxtCessPer.Left = Val(TxtCessPer.Left) - 1300
            Label1(62).Left = Val(Label1(62).Left) - 1300
            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 1300
            Label1(14).Left = Val(Label1(14).Left) - 1300
            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 1300

        End If
        
        If RSTCOMPANY!hide_disc = "Y" Then
            Label1(65).Visible = False
            TxtDiscAmt.Visible = False
            Label1(29).Left = Val(Label1(29).Left) - 760
            TxtCessPer.Left = Val(TxtCessPer.Left) - 760
            Label1(62).Left = Val(Label1(62).Left) - 760
            TxtCessAmt.Left = Val(TxtCessAmt.Left) - 760
            Label1(14).Left = Val(Label1(14).Left) - 760
            LBLSUBTOTAL.Left = Val(LBLSUBTOTAL.Left) - 760
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
    
    If MDIMAIN.lblgst.Caption <> "R" Then
        OptQtn.Value = True
        OptPI.Visible = False
    End If
    BR_FLAG = True
    NEW_BILL = True
    SERIAL_FLAG = False
    lblactqty.Caption = ""
    lblbarcode.Caption = ""
    ACT_FLAG = True
    AGNT_FLAG = True
    M_ADD = False
    lblcredit.Caption = "0"
    txtcrdays.text = ""
    lblP_Rate.Caption = "0"
    flagchange2.Caption = ""
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
    grdsales.ColWidth(47) = 400
    
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
    grdsales.TextArray(47) = "KFC"
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
    grdsales.ColAlignment(47) = 4
    
    If frmLogin.rs!Level <> "0" Then
'        Label1(21).Visible = False
'        lblretail.Visible = False
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
    TxtKFC.Enabled = False
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
    TXTCNTYPE.text = ""
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
    If ITEM_REC.State = 1 Then ITEM_REC.Close
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
                If (MsgBox("AVAILABLE STOCK IS  " & i & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                    TXTQTY.SelStart = 0
                    TXTQTY.SelLength = Len(TXTQTY.text)
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
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE, CATEGORY, WARRANTY, WARRANTY_TYPE, MRP, P_LWS, PRICE5, PRICE6, PRICE7  From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE, CATEGORY, WARRANTY, WARRANTY_TYPE, MRP, P_LWS, PRICE5, PRICE6, PRICE7  From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
                    txtNetrate.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                Case 1
                    txtNetrate.text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)))
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)))
                Case 2
                    txtNetrate.text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", Val(GRDPOPUPITEM.Columns(13)))
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", Val(GRDPOPUPITEM.Columns(13)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", Val(GRDPOPUPITEM.Columns(13)))
                Case 3
                    txtNetrate.text = IIf(IsNull(GRDPOPUPITEM.Columns(20)), "", Val(GRDPOPUPITEM.Columns(20)))
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(20)), "", Val(GRDPOPUPITEM.Columns(20)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(20)), "", Val(GRDPOPUPITEM.Columns(20)))
                    If Val(txtretail.text) = 0 Then
                        txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                        txtNetrate.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                    End If
                Case 4
                    txtNetrate.text = IIf(IsNull(GRDPOPUPITEM.Columns(22)), "", Val(GRDPOPUPITEM.Columns(22)))
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(22)), "", Val(GRDPOPUPITEM.Columns(22)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(22)), "", Val(GRDPOPUPITEM.Columns(22)))
                Case 5
                    txtNetrate.text = IIf(IsNull(GRDPOPUPITEM.Columns(23)), "", Val(GRDPOPUPITEM.Columns(23)))
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(23)), "", Val(GRDPOPUPITEM.Columns(23)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(23)), "", Val(GRDPOPUPITEM.Columns(23)))
                Case 6
                    txtNetrate.text = IIf(IsNull(GRDPOPUPITEM.Columns(24)), "", Val(GRDPOPUPITEM.Columns(24)))
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(24)), "", Val(GRDPOPUPITEM.Columns(24)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(24)), "", Val(GRDPOPUPITEM.Columns(24)))
                Case Else
                    txtNetrate.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                    txtretail.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
                    TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
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
            lblprice5.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(22)), "", GRDPOPUPITEM.Columns(22))
            lblprice6.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(23)), "", GRDPOPUPITEM.Columns(23))
            lblprice7.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(24)), "", GRDPOPUPITEM.Columns(24))
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
            TXTITEMCODE.text = grdtmp.Columns(0)
            If Trim(TXTITEMCODE.text) = "" Then
                LblPack.text = "1"
                lblunit.text = "Nos"
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                TxtName1.Enabled = False
                TXTITEMCODE.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            End If
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
            TxtCessPer.text = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
            TxtCessAmt.text = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
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
            txtNetrate.text = (Val(lblLWPrice.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        ElseIf cmbtype.ListIndex = 2 Then
            txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        ElseIf cmbtype.ListIndex = 3 Then
            txtretail.text = (Val(LBLMRP.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(LBLMRP.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            txtNetrate.text = (Val(LBLMRP.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        ElseIf cmbtype.ListIndex = 4 Then
            txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        ElseIf cmbtype.ListIndex = 5 Then
            txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        ElseIf cmbtype.ListIndex = 6 Then
            txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        Else
            txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
        End If
    Else
        'TxtMRP.Text = Val(LBLMRP.Caption)
        If cmbtype.ListIndex = 1 Then
            txtretail.text = Val(lblwsale.Caption)
            TXTRETAILNOTAX.text = Val(lblwsale.Caption)
            txtNetrate.text = Val(lblwsale.Caption)
        ElseIf cmbtype.ListIndex = 2 Then
            txtretail.text = Val(lblvan.Caption)
            TXTRETAILNOTAX.text = Val(lblvan.Caption)
            txtNetrate.text = Val(lblvan.Caption)
        ElseIf cmbtype.ListIndex = 3 Then
            txtretail.text = Val(LBLMRP.Caption)
            TXTRETAILNOTAX.text = Val(LBLMRP.Caption)
            txtNetrate.text = Val(LBLMRP.Caption)
        ElseIf cmbtype.ListIndex = 4 Then
            txtretail.text = Val(lblprice5.Caption)
            TXTRETAILNOTAX.text = Val(lblprice5.Caption)
            txtNetrate.text = Val(lblprice5.Caption)
        ElseIf cmbtype.ListIndex = 5 Then
            txtretail.text = Val(lblprice6.Caption)
            TXTRETAILNOTAX.text = Val(lblprice6.Caption)
            txtNetrate.text = Val(lblprice6.Caption)
        ElseIf cmbtype.ListIndex = 6 Then
            txtretail.text = Val(lblprice7.Caption)
            TXTRETAILNOTAX.text = Val(lblprice7.Caption)
            txtNetrate.text = Val(lblprice7.Caption)
        Else
            txtretail.text = Val(lblretail.Caption)
            TXTRETAILNOTAX.text = Val(lblretail.Caption)
            txtNetrate.text = Val(lblretail.Caption)
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
        ElseIf Val(TxtMRP.text) <> 0 And Val(Round(Val(TxtMRP.text), 2)) = Val(Round(Val(txtretail.text), 2)) And mrpplus = False Then
            Call txtNetrate_LostFocus
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
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OPTTaxMRP_Click()
    TXTRETAILNOTAX_LostFocus
End Sub

Private Sub OPTVAT_Click()
    TXTRETAILNOTAX_LostFocus
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
            lblbalance.Caption = ""
            Txtrcvd.text = ""
            txtBillNo.Tag = "N"
'            Set TRXMAST = New ADODB.Recordset
'            TRXMAST.Open "Select * FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.Text) & " AND (ISNULL(BILL_FLAG) OR BILL_FLAG <>'Y') ", db, adOpenStatic, adLockReadOnly
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
            TRXFILE.Open "Select * FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " order by LINE_NO ", db, adOpenStatic, adLockReadOnly
            Do Until TRXFILE.EOF
                i = i + 1
                TXTINVDATE.text = Format(TRXFILE!VCH_DATE, "DD/MM/YYYY")
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = TRXFILE!ITEM_CODE
                grdsales.TextMatrix(i, 2) = TRXFILE!ITEM_NAME
                grdsales.TextMatrix(i, 3) = TRXFILE!QTY
'                Set TRXMAST = New ADODB.Recordset
'                TRXMAST.Open "SELECT UNIT FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(TRXSUB!R_TRX_TYPE) & "' AND RTRXFILE.VCH_NO = " & Val(TRXSUB!R_VCH_NO) & " AND RTRXFILE.LINE_NO = " & Val(TRXSUB!R_LINE_NO) & " AND RTRXFILE.TRX_YEAR = '" & Val(TRXSUB!R_TRX_YEAR) & "'", db, adOpenStatic, adLockReadOnly
'                If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
'                    grdsales.TextMatrix(i, 4) = Val(TRXMAST!UNIT)
'                End If
'                TRXMAST.Close
                Set TRXMAST = Nothing
                grdsales.TextMatrix(i, 4) = 1
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
'                grdsales.TextMatrix(i, 14) = Val(TRXSUB!R_VCH_NO)
'                grdsales.TextMatrix(i, 15) = Val(TRXSUB!R_LINE_NO)
'                grdsales.TextMatrix(i, 16) = Trim(TRXSUB!R_TRX_TYPE)
'                grdsales.TextMatrix(i, 43) = IIf(IsNull(TRXSUB!R_TRX_YEAR), "", Trim(TRXSUB!R_TRX_YEAR))
                grdsales.TextMatrix(i, 17) = IIf(IsNull(TRXFILE!check_flag), "", Trim(TRXFILE!check_flag))
                TXTDEALER.text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                'TxtCode.Text = TRXFILE!ACT_CODE
                'DataList2.Text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                LBLDATE.Caption = IIf(IsNull(TRXFILE!CREATE_DATE), Date, TRXFILE!CREATE_DATE)
                grdsales.TextMatrix(i, 19) = "B"
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
                'grdsales.TextMatrix(i, 37) = IIf(IsNull(TRXFILE!DN_LINENO), "", TRXFILE!DN_LINENO)
                grdsales.TextMatrix(i, 38) = IIf(IsDate(TRXFILE!EXP_DATE), Format(TRXFILE!EXP_DATE, "MM/YY"), "")
                grdsales.TextMatrix(i, 39) = IIf(IsNull(TRXFILE!RETAILER_PRICE), 0, TRXFILE!RETAILER_PRICE)
                grdsales.TextMatrix(i, 40) = IIf(IsNull(TRXFILE!CESS_PER), 0, TRXFILE!CESS_PER)
                grdsales.TextMatrix(i, 41) = IIf(IsNull(TRXFILE!cess_amt), 0, TRXFILE!cess_amt)
                grdsales.TextMatrix(i, 42) = IIf(IsNull(TRXFILE!BARCODE), "", TRXFILE!BARCODE)
                grdsales.TextMatrix(i, 44) = IIf(IsNull(TRXFILE!ITEM_SPEC), "", TRXFILE!ITEM_SPEC)
                grdsales.TextMatrix(i, 45) = IIf(IsNull(TRXFILE!PUR_TAX), "", TRXFILE!PUR_TAX)
                grdsales.TextMatrix(i, 46) = IIf(IsNull(TRXFILE!UN_BILL), "N", TRXFILE!UN_BILL)
                grdsales.TextMatrix(i, 47) = IIf(IsNull(TRXFILE!kfc_tax), "", TRXFILE!kfc_tax)
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
            TRXMAST.Open "Select * FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
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
                End If
                LBLRETAMT.Caption = IIf(IsNull(TRXMAST!ADD_AMOUNT), "", Format(TRXMAST!ADD_AMOUNT, "0.00"))
                lblcredit.Caption = "0"
                'txtcrdays.Text = IIf(IsNull(TRXMAST!cr_days), "", TRXMAST!cr_days)
                TxtBillName.text = IIf(IsNull(TRXMAST!BILL_NAME), "", TRXMAST!BILL_NAME)
                TxtBillAddress.text = IIf(IsNull(TRXMAST!BILL_ADDRESS), "", TRXMAST!BILL_ADDRESS)
                TxtVehicle.text = IIf(IsNull(TRXMAST!VEHICLE), "", TRXMAST!VEHICLE)
                'TxtOrder.Text = IIf(IsNull(TRXMAST!D_ORDER), "", TRXMAST!D_ORDER)
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
                'TxtCN.Text = IIf(IsNull(TRXMAST!CN_REF), "", TRXMAST!CN_REF)
                TxtCN.text = IIf(IsNull(TRXMAST!BILL_NO), "", TRXMAST!BILL_NO)
                TXTCNTYPE.text = IIf(IsNull(TRXMAST!BillType), "", TRXMAST!BillType)
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
                    Case "5"
                        cmbtype.ListIndex = 4
                        TXTTYPE.text = 5
                    Case "6"
                        cmbtype.ListIndex = 5
                        TXTTYPE.text = 6
                    Case "7"
                        cmbtype.ListIndex = 6
                        TXTTYPE.text = 7
                    Case Else
                        cmbtype.ListIndex = 0
                        TXTTYPE.text = 1
                End Select
                
                
'                If IsNull(TRXMAST!TERMS) Or TRXMAST!TERMS = "" Then
'                    chkTerms.value = 0
'                    Terms1.Text = ""
'                Else
'                    chkTerms.value = 1
'                    Terms1.Text = TRXMAST!TERMS
'                End If
'                Call CMBBRNCH_GotFocus
'                CMBBRNCH.Text = IIf(IsNull(TRXMAST!BR_NAME), "", TRXMAST!BR_NAME)
'                CMBBRNCH.BoundText = IIf(IsNull(TRXMAST!BR_CODE), "", TRXMAST!BR_CODE)
'                lblIGST.Caption = IIf(IsNull(TRXMAST!CUST_IGST), "N", TRXMAST!CUST_IGST)
                NEW_BILL = False
                OLD_BILL = True
                lbltime.Caption = Format(TRXMAST!C_TIME, "hh:mm AMPM")
            Else
                flagchange2.Caption = ""
                TxtCN.text = ""
                TXTCNTYPE.text = ""
                lbltime.Caption = Format(Time, "hh:mm AMPM")
                TXTTIN.text = ""
                CMBBRNCH.text = ""
                NEW_BILL = True
                OLD_BILL = False
                'TXTTYPE.Text = ""
                'cmbtype.ListIndex = -1
                cmbtype.ListIndex = 0
                TXTTYPE.text = 1
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            If OLD_BILL = False Then
                Dim rstBILL As ADODB.Recordset
                Set rstBILL = New ADODB.Recordset
                rstBILL.Open "Select MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'QT'", db, adOpenForwardOnly
                If Not (rstBILL.EOF And rstBILL.BOF) Then
                    cmbtype.Tag = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0))
                End If
                rstBILL.Close
                Set rstBILL = Nothing
                If Val(txtBillNo.text) < Val(cmbtype.Tag) And Val(cmbtype.Tag) <> 0 Then
                    Set rstBILL = New ADODB.Recordset
                    rstBILL.Open "Select * FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                    If (rstBILL.EOF And rstBILL.BOF) Then
                        rstBILL.AddNew
                        rstBILL!VCH_NO = Val(txtBillNo.text)
                        rstBILL!TRX_TYPE = "QT"
                        rstBILL!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                        rstBILL!C_USER_ID = frmLogin.rs!USER_ID
                        rstBILL!CREATE_DATE = Format(Date, "DD/MM/YYYY")
                        rstBILL!C_TIME = Format(Time, "HH:MM:SS")
                        rstBILL!C_USER_NAME = frmLogin.rs!USER_NAME
                    End If
                    rstBILL.Update
                    rstBILL.Close
                    Set rstBILL = Nothing
                    OLD_BILL = True
                End If
            End If
            'OLD_BILL = True
            LBLBILLNO.Caption = Val(txtBillNo.text)
            
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
                        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                        LBLFOT.Caption = ""
                    Case Else
                        If MDIMAIN.lblgst.Caption = "R" Then
                            LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                        End If
                        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
            
            TXTSLNO.text = grdsales.rows
            txtBillNo.Visible = False
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            If grdsales.rows > 1 Then
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
'    TRXMAST.Open "Select * FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.Text) & " AND (ISNULL(BILL_FLAG) OR BILL_FLAG <>'Y') ", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
'        txtBillNo.Tag = "Y"
'    Else
'        txtBillNo.Tag = "N"
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
    lblbalance.Caption = ""
    Txtrcvd.text = ""
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'QT'", db, adOpenStatic, adLockReadOnly
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
'    TRXMAST.Open "Select MIN(VCH_NO) FROM QTNSUB WHERE TRX_TYPE = 'QT'", db, adOpenStatic, adLockReadOnly
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
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                TxtName1.SetFocus
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
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            Else
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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
    Call CHANGEBOXCOLOR(TXTDEALER, False)
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
            
            If lblsubdealer.Caption = "D" Then
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
    
    TXTRETAILNOTAX.Tag = Val(TXTRETAILNOTAX.text) - (Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100)
    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(TXTTAX.text) + Val(TxtKFC.Caption)) / 100)) + (TXTRETAILNOTAX.Tag * Val(TxtCessPer.text) / 100) + Val(TxtCessAmt.text)
    'TxtCessPer.Tag = (TXTRETAILNOTAX.Tag * Val(TxtCessPer.Text) / 100) + Val(TxtCessAmt.Text)
    
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text), 2), ".000")
    Else
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(TXTQTY.text), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text) * Val(TXTQTY.text), 2), ".000")
    End If
    
    If MDIMAIN.lblgst.Caption = "R" Then Call cost_calculate
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
    TXTRETAILNOTAX.Tag = Val(TXTRETAILNOTAX.text) - (Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100)
    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(TXTTAX.text) + Val(TxtKFC.Caption)) / 100)) + (TXTRETAILNOTAX.Tag * Val(TxtCessPer.text) / 100) + Val(TxtCessAmt.text)
    'TxtCessPer.Tag = (TXTRETAILNOTAX.Tag * Val(TxtCessPer.Text) / 100) + Val(TxtCessAmt.Text)
    
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text), 2), ".000")
    Else
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(TXTQTY.text), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text) * Val(TXTQTY.text), 2), ".000")
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
            
            If lblsubdealer.Caption = "D" Then
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
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            Else
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY CATEGORY, ITEM_SLNO", db, adOpenStatic, adLockReadOnly
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE ucase(CATEGORY) = 'OWN' AND ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        If MDIMAIN.lblitemrepeat.Caption = "Y" Then
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7, ITEM_COST, ITEM_NET_COST From ITEMMAST WHERE ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '%" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        Else
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7, ITEM_COST, ITEM_NET_COST From ITEMMAST WHERE ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        End If
        PHYFLAG = False
    Else
        PHY.Close
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE ucase(CATEGORY) = 'OWN' AND ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' OR MRP Like '" & Trim(Me.TxtName1.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        If MDIMAIN.lblitemrepeat.Caption = "Y" Then
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7, ITEM_COST, ITEM_NET_COST From ITEMMAST WHERE ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '%" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        Else
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7, ITEM_COST, ITEM_NET_COST From ITEMMAST WHERE ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
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
    grdtmp.Columns(26).Width = 700
    grdtmp.Columns(27).Width = 900
    grdtmp.Columns(28).Width = 900
    grdtmp.Columns(29).Width = 900
    grdtmp.Columns(30).Width = 900
    grdtmp.Columns(31).Caption = "Cost"
    grdtmp.Columns(31).Width = 1000
    grdtmp.Columns(32).Caption = "Net Cost"
    grdtmp.Columns(32).Width = 1000
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
            If txtBillNo.Tag = "Y" Then
                MsgBox "Cannot modify here", vbOKOnly, "QUOTATION"
                Exit Sub
            End If
            If UCase(TxtName1.text) = "OT" Then TXTITEMCODE.text = "OT"
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            'TXTITEMCODE.Enabled = False
            'TxtName1.Enabled = False
            'TXTPRODUCT.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            Exit Sub
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
    txtNetrate.SelStart = 0
    txtNetrate.SelLength = Len(txtNetrate.text)
    Call CHANGEBOXCOLOR(txtNetrate, True)
    If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
End Sub

Private Sub txtNetrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.text) <> 0 And Val(txtNetrate.text) = 0 Then Exit Sub
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
                If TxtDiscAmt.Visible = True Then
                    TxtDiscAmt.Enabled = True
                    TxtDiscAmt.SetFocus
                Else
                    TXTDISC.Enabled = True
                    TXTDISC.SetFocus
                End If
            End If
        Case vbKeyEscape
            txtretail.Enabled = True
            txtretail.SetFocus
        Case 116
            Call FILL_PREVIIOUSRATE
        Case 117
            If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
        Case vbKeyDown
            If CMBBillType.ListIndex = -1 And M_EDIT = False And CMBBillType.Visible = True Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtNetrate_LostFocus()
    Call CHANGEBOXCOLOR(txtNetrate, False)
    If Val(txtNetrate.text) = 0 Then Exit Sub
    If Val(TxtCessPer.text) <> 0 Or Val(TxtCessAmt.text) <> 0 Then
        TXTRETAILNOTAX.text = (Val(txtNetrate.text) - Val(TxtCessAmt.text)) / (1 + ((Val(TXTTAX.text) + Val(TxtKFC.Caption)) / 100) + (Val(TxtCessPer.text) / 100))
        'TXTRETAILNOTAX.Text = (Val(txtNetrate.Text) - Val(TxtCessAmt.Text)) / (1 + ((Val(TXTTAX.Text) + Val(TxtKFC.Caption)) / 100) + (Val(TxtCessPer.Text) / 100))
        txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 4)
        TXTRETAILNOTAX.text = Round(Val(txtretail.text) * 100 / ((Val(TXTTAX.text)) + 100), 4)
        TXTRETAILNOTAX.text = Format(Val(TXTRETAILNOTAX.text), "0.0000")
    Else
        'TXTRETAILNOTAX.Text = (Val(txtNetrate.Text) - Val(TxtCessAmt.Text)) / (1 + ((Val(TXTTAX.Text) + Val(TxtKFC.Caption)) / 100) + (Val(TxtCessPer.Text) / 100))
        TXTRETAILNOTAX.text = (Val(txtNetrate.text) - Val(TxtCessAmt.text)) / (1 + ((Val(TXTTAX.text) + Val(TxtKFC.Caption)) / 100) + (Val(TxtCessPer.text) / 100))
        txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 4)
        TXTRETAILNOTAX.text = Round(Val(txtretail.text) * 100 / ((Val(TXTTAX.text)) + 100), 4)
        TXTRETAILNOTAX.text = Format(Val(TXTRETAILNOTAX.text), "0.0000")
    End If
    
    If MDIMAIN.lblgst.Caption = "R" Then Call cost_calculate
    Call CHANGEBOXCOLOR(txtNetrate, False)
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
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            Else
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If txtretail.Enabled = True Then txtretail.SetFocus
            If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY CATEGORY, ITEM_SLNO", db, adOpenStatic, adLockReadOnly
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE ucase(CATEGORY) = 'OWN'AND ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "' OR MRP Like '%" & Trim(Me.TxtName1.Text) & "') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        If Trim(TxtName1.text) = "" Then
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7, ITEM_COST, ITEM_NET_COST From ITEMMAST  WHERE (ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        Else
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7, ITEM_COST, ITEM_NET_COST From ITEMMAST  WHERE (ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%') AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.text) & "%' OR MRP Like '" & Trim(Me.TxtName1.text) & "') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        End If
        PHYFLAG = False
    Else
        PHY.Close
        'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION  From ITEMMAST  WHERE ucase(CATEGORY) = 'OWN'AND ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND (ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "' OR MRP Like '%" & Trim(Me.TxtName1.Text) & "') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        If Trim(TxtName1.text) = "" Then
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7, ITEM_COST, ITEM_NET_COST From ITEMMAST  WHERE (ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        Else
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, MANUFACTURER, BIN_LOCATION, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7, ITEM_COST, ITEM_NET_COST From ITEMMAST  WHERE (ITEM_CODE Like '" & Trim(Me.TxtName1.text) & "%' OR ITEM_NAME Like '%" & Trim(Me.TxtName1.text) & "%')  AND (ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.text) & "%' ) AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
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
    grdtmp.Columns(26).Width = 700
    grdtmp.Columns(27).Width = 900
    grdtmp.Columns(28).Width = 900
    grdtmp.Columns(29).Width = 900
    grdtmp.Columns(30).Width = 900
    grdtmp.Columns(31).Caption = "Cost"
    grdtmp.Columns(31).Width = 1000
    grdtmp.Columns(32).Caption = "Net Cost"
    grdtmp.Columns(32).Width = 1000
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
            
'            If frmLogin.rs!Level = "0" Then
'                If Val(lblcrlimit.Caption) > 0 And Val(lblOB.Caption) >= Val(lblcrlimit.Caption) Then
'                    If MsgBox("CREDIT LIMIT EXCEEDED. ARE YOU SURE YOU WANT TO CONTINUE...? ", vbYesNo + vbDefaultButton2, "EzBiz.....") = vbNo Then Exit Sub
'                End If
'            Else
'                If Val(lblcrlimit.Caption) > 0 And Val(lblOB.Caption) >= Val(lblcrlimit.Caption) Then
'                    MsgBox "CREDIT LIMIT EXCEEDED", vbOKOnly, "BILL"
'                    Exit Sub
'                End If
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
                    txtNetrate.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                Case 1
                    txtNetrate.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    txtretail.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                Case 2
                    txtNetrate.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    txtretail.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                Case 3
                    txtNetrate.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                    txtretail.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                    If Val(TXTRETAILNOTAX.text) = 0 Then
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    End If
                Case 4
                    txtNetrate.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                    txtretail.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                Case 5
                    txtNetrate.text = IIf(IsNull(grdtmp.Columns(29)), "", Val(grdtmp.Columns(29)))
                    txtretail.text = IIf(IsNull(grdtmp.Columns(29)), "", Val(grdtmp.Columns(29)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(29)), "", Val(grdtmp.Columns(29)))
                Case 6
                    txtNetrate.text = IIf(IsNull(grdtmp.Columns(30)), "", Val(grdtmp.Columns(30)))
                    txtretail.text = IIf(IsNull(grdtmp.Columns(30)), "", Val(grdtmp.Columns(30)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(30)), "", Val(grdtmp.Columns(30)))
                Case Else
                    txtNetrate.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
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
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC,  CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7 From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "' ", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                Else
                    PHY.Close
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC,  CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7  From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "' ", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                End If
                Set grdtmp.DataSource = PHY
                
                If PHY.RecordCount = 0 Then
                    If (MsgBox("Item not exists. Dou you want to continue", vbYesNo, "QUOTATION") = vbNo) Then Exit Sub
                    If Trim(TXTPRODUCT.text) = "" Then TXTPRODUCT.text = Trim(TxtName1.text)
                    TxtName1.Enabled = False
                    TXTPRODUCT.Enabled = False
                    TXTITEMCODE.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
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
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    Case 2
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    Case 3
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        If Val(TXTRETAILNOTAX.text) = 0 Then
                            txtNetrate.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                            txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                            TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        End If
                    Case 4
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                    Case 5
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                    Case 6
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                    Case Else
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
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
                lblprice5.Caption = IIf(IsNull(grdtmp.Columns(26)), "", grdtmp.Columns(26))
                lblprice6.Caption = IIf(IsNull(grdtmp.Columns(27)), "", grdtmp.Columns(27))
                lblprice7.Caption = IIf(IsNull(grdtmp.Columns(28)), "", grdtmp.Columns(28))
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
        
    If Trim(TXTTIN.text) = "" And Trim(lblIGST.Caption) <> "Y" And MDIMAIN.lblkfc.Caption = "Y" And IsDate(MDIMAIN.DTKFCSTART.Value) And IsDate(MDIMAIN.DTKFCEND.Value) Then
        If DateValue(TXTINVDATE.text) >= DateValue(MDIMAIN.DTKFCSTART.Value) And DateValue(TXTINVDATE.text) <= DateValue(MDIMAIN.DTKFCEND.Value) Then
            If Val(TXTTAX.text) = 12 Or Val(TXTTAX.text) = 18 Or Val(TXTTAX.text) = 28 Then
                TxtKFC.Caption = 1
            Else
                TxtKFC.Caption = 0
            End If
        End If
    Else
        TxtKFC.Caption = 0
    End If
    
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
            RSTITEMDISC.Open "SELECT CUST_DISC FROM ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "'", db, adOpenForwardOnly
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
                RSTTRXFILE.Open "SELECT CLOSE_QTY  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
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
                        If (MsgBox("AVAILABLE STOCK IN THE PARTICULAR STOCK IS  " & (Val(lblactqty.Caption) / Val(LblPack.text)) & " " & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                            TXTQTY.SelStart = 0
                            TXTQTY.SelLength = Len(TXTQTY.text)
                            Exit Sub
                        End If
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
            End If
SKIP:
                If UCase(TXTITEMCODE.text) = "OT" Then
                If TxtMRP.Visible = True Then
                    TxtMRP.Enabled = True
                    TxtMRP.SetFocus
                Else
                    If M_EDIT = False And CMBBillType.Visible = True Then
                        CMBBillType.Enabled = True
                        CMBBillType.SetFocus
                    Else
                        If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                            txtretail.Enabled = True
                            txtretail.SetFocus
                        Else
                            If MDIMAIN.LBLTAXWARN.Caption = "Y" And Val(TXTTAX.text) = 0 Then
                                TXTTAX.Enabled = True
                                TXTTAX.SetFocus
                            Else
                                If (Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(txtNetrate.text) And mrpplus = True) Then
                                    TXTRETAILNOTAX.Enabled = True
                                    TXTRETAILNOTAX.SetFocus
                                    Exit Sub
                                End If
                            
                                If MDIMAIN.LblKFCNet.Caption <> "N" Then
                                    txtNetrate.Enabled = True
                                    txtNetrate.SetFocus
                                Else
                                    If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                                        TXTRETAILNOTAX.Enabled = True
                                        TXTRETAILNOTAX.SetFocus
                                    Else
                                        If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
                                            TXTRETAILNOTAX.Enabled = True
                                            TXTRETAILNOTAX.SetFocus
                                        ElseIf Val(TxtMRP.text) <> 0 And Val(Round(Val(TxtMRP.text), 2)) = Val(Round(Val(txtretail.text), 2)) And mrpplus = False Then
                                            txtNetrate.Enabled = True
                                            txtNetrate.SetFocus
                                        Else
                                            txtretail.Enabled = True
                                            txtretail.SetFocus
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If M_EDIT = False And CMBBillType.Visible = True Then
                    CMBBillType.Enabled = True
                    CMBBillType.SetFocus
                Else
                    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                        txtretail.Enabled = True
                        txtretail.SetFocus
                    Else
                        If MDIMAIN.LBLTAXWARN.Caption = "Y" And Val(TXTTAX.text) = 0 Then
                            TXTTAX.Enabled = True
                            TXTTAX.SetFocus
                        Else
                            If (Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(txtNetrate.text) And mrpplus = True) Then
                                TXTRETAILNOTAX.Enabled = True
                                TXTRETAILNOTAX.SetFocus
                                Exit Sub
                            End If
                            
                            If MDIMAIN.LblKFCNet.Caption <> "N" Or cmbtype.ListIndex = 3 Or (Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(txtNetrate.text) And mrpplus = False) Then
                                txtNetrate.Enabled = True
                                txtNetrate.SetFocus
                            Else
                                If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                                    TXTRETAILNOTAX.Enabled = True
                                    TXTRETAILNOTAX.SetFocus
                                Else
                                    If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
                                        TXTRETAILNOTAX.Enabled = True
                                        TXTRETAILNOTAX.SetFocus
                                    ElseIf Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = False Then
                                        txtNetrate.Enabled = True
                                        txtNetrate.SetFocus
                                    Else
                                        txtretail.Enabled = True
                                        txtretail.SetFocus
                                    End If
                                End If
                            End If
                        End If
                    End If
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
            If CMBBillType.ListIndex = -1 And M_EDIT = False And CMBBillType.Visible = True Then Exit Sub
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
            TRXMAST.Open "SELECT PACK_TYPE FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
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
            TRXMAST.Open "SELECT FULL_PACK FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
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
    TXTRETAILNOTAX.Tag = Val(TXTRETAILNOTAX.text) - (Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100)
    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(TXTTAX.text) + Val(TxtKFC.Caption)) / 100)) + (TXTRETAILNOTAX.Tag * Val(TxtCessPer.text) / 100) + Val(TxtCessAmt.text)
    'TxtCessPer.Tag = (TXTRETAILNOTAX.Tag * Val(TxtCessPer.Text) / 100) + Val(TxtCessAmt.Text)
    
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text), 2), ".000")
    Else
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(TXTQTY.text), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text) * Val(TXTQTY.text), 2), ".000")
    End If
    
    On Error GoTo ERRHAND
    Set RSTITEMCOST = New ADODB.Recordset
    RSTITEMCOST.Open "SELECT ITEM_COST, SALES_PRICE, SALES_TAX, UN_BILL FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
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
        RSTITEMCOST.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.text) & " AND RTRXFILE.TRX_YEAR = '" & Val(TrxRYear.text) & "' AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly, adCmdText
        With RSTITEMCOST
            If Not (.EOF And .BOF) Then
                LBLITEMCOST.Caption = IIf(IsNull(RSTITEMCOST!ITEM_COST), "", RSTITEMCOST!ITEM_COST * Val(LblPack.text))
                LBLNETCOST.Caption = Round(Val(LBLITEMCOST.Caption) + (Val(LBLITEMCOST.Caption) * IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX / 100)), 2)
                LBLNETCOST.Tag = IIf(IsNull(RSTITEMCOST!SALES_TAX), 0, RSTITEMCOST!SALES_TAX)
            Else
                RSTITEMCOST.Close
                Set RSTITEMCOST = Nothing
                Set RSTITEMCOST = New ADODB.Recordset
                RSTITEMCOST.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockReadOnly, adCmdText
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
                    RSTITEMCOST.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
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
                MsgBox "Commission Rate greater than actual Rate", vbOKOnly, "QUOTATION"
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
            If Val(TXTSLNO.text) < grdsales.rows And frmLogin.rs!Level <> "0" And NEW_BILL = False Then
                MsgBox "Permission Denied", vbOKOnly, "QUOTATION"
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
                
                Dim discamnt As Double
                discamnt = 0
                If UCase(txtcategory.text) = "SERVICE CHARGE" Then
                    txtNetrate.text = Format(Val(txtretail.text), "0.00")
                Else
            '        discamnt = Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
            '        txtNetrate.Text = Format(Round(Val(TXTRETAILNOTAX.Text) - discamnt, 2), ".00")
                    
            '        txtNetrate.Text = Val(txtNetrate.Text) + (Val(TXTRETAILNOTAX.Text) - (Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100)) * Val(TxtCessPer) / 100
            '        txtNetrate.Text = Val(txtNetrate.Text) + Round(Val(TxtCessAmt.Text), 3)
                    
                    txtNetrate.text = (Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * (Val(TXTTAX.text) + Val(TxtKFC.Caption)) / 100)) + (Val(TXTRETAILNOTAX.text) * Val(TxtCessPer.text) / 100) + Val(TxtCessAmt.text)
                    txtNetrate.text = Format(Val(txtNetrate.text), "0.00")
                End If
                
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
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            Else
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
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

Private Sub TxtTax_Change()
    If Trim(TXTTIN.text) = "" And Trim(lblIGST.Caption) <> "Y" And MDIMAIN.lblkfc.Caption = "Y" And IsDate(MDIMAIN.DTKFCSTART.Value) And IsDate(MDIMAIN.DTKFCEND.Value) Then
        If DateValue(TXTINVDATE.text) >= DateValue(MDIMAIN.DTKFCSTART.Value) And DateValue(TXTINVDATE.text) <= DateValue(MDIMAIN.DTKFCEND.Value) Then
            If Val(TXTTAX.text) = 12 Or Val(TXTTAX.text) = 18 Or Val(TXTTAX.text) = 28 Then
                TxtKFC.Caption = 1
            Else
                TxtKFC.Caption = 0
            End If
        End If
    Else
        TxtKFC.Caption = 0
    End If
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
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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

Private Function COSTCALCULATION()
    Dim COST As Double
    Dim n As Integer
    'Dim RSTITEMMAST As ADODB.Recordset
    
    LBLTOTALCOST.Caption = ""
    LBLPROFIT.Caption = ""
    COST = 0
    On Error GoTo ERRHAND

    If MDIMAIN.lblgst.Caption = "R" Then
        For n = 1 To grdsales.rows - 1
            LBLTOTALCOST.Caption = Val(LBLTOTALCOST.Caption) + (Val(grdsales.TextMatrix(n, 11)) * (Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20))))
            
            COST = (Val(grdsales.TextMatrix(n, 11))) '* (Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20))))
            TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(n, 6)) - (Val(grdsales.TextMatrix(n, 6)) * Val(grdsales.TextMatrix(n, 8)) / 100)
            LBLPROFIT.Caption = Val(LBLPROFIT.Caption) + (Val(TXTRETAILNOTAX.Tag) - COST) * (Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20)))
        Next n
        
        LBLTOTALCOST.Caption = Round(Val(LBLTOTALCOST.Caption), 2)
        LBLPROFIT.Caption = Round(Val(LBLPROFIT.Caption) - (Val(LBLDISCAMT.Caption) + Val(lblcomamt.text)), 2)
        LBLTOTALCOST.Caption = Format(Val(LBLTOTALCOST.Caption), "0.00")
        LBLPROFIT.Caption = Format(Val(LBLPROFIT.Caption), "0.00")
    Else
    
        For n = 1 To grdsales.rows - 1
            'COST = COST + (Val(grdsales.TextMatrix(N, 11)) * Val(grdsales.TextMatrix(N, 3)))
            COST = COST + ((Val(grdsales.TextMatrix(n, 11)) + (Val(grdsales.TextMatrix(n, 11)) * Val(grdsales.TextMatrix(n, 45)) / 100)) * (Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20))))
        Next n
        
        LBLTOTALCOST.Caption = Round(COST, 2)
        LBLPROFIT.Caption = Round(Val(LBLTOTAL.Caption) - Val(LBLDISCAMT.Caption) - COST, 2)
    End If
    Exit Function
    
ERRHAND:
    MsgBox err.Description
End Function

Private Function COSTCALCULATE()
    
    If MDIMAIN.lblgst.Caption = "R" Then
        LBLNETCOST.Caption = Round(((Val(grdsales.TextMatrix(grdsales.Row, 11)) + (Val(grdsales.TextMatrix(grdsales.Row, 11)) * Val(grdsales.TextMatrix(grdsales.Row, 45)) / 100))), 2)
        LBLITEMCOST.Caption = Val(grdsales.TextMatrix(grdsales.Row, 11))
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
    Else
        LBLNETCOST.Caption = Round(((Val(grdsales.TextMatrix(grdsales.Row, 11)) + (Val(grdsales.TextMatrix(grdsales.Row, 11)) * Val(grdsales.TextMatrix(grdsales.Row, 45)) / 100))), 2)
        If Val(LBLNETCOST.Caption) = 0 Then
            LblProfitPerc.Caption = 0
        Else
            If Val(grdsales.TextMatrix(grdsales.Row, 3)) = 0 Then
                LblProfitPerc.Caption = 0
            Else
                LblProfitPerc.Caption = Round((((Val(grdsales.TextMatrix(grdsales.Row, 12)) / Val(grdsales.TextMatrix(grdsales.Row, 3))) - Val(LBLNETCOST.Caption)) * 100) / Val(LBLNETCOST.Caption), 2)
            End If
        End If
        If Val(grdsales.TextMatrix(grdsales.Row, 3)) = 0 Then
            LBLNETPROFIT.Caption = 0
        Else
            LBLNETPROFIT.Caption = Round(((Val(grdsales.TextMatrix(grdsales.Row, 12)) / Val(grdsales.TextMatrix(grdsales.Row, 3))) - Val(LBLNETCOST.Caption)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 2)
        End If
        LBLNETCOST.Caption = Format(Val(LBLNETCOST.Caption), "0.00")
        LblProfitPerc.Caption = Format(Val(LblProfitPerc.Caption), "0.00")
        LBLNETPROFIT.Caption = Format(Val(LBLNETPROFIT.Caption), "0.00")
    End If
    
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
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            Else
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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
    
    Call COSTCALCULATION
    'LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - Val(LBLTOTALCOST.Caption), 2)
    
End Sub

Private Function ReportGeneratION()
    
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


    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr (106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
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
        If Trim(TXTTIN.text) <> "" Then
            Print #1, Space(20) & "FORM NO. 8 See rule 58(10), TAX INVOICE"
        Else
            Print #1, Space(20) & "FORM NO. 8B See rule 58(10), RETAIL INVOICE"
        End If
        Print #1, Space(32) & AlignLeft("CASH / CREDIT SALE", 25)
        Print #1, RepeatString("-", 80)
        'Print #1, "D.N. NO & Date" & Space(5) & "P.O. NO. & Date" & Space(5) & "D.Doc.NO & Date" & Space(5) & "Del Terms" & Space(5) & "Veh. No"
        Print #1,
        Print #1, RepeatString("-", 80)
        'Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & AlignLeft("INVOICE FORM 8H", 16)
    
        'If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, LBLDATE.Caption)
        Print #1, "Qtn No. " & Trim(txtBillNo.text) & Space(2) & AlignRight("Date:" & TXTINVDATE.text, 67) '& Space(2) & LBLTIME.Caption
        Print #1, "TO: " & TxtBillName.text
        If Trim(TxtBillAddress.text) <> "" Then Print #1, TxtBillAddress.text
        If Trim(TxtPhone.text) <> "" Then Print #1, "Phone: " & TxtPhone.text
        If Trim(TXTTIN.text) <> "" Then Print #1, "TIN: " & TXTTIN.text
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
    
        For i = 1 To grdsales.rows - 1
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
        If Val(Txthandle.text) <> 0 Then
            Print #1, AlignRight("HANDLING CHARGE ", 65) & AlignRight((Format(Txthandle.text, "####.00")), 12)
        End If
        If Val(TxtFrieght.text) <> 0 Then
            Print #1, AlignRight("FRIEGHT CHARGE  ", 65) & AlignRight((Format(Txthandle.text, "####.00")), 12)
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
        If Trim(TXTTIN.text) <> "" Then
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

ERRHAND:
    MsgBox err.Description
End Function

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
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            Else
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
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
                MsgBox "Commission Rate greater than actual Rate", vbOKOnly, "QUOTATION"
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
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  FROM TRXFILE  WHERE (TRX_TYPE ='GI' OR TRX_TYPE ='HI' OR TRX_TYPE ='WO' OR TRX_TYPE ='SV') AND ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  FROM TRXFILE  WHERE (TRX_TYPE ='GI' OR TRX_TYPE ='HI' OR TRX_TYPE ='WO' OR TRX_TYPE ='SV') AND ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
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
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7 From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y')", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7 From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.text & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y')", db, adOpenStatic, adLockReadOnly
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
                    'PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_TAX, LINE_DISC, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, ITEM_SIZE, ITEM_COLOR, BARCODE, REF_NO, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, SALES_PRICE, P_WS, CRTN_PACK, P_CRTN, MRP, TAX_MODE, EXP_DATE  From RTRXFILE  WHERE BARCODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_TAX, LINE_DISC, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, BARCODE, REF_NO, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, SALES_PRICE, P_WS, P_VAN, CRTN_PACK, P_CRTN, MRP, TAX_MODE, EXP_DATE, TRX_YEAR, P_LWS, CUST_DISC, PRICE5, PRICE6, PRICE7, TRX_TYPE  From RTRXFILE  WHERE BARCODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ORDER BY TRX_TYPE DESC", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                Else
                    PHY.Close
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_TAX, LINE_DISC, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, BARCODE, REF_NO, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, SALES_PRICE, P_WS, P_VAN, CRTN_PACK, P_CRTN, MRP, TAX_MODE, EXP_DATE, TRX_YEAR, P_LWS, CUST_DISC, PRICE5, PRICE6, PRICE7, TRX_TYPE  From RTRXFILE  WHERE BARCODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ORDER BY TRX_TYPE DESC", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                End If
                Set grdtmp.DataSource = PHY
                If PHY.RecordCount = 0 Then
                    If PHYFLAG = True Then
                        PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7 From ITEMMAST  WHERE BARCODE = '" & Me.TXTITEMCODE.text & "' ORDER BY UN_BILL DESC", db, adOpenStatic, adLockReadOnly
                        PHYFLAG = False
                    Else
                        PHY.Close
                        PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7 From ITEMMAST  WHERE BARCODE = '" & Me.TXTITEMCODE.text & "' ORDER BY UN_BILL DESC", db, adOpenStatic, adLockReadOnly
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
                                txtNetrate.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                                txtretail.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                            Case 2
                                txtNetrate.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                                txtretail.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                            Case 3
                                txtNetrate.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                                txtretail.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                                If Val(txtretail.text) = 0 Then
                                    txtNetrate.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                                    txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                                    TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                                End If
                            Case 4
                                txtNetrate.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                                txtretail.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                            Case 5
                                txtNetrate.text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                                txtretail.text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                            Case 6
                                txtNetrate.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                                txtretail.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                                TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                            Case Else
                                txtNetrate.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
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
                        
                        If MDIMAIN.lblPriceSplit.Caption = "Y" Then GoTo SKIP_BATCH
                        Set RSTBATCH = New ADODB.Recordset
                        Select Case cmbtype.ListIndex
                            Case 1
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            Case 2
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_VAN >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            Case 3
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, MRP, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            Case 4
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, PRICE5, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            Case 5
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, PRICE6, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            Case 6
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, PRICE7, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                            Case Else
                                'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_RETAIL >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                                'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                                
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
SKIP_BATCH:
                        'Call CONTINUE
                        If Val(LblPack.text) > 1 Or cmbtype.ListIndex <> 0 Then
                            If MDIMAIN.LblKFCNet.Caption <> "N" Then
                                Call txtNetrate_LostFocus
                                Call TXTDISC_LostFocus
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
                            End If
                            TXTQTY.Enabled = True
                            TXTQTY.SetFocus
                            TxtName1.Enabled = False
                            TXTITEMCODE.Enabled = False
                            TXTPRODUCT.Enabled = False
                            TXTQTY.Enabled = True
                            TXTQTY.SetFocus
                            Exit Sub
                        End If
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
                
                If MDIMAIN.lblPriceSplit.Caption = "Y" Then GoTo SKIP_BATCH_2
                Set RSTBATCH = New ADODB.Recordset
                Select Case cmbtype.ListIndex
                    Case 1
                        RSTBATCH.Open "Select DISTINCT BARCODE, P_WS From RTRXFILE WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 2
                        RSTBATCH.Open "Select DISTINCT BARCODE, P_VAN From RTRXFILE WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 3
                        RSTBATCH.Open "Select DISTINCT BARCODE, MRP From RTRXFILE WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 4
                        RSTBATCH.Open "Select DISTINCT BARCODE, PRICE5 From RTRXFILE WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 5
                        RSTBATCH.Open "Select DISTINCT BARCODE, PRICE6 From RTRXFILE WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 6
                        RSTBATCH.Open "Select DISTINCT BARCODE, PRICE7 From RTRXFILE WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case Else
                        RSTBATCH.Open "Select DISTINCT BARCODE, P_RETAIL From RTRXFILE WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
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
SKIP_BATCH_2:
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
                LBLITEMCOST.Caption = grdtmp.Columns(20)
                TXTDISC.text = IIf(IsNull(grdtmp.Columns(31)), "", grdtmp.Columns(31))
                LblPack.text = IIf(IsNull(grdtmp.Columns(11)) Or Val(grdtmp.Columns(11)) = 0, "1", grdtmp.Columns(11))
                lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(11)) Or Val(grdtmp.Columns(11)) = 0, "1", grdtmp.Columns(11))
                Select Case cmbtype.ListIndex
                    Case 1
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(22)), "", Val(grdtmp.Columns(22)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(22)), "", Val(grdtmp.Columns(22)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(22)), "", Val(grdtmp.Columns(22)))
                    Case 2
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(23)), "", Val(grdtmp.Columns(23)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(23)), "", Val(grdtmp.Columns(23)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(23)), "", Val(grdtmp.Columns(23)))
                    Case 3
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                        If Val(txtretail.text) = 0 Then
                            txtNetrate.text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                            txtretail.text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                            'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                        End If
                    Case 4
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(32)), "", Val(grdtmp.Columns(32)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(32)), "", Val(grdtmp.Columns(32)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                    Case 5
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(33)), "", Val(grdtmp.Columns(33)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(33)), "", Val(grdtmp.Columns(33)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                    Case 6
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(34)), "", Val(grdtmp.Columns(34)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(34)), "", Val(grdtmp.Columns(34)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                    Case Else
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                        'TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(5)), "", Val(grdtmp.Columns(5)))
                End Select
                LBLSELPRICE.Caption = Val(txtretail.text)
                lblretail.Caption = IIf(IsNull(grdtmp.Columns(5)), "", grdtmp.Columns(5))
                lblwsale.Caption = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
                lblvan.Caption = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
                LBLMRP.Caption = IIf(IsNull(grdtmp.Columns(26)), "", grdtmp.Columns(26))
                lblprice5.Caption = IIf(IsNull(grdtmp.Columns(32)), "", grdtmp.Columns(32))
                lblprice6.Caption = IIf(IsNull(grdtmp.Columns(33)), "", grdtmp.Columns(33))
                lblprice7.Caption = IIf(IsNull(grdtmp.Columns(34)), "", grdtmp.Columns(34))
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
                RSTtax.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
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
                                txtcategory.text = IIf(IsNull(RSTtax!Category), "", RSTtax!Category)
                                If Val(TXTDISC.text) = 0 Then TXTDISC.text = IIf(IsNull(RSTtax!CUST_DISC), "", RSTtax!CUST_DISC)
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
                    If MDIMAIN.LblKFCNet.Caption <> "N" Then
                        Call txtNetrate_LostFocus
                        Call TXTDISC_LostFocus
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
                    End If
                    
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
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    Case 2
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    Case 3
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                        If Val(txtretail.text) = 0 Then
                            txtNetrate.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                            txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                            TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        End If
                    Case 4
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                    Case 5
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                    Case 6
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                    Case Else
                        txtNetrate.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        txtretail.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        TXTRETAILNOTAX.text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                End Select
                LblPack.text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                TXTDISC.text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
                'TXTEXPIRY.Text = IIf(isdate(grdtmp.Columns(22)),Format(grdtmp.Columns(22), "MM/YY"),"  /  ")
                TXTITEMCODE.text = grdtmp.Columns(0)
                
                If MDIMAIN.LBLITMWRN.Caption = "Y" Then
                    For i = 1 To grdsales.rows - 1
                        If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.text) Then
                            If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                                Exit Sub
                            Else
                                Exit For
                            End If
                        End If
                    Next i
                End If
                
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
                
                If MDIMAIN.lblPriceSplit.Caption = "Y" Then GoTo SKIP_BATCH_3
                Set RSTBATCH = New ADODB.Recordset
                Select Case cmbtype.ListIndex
                    Case 1
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 2
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_VAN >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_VAN, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 3
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, MRP, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 4
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, PRICE5, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 5
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, PRICE6, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case 6
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_WS, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_WS >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, PRICE7, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                    Case Else
                        'RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, EXP_DATE From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 AND (P_RETAIL >0 OR NOT ISNULL(EXP_DATE))", db, adOpenStatic, adLockReadOnly
                        'RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        RSTBATCH.Open "Select DISTINCT ITEM_CODE, P_RETAIL, REF_NO From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
                        
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
SKIP_BATCH_3:
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

'Function FILL_BATCHGRID()
'    FRMEMAIN.Enabled = False
'    FRMEGRDTMP.Visible = True
'
'    Set GRDPOPUP.DataSource = Nothing
'    Set GRDPOPUPITEM.DataSource = Nothing
'    FRMEITEM.Visible = False
'
'    If BATCH_FLAG = True Then
'        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME, PRICE5, PRICE6, PRICE7 From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
'        BATCH_FLAG = False
'    Else
'        PHY_BATCH.Close
'        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME, PRICE5, PRICE6, PRICE7 From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
'        BATCH_FLAG = False
'    End If
'
'    Set GRDPOPUP.DataSource = PHY_BATCH
'    GRDPOPUP.Columns(0).Caption = "Serial No."
'    GRDPOPUP.Columns(1).Caption = "QTY"
'    GRDPOPUP.Columns(2).Caption = "VCH No"
'    GRDPOPUP.Columns(3).Caption = "Line No"
'    GRDPOPUP.Columns(4).Caption = "Trx Type"
'    GRDPOPUP.Columns(7).Caption = "" '"Warranty"
'    GRDPOPUP.Columns(8).Caption = ""
'    GRDPOPUP.Columns(9).Caption = "RT"
'    GRDPOPUP.Columns(10).Caption = "WS"
'    GRDPOPUP.Columns(11).Caption = "VN"
'    GRDPOPUP.Columns(21).Caption = "MRP"
'    GRDPOPUP.Columns(25).Caption = "Expiry"
'    'GRDPOPUP.Columns(32).Caption = "Godown"
'
'    GRDPOPUP.Columns(0).Width = 3000
'    GRDPOPUP.Columns(1).Width = 900
'    GRDPOPUP.Columns(2).Width = 0
'    GRDPOPUP.Columns(3).Width = 0
'    GRDPOPUP.Columns(4).Width = 0
'    GRDPOPUP.Columns(5).Width = 0
'    GRDPOPUP.Columns(6).Width = 0
'    GRDPOPUP.Columns(7).Width = 0
'    GRDPOPUP.Columns(8).Width = 0
'    GRDPOPUP.Columns(9).Width = 1000
'    GRDPOPUP.Columns(10).Width = 1000
'    GRDPOPUP.Columns(11).Width = 1000
'    GRDPOPUP.Columns(12).Width = 0
'    GRDPOPUP.Columns(13).Width = 0
'    GRDPOPUP.Columns(14).Width = 0
'    GRDPOPUP.Columns(15).Width = 0
'    GRDPOPUP.Columns(16).Width = 0
'    GRDPOPUP.Columns(17).Width = 0
'    GRDPOPUP.Columns(18).Width = 0
'    GRDPOPUP.Columns(19).Width = 0
'    GRDPOPUP.Columns(20).Width = 0
'    GRDPOPUP.Columns(21).Width = 1000
'    GRDPOPUP.Columns(22).Width = 0
'    GRDPOPUP.Columns(23).Width = 0
'    GRDPOPUP.Columns(24).Width = 0
'    GRDPOPUP.Columns(25).Width = 1200
'    GRDPOPUP.Columns(26).Width = 0
'    GRDPOPUP.Columns(27).Width = 0
'    GRDPOPUP.Columns(28).Width = 0
'    GRDPOPUP.Columns(29).Width = 0
'    GRDPOPUP.Columns(30).Width = 0
'    GRDPOPUP.Columns(31).Width = 0
'    GRDPOPUP.Columns(32).Width = 1000
'    GRDPOPUP.Columns(33).Width = 1000
'    GRDPOPUP.Columns(34).Width = 1000
'    GRDPOPUP.SetFocus
'
'    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).Text
'    LBLHEAD(9).Visible = True
'    LBLHEAD(0).Visible = True
'
'
'End Function

Function FILL_PREVIIOUSRATE2()
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        'PHY_PRERATE.Open "Select TOP 10 ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  FROM QTNSUB  WHERE TRX_TYPE ='QT' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, LINE_DISC, VCH_NO, ITEM_NAME, M_USER_ID  FROM TRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC LIMIT 10", db, adOpenForwardOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        'PHY_PRERATE.Open "Select TOP 10 ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  FROM QTNSUB  WHERE TRX_TYPE ='QT' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, LINE_DISC, VCH_NO, ITEM_NAME, M_USER_ID  FROM TRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC LIMIT 10", db, adOpenForwardOnly
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
        LBLHEAD(2).Caption = GRDPRERATE.Columns(7).text
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
                If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                    TxtName1.SetFocus
                Else
                    TXTITEMCODE.SetFocus
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
    ElseIf Val(TXTTYPE.text) = 5 Then
        cmbtype.ListIndex = 4
    ElseIf Val(TXTTYPE.text) = 6 Then
        cmbtype.ListIndex = 5
    ElseIf Val(TXTTYPE.text) = 7 Then
        cmbtype.ListIndex = 6
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
     If Val(TXTTYPE.text) = 0 Or Val(TXTTYPE.text) > 7 Then
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
    If cmbtype.ListIndex = 4 And Val(TXTTYPE.text) <> 5 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 5 And Val(TXTTYPE.text) <> 6 Then
        MsgBox "Bill type doesnot match", vbOKOnly, "Sales"
        cmbtype.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = 6 And Val(TXTTYPE.text) <> 7 Then
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
                If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                    TxtName1.SetFocus
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
                MsgBox "Please enter Period for Warranty", , "QUOTATION"
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
    rstBILL.Open "Select MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'QT'", db, adOpenForwardOnly
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

Private Function ReportGeneratION_estimate()
    
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


    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr (106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
        Print #1, AlignLeft("QUOTATION", 25)
        Print #1, RepeatString("-", 80)
        Print #1, AlignLeft("Sl", 2) & Space(1) & _
                AlignLeft("Comm Code", 14) & Space(1) & _
                AlignLeft("Description", 35) & _
                AlignLeft("Qty", 4) & Space(3) & _
                AlignLeft("Rate", 10) & Space(3) & _
                AlignLeft("Amount", 12) '& _
                Chr (27) & Chr(72) '//Bold Ends
    
        Print #1, RepeatString("-", 80)
    
        For i = 1 To grdsales.rows - 1
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
        If Val(Txthandle.text) <> 0 Then
            Print #1, AlignRight("HANDLING CHARGE ", 65) & AlignRight((Format(Txthandle.text, "####.00")), 12)
        End If
        If Val(TxtFrieght.text) <> 0 Then
            Print #1, AlignRight("FRIEGHT CHARGE  ", 65) & AlignRight((Format(Txthandle.text, "####.00")), 12)
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

ERRHAND:
    MsgBox err.Description
End Function

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
        
    Dim PRNSPACE As Single
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        PRNSPACE = Int(54 - Len(IIf(IsNull(RSTCOMPANY!COMP_NAME) Or RSTCOMPANY!COMP_NAME = "", 1, RSTCOMPANY!COMP_NAME)) / 2)
        If PRNSPACE < 0 Then PRNSPACE = 0
        Print #1, Chr(27) & Chr(72) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(72) & Chr(10) & _
              Space(PRNSPACE) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) & _
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
                
        Print #1, "State: Kerala (32 - KL)"
        Print #1, RepeatString("-", 125)
        'Print #1,
        '''Print #1,  "TIN No. " & RSTCOMPANY!KGST
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    'Print #1, Space(31) & "The KVAT Rules 2005"
    PRNSPACE = Int(27)
    Print #1, Chr(27) & Chr(72) & Chr(10) & _
               Chr(14) & Chr(15) & Space(PRNSPACE) & "QUOTATON" & _
              Chr(27) & Chr(72)
    'Print #1, Space(57) & "QUOTATON"
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
    'Print #1, "D.N. NO & Date" & Space(10) & "P.O. NO. & Date" & Space(15) & "D.Doc.NO & Date" & Space(15) & "Del Terms" & Space(15) & "Veh. No"
    Print #1,
    Print #1, RepeatString("-", 125)
    'Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & AlignLeft("INVOICE FORM 8H", 16)

    'If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, LBLDATE.Caption)
    Print #1, "Qtn No. " & Trim(txtBillNo.text) & Space(2) & AlignRight("Date:" & TXTINVDATE.text, 110)   '& Space(2) & LBLTIME.Caption
    Print #1, "TO: " & TxtBillName.text
    If Trim(TxtBillAddress.text) <> "" Then Print #1, TxtBillAddress.text
    If Trim(TxtPhone.text) <> "" Then Print #1, "Phone: " & TxtPhone.text
    If Trim(TXTTIN.text) <> "" Then Print #1, "GSTIN: " & TXTTIN.text
    'LBLDATE.Caption = Date

   ' Print #1, Chr(27) & Chr(72) &  "Salesman: CS"

    Print #1, RepeatString("-", 125)
    Print #1, AlignLeft("Description", 50) & _
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
    Dim TotalTax, TaxAmt, DISCAMT, Totaldiscamt As Double
    Dim HSNCODE As String
    Dim RSTHSNCODE As ADODB.Recordset
    TaxAmt = 0
    TotalTax = 0
    DISCAMT = 0
    Totaldiscamt = 0
    For i = 1 To grdsales.rows - 1
        DISCAMT = (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 8)) / 100) * Val(grdsales.TextMatrix(i, 3))
        Totaldiscamt = Totaldiscamt + DISCAMT
        If Val(creditbill.grdsales.TextMatrix(i, 9)) > 0 Then
            TaxAmt = Round((Val(grdsales.TextMatrix(i, 6)) - DISCAMT) * Val(grdsales.TextMatrix(i, 9)) / 100, 2)
        End If
        TotalTax = TotalTax + TaxAmt
        
        Set RSTHSNCODE = New ADODB.Recordset
        RSTHSNCODE.Open "SELECT * from ITEMMAST WHERE ITEM_CODE = '" & creditbill.grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTHSNCODE.EOF And RSTHSNCODE.BOF) Then
            HSNCODE = IIf(IsNull(RSTHSNCODE!REMARKS), "", RSTHSNCODE!REMARKS)
        Else
            HSNCODE = ""
        End If
        RSTHSNCODE.Close
        Set RSTHSNCODE = Nothing
        
        Print #1, AlignLeft(grdsales.TextMatrix(i, 2), 50) & Space(0) & _
            AlignLeft(HSNCODE, 8) & _
            AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 8) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)), 2), "0.00"), 9) & _
            AlignRight(Format(Round(DISCAMT, 2), "0.00"), 8) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 9)) / 2, 2), "0.00"), 8) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 9)) / 2, 2), "0.00"), 8) & _
            AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 9)) / 100, 2), "0.00"), 9) & _
            AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 13) '& _
            Chr (27) & Chr(72) '//Bold Ends
    Next i
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, RepeatString("-", 125)
    Print #1, "TOTAL:" & AlignRight(Format(Round(Totaldiscamt, 2), "0.00"), 77) & AlignRight(Format(Round(TotalTax / 2, 2), "0.00"), 8) & AlignRight(Format(Round(TotalTax / 2, 2), "0.00"), 8) & AlignRight(Format(Round(TotalTax, 2), "0.00"), 9) & AlignRight(Format(Round(Val(LBLTOTAL.Caption), 2), "0.00"), 13)
    
    Print #1, RepeatString("-", 125)
    
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
    Num = CCur(Round(LBLTOTAL.Caption, 0))
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
    RSTTRXFILE.Open "SELECT DISTINCT SALES_TAX FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND SALES_TAX > 0 order by SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        TaxAmt = 0
        taxableamt = 0
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select * FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND SALES_TAX = " & RSTTRXFILE!SALES_TAX & " AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
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
                        txtNetrate.text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
                    Case 2
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
                        'Kannattu
                        txtNetrate.text = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
                    Case 3
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
                        'Kannattu
                        txtNetrate.text = IIf(IsNull(GRDPOPUP.Columns(21)), "", GRDPOPUP.Columns(21))
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(21)), "", GRDPOPUP.Columns(21))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(21)), "", GRDPOPUP.Columns(21))
                        If Val(txtretail.text) = 0 Then
                            txtNetrate.text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                            txtretail.text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                            TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
                        End If
                    Case 4
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
                        'Kannattu
                        txtNetrate.text = IIf(IsNull(GRDPOPUP.Columns(32)), "", GRDPOPUP.Columns(32))
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(32)), "", GRDPOPUP.Columns(32))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(32)), "", GRDPOPUP.Columns(32))
                    Case 5
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
                        'Kannattu
                        txtNetrate.text = IIf(IsNull(GRDPOPUP.Columns(33)), "", GRDPOPUP.Columns(33))
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(33)), "", GRDPOPUP.Columns(33))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(33)), "", GRDPOPUP.Columns(33))
                    Case 6
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(19)), "", GRDPOPUP.Columns(19))
                        'Kannattu
                        txtNetrate.text = IIf(IsNull(GRDPOPUP.Columns(34)), "", GRDPOPUP.Columns(34))
                        txtretail.text = IIf(IsNull(GRDPOPUP.Columns(34)), "", GRDPOPUP.Columns(34))
                        TXTRETAILNOTAX.text = IIf(IsNull(GRDPOPUP.Columns(34)), "", GRDPOPUP.Columns(34))
                    Case Else
                        'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(20)), "", GRDPOPUP.Columns(20))
                        'Kannattu
                        txtNetrate.text = IIf(IsNull(GRDPOPUP.Columns(9)), "", GRDPOPUP.Columns(9))
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
                lblprice5.Caption = IIf(IsNull(GRDPOPUP.Columns(32)), "", GRDPOPUP.Columns(32))
                lblprice6.Caption = IIf(IsNull(GRDPOPUP.Columns(33)), "", GRDPOPUP.Columns(33))
                lblprice7.Caption = IIf(IsNull(GRDPOPUP.Columns(34)), "", GRDPOPUP.Columns(34))
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
                MsgBox "Enter Bill Type", vbOKOnly, "QUOTATION"
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
'        MsgBox "Select Bill Type from the List", vbOKOnly, "QUOTATION"
'        cmbtype.SetFocus
'        Exit Sub
'    End If
'    If cmbtype.ListIndex = 0 And Val(TXTTYPE.Text) <> 1 Then
'        MsgBox "Bill type doesnot match", vbOKOnly, "QUOTATION"
'        TXTTYPE.SetFocus
'        Exit Sub
'    End If
'    If cmbtype.ListIndex = 1 And Val(TXTTYPE.Text) <> 2 Then
'        MsgBox "Bill type doesnot match", vbOKOnly, "QUOTATION"
'        TXTTYPE.SetFocus
'        Exit Sub
'    End If
'    If cmbtype.ListIndex = 2 And Val(TXTTYPE.Text) <> 3 Then
'        MsgBox "Bill type doesnot match", vbOKOnly, "QUOTATION"
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
        grdsales.TextMatrix(i, 47) = grdsales.TextMatrix(i + 1, 47)
    Next i
    grdsales.rows = grdsales.rows - 1
    
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
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                If MDIMAIN.lblgst.Caption = "R" Then
                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                End If
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
    If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
        TxtName1.SetFocus
    Else
        TXTITEMCODE.SetFocus
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
                    db.BeginTrans
                    
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.000")
                    
                    TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(grdsales.Row, 6)) - (Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100)
                    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(grdsales.TextMatrix(grdsales.Row, 9)) + Val(grdsales.TextMatrix(grdsales.Row, 47))) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(grdsales.Row, 40)) / 100) + Val(grdsales.TextMatrix(grdsales.Row, 41))
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)), 4), ".0000")
                    Else
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
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
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                If MDIMAIN.lblgst.Caption = "R" Then
                                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                End If
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
                    db.Execute "Update QTNSUB set QTY = " & Val(grdsales.TextMatrix(grdsales.Row, 3)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update QTNMAST set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.CommitTrans
                    
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                    
                Case 5  'MRP
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.000")
                    db.BeginTrans
                    db.Execute "Update QTNSUB set MRP = " & Val(grdsales.TextMatrix(grdsales.Row, 5)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    db.CommitTrans
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
            
                Case 6  'RATE
                    TXTDISC.Tag = 0
                    grdsales.TextMatrix(grdsales.Row, 7) = Format(Round(Val(TXTsample.text) + Val(TXTsample.text) * (Val(grdsales.TextMatrix(grdsales.Row, 9))) / 100, 4), "0.0000")
                    grdsales.TextMatrix(grdsales.Row, 21) = Format(Round(Val(TXTsample.text) + Val(TXTsample.text) * (Val(grdsales.TextMatrix(grdsales.Row, 9))) / 100, 4), "0.0000")
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.000")
                    
                    TXTDISC.Tag = 0
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "PARDHA" Or UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "CLOTHES" Then
                        If Val(grdsales.TextMatrix(grdsales.Row, 6)) < 1000 Then
                           grdsales.TextMatrix(grdsales.Row, 9) = "5"
                        Else
                            grdsales.TextMatrix(grdsales.Row, 9) = "12"
                        End If
                    End If
                    
                    If Trim(TXTTIN.text) = "" And Trim(lblIGST.Caption) <> "Y" And MDIMAIN.lblkfc.Caption = "Y" And IsDate(MDIMAIN.DTKFCSTART.Value) And IsDate(MDIMAIN.DTKFCEND.Value) Then
                        If DateValue(TXTINVDATE.text) >= DateValue(MDIMAIN.DTKFCSTART.Value) And DateValue(TXTINVDATE.text) <= DateValue(MDIMAIN.DTKFCEND.Value) Then
                            If Val(grdsales.TextMatrix(grdsales.Row, 9)) = 12 Or Val(grdsales.TextMatrix(grdsales.Row, 9)) = 18 Or Val(grdsales.TextMatrix(grdsales.Row, 9)) = 28 Then
                                grdsales.TextMatrix(grdsales.Row, 47) = 1
                            Else
                                grdsales.TextMatrix(grdsales.Row, 47) = 0
                            End If
                        End If
                    Else
                        grdsales.TextMatrix(grdsales.Row, 47) = 0
                    End If
                    
                    TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(grdsales.Row, 6)) - (Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100)
                    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(grdsales.TextMatrix(grdsales.Row, 9)) + Val(grdsales.TextMatrix(grdsales.Row, 47))) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(grdsales.Row, 40)) / 100) + Val(grdsales.TextMatrix(grdsales.Row, 41))
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)), 4), ".0000")
                    Else
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
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
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                If MDIMAIN.lblgst.Caption = "R" Then
                                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                End If
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
                    
                    db.BeginTrans
                    TXTDISC.Tag = (Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(grdsales.TextMatrix(grdsales.Row, 6))) * Val(grdsales.TextMatrix(grdsales.Row, 3))
                    db.Execute "Update QTNSUB set KFC_TAX = " & Val(grdsales.TextMatrix(grdsales.Row, 47)) & ", SALES_PRICE = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", PTR = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update QTNMAST set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.CommitTrans
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                    
                Case 7  'NET RATE
                    
                    'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                    grdsales.TextMatrix(grdsales.Row, 6) = Format(Round(Val(TXTsample.text) * 100 / ((Val(grdsales.TextMatrix(grdsales.Row, 9))) + 100), 4), "0.0000")
                    grdsales.TextMatrix(grdsales.Row, 22) = Format(Round(Val(TXTsample.text) * 100 / ((Val(grdsales.TextMatrix(grdsales.Row, 9))) + 100), 4), "0.0000")
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.000")
                    
                    TXTDISC.Tag = 0
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "PARDHA" Or UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "CLOTHES" Then
                        If Val(grdsales.TextMatrix(grdsales.Row, 6)) < 1000 Then
                           grdsales.TextMatrix(grdsales.Row, 9) = "5"
                        Else
                            grdsales.TextMatrix(grdsales.Row, 9) = "12"
                        End If
                    End If
                    
                    If Trim(TXTTIN.text) = "" And Trim(lblIGST.Caption) <> "Y" And MDIMAIN.lblkfc.Caption = "Y" And IsDate(MDIMAIN.DTKFCSTART.Value) And IsDate(MDIMAIN.DTKFCEND.Value) Then
                        If DateValue(TXTINVDATE.text) >= DateValue(MDIMAIN.DTKFCSTART.Value) And DateValue(TXTINVDATE.text) <= DateValue(MDIMAIN.DTKFCEND.Value) Then
                            If Val(grdsales.TextMatrix(grdsales.Row, 9)) = 12 Or Val(grdsales.TextMatrix(grdsales.Row, 9)) = 18 Or Val(grdsales.TextMatrix(grdsales.Row, 9)) = 28 Then
                                grdsales.TextMatrix(grdsales.Row, 47) = 1
                            Else
                                grdsales.TextMatrix(grdsales.Row, 47) = 0
                            End If
                        End If
                    Else
                        grdsales.TextMatrix(grdsales.Row, 47) = 0
                    End If
        
                    TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(grdsales.Row, 6)) - (Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100)
                    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(grdsales.TextMatrix(grdsales.Row, 9)) + Val(grdsales.TextMatrix(grdsales.Row, 47))) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(grdsales.Row, 40)) / 100) + Val(grdsales.TextMatrix(grdsales.Row, 41))
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)), 4), ".0000")
                    Else
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
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
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                If MDIMAIN.lblgst.Caption = "R" Then
                                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                End If
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
                    db.BeginTrans
                    db.Execute "Update QTNSUB set KFC_TAX = " & Val(grdsales.TextMatrix(grdsales.Row, 47)) & ", SALES_PRICE = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", PTR = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update QTNMAST set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.CommitTrans
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                
                Case 8  'Disc
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.00")
                    TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(grdsales.Row, 6)) - (Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100)
                    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(grdsales.TextMatrix(grdsales.Row, 9)) + Val(grdsales.TextMatrix(grdsales.Row, 47))) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(grdsales.Row, 40)) / 100) + Val(grdsales.TextMatrix(grdsales.Row, 41))
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)), 4), ".0000")
                    Else
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
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
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                If MDIMAIN.lblgst.Caption = "R" Then
                                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                End If
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
                    
                    db.BeginTrans
                    db.Execute "Update QTNSUB set LINE_DISC = " & Val(grdsales.TextMatrix(grdsales.Row, 8)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update QTNMAST set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.CommitTrans
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                    
                Case 9  'TAX
                    If Val(TXTsample.text) = 12 Or Val(TXTsample.text) = 18 Or Val(TXTsample.text) = 28 Then
                        grdsales.TextMatrix(grdsales.Row, 47) = "1"
                    Else
                        grdsales.TextMatrix(grdsales.Row, 47) = "0"
                    End If
                    TXTDISC.Tag = 0
                    If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                        grdsales.TextMatrix(grdsales.Row, 7) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) + Val(grdsales.TextMatrix(grdsales.Row, 6)) * (Val(TXTsample.text)) / 100, 3), "0.000")
                        grdsales.TextMatrix(grdsales.Row, 21) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) + Val(grdsales.TextMatrix(grdsales.Row, 6)) * (Val(TXTsample.text)) / 100, 3), "0.000")
                    End If
                    grdsales.TextMatrix(grdsales.Row, 6) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) * 100 / ((Val(TXTsample.text)) + 100), 3), "0.000")
                    grdsales.TextMatrix(grdsales.Row, 22) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) * 100 / ((Val(TXTsample.text)) + 100), 3), "0.000")
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.00")
                    TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(grdsales.Row, 6)) - (Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100)
                    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(grdsales.TextMatrix(grdsales.Row, 9)) + Val(grdsales.TextMatrix(grdsales.Row, 47))) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(grdsales.Row, 40)) / 100) + Val(grdsales.TextMatrix(grdsales.Row, 41))
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)), 4), ".0000")
                    Else
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
                    End If
                    grdsales.TextMatrix(grdsales.Row, 17) = "V"
                    
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
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                If MDIMAIN.lblgst.Caption = "R" Then
                                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                End If
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
                    db.BeginTrans
                    db.Execute "Update QTNSUB set KFC_TAX = " & Val(grdsales.TextMatrix(grdsales.Row, 47)) & ", SALES_TAX = " & Val(grdsales.TextMatrix(grdsales.Row, 9)) & ", SALES_PRICE = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", PTR = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update QTNMAST set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    db.CommitTrans
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                Case 24  'Commission
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.text), "0.00")
                    
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
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                If MDIMAIN.lblgst.Caption = "R" Then
                                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                End If
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
                    
                    db.BeginTrans
                    db.Execute "Update QTNSUB set COM_AMT = " & Val(grdsales.TextMatrix(grdsales.Row, 24)) & "  WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update QTNMAST set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    'db.Execute "Update QTNMAST set PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.Text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.Text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.Text) & ""
                    db.CommitTrans
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
            End Select
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
    db.RollbackTrans
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
    TxtItemName.Visible = False
    DataList3.Visible = False
    
    grdsales.SetFocus
    If grdsales.rows <= 1 Then Exit Sub
    Call COSTCALCULATE
    If M_EDIT = True Then Exit Sub
    If TXTQTY.Enabled = True Then Exit Sub
    lblretail.Caption = ""
    lblwsale.Caption = ""
    lblvan.Caption = ""
    LBLMRP.Caption = ""
    lblprice5.Caption = ""
    lblprice6.Caption = ""
    lblprice7.Caption = ""
    lblcase.Caption = ""
    lblLWPrice.Caption = ""
    lblcrtnpack.Caption = 1
    Dim check_flag As Boolean
    
    check_flag = False
    Dim RSTITEMMAST As ADODB.Recordset
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "' AND PRICE_CHANGE ='Y'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        lblretail.Caption = IIf(IsNull(RSTITEMMAST!P_RETAIL) Or RSTITEMMAST!P_RETAIL = 0, "", RSTITEMMAST!P_RETAIL)
        lblwsale.Caption = IIf(IsNull(RSTITEMMAST!P_WS) Or RSTITEMMAST!P_WS = 0, "", RSTITEMMAST!P_WS)
        lblvan.Caption = IIf(IsNull(RSTITEMMAST!P_VAN) Or RSTITEMMAST!P_VAN = 0, "", RSTITEMMAST!P_VAN)
        LBLMRP.Caption = IIf(IsNull(RSTITEMMAST!MRP) Or RSTITEMMAST!MRP = 0, "", RSTITEMMAST!MRP)
        lblprice5.Caption = IIf(IsNull(RSTITEMMAST!PRICE5) Or RSTITEMMAST!PRICE5 = 0, "", RSTITEMMAST!PRICE5)
        lblprice6.Caption = IIf(IsNull(RSTITEMMAST!PRICE6) Or RSTITEMMAST!PRICE6 = 0, "", RSTITEMMAST!PRICE6)
        lblprice7.Caption = IIf(IsNull(RSTITEMMAST!PRICE7) Or RSTITEMMAST!PRICE7 = 0, "", RSTITEMMAST!PRICE7)
        lblcase.Caption = IIf(IsNull(RSTITEMMAST!P_CRTN) Or RSTITEMMAST!P_CRTN = 0, "", RSTITEMMAST!P_CRTN)
        lblLWPrice.Caption = IIf(IsNull(RSTITEMMAST!P_LWS) Or RSTITEMMAST!P_LWS = 0, "", RSTITEMMAST!P_LWS)
        lblcrtnpack.Caption = IIf(IsNull(RSTITEMMAST!CRTN_PACK) Or RSTITEMMAST!CRTN_PACK = 0, 1, RSTITEMMAST!CRTN_PACK)
        check_flag = True
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        GoTo SKIP
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    If MDIMAIN.lblPriceSplit.Caption = "Y" Or PC_FLAG = "Y" Then
        RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            lblretail.Caption = IIf(IsNull(RSTITEMMAST!P_RETAIL) Or RSTITEMMAST!P_RETAIL = 0, "", RSTITEMMAST!P_RETAIL)
            lblwsale.Caption = IIf(IsNull(RSTITEMMAST!P_WS) Or RSTITEMMAST!P_WS = 0, "", RSTITEMMAST!P_WS)
            lblvan.Caption = IIf(IsNull(RSTITEMMAST!P_VAN) Or RSTITEMMAST!P_VAN = 0, "", RSTITEMMAST!P_VAN)
            LBLMRP.Caption = IIf(IsNull(RSTITEMMAST!MRP) Or RSTITEMMAST!MRP = 0, "", RSTITEMMAST!MRP)
            lblprice5.Caption = IIf(IsNull(RSTITEMMAST!PRICE5) Or RSTITEMMAST!PRICE5 = 0, "", RSTITEMMAST!PRICE5)
            lblprice6.Caption = IIf(IsNull(RSTITEMMAST!PRICE6) Or RSTITEMMAST!PRICE6 = 0, "", RSTITEMMAST!PRICE6)
            lblprice7.Caption = IIf(IsNull(RSTITEMMAST!PRICE7) Or RSTITEMMAST!PRICE7 = 0, "", RSTITEMMAST!PRICE7)
            lblcase.Caption = IIf(IsNull(RSTITEMMAST!P_CRTN) Or RSTITEMMAST!P_CRTN = 0, "", RSTITEMMAST!P_CRTN)
            lblLWPrice.Caption = IIf(IsNull(RSTITEMMAST!P_LWS) Or RSTITEMMAST!P_LWS = 0, "", RSTITEMMAST!P_LWS)
            lblcrtnpack.Caption = IIf(IsNull(RSTITEMMAST!CRTN_PACK) Or RSTITEMMAST!CRTN_PACK = 0, 1, RSTITEMMAST!CRTN_PACK)
            check_flag = True
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    Else
        RSTITEMMAST.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(grdsales.Row, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(grdsales.Row, 43)) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            lblretail.Caption = IIf(IsNull(RSTITEMMAST!P_RETAIL) Or RSTITEMMAST!P_RETAIL = 0, "", RSTITEMMAST!P_RETAIL)
            lblwsale.Caption = IIf(IsNull(RSTITEMMAST!P_WS) Or RSTITEMMAST!P_WS = 0, "", RSTITEMMAST!P_WS)
            lblvan.Caption = IIf(IsNull(RSTITEMMAST!P_VAN) Or RSTITEMMAST!P_VAN = 0, "", RSTITEMMAST!P_VAN)
            LBLMRP.Caption = IIf(IsNull(RSTITEMMAST!MRP) Or RSTITEMMAST!MRP = 0, "", RSTITEMMAST!MRP)
            lblprice5.Caption = IIf(IsNull(RSTITEMMAST!PRICE5) Or RSTITEMMAST!PRICE5 = 0, "", RSTITEMMAST!PRICE5)
            lblprice6.Caption = IIf(IsNull(RSTITEMMAST!PRICE6) Or RSTITEMMAST!PRICE6 = 0, "", RSTITEMMAST!PRICE6)
            lblprice7.Caption = IIf(IsNull(RSTITEMMAST!PRICE7) Or RSTITEMMAST!PRICE7 = 0, "", RSTITEMMAST!PRICE7)
            lblcase.Caption = IIf(IsNull(RSTITEMMAST!P_CRTN) Or RSTITEMMAST!P_CRTN = 0, "", RSTITEMMAST!P_CRTN)
            lblLWPrice.Caption = IIf(IsNull(RSTITEMMAST!P_LWS) Or RSTITEMMAST!P_LWS = 0, "", RSTITEMMAST!P_LWS)
            lblcrtnpack.Caption = IIf(IsNull(RSTITEMMAST!CRTN_PACK) Or RSTITEMMAST!CRTN_PACK = 0, 1, RSTITEMMAST!CRTN_PACK)
            check_flag = True
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    End If
    
    If check_flag = False Then
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            lblretail.Caption = IIf(IsNull(RSTITEMMAST!P_RETAIL) Or RSTITEMMAST!P_RETAIL = 0, "", RSTITEMMAST!P_RETAIL)
            lblwsale.Caption = IIf(IsNull(RSTITEMMAST!P_WS) Or RSTITEMMAST!P_WS = 0, "", RSTITEMMAST!P_WS)
            lblvan.Caption = IIf(IsNull(RSTITEMMAST!P_VAN) Or RSTITEMMAST!P_VAN = 0, "", RSTITEMMAST!P_VAN)
            LBLMRP.Caption = IIf(IsNull(RSTITEMMAST!MRP) Or RSTITEMMAST!MRP = 0, "", RSTITEMMAST!MRP)
            lblprice5.Caption = IIf(IsNull(RSTITEMMAST!PRICE5) Or RSTITEMMAST!PRICE5 = 0, "", RSTITEMMAST!PRICE5)
            lblprice6.Caption = IIf(IsNull(RSTITEMMAST!PRICE6) Or RSTITEMMAST!PRICE6 = 0, "", RSTITEMMAST!PRICE6)
            lblprice7.Caption = IIf(IsNull(RSTITEMMAST!PRICE7) Or RSTITEMMAST!PRICE7 = 0, "", RSTITEMMAST!PRICE7)
            lblcase.Caption = IIf(IsNull(RSTITEMMAST!P_CRTN) Or RSTITEMMAST!P_CRTN = 0, "", RSTITEMMAST!P_CRTN)
            lblLWPrice.Caption = IIf(IsNull(RSTITEMMAST!P_LWS) Or RSTITEMMAST!P_LWS = 0, "", RSTITEMMAST!P_LWS)
            lblcrtnpack.Caption = IIf(IsNull(RSTITEMMAST!CRTN_PACK) Or RSTITEMMAST!CRTN_PACK = 0, 1, RSTITEMMAST!CRTN_PACK)
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    End If
SKIP:
    
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
    MsgBox err.Description, , "EzBiz"
End Sub
Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If grdsales.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then Exit Sub
            Select Case grdsales.Col
                Case 2
                    TxtItemName.Visible = True
                    DataList3.Visible = True
                    DataList3.Top = grdsales.CellTop + 510
                    DataList3.Left = grdsales.CellLeft '+ 50
                    TxtItemName.Top = grdsales.CellTop + 100
                    TxtItemName.Left = grdsales.CellLeft '+ 50
                    TxtItemName.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TxtItemName.Height = grdsales.CellHeight
                    TxtItemName.SetFocus
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
                    TXTsample.Visible = True
                    TXTsample.Top = grdsales.CellTop + 100
                    TXTsample.Left = grdsales.CellLeft + 0
                    TXTsample.Width = grdsales.CellWidth
                    TXTsample.Height = grdsales.CellHeight
                    TXTsample.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TXTsample.SetFocus
                Case 8
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
    TxtItemName.Visible = False
    DataList3.Visible = False
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
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            Else
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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
    On Error GoTo ERRHAND
    
    If txtBillNo.Tag = "Y" Then
        MsgBox "Any changes made will not be saved", vbOKOnly, "QUOTATION"
        GoTo SKIP
    End If
    
    'If OLD_BILL = False Then Call checklastbill
    Set RSTTRXFILE = New ADODB.Recordset
    If OLD_BILL = False And Val(txtBillNo.text) <> 1 Then
        RSTTRXFILE.Open "Select * FROM QTNMAST WHERE VCH_NO= (SELECT MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'QT')", db, adOpenStatic, adLockOptimistic, adCmdText
        txtBillNo.text = RSTTRXFILE!VCH_NO + 1
        LBLBILLNO.Caption = Val(txtBillNo.text)
        db.BeginTrans
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "QT"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = txtBillNo.text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
    Else
        RSTTRXFILE.Open "Select * FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "QT"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = txtBillNo.text
            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
            RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
        End If
    End If
    'RSTTRXFILE!CUST_IGST = lblIGST.Caption
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
    'RSTTRXFILE!BILL_FLAG = "Y"
'    If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
'        RSTTRXFILE!TERMS = Trim(Terms1.Text)
'    Else
'        RSTTRXFILE!TERMS = ""
'    End If
    'RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
    'RSTTRXFILE!BR_NAME = CMBBRNCH.Text
    'RSTTRXFILE!cr_days = Val(txtcrdays.Text)
    RSTTRXFILE.Update
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    OLD_BILL = True
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    If grdsales.rows = 1 Then
        If OLD_BILL = True Then
            Dim LASTBILL As Long
            LASTBILL = 1
            Set rstBILL = New ADODB.Recordset
            rstBILL.Open "Select MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'QT'", db, adOpenForwardOnly
            If Not (rstBILL.EOF And rstBILL.BOF) Then
                LASTBILL = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0))
            End If
            rstBILL.Close
            Set rstBILL = Nothing
            
            'If LASTBILL = 1 Then GoTo SKIP
            If Val(txtBillNo.text) >= LASTBILL Then
                db.Execute "delete FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
            End If
        End If
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT *  FROM QTNMAST WHERE BILL_NO = " & Val(txtBillNo.Text) & " AND BILLTYPE = 'QT' ", db, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'            RSTTRXFILE!BILL_NO = Null
'            RSTTRXFILE!billtype = Null
'            RSTTRXFILE.Update
'        End If
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
        GoTo SKIP
    End If
    
    
    i = 0
    
    
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
    RSTTRXFILE.Open "Select * FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!TRX_TYPE = "QT"
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
    'RSTTRXFILE!CUST_IGST = lblIGST.Caption
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
    RSTTRXFILE!POST_FLAG = "Y"
    RSTTRXFILE!CFORM_NO = Time
    RSTTRXFILE!REMARKS = Left(DataList2.text, 50)
    RSTTRXFILE!DISC_PERS = Val(TXTTOTALDISC.text)
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!VEHICLE = Trim(TxtVehicle.text)
    'RSTTRXFILE!D_ORDER = Trim(TxtOrder.Text)
    RSTTRXFILE!PHONE = Trim(TxtPhone.text)
    RSTTRXFILE!TIN = Trim(TXTTIN.text)
    RSTTRXFILE!FRIEGHT = Val(TxtFrieght.text)
    RSTTRXFILE!HANDLE = Val(Txthandle.text)
    RSTTRXFILE!Area = Trim(TXTAREA.text)
    RSTTRXFILE!MODIFY_DATE = Date
    'RSTTRXFILE!cr_days = Val(txtcrdays.Text)
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
        Case 4
            RSTTRXFILE!BILL_TYPE = "5"
        Case 5
            RSTTRXFILE!BILL_TYPE = "6"
        Case 6
            RSTTRXFILE!BILL_TYPE = "7"
        Case Else
            RSTTRXFILE!BILL_TYPE = "R"
    End Select
    If Val(TxtCN.text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.text)
    RSTTRXFILE!RCPT_AMOUNT = 0
    RSTTRXFILE!RCPT_REFNO = ""
    RSTTRXFILE!BANK_FLAG = "N"
    'RSTTRXFILE!CHQ_NO = Null
    'RSTTRXFILE!BANK_CODE = Null
    'RSTTRXFILE!BANK_NAME = Null
    'RSTTRXFILE!CHQ_DATE = Null
    RSTTRXFILE!CHQ_STATUS = "N"
    'RSTTRXFILE!BILL_FLAG = "Y"
'    If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
'        RSTTRXFILE!TERMS = Trim(Terms1.Text)
'    Else
'        RSTTRXFILE!TERMS = ""
'    End If
'    RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
'    RSTTRXFILE!BR_NAME = CMBBRNCH.Text
    RSTTRXFILE.Update
    db.CommitTrans
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
'    rstBILL.Open "Select MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'QT'", db, adOpenStatic, adLockReadOnly
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
    rstBILL.Open "Select MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'QT'", db, adOpenStatic, adLockReadOnly
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
    LBLFOT.Caption = ""
    LBLRETAMT.Caption = ""
    LBLPROFIT.Caption = ""
    LBLDATE.Caption = Date
    LBLGROSSAMT.Caption = ""
    LBLTOTAL.Caption = ""
    lblcomamt.text = ""
    TXTTOTALDISC.text = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
    TXTAMOUNT.text = ""
    LBLDISCAMT.Caption = ""
    lblbalance.Caption = ""
    Txtrcvd.text = ""
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
        
    flagchange2.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Tag = ""
    TxtCode.text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblcredit.Caption = "0"
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
    M_ADD = False
    TXTDEALER.text = ""
    txtphsearch.text = ""
    Terms1.text = Terms1.Tag
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    cmbtype.ListIndex = 0
    TXTTYPE.text = 1
    'TXTDEALER.SetFocus
    
    
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
    TXTITEMCODE.Enabled = True
    If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
        If TxtName1.Enabled = True Then TxtName1.SetFocus
    Else
        If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
    End If
    
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Function

Private Function Print_A4()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim DN_CN_FLag As Boolean
    Dim oRs As ADODB.Recordset
    
    Dim i As Long
    Dim CN As Integer
    Dim DN As Integer
    Dim b As Integer
    Dim Num, Figre As Currency
    
    On Error GoTo ERRHAND
    
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
    
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompAddress4, CompAddress5, CompTin, CompCST, BIL_PRE, BILL_SUF, DL, ML, DL1, DL2, INV_TERMS, INV_MSG, BANK_DET, PAN_NO  As String
    Dim QtnTerms, QtnTerms1, QtnTerms2, QtnTerms3, QtnTerms4 As String
    Dim RSTCOMPANY As ADODB.Recordset
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
        If Trim(TxtVehicle.text) = "" Then TxtVehicle.text = IIf(IsNull(RSTCOMPANY!VEHICLE), "", RSTCOMPANY!VEHICLE)
        INV_TERMS = IIf(IsNull(RSTCOMPANY!INV_TERMS) Or RSTCOMPANY!INV_TERMS = "", "", RSTCOMPANY!INV_TERMS)
        INV_MSG = IIf(IsNull(RSTCOMPANY!INV_MSGS) Or RSTCOMPANY!INV_MSGS = "", "", RSTCOMPANY!INV_MSGS)
        BANK_DET = IIf(IsNull(RSTCOMPANY!bank_details) Or RSTCOMPANY!bank_details = "", "", RSTCOMPANY!bank_details)
        PAN_NO = IIf(IsNull(RSTCOMPANY!PAN_NO) Or RSTCOMPANY!PAN_NO = "", "", RSTCOMPANY!PAN_NO)
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
    

    Set RSTTRXFILE = New ADODB.Recordset
    If OLD_BILL = False And Val(txtBillNo.text) <> 1 Then
        RSTTRXFILE.Open "Select * FROM QTNMAST WHERE VCH_NO= (SELECT MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'QT')", db, adOpenStatic, adLockOptimistic, adCmdText
        txtBillNo.text = RSTTRXFILE!VCH_NO + 1
        LBLBILLNO.Caption = Val(txtBillNo.text)
        db.BeginTrans
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "QT"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = txtBillNo.text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
        RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
    Else
        RSTTRXFILE.Open "Select * FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "QT"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = txtBillNo.text
            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
            RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
        End If
    End If
    'RSTTRXFILE!CUST_IGST = lblIGST.Caption
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
    'RSTTRXFILE!BILL_FLAG = "Y"
'    If chkTerms.value = 1 And Trim(Terms1.Text) <> "" Then
'        RSTTRXFILE!TERMS = Trim(Terms1.Text)
'    Else
'        RSTTRXFILE!TERMS = ""
'    End If
'    RSTTRXFILE!BR_CODE = CMBBRNCH.BoundText
'    RSTTRXFILE!BR_NAME = CMBBRNCH.Text
'    RSTTRXFILE!cr_days = Val(txtcrdays.Text)
    RSTTRXFILE.Update
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    OLD_BILL = True
        
    db.Execute "delete from TEMPTRXFILE WHERE VCH_NO = " & Val(txtBillNo.text) & " "
    'Dim RSTUNBILL As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        
'        If thermalprn = False And grdsales.TextMatrix(i, 46) = "Y" Then GoTo SKIP_UNBILL
'            Set RSTUNBILL = New ADODB.Recordset
'            RSTUNBILL.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "' AND UN_BILL = 'Y'", db, adOpenStatic, adLockReadOnly, adCmdText
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
        RSTTRXFILE!TRX_TYPE = "QT"
        'RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 13)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 5))
        
        If MDIMAIN.lblgst.Caption = "R" Then
            If Trim(TXTTIN.text) <> "" Then
                RSTTRXFILE!SALES_PRICE = Round(Val(grdsales.TextMatrix(i, 7)) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 47)) / 100), 3)
                RSTTRXFILE!P_RETAIL = Round(Val(grdsales.TextMatrix(i, 7)) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 47)) / 100), 3)
                RSTTRXFILE!PTR = Round(RSTTRXFILE!P_RETAIL * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 3)
                RSTTRXFILE!P_RETAILWOTAX = Round(RSTTRXFILE!P_RETAIL * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 3)
                RSTTRXFILE!kfc_tax = 0
            Else
                RSTTRXFILE!kfc_tax = Val(grdsales.TextMatrix(i, 47))
                RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
                RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
                RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 7)) 'Val(grdsales.TextMatrix(i, 21))
                RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 6)) 'Val(grdsales.TextMatrix(i, 22))
            End If
            RSTTRXFILE!CESS_PER = Val(grdsales.TextMatrix(i, 40))
            RSTTRXFILE!cess_amt = Val(grdsales.TextMatrix(i, 41))
            RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        Else
            RSTTRXFILE!kfc_tax = 0
            RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 7))
            RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
            RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 7)) 'Val(grdsales.TextMatrix(i, 21))
            RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 7)) 'Val(grdsales.TextMatrix(i, 22))
            RSTTRXFILE!SALES_TAX = 0
            RSTTRXFILE!CESS_PER = 0
            RSTTRXFILE!cess_amt = 0
        End If
        'RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!VCH_DESC = "" '"Issued to     " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(i, 17))
        'RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!C_USER_ID = IIf(IsNull(RSTITEMMAST!REMARKS), "", Left(RSTITEMMAST!REMARKS, 8))
            RSTTRXFILE!MFGR = IIf(IsNull(RSTITEMMAST!ITEM_MAL), "", RSTITEMMAST!ITEM_MAL)
            RSTTRXFILE!M_USER_ID = IIf(IsNull(RSTITEMMAST!FULL_PACK), RSTTRXFILE!PACK_TYPE, RSTITEMMAST!FULL_PACK)
        Else
            RSTTRXFILE!Category = ""
            RSTTRXFILE!MFGR = ""
            RSTTRXFILE!M_USER_ID = ""
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
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        
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
    rstTRXMAST.Open "SELECT * from RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(TxtCN.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
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
    RSTTRXFILE.Open "SELECT DISTINCT SALES_TAX From TEMPTRXFILE WHERE TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " order by SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        TaxAmt = 0
        taxableamt = 0
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select * From TEMPTRXFILE WHERE SALES_TAX = " & RSTTRXFILE!SALES_TAX & " AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
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
        
        ReportNameVar = Rptpath & "RPTQTN"
'        If Small_Print = True Then
'
'        Else
'            ReportNameVar = Rptpath & "RPTQUTN"
'        End If
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
            Report.Database.SetDataSource oRs, 3, 1
            Set oRs = Nothing
        End If
        
        Report.RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & ")"
        For i = 1 To Report.Database.Tables.COUNT
            Report.Database.Tables.Item(i).SetLogOnInfo strConnection
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
            If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
            If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
            If CRXFormulaField.Name = "{@Terms}" Then CRXFormulaField.text = "'" & QtnTerms & "'"
            If CRXFormulaField.Name = "{@Terms1}" Then CRXFormulaField.text = "'" & QtnTerms1 & "'"
            If CRXFormulaField.Name = "{@Terms2}" Then CRXFormulaField.text = "'" & QtnTerms2 & "'"
            If CRXFormulaField.Name = "{@Terms3}" Then CRXFormulaField.text = "'" & QtnTerms3 & "'"
            If CRXFormulaField.Name = "{@Terms4}" Then CRXFormulaField.text = "'" & QtnTerms4 & "'"
            If CRXFormulaField.Name = "{@Rcvd}" Then CRXFormulaField.text = " " & Val(Txtrcvd.text) & " "
            If CRXFormulaField.Name = "{@BalAmt}" Then CRXFormulaField.text = " " & Abs(Val(lblbalance.Caption)) & " "
            If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.text = "'" & BANK_DET & "'"
            If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.text = "'" & PAN_NO & "'"
            If CRXFormulaField.Name = "{@DL1}" Then CRXFormulaField.text = "'" & DL1 & "'"
            If CRXFormulaField.Name = "{@DL2}" Then CRXFormulaField.text = "'" & DL2 & "'"
        
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
            If Trim(TXTTIN.text) <> "" Then
                If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.text = "'GSTIN: ' & '" & Trim(TXTTIN.text) & "'"
            Else
                'If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'UID: ' & '" & Trim(TxtUID.Text) & "'"
            End If
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
        
        'Thermal
        If MDIMAIN.LBLTHPREVIEW.Caption = "Y" Then
            'Preview
            frmreport.Caption = "BILL"
            Call GENERATEREPORT
            Screen.MousePointer = vbNormal
        Else
            '    '''No Preview
            Report.PrintOut (False)
            'Report.PrinterName = thermalprinter
            Set CRXFormulaFields = Nothing
            Set CRXFormulaField = Nothing
            Set crxApplication = Nothing
            Set Report = Nothing
            Call cmdRefresh_Click
        End If
        Screen.MousePointer = vbNormal
        
        Exit Function
    End If
    
    If Small_Print = True Then
        'ReportNameVar = Rptpath & "rptbillretail"
        If MDIMAIN.lblgst.Caption = "R" Then
            If OptPI.Value = True Then
                ReportNameVar = Rptpath & "RPTQUTNPA5"
            Else
                ReportNameVar = Rptpath & "RPTQUTNA5"
            End If
        Else
            ReportNameVar = Rptpath & "RPTQUTNCA5"
        End If
    Else
        If MDIMAIN.lblgst.Caption = "R" Then
            If OptPI.Value = True Then
                ReportNameVar = Rptpath & "RPTQUTNP"
            Else
                ReportNameVar = Rptpath & "RPTQUTN"
            End If
        Else
            ReportNameVar = Rptpath & "RPTQUTNC"
        End If
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
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
    Report.RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & ")"
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
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
        If CRXFormulaField.Name = "{@ML}" Then CRXFormulaField.text = "'" & ML & "'"
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
        If Trim(TXTTIN.text) <> "" Then
            If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.text = "'GSTIN: ' & '" & Trim(TXTTIN.text) & "'"
        Else
            'If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'UID: ' & '" & Trim(TxtUID.Text) & "'"
        End If
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
    
    
    
'    If MDIMAIN.StatusBar.Panels(13).Text = "Y" Then
        'Preview
        frmreport.Caption = "BILL"
        Call GENERATEREPORT
        Screen.MousePointer = vbNormal
'    Else
'        '    '''No Preview
'        Report.PrintOut (False)
'        Set CRXFormulaFields = Nothing
'        Set CRXFormulaField = Nothing
'        Set crxApplication = Nothing
'        Set Report = Nothing
'        Call cmdRefresh_Click
'        Exit Function
'    End If
    
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
    OLD_BILL = True
    ''rptPRINT.Action = 1
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Function

'Private Sub CmdTax_Click()
'    If grdsales.Rows <= 1 Then Exit Sub
'    If Trim(TxtTaxPrint.Text) = "" Then Exit Sub
'    Tax_Print = True
'    Call Generateprint
'    TxtTaxPrint.Text = ""
'End Sub

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
            RSTTRXFILE.Open "SELECT CLOSE_QTY  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!CLOSE_QTY)) Then RSTTRXFILE!CLOSE_QTY = 0
                i = RSTTRXFILE!CLOSE_QTY / Val(LblPack.text)
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
'            If M_EDIT = False And (Val(TXTQTY.Text) + Val(TXTFREE.Text) > i) Then
'                MsgBox "AVAILABLE STOCK IS  " & i & " ", , "QUOTATION"
'                TXTQTY.SelStart = 0
'                TXTQTY.SelLength = Len(TXTQTY.Text)
'                Exit Sub
'            End If
'            If i <> 0 And Val(TXTFREE.Text) <> 0 Then
                If M_EDIT = False And SERIAL_FLAG = True And (Val(TXTFREE.text) + Val(TXTQTY.text)) > (Val(lblactqty.Caption) / Val(LblPack.text)) Then
                    If (MsgBox("AVAILABLE STOCK IN THE PARTICULAR STOCK IS  " & (Val(lblactqty.Caption) / Val(LblPack.text)) & " " & "  Do you want to CONTINUE", vbYesNo, "QUOTATION") = vbNo) Then
                        TXTQTY.SelStart = 0
                        TXTQTY.SelLength = Len(TXTQTY.text)
                        Exit Sub
                    End If
                End If
                If M_EDIT = False And (Val(TXTFREE.text) + Val(TXTQTY.text)) > i Then
                    If (MsgBox("AVAILABLE STOCK IS  " & i & "  Do you want to CONTINUE", vbYesNo, "QUOTATION") = vbNo) Then
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

'Function FILL_BATCHGRID2()
'    FRMEMAIN.Enabled = False
'    FRMEGRDTMP.Visible = True
'
'    Set GRDPOPUP.DataSource = Nothing
'    Set GRDPOPUPITEM.DataSource = Nothing
'    FRMEITEM.Visible = False
'
'    If BATCH_FLAG = True Then
'        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME, PRICE5, PRICE6, PRICE7 From RTRXFILE  WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
'        BATCH_FLAG = False
'    Else
'        PHY_BATCH.Close
'        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME, PRICE5, PRICE6, PRICE7 From RTRXFILE  WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
'        BATCH_FLAG = False
'    End If
'    Set GRDPOPUP.DataSource = PHY_BATCH
'    'TXTITEMCODE.Text = GRDPOPUP.Columns(26)
'    item_change = True
'    TXTPRODUCT.Text = GRDPOPUP.Columns(6)
'    item_change = False
'    GRDPOPUP.Columns(0).Caption = "Serial No."
'    GRDPOPUP.Columns(1).Caption = "QTY"
'    GRDPOPUP.Columns(2).Caption = "VCH No"
'    GRDPOPUP.Columns(3).Caption = "Line No"
'    GRDPOPUP.Columns(4).Caption = "Trx Type"
'    GRDPOPUP.Columns(7).Caption = "" '"Warranty"
'    GRDPOPUP.Columns(8).Caption = ""
'    GRDPOPUP.Columns(21).Caption = "MRP"
'    GRDPOPUP.Columns(25).Caption = "Expiry"
'
'    GRDPOPUP.Columns(0).Width = 1500
'    GRDPOPUP.Columns(1).Width = 900
'    GRDPOPUP.Columns(2).Width = 0
'    GRDPOPUP.Columns(3).Width = 0
'    GRDPOPUP.Columns(4).Width = 0
'    GRDPOPUP.Columns(5).Width = 0
'    GRDPOPUP.Columns(6).Width = 0
'    GRDPOPUP.Columns(7).Width = 0
'    GRDPOPUP.Columns(8).Width = 0
'    GRDPOPUP.Columns(9).Width = 1000
'    GRDPOPUP.Columns(10).Width = 1000
'    GRDPOPUP.Columns(11).Width = 1000
'    GRDPOPUP.Columns(12).Width = 0
'    GRDPOPUP.Columns(13).Width = 0
'    GRDPOPUP.Columns(14).Width = 0
'    GRDPOPUP.Columns(15).Width = 0
'    GRDPOPUP.Columns(16).Width = 0
'    GRDPOPUP.Columns(17).Width = 0
'    GRDPOPUP.Columns(18).Width = 0
'    GRDPOPUP.Columns(19).Width = 0
'    GRDPOPUP.Columns(20).Width = 0
'    GRDPOPUP.Columns(21).Width = 1000
'    GRDPOPUP.Columns(22).Width = 0
'    GRDPOPUP.Columns(23).Width = 0
'    GRDPOPUP.Columns(24).Width = 0
'    GRDPOPUP.Columns(25).Width = 1200
'    GRDPOPUP.Columns(26).Width = 0
'    GRDPOPUP.Columns(27).Width = 0
'    GRDPOPUP.Columns(28).Width = 0
'    GRDPOPUP.Columns(29).Width = 0
'    GRDPOPUP.Columns(30).Width = 0
'    GRDPOPUP.Columns(31).Width = 3000
'    GRDPOPUP.Columns(32).Width = 1000
'    GRDPOPUP.Columns(33).Width = 1000
'    GRDPOPUP.Columns(34).Width = 1000
'    GRDPOPUP.SetFocus
'    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).Text
'    LBLHEAD(9).Visible = True
'    LBLHEAD(0).Visible = True
'End Function

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
                If lblsubdealer.Caption = "D" Then
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
    
    TXTRETAILNOTAX.Tag = Val(TXTRETAILNOTAX.text) - (Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100)
    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(TXTTAX.text) + Val(TxtKFC.Caption)) / 100)) + (TXTRETAILNOTAX.Tag * Val(TxtCessPer.text) / 100) + Val(TxtCessAmt.text)
    'TxtCessPer.Tag = (TXTRETAILNOTAX.Tag * Val(TxtCessPer.Text) / 100) + Val(TxtCessAmt.Text)
    
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text), 2), ".000")
    Else
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(TXTQTY.text), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text) * Val(TXTQTY.text), 2), ".000")
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
            If lblsubdealer.Caption = "D" Then
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
    
    TXTRETAILNOTAX.Tag = Val(TXTRETAILNOTAX.text) - (Val(TXTRETAILNOTAX.text) * Val(TXTDISC.text) / 100)
    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(TXTTAX.text) + Val(TxtKFC.Caption)) / 100)) + (TXTRETAILNOTAX.Tag * Val(TxtCessPer.text) / 100) + Val(TxtCessAmt.text)
    'TxtCessPer.Tag = (TXTRETAILNOTAX.Tag * Val(TxtCessPer.Text) / 100) + Val(TxtCessAmt.Text)
    
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text), 2), ".000")
    Else
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(TXTQTY.text), 2), ".000")
        LblGross.Caption = Format(Round(Val(TXTRETAILNOTAX.text) * Val(TXTQTY.text), 2), ".000")
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
        'Print #1, "No. " & Trim(LBLBILLNO.Caption) & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 58) '& Space(2) & LBLTIME.Caption
        Print #1, AlignRight("Date:" & TXTINVDATE.text, 58)
        Print #1, "TO: " & TxtBillName.text
        Print #1, RepeatString("-", 67)
        Print #1, AlignLeft("Sl", 2) & Space(3) & _
                AlignLeft("Description", 32) & _
                AlignLeft("Qty", 10) & Space(3) & _
                AlignLeft("Rate", 10) & Space(0) & _
                AlignLeft("Amount", 12) '& _
                Chr (27) & Chr(72) '//Bold Ends
    
        Print #1, RepeatString("-", 67)
        'Print #1, RepeatString("-", 67)
    
        For i = 1 To grdsales.rows - 1
            Print #1, AlignLeft(Val(i), 5) & _
                Space(0) & AlignLeft(Mid(grdsales.TextMatrix(i, 2), 1, 31), 31) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 7) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 11) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 12) '& _
                Chr (27) & Chr(72) '//Bold Ends
            Print #1,
        Next i
        Print #1, RepeatString("-", 67)
        
        'Print #1, AlignRight("-------------", 47)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 54) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
            Print #1, AlignRight("DISC AMOUNT ", 54) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 54) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
        End If
        If Val(Txthandle.text) <> 0 Then
            Print #1, AlignRight("HANDLING CHARGE ", 54) & AlignRight((Format(Txthandle.text, "####.00")), 12)
        End If
        If Val(TxtFrieght.text) <> 0 Then
            Print #1, AlignRight("FRIEGHT CHARGE  ", 54) & AlignRight((Format(Txthandle.text, "####.00")), 12)
        End If
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 47) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, AlignRight("Round off ", 54) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
        Print #1, Chr(13)
        If Val(LBLRETAMT.Caption) <> 0 Then Print #1, AlignRight("RETURN AMOUNT ", 54) & AlignRight((Format(Round(LBLRETAMT.Caption, 0), "####.00")), 12)
        Print #1, AlignRight("NET AMOUNT ", 54) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 12)
        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
        Num = CCur(Round(LBLTOTAL.Caption, 0))
        Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & ")", 55)
        Print #1, RepeatString("-", 67)
        'Print #1, RepeatString("-", 67)
'        If OP_BAL > 0 Then
'            Print #1, AlignRight("Old Outstanding", 54) & AlignRight((Format(OP_BAL, "####.00")), 12)
'        End If
'        If RCPT_AMT > 0 Then
'            Print #1, AlignRight("Received Amt", 54) & AlignRight((Format(RCPT_AMT, "####.00")), 12)
'        End If
'        If Not (RCPT_AMT = 0 And OP_BAL = 0) Then
'            Print #1, AlignRight("Total Bal Amt", 54) & AlignRight((Format((Val(lblnetamount.Caption) + OP_BAL) - RCPT_AMT, "####.00")), 12)
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
        Print #1, Space(0) & "State: Kerala(32)"
        Print #1, RepeatString("-", 60)
        Print #1,
        
        Print #1, "Qtn No. " & Trim(LBLBILLNO.Caption) & Space(2) & AlignRight("Date:" & TXTINVDATE.text, 40) '& Space(2) & LBLTIME.Caption
        Print #1, "TO: " & TxtBillName.text '& " (" & Trim(TxtCode.Text) & ")"
        If Trim(TxtBillAddress.text) <> "" Then Print #1, TxtBillAddress.text
        If Trim(TxtPhone.text) <> "" Then Print #1, "Phone: " & TxtPhone.text
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
            RSTHSNCODE.Open "SELECT * from ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTHSNCODE.EOF And RSTHSNCODE.BOF) Then
                HSNCODE = IIf(IsNull(RSTHSNCODE!REMARKS), "", RSTHSNCODE!REMARKS)
            End If
            RSTHSNCODE.Close
            Set RSTHSNCODE = Nothing
            
            Print #1, AlignLeft(Val(i), 3) & _
                Space(0) & AlignLeft(HSNCODE, 8) & _
                Space(0) & AlignLeft(Mid(grdsales.TextMatrix(i, 2), 1, 18), 18) & _
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
        Num = CCur(Round(LBLTOTAL.Caption, 0))
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

Public Function Make_Invoice(BillType As String)
    'If BillType = "HI" Then Exit Function
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTP_RATE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim i, BILL_NUM As Double
    Dim TRXVALUE As Double
    Dim DAY_DATE As String
    Dim MONTH_DATE As String
    Dim YEAR_DATE As String
    Dim E_DATE As Date
    Dim rstBill_Type As String
    
    i = 0
    On Error GoTo ERRHAND
    
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 13) = "" Then GoTo SKIP_1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "' AND BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            grdsales.TextMatrix(i, 14) = RSTTRXFILE!VCH_NO
            grdsales.TextMatrix(i, 15) = RSTTRXFILE!LINE_NO
            grdsales.TextMatrix(i, 16) = RSTTRXFILE!TRX_TYPE
            grdsales.TextMatrix(i, 43) = RSTTRXFILE!TRX_YEAR
            RSTTRXFILE.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
SKIP_1:
    Next i
        
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = '" & BillType & "'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        BILL_NUM = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    TxtCN.text = BILL_NUM
    TXTCNTYPE.text = rstBill_Type
    

    db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & ""
    db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & ""
    
    i = 0
    
    Dim disctype, crdtype As String
    Dim DISCAMT As Double
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        disctype = "A"
        DISCAMT = Val(TXTTOTALDISC.text)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.text) > 0 Then
        disctype = "P"
        DISCAMT = Round(Val(LBLTOTAL.Caption) * Val(TXTTOTALDISC.text) / 100, 2)
    Else
        disctype = "P"
        DISCAMT = 0
    End If
    If lblcredit.Caption = "0" Then crdtype = "Y" Else crdtype = "N"
    
    db.Execute "delete FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & BillType & "' AND VCH_NO = " & BILL_NUM & ""
                                 
    Dim RSTACTCODE As ADODB.Recordset
    Dim CUSTTYPE As String
    CUSTTYPE = "R"
    Set RSTACTCODE = New ADODB.Recordset
    RSTACTCODE.Open "SELECT Type FROM CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTACTCODE.EOF And RSTACTCODE.BOF) Then
        CUSTTYPE = IIf(IsNull(RSTACTCODE!Type) Or RSTACTCODE!Type = "", "R", RSTACTCODE!Type)
    End If
    RSTACTCODE.Close
    Set RSTACTCODE = Nothing
    
    db.Execute "INSERT INTO TRXMAST (TRX_TYPE, TRX_YEAR, VCH_NO, VCH_AMOUNT, NET_AMOUNT, VCH_DATE, ACT_CODE, ACT_NAME, DISCOUNT, C_USER_ID, CREATE_DATE, C_TIME, C_USER_NAME, ADD_AMOUNT, ROUNDED_OFF, PAY_AMOUNT, REF_NO, SLSM_CODE, CHECK_FLAG, POST_FLAG, CFORM_NO, REMARKS, DISC_PERS, AST_PERS, AST_AMNT, BANK_CHARGE, VEHICLE, PHONE, TIN, UID_NO, FRIEGHT, MODIFY_DATE, cr_days, AGENT_CODE, AGENT_NAME, COMM_AMT, BILL_TYPE, CN_REF, BILL_NAME, BILL_ADDRESS)" & _
                            "VALUES ('" & BillType & "', '" & Year(MDIMAIN.DTFROM.Value) & "', " & BILL_NUM & ", " & Val(LBLTOTAL.Caption) & ", " & Val(lblnetamount.Caption) & ", CURDATE(), '" & DataList2.BoundText & "', '" & DataList2.text & "', " & DISCAMT & ", '" & frmLogin.rs!USER_ID & "', CURDATE(), '" & Format(Time, "HH:MM:SS") & "', '" & frmLogin.rs!USER_NAME & "', " & Val(LBLRETAMT.Caption) & ", 0, " & Val(LBLTOTALCOST.Caption) & ", '', " & _
                            " '" & disctype & "', 'I', '" & crdtype & "', '" & Format(Time, "HH:MM:SS") & "', '" & DataList2.text & "', 0, 0, 0, 0, '', '" & Trim(TxtPhone.text) & "', '" & Trim(TXTTIN.text) & "', '', " & Val(TxtFrieght.text) & ", CURDATE(), 0,'','',0,'" & CUSTTYPE & "',Null, '" & Trim(TxtBillName.text) & "', '" & Trim(TxtBillAddress.text) & "')"
                            
                            
    
    Dim rvchno As Long
    Dim rlineno As Long
    Dim rtrxtype As String
    Dim rtrxyear As String
    
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 13) = "" Then GoTo SKIP_3
        rvchno = IIf(grdsales.TextMatrix(i, 14) = "", 0, Val(grdsales.TextMatrix(i, 14)))
        rlineno = IIf(grdsales.TextMatrix(i, 15) = "", 0, Val(grdsales.TextMatrix(i, 15)))
        rtrxtype = IIf(grdsales.TextMatrix(i, 16) = "", "MI", grdsales.TextMatrix(i, 16))
        rtrxyear = IIf(grdsales.TextMatrix(i, 43) = "", "", grdsales.TextMatrix(i, 43))

        db.Execute "INSERT INTO TRXSUB (TRX_TYPE, TRX_YEAR, VCH_NO, line_no, R_VCH_NO, R_LINE_NO, R_TRX_TYPE, R_TRX_YEAR, QTY) VALUES ('" & BillType & "', '" & Year(MDIMAIN.DTFROM.Value) & "', " & BILL_NUM & ", " & i & ", " & rvchno & ", " & rlineno & ", '" & rtrxtype & "', '" & rtrxyear & "', " & Val(grdsales.TextMatrix(i, 3)) & ")"
SKIP_3:
    Next i
    Dim salesprice As Double
    Dim ptrprice As Double
    Dim bill_CST As Double
    Dim Bill_SCHEME As Double
    Dim kfctax As Double
    Dim VCHDESCCRP As String
    VCHDESCCRP = "Issued to     " & Mid(Trim(DataList2.text), 1, 30)
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 13) = "" Then GoTo SKIP_4
        
        Bill_SCHEME = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        Select Case grdsales.TextMatrix(i, 19)
            Case "DN"
                bill_CST = 1
            Case "CN"
                bill_CST = 2
            Case Else
                bill_CST = 0
        End Select
        
        If BillType = "GI" Then
            salesprice = Round(Val(grdsales.TextMatrix(i, 7)) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 47)) / 100), 3)
            ptrprice = Round(salesprice * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 3)
            kfctax = 0
        ElseIf BillType = "WO" Then
            TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(i, 7)) - (Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100)
            TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(grdsales.TextMatrix(i, 9)) + Val(grdsales.TextMatrix(i, 47))) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(i, 40)) / 100) + Val(grdsales.TextMatrix(i, 41))
            salesprice = Val(TXTRETAILNOTAX.Tag)
            ptrprice = Round(salesprice * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 3)
            kfctax = 0
        Else
            If MDIMAIN.lblkfc.Caption = "Y" And IsDate(MDIMAIN.DTKFCSTART.Value) And IsDate(MDIMAIN.DTKFCEND.Value) Then
                If DateValue(TXTINVDATE.text) >= DateValue(MDIMAIN.DTKFCSTART.Value) And DateValue(TXTINVDATE.text) <= DateValue(MDIMAIN.DTKFCEND.Value) Then
                    If Val(grdsales.TextMatrix(i, 9)) = 12 Or Val(grdsales.TextMatrix(i, 9)) = 18 Or Val(grdsales.TextMatrix(i, 9)) = 28 Then
                        kfctax = 1
                        TXTAREA.Tag = 1
                        If Val(grdsales.TextMatrix(i, 47)) = 1 Then
                            ptrprice = Val(grdsales.TextMatrix(i, 6))
                            salesprice = Val(grdsales.TextMatrix(i, 7))
                        Else
                            ptrprice = (Val(grdsales.TextMatrix(i, 7))) / (1 + ((Val(grdsales.TextMatrix(i, 9)) + 1) / 100))
                            salesprice = Round(ptrprice + (ptrprice * Val(grdsales.TextMatrix(i, 9)) / 100), 4)
                            ptrprice = Round(salesprice * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 4)
                            
'                            salesprice = Round(Val(grdsales.TextMatrix(i, 7)) - (Val(grdsales.TextMatrix(i, 6)) * 1 / 100), 4)
'                            ptrprice = Round(salesprice * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 3)
                        End If
                    Else
                        kfctax = 0
                        TXTAREA.Tag = 0
                        salesprice = Round(Val(grdsales.TextMatrix(i, 7)) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 47)) / 100), 3)
                        ptrprice = Round(salesprice * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 3)
                    End If
                Else
                    kfctax = 0
                    TXTAREA.Tag = 0
                    salesprice = Round(Val(grdsales.TextMatrix(i, 7)) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 47)) / 100), 3)
                    ptrprice = Round(salesprice * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 3)
                End If
            Else
                kfctax = 0
                TXTAREA.Tag = 0
                salesprice = Round(Val(grdsales.TextMatrix(i, 7)) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 47)) / 100), 3)
                ptrprice = Round(salesprice * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 3)
            End If
        End If
        
        If BillType = "WO" Then
            'RSTTRXFILE!CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 40))
            'RSTTRXFILE!cess_amt = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 41))
            'db.Execute "INSERT INTO TRXFILE (TRX_TYPE, TRX_YEAR, VCH_NO, VCH_DATE, LINE_NO, CATEGORY, ITEM_CODE, ITEM_NAME, QTY, ITEM_COST, MRP, SALES_PRICE, P_RETAIL, PTR, P_RETAILWOTAX, COM_AMT, COM_FLAG, LOOSE_FLAG, LOOSE_PACK, SALES_TAX, UNIT, VCH_DESC, REF_NO, ISSUE_QTY, CHECK_FLAG, MFGR, CST, BAL_QTY, TRX_TOTAL, LINE_DISC, SCHEME, FREE_QTY, MODIFY_DATE, CREATE_DATE, C_USER_ID, M_USER_ID, SALE_1_FLAG, PACK_TYPE, AREA, KFC_TAX )" & _
                            "VALUES ('" & BillType & "', '" & Year(MDIMAIN.DTFROM.Value) & "', " & BILL_NUM & ", CURDATE(), " & i & ", '" & Trim(grdsales.TextMatrix(i, 25)) & "', '" & grdsales.TextMatrix(i, 13) & "', '" & grdsales.TextMatrix(i, 2) & "', " & Val(grdsales.TextMatrix(i, 3)) & ", " & Val(grdsales.TextMatrix(i, 11)) & ", " & Val(grdsales.TextMatrix(i, 5)) & ", " & salesprice & ", " & salesprice & ", " & ptrprice & ", " & ptrprice & ", " & _
                            " " & Val(grdsales.TextMatrix(i, 24)) & ", 'N', '" & grdsales.TextMatrix(i, 26) & "', " & Val(grdsales.TextMatrix(i, 27)) & ", " & Val(grdsales.TextMatrix(i, 9)) & ", '" & Val(grdsales.TextMatrix(i, 4)) & "',  '" & VCHDESCCRP & "', '" & Trim(grdsales.TextMatrix(i, 10)) & "', 0, '" & Trim(grdsales.TextMatrix(i, 17)) & "', '" & Trim(grdsales.TextMatrix(i, 18)) & "', " & bill_CST & ", 0, " & Val(grdsales.TextMatrix(i, 12)) & ", " & Val(grdsales.TextMatrix(i, 8)) & ", " & _
                            " " & Bill_SCHEME & ", " & Val(grdsales.TextMatrix(i, 20)) & ", CURDATE(), CURDATE(), 'SM', '" & DataList2.BoundText & "', '" & Trim(grdsales.TextMatrix(i, 23)) & "', '" & Trim(grdsales.TextMatrix(i, 30)) & "', '" & Trim(TXTAREA.text) & "', " & kfctax & ")"
            
            db.Execute "INSERT INTO TRXFILE (TRX_TYPE, TRX_YEAR, VCH_NO, VCH_DATE, LINE_NO, CATEGORY, ITEM_CODE, ITEM_NAME, QTY, ITEM_COST, MRP, SALES_PRICE, P_RETAIL, PTR, P_RETAILWOTAX, COM_AMT, COM_FLAG, LOOSE_FLAG, LOOSE_PACK, SALES_TAX, UNIT, VCH_DESC, REF_NO, ISSUE_QTY, CHECK_FLAG, MFGR, CST, BAL_QTY, TRX_TOTAL, LINE_DISC, SCHEME, FREE_QTY, MODIFY_DATE, CREATE_DATE, C_USER_ID, M_USER_ID, SALE_1_FLAG, PACK_TYPE, AREA, KFC_TAX, UN_BILL )" & _
                            "VALUES ('" & BillType & "', '" & Year(MDIMAIN.DTFROM.Value) & "', " & BILL_NUM & ", CURDATE(), " & i & ", '" & Trim(grdsales.TextMatrix(i, 25)) & "', '" & grdsales.TextMatrix(i, 13) & "', '" & grdsales.TextMatrix(i, 2) & "', " & Val(grdsales.TextMatrix(i, 3)) & ", " & Val(grdsales.TextMatrix(i, 11)) & ", " & Val(grdsales.TextMatrix(i, 5)) & ", " & salesprice & ", " & salesprice & ", " & ptrprice & ", " & ptrprice & ", " & _
                            " " & Val(grdsales.TextMatrix(i, 24)) & ", 'N', '" & grdsales.TextMatrix(i, 26) & "', " & Val(grdsales.TextMatrix(i, 27)) & ", " & Val(grdsales.TextMatrix(i, 9)) & ", '" & Val(grdsales.TextMatrix(i, 4)) & "',  '" & VCHDESCCRP & "', '" & Trim(grdsales.TextMatrix(i, 10)) & "', 0, '" & Trim(grdsales.TextMatrix(i, 17)) & "', '" & Trim(grdsales.TextMatrix(i, 18)) & "', " & bill_CST & ", 0, " & Val(grdsales.TextMatrix(i, 12)) & ", 0, " & _
                            " " & Bill_SCHEME & ", " & Val(grdsales.TextMatrix(i, 20)) & ", CURDATE(), CURDATE(), 'SM', '" & DataList2.BoundText & "', '" & Trim(grdsales.TextMatrix(i, 23)) & "', '" & Trim(grdsales.TextMatrix(i, 30)) & "', '" & Trim(TXTAREA.text) & "', " & kfctax & ", '" & Trim(grdsales.TextMatrix(i, 46)) & "')"
        Else
            db.Execute "INSERT INTO TRXFILE (TRX_TYPE, TRX_YEAR, VCH_NO, VCH_DATE, LINE_NO, CATEGORY, ITEM_CODE, ITEM_NAME, QTY, ITEM_COST, MRP, SALES_PRICE, P_RETAIL, PTR, P_RETAILWOTAX, COM_AMT, COM_FLAG, LOOSE_FLAG, LOOSE_PACK, SALES_TAX, UNIT, VCH_DESC, REF_NO, ISSUE_QTY, CHECK_FLAG, MFGR, CST, BAL_QTY, TRX_TOTAL, LINE_DISC, SCHEME, FREE_QTY, MODIFY_DATE, CREATE_DATE, C_USER_ID, M_USER_ID, SALE_1_FLAG, PACK_TYPE, AREA, KFC_TAX, CESS_PER, CESS_AMT, UN_BILL )" & _
                            "VALUES ('" & BillType & "', '" & Year(MDIMAIN.DTFROM.Value) & "', " & BILL_NUM & ", CURDATE(), " & i & ", '" & Trim(grdsales.TextMatrix(i, 25)) & "', '" & grdsales.TextMatrix(i, 13) & "', '" & grdsales.TextMatrix(i, 2) & "', " & Val(grdsales.TextMatrix(i, 3)) & ", " & Val(grdsales.TextMatrix(i, 11)) & ", " & Val(grdsales.TextMatrix(i, 5)) & ", " & salesprice & ", " & salesprice & ", " & ptrprice & ", " & ptrprice & ", " & _
                            " " & Val(grdsales.TextMatrix(i, 24)) & ", 'N', '" & grdsales.TextMatrix(i, 26) & "', " & Val(grdsales.TextMatrix(i, 27)) & ", " & Val(grdsales.TextMatrix(i, 9)) & ", '" & Val(grdsales.TextMatrix(i, 4)) & "',  '" & VCHDESCCRP & "', '" & Trim(grdsales.TextMatrix(i, 10)) & "', 0, '" & Trim(grdsales.TextMatrix(i, 17)) & "', '" & Trim(grdsales.TextMatrix(i, 18)) & "', " & bill_CST & ", 0, " & Val(grdsales.TextMatrix(i, 12)) & ", " & Val(grdsales.TextMatrix(i, 8)) & ", " & _
                            " " & Bill_SCHEME & ", " & Val(grdsales.TextMatrix(i, 20)) & ", CURDATE(), CURDATE(), 'SM', '" & DataList2.BoundText & "', '" & Trim(grdsales.TextMatrix(i, 23)) & "', '" & Trim(grdsales.TextMatrix(i, 30)) & "', '" & Trim(TXTAREA.text) & "', " & kfctax & ", " & Val(grdsales.TextMatrix(i, 40)) & ", " & Val(grdsales.TextMatrix(i, 41)) & ", '" & Trim(grdsales.TextMatrix(i, 46)) & "')"
                            
        End If
SKIP_4:
    Next i


    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 13) = "" Then GoTo SKIP_2
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
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
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 43)) & "' AND BAL_QTY > 0", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 20))) * Val(grdsales.TextMatrix(i, 27)), 3)
                !BAL_QTY = !BAL_QTY - Round((Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 20))) * Val(grdsales.TextMatrix(i, 27)), 3)
                RSTTRXFILE.Update
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            Else
                'BALQTY = 0
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "' AND BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    If (IsNull(RSTTRXFILE!ISSUE_QTY)) Then RSTTRXFILE!ISSUE_QTY = 0
                    If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                    'BALQTY = RSTTRXFILE!BAL_QTY
                    RSTTRXFILE!ISSUE_QTY = RSTTRXFILE!ISSUE_QTY + Round((Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 20))) * Val(grdsales.TextMatrix(i, 27)), 3)
                    RSTTRXFILE!BAL_QTY = RSTTRXFILE!BAL_QTY - Round((Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 20))) * Val(grdsales.TextMatrix(i, 27)), 3)
                    
                    grdsales.TextMatrix(i, 14) = RSTTRXFILE!VCH_NO
                    grdsales.TextMatrix(i, 15) = RSTTRXFILE!LINE_NO
                    grdsales.TextMatrix(i, 16) = RSTTRXFILE!TRX_TYPE
        
                    RSTTRXFILE.Update
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            End If
        End With
SKIP_2:
    Next i
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select * From QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        If Val(TxtCN.text) = 0 Then
            rstBILL!BILL_NO = 0 'Null
        Else
            rstBILL!BILL_NO = TxtCN.text
        End If
        rstBILL!BillType = TXTCNTYPE.text
        rstBILL.Update
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
'    grdsales.FixedRows = 0
'    grdsales.Rows = 1
'    LBLTOTAL.Caption = ""
'    lblnetamount.Caption = ""
'    TXTTOTALDISC.Text = ""
'    txtcommper.Text = ""
'    LBLTOTALCOST.Caption = ""
'    Call AppendSale
    Chkcancel.Value = 0
    MsgBox "Success", vbOKOnly, "EzBiz"
SKIP:
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
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If MDIMAIN.StatusBar.Panels(15).text = "Y" Then
                If TxtName1.Enabled = True Then TxtName1.SetFocus
            Else
                If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            End If
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If txtretail.Enabled = True Then txtretail.SetFocus
            If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
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
    If CMDEXIT.Enabled = False Then cmdRefresh_Click
    
    If currbillno <= 1 Then
        txtBillNo.text = 1
    Else
        txtBillNo.text = currbillno - 1
    End If
    LBLBILLNO.Caption = Val(txtBillNo.text)
    
    flagchange2.Caption = ""
    TXTTIN.text = ""
    TxtCN.text = ""
    TXTCNTYPE.text = ""
    TXTTOTALDISC.text = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
    Chkcancel.Value = 0
    
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    lbltime.Caption = Format(Time, "hh:mm AMPM")
    
    Call txtBillNo_KeyDown(13, 0)
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdnxt_Click()
        
    Dim currbillno As Long
    currbillno = Val(txtBillNo.text)
    If CMDEXIT.Enabled = False Then cmdRefresh_Click
    Dim rstBILL As ADODB.Recordset
    Dim lastbillno As Long
    
    On Error GoTo ERRHAND
    
    TXTDEALER.text = ""
    txtphsearch.text = ""
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) FROM QTNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'QT'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        lastbillno = IIf(IsNull(rstBILL.Fields(0)), 0, rstBILL.Fields(0))
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    If currbillno > lastbillno Then Exit Sub
    txtBillNo.text = currbillno + 1
    LBLBILLNO.Caption = Val(txtBillNo.text)
    
    flagchange2.Caption = ""
    TXTTIN.text = ""
    TxtCN.text = ""
    TXTCNTYPE.text = ""
    TXTTOTALDISC.text = ""
    txtcommper.text = ""
    LBLTOTALCOST.Caption = ""
    Chkcancel.Value = 0
    
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    lbltime.Caption = Format(Time, "hh:mm AMPM")
    
    Call txtBillNo_KeyDown(13, 0)
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
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
                db.Execute "Update QTNSUB set COM_AMT = " & Val(grdsales.TextMatrix(i, 24)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & " "
'                grossamt = Round((Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 14))) * (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 5))), 3)
'                grdsales.TextMatrix(i, 8) = Format(Round(((grossamt / (Val(grdsales.TextMatrix(i, 5)) * (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))))) + ((Val(grdsales.TextMatrix(i, 32)) / Val(grdsales.TextMatrix(i, 5))))), 4), ".0000")
'                LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(i, 32)), ".00")
'                db.Execute "Update RTRXFILE set ITEM_COST = " & Val(grdsales.TextMatrix(i, 8)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
'                db.Execute "Update RTRXFILE set EXPENSE = " & Val(grdsales.TextMatrix(i, 32)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
               
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
                db.Execute "Update QTNSUB set COM_AMT = " & Val(grdsales.TextMatrix(i, 24)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & " "
'                grossamt = Round((Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 14))) * (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 5))), 3)
'                grdsales.TextMatrix(i, 8) = Format(Round(((grossamt / (Val(grdsales.TextMatrix(i, 5)) * (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))))) + ((Val(grdsales.TextMatrix(i, 32)) / Val(grdsales.TextMatrix(i, 5))))), 4), ".0000")
'                LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(i, 32)), ".00")
'                db.Execute "Update RTRXFILE set ITEM_COST = " & Val(grdsales.TextMatrix(i, 8)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
'                db.Execute "Update RTRXFILE set EXPENSE = " & Val(grdsales.TextMatrix(i, 32)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
               
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
            txtretail.text = Round(Val(TXTRETAILNOTAX.text) + Val(txtmrpbt.text) * (Val(TXTTAX.text)) / 100, 4)
        End If
        If OPTVAT.Value = True Then
            txtretail.text = Round(Val(TXTRETAILNOTAX.text) + Val(TXTRETAILNOTAX.text) * (Val(TXTTAX.text)) / 100, 4)
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
    If MDIMAIN.lblgst.Caption = "R" Then Call cost_calculate
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
    CMBBillType.Enabled = True
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
            If CMBBillType.ListIndex = -1 And M_EDIT = False And CMBBillType.Visible = True Then Exit Sub
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
                TXTDISC.SetFocus
            End If
        Case vbKeyEscape
            TXTRETAILNOTAX.Enabled = True
            TXTRETAILNOTAX.SetFocus
        Case 116
            Call FILL_PREVIIOUSRATE
        Case 117
            If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
        Case vbKeyDown
            If CMBBillType.ListIndex = -1 And M_EDIT = False And CMBBillType.Visible = True Then Exit Sub
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
'    If Val(txtretail.Text) = 0 Then
'        optnet.value = True
'        TXTTAX.Text = 0
'    End If
    If OPTVAT.Value = False Then TXTTAX.text = 0
    TXTRETAILNOTAX.text = Round(Val(txtretail.text) * 100 / ((Val(TXTTAX.text)) + 100), 4)
    TXTRETAILNOTAX.text = Format(Val(TXTRETAILNOTAX.text), "0.0000")
    txtretail.text = Format(Val(txtretail.text), "0.0000")
    
    If UCase(txtcategory.text) = "PARDHA" Or UCase(txtcategory.text) = "CLOTHES" Then
        If Val(TXTRETAILNOTAX.text) < 1000 Then
            TXTTAX.text = "5"
        Else
            TXTTAX.text = "12"
        End If
    End If
    
    TXTRETAILNOTAX.text = Round(Val(txtretail.text) * 100 / ((Val(TXTTAX.text)) + 100), 4)
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
    
    Dim discamnt As Double
    discamnt = 0
    If UCase(txtcategory.text) = "SERVICE CHARGE" Then
        txtNetrate.text = Format(Val(txtretail.text), "0.00")
    Else
'        discamnt = Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
'        txtNetrate.Text = Format(Round(Val(TXTRETAILNOTAX.Text) - discamnt, 2), ".00")
        
'        txtNetrate.Text = Val(txtNetrate.Text) + (Val(TXTRETAILNOTAX.Text) - (Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100)) * Val(TxtCessPer) / 100
'        txtNetrate.Text = Val(txtNetrate.Text) + Round(Val(TxtCessAmt.Text), 3)
        
        txtNetrate.text = (Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * (Val(TXTTAX.text) + Val(TxtKFC.Caption)) / 100)) + (Val(TXTRETAILNOTAX.text) * Val(TxtCessPer.text) / 100) + Val(TxtCessAmt.text)
        txtNetrate.text = Format(Val(txtNetrate.text), "0.00")
    End If
    
    If MDIMAIN.lblgst.Caption = "R" Then Call cost_calculate
    Call CHANGEBOXCOLOR(txtretail, False)
    'TXTDISC.Tag = 0
    'TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
    'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag), ".000")
End Sub

Private Sub CMBBillType_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMBBillType.ListIndex = -1 And M_EDIT = False Then Exit Sub
            If Val(lblOr_Pack.Caption) = 0 Then lblOr_Pack.Caption = 1
            If Val(LblPack.text) <> Val(lblOr_Pack.Caption) Then
                'txtretail.Text = Val(lblcase.Caption) * Val(LblPack.Text)
                If CMBBillType.ListIndex = 0 Then
                    txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                ElseIf CMBBillType.ListIndex = 1 Then
                    txtretail.text = (Val(lblLWPrice.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    TXTRETAILNOTAX.text = (Val(lblLWPrice.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    txtNetrate.text = (Val(lblLWPrice.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                ElseIf CMBBillType.ListIndex = 2 Then
                    txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                ElseIf CMBBillType.ListIndex = 4 Then
                    txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                ElseIf CMBBillType.ListIndex = 5 Then
                    txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                ElseIf CMBBillType.ListIndex = 6 Then
                    txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                ElseIf CMBBillType.ListIndex = 3 Then
                    txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                    txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                End If
            Else
                If CMBBillType.ListIndex = 0 Then
                    txtretail.text = Val(lblretail.Caption)
                    TXTRETAILNOTAX.text = Val(lblretail.Caption)
                    txtNetrate.text = Val(lblretail.Caption)
                ElseIf CMBBillType.ListIndex = 1 Then
                    txtretail.text = Val(lblwsale.Caption)
                    TXTRETAILNOTAX.text = Val(lblwsale.Caption)
                    txtNetrate.text = Val(lblwsale.Caption)
                ElseIf CMBBillType.ListIndex = 2 Then
                    txtretail.text = Val(lblvan.Caption)
                    TXTRETAILNOTAX.text = Val(lblvan.Caption)
                    txtNetrate.text = Val(lblvan.Caption)
                ElseIf CMBBillType.ListIndex = 4 Then
                    txtretail.text = Val(lblprice5.Caption)
                    TXTRETAILNOTAX.text = Val(lblprice5.Caption)
                    txtNetrate.text = Val(lblprice5.Caption)
                ElseIf CMBBillType.ListIndex = 5 Then
                    txtretail.text = Val(lblprice6.Caption)
                    TXTRETAILNOTAX.text = Val(lblprice6.Caption)
                    txtNetrate.text = Val(lblprice6.Caption)
                ElseIf CMBBillType.ListIndex = 6 Then
                    txtretail.text = Val(lblprice7.Caption)
                    TXTRETAILNOTAX.text = Val(lblprice7.Caption)
                    txtNetrate.text = Val(lblprice7.Caption)
                ElseIf CMBBillType.ListIndex = 3 Then
                    txtretail.text = Val(LBLMRP.Caption)
                    TXTRETAILNOTAX.text = Val(LBLMRP.Caption)
                    txtNetrate.text = Val(LBLMRP.Caption)
                End If
            End If
            
            If Val(TxtCessPer.text) <> 0 Or Val(TxtCessAmt.text) <> 0 Then
                TXTRETAILNOTAX.text = (Val(txtretail.text) - Val(TxtCessAmt.text)) / (1 + (Val(TXTTAX.text) / 100) + (Val(TxtCessPer.text) / 100))
                txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 3)
                TXTRETAILNOTAX.text = Val(txtretail.text)
            End If
                        
            If MDIMAIN.LblKFCNet.Caption <> "N" Then
                Call txtNetrate_LostFocus
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
            End If
    
'            If MDIMAIN.StatusBar.Panels(14).Text <> "Y" Then
'                Call TXTRETAILNOTAX_LostFocus
'            Else
'                Call TXTRETAIL_LostFocus
'            End If
            
            'If M_EDIT = False Then
'                Select Case CMBBillType.ListIndex
'                    Case 0
'                        TXTRETAILNOTAX.Text = Val(lblretail.Caption)
'                        txtretail.Text = Val(lblretail.Caption)
'                    Case 1
'                        TXTRETAILNOTAX.Text = Val(lblwsale.Caption)
'                        txtretail.Text = Val(lblwsale.Caption)
'                    Case 2
'                        TXTRETAILNOTAX.Text = Val(lblvan.Caption)
'                        txtretail.Text = Val(lblvan.Caption)
'                End Select
            'End If
            'CMBBillType.Enabled = False
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                txtretail.Enabled = True
                txtretail.SetFocus
            Else
                If Val(TXTTAX.text) = 0 Then
                    TXTTAX.Enabled = True
                    TXTTAX.SetFocus
                Else
                    If MDIMAIN.LblKFCNet.Caption <> "N" Then
                        txtNetrate.Enabled = True
                        txtNetrate.SetFocus
                    Else
                        If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                            TXTRETAILNOTAX.Enabled = True
                            TXTRETAILNOTAX.SetFocus
                        Else
                            If Val(TxtMRP.text) <> 0 And Val(TxtMRP.text) = Val(TXTRETAILNOTAX.text) And mrpplus = True Then
                                TXTRETAILNOTAX.Enabled = True
                                TXTRETAILNOTAX.SetFocus
                            ElseIf Val(TxtMRP.text) <> 0 And Val(Round(Val(TxtMRP.text), 2)) = Val(Round(Val(txtretail.text), 2)) And mrpplus = False Then
                                txtNetrate.Enabled = True
                                txtNetrate.SetFocus
                            Else
                                txtretail.Enabled = True
                                txtretail.SetFocus
                            End If
                        End If
                    End If
                End If
            End If
        Case vbKeyEscape
            'CMBBillType.Enabled = False
            TXTFREE.Enabled = True
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub CMBBillType_LostFocus()
    If CMBBillType.ListIndex = -1 And M_EDIT = False Then Exit Sub
    If M_EDIT = False Then
        If Val(LblPack.text) <> Val(lblOr_Pack.Caption) Then
            'txtretail.Text = Val(lblcase.Caption) * Val(LblPack.Text)
            If CMBBillType.ListIndex = 0 Then
                txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            ElseIf CMBBillType.ListIndex = 1 Then
                txtretail.text = (Val(lblLWPrice.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                TXTRETAILNOTAX.text = (Val(lblLWPrice.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                txtNetrate.text = (Val(lblLWPrice.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            ElseIf CMBBillType.ListIndex = 2 Then
                txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            ElseIf CMBBillType.ListIndex = 4 Then
                txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            ElseIf CMBBillType.ListIndex = 5 Then
                txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            ElseIf CMBBillType.ListIndex = 6 Then
                txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            ElseIf CMBBillType.ListIndex = 3 Then
                txtretail.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                TXTRETAILNOTAX.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
                txtNetrate.text = (Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.text)
            End If
        Else
            If CMBBillType.ListIndex = 0 Then
                txtretail.text = Val(lblretail.Caption)
                TXTRETAILNOTAX.text = Val(lblretail.Caption)
                txtNetrate.text = Val(lblretail.Caption)
            ElseIf CMBBillType.ListIndex = 1 Then
                txtretail.text = Val(lblwsale.Caption)
                TXTRETAILNOTAX.text = Val(lblwsale.Caption)
                txtNetrate.text = Val(lblwsale.Caption)
            ElseIf CMBBillType.ListIndex = 2 Then
                txtretail.text = Val(lblvan.Caption)
                TXTRETAILNOTAX.text = Val(lblvan.Caption)
                txtNetrate.text = Val(lblvan.Caption)
            ElseIf CMBBillType.ListIndex = 4 Then
                txtretail.text = Val(lblprice5.Caption)
                TXTRETAILNOTAX.text = Val(lblprice5.Caption)
                txtNetrate.text = Val(lblprice5.Caption)
            ElseIf CMBBillType.ListIndex = 5 Then
                txtretail.text = Val(lblprice6.Caption)
                TXTRETAILNOTAX.text = Val(lblprice6.Caption)
                txtNetrate.text = Val(lblprice6.Caption)
            ElseIf CMBBillType.ListIndex = 6 Then
                txtretail.text = Val(lblprice7.Caption)
                TXTRETAILNOTAX.text = Val(lblprice7.Caption)
                txtNetrate.text = Val(lblprice7.Caption)
            ElseIf CMBBillType.ListIndex = 3 Then
                txtretail.text = Val(LBLMRP.Caption)
                TXTRETAILNOTAX.text = Val(LBLMRP.Caption)
                txtNetrate.text = Val(LBLMRP.Caption)
            End If
        End If
        
        If Val(TxtCessPer.text) <> 0 Or Val(TxtCessAmt.text) <> 0 Then
            TXTRETAILNOTAX.text = (Val(txtretail.text) - Val(TxtCessAmt.text)) / (1 + (Val(TXTTAX.text) / 100) + (Val(TxtCessPer.text) / 100))
            txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 3)
            TXTRETAILNOTAX.text = Val(txtretail.text)
        End If
                    
        If MDIMAIN.LblKFCNet.Caption <> "N" Then
            Call txtNetrate_LostFocus
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
        End If
        
    End If
End Sub

'Private Sub CMDRESET_Click()
'
'    If grdsales.Rows <= 1 Then Exit Sub
'    If MsgBox("Are you sure you want to reset the items?", vbYesNo, "BILL..") = vbNo Then Exit Sub
'    Dim RSTUNBILL As ADODB.Recordset
'    Dim i As Long
'
'    On Error GoTo Errhand
'    For i = 1 To grdsales.Rows - 1
'        Set RSTUNBILL = New ADODB.Recordset
'        RSTUNBILL.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'        With RSTUNBILL
'            If Not (.EOF And .BOF) Then
'                grdsales.TextMatrix(i, 46) = IIf(IsNull(RSTUNBILL!UN_BILL), "N", RSTUNBILL!UN_BILL)
'                db.Execute "Update TRXFILE set UN_BILL = '" & Trim(grdsales.TextMatrix(i, 46)) & "' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & ""
'            End If
'        End With
'        RSTUNBILL.Close
'        Set RSTUNBILL = Nothing
'    Next i
'
'    Exit Sub
'Errhand:
'    MsgBox Err.Description, , "EzBiz"
'End Sub

Private Sub TxtAddRate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("."), Asc("-")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CmdAddRate_Click()
    If TxtAddRate.text = "" Then Exit Sub
    Dim i As Integer
    For i = 1 To grdsales.rows - 1
        Screen.MousePointer = vbHourglass
        TxtAddRate.Tag = Val(grdsales.TextMatrix(i, 7)) + (Val(grdsales.TextMatrix(i, 7)) * Val(TxtAddRate.text) / 100)
        grdsales.TextMatrix(i, 6) = Format(Round(Val(TxtAddRate.Tag) * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 4), "0.0000")
        grdsales.TextMatrix(i, 22) = Format(Round(Val(TxtAddRate.Tag) * 100 / ((Val(grdsales.TextMatrix(i, 9))) + 100), 4), "0.0000")
        grdsales.TextMatrix(i, 7) = Format(Val(TxtAddRate.Tag), "0.000")
        
        TXTDISC.Tag = 0
        If UCase(grdsales.TextMatrix(i, 25)) = "PARDHA" Or UCase(grdsales.TextMatrix(i, 25)) = "CLOTHES" Then
            If Val(grdsales.TextMatrix(i, 6)) < 1000 Then
               grdsales.TextMatrix(i, 9) = "5"
            Else
                grdsales.TextMatrix(i, 9) = "12"
            End If
        End If
        
        If Trim(lblIGST.Caption) <> "Y" And MDIMAIN.lblkfc.Caption = "Y" And IsDate(MDIMAIN.DTKFCSTART.Value) And IsDate(MDIMAIN.DTKFCEND.Value) Then
            If DateValue(TXTINVDATE.text) >= DateValue(MDIMAIN.DTKFCSTART.Value) And DateValue(TXTINVDATE.text) <= DateValue(MDIMAIN.DTKFCEND.Value) Then
                If Val(grdsales.TextMatrix(i, 9)) = 12 Or Val(grdsales.TextMatrix(i, 9)) = 18 Or Val(grdsales.TextMatrix(i, 9)) = 28 Then
                    grdsales.TextMatrix(i, 47) = 1
                Else
                    grdsales.TextMatrix(i, 47) = 0
                End If
            End If
        Else
7            grdsales.TextMatrix(i, 47) = 0
        End If

        TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(i, 6)) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 8)) / 100)
        TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(grdsales.TextMatrix(i, 9)) + Val(grdsales.TextMatrix(i, 47))) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(i, 40)) / 100) + Val(grdsales.TextMatrix(i, 41))
        If UCase(grdsales.TextMatrix(i, 25)) = "SERVICE CHARGE" Then
            'TXTDISC.Tag = Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
            grdsales.TextMatrix(i, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag), 3), ".000")
            grdsales.TextMatrix(i, 34) = Format(Round(Val(grdsales.TextMatrix(i, 6)), 4), ".0000")
        Else
            'TXTDISC.Tag = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 7)) * Val(grdsales.TextMatrix(i, 8)) / 100
            grdsales.TextMatrix(i, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(i, 3)), 3), ".000")
            grdsales.TextMatrix(i, 34) = Format(Round(Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3)), 4), ".0000")
        End If
        TXTDISC.Tag = (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 3))
        db.Execute "Update QTNSUB set KFC_TAX = " & Val(grdsales.TextMatrix(i, 47)) & ", SALES_PRICE = " & Val(grdsales.TextMatrix(i, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(i, 7)) & ", PTR = " & Val(grdsales.TextMatrix(i, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(i, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(i, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(i, 32)) & ""
    Next i
    
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
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
    
    
    'db.BeginTrans
    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
    Else
        db.Execute "Update QTNMAST set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
    End If
    'db.CommitTrans
    Screen.MousePointer = vbNormal
    
    TxtAddRate.text = ""
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub TxtItemName_Change()
    On Error GoTo ERRHAND
    If flagchange2.Caption <> "1" Then
        If ITEM_REC.State = 1 Then ITEM_REC.Close
        ITEM_REC.Open "Select DISTINCT ITEM_CODE,ITEM_NAME From ITEMMAST  WHERE ITEM_NAME Like '" & TxtItemName.text & "%'ORDER BY ITEM_NAME", db, adOpenForwardOnly
        
        If (ITEM_REC.EOF And ITEM_REC.BOF) Then
            LBLitem.Caption = ""
        Else
            LBLitem.Caption = ITEM_REC!ITEM_NAME
        End If
        Set DataList3.RowSource = ITEM_REC
        DataList3.ListField = "ITEM_NAME"
        DataList3.BoundColumn = "ITEM_CODE"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    'TxtItemName.Text = ""
    
    TxtItemName.text = ""
End Sub

Private Sub TxtItemName_GotFocus()
    TxtItemName.SelStart = 0
    TxtItemName.SelLength = Len(TxtItemName.text)
End Sub

Private Sub TxtItemName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList3.VisibleCount = 0 Then Exit Sub
            DataList3.SetFocus
        Case vbKeyEscape
            TxtItemName.Visible = False
            DataList3.Visible = False
            grdsales.SetFocus
    End Select
End Sub

Private Sub TxtItemName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList3_Click()
    TxtItemName.text = DataList3.text
    LBLitem.Caption = TxtItemName.text
End Sub

Private Sub DataList3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList3.text = "" Then Exit Sub
            If IsNull(DataList3.SelectedItem) Then
                MsgBox "Select Item From List", vbOKOnly, "Sales Bill..."
                DataList3.SetFocus
                Exit Sub
            End If
            
            If DataList3.BoundText = grdsales.TextMatrix(grdsales.Row, 13) Then GoTo SKIP_ITEMCHANGE
            Dim RSTTRXFILE As ADODB.Recordset
            Dim i As Long
            
            
            '====NEW
            grdsales.TextMatrix(grdsales.Row, 1) = DataList3.BoundText
            grdsales.TextMatrix(grdsales.Row, 2) = DataList3.text
            
            grdsales.TextMatrix(grdsales.Row, 33) = DataList3.text
            grdsales.TextMatrix(grdsales.Row, 13) = DataList3.BoundText
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With RSTTRXFILE
                If Not (.EOF And .BOF) Then
                    grdsales.TextMatrix(grdsales.Row, 25) = IIf(IsNull(RSTTRXFILE!Category), "", RSTTRXFILE!Category)
                    grdsales.TextMatrix(grdsales.Row, 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(grdsales.TextMatrix(grdsales.Row, 27)))
                    grdsales.TextMatrix(grdsales.Row, 9) = IIf(IsNull(RSTTRXFILE!SALES_TAX), 0, RSTTRXFILE!SALES_TAX)
                    grdsales.TextMatrix(grdsales.Row, 45) = IIf(IsNull(RSTTRXFILE!SALES_TAX), 0, RSTTRXFILE!SALES_TAX)
                    grdsales.TextMatrix(grdsales.Row, 5) = IIf(IsNull(RSTTRXFILE!MRP), 0, RSTTRXFILE!MRP)
                    grdsales.TextMatrix(grdsales.Row, 40) = IIf(IsNull(RSTTRXFILE!CESS_PER), 0, RSTTRXFILE!CESS_PER)
                    grdsales.TextMatrix(grdsales.Row, 41) = IIf(IsNull(RSTTRXFILE!cess_amt), 0, RSTTRXFILE!cess_amt)
                                        
                    Dim SELL_PRICE As Double
                    SELL_PRICE = 0
                    Select Case cmbtype.ListIndex
                        Case 0
                            SELL_PRICE = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                        Case 1
                            SELL_PRICE = IIf(IsNull(RSTTRXFILE!P_WS), 0, RSTTRXFILE!P_WS)
                        Case 2
                            SELL_PRICE = IIf(IsNull(RSTTRXFILE!P_VAN), 0, RSTTRXFILE!P_VAN)
                            
                        Case 3
                            SELL_PRICE = IIf(IsNull(RSTTRXFILE!MRP), 0, RSTTRXFILE!MRP)
                            If SELL_PRICE = 0 Then
                                SELL_PRICE = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                            End If
                        Case 4
                            SELL_PRICE = IIf(IsNull(RSTTRXFILE!PRICE5), 0, RSTTRXFILE!PRICE5)
                        Case 5
                            SELL_PRICE = IIf(IsNull(RSTTRXFILE!PRICE6), 0, RSTTRXFILE!PRICE6)
                        Case 6
                            SELL_PRICE = IIf(IsNull(RSTTRXFILE!PRICE7), 0, RSTTRXFILE!PRICE7)
                        Case Else
                            SELL_PRICE = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
                    End Select
'                    If Val(TxtCessPer.text) <> 0 Or Val(TxtCessAmt.text) <> 0 Then
'                        TXTRETAILNOTAX.text = (Val(txtretail.text) - Val(TxtCessAmt.text)) / (1 + (Val(TXTTAX.text) / 100) + (Val(TxtCessPer.text) / 100))
'                        txtretail.text = Round(Val(TXTRETAILNOTAX.text) + (Val(TXTRETAILNOTAX.text) * Val(TXTTAX.text) / 100), 3)
'                        TXTRETAILNOTAX.text = Val(txtretail.text)
'                    End If
'
                    If Val(grdsales.TextMatrix(grdsales.Row, 40)) <> 0 Or Val(grdsales.TextMatrix(grdsales.Row, 41)) <> 0 Then
                        If MDIMAIN.StatusBar.Panels(14).text = "Y" Then
                            If SELL_PRICE <> 0 Then
                                TXTRETAILNOTAX.Tag = (SELL_PRICE - Val(grdsales.TextMatrix(grdsales.Row, 41))) / (1 + ((Val(grdsales.TextMatrix(grdsales.Row, 9))) / 100) + (Val(grdsales.TextMatrix(grdsales.Row, 40)) / 100))
                                grdsales.TextMatrix(grdsales.Row, 7) = Round(Val(TXTRETAILNOTAX.Tag) + (Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(grdsales.Row, 9)) / 100), 4)
                                grdsales.TextMatrix(grdsales.Row, 7) = Round(Val(txtretail.text) * 100 / ((Val(TXTTAX.text)) + 100), 4)
                                grdsales.TextMatrix(grdsales.Row, 6) = Format(Val(TXTRETAILNOTAX.Tag), "0.0000")
                                grdsales.TextMatrix(grdsales.Row, 22) = Val(grdsales.TextMatrix(grdsales.Row, 7))
                            End If
                            grdsales.TextMatrix(grdsales.Row, 7) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) + Val(grdsales.TextMatrix(grdsales.Row, 6)) * (Val(grdsales.TextMatrix(grdsales.Row, 9))) / 100, 4), "0.0000")
                            grdsales.TextMatrix(grdsales.Row, 21) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) + Val(grdsales.TextMatrix(grdsales.Row, 6)) * (Val(grdsales.TextMatrix(grdsales.Row, 9))) / 100, 4), "0.0000")
                            grdsales.TextMatrix(grdsales.Row, 39) = Val(grdsales.TextMatrix(grdsales.Row, 7))
                        Else
                            If SELL_PRICE <> 0 Then
                                grdsales.TextMatrix(grdsales.Row, 7) = SELL_PRICE
                                grdsales.TextMatrix(grdsales.Row, 21) = SELL_PRICE
                            End If
                            grdsales.TextMatrix(grdsales.Row, 6) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) * 100 / ((Val(grdsales.TextMatrix(grdsales.Row, 9))) + 100), 4), "0.0000")
                            grdsales.TextMatrix(grdsales.Row, 22) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) * 100 / ((Val(grdsales.TextMatrix(grdsales.Row, 9))) + 100), 4), "0.0000")
                        End If
                    Else
                        If MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
                            If SELL_PRICE <> 0 Then
                                grdsales.TextMatrix(grdsales.Row, 6) = SELL_PRICE
                                grdsales.TextMatrix(grdsales.Row, 22) = SELL_PRICE
                            End If
                            grdsales.TextMatrix(grdsales.Row, 7) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) + Val(grdsales.TextMatrix(grdsales.Row, 6)) * (Val(grdsales.TextMatrix(grdsales.Row, 9))) / 100, 4), "0.0000")
                            grdsales.TextMatrix(grdsales.Row, 21) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) + Val(grdsales.TextMatrix(grdsales.Row, 6)) * (Val(grdsales.TextMatrix(grdsales.Row, 9))) / 100, 4), "0.0000")
                        Else
                            If SELL_PRICE <> 0 Then
                                grdsales.TextMatrix(grdsales.Row, 7) = SELL_PRICE
                                grdsales.TextMatrix(grdsales.Row, 21) = SELL_PRICE
                            End If
                            grdsales.TextMatrix(grdsales.Row, 6) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) * 100 / ((Val(grdsales.TextMatrix(grdsales.Row, 9))) + 100), 4), "0.0000")
                            grdsales.TextMatrix(grdsales.Row, 22) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 7)) * 100 / ((Val(grdsales.TextMatrix(grdsales.Row, 9))) + 100), 4), "0.0000")
                        End If
                        grdsales.TextMatrix(grdsales.Row, 39) = SELL_PRICE
                    End If
                    
                    grdsales.TextMatrix(grdsales.Row, 24) = IIf(IsNull(RSTTRXFILE!COM_AMT), 0, RSTTRXFILE!COM_AMT)
                    grdsales.TextMatrix(grdsales.Row, 17) = IIf(IsNull(RSTTRXFILE!check_flag), "V", RSTTRXFILE!check_flag)
                    grdsales.TextMatrix(grdsales.Row, 18) = IIf(IsNull(RSTTRXFILE!MANUFACTURER), "", RSTTRXFILE!MANUFACTURER)
                    grdsales.TextMatrix(grdsales.Row, 28) = IIf(IsNull(RSTTRXFILE!WARRANTY), 0, RSTTRXFILE!WARRANTY)
                    grdsales.TextMatrix(grdsales.Row, 29) = IIf(IsNull(RSTTRXFILE!WARRANTY_TYPE), "", RSTTRXFILE!WARRANTY_TYPE)
                    grdsales.TextMatrix(grdsales.Row, 30) = IIf(IsNull(RSTTRXFILE!PACK_TYPE), "Nos", RSTTRXFILE!PACK_TYPE)
                    'grdsales.TextMatrix(grdsales.Row, 31) = IIf(IsNull(RSTTRXFILE!ST_RATE), "", RSTTRXFILE!ST_RATE)
                    
                    
                    grdsales.TextMatrix(grdsales.Row, 42) = IIf(IsNull(RSTTRXFILE!BARCODE), "", RSTTRXFILE!BARCODE)
                    grdsales.TextMatrix(grdsales.Row, 44) = "" 'IIf(IsNull(RSTTRXFILE!ITEM_SPEC), "", RSTTRXFILE!ITEM_SPEC)
                    grdsales.TextMatrix(grdsales.Row, 46) = IIf(IsNull(RSTTRXFILE!UN_BILL), "N", RSTTRXFILE!UN_BILL)
                    
                    TXTDISC.Tag = 0
                    
                    
                    
                    TXTDISC.Tag = 0
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "PARDHA" Or UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "CLOTHES" Then
                        If Val(grdsales.TextMatrix(grdsales.Row, 6)) < 1000 Then
                           grdsales.TextMatrix(grdsales.Row, 9) = "5"
                        Else
                            grdsales.TextMatrix(grdsales.Row, 9) = "12"
                        End If
                    End If
                    
                    If Trim(lblIGST.Caption) <> "Y" And MDIMAIN.lblkfc.Caption = "Y" And IsDate(MDIMAIN.DTKFCSTART.Value) And IsDate(MDIMAIN.DTKFCEND.Value) Then
                        If DateValue(TXTINVDATE.text) >= DateValue(MDIMAIN.DTKFCSTART.Value) And DateValue(TXTINVDATE.text) <= DateValue(MDIMAIN.DTKFCEND.Value) Then
                            If Val(grdsales.TextMatrix(grdsales.Row, 9)) = 12 Or Val(grdsales.TextMatrix(grdsales.Row, 9)) = 18 Or Val(grdsales.TextMatrix(grdsales.Row, 9)) = 28 Then
                                grdsales.TextMatrix(grdsales.Row, 47) = 1
                            Else
                                grdsales.TextMatrix(grdsales.Row, 47) = 0
                            End If
                        End If
                    Else
                        grdsales.TextMatrix(grdsales.Row, 47) = 0
                    End If
                    
                    TXTRETAILNOTAX.Tag = Val(grdsales.TextMatrix(grdsales.Row, 6)) - (Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100)
                    TXTRETAILNOTAX.Tag = (TXTRETAILNOTAX.Tag + (TXTRETAILNOTAX.Tag * (Val(grdsales.TextMatrix(grdsales.Row, 9)) + Val(grdsales.TextMatrix(grdsales.Row, 47))) / 100)) + (TXTRETAILNOTAX.Tag * Val(grdsales.TextMatrix(grdsales.Row, 40)) / 100) + Val(grdsales.TextMatrix(grdsales.Row, 41))
                    If UCase(grdsales.TextMatrix(grdsales.Row, 25)) = "SERVICE CHARGE" Then
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag), 3), ".000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)), 4), ".0000")
                    Else
                        'TXTDISC.Tag = Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 7)) * Val(grdsales.TextMatrix(grdsales.Row, 8)) / 100
                        grdsales.TextMatrix(grdsales.Row, 12) = Format(Round(Val(TXTRETAILNOTAX.Tag) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 3), ".000")
                        grdsales.TextMatrix(grdsales.Row, 34) = Format(Round(Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 3)), 4), ".0000")
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
                                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                                LBLFOT.Caption = ""
                            Case Else
                                If MDIMAIN.lblgst.Caption = "R" Then
                                    LBLGROSSAMT.Caption = Round(Val(LBLGROSSAMT.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 8)) / 100, 2)
                                End If
                                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
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
                    'lbltotbal.Caption = Format(Round(Val(lblOB.Caption) + Val(lblnetamount.Caption), 2), "0.00")
                    Call COSTCALCULATION
                    
                    'db.BeginTrans
                    TXTDISC.Tag = (Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(grdsales.TextMatrix(grdsales.Row, 6))) * Val(grdsales.TextMatrix(grdsales.Row, 3))
                    db.Execute "Update TRXFILE set KFC_TAX = " & Val(grdsales.TextMatrix(grdsales.Row, 47)) & ", SALES_PRICE = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", P_RETAIL = " & Val(grdsales.TextMatrix(grdsales.Row, 7)) & ", PTR = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", P_RETAILWOTAX = " & Val(grdsales.TextMatrix(grdsales.Row, 6)) & ", TRX_TOTAL = " & Val(grdsales.TextMatrix(grdsales.Row, 12)) & ", SCHEME = " & Val(TXTDISC.Tag) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
                    If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                        db.Execute "Update QTNMAST set COMM_AMT = " & Val(lblcomamt.text) & ", PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    Else
                        db.Execute "Update QTNMAST set COMM_AMT = 0, PAY_AMOUNT = " & Val(LBLTOTALCOST.Caption) & ", ADD_AMOUNT = " & Val(LBLRETAMT.Caption) & ", DISCOUNT = " & Val(TXTTOTALDISC.text) & ", DISC_PERS  = " & Val(TXTTOTALDISC.text) & ", VCH_AMOUNT = " & Val(LBLTOTAL.Caption) & ", NET_AMOUNT = " & Val(lblnetamount.Caption) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & ""
                    End If
                    'db.Execute "Update DBTPYMT set INV_AMT = " & Val(lblnetamount.Caption) & ", COMM_AMT = " & Val(lblcomamt.text) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'DR' AND INV_TRX_TYPE = 'QT'"
                    
                    grdsales.TextMatrix(grdsales.Row, 14) = ""
                    grdsales.TextMatrix(grdsales.Row, 15) = ""
                    grdsales.TextMatrix(grdsales.Row, 16) = ""
                    grdsales.TextMatrix(grdsales.Row, 43) = ""
                    
'                    RSTTRXFILE!GROSS_AMOUNT = Val(grdsales.TextMatrix(grdsales.Row, 34))
'                    RSTTRXFILE!DN_NO = Val(grdsales.TextMatrix(grdsales.Row, 35))
'                    If IsDate(grdsales.TextMatrix(grdsales.Row, 36)) Then
'                        RSTTRXFILE!DN_DATE = IIf(IsDate(grdsales.TextMatrix(grdsales.Row, 36)), Format(grdsales.TextMatrix(grdsales.Row, 36), "DD/MM/YYYY"), Null)
'                    End If
'                    RSTTRXFILE!DN_LINENO = Val(grdsales.TextMatrix(grdsales.Row, 37))
'
                    
                End If
            End With
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
                                                                
            db.Execute "delete FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & ""
    
            db.Execute "START TRANSACTION"
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * FROM QTNSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='QT' AND VCH_NO = " & Val(txtBillNo.text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(grdsales.Row, 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE.AddNew
                RSTTRXFILE!TRX_TYPE = "QT"
                RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
                RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(grdsales.Row, 32))
            End If
            RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
            RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(grdsales.Row, 13)
            RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(grdsales.Row, 2)
            RSTTRXFILE!QTY = Val(grdsales.TextMatrix(grdsales.Row, 3))
            RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(grdsales.Row, 11))
            If Val(grdsales.TextMatrix(grdsales.Row, 11)) = 0 Then
                RSTTRXFILE!PUR_TAX = grdsales.TextMatrix(grdsales.Row, 9)
            Else
                RSTTRXFILE!PUR_TAX = Val(LBLNETCOST.Tag)
            End If
            RSTTRXFILE!MRP = Val(grdsales.TextMatrix(grdsales.Row, 5))
            RSTTRXFILE!PTR = Val(grdsales.TextMatrix(grdsales.Row, 6))
            RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(grdsales.Row, 7))
            RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(grdsales.Row, 21))
            RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(grdsales.Row, 22))
            RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(grdsales.Row, 24))
            RSTTRXFILE!Category = grdsales.TextMatrix(grdsales.Row, 25)
            If CMBDISTI.BoundText <> "" Or (lblsubdealer.Caption = "D" And Val(lblActAmt.Caption) <> 0) Then
                RSTTRXFILE!COM_FLAG = "Y"
                'RSTTRXFILE!AGENT_CODE = CMBDISTI.BoundText
                'RSTTRXFILE!AGENT_NAME = CMBDISTI.Text
            Else
                RSTTRXFILE!COM_FLAG = "N"
                'RSTTRXFILE!AGENT_CODE = ""
                'RSTTRXFILE!AGENT_NAME = ""
            End If
            RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(grdsales.Row, 26)
            RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(grdsales.Row, 27))
            RSTTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(grdsales.Row, 9))
            RSTTRXFILE!kfc_tax = Val(grdsales.TextMatrix(grdsales.Row, 47))
            RSTTRXFILE!UNIT = grdsales.TextMatrix(grdsales.Row, 4)
            If Trim(TxtBillName.text) = "" Then
                RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.text), 1, 30)
            Else
                RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(TxtBillName.text), 1, 30)
            End If
            RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(grdsales.Row, 10))
            RSTTRXFILE!ISSUE_QTY = 0
            RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(grdsales.Row, 17))
            RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(grdsales.Row, 18))
            Select Case grdsales.TextMatrix(grdsales.Row, 19)
                Case "DN"
                    RSTTRXFILE!CST = 1
                Case "CN"
                    RSTTRXFILE!CST = 2
                Case Else
                    RSTTRXFILE!CST = 0
            End Select
            RSTTRXFILE!BAL_QTY = 0
            RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(grdsales.Row, 12)
            RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(grdsales.Row, 8))
            RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(grdsales.Row, 7)) - Val(grdsales.TextMatrix(grdsales.Row, 6))) * Val(grdsales.TextMatrix(grdsales.Row, 3))
            RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(grdsales.Row, 20))
            RSTTRXFILE!MODIFY_DATE = Date
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE!C_USER_ID = "SM"
            RSTTRXFILE!M_USER_ID = DataList2.BoundText
            RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(grdsales.Row, 23))
            RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(grdsales.Row, 28) = "", 0, grdsales.TextMatrix(grdsales.Row, 28))
            RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(grdsales.Row, 29)
            RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(grdsales.Row, 30)
            RSTTRXFILE!ST_RATE = Val(grdsales.TextMatrix(grdsales.Row, 31))
            RSTTRXFILE!RETAILER_PRICE = Val(grdsales.TextMatrix(grdsales.Row, 39))
            RSTTRXFILE!CESS_PER = Val(grdsales.TextMatrix(grdsales.Row, 40))
            RSTTRXFILE!cess_amt = Val(grdsales.TextMatrix(grdsales.Row, 41))
            RSTTRXFILE!BARCODE = Trim(grdsales.TextMatrix(grdsales.Row, 42))
            RSTTRXFILE!ITEM_SPEC = Trim(grdsales.TextMatrix(grdsales.Row, 44))
            RSTTRXFILE!UN_BILL = Trim(grdsales.TextMatrix(grdsales.Row, 46))
            
            If Not IsDate(grdsales.TextMatrix(grdsales.Row, 38)) Then
                'RSTTRXFILE!EXP_DATE = Null
            Else
                RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(grdsales.Row, 38))) & "/" & Trim(grdsales.TextMatrix(grdsales.Row, 38))
            End If
            'RSTTRXFILE!EXP_DATE = Date
            If Trim(grdsales.TextMatrix(grdsales.Row, 33)) = "" Then
                RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(grdsales.Row, 2))
            Else
                RSTTRXFILE!PRINT_NAME = Trim(grdsales.TextMatrix(grdsales.Row, 33))
            End If
            RSTTRXFILE!GROSS_AMOUNT = Val(grdsales.TextMatrix(grdsales.Row, 34))
            RSTTRXFILE!DN_NO = Val(grdsales.TextMatrix(grdsales.Row, 35))
            If IsDate(grdsales.TextMatrix(grdsales.Row, 36)) Then
                RSTTRXFILE!DN_DATE = IIf(IsDate(grdsales.TextMatrix(grdsales.Row, 36)), Format(grdsales.TextMatrix(grdsales.Row, 36), "DD/MM/YYYY"), Null)
            End If
            'RSTTRXFILE!DN_LINENO = Val(grdsales.TextMatrix(grdsales.Row, 37))
            
        '    Dim RSTUNBILL As ADODB.Recordset
        '    Set RSTUNBILL = New ADODB.Recordset
        '    RSTUNBILL.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 13) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
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
            
SKIP_ITEMCHANGE:
            grdsales.Enabled = True
                    
        
            TxtItemName.Visible = False
            DataList3.Visible = False
            grdsales.SetFocus
        Case vbKeyEscape
            TxtItemName.SetFocus
    End Select
End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList3_GotFocus()
    flagchange2.Caption = 1
    TxtItemName.text = LBLitem.Caption
    DataList3.text = TxtItemName.text
    Call DataList3_Click
End Sub

Private Sub DataList3_LostFocus()
     flagchange2.Caption = ""
End Sub

Function FILL_BATCHGRID()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME, PRICE5, PRICE6, PRICE7 From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME, PRICE5, PRICE6, PRICE7 From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
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
    GRDPOPUP.Columns(21).Caption = "MRP"
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
    GRDPOPUP.Columns(21).Width = 1000
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
    GRDPOPUP.Columns(32).Width = 1000
    GRDPOPUP.Columns(33).Width = 1000
    GRDPOPUP.Columns(34).Width = 1000
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
    
    
End Function

Function FILL_BATCHGRID2()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME, PRICE5, PRICE6, PRICE7 From RTRXFILE  WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, P_CRTN, BARCODE, EXP_DATE, CESS_PER, CESS_AMT, TRX_YEAR, P_LWS, ITEM_CODE, ITEM_NAME, PRICE5, PRICE6, PRICE7 From RTRXFILE  WHERE BARCODE = '" & lblbarcode.Caption & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ASC", db, adOpenForwardOnly
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
    GRDPOPUP.Columns(21).Caption = "MRP"
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
    GRDPOPUP.Columns(21).Width = 1000
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
    GRDPOPUP.Columns(32).Width = 1000
    GRDPOPUP.Columns(33).Width = 1000
    GRDPOPUP.Columns(34).Width = 1000
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True

End Function

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
