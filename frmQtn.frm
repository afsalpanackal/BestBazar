VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMQTN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quotation"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12255
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   12255
   Begin VB.CommandButton CmdItemmast 
      Caption         =   "Create Item"
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
      Left            =   8850
      TabIndex        =   50
      Top             =   8085
      Width           =   1305
   End
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   5370
      Left            =   30
      TabIndex        =   44
      Top             =   1875
      Visible         =   0   'False
      Width           =   12195
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   5340
         Left            =   30
         TabIndex        =   45
         Top             =   30
         Width           =   12165
         _ExtentX        =   21458
         _ExtentY        =   9419
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   27
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
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   14010
      TabIndex        =   34
      Top             =   3090
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1482
      _Version        =   393216
   End
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   9315
      Left            =   -135
      TabIndex        =   12
      Top             =   0
      Width           =   13035
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00FFC0FF&
         Height          =   1980
         Left            =   150
         TabIndex        =   13
         Top             =   -75
         Width           =   12255
         Begin VB.OptionButton OptRetail 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Retail"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   8010
            TabIndex        =   60
            Top             =   1080
            Value           =   -1  'True
            Width           =   2115
         End
         Begin VB.OptionButton Optws 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Wholesale"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   8010
            TabIndex        =   59
            Top             =   330
            Width           =   2115
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
            Height          =   330
            Left            =   1065
            MaxLength       =   35
            TabIndex        =   46
            Top             =   720
            Width           =   4620
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
            Left            =   1050
            TabIndex        =   10
            Top             =   270
            Visible         =   0   'False
            Width           =   885
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   345
            Left            =   3015
            TabIndex        =   30
            Top             =   270
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
            Height          =   300
            Index           =   2
            Left            =   75
            TabIndex        =   52
            Top             =   1095
            Width           =   930
         End
         Begin MSForms.TextBox TxtBillAddress 
            Height          =   840
            Left            =   1065
            TabIndex        =   51
            Top             =   1080
            Width           =   4620
            VariousPropertyBits=   746604571
            MaxLength       =   100
            BorderStyle     =   1
            Size            =   "8149;1482"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
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
            Height          =   300
            Index           =   35
            Left            =   90
            TabIndex        =   47
            Top             =   735
            Width           =   930
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Qtn Date"
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
            Left            =   2040
            TabIndex        =   31
            Top             =   330
            Width           =   885
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Qtn NO."
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
            TabIndex        =   17
            Top             =   315
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
            Height          =   300
            Index           =   1
            Left            =   4470
            TabIndex        =   16
            Top             =   255
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
            Height          =   360
            Left            =   5640
            TabIndex        =   15
            Top             =   270
            Width           =   1350
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
            Left            =   1050
            TabIndex        =   14
            Top             =   270
            Width           =   885
         End
      End
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   465
         Left            =   7725
         TabIndex        =   29
         Top             =   9225
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
         Height          =   5430
         Left            =   150
         TabIndex        =   18
         Top             =   1830
         Width           =   12255
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
            Left            =   5475
            TabIndex        =   40
            Top             =   6330
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
            Left            =   7410
            TabIndex        =   39
            Top             =   5025
            Width           =   930
         End
         Begin VB.OptionButton OPTDISCPERCENT 
            BackColor       =   &H00FFC0FF&
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
            Left            =   5250
            TabIndex        =   38
            Top             =   5025
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton OptDiscAmt 
            BackColor       =   &H00FFC0FF&
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
            Left            =   6270
            TabIndex        =   37
            Top             =   5025
            Width           =   1125
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0FF&
            Height          =   4830
            Left            =   8625
            TabIndex        =   19
            Top             =   135
            Width           =   1635
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
               Height          =   435
               Left            =   90
               TabIndex        =   42
               Top             =   1815
               Width           =   1440
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
               ForeColor       =   &H00FF00FF&
               Height          =   375
               Index           =   4
               Left            =   60
               TabIndex        =   41
               Top             =   1560
               Width           =   1515
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
               ForeColor       =   &H00800080&
               Height          =   375
               Index           =   23
               Left            =   45
               TabIndex        =   33
               Top             =   810
               Width           =   1515
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
               Height          =   480
               Left            =   75
               TabIndex        =   32
               Top             =   1020
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
               ForeColor       =   &H00800080&
               Height          =   375
               Index           =   6
               Left            =   45
               TabIndex        =   21
               Top             =   120
               Width           =   1755
            End
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
               Height          =   480
               Left            =   90
               TabIndex        =   20
               Top             =   330
               Width           =   1440
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4785
            Left            =   60
            TabIndex        =   11
            Top             =   195
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   8440
            _Version        =   393216
            Rows            =   1
            Cols            =   8
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
            TabIndex        =   56
            Top             =   5055
            Width           =   645
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
            Left            =   1620
            TabIndex        =   55
            Top             =   5055
            Width           =   690
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
            Left            =   705
            TabIndex        =   54
            Top             =   5055
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
            Left            =   2310
            TabIndex        =   53
            Top             =   5055
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0FF&
         Height          =   2145
         Left            =   150
         TabIndex        =   22
         Top             =   7170
         Width           =   12270
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   345
            Left            =   10500
            TabIndex        =   61
            Top             =   555
            Width           =   1635
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
            Height          =   360
            Left            =   7275
            MaxLength       =   4
            TabIndex        =   48
            Top             =   375
            Width           =   885
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
            Height          =   360
            Left            =   6000
            MaxLength       =   6
            TabIndex        =   43
            Top             =   375
            Width           =   1245
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
            Left            =   5415
            TabIndex        =   4
            Top             =   915
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
            Left            =   60
            TabIndex        =   0
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
            Height          =   360
            Left            =   555
            TabIndex        =   1
            Top             =   375
            Width           =   4515
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
            Height          =   360
            Left            =   5100
            MaxLength       =   7
            TabIndex        =   2
            Top             =   375
            Width           =   870
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
            Left            =   6615
            TabIndex        =   7
            Top             =   1650
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
            Left            =   9015
            TabIndex        =   9
            Top             =   1650
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
            Left            =   5415
            TabIndex        =   6
            Top             =   1650
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
            Left            =   6615
            TabIndex        =   5
            Top             =   915
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
            TabIndex        =   23
            Top             =   2010
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
            Left            =   7815
            TabIndex        =   8
            Top             =   1650
            Width           =   1125
         End
         Begin MSForms.TextBox TxtSpec 
            Height          =   855
            Left            =   30
            TabIndex        =   58
            Top             =   1035
            Width           =   5040
            VariousPropertyBits=   746604571
            ForeColor       =   255
            MaxLength       =   99
            BorderStyle     =   1
            Size            =   "8890;1508"
            BorderColor     =   0
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            Caption         =   "Specifications"
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
            Left            =   45
            TabIndex        =   57
            Top             =   750
            Width           =   5025
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
            Height          =   300
            Index           =   13
            Left            =   7275
            TabIndex        =   49
            Top             =   150
            Width           =   885
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
            Height          =   300
            Index           =   30
            Left            =   6000
            TabIndex        =   36
            Top             =   150
            Width           =   1245
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
            Height          =   300
            Index           =   8
            Left            =   60
            TabIndex        =   28
            Top             =   150
            Width           =   480
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            Caption         =   " Item Description"
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
            Left            =   555
            TabIndex        =   27
            Top             =   150
            Width           =   4515
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
            Index           =   10
            Left            =   5100
            TabIndex        =   26
            Top             =   150
            Width           =   870
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
            Height          =   240
            Index           =   14
            Left            =   8190
            TabIndex        =   25
            Top             =   150
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
            Left            =   930
            TabIndex        =   24
            Top             =   2025
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
            Height          =   360
            Left            =   8190
            TabIndex        =   3
            Top             =   375
            Width           =   1530
         End
      End
   End
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   -15
      TabIndex        =   35
      Top             =   -225
      Width           =   915
   End
End
Attribute VB_Name = "FRMQTN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean

Dim CLOSEALL As Integer
Dim M_EDIT As Boolean
Dim N_EDIT As Boolean

Private Sub CmdItemmast_Click()
    'frmItemmast.Show
    'MDIMAIN.Enabled = False
End Sub

Private Sub CmdItemmast_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            If TxtSpec.Enabled = True Then TxtSpec.SetFocus
            If txtremarks.Enabled = True Then txtremarks.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub CmdPrint_Click()
    
    Call AppendSale
    
    MsgBox "Click OK", vbOKOnly, "Quotation"
    
    ReportNameVar = App.Path & "\RptQuotation.RPT"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "({POSUB.VCH_NO}= " & Val(txtBillNo.Text) & " )"
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.Count
        Report.Database.Tables(i).SetLogOnInfo "ConnectionName", "D:\dbase\INVSOFT.SQL", "admin", "!@#$%^&*())(*&^%$#@!"
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & txtremarks.Text & "'"
        If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
        'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
        If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
        If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
        If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
        'If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
        'If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & TxtPhone.Text & "'"
        'If CRXFormulaField.Name = "{@VCH_NO}" Then CRXFormulaField.Text = "'" & Trim(txtBillNo.Text) & "'"
        'If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
        If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = "'" & Format(Val(LBLDISCAMT.Caption), "0.00") & "'"
        'If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.Text = "'" & Trim(TxtCode.Text) & "'"
        If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.Text = " " & Val(txtOutstanding.Text) & " "
        If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
        
        'If CRXFormulaField.Name = "{@unit}" Then CRXFormulaField.Text = "'" & Trim(lblunit.Caption) & "'"
    Next
    frmreport.Caption = "BILL"
    Call GENERATEREPORT
End Sub

Private Sub CMDPRINT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTDISC.Text = ""
            TxtSpec.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtSpec.Enabled = False
            CMDMODIFY.Enabled = False
            cmddelete.Enabled = False
    End Select
End Sub

Private Sub Form_Activate()
    If txtremarks.Enabled = True Then txtremarks.SetFocus
    If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
    If TxtSpec.Enabled = True Then TxtSpec.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
    If CmdPrint.Enabled = True Then CmdPrint.SetFocus
    If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo eRRHAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(VCH_NO)) From QTNMAST ", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    N_EDIT = False
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 500
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 3300
    grdsales.ColWidth(3) = 900
    grdsales.ColWidth(4) = 1200
    grdsales.ColWidth(5) = 800
    grdsales.ColWidth(6) = 1400
    grdsales.ColWidth(7) = 6400
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM DESCRIPTION"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "RATE"
    grdsales.TextArray(5) = "DISC %"
    grdsales.TextArray(6) = "SUB TOTAL"
    grdsales.TextArray(7) = "Specifications"
    
    LBLTOTAL.Caption = 0
    
    PHYFLAG = True
    ITEM_FLAG = True
    
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTDISC.Enabled = False
    TxtSpec.Enabled = False
    cmddelete.Enabled = False
    CMDMODIFY.Enabled = False
    CmdPrint.Enabled = False
    TXTSLNO.Text = 1
    TXTSLNO.Enabled = False
    CLOSEALL = 1
    M_EDIT = False
'    Me.Width = 11700
'    Me.Height = 10185
    Me.Left = 0
    Me.Top = 0
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If ITEM_FLAG = False Then PHY_ITEM.Close
        'MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        'MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub TxtBillAddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtremarks.Text) = "" Then
                MsgBox "Enter the name for the customer", , "Quotation"
                txtremarks.SetFocus
                Exit Sub
            End If
            FRMEHEAD.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
    End Select
End Sub

Private Sub TxtBillAddress_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TRXMAST As ADODB.Recordset
    Dim TRXSUB As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    
    Dim i As Integer
    Dim N As Integer
    Dim M As Integer

    On Error GoTo eRRHAND

    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtBillNo.Text) = 0 Then Exit Sub
            grdsales.Rows = 1
             grdsales.FixedRows = 0
            i = 0
            Set TRXSUB = New ADODB.Recordset
            TRXSUB.Open "Select * From QTNSUB WHERE VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until TRXSUB.EOF
                Set TRXFILE = New ADODB.Recordset
                TRXFILE.Open "Select * From QTNSUB WHERE VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(TRXSUB!LINE_NO) & "", db, adOpenStatic, adLockReadOnly
                If Not (TRXFILE.EOF And TRXFILE.BOF) Then
                    i = i + 1
                    grdsales.Rows = grdsales.Rows + 1
                    grdsales.FixedRows = 1
                    grdsales.TextMatrix(i, 0) = i
                    grdsales.TextMatrix(i, 1) = TRXFILE!ITEM_CODE
                    grdsales.TextMatrix(i, 2) = TRXFILE!ITEM_NAME
                    grdsales.TextMatrix(i, 3) = TRXFILE!QTY
                    grdsales.TextMatrix(i, 4) = Format(TRXFILE!SALES_PRICE, ".000")
                    grdsales.TextMatrix(i, 5) = Format(TRXFILE!LINE_DISC, ".00") 'DISC
                    grdsales.TextMatrix(i, 6) = Format(Val(TRXFILE!TRX_TOTAL), ".000")
                    grdsales.TextMatrix(i, 7) = IIf(IsNull(TRXFILE!SPEC), "", TRXFILE!SPEC)
                    'txtBillNo.Text = ""
                    'LBLBILLNO.Caption = ""
                End If
                TRXFILE.Close
                Set TRXFILE = Nothing
                TRXSUB.MoveNext
            Loop
            TRXSUB.Close
            Set TRXSUB = Nothing
            
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "Select * From QTNMAST  WHERE VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                TXTINVDATE.Text = Format(TRXMAST!VCH_DATE, "DD/MM/YYYY")
                LBLDATE.Caption = Format(TRXMAST!CREATE_DATE, "DD/MM/YYYY")
                txtremarks.Text = IIf(IsNull(TRXMAST!ACT_NAME), "", TRXMAST!ACT_NAME)
                TxtBillAddress.Text = IIf(IsNull(TRXMAST!ADDRESS), "", TRXMAST!ADDRESS)
                If TRXMAST!SLSM_CODE = "A" Then
                    TXTTOTALDISC.Text = IIf(IsNull(TRXMAST!DISCOUNT), "", TRXMAST!DISCOUNT)
                    OptDiscAmt.value = True
                ElseIf TRXMAST!SLSM_CODE = "P" Then
                    TXTTOTALDISC.Text = IIf(IsNull(TRXMAST!DISCOUNT), "", Round((TRXMAST!DISCOUNT * 100 / TRXMAST!VCH_AMOUNT), 2))
                    OPTDISCPERCENT.value = True
                End If
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            LBLBILLNO.Caption = Val(txtBillNo.Text)
            
            
            LBLTOTAL.Caption = ""
            lblnetamount.Caption = ""
            For i = 1 To grdsales.Rows - 1
                grdsales.TextMatrix(i, 0) = i
                LBLTOTAL.Caption = Format(Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 6)), 2), "0.00")
            Next i
            LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
            TXTAMOUNT.Text = ""
            If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
                TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
            ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
                TXTAMOUNT.Text = Round((Val(LBLTOTAL.Caption) * Val(TXTTOTALDISC.Text) / 100), 2)
            End If
            LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
            lblnetamount.Caption = Format(Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2), "0.00")
            
            
            TXTSLNO.Text = grdsales.Rows
            txtBillNo.Visible = False
            TXTSLNO.Enabled = True
            
            If grdsales.Rows > 1 Then
                TXTSLNO.SetFocus
            Else
                TXTSLNO.Enabled = False
                TXTINVDATE.SetFocus
            End If
    
    End Select
    N_EDIT = False
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTBILLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBillNo_LostFocus()
    Dim TRXMAST As ADODB.Recordset
    Dim i As Integer

    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(Val(VCH_NO)) From QTNMAST ", db, adOpenStatic, adLockReadOnly
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
    TRXMAST.Open "Select MIN(Val(VCH_NO)) From QTNSUB ", db, adOpenStatic, adLockReadOnly
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

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.Text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.Text = "  /  /    " Then
                TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
                txtremarks.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                txtremarks.SetFocus
            End If
        Case vbKeyEscape
            If M_EDIT = False Then
                txtBillNo.Visible = True
                txtBillNo.SetFocus
            End If
    End Select
End Sub

Private Sub TXTINVDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTREMARKS_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtremarks.Text) = "" Then
                MsgBox "Enter the name for the customer", , "Quotation"
                txtremarks.SetFocus
                Exit Sub
            End If
            TxtBillAddress.SetFocus
        Case vbKeyEscape
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
    End Select
End Sub

Private Sub CMBDISTI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTREMARKS_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.Text) = 0 Then
                TXTSLNO.Text = ""
                TXTPRODUCT.Text = ""
                TXTQTY.Text = ""
                TXTDISC.Text = ""
                TxtSpec.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTITEMCODE.Text = ""
                TXTSLNO.Text = grdsales.Rows
                cmddelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.Text) >= grdsales.Rows Then
                TXTSLNO.Text = grdsales.Rows
                cmddelete.Enabled = False
                CMDMODIFY.Enabled = False
            End If
            If Val(TXTSLNO.Text) < grdsales.Rows Then
                TXTSLNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 0)
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 1)
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 3)
                TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 6), ".000")
                TxtSpec.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                TXTRETAILNOTAX.Enabled = False
                TXTDISC.Enabled = False
                TxtSpec.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                cmddelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTDISC.Enabled = False
            TxtSpec.Enabled = False
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If cmddelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TXTQTY.Text = ""
                TXTRETAILNOTAX.Text = ""
                TXTDISC.Text = ""
                TxtSpec.Text = ""
                LBLSUBTOTAL.Caption = ""
                cmdadd.Enabled = False
                cmddelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.Rows > 1 Then
                TXTSLNO.Enabled = False
                CmdPrint.Enabled = True
                cmdRefresh.Enabled = True
                CmdPrint.SetFocus
            Else
                TXTSLNO.Enabled = False
                FRMEHEAD.Enabled = True
                txtremarks.Enabled = True
                txtremarks.SetFocus
            End If
    End Select
End Sub

Private Sub TXTSLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case vbKeyTab
            KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    On Error GoTo eRRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            cmddelete.Enabled = False
            TXTQTY.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTDISC.Text = ""
            TxtSpec.Text = ""
            LBLSUBTOTAL.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
           
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY], [P_RETAIL], [P_WS], [P_VAN], [ITEM_COST] From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY], [P_RETAIL], [P_WS], [P_VAN], [ITEM_COST] From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            Set grdtmp.DataSource = PHY
            
            If PHY.RecordCount = 0 Then
                If MsgBox("Item not found!!!! Do yo want to add this item", vbYesNo, "Quotation Entry..") = vbNo Then
                    Exit Sub
                Else
                    TXTPRODUCT.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                    Exit Sub
                End If
            End If
            
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                If OptRetail.value = True Then
                    TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(3)), "", Format(grdtmp.Columns(3), "0.00"))
                Else
                    TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(4)), "", Format(grdtmp.Columns(4), "0.00"))
                End If
                lblretail.Caption = IIf(IsNull(grdtmp.Columns(3)), "", Format(grdtmp.Columns(3), "0.00"))
                lblwsale.Caption = IIf(IsNull(grdtmp.Columns(4)), "", Format(grdtmp.Columns(4), "0.00"))
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "Quotation Entry..") = vbNo Then Exit Sub
                    End If
                Next i
                
                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
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
            
JUMPNONSTOCK:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTDISC.Enabled = False
            cmddelete.Enabled = False
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTDISC.Enabled = False
            TXTSLNO.SetFocus
            cmddelete.Enabled = False
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTQTY_GotFocus()
    
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
    TXTQTY.Tag = Trim(TXTPRODUCT.Text)
    
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTQTY.Text) = 0 Then Exit Sub
            TXTQTY.Enabled = False
            TXTRETAILNOTAX.Enabled = True
            TXTRETAILNOTAX.SetFocus
         Case vbKeyEscape
            If N_EDIT = True Then Exit Sub
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    TXTQTY.Text = Format(TXTQTY.Text, ".000")
    TXTDISC.Tag = 0
    TXTDISC.Tag = Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
    If Val(TXTQTY.Text) = 0 Then
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Text) - Val(TXTDISC.Tag), 3), "0.00")
    Else
        LBLSUBTOTAL.Caption = Format(Round((Val(TXTRETAILNOTAX.Text) * Val(TXTQTY.Text)) - (Val(TXTDISC.Tag) * Val(TXTQTY.Text)), 3), "0.00")
    End If
End Sub

Private Sub TXTRETAILNOTAX_GotFocus()
    TXTRETAILNOTAX.SelStart = 0
    TXTRETAILNOTAX.SelLength = Len(TXTRETAILNOTAX.Text)
End Sub

Private Sub TXTRETAILNOTAX_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 And Val(TXTRETAILNOTAX.Text) = 0 Then Exit Sub
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyEscape
            TXTRETAILNOTAX.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub TXTRETAILNOTAX_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub TXTRETAILNOTAX_LostFocus()
    TXTDISC.Tag = 0
    TXTDISC.Tag = Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
    If Val(TXTQTY.Text) = 0 Then
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Text) - Val(TXTDISC.Tag), 3), "0.00")
    Else
        LBLSUBTOTAL.Caption = Format(Round((Val(TXTRETAILNOTAX.Text) * Val(TXTQTY.Text)) - (Val(TXTDISC.Tag) * Val(TXTQTY.Text)), 3), "0.00")
    End If
    TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.000")
End Sub

Private Sub TXTDISC_GotFocus()
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.Text)
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtSpec.Enabled = True
            TXTDISC.Enabled = False
            TxtSpec.SetFocus
        Case vbKeyEscape
            TXTDISC.Enabled = False
            TXTRETAILNOTAX.Enabled = True
            TXTRETAILNOTAX.SetFocus
    End Select
End Sub

Private Sub TXTDISC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDISC_LostFocus()
    TXTDISC.Tag = 0
    TXTDISC.Tag = Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
    If Val(TXTQTY.Text) = 0 Then
        LBLSUBTOTAL.Caption = Format(Round(Val(TXTRETAILNOTAX.Text) - Val(TXTDISC.Tag), 3), "0.00")
    Else
        LBLSUBTOTAL.Caption = Format(Round((Val(TXTRETAILNOTAX.Text) * Val(TXTQTY.Text)) - (Val(TXTDISC.Tag) * Val(TXTQTY.Text)), 3), "0.00")
    End If
End Sub

Private Sub CMDMODIFY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    On Error GoTo eRRHAND
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            '!ISSUE_QTY = !ISSUE_QTY - ((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))))
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    CMDMODIFY.Enabled = False
    cmddelete.Enabled = False
    cmdexit.Enabled = False
    M_EDIT = True
    N_EDIT = True
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTDISC.Text = ""
            TxtSpec.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtSpec.Enabled = False
            CMDMODIFY.Enabled = False
            cmddelete.Enabled = False
    End Select
End Sub

Private Sub CmdDelete_Click()
    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    For i = Val(TXTSLNO.Text) - 1 To grdsales.Rows - 2
        grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = i
        grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(i + 1, 6)
    Next i
    grdsales.Rows = grdsales.Rows - 1
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 6)), 2)
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Format(Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2), "0.00")
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTQTY.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTDISC.Text = ""
    TxtSpec.Text = ""
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    cmddelete.Enabled = False
    CMDMODIFY.Enabled = False
    cmdexit.Enabled = False
    M_EDIT = True
    If grdsales.Rows = 1 Then
'        CMDEXIT.Enabled = True
        CmdPrint.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    
End Sub

Private Sub CMDADD_Click()
    Dim i As Integer
    
    On Error GoTo eRRHAND
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Format(Val(TXTRETAILNOTAX.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Format(Val(TXTDISC.Text), "0.00")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Trim(TxtSpec.Text)
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 6)), 2)
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    
    
    Call TXTTOTALDISC_LostFocus
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TXTQTY.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTDISC.Text = ""
    TxtSpec.Text = ""
    LBLSUBTOTAL.Caption = ""
    lblretail.Caption = ""
    lblwsale.Caption = ""
    cmdadd.Enabled = False
    cmddelete.Enabled = False
    cmdexit.Enabled = False
    TXTSLNO.Enabled = True
    M_EDIT = True
    N_EDIT = False
    TXTSLNO.SetFocus
    If grdsales.Rows >= 15 Then grdsales.TopRow = grdsales.Rows - 1
Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            TxtSpec.Enabled = True
            TxtSpec.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub cmdRefresh_Click()
    
   ' If grdsales.Rows = 1 Then GoTo SKIP
    
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
    Call AppendSale
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(VCH_NO)) From QTNMAST ", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    lblnetamount.Caption = ""
    LBLDATE.Caption = Date
    LBLTOTAL.Caption = ""
    TXTTOTALDISC.Text = ""
    TXTAMOUNT.Text = ""
    LBLDISCAMT.Caption = ""
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    M_EDIT = False
    cmdRefresh.Enabled = False
    cmdexit.Enabled = True
    CmdPrint.Enabled = False
    cmdexit.Enabled = True
    TXTSLNO.Enabled = False
    FRMEHEAD.Enabled = True
    txtremarks.Enabled = True
    txtremarks.SetFocus
    TXTQTY.Tag = ""
    txtremarks.Text = ""
    TxtBillAddress.Text = ""
    'Me.Enabled = False
    'FRMDEBIT.Show
    
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTDISC.Text = ""
            TxtSpec.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            TxtSpec.Enabled = False
            CMDMODIFY.Enabled = False
            cmddelete.Enabled = False
    End Select
End Sub

Function FILL_ITEMGRID()
    FrmeMain.Enabled = False
    FRMEITEM.Visible = True
    Set GRDPOPUPITEM.DataSource = Nothing
    
    If ITEM_FLAG = True Then
        PHY_ITEM.Open "Select [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY], [P_RETAIL], [P_WS], [MRP], [ITEM_COST] From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%' ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY], [P_RETAIL], [P_WS], [MRP], [ITEM_COST] From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%' ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    End If

    Set GRDPOPUPITEM.DataSource = PHY_ITEM
    'GRDPOPUPITEM.RowHeight = 350
    GRDPOPUPITEM.Columns(0).Visible = False
    GRDPOPUPITEM.Columns(1).Caption = "ITEM NAME"
    GRDPOPUPITEM.Columns(1).Width = 4600
    GRDPOPUPITEM.Columns(2).Caption = "QTY"
    GRDPOPUPITEM.Columns(2).Width = 1000
    GRDPOPUPITEM.Columns(3).Caption = "RT"
    GRDPOPUPITEM.Columns(3).Width = 1100
    GRDPOPUPITEM.Columns(4).Caption = "WS"
    GRDPOPUPITEM.Columns(4).Width = 1100
    GRDPOPUPITEM.Columns(5).Caption = "MRP"
    GRDPOPUPITEM.Columns(5).Width = 1100
    GRDPOPUPITEM.Columns(6).Caption = "COST"
    GRDPOPUPITEM.Columns(6).Width = 1100
    GRDPOPUPITEM.SetFocus

End Function

Private Function AppendSale()
    Dim RSTTRXFILE As ADODB.Recordset
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
    On Error GoTo eRRHAND
    
    db.Execute "delete * From QTNMAST  WHERE VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete * From QTNSUB  WHERE VCH_NO = " & Val(txtBillNo.Text) & ""
    'DB.Execute "delete * From P_Rate  AND VCH_NO = " & Val(txtBillNo.Text) & ""
    
    E_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'    If Day(E_DATE) <= 12 Then
'        DAY_DATE = Format(Month(E_DATE), "00")
'        MONTH_DATE = Format(Day(E_DATE), "00")
'        YEAR_DATE = Format(Year(E_DATE), "0000")
'        E_DATE = DAY_DATE & "/" & MONTH_DATE & "/" & YEAR_DATE
'    End If
'    E_DATE = Format(E_DATE, "MM/DD/YYYY")
    
    TRXVALUE = 0
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From QTNSUB WHERE VCH_DATE = # " & E_DATE & " # ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        TRXVALUE = TRXVALUE + RSTTRXFILE!TRX_TOTAL
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From QTNMAST WHERE VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    Else
        RSTTRXFILE!VCH_AMOUNT = RSTTRXFILE!VCH_AMOUNT + Val(LBLTOTAL.Caption)
    End If
        
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = "10111"
    RSTTRXFILE!ACT_NAME = Trim(txtremarks.Text)
    RSTTRXFILE!ADDRESS = Trim(TxtBillAddress.Text)
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!PAY_AMOUNT = 0
    RSTTRXFILE!REF_NO = ""
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "A"
        RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "P"
        RSTTRXFILE!DISCOUNT = Round(RSTTRXFILE!VCH_AMOUNT * Val(TXTTOTALDISC.Text) / 100, 2)
    End If
    RSTTRXFILE!CHECK_FLAG = "I"
    RSTTRXFILE!POST_FLAG = "N"
    RSTTRXFILE!Remarks = Trim(txtremarks.Text)
    RSTTRXFILE!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From QTNSUB", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "Select * From ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTITEMMAST.BOF Or RSTITEMMAST.EOF) Then
            RSTITEMMAST!SALES_PRICE = Val(grdsales.TextMatrix(i, 4))
            RSTITEMMAST.Update
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = ""
        RSTTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(i, 1))
        RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 4))
        RSTTRXFILE!VCH_DESC = "Issued to     " & Trim(txtremarks.Text)
        RSTTRXFILE!SPEC = Trim(grdsales.TextMatrix(i, 7))
        
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 5))
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 6)
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "AM"
        
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Exit Function
eRRHAND:
    MsgBox Err.Description
End Function

Private Sub GRDPOPUPITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTtax As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            If OptRetail.value = True Then
                TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(3)), "", Format(GRDPOPUPITEM.Columns(3), "0.00"))
            Else
                TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUPITEM.Columns(4)), "", Format(GRDPOPUPITEM.Columns(4), "0.00"))
            End If
            lblretail.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(3)), "", Format(GRDPOPUPITEM.Columns(3), "0.00"))
            lblwsale.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(4)), "", Format(GRDPOPUPITEM.Columns(4), "0.00"))
            i = 0
            For i = 1 To grdsales.Rows - 1
                If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists.... Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then
                        Set GRDPOPUPITEM.DataSource = Nothing
                        FRMEITEM.Visible = False
                        FrmeMain.Enabled = True
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
            FrmeMain.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
            Exit Sub
        Case vbKeyEscape
            TXTQTY.Text = ""
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FrmeMain.Enabled = True
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TxtSpec_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.Enabled = True
            TxtSpec.Enabled = False
            'cmdadd.SetFocus
        Case vbKeyEscape
            TxtSpec.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
    End Select
End Sub

Private Sub TxtSpec_KeyPress(KeyAscii As MSForms.ReturnInteger)
     Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTTOTALDISC_GotFocus()
    TXTTOTALDISC.SelStart = 0
    TXTTOTALDISC.SelLength = Len(TXTTOTALDISC.Text)
End Sub

Private Sub TXTTOTALDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            If TxtSpec.Enabled = True Then TxtSpec.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
        End Select
End Sub

Private Sub TXTTOTALDISC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTTOTALDISC_LostFocus()
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        lblnetamount.Caption = Val(lblnetamount.Caption) + Val(grdsales.TextMatrix(i, 6))
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = 0
    If OptDiscAmt.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round((Val(LBLTOTAL.Caption) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2)
    
End Sub

Private Sub TxtSpec_GotFocus()
    TxtSpec.SelStart = 0
    TxtSpec.SelLength = Len(TxtSpec.Text)
End Sub
