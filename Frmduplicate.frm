VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMdupsale 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " DUMMY BILL"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   ControlBox      =   0   'False
   Icon            =   "Frmduplicate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11355
   Begin VB.Timer TimerPRINT 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5115
      Top             =   7650
   End
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   8040
      Top             =   7230
   End
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3120
      Left            =   1950
      TabIndex        =   55
      Top             =   1545
      Visible         =   0   'False
      Width           =   6030
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   2835
         Left            =   120
         TabIndex        =   56
         Top             =   180
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5001
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
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   2040
      TabIndex        =   51
      Top             =   1575
      Visible         =   0   'False
      Width           =   5835
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2535
         Left            =   90
         TabIndex        =   54
         Top             =   360
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   4471
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
      Begin VB.Label lblhead 
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
         TabIndex        =   53
         Top             =   105
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.Label lblhead 
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
         TabIndex        =   52
         Top             =   105
         Visible         =   0   'False
         Width           =   2610
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   7830
      Left            =   -90
      TabIndex        =   21
      Top             =   -240
      Width           =   11295
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Height          =   690
         Left            =   210
         TabIndex        =   22
         Top             =   180
         Width           =   11085
         Begin VB.ComboBox TXTPATIENT 
            Height          =   315
            Left            =   4980
            TabIndex        =   3
            Top             =   225
            Width           =   2070
         End
         Begin VB.TextBox txttime 
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
            Left            =   3195
            TabIndex        =   2
            Top             =   225
            Width           =   885
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
            Left            =   420
            TabIndex        =   0
            Top             =   225
            Width           =   885
         End
         Begin MSMask.MaskEdBox txtdate 
            Height          =   330
            Left            =   1920
            TabIndex        =   1
            Top             =   225
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
         Begin MSForms.ComboBox TXTDOCTOR 
            Height          =   360
            Left            =   7845
            TabIndex        =   59
            Top             =   195
            Width           =   3195
            VariousPropertyBits=   746604571
            ForeColor       =   255
            MaxLength       =   35
            DisplayStyle    =   3
            Size            =   "5636;635"
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
            Left            =   75
            TabIndex        =   26
            Top             =   240
            Width           =   390
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
            Height          =   300
            Index           =   1
            Left            =   1350
            TabIndex        =   25
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "PATIENT"
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
            Left            =   4140
            TabIndex        =   24
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DOCTOR"
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
            Index           =   3
            Left            =   7065
            TabIndex        =   23
            Top             =   255
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   4680
         Left            =   195
         TabIndex        =   27
         Top             =   795
         Width           =   11100
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Height          =   4455
            Left            =   9015
            TabIndex        =   28
            Top             =   180
            Width           =   2040
            Begin VB.Label lblhead 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "DUMMY BILL"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   1110
               Index           =   1
               Left            =   180
               TabIndex        =   58
               Top             =   1740
               Width           =   1650
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL AMOUNT"
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
               Height          =   375
               Index           =   6
               Left            =   210
               TabIndex        =   30
               Top             =   255
               Width           =   1590
            End
            Begin VB.Label LBLTOTAL 
               Alignment       =   2  'Center
               BackColor       =   &H00800000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label2"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   15
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   540
               Left            =   255
               TabIndex        =   29
               Top             =   675
               Width           =   1395
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4335
            Left            =   105
            TabIndex        =   4
            Top             =   270
            Width           =   8850
            _ExtentX        =   15610
            _ExtentY        =   7646
            _Version        =   393216
            Rows            =   1
            Cols            =   18
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            GridLineWidth   =   2
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Height          =   1575
         Left            =   195
         TabIndex        =   31
         Top             =   5415
         Width           =   11100
         Begin VB.CommandButton CMDREFRESH 
            Caption         =   "&REFRESH"
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
            Left            =   8610
            TabIndex        =   19
            Top             =   810
            Width           =   1125
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
            Height          =   480
            Left            =   3915
            TabIndex        =   15
            Top             =   795
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
            Height          =   285
            Left            =   120
            TabIndex        =   5
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
            Left            =   705
            TabIndex        =   6
            Top             =   435
            Width           =   4440
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
            Left            =   5175
            MaxLength       =   7
            TabIndex        =   7
            Top             =   450
            Width           =   675
         End
         Begin VB.TextBox TXTRATE 
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
            Left            =   5880
            MaxLength       =   6
            TabIndex        =   8
            Top             =   450
            Width           =   630
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
            Height          =   300
            Left            =   6540
            MaxLength       =   4
            TabIndex        =   9
            Top             =   450
            Width           =   600
         End
         Begin VB.TextBox TXTDISC 
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
            Left            =   9285
            MaxLength       =   4
            TabIndex        =   13
            Top             =   450
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
            Height          =   465
            Left            =   7455
            TabIndex        =   18
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
            Left            =   9810
            TabIndex        =   20
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
            Left            =   6270
            TabIndex        =   17
            Top             =   810
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
            Height          =   480
            Left            =   5100
            TabIndex        =   16
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
            Left            =   1230
            TabIndex        =   36
            Top             =   825
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
            Height          =   285
            Left            =   8310
            MaxLength       =   15
            TabIndex        =   12
            Top             =   450
            Width           =   930
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
            Left            =   1215
            TabIndex        =   35
            Top             =   1200
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
            Left            =   3075
            TabIndex        =   34
            Top             =   1200
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
            Left            =   4905
            TabIndex        =   33
            Top             =   1230
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
            Left            =   6780
            TabIndex        =   32
            Top             =   1275
            Visible         =   0   'False
            Width           =   690
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   285
            Left            =   7185
            TabIndex        =   10
            Top             =   465
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   503
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
         Begin VB.TextBox txtexpdate 
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
            Left            =   7185
            MaxLength       =   10
            TabIndex        =   11
            Top             =   450
            Width           =   1080
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
            Left            =   120
            TabIndex        =   50
            Top             =   225
            Width           =   570
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
            Left            =   705
            TabIndex        =   49
            Top             =   225
            Width           =   4440
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
            Left            =   5175
            TabIndex        =   48
            Top             =   225
            Width           =   675
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
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   11
            Left            =   5880
            TabIndex        =   47
            Top             =   225
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Tax %"
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
            Left            =   6540
            TabIndex        =   46
            Top             =   225
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Disc %"
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
            Left            =   9285
            TabIndex        =   45
            Top             =   240
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
            Left            =   9990
            TabIndex        =   44
            Top             =   240
            Width           =   930
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
            Left            =   90
            TabIndex        =   43
            Top             =   840
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
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   16
            Left            =   7170
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Batch"
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
            Left            =   8310
            TabIndex        =   41
            Top             =   240
            Width           =   930
         End
         Begin VB.Label LBLSUBTOTAL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   9990
            TabIndex        =   14
            Top             =   450
            Width           =   930
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
            Left            =   75
            TabIndex        =   40
            Top             =   1215
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
            Left            =   3765
            TabIndex        =   39
            Top             =   1245
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
            Left            =   5640
            TabIndex        =   38
            Top             =   1290
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
            Left            =   1935
            TabIndex        =   37
            Top             =   1215
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
   Begin MSDataGridLib.DataGrid grdtmp 
      Height          =   465
      Left            =   11070
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
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
End
Attribute VB_Name = "FRMdupsale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim TMPREC As New ADODB.Recordset
Dim TMPFLAG As Boolean

Dim PHY_BATCH As New ADODB.Recordset
Dim BATCH_FLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean

Dim M_EDIT As Boolean


Private Sub CMDADD_Click()
    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Format(Val(TXTRATE.Text) * Val(TXTUNIT.Text), ".00")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTRATE.Text), ".00")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Val(TXTDISC.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(TXTTAX.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = Trim(txtexpdate.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Format(Val(LBLSUBTOTAL.Caption), ".00")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTTRXTYPE.Text)
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = Trim(RSTTRXFILE!MANUFACTURER)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    LBLTOTAL.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Format(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 11)), ".00")
    Next i
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtBatch.Text = ""
    txtexpdate.Text = ""
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    TXTPATIENT.Enabled = True
    TXTDOCTOR.Enabled = True
    TXTSLNO.Enabled = True
    M_EDIT = False
    TXTSLNO.SetFocus
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    
    For i = Val(TXTSLNO.Text) - 1 To grdsales.Rows - 2
        grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = i
        grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(i + 1, 6)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = grdsales.TextMatrix(i + 1, 7)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = grdsales.TextMatrix(i + 1, 8)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = grdsales.TextMatrix(i + 1, 9)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = grdsales.TextMatrix(i + 1, 10)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = grdsales.TextMatrix(i + 1, 11)
    Next i
    grdsales.Rows = grdsales.Rows - 1
    LBLTOTAL.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 11))
    Next i
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtexpdate.Text = ""
    txtBatch.Text = ""
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    TXTPATIENT.Enabled = True
    TXTDOCTOR.Enabled = True
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    M_EDIT = False
    If grdsales.Rows = 1 Then CMDPRINT.Enabled = False
    
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    M_EDIT = True
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
            TXTRATE.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            txtexpdate.Text = ""
            txtBatch.Text = ""
            
            TXTPATIENT.Enabled = True
            TXTDOCTOR.Enabled = True
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub cmdprint_Click()
    If grdsales.Rows = 1 Then Exit Sub
    
    If Not IsDate(txtdate.Text) Then
        MsgBox "ENTER PROPER DATE", vbOKOnly, "BILL.."
        txtdate.SetFocus
        Exit Sub
    End If
    
    If Weekday(txtdate.Text) = 1 Then
        MsgBox "THE ENTERED DATE IS SUNDAY", vbOKOnly, "BILL.."
        txtdate.Enabled = True
        txtdate.SetFocus
        Exit Sub
    End If
    
    Call cmdReportGenerate_Click
    
    Open App.Path & "\repo.bat" For Output As #1 '//Creating Batch file
    
    Print #1, "TYPE " & App.Path & "\Report.txt > PRN"
    Print #1, "EXIT"
    Close #1
    
    '//HERE write the proper path where your command.com file exist
    'Shell "C:\WINDOWS\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
    Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
    
    TimerPRINT.Enabled = True
    
    TXTPATIENT.Enabled = True
    TXTDOCTOR.Enabled = True
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTTAX.Enabled = False
    txtexpdate.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
End Sub

Private Sub cmdRefresh_Click()
    Dim rstrxduplicate As ADODB.Recordset
    
    db2.Execute ("DELETE * FROM DUMMYBILL WHERE PATIENT = '" & TXTPATIENT.Text & "'")
    Set rstrxduplicate = New ADODB.Recordset
    rstrxduplicate.Open "SELECT *  FROM DUMMYBILL", db2, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        rstrxduplicate.AddNew
        rstrxduplicate!PATIENT = TXTPATIENT.Text
        rstrxduplicate!DOCTOR = TXTDOCTOR.Text
        rstrxduplicate!LINE_NO = i
        rstrxduplicate!ITEM_NAME = (grdsales.TextMatrix(i, 2))
        rstrxduplicate!ITEM_CODE = (grdsales.TextMatrix(i, 12))
        rstrxduplicate!MFGR = (grdsales.TextMatrix(i, 17))
        rstrxduplicate!QTY = (grdsales.TextMatrix(i, 3))
        rstrxduplicate!MRP = (grdsales.TextMatrix(i, 5))
        rstrxduplicate!SALES_PRICE = (grdsales.TextMatrix(i, 6))
        rstrxduplicate!REF_NO = (grdsales.TextMatrix(i, 9))
        rstrxduplicate!TRX_TOTAL = Val(grdsales.TextMatrix(i, 11))
        rstrxduplicate!EXP_DATE = (grdsales.TextMatrix(i, 10))
        rstrxduplicate.Update
    Next i
    rstrxduplicate.Close
    Set rstrxduplicate = Nothing
    TXTSLNO.Enabled = False
    txtdate.Enabled = False
    TXTPATIENT.Enabled = False
    TXTDOCTOR.Enabled = False
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTTAX.Enabled = False
    txtexpdate.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    CMDREFRESH.Enabled = False
    TXTSLNO.Text = 1
    grdsales.Rows = 1
    TXTPRODUCT.Text = ""
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    TXTTAX.Text = ""
    txtexpdate.Text = ""
    TXTEXPIRY.Text = "  /  "
    txtBatch.Text = ""
    TXTDISC.Text = ""
    LBLSUBTOTAL.Caption = ""
    txttime.Text = Time
    txtBillNo.Enabled = True
    txtBillNo.SetFocus
    txtBillNo.Text = ""
    TXTPATIENT.Text = ""
    TXTDOCTOR.Text = ""
    txtdate.Text = "  /  /    "
End Sub

Private Sub Form_Load()
    txttime.Text = Time
    grdsales.ColWidth(0) = 400
    grdsales.ColAlignment(0) = 1
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2000
    grdsales.ColAlignment(2) = 1
    grdsales.ColWidth(3) = 500
    grdsales.ColAlignment(3) = 1
    grdsales.ColWidth(4) = 500
    grdsales.ColAlignment(4) = 1
    grdsales.ColWidth(5) = 700
    grdsales.ColAlignment(5) = 1
    grdsales.ColWidth(6) = 600
    grdsales.ColAlignment(6) = 1
    grdsales.ColWidth(7) = 500
    grdsales.ColAlignment(7) = 1
    grdsales.ColWidth(8) = 600
    grdsales.ColAlignment(8) = 1
    grdsales.ColWidth(9) = 800
    grdsales.ColAlignment(9) = 1
    grdsales.ColWidth(10) = 1100
    grdsales.ColAlignment(10) = 1
    grdsales.ColWidth(11) = 1000
    grdsales.ColAlignment(11) = 1
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "UNIT"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "RATE"
    grdsales.TextArray(7) = "DISC %"
    grdsales.TextArray(8) = "TAX %"
    grdsales.TextArray(9) = "BATCH"
    grdsales.TextArray(10) = "EXPIRY"
    grdsales.TextArray(11) = "SUB TOTAL"
    grdsales.TextArray(12) = "ITEM CODE"
    grdsales.TextArray(13) = "Vch No"
    grdsales.TextArray(14) = "Line No"
    grdsales.TextArray(15) = "Trx Type"
    grdsales.TextArray(16) = "Non Stock"
    grdsales.TextArray(17) = "MFGR"
    
    grdsales.ColWidth(12) = 0
    grdsales.ColWidth(13) = 0
    grdsales.ColWidth(14) = 0
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(16) = 0
    
    LBLTOTAL.Caption = 0
    
    PHYFLAG = True
    TMPFLAG = True
    BATCH_FLAG = True
    ITEM_FLAG = True
    
    TXTSLNO.Enabled = False
    txtdate.Enabled = False
    TXTPATIENT.Enabled = False
    TXTDOCTOR.Enabled = False
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTTAX.Enabled = False
    txtexpdate.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    TXTSLNO.Text = 1
    M_EDIT = False
    Me.Width = 11400
    Me.Height = 8100
    Me.Left = 0
    Me.Top = 0
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PHYFLAG = False Then PHY.Close
    If TMPFLAG = False Then TMPREC.Close
    If BATCH_FLAG = False Then PHY_BATCH.Close
    If ITEM_FLAG = False Then PHY_ITEM.Close
    
    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDPOPUP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTQTY.Text = GRDPOPUP.Columns(1)
            TXTRATE.Text = GRDPOPUP.Columns(3)
            TXTTAX.Text = GRDPOPUP.Columns(4)
            txtexpdate.Text = GRDPOPUP.Columns(2)
            txtBatch.Text = GRDPOPUP.Columns(0)
            
            TXTVCHNO.Text = GRDPOPUP.Columns(7)
            TXTLINENO.Text = GRDPOPUP.Columns(8)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(9)
            TXTUNIT.Text = GRDPOPUP.Columns(10)
            
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

Private Sub GRDPOPUPITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo eRRHAND
    Select Case KeyCode
        Case vbKeyReturn

            TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            For i = 1 To grdsales.Rows - 1
                If Trim(grdsales.TextMatrix(i, 12)) = Trim(TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then
                        
                        Set GRDPOPUPITEM.DataSource = Nothing
                        FRMEITEM.Visible = False
                        FRMEMAIN.Enabled = True
                        TXTPRODUCT.Enabled = True
                        TXTQTY.Enabled = False
                        TXTPRODUCT.SetFocus
                        Exit Sub
                    End If
                End If
            Next i
            Set GRDPOPUPITEM.DataSource = Nothing
            If ITEM_FLAG = True Then
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            End If
            If PHY_ITEM.RecordCount = 0 Then
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            End If
            Set GRDPOPUPITEM.DataSource = PHY_ITEM
            If PHY_ITEM.RecordCount = 1 Then
                TXTQTY.Text = GRDPOPUPITEM.Columns(2)
                TXTRATE.Text = GRDPOPUPITEM.Columns(3)
                TXTTAX.Text = GRDPOPUPITEM.Columns(4)
                txtexpdate.Text = GRDPOPUPITEM.Columns(7)
                txtBatch.Text = GRDPOPUPITEM.Columns(6)
                
                TXTVCHNO.Text = GRDPOPUPITEM.Columns(8)
                TXTLINENO.Text = GRDPOPUPITEM.Columns(9)
                TXTTRXTYPE.Text = GRDPOPUPITEM.Columns(10)
                TXTUNIT.Text = GRDPOPUPITEM.Columns(5)
                
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            ElseIf PHY_ITEM.RecordCount > 1 Then
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEGRDTMP.Visible = False
                Call FILL_BATCHGRID
            End If
        Case vbKeyEscape
            TXTQTY.Text = ""
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
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub Timer_Timer()
    'Timer.Interval = 30
    If lblhead(1).Visible = True Then
        lblhead(1).Visible = False
    Else
        lblhead(1).Visible = True
    End If
End Sub

Private Sub TimerPRINT_Timer()
    If MsgBox("DO YOU WANT ONE MORE PRINT", vbYesNo, "PRINT..") = vbYes Then Shell "C:\WINDOWS\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
    TimerPRINT.Enabled = False
    CMDREFRESH.Enabled = True
    CMDREFRESH.SetFocus
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyEscape
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = True
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            txtexpdate.SetFocus
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
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtBillNo.Text) = 0 Then Exit Sub
            txtdate.Enabled = True
            txtBillNo.Enabled = False
            txtdate.SetFocus
    End Select
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

Private Sub txtdate_GotFocus()
    txtdate.SelStart = 0
    txtdate.SelLength = Len(txtdate.Text)
End Sub

Private Sub txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstDUMMYNAME As ADODB.Recordset
    Dim i As Integer

    Select Case KeyCode
        Case vbKeyReturn
            If Not IsDate(txtdate.Text) Then Exit Sub
            i = 0
            TXTPATIENT.Clear
            grdsales.Rows = 1
            TXTSLNO.Text = 1
            Set rstDUMMYNAME = New ADODB.Recordset
            rstDUMMYNAME.Open "SELECT DISTINCT PATIENT FROM DUMMYBILL", db2, adOpenStatic, adLockReadOnly
            Do Until rstDUMMYNAME.EOF
                TXTPATIENT.AddItem (Trim(rstDUMMYNAME!PATIENT))
                i = i + 1
                rstDUMMYNAME.MoveNext
            Loop
            Set rstDUMMYNAME = Nothing
            txtdate.Enabled = False
            TXTPATIENT.Enabled = True
            TXTPATIENT.SetFocus
        Case vbKeyEscape
            txtdate.Enabled = False
            txtBillNo.Enabled = True
            txtBillNo.SetFocus
    End Select
End Sub

Private Sub txtdate_LostFocus()
    txtdate.Text = Format(txtdate.Text, "DD/MM/YYYY")
End Sub

Private Sub TXTDISC_GotFocus()
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.Text)
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            cmdadd.Enabled = True
            TXTDISC.Enabled = False
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = True
            TXTDISC.Enabled = False
            txtBatch.SetFocus
    End Select
End Sub

Private Sub TXTDISC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDISC_LostFocus()
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 2)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".00")
End Sub

Private Sub TXTDOCTOR_GotFocus()
    TXTDOCTOR.SelStart = 0
    TXTDOCTOR.SelLength = Len(TXTDOCTOR.Text)
End Sub

Private Sub TXTDOCTOR_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(TXTDOCTOR.Text) = "" Then Exit Sub
            TXTSLNO.Enabled = True
            TXTSLNO.Text = grdsales.Rows
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTPATIENT.Enabled = True
            TXTDOCTOR.Enabled = False
            TXTPATIENT.SetFocus
    End Select
End Sub

Private Sub TXTDOCTOR_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTEXPDATE_GotFocus()
    txtexpdate.SelStart = 0
    txtexpdate.SelLength = Len(txtexpdate.Text)
End Sub

Private Sub TXTEXPDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Not IsDate(txtexpdate.Text) Then Exit Sub
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = True
            TXTDISC.Enabled = False
            txtBatch.SetFocus
        Case vbKeyEscape
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = True
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTTAX.SetFocus
    End Select
End Sub

Private Sub TXTEXPDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKey0 To vbKey9, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPATIENT_GotFocus()
    TXTPATIENT.SelStart = 0
    TXTPATIENT.SelLength = Len(TXTPATIENT.Text)
End Sub

Private Sub TXTPATIENT_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RSTDUMMY As ADODB.Recordset
'Dim i As Integer
    Select Case KeyCode
        Case vbKeyReturn
'            If Trim(TXTPATIENT.Text) = "" Then TXTPATIENT.Text = "CASH"
'            grdsales.Rows = 1
'            i = 0
'            LBLTOTAL.Caption = ""
'            Set RSTDUMMY = New ADODB.Recordset
'            RSTDUMMY.Open "SELECT *  FROM DUMMYBILL WHERE PATIENT = '" & TXTPATIENT.Text & "' ORDER BY LINE_NO", db2, adOpenStatic, adLockReadOnly
'            Do Until RSTDUMMY.EOF
'                grdsales.Rows = grdsales.Rows + 1
'                grdsales.FixedRows = 1
'                i = i + 1
'                TXTDOCTOR.Text = IIf(IsNull(RSTDUMMY!DOCTOR), "", RSTDUMMY!DOCTOR)
'                grdsales.TextMatrix(i, 0) = i
'                grdsales.TextMatrix(i, 1) = RSTDUMMY!ITEM_CODE
'                grdsales.TextMatrix(i, 2) = RSTDUMMY!ITEM_NAME
'                grdsales.TextMatrix(i, 12) = RSTDUMMY!ITEM_CODE
'                grdsales.TextMatrix(i, 17) = RSTDUMMY!MFGR
'                grdsales.TextMatrix(i, 3) = RSTDUMMY!QTY
'                grdsales.TextMatrix(i, 5) = RSTDUMMY!MRP
'                grdsales.TextMatrix(i, 6) = RSTDUMMY!SALES_PRICE
'                grdsales.TextMatrix(i, 9) = RSTDUMMY!REF_NO
'                grdsales.TextMatrix(i, 11) = RSTDUMMY!TRX_TOTAL
'                grdsales.TextMatrix(i, 10) = RSTDUMMY!EXP_DATE
'                LBLTOTAL.Caption = Format(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 11)), ".00")
'                RSTDUMMY.MoveNext
'            Loop
'            RSTDUMMY.Close
'            Set RSTDUMMY = Nothing
'
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = True
            TXTDOCTOR.SetFocus
        Case vbKeyEscape
            TXTPATIENT.Enabled = False
            txtdate.Enabled = True
            txtdate.SetFocus
    End Select
End Sub

Private Sub TXTPATIENT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
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
            CmdDelete.Enabled = False
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            txtexpdate.Text = ""
            txtBatch.Text = ""
            LBLSUBTOTAL.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
           
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 12)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then Exit Sub
                    End If
                Next i
                Call FILL_BATCHGRID
                Exit Sub
            ElseIf PHY.RecordCount > 1 Then
                'FRMSUB.grdsub.Columns(0).Visible = True
                'FRMSUB.grdsub.Columns(1).Caption = "ITEM NAME"
                'FRMSUB.grdsub.Columns(1).Width = 3200
                'FRMSUB.grdsub.Columns(2).Caption = "QTY"
                'FRMSUB.grdsub.Columns(2).Width = 1300
                Call FILL_ITEMGRID
            End If
        
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CmdDelete.Enabled = False
    
        Case vbKeyEscape
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            TXTTAX.Text = ""
            txtexpdate.Text = ""
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            
            TXTPATIENT.Enabled = True
            TXTDOCTOR.Enabled = True
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
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
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 Then Exit Sub

            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = True
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTRATE.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
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
    TXTQTY.Text = Format(TXTQTY.Text, ".00")
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 2)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".00")
End Sub

Private Sub TXTRATE_GotFocus()
    TXTRATE.SelStart = 0
    TXTRATE.SelLength = Len(TXTRATE.Text)
End Sub

Private Sub TXTRATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTRATE.Text) = 0 Then Exit Sub
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = True
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTTAX.SetFocus
        Case vbKeyEscape
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub TXTRATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRATE_LostFocus()
    TXTRATE.Text = Format(TXTRATE.Text, ".00")
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 2)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".00")
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
                TXTRATE.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTUNIT.Text = ""
                LBLSUBTOTAL.Caption = ""
                txtexpdate.Text = ""
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
                TXTRATE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 11), ".00")
                
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 12)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 15)
                TXTUNIT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                LBLSUBTOTAL.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 11)
                txtexpdate.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                
                TXTPATIENT.Enabled = False
                TXTDOCTOR.Enabled = False
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                TXTRATE.Enabled = False
                TXTTAX.Enabled = False
                txtexpdate.Enabled = False
                txtBatch.Enabled = False
                TXTDISC.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTPATIENT.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TXTUNIT.Text = ""
                TXTQTY.Text = ""
                TXTRATE.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                txtexpdate.Text = ""
                txtBatch.Text = ""
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTPATIENT.Enabled = True
                TXTDOCTOR.Enabled = True
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.Rows > 1 Then
                CMDREFRESH.Enabled = True
                CMDPRINT.Enabled = True
                CMDPRINT.SetFocus
            Else
                TXTDOCTOR.Enabled = True
                TXTDOCTOR.SetFocus
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

Private Sub TXTTAX_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
End Sub

Private Sub TXTTAX_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            If txtexpdate.Text = "" Then
                TXTEXPIRY.Enabled = True
                TXTEXPIRY.SetFocus
            Else
                txtexpdate.Enabled = True
                txtexpdate.SetFocus
            End If
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            
        Case vbKeyEscape
            TXTPATIENT.Enabled = False
            TXTDOCTOR.Enabled = False
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = True
            TXTTAX.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTRATE.SetFocus
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

Private Sub TXTTAX_LostFocus()
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 2)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".00")
End Sub

Private Sub TXTTOTALDISC_GotFocus()
    TXTTOTALDISC.SelStart = 0
    TXTTOTALDISC.SelLength = Len(TXTTOTALDISC.Text)
End Sub

Private Sub TXTTOTALDISC_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cmdReportGenerate_Click()
    Dim vlineCount As Integer
    Dim vpageCount As Integer
    Dim SN As Integer
    Dim i As Integer
    
    vlineCount = 0
    vpageCount = 1
    SN = 0
    
    '//NOTE : Report file name should never contain blank space.
    Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    
    Print #1, Chr(27) & Chr(71) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & "SARAS MEDICALS" & _
              Chr(27) & Chr(72)
              
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & " Kaichoondy Junction" & Space(22) & "Phone: 0477-3290525"
      
    Print #1, Space(7) & "Alappuzha 688006" & Space(15) & "DL No. 6-176/20/2003 Dtd. 31.10.2003"
              
    Print #1, Space(7) & "TIN No.32041339615"
              
    Print #1, Chr(27) & Chr(71) & Chr(10) & Space(7) & AlignRight("INVOICE FORM 8BF", 38)
        
    Print #1, Space(7) & "Bill No. " & Trim(txtBillNo.Text) & Chr(27) & Chr(72) & Space(29) & "Date:" & txtdate.Text '& Space(2) & txttime.Text
    Print #1, Chr(27) & Chr(72) & Space(7) & "Patient: " & Trim(TXTPATIENT.Text) & Space(27); "Doctor: " & TXTDOCTOR.Text
              
    Print #1, Chr(27) & Chr(72) & Space(7) & "Salesman: CS"
    
    Print #1, Space(7) & RepeatString("-", 65)
    Print #1, Space(7) & AlignLeft(" SN", 2) & Space(2) & _
            AlignLeft("Description", 11) & Space(17) & _
            AlignLeft("MFR", 3) & Space(3) & _
            AlignLeft("Batch", 6) & Space(2) & _
            AlignLeft("Ex.Dt", 6) & Space(1) & _
            AlignLeft("Qty", 4) & Space(1) & _
            AlignLeft("Amount", 6) & _
            Chr(27) & Chr(72)  '//Bold Ends
            
    Print #1, Space(7) & RepeatString("-", 65)
              
    For i = 1 To grdsales.Rows - 1
        Print #1, Space(7) & AlignRight(Str(i), 2) & Space(1) & _
            AlignLeft(grdsales.TextMatrix(i, 2), 28) & _
            AlignLeft(grdsales.TextMatrix(i, 17), 5) & Space(2) & _
            AlignLeft(grdsales.TextMatrix(i, 9), 9) & _
            AlignLeft(Format(grdsales.TextMatrix(i, 10), "mm/yy"), 5) & _
            AlignRight(grdsales.TextMatrix(i, 3), 4) & Space(1) & _
            AlignRight(grdsales.TextMatrix(i, 11), 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Next i
    
    Print #1, Space(7) & AlignRight("-------------", 65)
    Print #1, Chr(27) & Chr(71) & Chr(10) & Space(9) & AlignRight("NET AMOUNT", 52) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 10)
    Print #1, Chr(27) & Chr(67) & Chr(0)
    Print #1, Chr(27) & Chr(72) & Space(7) & AlignRight("**** WISH YOU A SPEEDY RECOVERY ****", 40) & AlignRight("Pharmacist:", 13)
   
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
    
End Sub

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
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
             If Len(Trim(TXTEXPIRY.Text)) = 1 Then GoTo Nextstep
            If Len(Trim(TXTEXPIRY.Text)) < 5 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
Nextstep:
            TXTEXPIRY.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    'TXTEXPIRY.Text = Format(TXTEXPIRY.Text, "MM/YY")
    'TXTEXPIRY.Visible = False
End Sub
Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Function FILL_BATCHGRID()
    Dim RSTBALQTY As ADODB.Recordset
    
    Set RSTBALQTY = New ADODB.Recordset
    RSTBALQTY.Open "Select BAL_QTY From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
    If Not (RSTBALQTY.EOF Or RSTBALQTY.BOF) Then

    Else
        GoTo SKIP
    End If
    RSTBALQTY.Close
    Set RSTBALQTY = Nothing
    
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "BATCH NO."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "EXP DATE"
    GRDPOPUP.Columns(3).Caption = "PRICE"
    GRDPOPUP.Columns(4).Caption = "TAX"
    GRDPOPUP.Columns(5).Caption = "VCH No"
    GRDPOPUP.Columns(6).Caption = "Line No"
    GRDPOPUP.Columns(7).Caption = "Trx Type"
    
    GRDPOPUP.Columns(0).Width = 1400
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 1400
    GRDPOPUP.Columns(3).Width = 1000
    GRDPOPUP.Columns(4).Width = 900
    
    GRDPOPUP.Columns(5).Visible = False
    GRDPOPUP.Columns(6).Visible = False
    
    GRDPOPUP.SetFocus
    lblhead(0).Caption = GRDPOPUP.Columns(6).Text
    lblhead(9).Visible = True
    lblhead(0).Visible = True
    Exit Function
SKIP:
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
End Function

Function FILL_ITEMGRID()
    FRMEMAIN.Enabled = False
    FRMEITEM.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEGRDTMP.Visible = False
    
    
    If ITEM_FLAG = True Then
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    End If

    Set GRDPOPUPITEM.DataSource = PHY_ITEM
    GRDPOPUPITEM.RowHeight = 250
    GRDPOPUPITEM.Columns(0).Visible = False
    GRDPOPUPITEM.Columns(1).Caption = "ITEM NAME"
    GRDPOPUPITEM.Columns(1).Width = 3800
    GRDPOPUPITEM.Columns(2).Caption = "QTY"
    GRDPOPUPITEM.Columns(2).Width = 1300
    GRDPOPUPITEM.SetFocus
End Function
