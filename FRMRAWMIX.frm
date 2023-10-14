VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMRAWMIX 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRODUCTION - I"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12345
   Icon            =   "FRMRAWMIX.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   12345
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   8400
      Left            =   -135
      TabIndex        =   30
      Top             =   -15
      Width           =   12510
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
         TabIndex        =   52
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00FAF2F1&
         Height          =   1065
         Left            =   210
         TabIndex        =   31
         Top             =   -30
         Width           =   12270
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
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
            Left            =   9495
            ScaleHeight     =   240
            ScaleWidth      =   1335
            TabIndex        =   72
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
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
            Left            =   9510
            ScaleHeight     =   240
            ScaleWidth      =   855
            TabIndex        =   71
            Top             =   870
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
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
            Left            =   7680
            ScaleHeight     =   240
            ScaleWidth      =   1965
            TabIndex        =   70
            Top             =   600
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
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
            Left            =   7680
            ScaleHeight     =   240
            ScaleWidth      =   1800
            TabIndex        =   69
            Top             =   870
            Visible         =   0   'False
            Width           =   1800
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
            Height          =   360
            Left            =   3990
            MaxLength       =   100
            TabIndex        =   2
            Top             =   615
            Width           =   2670
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
            Left            =   1590
            TabIndex        =   0
            Top             =   195
            Visible         =   0   'False
            Width           =   885
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   345
            Left            =   1545
            TabIndex        =   1
            Top             =   660
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
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
            Height          =   300
            Index           =   2
            Left            =   3045
            TabIndex        =   59
            Top             =   660
            Width           =   900
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Prodn Date"
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
            Left            =   105
            TabIndex        =   53
            Top             =   690
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Production No."
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
            TabIndex        =   36
            Top             =   240
            Width           =   1440
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
            Left            =   3390
            TabIndex        =   35
            Top             =   195
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
            Left            =   3990
            TabIndex        =   34
            Top             =   165
            Width           =   1335
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
            Left            =   5445
            TabIndex        =   33
            Top             =   165
            Width           =   1230
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
            Left            =   1590
            TabIndex        =   32
            Top             =   210
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0FF&
         Height          =   5040
         Left            =   210
         TabIndex        =   37
         Top             =   945
         Width           =   12270
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
            Height          =   315
            Left            =   4650
            TabIndex        =   68
            Top             =   1725
            Visible         =   0   'False
            Width           =   1110
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4860
            Left            =   45
            TabIndex        =   29
            Top             =   135
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   8573
            _Version        =   393216
            Rows            =   1
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
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
         TabIndex        =   51
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00FAF2F1&
         Height          =   2475
         Left            =   210
         TabIndex        =   38
         Top             =   5895
         Width           =   12270
         Begin VB.TextBox TxtLRate 
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
            Left            =   7875
            MaxLength       =   7
            TabIndex        =   16
            Top             =   1140
            Width           =   915
         End
         Begin VB.TextBox txtTotalLoose 
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
            Left            =   10800
            MaxLength       =   7
            TabIndex        =   10
            Top             =   465
            Width           =   855
         End
         Begin VB.TextBox txtreference 
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
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   8685
            MaxLength       =   7
            TabIndex        =   22
            Top             =   1515
            Width           =   3540
         End
         Begin VB.TextBox Txtbatch 
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
            Left            =   10680
            MaxLength       =   25
            TabIndex        =   20
            Top             =   1140
            Width           =   1545
         End
         Begin VB.TextBox txtMRP 
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
            Left            =   4200
            MaxLength       =   7
            TabIndex        =   12
            Top             =   1140
            Width           =   885
         End
         Begin VB.TextBox TxtPack 
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
            Left            =   10185
            MaxLength       =   7
            TabIndex        =   9
            Top             =   465
            Width           =   600
         End
         Begin VB.TextBox txtBarcode 
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
            Left            =   5190
            MaxLength       =   30
            TabIndex        =   21
            Top             =   1515
            Width           =   2790
         End
         Begin VB.CommandButton CmDCancel 
            Caption         =   "&Cancel"
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
            Left            =   5655
            TabIndex        =   26
            Top             =   1920
            Width           =   1125
         End
         Begin VB.TextBox TxttaxMRP 
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
            Left            =   8805
            MaxLength       =   7
            TabIndex        =   17
            Top             =   1140
            Width           =   675
         End
         Begin VB.TextBox TXTRETAIL 
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
            TabIndex        =   13
            Top             =   1140
            Width           =   915
         End
         Begin VB.TextBox txtWS 
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
            Left            =   6030
            MaxLength       =   7
            TabIndex        =   14
            Top             =   1140
            Width           =   960
         End
         Begin VB.TextBox txtvanrate 
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
            Left            =   7005
            MaxLength       =   7
            TabIndex        =   15
            Top             =   1140
            Width           =   855
         End
         Begin VB.TextBox TxtResult 
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
            Left            =   8925
            MaxLength       =   7
            TabIndex        =   7
            Top             =   465
            Width           =   780
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
            TabIndex        =   58
            Top             =   2460
            Visible         =   0   'False
            Width           =   795
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
            TabIndex        =   57
            Top             =   2625
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox TXTPRODUCT2 
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
            Left            =   75
            TabIndex        =   3
            Top             =   465
            Width           =   4095
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
            Height          =   375
            Left            =   4185
            MaxLength       =   7
            TabIndex        =   5
            Top             =   465
            Width           =   960
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
            Height          =   450
            Left            =   3255
            TabIndex        =   24
            Top             =   1920
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
            Height          =   450
            Left            =   7005
            TabIndex        =   27
            Top             =   1920
            Width           =   1140
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
            Left            =   2055
            TabIndex        =   23
            Top             =   1920
            Width           =   1125
         End
         Begin VB.CommandButton CmdDelete 
            Caption         =   "&Delete"
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
            Left            =   825
            TabIndex        =   28
            Top             =   1920
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
            Left            =   8700
            TabIndex        =   43
            Top             =   2625
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
            TabIndex        =   42
            Top             =   2775
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
            TabIndex        =   41
            Top             =   2775
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox TxtActqty 
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
            Left            =   1185
            TabIndex        =   40
            Top             =   2505
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
            TabIndex        =   39
            Top             =   2535
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
            Height          =   450
            Left            =   4455
            TabIndex        =   25
            Top             =   1920
            Width           =   1125
         End
         Begin MSDataListLib.DataList DataList1 
            Height          =   1035
            Left            =   75
            TabIndex        =   4
            Top             =   855
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1826
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
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   345
            Left            =   9495
            TabIndex        =   18
            Top             =   1155
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   609
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
            Height          =   345
            Left            =   9495
            TabIndex        =   19
            Top             =   1155
            Width           =   1170
            _ExtentX        =   2064
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
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L. Rate"
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
            Left            =   7875
            TabIndex        =   81
            Top             =   870
            Width           =   915
         End
         Begin VB.Label LBLLPACK 
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
            Height          =   375
            Left            =   11670
            TabIndex        =   11
            Top             =   465
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Total Qty"
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
            Height          =   300
            Index           =   8
            Left            =   10800
            TabIndex        =   80
            Top             =   150
            Width           =   1410
         End
         Begin VB.Label lblcost 
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
            Height          =   360
            Left            =   9105
            TabIndex        =   79
            Top             =   1905
            Width           =   1290
         End
         Begin VB.Label lblreference 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Cost"
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
            Height          =   360
            Index           =   0
            Left            =   8160
            TabIndex        =   78
            Top             =   1905
            Width           =   930
         End
         Begin VB.Label lblreference 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Ref."
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
            Height          =   360
            Index           =   8
            Left            =   7995
            TabIndex        =   77
            Top             =   1515
            Width           =   690
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
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   7
            Left            =   10680
            TabIndex        =   76
            Top             =   870
            Width           =   1545
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
            Index           =   6
            Left            =   4200
            TabIndex        =   75
            Top             =   870
            Width           =   885
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
            Left            =   9495
            TabIndex        =   74
            Top             =   870
            Width           =   1170
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
            Height          =   300
            Index           =   5
            Left            =   10185
            TabIndex        =   73
            Top             =   150
            Width           =   600
         End
         Begin VB.Label lblpack 
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
            Height          =   375
            Left            =   9720
            TabIndex        =   8
            Top             =   465
            Width           =   450
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
            Left            =   8805
            TabIndex        =   67
            Top             =   870
            Width           =   675
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Retail"
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
            Left            =   5100
            TabIndex        =   66
            Top             =   870
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "W.Rate"
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
            Index           =   27
            Left            =   6030
            TabIndex        =   65
            Top             =   870
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "V.Rate"
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
            Index           =   32
            Left            =   7005
            TabIndex        =   64
            Top             =   870
            Width           =   855
         End
         Begin VB.Label lBLpRODUCT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   5160
            TabIndex        =   6
            Top             =   465
            Width           =   3750
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Output Product"
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
            Height          =   300
            Index           =   4
            Left            =   5160
            TabIndex        =   63
            Top             =   150
            Width           =   3750
         End
         Begin VB.Label flagchange2 
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   7710
            TabIndex        =   62
            Top             =   915
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbldealer2 
            BackColor       =   &H00FAF2F1&
            Height          =   315
            Left            =   8370
            TabIndex        =   61
            Top             =   1800
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Output Qty"
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
            Height          =   300
            Index           =   3
            Left            =   8925
            TabIndex        =   60
            Top             =   150
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Mixture Name"
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
            Index           =   9
            Left            =   75
            TabIndex        =   50
            Top             =   150
            Width           =   4095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "No. of Mix"
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
            Left            =   4185
            TabIndex        =   49
            Top             =   150
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
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   15
            Left            =   4185
            TabIndex        =   48
            Top             =   1515
            Width           =   990
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
            TabIndex        =   47
            Top             =   2790
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Act. Qty"
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
            Left            =   45
            TabIndex        =   46
            Top             =   2520
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
            TabIndex        =   45
            Top             =   2745
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
            TabIndex        =   44
            Top             =   2790
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   -15
      TabIndex        =   56
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11355
      TabIndex        =   55
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   54
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMRAWMIX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim MIX_ITEM As New ADODB.Recordset
Dim MIX_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_STOCK As Integer
Dim M_EDIT As Boolean
Dim EDIT_INV, OLD_INV As Boolean

Private Sub CMDADD_Click()
    Dim rststock As ADODB.Recordset
    'Dim RSTMINQTY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Long
    
    If DataList1.BoundText = "" Then
        MsgBox "Select Mixture from the List", , "Production"
        TXTPRODUCT2.SetFocus
        Exit Sub
    End If
    
    If Val(TXTQTY.Text) = 0 Then
        MsgBox "Enter the number of mixture", , "Production"
        TXTQTY.SetFocus
        Exit Sub
    End If
    
    If Val(TxtResult.Text) = 0 Then
        MsgBox "Output Qty could not be Zero", , "Production"
        TxtResult.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ERRHAND
    grdsales.FixedRows = 0
    grdsales.rows = 1
        
    i = 1
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM  TRXFORMULASUB WHERE FOR_NO = " & DataList1.BoundText & " AND TRX_TYPE='FR'", db, adOpenStatic, adLockReadOnly, adCmdText
    With RSTITEMMAST
        Do Until .EOF
            grdsales.rows = grdsales.rows + 1
            grdsales.FixedRows = 1
            grdsales.TextMatrix(i, 0) = i
            grdsales.TextMatrix(i, 1) = RSTITEMMAST!ITEM_NAME
            grdsales.TextMatrix(i, 4) = RSTITEMMAST!ITEM_CODE
            If UCase(RSTITEMMAST!Category) = "SERVICE CHARGE" Then
                grdsales.TextMatrix(i, 2) = ""
                grdsales.TextMatrix(i, 3) = ""
                grdsales.TextMatrix(i, 5) = ""
                grdsales.TextMatrix(i, 6) = ""
                grdsales.TextMatrix(i, 7) = RSTITEMMAST!QTY
            Else
                grdsales.TextMatrix(i, 2) = RSTITEMMAST!QTY * Val(TXTQTY)
                grdsales.TextMatrix(i, 3) = 1
                grdsales.TextMatrix(i, 5) = IIf(IsNull(RSTITEMMAST!LOOSE_PACK), "1", RSTITEMMAST!LOOSE_PACK)
                grdsales.TextMatrix(i, 6) = IIf(IsNull(RSTITEMMAST!PACK_TYPE), "", RSTITEMMAST!PACK_TYPE)
            End If
            grdsales.TextMatrix(i, 8) = IIf(IsNull(RSTITEMMAST!Category), "", RSTITEMMAST!Category)
            grdsales.TextMatrix(i, 9) = IIf(IsNull(RSTITEMMAST!WASTE_QTY), 0, RSTITEMMAST!WASTE_QTY) * Val(TXTQTY)
            i = i + 1
            .MoveNext
        Loop
    End With
    Set RSTITEMMAST = Nothing
    
    
    cmdadd.Enabled = False
    ''CmdDelete.Enabled = False
    CmdExit.Enabled = False
    M_EDIT = False
    Call cost_calculate
    'grdsales.TopRow = grdsales.Rows - 1

    cmdRefresh.Enabled = True
    cmdRefresh.SetFocus
Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            'cmdadd.Enabled = False
            txtBatch.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub cmdcancel_Click()
    If MsgBox("Are you sure you want to cancel?", vbYesNo, "Production") = vbNo Then Exit Sub
    Call cancel_bill
End Sub

Private Function cancel_bill()
    On Error GoTo ERRHAND
    Dim rstBILL As ADODB.Recordset
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'MI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
        
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    LBLDATE.Caption = Date
    lbltime.Caption = Time
    'LBLTOTALCOST.Caption = ""
    grdsales.rows = 1
    M_EDIT = False
    EDIT_INV = False
    TXTPRODUCT2.Text = ""
    TXTQTY.Text = ""
    lBLpRODUCT.Caption = ""
    TXTITEMCODE.Text = ""
    TxtResult.Text = ""
    txtTotalLoose.Text = ""
    TxtBarcode.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LblPack.Caption = ""
    LBLLPACK.Caption = ""
    txtBatch.Text = ""
    txtretail.Text = ""
    TxtMRP.Text = ""
    Txtpack.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    TxttaxMRP.Text = ""
    
    cmdRefresh.Enabled = False
    CmdExit.Enabled = True
    CmdPrint.Enabled = False
    CmdExit.Enabled = True
    FRMEHEAD.Enabled = True
    OLD_INV = False
    TXTPRODUCT2.SetFocus
    'LBLITEMCOST.Caption = ""
    TXTQTY.Tag = ""
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub CmdDelete_Click()
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstTRXMAST As ADODB.Recordset

    If grdsales.rows <= 1 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE THE ENTIRE PRODUCTION", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    If OLD_INV = False Then
        Call cancel_bill
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'    With RSTTRXFILE
'        If Not (.EOF And .BOF) Then
'            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
'            !ISSUE_QTY = !ISSUE_QTY + Val(TxtResult.Text)
'            !FREE_QTY = 0
'            !ISSUE_VAL = 0
'            !CLOSE_QTY = !CLOSE_QTY - Val(TxtResult.Text)
'            !CLOSE_VAL = 0
'            RSTTRXFILE.Update
'        End If
'    End With
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing

    For i = 1 To grdsales.rows - 1
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        With RSTTRXFILE
'            If Not (.EOF And .BOF) Then
'                !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(i, 2)) '/ Val(grdsales.TextMatrix(i, 5)))
'                !FREE_QTY = 0
'                !ISSUE_VAL = 0
'                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(i, 2)) '/ Val(grdsales.TextMatrix(i, 5)))
'                !CLOSE_VAL = 0
'                !LOOSE_PACK = Val(grdsales.TextMatrix(i, 5))
'                !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
'                RSTTRXFILE.Update
'            End If
'        End With
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
                
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                !LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
                !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
                !PD_NO = Val(txtBillNo.Text)
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    Next i
    
    db.Execute "delete From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete FROM RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    
    Dim RSTITEMMAST, rststock As ADODB.Recordset
    Dim INWARD, OUTWARD As Double
    INWARD = 0
    OUTWARD = 0
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        INWARD = 0
        OUTWARD = 0
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Do Until rststock.EOF
            INWARD = INWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY)) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
            
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockOptimistic, adCmdText
        'rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Do Until rststock.EOF
            OUTWARD = OUTWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)) + IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
            OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        
        RSTITEMMAST!CLOSE_QTY = INWARD - OUTWARD
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    For i = 1 To grdsales.rows - 1
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTITEMMAST.EOF
            INWARD = 0
            OUTWARD = 0
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            Do Until rststock.EOF
                INWARD = INWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY)) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            Set rststock = New ADODB.Recordset
            'rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until rststock.EOF
                OUTWARD = OUTWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)) + IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
                OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            RSTITEMMAST!CLOSE_QTY = INWARD - OUTWARD
            RSTITEMMAST.Update
            RSTITEMMAST.MoveNext
        Loop
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    Next i
    
    Call cancel_bill
    Screen.MousePointer = vbNormal
    Exit Sub
'    For i = Val(TXTSLNO.Text) - 1 To grdsales.Rows - 2
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = i
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = grdsales.TextMatrix(i + 1, 1)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = grdsales.TextMatrix(i + 1, 2)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = grdsales.TextMatrix(i + 1, 3)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = grdsales.TextMatrix(i + 1, 4)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(i + 1, 6)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = grdsales.TextMatrix(i + 1, 5)
'        grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = grdsales.TextMatrix(i + 1, 7)
'    Next i
'    grdsales.Rows = grdsales.Rows - 1
    
    Call cost_calculate
    
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    cmdadd.Enabled = False
    'CmdDelete.Enabled = False
    CmdExit.Enabled = False
    M_EDIT = False
    EDIT_INV = True
    If grdsales.rows = 1 Then
'        CMDEXIT.Enabled = True
        CmdPrint.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdExit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    
    If grdsales.rows = 1 Then Exit Sub
    
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
    Call Generateprint
    
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    Dim Num As Currency
    
    On Error GoTo ERRHAND
    
    db.Execute "delete From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "MI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = "GENERAL"
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 2))
        RSTTRXFILE!WASTE_QTY = Val(grdsales.TextMatrix(i, 9))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = 0
        RSTTRXFILE!PTR = 0
        RSTTRXFILE!SALES_PRICE = 0
        RSTTRXFILE!SALES_TAX = 0
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 3)
        RSTTRXFILE!VCH_DESC = "Issued to      Press"
        RSTTRXFILE!REF_NO = ""
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = ""
        RSTTRXFILE!MFGR = ""
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = 0
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!P_RETAIL = 0
        RSTTRXFILE!P_RETAILWOTAX = 0
        RSTTRXFILE!SALE_1_FLAG = ""
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "" 'DataList2.BoundText
        
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Call ReportGeneratION
    ReportNameVar = Rptpath & "rptRAWBILL"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "( {TRXFILE.TRX_TYPE}='MI' AND {TRXFILE.VCH_NO}= " & Val(txtBillNo.Text) & " )"
    
    Set CRXFormulaFields = Report.FormulaFields
    
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
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
    
    Set Printer = Printers(barcodeprinter)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        'If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & DataList2.Text & "'"
'        If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & lbladdress.Caption & "'"
'        If CRXFormulaField.Name = "{@DLNO2}" Then CRXFormulaField.Text = "'" & LBLDLNO2.Caption & "'"
'        If CRXFormulaField.Name = "{@DLNO}" Then CRXFormulaField.Text = "'" & lbldlno.Caption & "'"
'        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
'        If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'TAX INVOICE FORM 8H/8B/8'"
'        If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & lbltin.Caption & "'"
'        If lblcredit.Caption = "0" Then
'            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'CASH'"
'        Else
'            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'" & txtcrdays.Text & "'" & "' Days'"
'        End If
    Next
    frmreport.Caption = "BILL"
    Call GENERATEREPORT
    
    ''cmdRefresh.SetFocus
'
    
    CmdExit.Enabled = False
    'TXTQTY.Enabled = False
    
    ''rptPRINT.Action = 1
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub cmdRefresh_Click()
    
   ' If grdsales.Rows = 1 Then GoTo SKIP
    
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
    
    If DataList1.BoundText = "" Then
        MsgBox "Please select the output Product", vbOKOnly, "Production"
        Exit Sub
    End If
    If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
    
    If Val(TxtLRate.Text) > Val(txtretail.Text) And Val(Txtpack.Text) <> 1 Then
        MsgBox "Loose Price cannot be greater than Retail Price", vbOKOnly, "Production"
        TxtLRate.SetFocus
        Exit Sub
    End If
    
    If Val(txtretail.Text) > Val(TxtMRP.Text) And Val(TxtMRP.Text) <> 0 Then
        MsgBox "Price cannot be greater than MRP", vbOKOnly, "Production"
        txtretail.SetFocus
        Exit Sub
    End If
    
    If Val(txtWS.Text) > Val(TxtMRP.Text) And Val(TxtMRP.Text) <> 0 Then
        MsgBox "Price cannot be greater than MRP", vbOKOnly, "Production"
        txtWS.SetFocus
        Exit Sub
    End If
    
    If Val(txtvanrate.Text) > Val(TxtMRP.Text) And Val(TxtMRP.Text) <> 0 Then
        MsgBox "Price cannot be greater than MRP", vbOKOnly, "Production"
        txtvanrate.SetFocus
        Exit Sub
    End If
    
    Dim i As Single
    Dim M, n As Integer
    Dim ObjFile, objText, Text
    Dim rstformula As ADODB.Recordset
    Dim pergr As Integer
    pergr = 0
    Set rstformula = New ADODB.Recordset
    rstformula.Open "select * from ITEMMAST where ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstformula.EOF Or rstformula.BOF) Then
        If IsNull(rstformula!item_spec) Then
            pergr = 0
        Else
            pergr = IIf(IsNull(rstformula!item_spec), 0, Val(rstformula!item_spec))
        End If
    End If
    rstformula.Close
    Set rstformula = Nothing
    If MDIMAIN.StatusBar.Panels(6).Text = "Y" Then
    
        If Trim(TxtBarcode.Text) = "" Then TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Int(Val(txtretail.Text))
        If MsgBox("Do you want to Print Barcode Labels", vbYesNo, "Production.....") = vbYes Then
            i = Val(InputBox("Enter number of lables to be print", "No. of labels..", Val(TxtResult.Text)))
            If BARTEMPLATE = "Y" Then
                If i = 0 Then GoTo SKIP_BARCODE
                If Val(MDIMAIN.LBLLABELNOS.Caption) = 0 Then MDIMAIN.LBLLABELNOS.Caption = 1
                i = i / Val(MDIMAIN.LBLLABELNOS.Caption)
                If Math.Abs(i - Fix(i)) > 0 Then
                    i = Int(i) + 1
                End If
                If FileExists(App.Path & "\template1.txt") Then
                    Set ObjFile = CreateObject("Scripting.FileSystemObject")
                    Set objText = ObjFile.OpenTextFile(App.Path & "\template1.txt")
                    Text = objText.ReadAll
                    objText.Close
                
                    Set objText = Nothing
                    Set ObjFile = Nothing
                    
                    If pergr > 1 And Val(txtretail.Text) <> 0 Then
                        Text = Replace(Text, "[PPPPPPPP]", "" & Round(Val(txtretail.Text) / pergr, 3) & "") 'pergram
                    Else
                        Text = Replace(Text, "[PPPPPPPP]", "")   'REF (SPEC)
                    End If
                    
                    Text = Replace(Text, "[AAAAAAAA]", "" & Trim(txtreference.Text) & "")  'REF (SPEC)
                    Text = Replace(Text, "[BBBBBBBB]", "") 'PACK
                    Text = Replace(Text, "[DDDDDDDD]", "" & Format(Val(TxtMRP.Text), "0.00") & "")  'MRP
                    If IsDate(TXTEXPDATE.Text) Then
                        Text = Replace(Text, "[EEEEEEEE]", "" & Format(TXTEXPDATE.Text, "dd/mm/yyyy") & "")  'EXP DATE
                        If IsDate(TXTINVDATE.Text) Then
                            Text = Replace(Text, "[CCCCCCCC]", "" & Format(TXTINVDATE.Text, "dd/mm/yyyy") & "")  'PROD DATE
                        Else
                            Text = Replace(Text, "[CCCCCCCC]", "")  'PROD DATE
                        End If
                    Else
                        Text = Replace(Text, "[EEEEEEEE]", "")   'EXP DATE
                        Text = Replace(Text, "[CCCCCCCC]", "")  'PROD DATE
                    End If
                    Text = Replace(Text, "[FFFFFFFF]", "" & Left(Trim(lBLpRODUCT.Caption), 30) & "")  'ITEM NAME
                    Text = Replace(Text, "[NNNNNNNN]", "" & Trim(TXTITEMCODE.Text) & "")  'ITEM CODE
                    Text = Replace(Text, "[GGGGGGGG]", "" & Trim(TxtBarcode.Text) & "")  'BARCODE
                    'If BARFORMAT = "Y" Then
                        If Len(Trim(TxtBarcode.Text)) Mod 2 = 0 Then
                            Text = Replace(Text, "[LLLLLLLL]", "" & Trim(TxtBarcode.Text) & "")  'BARCODE
                            Text = Replace(Text, "[MMMMMMMM]", "" & Trim(TxtBarcode.Text) & "")  'BARCODE
                        Else
                            Text = Replace(Text, "[LLLLLLLL]", "" & Mid(Trim(TxtBarcode.Text), 1, Len(Trim(TxtBarcode.Text)) - 1) & "!100" & Right(Trim(TxtBarcode.Text), 1) & "") 'BARCODE
                            Text = Replace(Text, "[MMMMMMMM]", "" & Mid(Trim(TxtBarcode.Text), 1, Len(Trim(TxtBarcode.Text)) - 1) & ">6" & Right(Trim(TxtBarcode.Text), 1) & "") 'BARCODE
                        End If
                    'End If
                    Text = Replace(Text, "[HHHHHHHH]", "" & Format(Val(txtretail.Text), "0.00") & "")  'PRICE
                    Text = Replace(Text, "[IIIIIIII]", "" & Trim(txtBatch.Text) & "")  'BATCH
                    Text = Replace(Text, "[JJJJJJJJ]", "" & Trim(MDIMAIN.StatusBar.Panels(5).Text) & "")  'COMP NAME
                    
                    Dim intFile As Integer
                    Dim strFile As String
                    If FileExists(App.Path & "\BARCODE.PRN") Then
                        Kill (App.Path & "\BARCODE.PRN")
                    End If
                    strFile = App.Path & "\BARCODE.PRN" 'the file you want to save to
                    intFile = FreeFile
                    Open strFile For Output As #intFile
                        Print #intFile, Text 'the data you want to save
                    Close #intFile
                    
                    On Error GoTo CLOSEFILE
                    Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
                    If err.Number = 55 Then
                        Close #1
                        Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
                    End If
                    On Error GoTo ERRHAND
                    
                    'Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
                    Print #1, "COPY/B " & App.Path & "\BARCODE.PRN " & BarPrint
                    Print #1, "EXIT"
                    Close #1
                    
                    '//HERE write the proper path where your command.com file exist
                    For M = 1 To i
                        Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & Rptpath & "REPO.BAT N", vbHide
                    Next M
                Else
                    MsgBox "No template exists", , "EzBiz"
                    Exit Sub
                End If
                
            Else
                db.Execute "Delete from barprint"
                
                Dim RSTTRXFILE As ADODB.Recordset
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "Select * From barprint", db, adOpenStatic, adLockOptimistic, adCmdText
                For M = 1 To i
                    RSTTRXFILE.AddNew
                    RSTTRXFILE!BARCODE = "*" & Trim(TxtBarcode.Text) & "*"
                    RSTTRXFILE!ITEM_NAME = Trim(lBLpRODUCT.Caption)
                    RSTTRXFILE!item_Price = Val(txtretail.Text)
                    If IsDate(TXTEXPDATE.Text) Then
                        RSTTRXFILE!expdate = Format(TXTEXPDATE.Text, "dd/mm/yyyy")
                    End If
                    If IsDate(TXTINVDATE.Text) Then
                        RSTTRXFILE!pckdate = Format(TXTINVDATE.Text, "dd/mm/yyyy")
                    End If
                    RSTTRXFILE!item_MRP = Val(TxtMRP.Text)
                    RSTTRXFILE!item_color = Trim(txtBatch.Text)
                    RSTTRXFILE!REMARKS = Trim(txtreference.Text)
                    RSTTRXFILE!item_spec = pergr
                    RSTTRXFILE!COMP_NAME = Trim(MDIMAIN.StatusBar.Panels(5).Text)
                    RSTTRXFILE.Update
                Next M
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
                ReportNameVar = Rptpath & "Rptbarprn"
                Set Report = crxApplication.OpenReport(ReportNameVar, 1)
                Set CRXFormulaFields = Report.FormulaFields
            
                For n = 1 To Report.Database.Tables.COUNT
                    Report.Database.Tables.Item(n).SetLogOnInfo strConnection
                    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                        Set oRs = New ADODB.Recordset
                        Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(n).Name & " ")
                        Report.Database.SetDataSource oRs, 3, n
                        Set oRs = Nothing
                    End If
                Next n
                
                Set Printer = Printers(barcodeprinter)
                Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
                Report.DiscardSavedData
                Report.VerifyOnEveryPrint = True
                Report.PrintOut (False)
                Set CRXFormulaFields = Nothing
                Set crxApplication = Nothing
                Set Report = Nothing
            End If
        End If
    End If
SKIP_BARCODE:
    Call AppendSale
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
'    Me.Enabled = False
'    FRMDEBIT.Show
    
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
'            TXTPRODUCT.Text = ""
'            TXTQTY.Text = ""
'            TXTITEMCODE.Text = ""
'            TXTVCHNO.Text = ""
'            TXTLINENO.Text = ""
'            TXTTRXTYPE.Text = ""
'            TXTUNIT.Text = ""
'            TXTPRODUCT.SetFocus
'            TXTQTY.Enabled = False
'            CMDMODIFY.Enabled = False
'            'CmdDelete.Enabled = False
    End Select
End Sub

Private Sub Form_Activate()
    If TXTPRODUCT2.Enabled = True Then TXTPRODUCT2.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
    If TXTREMARKS.Enabled = True Then TXTREMARKS.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ERRHAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'MI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
        
    MIX_FLAG = True
    ACT_FLAG = True
    LBLDATE.Caption = Date
    lbltime.Caption = Time
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 4000
    grdsales.ColWidth(2) = 1500
    'grdsales.ColWidth(3) = 0
    'grdsales.ColWidth(4) = 0
    'grdsales.ColWidth(5) = 0
        
    If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
        LBLCOST.Visible = True
        lblreference(0).Visible = True
    Else
        LBLCOST.Visible = False
        lblreference(0).Visible = False
    End If
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM NAME"
    grdsales.TextArray(2) = "QTY"
    grdsales.TextArray(3) = "PACK"
    grdsales.TextArray(4) = "ITEM CODE"
    grdsales.TextArray(5) = "Loose Pack"
    grdsales.TextArray(6) = "Pack Type"
    grdsales.TextArray(7) = "Amount"
    grdsales.TextArray(8) = "Category"
    grdsales.TextArray(9) = "Waste"
    
    
    'TXTQTY.Enabled = False
    'CmdDelete.Enabled = False
    CmdPrint.Enabled = False
    
    CLOSEALL = 1
    M_EDIT = False
    EDIT_INV = False
'    Me.Width = 11700
'    Me.Height = 10185
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If MIX_FLAG = False Then MIX_ITEM.Close
        If ACT_FLAG = False Then ACT_REC.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdsales_Click()
    TXTsample.Visible = False
    grdsales.SetFocus
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    If grdsales.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113
            'If OLD_INV = True Then Exit Sub
            Select Case grdsales.Col
                Case 2
                    TXTsample.Visible = True
                    TXTsample.Top = grdsales.CellTop + 130
                    TXTsample.Left = grdsales.CellLeft + 50
                    TXTsample.Width = grdsales.CellWidth
                    TXTsample.Height = grdsales.CellHeight
                    TXTsample.Text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TXTsample.SetFocus
                Case 9
                    TXTsample.Visible = True
                    TXTsample.Top = grdsales.CellTop + 130
                    TXTsample.Left = grdsales.CellLeft + 50
                    TXTsample.Width = grdsales.CellWidth
                    TXTsample.Height = grdsales.CellHeight
                    TXTsample.Text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TXTsample.SetFocus
            End Select
    End Select
End Sub

Private Sub grdsales_Scroll()
    TXTsample.Visible = False
    grdsales.SetFocus
End Sub

Private Sub txtBillNo_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim TRXMAST As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    
    Dim i As Long
    Dim n As Integer
    Dim M As Integer

    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            CmdExit.Enabled = True
            If Val(txtBillNo.Text) = 0 Then Exit Sub
            grdsales.rows = 1
            Set TRXFILE = New ADODB.Recordset
            TRXFILE.Open "Select * FROM RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockReadOnly
            If Not (TRXFILE.EOF And TRXFILE.BOF) Then
                i = i + 1
                TXTINVDATE.Text = Format(TRXFILE!VCH_DATE, "DD/MM/YYYY")
                LBLDATE.Caption = Format(Date, "DD/MM/YYYY")
                lbltime.Caption = Time
                lBLpRODUCT.Caption = IIf(IsNull(TRXFILE!ITEM_NAME), "", TRXFILE!ITEM_NAME)
                TXTITEMCODE.Text = IIf(IsNull(TRXFILE!ITEM_CODE), "", TRXFILE!ITEM_CODE)
                TXTPRODUCT2.Text = IIf(IsNull(TRXFILE!FORM_NAME), "", TRXFILE!FORM_NAME)
                TXTQTY.Text = IIf(IsNull(TRXFILE!FORM_QTY), "", TRXFILE!FORM_QTY)
                txtretail.Text = IIf(IsNull(TRXFILE!P_RETAIL), "", TRXFILE!P_RETAIL)
                TxtLRate.Text = IIf(IsNull(TRXFILE!P_CRTN), "", TRXFILE!P_CRTN)
                txtBatch.Text = IIf(IsNull(TRXFILE!REF_NO), "", TRXFILE!REF_NO)
                TxtMRP.Text = IIf(IsNull(TRXFILE!MRP), "", TRXFILE!MRP)
                Txtpack.Text = IIf(IsNull(TRXFILE!LOOSE_PACK), "1", TRXFILE!LOOSE_PACK)
                If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
                txtWS.Text = IIf(IsNull(TRXFILE!P_WS), "", TRXFILE!P_WS)
                txtvanrate.Text = IIf(IsNull(TRXFILE!P_VAN), "", TRXFILE!P_VAN)
                TxttaxMRP.Text = IIf(IsNull(TRXFILE!SALES_TAX), "", TRXFILE!SALES_TAX)
                TxtResult.Text = IIf(IsNull(TRXFILE!QTY), "", Format(TRXFILE!QTY / Val(Txtpack.Text), "0.00"))
                LblPack.Caption = IIf(IsNull(TRXFILE!PACK_TYPE), "", TRXFILE!PACK_TYPE)
                TxtBarcode.Text = IIf(IsNull(TRXFILE!BARCODE), "", TRXFILE!BARCODE)
                TXTEXPDATE.Text = IIf(IsDate(TRXFILE!EXP_DATE), TRXFILE!EXP_DATE, "  /  /    ")
                TXTEXPIRY.Text = IIf(IsDate(TRXFILE!EXP_DATE), Format(TRXFILE!EXP_DATE, "mm/yy"), "  /  ")
                i = 0
                Set TRXMAST = New ADODB.Recordset
                TRXMAST.Open "Select * From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockReadOnly
                Do Until TRXMAST.EOF
                    i = i + 1
                    grdsales.rows = grdsales.rows + 1
                    grdsales.FixedRows = 1
                    grdsales.TextMatrix(i, 0) = i
                    grdsales.TextMatrix(i, 1) = TRXMAST!ITEM_NAME
                    grdsales.TextMatrix(i, 4) = TRXMAST!ITEM_CODE
                    grdsales.TextMatrix(i, 8) = IIf(IsNull(TRXMAST!Category), "", TRXMAST!Category)
                    grdsales.TextMatrix(i, 9) = IIf(IsNull(TRXMAST!WASTE_QTY), 0, TRXMAST!WASTE_QTY)
                    If UCase(grdsales.TextMatrix(i, 8)) <> "SERVICE CHARGE" Then
                        grdsales.TextMatrix(i, 2) = TRXMAST!QTY
                        grdsales.TextMatrix(i, 3) = 1 'TRXMAST!UNIT
                        grdsales.TextMatrix(i, 5) = IIf(IsNull(TRXMAST!LOOSE_PACK), 1, TRXMAST!LOOSE_PACK)
                        grdsales.TextMatrix(i, 6) = TRXMAST!PACK_TYPE
                    Else
                        grdsales.TextMatrix(i, 7) = IIf(IsNull(TRXMAST!ITEM_COST), 0, TRXMAST!ITEM_COST)
                    End If
                    TRXMAST.MoveNext
                Loop
                TRXMAST.Close
                Set TRXMAST = Nothing
                'txtBillNo.Text = ""
                'LBLBILLNO.Caption = ""
                'Call ResetStock
                CmdExit.Enabled = False
                OLD_INV = True
                cmdadd.Enabled = False
                cmdRefresh.Enabled = False
            Else
                OLD_INV = False
            End If
            TRXFILE.Close
            Set TRXFILE = Nothing
            'LBLTIME.Caption = IIf(IsNull(TRXMAST!CFORM_NO), Time, TRXMAST!CFORM_NO)
            
            LBLBILLNO.Caption = Val(txtBillNo.Text)
            
            Call cost_calculate
            
            txtBillNo.Visible = False
            
            If OLD_INV = True Then
                If grdsales.rows > 1 Then
                    cmdRefresh.Enabled = True
                    cmdRefresh.SetFocus
                Else
                    TXTREMARKS.SetFocus
                End If
            End If
    End Select
    DataList1.Text = TXTPRODUCT2.Text
    Call DataList1_Click
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

    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'MI'", db, adOpenStatic, adLockReadOnly
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
    TRXMAST.Open "Select MIN(VCH_NO) From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'MI'", db, adOpenStatic, adLockReadOnly
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
    Call txtBillNo_KeyDown(13, 0)
    Exit Sub
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
            
'            If DateDiff("d", Date, TXTEXPDATE.Text) < 60 Then
'                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days", vbOKOnly, "EzBiz"
'                TXTEXPDATE.Text = "  /  /    "
'                TXTEXPIRY.SelStart = 0
'                TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
'                TXTEXPIRY.SetFocus
'                Exit Sub
'            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 60 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "EzBiz") = vbNo Then
                    TXTEXPDATE.Text = "  /  /    "
                    TXTEXPIRY.SelStart = 0
                    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
                    TXTEXPIRY.SetFocus
                    Exit Sub
                End If
            End If
SKIP:
'            If OLD_INV = False Then
                TXTEXPIRY.Visible = False
                'TXTEXPDATE.Enabled = False
                txtBatch.Enabled = True
                txtBatch.SetFocus
'            Else
'                cmdadd.Enabled = False
'                cmdRefresh.Enabled = False
'            End If
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            'TXTEXPDATE.Enabled = False
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
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

Private Sub TXTEXPDATE_GotFocus()
    'TXTEXPDATE.BackColor = &H98F3C1
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
            
'            If DateDiff("d", Date, TXTEXPDATE.Text) < 60 Then
'                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days", vbOKOnly, "EzBiz"
'                TXTEXPDATE.SelStart = 0
'                TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
'                TXTEXPDATE.SetFocus
'                Exit Sub
'            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 60 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "EzBiz") = vbNo Then
                    TXTEXPDATE.SelStart = 0
                    TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
                    TXTEXPDATE.SetFocus
                    Exit Sub
                End If
            End If
SKID:
'            If OLD_INV = False Then
                TXTEXPIRY.Visible = False
                'TXTEXPDATE.Enabled = False
                cmdadd.Enabled = True
                cmdadd.SetFocus
'            Else
'                cmdadd.Enabled = False
'                cmdRefresh.Enabled = False
'            End If
        Case vbKeyEscape
            If TXTEXPDATE.Text = "  /  /    " Then GoTo SKIP
            If Not IsDate(TXTEXPDATE.Text) Then Exit Sub
SKIP:
            TxttaxMRP.Enabled = True
            'TXTEXPDATE.Enabled = False
            TXTEXPIRY.Visible = False
            TxttaxMRP.SetFocus
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
    'TXTEXPDATE.BackColor = vbWhite
    TXTEXPDATE.Text = Format(TXTEXPDATE.Text, "DD/MM/YYYY")
    If IsDate(TXTEXPDATE.Text) Then TXTEXPIRY.Text = Format(TXTEXPDATE.Text, "MM/YY")
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
                TXTREMARKS.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
            End If
        Case vbKeyEscape
            'Exit Sub
            'If EDIT_INV = True Then Exit Sub
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

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            i = 0
'            Set RSTTRXFILE = New ADODB.Recordset
'            RSTTRXFILE.Open "SELECT BAL_QTY  FROM RTRXFILE WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND VCH_NO = " & Val(TXTVCHNO.Text) & " AND LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockReadOnly
'            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
'                If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
'                i = RSTTRXFILE!BAL_QTY
'            End If
'            RSTTRXFILE.Close
'            Set RSTTRXFILE = Nothing
'
'            If Val(TXTQTY.Text) = 0 Then Exit Sub
'            If i > 0 Then
'                If Val(TXTQTY.Text) > i Then
'                    MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
'                    TXTQTY.SelStart = 0
'                    TXTQTY.SelLength = Len(TXTQTY.Text)
'                    Exit Sub
'                End If
'            End If
'SKIP:
            TxtResult.SetFocus
         Case vbKeyEscape
            TXTPRODUCT2.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    TXTQTY.Text = Format(TXTQTY.Text, ".00")
    If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
    TxtResult.Text = Format(Val(TXTQTY) * Val(TxtActqty.Text), "0.00")
End Sub

'Private Function COSTCALCULATION()
'    Dim RSTCOST As ADODB.Recordset
'    Dim COST As Double
'    Dim N As Integer
'    'Dim RSTITEMMAST As ADODB.Recordset
'
'     'LBLTOTALCOST.Caption = ""
'     'LBLPROFIT.Caption = ""
'        COST = 0
'    On Error GoTo eRRHAND
'    For N = 1 To grdsales.Rows - 1
'        Set RSTCOST = New ADODB.Recordset
'        RSTCOST.Open "SELECT ITEM_COST FROM RTRXFILE WHERE TRX_TYPE = '" & Trim(grdsales.TextMatrix(N, 7)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(N, 5)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(N, 6)) & "", db, adOpenStatic, adLockReadOnly, adCmdText
'        Do Until RSTCOST.EOF
'            'COST = COST + (RSTCOST!ITEM_COST) * Val(grdsales.TextMatrix(N, 3))
'            RSTCOST.MoveNext
'        Loop
'        RSTCOST.Close
'        Set RSTCOST = Nothing
'    Next N
'
'    'LBLTOTALCOST.Caption = Round(COST, 2)
'    'LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - COST, 2)
'
'    Exit Function
'
'eRRHAND:
'    MsgBox Err.Description
'End Function

Private Function AppendSale()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTP_RATE As ADODB.Recordset
    Dim RSTITEMMAST, RSTRTRXFILE, rststock As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim i, M_DATA As Double
    Dim TRXVALUE As Double
    
    Dim DAY_DATE As String
    Dim MONTH_DATE As String
    Dim YEAR_DATE As String
    Dim E_DATE As Date
    i = 0
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    
    ''db.RollbackTrans
    
    If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
    db.BeginTrans
    db.Execute "delete From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & ""

    
    E_DATE = Format(TXTINVDATE.Text, "MM/DD/YYYY")
    If Day(E_DATE) <= 12 Then
        DAY_DATE = Format(Month(E_DATE), "00")
        MONTH_DATE = Format(Day(E_DATE), "00")
        YEAR_DATE = Format(Year(E_DATE), "0000")
        E_DATE = DAY_DATE & "/" & MONTH_DATE & "/" & YEAR_DATE
    End If
    E_DATE = Format(E_DATE, "MM/DD/YYYY")
    
    
    Dim rstTRXMAST As ADODB.Recordset
    Dim ITEMCOST As Double
    ITEMCOST = 0
    For i = 1 To grdsales.rows - 1
        If UCase(grdsales.TextMatrix(i, 8)) <> "SERVICE CHARGE" Then
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
            With rstTRXMAST
                If Not (.EOF And .BOF) Then
                    .Properties("Update Criteria").Value = adCriteriaKey
                    !ISSUE_QTY = !ISSUE_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                    !FREE_QTY = 0
                    !ISSUE_VAL = 0
                    !CLOSE_QTY = !CLOSE_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                    !CLOSE_VAL = 0
                    '!LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
                    !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
                    ITEMCOST = ITEMCOST + IIf(IsNull(!ITEM_COST), 0, !ITEM_COST * Val(grdsales.TextMatrix(i, 2)) + Val(grdsales.TextMatrix(i, 9)))
                    rstTRXMAST.Update
                End If
            End With
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            If OLD_INV = True Then
                Set rstTRXMAST = New ADODB.Recordset
                rstTRXMAST.Open "SELECT *  FROM RTRXFILE WHERE PD_NO = '" & Val(txtBillNo.Text) & "' AND  TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' and ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
                With rstTRXMAST
                    If Not (.EOF And .BOF) Then
                        .Properties("Update Criteria").Value = adCriteriaKey
    '                    If OLD_INV = True Then
    '                        M_DATA = Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9))
    '                        M_DATA = M_DATA - (!QTY - !BAL_QTY)
    '                    Else
    '                        M_DATA = Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9))
    '                    End If
                        If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                        !ISSUE_QTY = !ISSUE_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                        If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                        !BAL_QTY = !BAL_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                        '!BAL_QTY = M_DATA
                        '!LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
                        !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
                        !PD_NO = Val(txtBillNo.Text)
                        rstTRXMAST.Update
                    Else
                        rstTRXMAST.Close
                        Set rstTRXMAST = Nothing
                        GoTo SKIP
                    End If
                End With
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
            Else
SKIP:
                Set rstTRXMAST = New ADODB.Recordset
                rstTRXMAST.Open "SELECT *  FROM RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' and ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
                With rstTRXMAST
                    If Not (.EOF And .BOF) Then
                        .Properties("Update Criteria").Value = adCriteriaKey
    '                    If OLD_INV = True Then
    '                        M_DATA = Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9))
    '                        M_DATA = M_DATA - (!QTY - !BAL_QTY)
    '                    Else
    '                        M_DATA = Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9))
    '                    End If
                        If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                        !ISSUE_QTY = !ISSUE_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                        If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                        !BAL_QTY = !BAL_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                        '!BAL_QTY = M_DATA
                        '!LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
                        !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
                        !PD_NO = Val(txtBillNo.Text)
                        rstTRXMAST.Update
                    End If
                End With
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
            End If
        Else
            ITEMCOST = ITEMCOST + Val(grdsales.TextMatrix(i, 7))
        End If
    Next i
    ITEMCOST = Round(ITEMCOST / (Val(TxtResult.Text) * Val(Txtpack.Text)), 3)
    'ITEMCOST = Round(ITEMCOST * 100 / (Val(TxttaxMRP.Text) + 100), 3)
    M_DATA = 0
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
    'RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!TRX_TYPE = "MI"
        RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTRTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTRTRXFILE!LINE_NO = 1
        RSTRTRXFILE!ITEM_CODE = TXTITEMCODE.Text
        RSTRTRXFILE!QTY = Round(Val(TxtResult.Text) * Val(Txtpack.Text), 3)
        RSTRTRXFILE!BAL_QTY = Round(Val(TxtResult.Text) * Val(Txtpack.Text), 3)
        
        RSTRTRXFILE!ITEM_COST = ITEMCOST
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  From ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").Value = adCriteriaKey
                rststock!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
                !CLOSE_QTY = !CLOSE_QTY + (Val(TxtResult.Text) * Val(Txtpack.Text))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = 0
                
                !RCPT_QTY = !RCPT_QTY + (Val(TxtResult.Text) * Val(Txtpack.Text))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                !RCPT_VAL = 0 ' !RCPT_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
                
                If Val(txtretail.Text) <> 0 Then
                    !P_RETAIL = Val(txtretail.Text)
                End If
                If Val(TxtLRate.Text) = 0 Then
                    !P_CRTN = Val(txtretail.Text) / Val(Txtpack.Text)
                Else
                    !P_CRTN = Val(TxtLRate.Text)
                End If
                If Val(TxtMRP.Text) <> 0 Then !MRP = Val(TxtMRP.Text)
                If Val(txtWS.Text) <> 0 Then !P_WS = Val(txtWS.Text)
                If Val(txtvanrate.Text) <> 0 Then !P_VAN = Val(txtvanrate.Text)
                If Val(Txtpack.Text) > 1 Then !LOOSE_PACK = Val(Txtpack.Text)
                !CRTN_PACK = 1
                !ITEM_COST = ITEMCOST
                RSTRTRXFILE!MFGR = !MANUFACTURER
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        
    Else
        M_DATA = Round(Val(TxtResult.Text) * Val(Txtpack.Text), 3)
        M_DATA = M_DATA - (RSTRTRXFILE!QTY - RSTRTRXFILE!BAL_QTY)
        RSTRTRXFILE!BAL_QTY = M_DATA
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  From ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").Value = adCriteriaKey
                rststock!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
                !CLOSE_QTY = !CLOSE_QTY - RSTRTRXFILE!QTY
                !CLOSE_QTY = !CLOSE_QTY + (Val(TxtResult.Text) * Val(Txtpack.Text))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = 0 '!CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
                
                !RCPT_QTY = !RCPT_QTY - RSTRTRXFILE!QTY
                !RCPT_QTY = !RCPT_QTY + (Val(TxtResult.Text) * Val(Txtpack.Text))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                !RCPT_VAL = 0 '!RCPT_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
                !ITEM_COST = ITEMCOST
                RSTRTRXFILE!MFGR = !MANUFACTURER
                If Val(txtretail.Text) <> 0 Then
                    !P_RETAIL = Val(txtretail.Text)
                End If
                If Val(TxtLRate.Text) = 0 Then
                    !P_CRTN = Val(txtretail.Text) / Val(Txtpack.Text)
                Else
                    !P_CRTN = Val(TxtLRate.Text)
                End If
                If Val(Txtpack.Text) > 1 Then !LOOSE_PACK = Val(Txtpack.Text)
                !CRTN_PACK = 1
                If Val(TxtMRP.Text) <> 0 Then !MRP = Val(TxtMRP.Text)
                If Val(txtWS.Text) <> 0 Then !P_WS = Val(txtWS.Text)
                If Val(txtvanrate.Text) <> 0 Then !P_VAN = Val(txtvanrate.Text)
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        RSTRTRXFILE!QTY = Round((Val(TxtResult.Text) * Val(Txtpack.Text)), 3)
    End If
    RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    RSTRTRXFILE!ITEM_NAME = lBLpRODUCT.Caption
    RSTRTRXFILE!FORM_CODE = DataList1.BoundText
    RSTRTRXFILE!FORM_QTY = Val(TXTQTY.Text)
    RSTRTRXFILE!FORM_NAME = DataList1.Text
    RSTRTRXFILE!TRX_TOTAL = 0 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
    RSTRTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "dd/mm/yyyy")
    RSTRTRXFILE!BARCODE = Trim(TxtBarcode.Text)
    RSTRTRXFILE!P_RETAIL = Val(txtretail.Text)
    RSTRTRXFILE!LOOSE_PACK = Val(Txtpack.Text)
    RSTRTRXFILE!P_WS = Val(txtWS.Text)
    RSTRTRXFILE!P_VAN = Val(txtvanrate.Text)
    RSTRTRXFILE!SALES_TAX = Val(TxttaxMRP.Text)
    RSTRTRXFILE!CRTN_PACK = 1
    If Val(TxtLRate.Text) <> 0 And Val(Txtpack.Text) <> 1 Then
        RSTRTRXFILE!P_CRTN = Val(TxtLRate.Text)
    Else
        RSTRTRXFILE!P_CRTN = Val(txtretail.Text) / Val(Txtpack.Text)
    End If
    'RSTRTRXFILE!LOOSE_PACK = 1
    RSTRTRXFILE!LINE_DISC = 1 ' Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
    RSTRTRXFILE!P_DISC = 0 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
    RSTRTRXFILE!MRP = Val(TxtMRP.Text)
    RSTRTRXFILE!PTR = 0 ' Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 9))
    RSTRTRXFILE!SALES_PRICE = 0
    RSTRTRXFILE!Category = "OWN"
    RSTRTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
    RSTRTRXFILE!REF_NO = Trim(txtBatch.Text)
    RSTRTRXFILE!CST = 0
        
    RSTRTRXFILE!SCHEME = 0
    If IsDate(TXTEXPDATE.Text) Then
        RSTRTRXFILE!EXP_DATE = IIf(IsDate(TXTEXPDATE.Text), Format(TXTEXPDATE.Text, "dd/mm/yyyy"), Null)
    End If
    RSTRTRXFILE!FREE_QTY = 0
    RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!C_USER_ID = "SM"
    RSTRTRXFILE!check_flag = "V"
    RSTRTRXFILE.Update
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    TRXVALUE = 0
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "Select * From TRXSUB ", db, adOpenStatic, adLockOptimistic, adCmdText
'    For i = 1 To grdsales.Rows - 1
'        RSTTRXFILE.AddNew
'        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
'        RSTTRXFILE!TRX_TYPE = "MI"
'        RSTTRXFILE!LINE_NO = i
'        RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(i, 5) = "", 0, grdsales.TextMatrix(i, 5))
'        RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(i, 6) = "", 0, grdsales.TextMatrix(i, 6))
'        RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(i, 7) = "", "MI", grdsales.TextMatrix(i, 7))
'        RSTTRXFILE!QTY = grdsales.TextMatrix(i, 3)
'        RSTTRXFILE.Update
'    Next i
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    
    Dim RSTITEMCOST As ADODB.Recordset
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "MI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = Trim(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!WASTE_QTY = Val(grdsales.TextMatrix(i, 9))
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 1)
        If UCase(grdsales.TextMatrix(i, 8)) <> "SERVICE CHARGE" Then
            RSTTRXFILE!QTY = (Val(grdsales.TextMatrix(i, 2))) '* Val(grdsales.TextMatrix(i, 5)))
            RSTTRXFILE!LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
            RSTTRXFILE!PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
            RSTTRXFILE!LOOSE_FLAG = "L"
            RSTTRXFILE!ITEM_COST = 0
            Set RSTITEMCOST = New ADODB.Recordset
            RSTITEMCOST.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            With RSTITEMCOST
                If Not (.EOF And .BOF) Then
                    RSTTRXFILE!ITEM_COST = RSTITEMCOST!ITEM_COST
                End If
            End With
            RSTITEMCOST.Close
            Set RSTITEMCOST = Nothing
        Else
            RSTTRXFILE!QTY = 0
            RSTTRXFILE!LOOSE_PACK = 0
            RSTTRXFILE!PACK_TYPE = ""
            RSTTRXFILE!LOOSE_FLAG = ""
            RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(i, 7))
        End If
        RSTTRXFILE!MRP = 0
        RSTTRXFILE!PTR = 0
        RSTTRXFILE!SALES_PRICE = 0
        RSTTRXFILE!P_RETAIL = 0
        RSTTRXFILE!P_RETAILWOTAX = 0
        RSTTRXFILE!SALES_TAX = 0
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 3)
        RSTTRXFILE!VCH_DESC = "Issued to      Factory" '& Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = ""
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = ""
        RSTTRXFILE!MFGR = ""
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = 0
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "" 'DataList2.BoundText
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='MI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "MI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    End If
    RSTTRXFILE!VCH_AMOUNT = 0
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = "10111"
    RSTTRXFILE!ACT_NAME = "PRESS"
    RSTTRXFILE!DISCOUNT = 0
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = 0
    RSTTRXFILE!REF_NO = ""
    RSTTRXFILE!SLSM_CODE = ""
    RSTTRXFILE!PAY_AMOUNT = 0
    RSTTRXFILE!check_flag = ""
    RSTTRXFILE!POST_FLAG = ""
    RSTTRXFILE!CFORM_NO = lbltime.Caption
    RSTTRXFILE!REMARKS = ""
    RSTTRXFILE!DISC_PERS = 0
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!BILL_NAME = DataList1.Text
    RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!C_USER_ID = "SM"
    
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    i = 0
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'MI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    Dim INWARD, OUTWARD, BALQTY, DIFFQTY As Double
    
    For i = 1 To grdsales.rows - 1
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTITEMMAST.EOF
            INWARD = 0
            OUTWARD = 0
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            Do Until rststock.EOF
                INWARD = INWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                INWARD = INWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) ' * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockReadOnly
            Do Until rststock.EOF
                OUTWARD = OUTWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)) + IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
                'OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            '=============
            
            BALQTY = 0
            db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY <0"
            If Round(INWARD - OUTWARD, 2) = 0 Then
                db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY >0"
            Else
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT SUM(BAL_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY > 0", db, adOpenForwardOnly
                If Not (rststock.EOF And rststock.BOF) Then
                    BALQTY = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
                End If
                rststock.Close
                Set rststock = Nothing
            End If
        
            If Round(INWARD - OUTWARD, 2) < BALQTY Then
                DIFFQTY = BALQTY - (Round(INWARD - OUTWARD, 2))
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ", db, adOpenStatic, adLockOptimistic, adCmdText
                Do Until rststock.EOF
                    If DIFFQTY - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) >= 0 Then
                        DIFFQTY = DIFFQTY - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY)
                        rststock!BAL_QTY = 0
                        rststock.Update
                    Else
                        rststock!BAL_QTY = Round(IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) - DIFFQTY, 2)
                        DIFFQTY = 0
                        rststock.Update
                    End If
                    If DIFFQTY <= 0 Then Exit Do
                    rststock.MoveNext
                Loop
                rststock.Close
                Set rststock = Nothing
            ElseIf Round(INWARD - OUTWARD, 2) > BALQTY Then
                DIFFQTY = Round((INWARD - OUTWARD), 2) - BALQTY
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockOptimistic, adCmdText
                Do Until rststock.EOF
                    If DIFFQTY <= IIf(IsNull(rststock!QTY), 0, rststock!QTY) - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) Then
                        rststock!BAL_QTY = Round(IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) + DIFFQTY, 2)
                        DIFFQTY = 0
                    Else
                        If Not rststock!BAL_QTY = IIf(IsNull(rststock!QTY), 0, rststock!QTY) Then
                            rststock!BAL_QTY = Round(IIf(IsNull(rststock!QTY), 0, rststock!QTY), 2)
                            DIFFQTY = DIFFQTY - IIf(IsNull(rststock!QTY), 0, rststock!QTY)
                        End If
                    End If
                    rststock.Update
                    If DIFFQTY <= 0 Then Exit Do
                    rststock.MoveNext
                Loop
                rststock.Close
                Set rststock = Nothing
                'MsgBox ""
            End If
            
            '============
            
            RSTITEMMAST!CLOSE_QTY = INWARD - OUTWARD
            RSTITEMMAST!RCPT_QTY = INWARD
            RSTITEMMAST!ISSUE_QTY = OUTWARD
            RSTITEMMAST.Update
            RSTITEMMAST.MoveNext
        Loop
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    Next i
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & TXTITEMCODE.Text & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTITEMMAST.EOF
        INWARD = 0
        OUTWARD = 0
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Do Until rststock.EOF
            INWARD = INWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) ' * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
            INWARD = INWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until rststock.EOF
            'OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
            OUTWARD = OUTWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)) + IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
            OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        
        If Val(TxttaxMRP.Text) <> 0 Then RSTITEMMAST!SALES_TAX = Val(TxttaxMRP.Text)
        RSTITEMMAST!check_flag = "V"
        RSTITEMMAST!ITEM_COST = ITEMCOST
        RSTITEMMAST!CLOSE_QTY = INWARD - OUTWARD
        RSTITEMMAST!RCPT_QTY = INWARD
        RSTITEMMAST!ISSUE_QTY = OUTWARD
        RSTITEMMAST!PACK_TYPE = LblPack.Caption
        RSTITEMMAST.Update
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    db.CommitTrans
    
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    LBLDATE.Caption = Date
    lbltime.Caption = Time
    'LBLTOTALCOST.Caption = ""
    grdsales.rows = 1
    M_EDIT = False
    EDIT_INV = False
    TXTPRODUCT2.Text = ""
    TXTQTY.Text = ""
    lBLpRODUCT.Caption = ""
    TXTITEMCODE.Text = ""
    TxtResult.Text = ""
    txtTotalLoose.Text = ""
    TxtBarcode.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LblPack.Caption = ""
    LBLLPACK.Caption = ""
    txtBatch.Text = ""
    txtretail.Text = ""
    TxtLRate.Text = ""
    TxtMRP.Text = ""
    Txtpack.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    TxttaxMRP.Text = ""
    
    cmdRefresh.Enabled = False
    CmdExit.Enabled = True
    CmdPrint.Enabled = False
    CmdExit.Enabled = True
    FRMEHEAD.Enabled = True
    TXTPRODUCT2.SetFocus
    'LBLITEMCOST.Caption = ""
    TXTQTY.Tag = ""
    OLD_INV = False
    Screen.MousePointer = vbNormal
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Function

Private Sub ReportGeneratION()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    Dim rstSUBfile As ADODB.Recordset
    Dim SN As Integer
    Dim TRXTOTAL As Double
    
    SN = 0
    TRXTOTAL = 0
   ' On Error GoTo errHand
    '//NOTE : Report file name should never contain blank space.
    On Error GoTo CLOSEFILE
    Open App.Path & "\Report.PRN" For Output As #1 '//Report file Creation
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open App.Path & "\Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)

    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenForwardOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!Address & ", " & RSTCOMPANY!HO_NAME, 140)
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(48) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(48) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(48) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "SALES SUMMARY FOR THE PERIOD"
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    'Set RSTTRXFILE = New ADODB.Recordset
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(13) & RepeatString("-", 59)
    Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SN", 3) & Space(2) & _
            AlignLeft("INV DATE", 8) & Space(10) & _
            AlignLeft("INV AMT", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Print #1, Space(12) & RepeatString("-", 59)
    SN = 0
'    RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_NO", Conn, adOpenForwardOnly
'    Do Until RSTTRXFILE.EOF
'        SN = SN + 1
'        Print #1, Chr(27) & Chr(71) & Space(5) & Chr(14) & Chr(15) & AlignRight(Str(SN), 4) & ". " & Space(1) & _
'            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
'            AlignRight(Format(Round(RSTTRXFILE!VCH_AMOUNT, 0), "0.00"), 16)
'        'Print #1, Chr(13)
'        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
'        RSTTRXFILE.MoveNext
'    Loop
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    Print #1,
    
    Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(13) & AlignRight("TOTAL AMOUNT", 12) & AlignRight((Format(TRXTOTAL, "####.00")), 11)
    Print #1, Space(56) & RepeatString("-", 16)
    'Print #1, Chr(27) & Chr(67) & Chr(0)
    'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** WISH YOU A SPEEDY RECOVERY ****", 40)


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
    
    'Print #1, Chr(27) & Chr(80)
    Close #1 '//Closing the file
    'MsgBox "Report file generated at " & App.Path & "\Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Sub

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            FRMEHEAD.Enabled = False
            TXTPRODUCT2.SetFocus
        Case vbKeyEscape
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
    End Select
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

Private Sub TxtProduct2_Change()
    Dim rstCharge As ADODB.Recordset
    On Error GoTo ERRHAND
    If flagchange2.Caption <> "1" Then
        If MIX_FLAG = True Then
            MIX_ITEM.Open "select FOR_NO, FOR_NAME from TRXFORMULAMAST where FOR_NAME Like '" & TXTPRODUCT2.Text & "%' AND TRX_TYPE='FR' ORDER BY FOR_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            MIX_FLAG = False
        Else
            MIX_ITEM.Close
            MIX_ITEM.Open "select FOR_NO, FOR_NAME from TRXFORMULAMAST  where FOR_NAME Like '" & TXTPRODUCT2.Text & "%' AND TRX_TYPE='FR' ORDER BY FOR_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            MIX_FLAG = False
        End If
        If (MIX_ITEM.EOF And MIX_ITEM.BOF) Then
            LBLDEALER2.Caption = ""
        Else
            LBLDEALER2.Caption = MIX_ITEM!FOR_NAME
            'TxtActqty.Text = IIf(IsNull(MIX_ITEM!QTY), "", MIX_ITEM!QTY)
            'TXTITEMCODE.Text = IIf(IsNull(MIX_ITEM!ITEM_CODE), "", MIX_ITEM!ITEM_CODE)
        End If
        Set DataList1.RowSource = MIX_ITEM
        DataList1.ListField = "FOR_NAME"
        DataList1.BoundColumn = "FOR_NO"
       
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TxtProduct2_GotFocus()
    TXTPRODUCT2.SelStart = 0
    TXTPRODUCT2.SelLength = Len(TXTPRODUCT2.Text)
End Sub

Private Sub TxtProduct2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList1.VisibleCount = 0 Then Exit Sub
            DataList1.Enabled = True
            DataList1.SetFocus
    End Select
End Sub

Private Sub TxtProduct2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_Click()
    TXTPRODUCT2.Text = DataList1.Text
    LBLDEALER2.Caption = TXTPRODUCT2.Text
    
    Dim rstformula As ADODB.Recordset
    On Error GoTo ERRHAND
    
    If DataList1.BoundText = "" Then Exit Sub
    Set rstformula = New ADODB.Recordset
    rstformula.Open "select * from TRXFORMULAMAST where FOR_NO = " & DataList1.BoundText & " and TRX_TYPE='FR'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstformula.EOF Or rstformula.BOF) Then
        TxtActqty.Text = IIf(IsNull(rstformula!QTY), "", rstformula!QTY)
        Txtpack.Text = IIf(IsNull(rstformula!LOOSE_PACK), "", rstformula!LOOSE_PACK)
        LblPack.Caption = IIf(IsNull(rstformula!PACK_TYPE), "", rstformula!PACK_TYPE)
        TXTITEMCODE.Text = IIf(IsNull(rstformula!ITEM_CODE), "", rstformula!ITEM_CODE)
        lBLpRODUCT.Caption = IIf(IsNull(rstformula!ITEM_NAME), "", rstformula!ITEM_NAME)
    Else
        TxtActqty.Text = ""
        Txtpack.Text = ""
        TXTITEMCODE.Text = ""
        lBLpRODUCT.Caption = ""
        LblPack.Caption = ""
        LBLLPACK.Caption = ""
    End If
    rstformula.Close
    Set rstformula = Nothing
    
    Set rstformula = New ADODB.Recordset
    rstformula.Open "select * from ITEMMAST where ITEM_CODE = '" & TXTITEMCODE.Text & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstformula.EOF Or rstformula.BOF) Then
        LBLLPACK.Caption = IIf(IsNull(rstformula!FULL_PACK), "Nos", rstformula!FULL_PACK)
        txtretail.Text = IIf(IsNull(rstformula!P_RETAIL), "", rstformula!P_RETAIL)
        TxtLRate.Text = IIf(IsNull(rstformula!P_CRTN), "", rstformula!P_CRTN)
        TxtMRP.Text = IIf(IsNull(rstformula!MRP), "", rstformula!MRP)
        Txtpack.Text = IIf(IsNull(rstformula!LOOSE_PACK), "1", rstformula!LOOSE_PACK)
        txtWS.Text = IIf(IsNull(rstformula!P_WS), "", rstformula!P_WS)
        txtvanrate.Text = IIf(IsNull(rstformula!P_VAN), "", rstformula!P_VAN)
        TxttaxMRP.Text = IIf(IsNull(rstformula!SALES_TAX), "", rstformula!SALES_TAX)
        txtreference.Text = IIf(IsNull(rstformula!item_spec), "", rstformula!item_spec)
    End If
    rstformula.Close
    Set rstformula = Nothing
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            If DataList1.Text = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Mixture from the List", vbOKOnly, "Production"
                DataList1.SetFocus
                Exit Sub
            End If
            
            'TXTPRODUCT2.Enabled = False
            'DataList1.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
            
        Case vbKeyEscape
            TXTPRODUCT2.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_GotFocus()
    flagchange2.Caption = 1
    TXTPRODUCT2.Text = LBLDEALER2.Caption
    DataList1.Text = TXTPRODUCT2.Text
    Call DataList1_Click
End Sub

Private Sub DataList1_LostFocus()
    flagchange2.Caption = ""
    
    Dim rstformula As ADODB.Recordset
    On Error GoTo ERRHAND
    If DataList1.BoundText = "" Then Exit Sub
    
    Set rstformula = New ADODB.Recordset
    rstformula.Open "select * from TRXFORMULAMAST where FOR_NO = " & DataList1.BoundText & " and TRX_TYPE='FR'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstformula.EOF Or rstformula.BOF) Then
        TxtActqty.Text = IIf(IsNull(rstformula!QTY), "", rstformula!QTY)
        Txtpack.Text = IIf(IsNull(rstformula!LOOSE_PACK), "", rstformula!LOOSE_PACK)
        LblPack.Caption = IIf(IsNull(rstformula!PACK_TYPE), "", rstformula!PACK_TYPE)
        TXTITEMCODE.Text = IIf(IsNull(rstformula!ITEM_CODE), "", rstformula!ITEM_CODE)
        lBLpRODUCT.Caption = IIf(IsNull(rstformula!ITEM_NAME), "", rstformula!ITEM_NAME)
    Else
        TxtActqty.Text = ""
        Txtpack.Text = ""
        TXTITEMCODE.Text = ""
        lBLpRODUCT.Caption = ""
        LblPack.Caption = ""
        LBLLPACK.Caption = ""
    End If
    rstformula.Close
    Set rstformula = Nothing
    
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub


Private Sub TxtResult_GotFocus()
    TxtResult.SelStart = 0
    TxtResult.SelLength = Len(TxtResult.Text)

End Sub

Private Sub TxtResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtResult.Text) = 0 Then Exit Sub
            TxtMRP.SetFocus
         Case vbKeyEscape
            TXTQTY.SetFocus
    End Select

End Sub

Private Sub TxtResult_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub TxtResult_LostFocus()
    txtTotalLoose.Text = Round(Val(TxtResult.Text) * Val(Txtpack.Text), 2)
End Sub

Private Sub TXTRETAIL_GotFocus()
    txtretail.SelStart = 0
    txtretail.SelLength = Len(txtretail.Text)
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtWS.SetFocus
         Case vbKeyEscape
            TxtMRP.SetFocus
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
    txtretail.Text = Format(txtretail.Text, "0.00")
    
    TxtLRate.Text = Round(Val(txtretail.Text) / Val(Txtpack.Text), 2)
    TxtLRate.Text = Format(TxtLRate.Text, "0.00")
End Sub

Private Sub TXTsample_LostFocus()
    TXTsample.Visible = False
End Sub

Private Sub txtTotalLoose_LostFocus()
    If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
    TxtResult.Text = Round(Val(txtTotalLoose.Text) / Val(Txtpack.Text), 2)
End Sub

Private Sub txtws_GotFocus()
    txtWS.SelStart = 0
    txtWS.SelLength = Len(txtWS.Text)
End Sub

Private Sub txtws_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtvanrate.SetFocus
         Case vbKeyEscape
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
    txtWS.Text = Format(txtWS.Text, "0.00")
End Sub

Private Sub txtvanrate_GotFocus()
    txtvanrate.SelStart = 0
    txtvanrate.SelLength = Len(txtvanrate.Text)
End Sub

Private Sub txtvanrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtpack.Text) = 1 Then
                TxttaxMRP.SetFocus
            Else
                TxtLRate.SetFocus
            End If
         Case vbKeyEscape
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
    txtvanrate.Text = Format(txtvanrate.Text, "0.00")
End Sub

Private Sub TxttaxMRP_GotFocus()
    TxttaxMRP.SelStart = 0
    TxttaxMRP.SelLength = Len(TxttaxMRP.Text)
End Sub

Private Sub TxttaxMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTEXPIRY.Visible = True
            TXTEXPIRY.SetFocus
         Case vbKeyEscape
            txtvanrate.SetFocus
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
    TxttaxMRP.Text = Format(TxttaxMRP.Text, "0.00")
End Sub

Private Function ResetStock()
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    For i = 1 To grdsales.rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                !ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                !FREE_QTY = 0
                !ISSUE_VAL = 0
                !CLOSE_QTY = !CLOSE_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                !CLOSE_VAL = 0
                '!LOOSE_PACK = Val(grdsales.TextMatrix(i, 5))
                !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND PD_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY + (Val(grdsales.TextMatrix(i, 2)) * Val(grdsales.TextMatrix(i, 5)) + Val(grdsales.TextMatrix(i, 9)))
                !LOOSE_PACK = IIf((Val(grdsales.TextMatrix(i, 5)) = 0), 1, Val(grdsales.TextMatrix(i, 5)))
                !PACK_TYPE = Trim(grdsales.TextMatrix(i, 6))
    
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    Next i
    
End Function

Private Sub TXTsample_GotFocus()
    
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 2  ' QTY
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = TXTsample.Text
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                    Call cost_calculate
                Case 9  ' WASTAGE
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = TXTsample.Text
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                    Call cost_calculate
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            grdsales.SetFocus
    End Select
        Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case grdsales.Col
        Case 2
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

Private Sub Txtpack_GotFocus()
    If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.Text)
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtretail.SetFocus
         Case vbKeyEscape
            TxtResult.SetFocus
    End Select
End Sub

Private Sub Txtpack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtPack_LostFocus()
    txtTotalLoose.Text = Round(Val(TxtResult.Text) * Val(Txtpack.Text), 2)
    Txtpack.Text = Format(Txtpack.Text, "0.00")
End Sub

Private Sub TXTMRP_GotFocus()
    TxtMRP.SelStart = 0
    TxtMRP.SelLength = Len(TxtMRP.Text)
End Sub

Private Sub TXTMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtretail.SetFocus
         Case vbKeyEscape
            TxtResult.SetFocus
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
    TxtMRP.Text = Format(TxtMRP.Text, "0.00")
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            TXTEXPIRY.Visible = True
            TXTEXPIRY.SetFocus
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub


Private Function cost_calculate()
    Dim i As Integer
    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    LBLCOST.Caption = 0
    For i = 1 To grdsales.rows - 1
        If UCase(grdsales.TextMatrix(i, 8)) <> "SERVICE CHARGE" Then
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 4) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            With rstTRXMAST
                If Not (.EOF And .BOF) Then
                    LBLCOST.Caption = Val(LBLCOST.Caption) + IIf(IsNull(!ITEM_COST), 0, !ITEM_COST * Val(grdsales.TextMatrix(i, 2)) + Val(grdsales.TextMatrix(i, 9)))
                End If
            End With
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
        Else
            LBLCOST.Caption = Val(LBLCOST.Caption) + Val(grdsales.TextMatrix(i, 7))
        End If
    Next i
    If Val(TxtResult.Text) = 0 Then
        LBLCOST.Caption = 0
    Else
        LBLCOST.Caption = Round(Val(LBLCOST.Caption) / (Val(TxtResult.Text) * Val(Txtpack.Text)), 3)
    End If
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub txtTotalLoose_GotFocus()
    txtTotalLoose.SelStart = 0
    txtTotalLoose.SelLength = Len(txtTotalLoose.Text)

End Sub

Private Sub txtTotalLoose_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtTotalLoose.Text) = 0 Then Exit Sub
            TxtMRP.SetFocus
         Case vbKeyEscape
            TXTQTY.SetFocus
    End Select

End Sub

Private Sub txtTotalLoose_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub TxtLRate_GotFocus()
    TxtLRate.SelStart = 0
    TxtLRate.SelLength = Len(TxtLRate.Text)
End Sub

Private Sub TxtLRate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxttaxMRP.SetFocus
         Case vbKeyEscape
            txtvanrate.SetFocus
    End Select
End Sub

Private Sub TxtLRate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLRate_LostFocus()
    TxtLRate.Text = Format(TxtLRate.Text, "0.00")
End Sub

