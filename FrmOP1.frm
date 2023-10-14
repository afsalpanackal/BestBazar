VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPENING STOCK ENTRY"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18645
   ControlBox      =   0   'False
   Icon            =   "FrmOP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10995
   ScaleWidth      =   18645
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   60
      TabIndex        =   101
      Top             =   1680
      Visible         =   0   'False
      Width           =   14820
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   3855
         Left            =   30
         TabIndex        =   102
         Top             =   390
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6800
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
         TabIndex        =   104
         Top             =   15
         Width           =   10995
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
         TabIndex        =   103
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
      Left            =   1260
      TabIndex        =   62
      Top             =   105
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
      Left            =   4455
      TabIndex        =   35
      Top             =   7710
      Width           =   1200
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   2040
      TabIndex        =   50
      Top             =   1815
      Visible         =   0   'False
      Width           =   10320
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   4080
         Left            =   15
         TabIndex        =   51
         Top             =   15
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   7197
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
   End
   Begin VB.Frame Fram 
      BackColor       =   &H00B8F3E4&
      Caption         =   "Frame1"
      Height          =   11040
      Left            =   -135
      TabIndex        =   36
      Top             =   -90
      Width           =   18690
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   14640
         ScaleHeight     =   240
         ScaleWidth      =   1800
         TabIndex        =   139
         Top             =   435
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   14640
         ScaleHeight     =   240
         ScaleWidth      =   1965
         TabIndex        =   138
         Top             =   165
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   16470
         ScaleHeight     =   240
         ScaleWidth      =   855
         TabIndex        =   137
         Top             =   435
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   16455
         ScaleHeight     =   240
         ScaleWidth      =   1335
         TabIndex        =   136
         Top             =   165
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00B8F3E4&
         Height          =   600
         Left            =   135
         TabIndex        =   53
         Top             =   0
         Width           =   11265
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
            Left            =   11475
            TabIndex        =   61
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
            Height          =   330
            Left            =   7755
            MaxLength       =   100
            TabIndex        =   58
            Top             =   210
            Width           =   3450
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
            TabIndex        =   57
            Top             =   210
            Width           =   1260
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   5175
            TabIndex        =   60
            Top             =   210
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
         Begin VB.Label lblcredit 
            Height          =   525
            Left            =   11460
            TabIndex        =   78
            Top             =   645
            Visible         =   0   'False
            Width           =   630
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
            Left            =   6780
            TabIndex        =   59
            Top             =   225
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
            TabIndex        =   56
            Top             =   210
            Width           =   780
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "OP NO."
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
            TabIndex        =   55
            Top             =   210
            Width           =   870
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "OP. DATE"
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
            Left            =   4140
            TabIndex        =   54
            Top             =   225
            Width           =   1155
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdsales 
         Height          =   5265
         Left            =   150
         TabIndex        =   94
         Top             =   645
         Width           =   18495
         _ExtentX        =   32623
         _ExtentY        =   9287
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
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00B8F3E4&
         Height          =   5160
         Left            =   150
         TabIndex        =   37
         Top             =   5835
         Width           =   18480
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
            Height          =   450
            Left            =   15090
            TabIndex        =   142
            Top             =   1860
            Width           =   1335
         End
         Begin VB.TextBox TxtCessPer 
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
            Left            =   10860
            MaxLength       =   7
            TabIndex        =   140
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   11520
            MaxLength       =   7
            TabIndex        =   134
            Top             =   1140
            Width           =   1095
         End
         Begin VB.TextBox txtHSN 
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
            Left            =   16185
            MaxLength       =   15
            TabIndex        =   131
            Top             =   480
            Width           =   1530
         End
         Begin VB.TextBox TxtBarcode 
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
            Left            =   600
            MaxLength       =   20
            TabIndex        =   129
            Top             =   480
            Width           =   1920
         End
         Begin VB.TextBox TxtLWRate 
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
            Left            =   5565
            MaxLength       =   7
            TabIndex        =   22
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   9840
            MaxLength       =   7
            TabIndex        =   27
            Top             =   1140
            Width           =   1005
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
            Left            =   9210
            MaxLength       =   7
            TabIndex        =   26
            Top             =   1140
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
            Left            =   8250
            MaxLength       =   7
            TabIndex        =   25
            Top             =   1140
            Width           =   945
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
            Left            =   16455
            TabIndex        =   121
            Top             =   1545
            Width           =   1320
         End
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
            Left            =   16455
            TabIndex        =   120
            Top             =   1860
            Width           =   1335
         End
         Begin VB.TextBox TxtExpense 
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
            Left            =   2490
            MaxLength       =   7
            TabIndex        =   13
            Top             =   3870
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox txtcategory 
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
            Left            =   2535
            TabIndex        =   1
            Top             =   480
            Width           =   1260
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
            Left            =   14025
            MaxLength       =   4
            TabIndex        =   117
            Top             =   1140
            Width           =   330
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
            ItemData        =   "FrmOP.frx":030A
            Left            =   14355
            List            =   "FrmOP.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   1140
            Width           =   825
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
            Left            =   990
            MaxLength       =   7
            TabIndex        =   15
            Top             =   1530
            Width           =   885
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
            Height          =   420
            Left            =   1890
            MaxLength       =   7
            TabIndex        =   17
            Top             =   1515
            Width           =   930
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
            Height          =   420
            Left            =   2835
            MaxLength       =   7
            TabIndex        =   19
            Top             =   1515
            Width           =   990
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
            ItemData        =   "FrmOP.frx":0325
            Left            =   8805
            List            =   "FrmOP.frx":0368
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   480
            Width           =   960
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
            Left            =   8175
            MaxLength       =   7
            TabIndex        =   3
            Top             =   465
            Width           =   615
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
            Left            =   13935
            MaxLength       =   10
            TabIndex        =   9
            Top             =   465
            Width           =   1470
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
            Left            =   2835
            MaxLength       =   7
            TabIndex        =   18
            Top             =   1140
            Width           =   990
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
            Left            =   3840
            MaxLength       =   7
            TabIndex        =   20
            Top             =   1140
            Width           =   750
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
            Left            =   6585
            MaxLength       =   7
            TabIndex        =   23
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   7320
            MaxLength       =   7
            TabIndex        =   24
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   4605
            MaxLength       =   7
            TabIndex        =   21
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1890
            MaxLength       =   7
            TabIndex        =   16
            Top             =   1125
            Width           =   930
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
            Height          =   375
            Left            =   990
            MaxLength       =   7
            TabIndex        =   14
            Top             =   1140
            Width           =   885
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
            Left            =   8415
            MaxLength       =   7
            TabIndex        =   11
            Top             =   3765
            Visible         =   0   'False
            Width           =   825
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
            TabIndex        =   85
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
            Left            =   14970
            MaxLength       =   7
            TabIndex        =   83
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
            Left            =   6735
            TabIndex        =   80
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
            Left            =   7860
            TabIndex        =   79
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
            TabIndex        =   72
            Top             =   2205
            Width           =   945
         End
         Begin VB.OptionButton OPTNET 
            BackColor       =   &H00B8F3E4&
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
            Left            =   12630
            TabIndex        =   70
            Top             =   2145
            Value           =   -1  'True
            Width           =   675
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
            Left            =   7155
            MaxLength       =   7
            TabIndex        =   6
            Top             =   3645
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00B8F3E4&
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
            Left            =   12615
            TabIndex        =   68
            Top             =   1575
            Width           =   1410
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00B8F3E4&
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
            Left            =   12615
            TabIndex        =   69
            Top             =   1860
            Width           =   1395
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
            Left            =   15420
            MaxLength       =   7
            TabIndex        =   10
            Top             =   480
            Width           =   750
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
            Left            =   3915
            MaxLength       =   7
            TabIndex        =   64
            Top             =   2895
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
            Height          =   375
            Left            =   11640
            MaxLength       =   7
            TabIndex        =   8
            Top             =   465
            Width           =   1050
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
            Height          =   375
            Left            =   10665
            MaxLength       =   7
            TabIndex        =   7
            Top             =   465
            Width           =   960
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
            Left            =   60
            TabIndex        =   31
            Top             =   1980
            Width           =   1095
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
            Left            =   3810
            TabIndex        =   2
            Top             =   480
            Width           =   4350
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
            Left            =   9765
            MaxLength       =   8
            TabIndex        =   5
            Top             =   480
            Width           =   885
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
            Left            =   2280
            TabIndex        =   33
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
            Left            =   1185
            TabIndex        =   32
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
            TabIndex        =   39
            Top             =   2895
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   12630
            MaxLength       =   15
            TabIndex        =   29
            Top             =   1140
            Width           =   1380
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
            TabIndex        =   38
            Top             =   3945
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
            Left            =   3375
            TabIndex        =   34
            Top             =   1980
            Width           =   975
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   360
            Left            =   12705
            TabIndex        =   28
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
            Left            =   12705
            TabIndex        =   63
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
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8F3E4&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   3840
            TabIndex        =   95
            Top             =   1425
            Width           =   2745
            Begin VB.OptionButton OptComAmt 
               BackColor       =   &H00B8F3E4&
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
               Left            =   1350
               TabIndex        =   97
               Top             =   180
               Width           =   1335
            End
            Begin VB.OptionButton OptComper 
               BackColor       =   &H00B8F3E4&
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
               Left            =   60
               TabIndex        =   96
               Top             =   180
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8F3E4&
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   9030
            TabIndex        =   110
            Top             =   2295
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
               TabIndex        =   112
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
               TabIndex        =   111
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
               TabIndex        =   114
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
               TabIndex        =   113
               Top             =   195
               Width           =   1050
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8F3E4&
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   4305
            TabIndex        =   105
            Top             =   3750
            Visible         =   0   'False
            Width           =   2070
            Begin VB.OptionButton optdiscper 
               BackColor       =   &H00B8F3E4&
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
               TabIndex        =   107
               Top             =   120
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton Optdiscamt 
               BackColor       =   &H00B8F3E4&
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
               TabIndex        =   106
               Top             =   135
               Width           =   1125
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00B8F3E4&
            Height          =   2415
            Left            =   11610
            TabIndex        =   122
            Top             =   2295
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
            Left            =   10860
            TabIndex        =   141
            Top             =   885
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Adl. Cess Rate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   49
            Left            =   11520
            TabIndex        =   135
            Top             =   885
            Width           =   1095
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
            Left            =   2535
            TabIndex        =   133
            Top             =   195
            Width           =   1260
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
            Height          =   285
            Index           =   48
            Left            =   16185
            TabIndex        =   132
            Top             =   195
            Width           =   1530
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
            Left            =   600
            TabIndex        =   130
            Top             =   195
            Width           =   1920
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
            Left            =   5565
            TabIndex        =   128
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
            Left            =   4710
            TabIndex        =   127
            Top             =   2865
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
            Left            =   4710
            TabIndex        =   126
            Top             =   2610
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Trade Disc"
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
            Left            =   9840
            TabIndex        =   125
            Top             =   885
            Width           =   1005
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
            Left            =   9210
            TabIndex        =   124
            Top             =   885
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
            Left            =   8250
            TabIndex        =   123
            Top             =   885
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   45
            TabIndex        =   12
            Top             =   1140
            Width           =   930
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
            Left            =   2490
            TabIndex        =   119
            Top             =   3615
            Visible         =   0   'False
            Width           =   840
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
            Left            =   14025
            TabIndex        =   118
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
            Left            =   465
            TabIndex        =   115
            Top             =   1530
            Width           =   510
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
            TabIndex        =   109
            Top             =   2610
            Visible         =   0   'False
            Width           =   3300
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
            Left            =   8175
            TabIndex        =   108
            Top             =   195
            Width           =   1575
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
            Left            =   13935
            TabIndex        =   100
            Top             =   195
            Width           =   1470
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
            Left            =   2835
            TabIndex        =   99
            Top             =   885
            Width           =   990
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
            Left            =   3840
            TabIndex        =   98
            Top             =   885
            Width           =   750
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
            Left            =   6585
            TabIndex        =   93
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
            Left            =   7320
            TabIndex        =   92
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
            Left            =   4605
            TabIndex        =   91
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
            Left            =   1890
            TabIndex        =   90
            Top             =   885
            Width           =   930
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
            TabIndex        =   89
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
            Left            =   8415
            TabIndex        =   86
            Top             =   3510
            Visible         =   0   'False
            Width           =   825
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
            Left            =   990
            TabIndex        =   84
            Top             =   885
            Width           =   885
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
            Left            =   6765
            TabIndex        =   82
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
            Left            =   7860
            TabIndex        =   81
            Top             =   1965
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Net Amount"
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
            Left            =   11085
            TabIndex        =   77
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
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   570
            Left            =   10770
            TabIndex        =   76
            Top             =   1725
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
            Left            =   5745
            TabIndex        =   75
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
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   570
            Left            =   9045
            TabIndex        =   74
            Top             =   1725
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Amount"
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
            Left            =   9075
            TabIndex        =   73
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
            Left            =   7155
            TabIndex        =   71
            Top             =   3360
            Visible         =   0   'False
            Width           =   705
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
            Left            =   45
            TabIndex        =   67
            Top             =   885
            Width           =   930
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
            TabIndex        =   66
            Top             =   195
            Width           =   750
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
            Left            =   3915
            TabIndex        =   65
            Top             =   2610
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
            Left            =   11640
            TabIndex        =   52
            Top             =   195
            Width           =   1050
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
            TabIndex        =   49
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
            Left            =   3795
            TabIndex        =   48
            Top             =   195
            Width           =   4365
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
            Left            =   9765
            TabIndex        =   47
            Top             =   195
            Width           =   885
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
            Left            =   10665
            TabIndex        =   46
            Top             =   195
            Width           =   960
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
            Left            =   15195
            TabIndex        =   45
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
            TabIndex        =   44
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
            Left            =   12705
            TabIndex        =   43
            Top             =   195
            Width           =   1215
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
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   7
            Left            =   12630
            TabIndex        =   42
            Top             =   885
            Width           =   1380
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
            Left            =   15195
            TabIndex        =   30
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
            TabIndex        =   41
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
            TabIndex        =   40
            Top             =   3615
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   135
         TabIndex        =   88
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   705
         TabIndex        =   87
         Top             =   45
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytData() As Byte
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim PHY_CODE As New ADODB.Recordset
Dim PHYCODE_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT, M_ADD, OLD_BILL As Boolean
Dim PHY_PRERATE As New ADODB.Recordset
Dim PRERATE_FLAG As Boolean
Dim CHANGE_FLAG As Boolean

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = "1"
            TXTQTY.SetFocus
         Case vbKeyEscape
            'TXTUNIT.Text = ""
            Los_Pack.SetFocus
    End Select
End Sub

Private Sub CMDADD_Click()
        
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
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
    'Call TXTPTR_LostFocus
    'Call Txtgrossamt_LostFocus
    Call txtPD_LostFocus
    Call txtcrtn_GotFocus
    Call TxtLWRate_GotFocus
    
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim M_DATA As Long
    
    M_DATA = 0
    Txtpack.Text = 1
    If MDIMAIN.StatusBar.Panels(6).Text = "Y" Then If Trim(TxtBarcode.Text) = "" Then TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text)
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text) + Val(TxtFree.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = 1 'Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Val(Los_Pack.Text) ' 1 'Val(TxtPack.Text)
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Round(Val(TXTRATE.Text) / Val(Los_Pack.Text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTRATE.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Format(Round(((Val(LblGross.Caption) / (Val(Los_Pack.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)))) + ((Val(TxtExpense.Text) / Val(Los_Pack.Text)))), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Format(Round(Val(TXTPTR.Text) / Val(Los_Pack.Text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Format((Val(txtprofit.Text)), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = IIf(Val(TxttaxMRP.Text) = 0, "", Format(Val(TxttaxMRP.Text), ".00")) 'TAX
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = IIf(Trim(TXTEXPDATE.Text) = "/  /", "", TXTEXPDATE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Val(TxtFree.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = Val(txtPD.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = Format(Val(TXTRETAIL.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = Format(Val(txtWS.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = Format(Val(txtvanrate.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 26) = Format(Val(Txtgrossamt.Text), ".00")
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
    
    If OLD_BILL = False Then Call checklastbill
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!TRX_TYPE = "OP"
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
                RSTRTRXFILE!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
'                If UCase(rststock!CATEGORY) = "CUTSHEET" Then
'                Else
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
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)) <> 0 Then !P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) <> 0 Then !P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37)) <> 0 Then !P_LWS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) <> 0 Then !P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) <> 0 Then !CESS_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) <> 0 Then !CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) ' / Val(Los_Pack.Text), 3)
                
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
                RSTRTRXFILE!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
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
                !MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)) <> 0 Then !P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) <> 0 Then !P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37)) <> 0 Then !P_LWS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 37)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) <> 0 Then !P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) <> 0 Then !CESS_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) <> 0 Then !CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) ' / Val(Los_Pack.Text), 3)

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
    
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "OP"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE.Update
    End If
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
           
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    For i = 1 To grdsales.Rows - 1
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
    Next i
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    
    If MDIMAIN.StatusBar.Panels(6).Text = "Y" Then
        If MsgBox("Do you want to Print Barcode Labels now?", vbYesNo, "Purchase.....") = vbYes Then
            i = Val(InputBox("Enter number of lables to be print", "No. of labels..", Val(TXTQTY.Text) + Val(TxtFree.Text)))
            If i > 0 Then Call print_labels(i, Trim(TxtBarcode.Text), "")
        Else
            grdsales.TextMatrix(Val(TXTSLNO.Text), 41) = Val(InputBox("Enter number of lables to be print", "No. of labels..", Val(TXTQTY.Text) + Val(TxtFree.Text)))
        End If
    End If
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TxtBarcode.Text = ""
    TXTPTR.Text = ""
    Txtgrossamt.Text = ""
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TxtFree.Text = ""
    TxttaxMRP.Text = ""
    TxtExDuty.Text = ""
    TxtCSTper.Text = ""
    TxtTrDisc.Text = ""
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
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    OPTNET.value = True
    OptComper.value = True
    M_ADD = True
    Chkcancel.value = 0
    OLD_BILL = True
    txtcategory.Enabled = True
    txtBillNo.Enabled = False
    FRMEGRDTMP.Visible = False
    cmdRefresh.Enabled = True
    Los_Pack.Enabled = False
    CmbPack.Enabled = False
    TXTQTY.Enabled = False
    TxtFree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
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
    TxtComper.Enabled = False
    TxtComAmt.Enabled = False
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
        txtcategory.SetFocus
    End If
    M_EDIT = False
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TxtLWRate.SetFocus
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
   
    On Error GoTo Errhand
    db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & ""
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    With rststock
        If Not (.EOF And .BOF) Then
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
    
    i = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(LINE_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND  TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTRTRXFILE.EOF
        RSTRTRXFILE!line_no = i
        i = i + 1
        RSTRTRXFILE.Update
        RSTRTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    i = 1
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND  TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTRTRXFILE.EOF
        RSTRTRXFILE!line_no = i
        i = i + 1
        RSTRTRXFILE.Update
        RSTRTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    grdsales.Rows = 1
    i = 0
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    grdsales.Rows = 1
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTRTRXFILE.EOF
        grdsales.Rows = grdsales.Rows + 1
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
        grdsales.TextMatrix(i, 16) = RSTRTRXFILE!line_no
        grdsales.TextMatrix(i, 17) = IIf(IsNull(RSTRTRXFILE!P_DISC), 0, RSTRTRXFILE!P_DISC)
        grdsales.TextMatrix(i, 18) = IIf(IsNull(RSTRTRXFILE!P_RETAIL), 0, RSTRTRXFILE!P_RETAIL)
        grdsales.TextMatrix(i, 19) = IIf(IsNull(RSTRTRXFILE!P_WS), 0, RSTRTRXFILE!P_WS)
        grdsales.TextMatrix(i, 20) = IIf(IsNull(RSTRTRXFILE!P_CRTN), 0, RSTRTRXFILE!P_CRTN)
        grdsales.TextMatrix(i, 37) = IIf(IsNull(RSTRTRXFILE!P_LWS), 0, RSTRTRXFILE!P_LWS)
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
        grdsales.TextMatrix(i, 32) = IIf(IsNull(RSTRTRXFILE!EXPENSE), "", RSTRTRXFILE!EXPENSE)
        grdsales.TextMatrix(i, 33) = IIf(IsNull(RSTRTRXFILE!EXDUTY), "", RSTRTRXFILE!EXDUTY)
        grdsales.TextMatrix(i, 34) = IIf(IsNull(RSTRTRXFILE!CSTPER), "", RSTRTRXFILE!CSTPER)
        grdsales.TextMatrix(i, 35) = IIf(IsNull(RSTRTRXFILE!TR_DISC), "", RSTRTRXFILE!TR_DISC)
        grdsales.TextMatrix(i, 36) = IIf(IsNull(RSTRTRXFILE!GROSS_AMOUNT), "", RSTRTRXFILE!GROSS_AMOUNT)
        grdsales.TextMatrix(i, 38) = IIf(IsNull(RSTRTRXFILE!BARCODE), "", RSTRTRXFILE!BARCODE)
        grdsales.TextMatrix(i, 39) = IIf(IsNull(RSTRTRXFILE!CESS_AMT), "", RSTRTRXFILE!CESS_AMT)
        grdsales.TextMatrix(i, 40) = IIf(IsNull(RSTRTRXFILE!CESS_PER), "", RSTRTRXFILE!CESS_PER)
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    
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
    TxtFree.Text = ""
    TxttaxMRP.Text = ""
    TxtExDuty.Text = ""
    TxtCSTper.Text = ""
    TxtTrDisc.Text = ""
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
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_ADD = True
    OLD_BILL = True
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub CmdDeleteAll_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    On Error GoTo Errhand
    If Chkcancel.value = 0 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE ALL", vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then Exit Sub
    

    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    TXTREMARKS.Text = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTAL.Caption = ""
    TXTDISCAMOUNT.Text = ""
    txtaddlamt.Text = ""
    
    For i = 1 To grdsales.Rows - 1
        db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & ""
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        With rststock
            If Not (.EOF And .BOF) Then
                !RCPT_QTY = !RCPT_QTY - Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                !RCPT_VAL = !RCPT_VAL - Val(grdsales.TextMatrix(i, 13))
                
                !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(i, 13))
                rststock.Update
            End If
        End With
        db.CommitTrans
        rststock.Close
        Set rststock = Nothing
    Next i
    
    grdsales.FixedRows = 0
    grdsales.Rows = 1
    Call appendpurchase
    Screen.MousePointer = vbNormal
    Exit Sub
Errhand:
    MsgBox Err.Description
    
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    Los_Pack.Enabled = True
    CmbPack.Enabled = True
    TXTQTY.Enabled = True
    TxtFree.Enabled = True
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
    TXTQTY.SetFocus
    
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
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
            TxtFree.Text = ""
            TxttaxMRP.Text = ""
            TxtExDuty.Text = ""
            TxtCSTper.Text = ""
            TxtTrDisc.Text = ""
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

Private Sub cmdRefresh_Click()
    On Error GoTo Errhand
    If grdsales.Rows <= 1 Then
        lblcredit.Caption = "0"
        Call appendpurchase
    Else
        If Not IsDate(TXTINVDATE.Text) Then
            MsgBox "Enter Opening Stock Date", vbOKOnly, "EzBiz"
            Exit Sub
        End If
        'Me.Enabled = False
        'MDIMAIN.cmdpurchase.Enabled = False
        'Set creditbill = Me
        'frmCREDIT.Show
        Call appendpurchase
    End If
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub cmdRefresh_GotFocus()
    FRMEGRDTMP.Visible = False
End Sub

Private Sub Form_Activate()
    On Error GoTo Errhand
    txtBillNo.SetFocus
    Exit Sub
Errhand:
    If Err.Number = 5 Then Exit Sub
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo Errhand
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'OP'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    PRERATE_FLAG = True
    OLD_BILL = False
    grdsales.ColWidth(0) = 500
    grdsales.ColWidth(1) = 0
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
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(17) = 800
    grdsales.ColWidth(18) = 800
    grdsales.ColWidth(19) = 800
    grdsales.ColWidth(20) = 800
    grdsales.ColWidth(37) = 800
    grdsales.ColWidth(21) = 0
    grdsales.ColWidth(22) = 0
    grdsales.ColWidth(23) = 0
    grdsales.ColWidth(24) = 800
    grdsales.ColWidth(25) = 0
    grdsales.ColWidth(26) = 1700
    grdsales.ColWidth(27) = 0
    grdsales.ColWidth(28) = 1100
    grdsales.ColWidth(29) = 0
    grdsales.ColWidth(30) = 0
    grdsales.ColWidth(31) = 0
    
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
    grdsales.TextArray(24) = "Loose Pck"
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
    M_ADD = False
    'Me.Width = 15135
    'Me.Height = 9660
    Me.Left = 0
    Me.Top = 0
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If PHYCODE_FLAG = False Then PHY_CODE.Close
        If PRERATE_FLAG = False Then PHY_PRERATE.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo Errhand
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
                    Set Image1.DataSource = PHY 'setting image1’s datasource
                    Image1.DataField = "PHOTO"
                    bytData = PHY!PHOTO
                End If
            End If
            On Error GoTo Errhand
            For i = 1 To grdsales.Rows - 1
                If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                    Exit For
                End If
            Next i
            
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                'RSTRXFILE.MoveLast
                TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
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
                    TXTPTR.Text = Format(Round(Val(RSTRXFILE!PTR), 3), ".000")
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
                On Error GoTo Errhand
                
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
                    On Error GoTo Errhand
                    
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.Text = ""
                    TXTEXPIRY.Text = "  /  "
                    TXTRATE.Text = ""
                    txtmrpbt.Text = ""
                    TXTPTR.Text = ""
                    TXTRETAIL.Text = ""
                    txtWS.Text = ""
                    txtvanrate.Text = ""
                    txtcrtn.Text = ""
                    TxtLWRate.Text = ""
                    txtcrtnpack.Text = ""
                    txtprofit.Text = ""
                    TxttaxMRP.Text = "5"
                    Los_Pack.Text = "1"
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo Errhand
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
            Los_Pack.Enabled = True
            Los_Pack.SetFocus
            'TxtPack.Enabled = True
            'TxtPack.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TxtFree.Text = ""
            Fram.Enabled = True
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
Errhand:
    MsgBox Err.Description
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

Private Sub txtbarcode_GotFocus()
    TxtBarcode.SelStart = 0
    TxtBarcode.SelLength = Len(TxtBarcode.Text)
    FRMEGRDTMP.Visible = False
    TXTSLNO.Enabled = False
    CmbPack.Enabled = False
    TXTQTY.Enabled = False
    TxtFree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
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

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(txtBatch.Text) = "" Then Exit Sub
            TxtWarranty.SetFocus
        Case vbKeyEscape
            TxtTrDisc.SetFocus
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
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTDIST As ADODB.Recordset
    Dim RSTTRNSMAST As ADODB.Recordset
    Dim i As Long

    On Error GoTo Errhand
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
            i = 0
            LBLTOTAL.Caption = ""
            lbltotalwodiscount = ""
            grdsales.Rows = 1
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until rstTRXMAST.EOF
                grdsales.Rows = grdsales.Rows + 1
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
                grdsales.TextMatrix(i, 16) = rstTRXMAST!line_no
                grdsales.TextMatrix(i, 17) = IIf(IsNull(rstTRXMAST!P_DISC), 0, rstTRXMAST!P_DISC)
                grdsales.TextMatrix(i, 18) = IIf(IsNull(rstTRXMAST!P_RETAIL), 0, rstTRXMAST!P_RETAIL)
                grdsales.TextMatrix(i, 19) = IIf(IsNull(rstTRXMAST!P_WS), 0, rstTRXMAST!P_WS)
                grdsales.TextMatrix(i, 20) = IIf(IsNull(rstTRXMAST!P_CRTN), 0, rstTRXMAST!P_CRTN)
                grdsales.TextMatrix(i, 37) = IIf(IsNull(rstTRXMAST!P_LWS), 0, rstTRXMAST!P_LWS)
                
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
                grdsales.TextMatrix(i, 32) = IIf(IsNull(rstTRXMAST!EXPENSE), "", rstTRXMAST!EXPENSE)
                grdsales.TextMatrix(i, 33) = IIf(IsNull(rstTRXMAST!EXDUTY), "", rstTRXMAST!EXDUTY)
                grdsales.TextMatrix(i, 34) = IIf(IsNull(rstTRXMAST!CSTPER), "", rstTRXMAST!CSTPER)
                grdsales.TextMatrix(i, 35) = IIf(IsNull(rstTRXMAST!TR_DISC), "", rstTRXMAST!TR_DISC)
                grdsales.TextMatrix(i, 36) = IIf(IsNull(rstTRXMAST!GROSS_AMOUNT), "", rstTRXMAST!GROSS_AMOUNT)
                grdsales.TextMatrix(i, 38) = IIf(IsNull(rstTRXMAST!BARCODE), "", rstTRXMAST!BARCODE)
                grdsales.TextMatrix(i, 39) = IIf(IsNull(rstTRXMAST!CESS_AMT), "", rstTRXMAST!CESS_AMT)
                grdsales.TextMatrix(i, 40) = IIf(IsNull(rstTRXMAST!CESS_PER), "", rstTRXMAST!CESS_PER)
                lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
                On Error Resume Next
                TXTINVDATE.Text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                On Error GoTo Errhand
                TXTREMARKS.Text = IIf(IsNull(rstTRXMAST!VCH_DESC), "", Mid(rstTRXMAST!VCH_DESC, 15))
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                TXTDISCAMOUNT.Text = IIf(IsNull(rstTRXMAST!DISCOUNT), "", Format(rstTRXMAST!DISCOUNT, ".00"))
                txtaddlamt.Text = IIf(IsNull(rstTRXMAST!ADD_AMOUNT), "", Format(rstTRXMAST!ADD_AMOUNT, ".00"))
                txtcramt.Text = IIf(IsNull(rstTRXMAST!DISC_PERS), "", Format(rstTRXMAST!DISC_PERS, ".00"))
                TxtCST.Text = IIf(IsNull(rstTRXMAST!CST_PER), "", Format(rstTRXMAST!CST_PER, ".00"))
                TxtInsurance.Text = IIf(IsNull(rstTRXMAST!INS_PER), "", Format(rstTRXMAST!INS_PER, ".00"))
                If rstTRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                On Error Resume Next
                TXTINVDATE.Text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                TXTDATE.Text = Format(rstTRXMAST!CREATE_DATE, "DD/MM/YYYY")
                On Error GoTo Errhand
                OLD_BILL = True
            Else
                TXTDATE.Text = Format(Date, "DD/MM/YYYY")
                OLD_BILL = False
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            ''''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
            'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
            LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
            
            TXTSLNO.Text = grdsales.Rows
            TXTSLNO.Enabled = True
            txtBillNo.Enabled = False
            FRMEMASTER.Enabled = True
            If i > 0 Or (Val(txtBillNo.Text) < Val(TXTLASTBILL.Text)) Then
                FRMEMASTER.Enabled = True
                FRMECONTROLS.Enabled = True
                cmdRefresh.Enabled = True
                cmdRefresh.SetFocus
            Else
                TXTINVDATE.SetFocus
            End If
            
'            Set RSTTRNSMAST = New ADODB.Recordset
'            RSTTRNSMAST.Open "Select CHECK_FLAG From TRANSMAST WHERE TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
'            If Not (RSTTRNSMAST.EOF Or RSTTRNSMAST.BOF) Then
'                If RSTTRNSMAST!CHECK_FLAG = "Y" Then FRMEMASTER.Enabled = False
'            End If
'            RSTTRNSMAST.Close
'            Set RSTTRNSMAST = Nothing
    
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
    If Val(txtBillNo.Text) = 0 Or Val(txtBillNo.Text) > Val(TXTLASTBILL.Text) Then txtBillNo.Text = TXTLASTBILL.Text
End Sub

Private Sub txtcategory_Change()
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo Errhand
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select * From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         Else
             PHY.Close
             'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHY.Open "Select * From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
Errhand:
        MsgBox Err.Description
End Sub

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.Text)
    FRMEGRDTMP.Visible = False
    'TXTSLNO.Enabled = False
    CmbPack.Enabled = False
    TXTQTY.Enabled = False
    TxtFree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
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

Private Sub TxtCSTper_LostFocus()
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TxtExDuty_LostFocus()
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TXTEXPDATE_GotFocus()
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
            TXTRATE.Enabled = True
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
            TXTRATE.SetFocus
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
    TXTEXPDATE.Text = Format(TXTEXPDATE.Text, "DD/MM/YYYY")
    If IsDate(TXTEXPDATE.Text) Then TXTEXPIRY.Text = Format(TXTEXPDATE.Text, "MM/YY")
End Sub

Private Sub TxtFree_GotFocus()
    TxtFree.SelStart = 0
    TxtFree.SelLength = Len(TxtFree.Text)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRATE.SetFocus
        Case vbKeyEscape
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
    If Val(TxtFree.Text) = 0 Then TxtFree.Text = 0
    TxtFree.Text = Format(TxtFree.Text, "0.00")
End Sub

Private Sub TXTINVDATE_GotFocus()
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
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
            End If
        Case vbKeyEscape
            If M_ADD = False Then
                FRMECONTROLS.Enabled = False
                FRMEMASTER.Enabled = False
                txtBillNo.Enabled = True
                txtBillNo.SetFocus
            End If
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
    Dim i As Long
    On Error GoTo Errhand
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select * From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.Text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         Else
             PHY.Close
             'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHY.Open "Select * From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.Text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
Errhand:
        MsgBox Err.Description
                
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
    If Trim(TXTPRODUCT.Text) <> "" Or Trim(txtcategory.Text) <> "" Then Call TXTPRODUCT_Change
    'TXTSLNO.Enabled = False
    CmbPack.Enabled = False
    TXTQTY.Enabled = False
    TxtFree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
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
    Dim RSTRXFILE, RSTITEMMAST  As ADODB.Recordset
    Dim i As Long
    On Error GoTo Errhand
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
        Case vbKeyReturn
            
            On Error Resume Next
            TXTITEMCODE.Text = ""
            TXTITEMCODE.Text = grdtmp.Columns(0)
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If Trim(TXTITEMCODE.Text) = "" Then
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
                RSTITEMMAST.Update
                db.CommitTrans
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                TXTITEMCODE.Text = TXTPRODUCT.Tag
                Call TxtItemcode_KeyDown(13, 0)
                'frmitemmaster.Show
                'frmitemmaster.TXTITEM.Text = Trim(TXTPRODUCT.Text)
                'frmitemmaster.LBLLP.Caption = "P"
                'MsgBox "Item not found!!!!", , "EzBiz"
                Exit Sub
            Else
                Call TxtItemcode_KeyDown(13, 0)
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
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' and ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                        Exit For
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    'RSTRXFILE.MoveLast
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.Text = 1
                    Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    TxtWarranty.Text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.Text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo Errhand
                    
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
                        TXTPTR.Text = Format(Round(Val(RSTRXFILE!PTR), 3), ".000")
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
                    On Error GoTo Errhand
                
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
                    On Error GoTo Errhand
                    
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.Text = ""
                    txtHSN.Text = ""
                    TXTEXPIRY.Text = "  /  "
                    TXTRATE.Text = ""
                    txtmrpbt.Text = ""
                    TXTPTR.Text = ""
                    TXTRETAIL.Text = ""
                    txtWS.Text = ""
                    txtvanrate.Text = ""
                    txtcrtn.Text = ""
                    TxtLWRate.Text = ""
                    txtcrtnpack.Text = ""
                    txtprofit.Text = ""
                    TxttaxMRP.Text = "5"
                    TxtExDuty.Text = ""
                    TxtCSTper.Text = ""
                    TxtTrDisc.Text = ""
                    TxtCessPer.Text = ""
                    txtCess.Text = ""
                    Los_Pack.Text = "1"
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo Errhand
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
Errhand:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPTR_GotFocus()
    TXTPTR.SelStart = 0
    TXTPTR.SelLength = Len(TXTPTR.Text)
    Call FILL_PREVIIOUSRATE
End Sub

Private Sub TXTPTR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            TXTEXPIRY.Visible = True
            TXTEXPIRY.SetFocus
        Case vbKeyEscape
            TXTRATE.SetFocus
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
    Txtgrossamt.Text = Val(TXTPTR.Text) * Val(TXTQTY.Text)
    TXTPTR.Text = Format(TXTPTR.Text, ".000")
    'TXTRETAIL.Text = Round(Val(txtmrpbt.Text) * 0.8, 2)
'    txtretail.Text = Format(Round(Val(TXTRATE.Text) - (Val(txtmrpbt.Text) * 20 / 100), 3), ".000")
'    txtprofit.Text = Format(Round(Val(txtretail.Text) - Val(txtretail.Text) * 10 / 100, 3), ".000")
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
    FRMEGRDTMP.Visible = False
    CmbPack.Enabled = True
    Los_Pack.Enabled = True
    TXTQTY.Enabled = True
    TxtFree.Enabled = True
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
        Else
            txtHSN.Text = ""
        End If
    End With
    rststock.Close
    Set rststock = Nothing
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            TXTRATE.SetFocus
        Case vbKeyEscape
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
    TXTQTY.Text = Format(TXTQTY.Text, ".00")
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTPTR.Text), 2)), ".000")
    LblGross.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTPTR.Text), 2)), ".000")
End Sub

Private Sub TXTRATE_Change()
    If Val(TXTRATE.Text) > 0 Then TXTRETAIL.Text = Val(TXTRATE.Text)
End Sub

Private Sub TXTRATE_GotFocus()
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
            TXTQTY.SetFocus
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
    TXTRATE.Text = Format(TXTRATE.Text, ".000")
    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105 '(100 + Val(TxttaxMRP.Text))
End Sub

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errhand
    Select Case KeyCode
        Case vbKeyReturn
            If txtBillNo.Text = "" Then Exit Sub
            'If TXTINVOICE.Text = "" Then Exit Sub
            If Not IsDate(TXTINVDATE.Text) Then Exit Sub
            FRMECONTROLS.Enabled = True
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTINVDATE.SetFocus
    End Select
    Exit Sub
Errhand:
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

Private Sub TxtRetailPercent_GotFocus()
    TxtRetailPercent.SelStart = 0
    TxtRetailPercent.SelLength = Len(TxtRetailPercent.Text)
End Sub

Private Sub TxtRetailPercent_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            txtWS.SetFocus
         Case vbKeyEscape
            TXTRETAIL.SetFocus
    End Select
End Sub

Private Sub TxtRetailPercent_LostFocus()
    On Error Resume Next
    If optdiscper.value = True Then
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) + ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
    Else
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) + ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
    End If
    TXTRETAIL.Text = Round((Val(TXTPTR.Tag) * Val(TxtRetailPercent.Text) / 100) + Val(TXTPTR.Tag), 2)
    'txtretail.Text = Round(((Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) + Val(TXTPTR.Text)) * Val(TxtRetailPercent.Text) / 100 + ((Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) + Val(TXTPTR.Text)), 2)
    TXTRETAIL.Text = Format(Val(TXTRETAIL.Text), "0.0000")
    
End Sub

Private Sub txtSchPercent_GotFocus()
    txtSchPercent.SelStart = 0
    txtSchPercent.SelLength = Len(txtSchPercent.Text)
End Sub

Private Sub txtSchPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcrtnpack.SetFocus
         Case vbKeyEscape
            txtvanrate.SetFocus
    End Select
End Sub

Private Sub txtSchPercent_LostFocus()
    On Error Resume Next
    If optdiscper.value = True Then
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) + ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
    Else
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) + ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
    End If
    txtvanrate.Text = Round((Val(TXTPTR.Tag) * Val(txtSchPercent.Text) / 100) + Val(TXTPTR.Tag), 2)
    txtvanrate.Text = Format(Val(txtvanrate.Text), "0.000")
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    TxtBarcode.Enabled = False
    txtcategory.Enabled = False
    TXTPRODUCT.Enabled = False
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
                TXTPTR.Text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 9)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)), 3), "0.000")
                txtprofit.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7)), "0.00")
                TXTRETAIL.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)), "0.00")
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
                TxtFree.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
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
                Los_Pack.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 28)
                CmbPack.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 29)
                TxtWarranty.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 30)
                CmbWrnty.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 31)
                TxtExpense.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32))
                TxtExDuty.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 33))
                TxtCSTper.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 34))
                TxtTrDisc.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 35))
                LblGross.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 36))
                TxtBarcode.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 38))
                txtCess.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39))
                TxtCessPer.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40))
                FRMEGRDTMP.Visible = False
                
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
                TxtFree.Text = ""
                TxttaxMRP.Text = ""
                TxtExDuty.Text = ""
                TxtCSTper.Text = ""
                TxtTrDisc.Text = ""
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
            TXTPTR.Enabled = True
            TXTEXPDATE.Enabled = False
            TXTPTR.SetFocus
                        
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
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

Private Sub TxttaxMRP_GotFocus()
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
                txtHSN.SetFocus
            Else
                TXTRETAIL.SetFocus
            End If
         Case vbKeyEscape
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
    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / (100 + Val(TxttaxMRP.Text))
    Txtgrossamt.Tag = Val(Txtgrossamt.Text) + (Val(Txtgrossamt.Text) * Val(TxtExDuty.Text) / 100)
    Txtgrossamt.Tag = Val(Txtgrossamt.Tag) + (Val(Txtgrossamt.Text) * Val(TxtCSTper.Text) / 100)
    'Txtgrossamt.Tag = Val(Txtgrossamt.Tag) + Val(txtCess.Text)
    If Val(TxttaxMRP.Text) = 0 Then
        
        TxttaxMRP.Text = 0
        lbltaxamount.Caption = 0
        lbltaxamount.Caption = ""
        If optdiscper.value = True Then
            LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag)) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.Text) / 100)
            LblGross.Caption = (Val(Txtgrossamt.Tag)) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.Text) / 100)
        Else
            LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) - Val(txtPD.Text))
            LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(txtPD.Text))
        End If
    Else
        If OPTTaxMRP.value = True Then
            lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100
            If optdiscper.value = True Then
                LBLSUBTOTAL.Caption = (Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption)
                LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - (Val(LBLSUBTOTAL.Caption) * Val(txtPD.Text) / 100)
            Else
                LBLSUBTOTAL.Caption = (Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption)
                LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - Val(txtPD.Text)
            End If
            LblGross.Caption = LBLSUBTOTAL.Caption
        ElseIf OPTVAT.value = True Then
           If optdiscper.value = True Then
                lbltaxamount.Caption = Round((Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.Text) / 100)
                LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - (Val(LBLSUBTOTAL.Caption) - Val(lbltaxamount.Caption)) * Val(TxtTrDisc.Text) / 100
                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.Text) / 100)
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.Text) / 100)
            Else
                lbltaxamount.Caption = Round((Val(Txtgrossamt.Tag) - Val(txtPD.Text)) * Val(TxttaxMRP.Text) / 100, 2)
                LBLSUBTOTAL.Caption = Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption) - Val(txtPD.Text)
                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.Text)
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.Text) / 100)
            End If
            LBLSUBTOTAL.Caption = LBLSUBTOTAL.Caption + (Val(LblGross.Caption) * Val(TxtCessPer.Text) / 100)
            LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) + (Val(txtCess.Text) * Val(TXTQTY.Text))
        Else
            TxttaxMRP.Text = 0
            lbltaxamount.Caption = 0
            lbltaxamount.Caption = ""
            If optdiscper.value = True Then
                LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag)) - Val(txtPD.Text)
            Else
                LBLSUBTOTAL.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.Text)
            End If
            LblGross.Caption = LBLSUBTOTAL.Caption
        End If
    End If
    'LBLSUBTOTAL.Caption = Round(Val(LBLSUBTOTAL.Caption) + Val(txtCess.Text), 2)
    LBLSUBTOTAL.Caption = Format(Round(LBLSUBTOTAL.Caption, 3), "0.00")
    LblGross.Caption = Format(LblGross.Caption, "0.00")
    TxttaxMRP.Text = Format(TxttaxMRP.Text, "0.00")
    lbltaxamount.Caption = Format(lbltaxamount.Caption, "0.00")
End Sub

Private Sub TxtTrDisc_LostFocus()
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TXTUNIT_GotFocus()
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
            TxtFree.Text = ""
            TxttaxMRP.Text = ""
            TxtExDuty.Text = ""
            TxtCSTper.Text = ""
            TxtTrDisc.Text = ""
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
            Txtgrossamt.Text = ""
            TXTEXPDATE.Text = "  /  /    "
            TXTEXPIRY.Text = "  /  "
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            lbltaxamount.Caption = ""
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
    Dim DISC As Currency
    
    On Error GoTo Errhand
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    ''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
Errhand:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    TXTDISCAMOUNT.SetFocus
End Sub

Private Sub TXTDISCAMOUNT_GotFocus()
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
    Dim i As Long
    
    On Error GoTo Errhand
    Screen.MousePointer = vbHourglass
    
    If OLD_BILL = False Then Call checklastbill
    db.Execute "delete From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    If grdsales.Rows = 1 Then GoTo SKIP
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        db.BeginTrans
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "OP"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = ""
        RSTTRXFILE!ACT_NAME = "OP Stock"
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
        If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
        RSTTRXFILE!CFORM_NO = ""
        RSTTRXFILE!CFORM_DATE = Date
        RSTTRXFILE!Remarks = ""
        RSTTRXFILE!DISC_PERS = Val(txtcramt.Text)
        RSTTRXFILE!CST_PER = Val(TxtCST.Text)
        RSTTRXFILE!INS_PER = Val(TxtInsurance.Text)
        RSTTRXFILE!LETTER_NO = 0
        RSTTRXFILE!LETTER_DATE = Date
        RSTTRXFILE!INV_MSGS = ""
        If Not IsDate(TXTDATE.Text) Then TXTDATE.Text = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!CREATE_DATE = Format(TXTDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!PINV = ""
        RSTTRXFILE.Update
        db.CommitTrans
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
            
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='OP' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(Trim(TXTINVDATE.Text), "dd/mm/yyyy")
        RSTTRXFILE!VCH_DESC = "Received From OP Stock"
        RSTTRXFILE!PINV = ""
        RSTTRXFILE!M_USER_ID = ""
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

SKIP:
    
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'OP'", db, adOpenStatic, adLockReadOnly
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
    TxtFree.Text = ""
    TxttaxMRP.Text = ""
    TxtExDuty.Text = ""
    TxtCSTper.Text = ""
    TxtTrDisc.Text = ""
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
    Txtgrossamt.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    txtaddlamt.Text = ""
    txtcramt.Text = ""
    TxtInsurance.Text = ""
    TxtCST.Text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    TXTDISCAMOUNT.Text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    lbldealer.Caption = ""
    grdsales.Rows = 1
    CMDEXIT.Enabled = True
    OptComper.value = True
    txtBillNo.SetFocus
    M_ADD = False
    OLD_BILL = False
    Chkcancel.value = 0
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "EzBiz"
    Exit Sub
Errhand:
    Screen.MousePointer = vbNormal
    If Err.Number = 7 Then
        MsgBox "Select Supplier from the list", vbOKOnly, "EzBiz"
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub txtaddlamt_GotFocus()
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
    Dim DISC As Currency
    
    On Error GoTo Errhand
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
Errhand:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    txtaddlamt.SetFocus
End Sub

Private Sub txtcramt_GotFocus()
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
    Dim DISC As Currency
    
    On Error GoTo Errhand
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    Exit Sub
Errhand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    txtcramt.SetFocus
End Sub

Private Sub OPTTaxMRP_GotFocus()
    'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100
    'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
    lbltaxamount.Caption = ((Val(TXTRATE.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)) * 55 / 100)) * Val(TxttaxMRP.Text) / 100
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
    lbltaxamount.Caption = ""
    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text), ".000")
    LblGross.Caption = Format(Val(Txtgrossamt.Text), ".000")
End Sub

Private Sub txtprofit_GotFocus()
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
    txtprofit.Text = Format(txtprofit.Text, "0.00")
End Sub

Private Sub txtPD_GotFocus()
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
            TxtExpense.SetFocus
         Case vbKeyEscape
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


Private Sub TXTRETAIL_GotFocus()
    TXTRETAIL.SelStart = 0
    TXTRETAIL.SelLength = Len(TXTRETAIL.Text)
    If Val(TXTRETAIL.Text) = 0 Then TXTRETAIL.Text = Val(TXTRATE.Text)
    Call FILL_PREVIIOUSRATE
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
            If Trim(txtHSN.Text) = "" Then
                txtHSN.SetFocus
            Else
                TxttaxMRP.SetFocus
            End If
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
    On Error Resume Next
    TXTRETAIL.Text = Format(TXTRETAIL.Text, "0.00")
    If optdiscper.value = True Then
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) + ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
    Else
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) + ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
    End If
    If Val(Val(TXTPTR.Tag)) <> 0 Then
        TxtRetailPercent.Text = Round(((Val(TXTRETAIL.Text) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
        TxtRetailPercent.Text = Format(Val(TxtRetailPercent.Text), "0.00")
    Else
         TxtRetailPercent.Text = Round(((Val(TXTRETAIL.Text) - Val(TXTPTR.Tag)) * 100), 2)
        TxtRetailPercent.Text = Format(Val(TxtRetailPercent.Text), "0.00")
    End If
End Sub

Private Sub txtws_GotFocus()
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
    On Error Resume Next
    txtWS.Text = Format(txtWS.Text, "0.00")
    If optdiscper.value = True Then
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) + ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
    Else
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) + ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
    End If
    If Val(Val(TXTPTR.Tag)) <> 0 Then
        txtWsalePercent.Text = Round(((Val(txtWS.Text) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
        txtWsalePercent.Text = Format(Val(txtWsalePercent.Text), "0.00")
    Else
         txtWsalePercent.Text = Round(((Val(txtWS.Text) - Val(TXTPTR.Tag)) * 100), 2)
        txtWsalePercent.Text = Format(Val(txtWsalePercent.Text), "0.00")
    End If
End Sub

Private Sub txtcrtn_GotFocus()
    If Val(txtcrtnpack.Text) = 0 Then txtcrtnpack.Text = "1"
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
    txtcrtn.Text = Format(txtcrtn.Text, "0.00")
End Sub

Private Sub TxtComper_GotFocus()
    TxtComper.SelStart = 0
    TxtComper.SelLength = Len(TxtComper.Text)
    OptComper.value = True
End Sub

Private Sub TxtComper_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtExDuty.SetFocus
         Case vbKeyEscape
            txtvanrate.SetFocus
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
    TxtComper.Text = Format(TxtComper.Text, "0.00")
End Sub

Private Sub TxtComAmt_GotFocus()
    TxtComAmt.SelStart = 0
    TxtComAmt.SelLength = Len(TxtComAmt.Text)
    OptComAmt.value = True
End Sub

Private Sub TxtComAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtExDuty.SetFocus
         Case vbKeyEscape
            txtvanrate.SetFocus
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
    TxtComAmt.Text = ""
    TxtComAmt.Enabled = False
    TxtComper.Enabled = True
    TxtComper.SetFocus
End Sub

Private Sub txtcrtnpack_GotFocus()
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
    txtcrtnpack.Text = Format(txtcrtnpack.Text, "0.00")
End Sub

Private Sub txtvanrate_GotFocus()
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
    On Error Resume Next
    txtvanrate.Text = Format(txtvanrate.Text, "0.00")
    If optdiscper.value = True Then
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) + ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
    Else
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) + ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
    End If
    If Val(Val(TXTPTR.Tag)) <> 0 Then
        txtSchPercent.Text = Round(((Val(txtvanrate.Text) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
        txtSchPercent.Text = Format(Val(txtSchPercent.Text), "0.00")
    Else
        txtSchPercent.Text = Round(((Val(txtvanrate.Text) - Val(TXTPTR.Tag)) * 100), 2)
        txtSchPercent.Text = Format(Val(txtSchPercent.Text), "0.00")
    End If
End Sub

Private Sub Txtgrossamt_GotFocus()
    Txtgrossamt.SelStart = 0
    Txtgrossamt.SelLength = Len(Txtgrossamt.Text)
End Sub

Private Sub Txtgrossamt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(Txtgrossamt.Text) = 0 Then Exit Sub
            TxttaxMRP.SetFocus
        Case vbKeyEscape
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
    If Val(Txtgrossamt.Text) <> 0 Then
        Txtgrossamt.Text = Format(Txtgrossamt.Text, ".000")
        If Val(TXTQTY.Text) <> 0 Then
            TXTPTR.Text = Format(Round(Val(Txtgrossamt.Text) / Val(TXTQTY.Text), 2), "0.00")
        ElseIf Val(TXTPTR.Text) <> 0 Then
            TXTQTY.Text = Format(Round(Val(Txtgrossamt.Text) / Val(TXTPTR.Text), 2), "0.00")
        End If
    End If
    Call TxttaxMRP_LostFocus
End Sub

Function FILL_PREVIIOUSRATE()
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, ITEM_COST_PRICE, ITEM_NET_COST_PRICE, P_RETAIL, P_WS, VCH_NO, VCH_DATE, VCH_DESC  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, ITEM_COST_PRICE, ITEM_NET_COST_PRICE, P_RETAIL, P_WS, VCH_NO, VCH_DATE, VCH_DESC  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
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
    Los_Pack.SelStart = 0
    Los_Pack.SelLength = Len(Los_Pack.Text)
    FRMEGRDTMP.Visible = False
    CmbPack.Enabled = True
    TXTQTY.Enabled = True
    TxtFree.Enabled = True
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
End Sub

Private Sub Los_Pack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            CmbPack.SetFocus
         Case vbKeyEscape
             If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            Los_Pack.Enabled = False
            TXTPRODUCT.Enabled = True
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
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo Errhand
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
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            Else
                PHY_CODE.Close
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' and ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
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
                        Set Image1.DataSource = PHY 'setting image1’s datasource
                        Image1.DataField = "PHOTO"
                        bytData = PHY!PHOTO
                    End If
                End If
                On Error GoTo Errhand
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                        Exit For
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    'RSTRXFILE.MoveLast
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.Text = 1
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.Text = IIf(IsNull(RSTRXFILE!REF_NO), "", RSTRXFILE!REF_NO)
                    TXTEXPIRY.Text = IIf(IsDate(RSTRXFILE!EXP_DATE), Format(RSTRXFILE!EXP_DATE, "MM/YY"), "  /  ")
                    Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
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
                    On Error GoTo Errhand
                    
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
                    On Error GoTo Errhand
                    
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.Text = ""
                    txtHSN.Text = ""
                    TXTEXPIRY.Text = "  /  "
                    TXTRATE.Text = ""
                    txtmrpbt.Text = ""
                    TXTPTR.Text = ""
                    TXTRETAIL.Text = ""
                    txtWS.Text = ""
                    txtvanrate.Text = ""
                    txtcrtn.Text = ""
                    TxtLWRate.Text = ""
                    txtcrtnpack.Text = ""
                    txtprofit.Text = ""
                    TxttaxMRP.Text = "5"
                    TxtExDuty.Text = ""
                    TxtCSTper.Text = ""
                    TxtTrDisc.Text = ""
                    TxtCessPer.Text = ""
                    txtCess.Text = ""
                    Los_Pack.Text = "1"
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo Errhand
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
                    TXTITEMCODE.Enabled = False
                    TXTPRODUCT.Enabled = False
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

Private Sub txtcst_GotFocus()
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
    Dim DISC As Currency
    
    On Error GoTo Errhand
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtCST.Text)), 0), ".00")
    Exit Sub
Errhand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtCST.SetFocus
End Sub

Private Sub TxtInsurance_GotFocus()
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
    Dim DISC As Currency
    
    On Error GoTo Errhand
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtInsurance.Text)), 0), ".00")
    Exit Sub
Errhand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtInsurance.SetFocus
End Sub

Private Sub txtWsalePercent_GotFocus()
    txtWsalePercent.SelStart = 0
    txtWsalePercent.SelLength = Len(txtWsalePercent.Text)
End Sub

Private Sub txtWsalePercent_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            txtvanrate.SetFocus
         Case vbKeyEscape
            txtWS.SetFocus
    End Select
End Sub

Private Sub txtWsalePercent_LostFocus()
    On Error Resume Next
    If optdiscper.value = True Then
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) + ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
    Else
        'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
        txtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
        TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) + ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
    End If
    txtWS.Text = Round((Val(TXTPTR.Tag) * Val(txtWsalePercent.Text) / 100) + Val(TXTPTR.Tag), 2)
    txtWS.Text = Format(Val(txtWS.Text), "0.000")
End Sub

Private Sub TxtWarranty_GotFocus()
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
            txtBatch.SetFocus
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
    On Error GoTo Errhand
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'OP'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
Exit Function
Errhand:
    MsgBox Err.Description
End Function

Private Sub TxtExpense_GotFocus()
    TxtExpense.SelStart = 0
    TxtExpense.SelLength = Len(TxtExpense.Text)
End Sub

Private Sub TxtExpense_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRETAIL.SetFocus
         Case vbKeyEscape
            txtPD.SetFocus
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
    TxtCSTper.SelStart = 0
    TxtCSTper.SelLength = Len(TxtCSTper.Text)
End Sub

Private Sub TxtCSTper_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtTrDisc.SetFocus
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
    TxtTrDisc.SelStart = 0
    TxtTrDisc.SelLength = Len(TxtTrDisc.Text)
End Sub

Private Sub TxtTrDisc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtCessPer.SetFocus
         Case vbKeyEscape
            TxtCSTper.SetFocus
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
            Frame1.Enabled = True
            If OptComper.value = True Then
                TxtComper.SetFocus
            Else
                TxtComAmt.SetFocus
            End If
         Case vbKeyEscape
            txtcrtn.SetFocus
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
    TxtLWRate.Text = Format(TxtLWRate.Text, "0.00")
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim rstTRXMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TxtBarcode.Text) = "" Then
                txtcategory.Enabled = True
                txtcategory.SetFocus
                Exit Sub
            End If
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From RTRXFILE WHERE BARCODE= '" & Trim(TxtBarcode.Text) & "' ORDER BY VCH_NO ", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                rstTRXMAST.MoveLast
                CHANGE_FLAG = True
                TXTITEMCODE.Text = IIf(IsNull(rstTRXMAST!ITEM_CODE), "", rstTRXMAST!ITEM_CODE)
                TXTPRODUCT.Text = IIf(IsNull(rstTRXMAST!ITEM_NAME), "", rstTRXMAST!ITEM_NAME)
                CHANGE_FLAG = False
                TXTUNIT.Text = 1 'IIf(IsNull(rstTRXMAST!UNIT), "", rstTRXMAST!UNIT)
                Txtpack.Text = IIf(IsNull(rstTRXMAST!LINE_DISC), "", rstTRXMAST!LINE_DISC)
                Txtpack.Text = 1
                Los_Pack.Text = IIf(IsNull(rstTRXMAST!LOOSE_PACK), "1", rstTRXMAST!LOOSE_PACK)
                TxtWarranty.Text = IIf(IsNull(rstTRXMAST!WARRANTY), "", rstTRXMAST!WARRANTY)
                On Error Resume Next
                CmbPack.Text = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)
                CmbWrnty.Text = IIf(IsNull(rstTRXMAST!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, rstTRXMAST!WARRANTY_TYPE)
                'cmbcolor.Text = IIf(IsNull(rstTRXMAST!ITEM_COLOR), CmbWrnty.ListIndex = -1, rstTRXMAST!ITEM_COLOR)
                On Error GoTo Errhand
                'Txtsize.Text = IIf(IsNull(rstTRXMAST!ITEM_SIZE), "", rstTRXMAST!ITEM_SIZE)
                TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(rstTRXMAST!EXP_DATE), "  /  /    ", Format(rstTRXMAST!EXP_DATE, "DD/MM/YYYY"))
                txtBatch.Text = IIf(IsNull(rstTRXMAST!REF_NO), "", rstTRXMAST!REF_NO)
                TXTEXPIRY.Text = IIf(IsDate(rstTRXMAST!EXP_DATE), Format(rstTRXMAST!EXP_DATE, "MM/YY"), "  /  ")
                TXTRATE.Text = IIf(IsNull(rstTRXMAST!MRP), "", Format(Round(Val(rstTRXMAST!MRP) * Val(Los_Pack.Text), 2), ".000"))
                If (IsNull(rstTRXMAST!MRP_BT)) Then
                    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105
                Else
                    txtmrpbt.Text = Val(TXTRATE.Text)
                End If
                If IsNull(rstTRXMAST!PTR) Then
                    TXTPTR.Text = ""
                Else
                    TXTPTR.Text = Format(Round(Val(rstTRXMAST!PTR) * Val(Los_Pack.Text), 2), ".000")
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
                TxtWarranty.Text = IIf(IsNull(rstTRXMAST!WARRANTY), "", rstTRXMAST!WARRANTY)
                On Error Resume Next
                CmbPack.Text = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)
                CmbWrnty.Text = IIf(IsNull(rstTRXMAST!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, rstTRXMAST!WARRANTY_TYPE)
                On Error GoTo Errhand
                txtPD.Text = IIf(IsNull(rstTRXMAST!P_DISC), "", rstTRXMAST!P_DISC)
                Select Case rstTRXMAST!DISC_FLAG
                    Case "P"
                        optdiscper.value = True
                    Case "A"
                        Optdiscamt.value = True
                End Select
                'TxttaxMRP.Text = IIf(IsNull(rstTRXMAST!SALES_TAX), "", Format(Val(rstTRXMAST!SALES_TAX), ".00"))
                If rstTRXMAST!CHECK_FLAG = "M" Then
                    OPTTaxMRP.value = True
                ElseIf rstTRXMAST!CHECK_FLAG = "V" Then
                    OPTVAT.value = True
                Else
                    OPTNET.value = True
                End If
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
            
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
    End Select
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCess_GotFocus()
    txtCess.SelStart = 0
    txtCess.SelLength = Len(txtCess.Text)
End Sub

Private Sub txtCess_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.SetFocus
         Case vbKeyEscape
            TxtTrDisc.SetFocus
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
    Call TxttaxMRP_LostFocus
End Sub

Private Function print_labels(i As Long, BAR_LABEL As String, P_RATE As String)
    Dim wid As Single
    Dim hgt As Single
    
    On Error GoTo Errhand
    
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
    
    Picture1.Cls
    Picture1.Picture = Nothing
    Picture1.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
    Picture1.CurrentY = 0 'Y2 + 0.25 * Th
    Picture1.FontName = "MS Sans Serif"
    Picture1.FontSize = 7
    Picture1.FontBold = True
    Picture1.Print Trim(MDIMAIN.StatusBar.Panels(5).Text) 'COMP NAME
    
    Picture2.Cls
    Picture2.Picture = Nothing
    Picture2.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
    Picture2.CurrentY = 0 'Y2 + 0.25 * Th
    Picture2.FontName = "MS Sans Serif"
    Picture2.FontSize = 6
    Picture2.FontBold = False
    Picture2.Print Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2)) 'ITEM NAME
        
    If Val(TXTRETAIL.Text) <> 0 Then
        Picture5.Cls
        Picture5.Picture = Nothing
        Picture5.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
        Picture5.CurrentY = 0 'Y2 + 0.25 * Th
        Picture2.FontName = "Arial"
        Picture2.FontSize = 6
        Picture2.FontBold = True
        Picture5.Print "Price: " & Format(Val(TXTRETAIL.Text), "0.00")
    End If
    
    If Val(TXTRATE.Text) > 0 And Val(TXTRETAIL.Text) < Val(TXTRATE.Text) Then
        Picture6.Cls
        Picture6.Picture = Nothing
        Picture6.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
        Picture6.CurrentY = 0 'Y2 + 0.25 * Th
        Picture2.FontName = "Arial"
        Picture2.FontSize = 6
        Picture2.FontBold = True
        Picture6.Print "MRP  : " & Format(Val(TXTRATE.Text), "0.00")
    End If
    

'    Picture3.Cls
'    Picture3.Picture = Nothing
'    Picture3.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'    Picture3.CurrentY = 0 'Y2 + 0.25 * Th
'    Picture3.FontName = "barcode font"
'    Picture3.FontSize = 14
'    Picture3.FontBold = False
'    Picture3.Print BAR_LABEL
    
    Do Until i <= 0
        Picture1.ScaleMode = vbPixels
        Picture2.ScaleMode = vbPixels
        Picture5.ScaleMode = vbPixels
        Picture6.ScaleMode = vbPixels
        
'        Printer.PaintPicture Picture1.Image, 200, 600 ', wid, hgt
'        Printer.PaintPicture Picture1.Image, 2100, 600 ', wid, hgt
'        Printer.PaintPicture Picture1.Image, 4000, 600 ', wid, hgt
'
'        Printer.PaintPicture Picture6.Image, 1300, 600 ', wid, hgt 'MRP
'        Printer.PaintPicture Picture6.Image, 3200, 600 ', wid, hgt 'MRP
'        Printer.PaintPicture Picture6.Image, 5100, 600 ', wid, hgt 'MRP
'
'        Printer.PaintPicture Picture5.Image, 200, 800 ', wid, hgt
'        Printer.PaintPicture Picture5.Image, 2100, 800 ', wid, hgt
'        Printer.PaintPicture Picture5.Image, 4000, 800 ', wid, hgt

        Printer.PaintPicture Picture2.Image, 900, 800 ', wid, hgt  'Item Name
        Printer.PaintPicture Picture2.Image, 3150, 800 ', wid, hgt  'Item Name
        
        Printer.PaintPicture Picture5.Image, 900, 960 ', wid, hgt  ' Price
        Printer.PaintPicture Picture5.Image, 3150, 960 ', wid, hgt
        
        'Printer.PaintPicture Picture6.Image, 2000, 950 ', wid, hgt 'MRP
        'Printer.PaintPicture Picture6.Image, 3500, 600 ', wid, hgt 'MRP
        
        Printer.PaintPicture Picture1.Image, 900, 1160 ', wid, hgt  ' Comp Name
        Printer.PaintPicture Picture1.Image, 3150, 1160 ', wid, hgt
            
        Printer.FontName = "Arial"
        'Printer.FontName = "barcode font"
        'Printer.FontSize = 1
        Printer.FontBold = False
        'Printer.Print ""
        
        Printer.FontName = "IDAutomationHC39M"
        'Printer.FontName = "barcode font"
        Printer.FontSize = 24
        Printer.FontBold = False
        Dim bar_space As Integer
        If Len(BAR_LABEL) > 13 Then
            bar_space = 0
            Printer.FontSize = 6
        ElseIf Len(BAR_LABEL) >= 12 Then
            bar_space = 0
            Printer.FontSize = 7
        Else
            bar_space = 9 - Len(BAR_LABEL)
            Printer.FontSize = 11
        End If
        'Printer.Print " (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")"
        Printer.Print "    (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")"
'        'Picture1.ScaleMode = vbPixels
'        Picture5.ScaleMode = vbPixels
'        Picture6.ScaleMode = vbPixels
        ' Finish printing.
        Printer.EndDoc
        i = i - 2
    Loop
    
    Exit Function
Errhand:
    MsgBox Err.Description
End Function

Private Sub TxtCessPer_GotFocus()
    TxtCessPer.SelStart = 0
    TxtCessPer.SelLength = Len(TxtCessPer.Text)
End Sub

Private Sub TxtCessPer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtCess.SetFocus
         Case vbKeyEscape
            TxtTrDisc.SetFocus
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
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TxtHSN_GotFocus()
    txtHSN.SelStart = 0
    txtHSN.SelLength = Len(txtHSN.Text)
End Sub

Private Sub TxtHSN_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRETAIL.SetFocus
         Case vbKeyEscape
            TxttaxMRP.SetFocus
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

Private Function print_3labels(i As Long, BAR_LABEL As String, P_RATE As String)
    Dim wid As Single
    Dim hgt As Single
    
    On Error GoTo Errhand
    
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
'
    'i = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3))
    
    Picture1.Cls
    Picture1.Picture = Nothing
    Picture1.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
    Picture1.CurrentY = 0 'Y2 + 0.25 * Th
    Picture1.FontName = "MS Sans Serif"
    Picture1.FontSize = 7
    Picture1.FontBold = True
    Picture1.Print Trim(MDIMAIN.StatusBar.Panels(5).Text) 'COMP NAME
    
    Picture2.Cls
    Picture2.Picture = Nothing
    Picture2.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
    Picture2.CurrentY = 0 'Y2 + 0.25 * Th
    Picture2.FontName = "MS Sans Serif"
    Picture2.FontSize = 6
    Picture2.FontBold = False
    Picture2.Print Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2)) 'ITEM NAME
        
    If Val(TXTRETAIL.Text) <> 0 Then
        Picture5.Cls
        Picture5.Picture = Nothing
        Picture5.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
        Picture5.CurrentY = 0 'Y2 + 0.25 * Th
        Picture2.FontName = "Arial"
        Picture2.FontSize = 6
        Picture2.FontBold = True
        Picture5.Print "Price: " & Format(Val(TXTRETAIL.Text), "0.00")
    End If
    
    If Val(TXTRATE.Text) > 0 And Val(TXTRETAIL.Text) < Val(TXTRATE.Text) Then
        Picture6.Cls
        Picture6.Picture = Nothing
        Picture6.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
        Picture6.CurrentY = 0 'Y2 + 0.25 * Th
        Picture2.FontName = "Arial"
        Picture2.FontSize = 6
        Picture2.FontBold = True
        Picture6.Print "MRP  : " & Format(Val(TXTRATE.Text), "0.00")
    End If
    

'    Picture3.Cls
'    Picture3.Picture = Nothing
'    Picture3.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'    Picture3.CurrentY = 0 'Y2 + 0.25 * Th
'    Picture3.FontName = "barcode font"
'    Picture3.FontSize = 14
'    Picture3.FontBold = False
'    Picture3.Print BAR_LABEL
    
    Do Until i <= 0
        Picture1.ScaleMode = vbPixels
        Picture2.ScaleMode = vbPixels
        Picture5.ScaleMode = vbPixels
        Picture6.ScaleMode = vbPixels
        
'        Printer.PaintPicture Picture1.Image, 200, 600 ', wid, hgt
'        Printer.PaintPicture Picture1.Image, 2100, 600 ', wid, hgt
'        Printer.PaintPicture Picture1.Image, 4000, 600 ', wid, hgt
'
'        Printer.PaintPicture Picture6.Image, 1300, 600 ', wid, hgt 'MRP
'        Printer.PaintPicture Picture6.Image, 3200, 600 ', wid, hgt 'MRP
'        Printer.PaintPicture Picture6.Image, 5100, 600 ', wid, hgt 'MRP
'
'        Printer.PaintPicture Picture5.Image, 200, 800 ', wid, hgt
'        Printer.PaintPicture Picture5.Image, 2100, 800 ', wid, hgt
'        Printer.PaintPicture Picture5.Image, 4000, 800 ', wid, hgt

        Printer.PaintPicture Picture2.Image, 200, 880 ', wid, hgt  'Item Name
        Printer.PaintPicture Picture2.Image, 2500, 880 ', wid, hgt  'Item Name
        
        Printer.PaintPicture Picture5.Image, 200, 1040 ', wid, hgt  ' Price
        Printer.PaintPicture Picture5.Image, 2500, 1040 ', wid, hgt
        
        'Printer.PaintPicture Picture6.Image, 2000, 950 ', wid, hgt 'MRP
        'Printer.PaintPicture Picture6.Image, 3500, 600 ', wid, hgt 'MRP
        
        Printer.PaintPicture Picture1.Image, 200, 1240 ', wid, hgt  ' Comp Name
        Printer.PaintPicture Picture1.Image, 2500, 1240 ', wid, hgt
            
        Printer.FontName = "Arial"
        'Printer.FontName = "barcode font"
        Printer.FontSize = 5
        Printer.FontBold = False
        Printer.Print ""
        
        Printer.FontName = "IDAutomationHC39M"
        'Printer.FontName = "barcode font"
        Printer.FontSize = 24
        Printer.FontBold = False
        Dim bar_space As Integer
        If Len(BAR_LABEL) > 13 Then
            bar_space = 0
            Printer.FontSize = 6
        ElseIf Len(BAR_LABEL) >= 12 Then
            bar_space = 0
            Printer.FontSize = 7
        Else
            bar_space = 9 - Len(BAR_LABEL)
            Printer.FontSize = 11
        End If
        'Printer.Print " (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")"
        Printer.Print " (" & BAR_LABEL & ")" & Space(bar_space) & "(" & BAR_LABEL & ")"
'        'Picture1.ScaleMode = vbPixels
'        Picture5.ScaleMode = vbPixels
'        Picture6.ScaleMode = vbPixels
        ' Finish printing.
        Printer.EndDoc
        i = i - 2
    Loop
    
    Exit Function
Errhand:
    MsgBox Err.Description
End Function
