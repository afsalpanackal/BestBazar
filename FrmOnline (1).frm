VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOnline 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Purchase from Suppliers / Branch"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17055
   ControlBox      =   0   'False
   Icon            =   "FrmOnline.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   17055
   Begin VB.CommandButton CmdCancel 
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
      Left            =   5100
      TabIndex        =   18
      Top             =   8325
      Width           =   1245
   End
   Begin VB.CommandButton cmditemcreate 
      Caption         =   "&Create Item"
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
      Left            =   150
      TabIndex        =   14
      Top             =   8325
      Width           =   1245
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
      Left            =   5070
      TabIndex        =   13
      Top             =   7545
      Width           =   1245
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   750
      TabIndex        =   4
      Top             =   3255
      Visible         =   0   'False
      Width           =   9060
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   3675
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   6482
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
      BackColor       =   &H00C0C000&
      Height          =   9240
      Left            =   -60
      TabIndex        =   2
      Top             =   -135
      Width           =   17130
      Begin VB.Frame Frmmain 
         BackColor       =   &H0080C0FF&
         Height          =   1515
         Left            =   90
         TabIndex        =   19
         Top             =   75
         Width           =   6375
         Begin VB.Frame Frame1 
            BackColor       =   &H0080C0FF&
            Height          =   825
            Left            =   5010
            TabIndex        =   97
            Top             =   135
            Width           =   1320
            Begin VB.OptionButton OptBranch 
               BackColor       =   &H0080C0FF&
               Caption         =   "&Branch"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   90
               TabIndex        =   99
               Top             =   495
               Width           =   1020
            End
            Begin VB.OptionButton OptOthers 
               BackColor       =   &H0080C0FF&
               Caption         =   "S&uppliers"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   90
               TabIndex        =   98
               Top             =   210
               Value           =   -1  'True
               Width           =   1125
            End
         End
         Begin VB.TextBox txtBillNo 
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   3840
            TabIndex        =   25
            Top             =   135
            Width           =   1140
         End
         Begin VB.CommandButton CmdLoadInv 
            Caption         =   "&Load Invoice"
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
            Left            =   5055
            TabIndex        =   24
            Top             =   990
            Width           =   1245
         End
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
            TabIndex        =   20
            Top             =   465
            Width           =   3735
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
            Height          =   315
            Left            =   1245
            TabIndex        =   0
            Top             =   135
            Width           =   2520
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1245
            TabIndex        =   21
            Top             =   810
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice No"
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
            TabIndex        =   23
            Top             =   195
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier"
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
            Left            =   90
            TabIndex        =   22
            Top             =   495
            Width           =   1005
         End
      End
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00C0C000&
         Height          =   1515
         Left            =   6465
         TabIndex        =   6
         Top             =   75
         Visible         =   0   'False
         Width           =   4260
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3795
            Top             =   270
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
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
            Left            =   1545
            MaxLength       =   100
            TabIndex        =   10
            Top             =   1020
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
            Left            =   1545
            MaxLength       =   10
            TabIndex        =   9
            Top             =   165
            Width           =   1260
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   1545
            TabIndex        =   12
            Top             =   585
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
            Left            =   9480
            TabIndex        =   15
            Top             =   645
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            Left            =   120
            TabIndex        =   11
            Top             =   975
            Width           =   1290
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Entry"
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
            Left            =   120
            TabIndex        =   8
            Top             =   150
            Width           =   1350
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Date"
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
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Height          =   7500
         Left            =   60
         TabIndex        =   3
         Top             =   1515
         Visible         =   0   'False
         Width           =   17070
         Begin MSMask.MaskEdBox TxtExp 
            Height          =   315
            Left            =   12255
            TabIndex        =   96
            Top             =   1305
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
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
            Left            =   10020
            TabIndex        =   31
            Top             =   1560
            Visible         =   0   'False
            Width           =   795
         End
         Begin MSDataListLib.DataList DataList1 
            Height          =   1425
            Left            =   1950
            TabIndex        =   28
            Top             =   1005
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
            Left            =   1950
            TabIndex        =   27
            Top             =   660
            Visible         =   0   'False
            Width           =   3735
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   5325
            Left            =   45
            TabIndex        =   1
            Top             =   120
            Width           =   16995
            _ExtentX        =   29977
            _ExtentY        =   9393
            _Version        =   393216
            Rows            =   1
            Cols            =   29
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            FocusRect       =   2
            AllowUserResizing=   1
            Appearance      =   0
            GridLineWidth   =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Frame FRMECONTROLS 
            BackColor       =   &H00C0C000&
            Height          =   2085
            Left            =   30
            TabIndex        =   32
            Top             =   5385
            Visible         =   0   'False
            Width           =   17010
            Begin VB.CommandButton CMDADDITEM 
               Caption         =   "Add &Item"
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
               Left            =   15705
               TabIndex        =   95
               Top             =   165
               Width           =   1245
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
               Height          =   450
               Left            =   3870
               TabIndex        =   56
               Top             =   780
               Width           =   1125
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
               Left            =   6210
               TabIndex        =   55
               Top             =   2310
               Visible         =   0   'False
               Width           =   825
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
               Height          =   285
               Left            =   6810
               MaxLength       =   15
               TabIndex        =   54
               Top             =   450
               Width           =   960
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
               Left            =   9690
               TabIndex        =   53
               Top             =   1950
               Visible         =   0   'False
               Width           =   690
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
               Left            =   1485
               TabIndex        =   52
               Top             =   780
               Visible         =   0   'False
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
               Height          =   450
               Left            =   2700
               TabIndex        =   51
               Top             =   780
               Width           =   1110
            End
            Begin VB.TextBox TXTQTY 
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
               Left            =   5415
               MaxLength       =   7
               TabIndex        =   50
               Top             =   435
               Width           =   660
            End
            Begin VB.TextBox TXTPRODUCT 
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
               Left            =   735
               TabIndex        =   49
               Top             =   450
               Width           =   3975
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
               Height          =   300
               Left            =   150
               TabIndex        =   48
               Top             =   450
               Width           =   570
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
               Left            =   135
               TabIndex        =   47
               Top             =   780
               Width           =   1245
            End
            Begin VB.TextBox TXTRATE 
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
               Left            =   9015
               MaxLength       =   6
               TabIndex        =   46
               Top             =   435
               Width           =   765
            End
            Begin VB.TextBox TXTPTR 
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
               Left            =   9795
               MaxLength       =   6
               TabIndex        =   45
               Top             =   435
               Width           =   765
            End
            Begin VB.TextBox Txtpack 
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
               Left            =   4740
               MaxLength       =   7
               TabIndex        =   44
               Top             =   435
               Width           =   660
            End
            Begin VB.TextBox TxttaxMRP 
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
               Left            =   10575
               MaxLength       =   7
               TabIndex        =   43
               Top             =   435
               Width           =   615
            End
            Begin VB.OptionButton OPTVAT 
               BackColor       =   &H00C0C000&
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
               Height          =   330
               Left            =   8490
               TabIndex        =   42
               Top             =   840
               Width           =   1005
            End
            Begin VB.OptionButton OPTTaxMRP 
               BackColor       =   &H00C0C000&
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
               Height          =   240
               Left            =   6705
               TabIndex        =   41
               Top             =   885
               Width           =   1680
            End
            Begin VB.TextBox TxtFree 
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
               Left            =   6090
               MaxLength       =   7
               TabIndex        =   40
               Top             =   435
               Width           =   705
            End
            Begin VB.OptionButton OPTNET 
               BackColor       =   &H00C0C000&
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
               Height          =   330
               Left            =   9540
               TabIndex        =   39
               Top             =   855
               Value           =   -1  'True
               Width           =   780
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
               Height          =   345
               Left            =   9330
               MaxLength       =   10
               TabIndex        =   38
               Top             =   1635
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
               Height          =   345
               Left            =   7965
               MaxLength       =   10
               TabIndex        =   37
               Top             =   1635
               Width           =   1095
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
               Height          =   345
               Left            =   6585
               MaxLength       =   10
               TabIndex        =   36
               Top             =   1635
               Width           =   1095
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
               Left            =   6240
               MaxLength       =   6
               TabIndex        =   35
               Top             =   2085
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox txtPD 
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
               Left            =   12060
               MaxLength       =   7
               TabIndex        =   34
               Top             =   435
               Width           =   585
            End
            Begin VB.TextBox Txtdisccust 
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
               Left            =   14055
               MaxLength       =   7
               TabIndex        =   33
               Top             =   435
               Width           =   885
            End
            Begin MSMask.MaskEdBox TXTEXPIRY 
               Height          =   315
               Left            =   7755
               TabIndex        =   57
               Top             =   420
               Visible         =   0   'False
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
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
               Height          =   315
               Left            =   7785
               TabIndex        =   58
               Top             =   420
               Width           =   1215
               _ExtentX        =   2143
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
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Act Profit%"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Index           =   55
               Left            =   3795
               TabIndex        =   94
               Top             =   1275
               Width           =   1305
            End
            Begin VB.Label lblactprofit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   3780
               TabIndex        =   93
               Top             =   1500
               Width           =   1170
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
               Left            =   4725
               TabIndex        =   92
               Top             =   2250
               Visible         =   0   'False
               Width           =   1080
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
               Left            =   6225
               TabIndex        =   91
               Top             =   2160
               Visible         =   0   'False
               Width           =   825
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
               Left            =   12660
               TabIndex        =   90
               Top             =   435
               Width           =   1380
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
               Left            =   6810
               TabIndex        =   89
               Top             =   195
               Width           =   960
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
               Height          =   255
               Index           =   16
               Left            =   7785
               TabIndex        =   88
               Top             =   195
               Width           =   1215
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
               Left            =   8580
               TabIndex        =   87
               Top             =   2025
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
               ForeColor       =   &H008080FF&
               Height          =   270
               Index           =   14
               Left            =   12660
               TabIndex        =   86
               Top             =   195
               Width           =   1380
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
               Left            =   9015
               TabIndex        =   85
               Top             =   195
               Width           =   765
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
               Height          =   255
               Index           =   10
               Left            =   5415
               TabIndex        =   84
               Top             =   195
               Width           =   660
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
               Height          =   270
               Index           =   9
               Left            =   735
               TabIndex        =   83
               Top             =   195
               Width           =   3975
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
               Height          =   255
               Index           =   8
               Left            =   150
               TabIndex        =   82
               Top             =   195
               Width           =   570
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "RATE"
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
               Index           =   2
               Left            =   9795
               TabIndex        =   81
               Top             =   195
               Width           =   765
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
               Height          =   270
               Index           =   4
               Left            =   4740
               TabIndex        =   80
               Top             =   195
               Width           =   660
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
               Height          =   255
               Index           =   12
               Left            =   10575
               TabIndex        =   79
               Top             =   195
               Width           =   615
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
               Left            =   11205
               TabIndex        =   78
               Top             =   195
               Width           =   825
            End
            Begin VB.Label lbltaxamount 
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
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   11205
               TabIndex        =   77
               Top             =   435
               Width           =   825
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
               Height          =   255
               Index           =   17
               Left            =   6090
               TabIndex        =   76
               Top             =   195
               Width           =   705
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Purchase Amt"
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
               Left            =   10905
               TabIndex        =   75
               Top             =   1365
               Width           =   1875
            End
            Begin VB.Label lbltotalwodiscount 
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
               Left            =   10890
               TabIndex        =   74
               Top             =   1575
               Width           =   1950
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
               Left            =   9435
               TabIndex        =   73
               Top             =   1365
               Width           =   945
            End
            Begin VB.Label LBLTOTAL 
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
               Height          =   450
               Left            =   10890
               TabIndex        =   72
               Top             =   930
               Width           =   1935
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
               Height          =   330
               Index           =   21
               Left            =   10875
               TabIndex        =   71
               Top             =   720
               Width           =   1245
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Cr. Note Amt"
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
               Left            =   7875
               TabIndex        =   70
               Top             =   1365
               Width           =   1305
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
               Left            =   6645
               TabIndex        =   69
               Top             =   1365
               Width           =   1080
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
               ForeColor       =   &H008080FF&
               Height          =   255
               Index           =   25
               Left            =   12060
               TabIndex        =   68
               Top             =   195
               Width           =   585
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Profit%"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Index           =   24
               Left            =   1440
               TabIndex        =   67
               Top             =   1275
               Width           =   1080
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Profit Amt"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Index           =   26
               Left            =   2460
               TabIndex        =   66
               Top             =   1275
               Width           =   1305
            End
            Begin VB.Label lblprftper 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   1425
               TabIndex        =   65
               Top             =   1500
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Sale Value"
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
               Height          =   330
               Index           =   27
               Left            =   12855
               TabIndex        =   64
               Top             =   720
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.Label LblSale_Val 
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
               Height          =   450
               Left            =   12960
               TabIndex        =   63
               Top             =   930
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.Label lblPrftAmt 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   2460
               TabIndex        =   62
               Top             =   1500
               Width           =   1290
            End
            Begin VB.Label LblProfittotal 
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
               Height          =   450
               Left            =   12960
               TabIndex        =   61
               Top             =   1575
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Profit Amount"
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
               Height          =   330
               Index           =   28
               Left            =   12975
               TabIndex        =   60
               Top             =   1365
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "Cust Disc"
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
               Index           =   1
               Left            =   14055
               TabIndex        =   59
               Top             =   195
               Width           =   885
            End
         End
      End
      Begin VB.Label LBLitem 
         Height          =   315
         Left            =   13155
         TabIndex        =   30
         Top             =   750
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label flagchange2 
         Height          =   315
         Left            =   11865
         TabIndex        =   29
         Top             =   915
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblitemname 
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
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   10740
         TabIndex        =   26
         Top             =   180
         Width           =   4350
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   12120
         TabIndex        =   17
         Top             =   210
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   11415
         TabIndex        =   16
         Top             =   165
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim ACT_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT As Boolean
Dim M_ADD As Boolean
Dim ITEM_REC As New ADODB.Recordset
Dim ITEM_FLAG As Boolean

Private Sub CMDADD_Click()
    Dim i As Integer
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim M_DATA As Double

    M_DATA = 0
    If grdsales.rows <= 1 Then Exit Sub
    
    If Val(Txtpack.text) = 0 Then
        Txtpack.Enabled = True
        Txtpack.SetFocus
        Exit Sub
    End If
    If (Val(TXTQTY.text) = 0 And Val(TXTFREE.text) = 0) Then
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
        Exit Sub
    End If
    If Trim(txtBatch.text) = "" Then
        txtBatch.Enabled = True
        txtBatch.SetFocus
        Exit Sub
    End If
    
    If Val(Mid(TXTEXPIRY.text, 1, 2)) = 0 Then
        TXTEXPIRY.Visible = True
        TXTEXPIRY.SetFocus
        Exit Sub
    End If
    If Val(Mid(TXTEXPIRY.text, 1, 2)) > 12 Then
        TXTEXPIRY.Visible = True
        TXTEXPIRY.SetFocus
        Exit Sub
    End If
    If Val(Mid(TXTEXPIRY.text, 4, 5)) = 0 Then
        TXTEXPIRY.Visible = True
        TXTEXPIRY.SetFocus
        Exit Sub
    End If
    If Val(Mid(TXTEXPIRY.text, 1, 2)) = 0 Then
        TXTEXPDATE.text = "  /  /    "
        TXTEXPIRY.Visible = True
        TXTEXPIRY.SetFocus
        Exit Sub
    End If
    If Val(Mid(TXTEXPIRY.text, 4, 5)) = 0 Then
        TXTEXPDATE.text = "  /  /    "
        TXTEXPIRY.Visible = True
        TXTEXPIRY.SetFocus
        Exit Sub
        Exit Sub
    End If
    
    If Val(Mid(TXTEXPIRY.text, 1, 2)) > 12 Then
        TXTEXPDATE.text = "  /  /    "
        TXTEXPIRY.Visible = True
        TXTEXPIRY.SetFocus
        Exit Sub
    End If
    
    Dim M_DATE As Date
    Dim D As Integer
    Dim M As Integer
    Dim Y As Integer

    M = Val(Mid(TXTEXPIRY.text, 1, 2))
    Y = Val(Right(TXTEXPIRY.text, 2))
    Y = 2000 + Y
    M_DATE = "01" & "/" & M & "/" & Y
    D = LastDayOfMonth(M_DATE)
    M_DATE = D & "/" & M & "/" & Y
    TXTEXPDATE.text = Format(M_DATE, "dd/mm/yyyy")
    
    If DateDiff("d", Date, TXTEXPDATE.text) < 0 Then
        MsgBox "Item Expired....", vbOKOnly, "PURCHASE.."
        TXTEXPDATE.text = "  /  /    "
        TXTEXPIRY.SelStart = 0
        TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
        TXTEXPIRY.Visible = True
        TXTEXPIRY.SetFocus
        Exit Sub
    End If
    
    If DateDiff("d", Date, TXTEXPDATE.text) < 60 Then
        MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days", vbOKOnly, "PURCHASE.."
        TXTEXPDATE.text = "  /  /    "
        TXTEXPIRY.SelStart = 0
        TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
        TXTEXPIRY.Visible = True
        TXTEXPIRY.SetFocus
        Exit Sub
    End If
    
    If DateDiff("d", Date, TXTEXPDATE.text) < 180 Then
        If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "PURCHASE..") = vbNo Then
            TXTEXPDATE.text = "  /  /    "
            TXTEXPIRY.SelStart = 0
            TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
            TXTEXPIRY.Visible = True
            TXTEXPIRY.SetFocus
            Exit Sub
        End If
    End If
    If Val(TXTRATE.text) = 0 Then
        TXTRATE.Enabled = True
        TXTRATE.SetFocus
        Exit Sub
    End If
    If Val(TXTPTR.text) = 0 Then
        TXTPTR.Enabled = True
        TXTPTR.SetFocus
        Exit Sub
    End If
    Call txtPD_LostFocus
'    For i = 1 To grdsales.Rows
'        If Trim(grdsales.TextMatrix(i, 1)) = "" Then
'            MsgBox " Item Code not being assigned to Sl No. " & i, vbOKOnly, "Online Entry"
'            Exit Sub
'        End If
'    Next i
    
    'For i = 1 To grdsales.Rows
        'TXTSLNO.Text = grdsales.TextMatrix(i, 0)
    grdsales.TextMatrix(Val(TXTSLNO.text), 1) = Trim(TXTITEMCODE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 2) = Trim(TXTPRODUCT.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 3) = Val(TXTQTY.text) + Val(TXTFREE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 4) = 1
    grdsales.TextMatrix(Val(TXTSLNO.text), 5) = Val(Txtpack.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 6) = Format((Val(TXTRATE.text)), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 7) = Format((Val(TXTRATE.text) / Val(Txtpack.text)), ".000")
    If Val(TXTQTY.text) = 0 Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 8) = "0.00"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.text), 8) = Format(Round((Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.text)) / Val(Txtpack.text), 3), ".000")
    End If
    grdsales.TextMatrix(Val(TXTSLNO.text), 9) = Format(Val(TXTPTR.text) / Val(Txtpack.text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 10) = IIf(Val(TxttaxMRP.text) = 0, "", Format(Val(TxttaxMRP.text), ".00")) 'TAX
    grdsales.TextMatrix(Val(TXTSLNO.text), 11) = Trim(txtBatch.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 12) = IIf(Trim(TXTEXPDATE.text) = "/  /", "", TXTEXPDATE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 13) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 14) = Val(TXTFREE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 17) = Val(txtPD.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 20) = Val(Txtdisccust.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 18) = Format((Val(txtmrpbt.text) / Val(Txtpack.text)), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 21) = "Y"

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

        
    grdsales.Row = Val(TXTSLNO.text)
    Call grdsales_Click
    TXTPRODUCT.text = ""
    
    TXTITEMCODE.text = ""
    Txtpack.text = ""
    TXTPTR.text = ""
    TXTQTY.text = ""
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    txtPD.text = ""
    TXTRATE.text = ""
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
    TXTSLNO.text = Val(TXTSLNO.text) + 1
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    FRMEMASTER.Enabled = True
    'CmdCancel.Enabled = False
    lblitemname.Caption = ""
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

Private Sub CMDADDITEM_Click()
    If grdsales.rows = 1 Then Exit Sub
    If grdsales.TextMatrix(grdsales.Row, 1) <> "" Then Exit Sub
    If Trim(Trim(grdsales.TextMatrix(grdsales.Row, 22))) = "" Then Exit Sub
    
    If MsgBox("ARE YOU SURE YOU WANT TO ADD THE ITEM " & Trim(grdsales.TextMatrix(grdsales.Row, 2)) & " AUTOMATICALLY !!!!", vbYesNo + vbDefaultButton2, "ONLINE PURCHASE ENTRY.....") = vbNo Then Exit Sub
    CMDADDITEM.Tag = ""
    On Error GoTo ErrHand
    Dim RSTITEMMAST As ADODB.Recordset
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        CMDADDITEM.Tag = IIf(IsNull(RSTITEMMAST.Fields(0)), 1, Val(RSTITEMMAST.Fields(0)) + 1)
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & CMDADDITEM.Tag & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!ITEM_CODE = CMDADDITEM.Tag
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
        RSTITEMMAST!UNIT = 1
        RSTITEMMAST!PACK_TYPE = "Nos"
        RSTITEMMAST!FULL_PACK = "Nos"
        RSTITEMMAST!DEAD_STOCK = "N"
        RSTITEMMAST!UN_BILL = "N"
        RSTITEMMAST!PRICE_CHANGE = "N"
        RSTITEMMAST!ITEM_MAL = ""
        RSTITEMMAST!check_flag = "V"
    End If
    RSTITEMMAST!ITEM_NAME = Trim(grdsales.TextMatrix(grdsales.Row, 22))
    RSTITEMMAST!Category = "GENERAL"
    RSTITEMMAST!MANUFACTURER = Trim(grdsales.TextMatrix(grdsales.Row, 24))
    RSTITEMMAST!SCHEDULE = "H"
    If Trim(Trim(grdsales.TextMatrix(grdsales.Row, 25))) <> "" Then RSTITEMMAST!REMARKS = Trim(grdsales.TextMatrix(grdsales.Row, 25))
    If Val(grdsales.TextMatrix(grdsales.Row, 5)) <= 0 Then grdsales.TextMatrix(grdsales.Row, 5) = 1
    RSTITEMMAST!REORDER_QTY = Val(grdsales.TextMatrix(grdsales.Row, 5))
    RSTITEMMAST!BIN_LOCATION = Left(Trim(grdsales.TextMatrix(grdsales.Row, 2)), 1)
    RSTITEMMAST!P_WS = 0
    RSTITEMMAST!CRTN_PACK = 1
    RSTITEMMAST!SALES_TAX = Val(grdsales.TextMatrix(grdsales.Row, 10))
    RSTITEMMAST!ITEM_COST = Val(grdsales.TextMatrix(grdsales.Row, 8))
    RSTITEMMAST!P_RETAIL = Val(grdsales.TextMatrix(grdsales.Row, 26))
    RSTITEMMAST!P_CRTN = Round(Val(grdsales.TextMatrix(grdsales.Row, 26)) / Val(grdsales.TextMatrix(grdsales.Row, 5)), 3)
    RSTITEMMAST!P_WS = Val(grdsales.TextMatrix(grdsales.Row, 27))
    RSTITEMMAST!P_LWS = Round(Val(grdsales.TextMatrix(grdsales.Row, 27)) / Val(grdsales.TextMatrix(grdsales.Row, 5)), 3)
    RSTITEMMAST!CRTN_PACK = 1
    RSTITEMMAST!MRP = Val(grdsales.TextMatrix(grdsales.Row, 6))
    RSTITEMMAST!PTR = 0
    RSTITEMMAST!LOOSE_PACK = Val(grdsales.TextMatrix(grdsales.Row, 5))
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    grdsales.TextMatrix(grdsales.Row, 1) = CMDADDITEM.Tag
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT MANUFACTURER FROM MANUFACT WHERE MANUFACTURER = '" & Trim(grdsales.TextMatrix(grdsales.Row, 24)) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!MANUFACTURER = Trim(grdsales.TextMatrix(grdsales.Row, 24))
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & CMDADDITEM.Tag & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTITEMMAST.EOF
        RSTITEMMAST!ITEM_NAME = Trim(grdsales.TextMatrix(grdsales.Row, 2))
        RSTITEMMAST!MFGR = Trim(grdsales.TextMatrix(grdsales.Row, 24))
        RSTITEMMAST!Category = "GENERAL"
        RSTITEMMAST.Update
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    On Error Resume Next
    grdsales.SetFocus
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub cmdcancel_Click()
        
    If grdsales.rows = 1 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO CANCEL !!!!", vbYesNo, "ONLINE PURCHASE ENTRY.....") = vbNo Then Exit Sub
    FRMEMASTER.Visible = False
    Frame2.Visible = False
    FRMECONTROLS.Visible = False
    Frmmain.Enabled = True
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    'cmdRefresh.Enabled = False
    'FRMEMASTER.Enabled = False
    'FRMECONTROLS.Enabled = False
    DataList2.text = ""
    TXTINVDATE.text = "  /  /    "
    TXTINVOICE.text = ""
    TXTREMARKS.text = ""
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TXTPRODUCT.text = ""
    Txtpack.text = ""
    TXTQTY.text = ""
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    txtPD.text = ""
    txtBatch.text = ""
    TXTRATE.text = ""
    txtmrpbt.text = ""
    TXTPTR.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    txtaddlamt.text = ""
    txtcramt.text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LblSale_Val.Caption = ""
    LblProfittotal.Caption = ""
    TXTDISCAMOUNT.text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    flagchange2.Caption = ""
    TXTDEALER.text = ""
    lbldealer.Caption = ""
    LBLitem.Caption = ""
    lblPrftAmt.Caption = ""
    lblprftper.Caption = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    M_ADD = False
    
    TXTINVOICE.SetFocus
End Sub

Private Sub CmdDelete_Click()
    Dim i As Integer
    Dim rststock, rstMaxNo, RSTRTRXFILE   As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
   
    If OptOthers.Value = True Then
        db.Execute "delete  From RTRXFILE WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & ""
    Else
        db.Execute "delete  From RTRXFILE WHERE TRX_TYPE='TF' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & ""
    End If
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With rststock
        If Not (.EOF And .BOF) Then
            If (IsNull(!RCPT_QTY)) Then !RCPT_QTY = 0
            !RCPT_QTY = !RCPT_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
            If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
            !RCPT_VAL = !RCPT_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 13))
            
            If (IsNull(!CLOSE_QTY)) Then !CLOSE_QTY = 0
            !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 13))
            rststock.Update
        End If
    End With
    rststock.Close
    Set rststock = Nothing
    
    i = 0
    Set rstMaxNo = New ADODB.Recordset
    If OptOthers.Value = True Then
        rstMaxNo.Open "Select MAX(Val(LINE_NO)) From RTRXFILE WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly
    Else
        rstMaxNo.Open "Select MAX(Val(LINE_NO)) From RTRXFILE WHERE TRX_TYPE='TF' AND VCH_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly
    End If
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    Set RSTRTRXFILE = New ADODB.Recordset
    If OptOthers.Value = True Then
        RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    Else
        RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_TYPE='TF' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    End If
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
    If OptOthers.Value = True Then
        RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    Else
        RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_TYPE='TF' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    End If
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
    LblSale_Val.Caption = ""
    LblProfittotal.Caption = ""
    
    grdsales.rows = 1
    i = 0
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    LblSale_Val.Caption = ""
    LblProfittotal.Caption = ""
    grdsales.rows = 1
    For i = Val(TXTSLNO.text) To grdsales.rows - 2
        grdsales.TextMatrix(i, 0) = i
        grdsales.TextMatrix(i, 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(i, 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(i, 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(i, 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(i, 5) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(i, 6) = grdsales.TextMatrix(i + 1, 6)
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
    Next i
    grdsales.rows = grdsales.rows - 1
    
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LblSale_Val.Caption = ""
    LblProfittotal.Caption = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        lbltotalwodiscount.Caption = Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13))
        LblSale_Val.Caption = Val(LblSale_Val.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3)))
    Next i
    
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.text)) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), ".00")
    LblProfittotal.Caption = Format(Val(LblSale_Val.Caption) - Val(LBLTOTAL.Caption), ".00")
    
    TXTSLNO.text = Val(grdsales.rows)
    TXTPRODUCT.text = ""
    TXTITEMCODE.text = ""
    Txtpack.text = ""
    TXTQTY.text = ""
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    txtPD.text = ""
    TXTRATE.text = ""
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
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub cmditemcreate_Click()
    frmitemmaster.Show
    frmitemmaster.TXTITEM.text = Trim(lblitemname.Caption)
End Sub

Private Sub cmditemcreate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If Txtpack.Enabled = True Then Txtpack.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTPTR.Enabled = True Then TXTPTR.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub CmdLoadInv_Click()
    On Error GoTo ErrHand
    Dim i As Long
    'MsgBox "Payment not done for this updation", vbOKOnly, "Import Purchase"
    'Exit Sub
    If Trim(TXTINVOICE.text) = "" Then
        MsgBox "Please enter the Invoice number", vbOKOnly, "Purchase Entry"
        TXTINVOICE.SetFocus
        Exit Sub
    End If
    
    If DataList2.BoundText = "" Then
        MsgBox "Please select the supplier from the list", vbOKOnly, "Purchase Entry"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
'    'ANANDA
'    If DataList2.BoundText = "311005" Then Call PARAGON
'    If DataList2.BoundText = "311002" Then Call OCEAN
'    If DataList2.BoundText = "311021" Then Call DAVIS
'    If DataList2.BoundText = "311008" Then Call ENGLISH
'    If DataList2.BoundText = "311011" Then Call PRABHU
    
    If OptOthers.Value = True Then
        Call Import_Bill
    Else
        Call Import_Br_Bill
    End If
    
    'If DataList2.BoundText = "311705" Then Call INTERPHARMA

    
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    LblSale_Val.Caption = ""
    LblProfittotal.Caption = ""
    For i = 1 To grdsales.rows - 1
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
        LblSale_Val.Caption = Format(Val(LblSale_Val.Caption) + (Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 3))), ".00")
    Next i
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.text)) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), ".00")
    LblProfittotal.Caption = Format(Val(LblSale_Val.Caption) - Val(LBLTOTAL.Caption), ".00")
    
    If grdsales.rows > 1 Then cmdRefresh.Enabled = True
    flagchange.Caption = ""
    flagchange2.Caption = ""
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.text) >= grdsales.rows Then Exit Sub
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 21)) = "Y" Then
        M_EDIT = True
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
    Else
        M_EDIT = False
        TXTPRODUCT.Enabled = True
        TXTPRODUCT.SetFocus
    End If
    
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            TXTQTY.text = ""
            TXTFREE.text = ""
            TxttaxMRP.text = ""
            txtPD.text = ""
            TXTRATE.text = ""
            txtmrpbt.text = ""
            TXTITEMCODE.text = ""
            Txtpack.text = ""
            LBLSUBTOTAL.Caption = ""
            lbltaxamount.Caption = ""
            TXTEXPDATE.text = "  /  /    "
            TXTEXPIRY.text = "  /  "
            txtBatch.text = ""
        
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
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

Private Sub cmdRefresh_Click()

    Dim i As Integer
    
    On Error GoTo ErrHand
    For i = 1 To grdsales.rows - 1
        If Trim(grdsales.TextMatrix(i, 1)) = "" Then
            MsgBox "Not Completed", vbOKOnly
            Exit Sub
        End If
    Next i
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Supplier From List", vbOKOnly, "PURCHASE..."
        FRMEMASTER.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
    
    If DataList2.BoundText = "" Then
        MsgBox "Select Supplier From List", vbOKOnly, "PURCHASE..."
        FRMEMASTER.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
    
    If TXTINVOICE.text = "" Then
        MsgBox "Enter Supplier Invoice No.", vbOKOnly, "PURCHASE"
        FRMEMASTER.Enabled = True
        Exit Sub
    End If
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Enter Supplier Invoice Date", vbOKOnly, "PURCHASE"
        FRMEMASTER.Enabled = True
        Exit Sub
    End If
    
    Call appendpurchase
    cmdcancel.Enabled = True
    
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrHand
    'Call FILLCOMBO
    Exit Sub
ErrHand:
    If err.Number = 5 Then Exit Sub
    MsgBox err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then Call CMDADDITEM_Click
End Sub

Private Sub Form_Load()

    On Error GoTo ErrHand
    
    Dim TRXMAST As ADODB.Recordset
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    ACT_FLAG = True
    ITEM_FLAG = True
    grdsales.ColWidth(0) = 500
    ''grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2700
    grdsales.ColWidth(3) = 1000
    grdsales.ColWidth(4) = 0
    grdsales.ColWidth(5) = 800
    grdsales.ColWidth(6) = 1200
    grdsales.ColWidth(7) = 800
    grdsales.ColWidth(8) = 800
    grdsales.ColWidth(9) = 800
    grdsales.ColWidth(10) = 1000
    grdsales.ColWidth(11) = 1000
    grdsales.ColWidth(12) = 1100
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(16) = 0
    grdsales.ColWidth(18) = 0
    grdsales.ColWidth(19) = 0
    grdsales.ColWidth(20) = 1100
    
    
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
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "TOTAL QTY"
    grdsales.TextArray(4) = "" '"UNIT"
    grdsales.TextArray(5) = "PACK"
    grdsales.TextArray(6) = "MRP"
    grdsales.TextArray(7) = "S. PRICE"
    grdsales.TextArray(8) = "COST"
    grdsales.TextArray(9) = "PTR"
    grdsales.TextArray(10) = "TAX %"
    grdsales.TextArray(11) = "BATCH"
    grdsales.TextArray(12) = "EXPIRY"
    grdsales.TextArray(13) = "SUB TOTAL"
    grdsales.TextArray(14) = "FREE"
    grdsales.TextArray(15) = "" '"TAX MODE"
    grdsales.TextArray(16) = "" '"Line No"
    grdsales.TextArray(17) = "P Disc"
    grdsales.TextArray(18) = "" '"MRP_BT"
     grdsales.TextArray(20) = "Cust Disc"

    PHYFLAG = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTDATE.text = Date
    TXTEXPDATE.Enabled = False
    txtBatch.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    TXTSLNO.text = 1
    TXTSLNO.Enabled = True
    'FRMECONTROLS.Enabled = False
    'FRMEMASTER.Enabled = False
    CLOSEALL = 1
    lblcredit.Caption = "1"
    TXTDEALER.text = ""
    M_ADD = False
    'Width = 15135
    'Height = 9660
    Left = 0
    Top = 0
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If ITEM_FLAG = False Then ITEM_REC.Close
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.cmdpurchase.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdsales_Click()
    On Error Resume Next
    lblPrftAmt.Caption = ""
    lblprftper.Caption = ""
    lblactprofit.Caption = ""
    TXTRATE.Tag = Val(grdsales.TextMatrix(grdsales.Row, 6)) - Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 20)) / 100
    If grdsales.rows > 1 Then
        lblprftper.Caption = Format(Round((((Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 3))) - Val(grdsales.TextMatrix(grdsales.Row, 13))) * 100) / (Val(grdsales.TextMatrix(grdsales.Row, 6)) * (Val(grdsales.TextMatrix(grdsales.Row, 3)))), 2), "0.00")
        lblPrftAmt.Caption = Format(Round((Val(grdsales.TextMatrix(grdsales.Row, 6)) * (Val(grdsales.TextMatrix(grdsales.Row, 3)))) - Val(grdsales.TextMatrix(grdsales.Row, 13)), 2), "0.00")
        lblactprofit.Caption = Format(Round((((Val(TXTRATE.Tag) * Val(grdsales.TextMatrix(grdsales.Row, 3))) - Val(grdsales.TextMatrix(grdsales.Row, 13))) * 100) / (Val(TXTRATE.Tag) * (Val(grdsales.TextMatrix(grdsales.Row, 3)))), 2), "0.00")
    End If

    lblitemname.Caption = Trim(grdsales.TextMatrix(grdsales.Row, 22))
    TxtItemName.Visible = False
    DataList1.Visible = False
    TXTsample.Visible = False
    TXTEXP.Visible = False
    On Error Resume Next
    grdsales.SetFocus
End Sub

Private Sub grdsales_GotFocus()
    lblPrftAmt.Caption = ""
    lblprftper.Caption = ""
    Call grdsales_Click
End Sub

Private Sub grdsales_RowColChange()
    Call grdsales_Click
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn

            TXTITEMCODE.text = grdtmp.Columns(0)
            TXTPRODUCT.text = grdtmp.Columns(1)
            
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Fram.Enabled = True
            TXTPRODUCT.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
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

Private Sub OptBranch_Click()
    flagchange.Caption = "0"
    Call TXTDEALER_Change
    TXTDEALER.SetFocus
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
                cmdadd.Enabled = True
                cmdadd.SetFocus
            ElseIf cmdadd.Enabled = True Then
                cmdadd.SetFocus
            End If
        Case vbKeyEscape
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
    End Select
End Sub

Private Sub OptOthers_Click()
    flagchange.Caption = "0"
    Call TXTDEALER_Change
    TXTDEALER.SetFocus
End Sub

Private Sub OPTTaxMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
'            If Val(TxttaxMRP.Text) <> 0 Then
'                If OPTTaxMRP.Value = False And OPTVAT.Value = False Then
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "PURCHASE"
'                    Exit Sub
'                End If
'            End If
            If TxttaxMRP.Enabled = True Then
                TxttaxMRP.Enabled = False
                cmdadd.Enabled = True
                cmdadd.SetFocus
            ElseIf cmdadd.Enabled = True Then
                cmdadd.SetFocus
            End If
        Case vbKeyEscape
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
    End Select
End Sub

Private Sub OPTVAT_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
        Case vbKeyReturn
'            If Val(TxttaxMRP.Text) <> 0 Then
'                If OPTTaxMRP.Value = False And OPTVAT.Value = False Then
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "PURCHASE"
'                    Exit Sub
'                End If
'            End If
            If TxttaxMRP.Enabled = True Then
                TxttaxMRP.Enabled = False
                cmdadd.Enabled = True
                cmdadd.SetFocus
            ElseIf cmdadd.Enabled = True Then
                cmdadd.SetFocus
            End If
        Case vbKeyEscape
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
    End Select

End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtBatch.text) = "" Then Exit Sub
            txtBatch.Enabled = False
            TXTEXPIRY.Visible = True
            TXTEXPIRY.SetFocus
        Case vbKeyEscape
            TXTFREE.Enabled = True
            txtBatch.Enabled = False
            TXTFREE.SetFocus
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTEXPDATE_GotFocus()
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
End Sub

Private Sub TXTEXPDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Len(Trim(TXTEXPDATE.text)) = 4 Then GoTo SKID
            If Not IsDate(TXTEXPDATE.text) Then Exit Sub
            If DateDiff("d", Date, TXTEXPDATE.text) < 0 Then
                MsgBox "Item Expired....", vbOKOnly, "PURCHASE.."
                TXTEXPDATE.SelStart = 0
                TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.text) < 60 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days", vbOKOnly, "PURCHASE.."
                TXTEXPDATE.SelStart = 0
                TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.text) < 180 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "PURCHASE..") = vbNo Then
                    TXTEXPDATE.SelStart = 0
                    TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
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
            If TXTEXPDATE.text = "  /  /    " Then GoTo SKIP
            If Not IsDate(TXTEXPDATE.text) Then Exit Sub
SKIP:
            txtBatch.Enabled = True
            TXTEXPDATE.Enabled = False
            TXTEXPIRY.Visible = False
            txtBatch.SetFocus
    End Select
End Sub

Private Sub TXTEXPDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
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
            If (Val(TXTQTY.text) = 0 And Val(TXTFREE.text) = 0) Then
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            End If
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
        Case Asc("'"), Asc("["), Asc("]")
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
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
            End If
        Case vbKeyEscape
            TXTINVOICE.SetFocus
    End Select
End Sub

Private Sub TXTINVDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTINVOICE_GotFocus()
    TXTINVOICE.SelStart = 0
    TXTINVOICE.SelLength = Len(TXTINVOICE.text)
End Sub

Private Sub TXTINVOICE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVOICE.text = "" Then Exit Sub
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub TXTINVOICE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("-")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub Txtpack_GotFocus()
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtpack.text) = 0 Then Exit Sub
            Txtpack.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            Txtpack.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub Txtpack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPRODUCT_Change()

    On Error GoTo ErrHand
    Set grdtmp.DataSource = Nothing
    If PHYFLAG = True Then
        PHY.Open "Select ITEM_CODE,ITEM_NAME,[CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & Trim(TXTPRODUCT.text) & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    Else
        PHY.Close
        PHY.Open "Select ITEM_CODE,ITEM_NAME,[CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & Trim(TXTPRODUCT.text) & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    End If
    
    Set grdtmp.DataSource = PHY
    grdtmp.Columns(0).Visible = False
    grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
    grdtmp.Columns(1).Width = 5000
    'grdtmp.Columns(2).Visible = False
    grdtmp.Columns(2).Caption = "QTY"
    grdtmp.Columns(2).Width = 1100
                          
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.text)
    FRMEGRDTMP.Visible = True
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Integer
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
            
        Case vbKeyReturn
        
            If Trim(TXTPRODUCT.text) = "" Then Exit Sub
            On Error Resume Next
            TXTPRODUCT.text = grdtmp.Columns(1)
            On Error GoTo ErrHand
            CmdDelete.Enabled = False
            TXTITEMCODE.text = ""
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT ITEM_CODE,ITEM_NAME,[CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & Trim(TXTPRODUCT.text) & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT ITEM_CODE,ITEM_NAME,[CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & Trim(TXTPRODUCT.text) & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.text = grdtmp.Columns(0)
                TXTPRODUCT.text = grdtmp.Columns(1)
                
                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    cmdadd.Enabled = True
                    cmdadd.SetFocus
                    Exit Sub
                End If
            ElseIf PHY.RecordCount > 1 Then
                FRMEGRDTMP.Visible = True
                Fram.Enabled = False
                grdtmp.Columns(0).Visible = False
                grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
                grdtmp.Columns(1).Width = 3000
                'grdtmp.Columns(2).Visible = False
                grdtmp.Columns(2).Caption = "QTY"
                grdtmp.Columns(2).Width = 1100
                grdtmp.SetFocus
            End If
            
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
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

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
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
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
    End Select
End Sub

Private Sub TXTPTR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPTR_LostFocus()
    'tXTptrdummy.Text = Format(Val(TXTPTR.Text) / Val(TXTUNIT.Text), ".000")
    TXTPTR.text = Format(TXTPTR.text, ".000")
    'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTPTR.Text) + Val(lbltaxamount.Caption), 2, ".000")))
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTQTY.Text) = 0 Then Exit Sub
            TXTQTY.Enabled = False
            TXTFREE.Enabled = True
            TXTFREE.SetFocus
        Case vbKeyEscape
            Txtpack.Enabled = True
            TXTQTY.Enabled = False
            Txtpack.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
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
End Sub

Private Sub TXTRATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTRATE.text) = 0 Then Exit Sub
            TXTRATE.Enabled = False
            TXTPTR.Enabled = True
            TXTPTR.SetFocus
         Case vbKeyEscape
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = True
            TXTEXPIRY.Visible = False
            TXTEXPDATE.SetFocus
    End Select
End Sub

Private Sub TXTRATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRATE_LostFocus()
    TXTRATE.text = Format(TXTRATE.text, ".000")
    txtmrpbt.text = 100 * Val(TXTRATE.text) / (100 + Val(TxttaxMRP.text))
End Sub

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVOICE.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then Exit Sub
            'If TXTINVOICE.Text = "" Then Exit Sub
            If Not IsDate(TXTINVDATE.text) Then Exit Sub
            
            Set rstTRXMAST = New ADODB.Recordset
            If OptOthers.Value = True Then
                rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_TYPE='PI' AND PINV = '" & Trim(TXTINVOICE.text) & "' AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
            Else
                rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_TYPE='TF' AND PINV = '" & Trim(TXTINVOICE.text) & "' AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
            End If
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                MsgBox "You have already entered this Invoice number for " & Trim(DataList2.text) & " as Computer Bill No. " & rstTRXMAST!VCH_NO, vbOKOnly, "Purchase Entry"
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                TXTINVOICE.SetFocus
                Exit Sub
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
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
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.text)
    lblitemname.Caption = ""
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            Exit Sub
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
                lblitemname.Caption = grdsales.TextMatrix(Val(TXTSLNO.text), 2)
                TXTPRODUCT.text = grdsales.TextMatrix(Val(TXTSLNO.text), 2)
                TXTQTY.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 14))
                Txtpack.text = grdsales.TextMatrix(Val(TXTSLNO.text), 5)
                TXTRATE.text = Format(Round(grdsales.TextMatrix(Val(TXTSLNO.text), 6), 2), "0.000")
                If Val(Txtpack.text) = 0 Then Txtpack.text = 1
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 9)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.text), 9) = 1
                TXTPTR.text = Format(Round(grdsales.TextMatrix(Val(TXTSLNO.text), 9) * Val(Txtpack.text), 2), "0.000")
                'TXTPTR.Text = Format((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))) * Val(Txtpack.Text), "0.000")

                txtBatch.text = grdsales.TextMatrix(Val(TXTSLNO.text), 11)
                TXTEXPDATE.text = IIf(grdsales.TextMatrix(Val(TXTSLNO.text), 12) = "", "  /  /    ", grdsales.TextMatrix(Val(TXTSLNO.text), 12))
                TXTEXPIRY.text = IIf(grdsales.TextMatrix(Val(TXTSLNO.text), 12) = "", "  /  ", Format(grdsales.TextMatrix(Val(TXTSLNO.text), 12), "mm/yy"))
                LBLSUBTOTAL.Caption = Format(Val(TXTQTY.text) * (Val(TXTPTR.text) + Val(lbltaxamount.Caption)), ".000")
                TXTFREE.text = grdsales.TextMatrix(Val(TXTSLNO.text), 14)
                TxttaxMRP.text = grdsales.TextMatrix(Val(TXTSLNO.text), 10)
                txtmrpbt.text = 100 * Val(TXTRATE.text) / (100 + Val(TxttaxMRP.text))
                txtPD.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 17))
                Txtdisccust.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) = "V" Then
                    OPTVAT.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) = "M" Then
                    OPTTaxMRP.Value = True
                Else
                    optnet.Value = True
                End If
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
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
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.text = Val(grdsales.rows)
                TXTPRODUCT.text = ""
                TXTITEMCODE.text = ""
                TXTQTY.text = ""
                TXTFREE.text = ""
                TxttaxMRP.text = ""
                txtPD.text = ""
                TXTRATE.text = ""
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
        Case Asc("'"), Asc("["), Asc("]")
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
            ''If Len(Trim(TXTEXPIRY.Text)) = 1 Then GoTo SKIP
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
            
            If DateDiff("d", Date, TXTEXPDATE.text) < 0 Then
                MsgBox "Item Expired....", vbOKOnly, "PURCHASE.."
                TXTEXPDATE.text = "  /  /    "
                TXTEXPIRY.SelStart = 0
                TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
                TXTEXPIRY.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.text) < 60 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days", vbOKOnly, "PURCHASE.."
                TXTEXPDATE.text = "  /  /    "
                TXTEXPIRY.SelStart = 0
                TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
                TXTEXPIRY.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.text) < 180 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "PURCHASE..") = vbNo Then
                    TXTEXPDATE.text = "  /  /    "
                    TXTEXPIRY.SelStart = 0
                    TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
                    TXTEXPIRY.SetFocus
                    Exit Sub
                End If
            End If
SKIP:
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            txtBatch.Enabled = True
            TXTEXPDATE.Enabled = False
            txtBatch.SetFocus
                        
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
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
        Case Asc("'"), Asc("["), Asc("]")
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
        LBLSUBTOTAL.Caption = Format(Val(TXTQTY.text) * Val(TXTPTR.text), ".000")

    Else
        If OPTTaxMRP.Value = True Then
            lbltaxamount.Caption = Val(txtmrpbt.text) * ((Val(TXTQTY.text) + Val(TXTFREE.text))) * Val(TxttaxMRP.text) / 100
            LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption), ".000")
            'lbltaxamount.Caption = (((Val(TXTRATE.Text) / Val(Txtpack.Text)))*VAL(TXTQTY.Text)+VAL(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100) * Val(Txtpack.Text)
        ElseIf OPTVAT.Value = True Then
            lbltaxamount.Caption = (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) * (Val(TXTQTY.text) + Val(TXTFREE.text))
            LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption), ".000")
        Else
            lbltaxamount.Caption = ""
            LBLSUBTOTAL.Caption = Format(Val(TXTQTY.text) * Val(TXTPTR.text), ".000")
        End If
    End If

    TxttaxMRP.text = Format(TxttaxMRP.text, "0.00")
    lbltaxamount.Caption = Format(lbltaxamount.Caption, "0.00")
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
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.text)) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), ".00")
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
        Case Asc("'"), Asc("["), Asc("]")
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
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Public Sub appendpurchase()
    
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTTRXFILE, RSTRTRXFILE, rststock As ADODB.Recordset
    Dim rstMaxNo, rstminus  As ADODB.Recordset
    Dim RSTLINK As ADODB.Recordset
    
    Dim M_DATA As Double
    Dim i As Integer
    
    'On Error GoTo eRRHAND
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    'db.Execute "delete From TRANSMAST WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    
    Dim TRXMAST As ADODB.Recordset
    Set TRXMAST = New ADODB.Recordset
    If OptOthers.Value = True Then
        TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    Else
        TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'TF'", db, adOpenStatic, adLockReadOnly
    End If
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    If OptOthers.Value = True Then
        db.Execute "delete FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'CR' AND INV_TRX_TYPE = 'PI'"
    End If
    
    For i = 1 To grdsales.rows - 1
        If Val(grdsales.TextMatrix(i, 5)) = 0 Then grdsales.TextMatrix(i, 5) = 1
        Set RSTRTRXFILE = New ADODB.Recordset
        If OptOthers.Value = True Then
            RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        Else
            RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='TF' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        End If
        db.BeginTrans
        If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
            RSTRTRXFILE.AddNew
            If OptOthers.Value = True Then
                RSTRTRXFILE!TRX_TYPE = "PI"
            Else
                RSTRTRXFILE!TRX_TYPE = "TF"
            End If
            RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTRTRXFILE!VCH_NO = Val(txtBillNo.text)
            RSTRTRXFILE!LINE_NO = Val(grdsales.TextMatrix(i, 16))
            RSTRTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(i, 1))
            RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
            RSTRTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With rststock
                If Not (.EOF And .BOF) Then
                    rststock!Category = IIf(IsNull(rststock!Category), "GENERAL", rststock!Category)
                    If (IsNull(!CLOSE_QTY)) Then !CLOSE_QTY = 0
                    !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(i, 13))
    
                    !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                    If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                    !RCPT_VAL = !RCPT_VAL + Val(grdsales.TextMatrix(i, 13))
    
                    !ITEM_COST = Val(grdsales.TextMatrix(i, 8))
                    !MRP = Val(grdsales.TextMatrix(i, 6))
                    !MRP_BT = Val(grdsales.TextMatrix(i, 18))
                    !SALES_TAX = Val(grdsales.TextMatrix(i, 10))
                    !UNIT = Val(grdsales.TextMatrix(i, 5))
                    !PTR = Val(grdsales.TextMatrix(i, 9))
                    !check_flag = "V" 'Trim(grdsales.TextMatrix(i, 15))  'MODE OF TAX
                    If Trim(grdsales.TextMatrix(i, 25)) <> "" Then !REMARKS = Trim(grdsales.TextMatrix(i, 25))
                    !CUST_DISC = Val(grdsales.TextMatrix(i, 20))
                    RSTRTRXFILE!MFGR = IIf(IsNull(!MANUFACTURER), "", !MANUFACTURER)
                    RSTRTRXFILE!Category = IIf(IsNull(!Category), "", !Category)
                    If Val(Val(grdsales.TextMatrix(i, 26))) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(i, 26))
                    If Val(Val(grdsales.TextMatrix(i, 27))) <> 0 Then !P_WS = Val(grdsales.TextMatrix(i, 27))
                    If Val(Val(grdsales.TextMatrix(i, 26))) <> 0 Then !P_CRTN = Round(Val(grdsales.TextMatrix(i, 26)) / Val(grdsales.TextMatrix(i, 5)), 3)
                    If Val(Val(grdsales.TextMatrix(i, 27))) <> 0 Then !P_LWS = Round(Val(grdsales.TextMatrix(i, 27)) / Val(grdsales.TextMatrix(i, 5)), 3)
                    !CRTN_PACK = 1

                    rststock.Update
                End If
            End With
            rststock.Close
            Set rststock = Nothing
    
        Else
            M_DATA = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
            M_DATA = M_DATA - (RSTRTRXFILE!QTY - RSTRTRXFILE!BAL_QTY)
            RSTRTRXFILE!BAL_QTY = M_DATA
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With rststock
                If Not (.EOF And .BOF) Then
                    rststock!Category = IIf(IsNull(rststock!Category), "GENERAL", rststock!Category)
                    If (IsNull(!CLOSE_QTY)) Then !CLOSE_QTY = 0
                    !CLOSE_QTY = !CLOSE_QTY - RSTRTRXFILE!QTY
                    !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(i, 13))
    
                    !RCPT_QTY = !RCPT_QTY - RSTRTRXFILE!QTY
                    !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                    If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                    !RCPT_VAL = !RCPT_VAL + Val(grdsales.TextMatrix(i, 13))
    
                    !ITEM_COST = Val(grdsales.TextMatrix(i, 8))
                    !MRP = Val(grdsales.TextMatrix(i, 6))
                    !MRP_BT = Val(grdsales.TextMatrix(i, 18))
                    !SALES_TAX = Val(grdsales.TextMatrix(i, 10))
                    !UNIT = Val(grdsales.TextMatrix(i, 5))
                    !PTR = Val(grdsales.TextMatrix(i, 9))
                    !check_flag = "V" 'Trim(grdsales.TextMatrix(i, 15))  'MODE OF TAX
                    If Trim(grdsales.TextMatrix(i, 25)) <> "" Then !REMARKS = Trim(grdsales.TextMatrix(i, 25))
                    !CUST_DISC = Val(grdsales.TextMatrix(i, 20))
                    RSTRTRXFILE!MFGR = IIf(IsNull(!MANUFACTURER), "", !MANUFACTURER)
                    RSTRTRXFILE!Category = IIf(IsNull(!Category), "", !Category)
                    If Val(Val(grdsales.TextMatrix(i, 26))) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(i, 26))
                    If Val(Val(grdsales.TextMatrix(i, 27))) <> 0 Then !P_WS = Val(grdsales.TextMatrix(i, 27))
                    If Val(Val(grdsales.TextMatrix(i, 26))) <> 0 Then !P_CRTN = Round(Val(grdsales.TextMatrix(i, 26)) / Val(grdsales.TextMatrix(i, 5)), 3)
                    If Val(Val(grdsales.TextMatrix(i, 27))) <> 0 Then !P_LWS = Round(Val(grdsales.TextMatrix(i, 27)) / Val(grdsales.TextMatrix(i, 5)), 3)
                    !CRTN_PACK = 1

                    rststock.Update
                End If
            End With
            rststock.Close
            Set rststock = Nothing
            RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
        End If
        RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTRTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(i, 13))
        RSTRTRXFILE!VCH_DATE = Format(Date, "dd/mm/yyyy")
        RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
        RSTRTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(i, 8))
        RSTRTRXFILE!ITEM_NET_COST_PRICE = Round((Val(grdsales.TextMatrix(i, 13)) / (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14)))), 3)
        RSTRTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 5))
        RSTRTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 5))
        RSTRTRXFILE!P_DISC = Val(grdsales.TextMatrix(i, 17))
        RSTRTRXFILE!MRP = Val(grdsales.TextMatrix(i, 6))
        RSTRTRXFILE!MRP_BT = Val(grdsales.TextMatrix(i, 18))
        RSTRTRXFILE!PTR = Val(grdsales.TextMatrix(i, 9))
        RSTRTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTRTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(i, 10))
        RSTRTRXFILE!CUST_DISC = Val(grdsales.TextMatrix(i, 20))
        RSTRTRXFILE!UNIT = 1
        'RSTRTRXFILE!VCH_DESC = "Received From " & DataList2.Text
        RSTRTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 11))
        RSTRTRXFILE!BARCODE = Trim(grdsales.TextMatrix(i, 28))
        'RSTRTRXFILE!ISSUE_QTY = 0
        RSTRTRXFILE!CST = 0
    
        RSTRTRXFILE!SCHEME = Val(grdsales.TextMatrix(i, 14))
        RSTRTRXFILE!FREE_QTY = 0
        RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTRTRXFILE!C_USER_ID = "SM"
    
        'RSTRTRXFILE!M_USER_ID = DataList2.BoundText
        RSTRTRXFILE!check_flag = "V" 'Trim(grdsales.TextMatrix(i, 15))  'MODE OF TAX
        'RSTRTRXFILE!PINV = Trim(TXTINVOICE.Text)
        RSTRTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 26))
        RSTRTRXFILE!P_CRTN = Round(Val(grdsales.TextMatrix(i, 26)) / Val(grdsales.TextMatrix(i, 5)), 3)
        RSTRTRXFILE!P_WS = Val(grdsales.TextMatrix(i, 27))
        RSTRTRXFILE!P_LWS = Round(Val(grdsales.TextMatrix(i, 27)) / Val(grdsales.TextMatrix(i, 5)), 3)
        RSTRTRXFILE!CRTN_PACK = 1
'        RSTRTRXFILE!GROSS_AMT = Val(grdsales.TextMatrix(i, 26))
        
        'RSTRTRXFILE!LOOSE_PACK = 1
        RSTRTRXFILE!PACK_TYPE = "Nos"
        RSTRTRXFILE!TR_DISC = 0
        RSTRTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(i, 4))
        'RSTRTRXFILE!VCH_DESC = "Received From " & DataList2.Text
        'RSTRTRXFILE!ISSUE_QTY = 0
        RSTRTRXFILE!CST = 0
        If Trim(grdsales.TextMatrix(i, 27)) = "P" Then
            RSTRTRXFILE!DISC_FLAG = "P"
        Else
            RSTRTRXFILE!DISC_FLAG = "A"
        End If
        RSTRTRXFILE!SCHEME = Val(grdsales.TextMatrix(i, 14))
        'RSTRTRXFILE!EXP_DATE = IIf(grdsales.TextMatrix(i, 12) = "", Null, Format(grdsales.TextMatrix(i, 12), "dd/mm/yyyy"))
        If IsDate(grdsales.TextMatrix(i, 12)) Then
            RSTRTRXFILE!EXP_DATE = IIf(grdsales.TextMatrix(i, 12) = "", Null, Format(grdsales.TextMatrix(i, 12), "dd/mm/yyyy"))
        End If
        RSTRTRXFILE!FREE_QTY = 0
        RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTRTRXFILE!C_USER_ID = "SM"
        'RSTRTRXFILE!M_USER_ID = DataList2.BoundText
        ''''RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 15))  'MODE OF TAX
        'RSTRTRXFILE!PINV = Trim(TXTINVOICE.Text)
        RSTRTRXFILE.Update
        db.CommitTrans
        RSTRTRXFILE.Close
        
        M_DATA = 0
        Set RSTRTRXFILE = Nothing
    Next i
       
    Set RSTTRXFILE = New ADODB.Recordset
    If OptOthers.Value = True Then
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    Else
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='TF' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    End If
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        db.BeginTrans
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        If OptOthers.Value = True Then
            RSTTRXFILE!TRX_TYPE = "PI"
        Else
            RSTTRXFILE!TRX_TYPE = "TF"
        End If
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = Trim(DataList2.text)
        RSTTRXFILE!VCH_AMOUNT = Val(lbltotalwodiscount.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!DISCOUNT = Val(TXTDISCAMOUNT.text)
        RSTTRXFILE!ADD_AMOUNT = Val(txtaddlamt.text)
        RSTTRXFILE!ROUNDED_OFF = 0
        RSTTRXFILE!OPEN_PAY = 0
        RSTTRXFILE!PAY_AMOUNT = 0
        RSTTRXFILE!REF_NO = ""
        RSTTRXFILE!SLSM_CODE = "CS"
        RSTTRXFILE!check_flag = "N"
        If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
        RSTTRXFILE!CFORM_NO = ""
        RSTTRXFILE!CFORM_DATE = Date
        RSTTRXFILE!REMARKS = Trim(DataList2.text)
        RSTTRXFILE!DISC_PERS = Val(txtcramt.text)
        'RSTTRXFILE!CST_PER = Val(TxtCST.Text)
        'RSTTRXFILE!INS_PER = Val(TxtInsurance.Text)
        RSTTRXFILE!LETTER_NO = 0
        RSTTRXFILE!LETTER_DATE = Date
        RSTTRXFILE!INV_MSGS = ""
        If Not IsDate(TXTDATE.text) Then TXTDATE.text = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!CREATE_DATE = Format(TXTDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!PINV = Trim(TXTINVOICE.text)
        RSTTRXFILE.Update
        db.CommitTrans
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
               
    i = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(CR_NO) From CRDTPYMT", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    If OptOthers.Value = True Then
        If lblcredit.Caption = "1" Then
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'CR' AND INV_TRX_TYPE = 'PI'", db, adOpenStatic, adLockOptimistic, adCmdText
            db.BeginTrans
            If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                RSTITEMMAST.AddNew
                RSTITEMMAST!TRX_TYPE = "CR"
                RSTITEMMAST!INV_TRX_TYPE = "PI"
                RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                RSTITEMMAST!CR_NO = i
                RSTITEMMAST!INV_NO = Val(txtBillNo.text)
                RSTITEMMAST!RCPT_AMOUNT = 0
            End If
            RSTITEMMAST!INV_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
            RSTITEMMAST!INV_AMT = Val(LBLTOTAL.Caption)
            If lblcredit.Caption = "0" Then
                RSTITEMMAST!check_flag = "Y"
                RSTITEMMAST!BAL_AMT = 0
            Else
                RSTITEMMAST!check_flag = "N"
                RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption) - RSTITEMMAST!RCPT_AMOUNT
            End If
            RSTITEMMAST!PINV = Trim(TXTINVOICE.text)
            RSTITEMMAST!ACT_CODE = DataList2.BoundText
            RSTITEMMAST!ACT_NAME = DataList2.text
            RSTITEMMAST.Update
            db.CommitTrans
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        End If
    End If
'    For i = 1 To grdsales.Rows - 1
'        db.Execute ("delete  FROM Tmporderlist WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'")
'        db.Execute ("delete  FROM NONRCVD WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'")
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
'            RSTLINK!UNIT = grdsales.TextMatrix(i, 5)
'            RSTLINK!Remarks = grdsales.TextMatrix(i, 4)
'            RSTLINK!ORD_QTY = 0
'            RSTLINK!CST = 0
'            RSTLINK!act_code = DataList2.BoundText
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
    
    'db.Execute "Update RTRXFILE set VCH_DATE = 'Y' WHERE ps_code = '" & MDIMAIN.StatusBar.Panels(8).Text & "' "
    
    Set RSTTRXFILE = New ADODB.Recordset
    If OptOthers.Value = True Then
        RSTTRXFILE.Open "SELECT * from RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    Else
        RSTTRXFILE.Open "SELECT * from RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='TF' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    End If
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(Trim(TXTINVDATE.text), "dd/mm/yyyy")
        RSTTRXFILE!VCH_DESC = "Received From " & Left(DataList2.text, 85)
        RSTTRXFILE!PINV = Trim(TXTINVOICE.text)
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
'
    i = 0
    
'    Dim slcount As Integer
'    For slcount = 1 To grdsales.Rows - 1
'        Set RSTITEMMAST = New ADODB.Recordset
'        RSTITEMMAST.Open "SELECT * from RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & grdsales.TextMatrix(slcount, 1) & "'  AND RTRXFILE.BAL_QTY > 0 ", db, adOpenForwardOnly
'        Do Until RSTITEMMAST.EOF
'            i = 0
'            Set rstMaxRec = New ADODB.Recordset
'            rstMaxRec.Open "SELECT * from RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND RTRXFILE.REF_NO = '" & RSTITEMMAST!REF_NO & "' AND RTRXFILE.EXP_DATE = # " & RSTITEMMAST!EXP_DATE & " # AND RTRXFILE.BAL_QTY > 0 ORDER BY RTRXFILE.VCH_NO", db, adOpenStatic, adLockOptimistic, adCmdText
'            '[EXP_DATE] <=# " & E_DATE & " #
'            If rstMaxRec.RecordCount > 1 Then
'                Do Until rstMaxRec.EOF
'                    i = i + rstMaxRec!BAL_QTY
'                    rstMaxRec!BAL_QTY = 0
'                    rstMaxRec.Update
'                    rstMaxRec.MoveNext
'                Loop
'                rstMaxRec.MoveLast
'                rstMaxRec!BAL_QTY = i
'                rstMaxRec.Update
'            End If
'            rstMaxRec.Close
'            Set rstMaxRec = Nothing
'            RSTITEMMAST.MoveNext
'        Loop
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
'    Next slcount
    
SKIP:
    Set rstMaxNo = New ADODB.Recordset
    If OptOthers.Value = True Then
        rstMaxNo.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    Else
        rstMaxNo.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'TF'", db, adOpenStatic, adLockReadOnly
    End If
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    'cmdRefresh.Enabled = False
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = True
    DataList2.text = ""
    TXTINVDATE.text = "  /  /    "
    TXTINVOICE.text = ""
    TXTREMARKS.text = ""
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TXTPRODUCT.text = ""
    Txtpack.text = ""
    TXTQTY.text = ""
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    txtPD.text = ""
    txtBatch.text = ""
    TXTRATE.text = ""
    txtmrpbt.text = ""
    TXTPTR.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    txtaddlamt.text = ""
    txtcramt.text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LblSale_Val.Caption = ""
    LblProfittotal.Caption = ""
    TXTDISCAMOUNT.text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    flagchange2.Caption = ""
    TXTDEALER.text = ""
    lbldealer.Caption = ""
    LBLitem.Caption = ""
    lblPrftAmt.Caption = ""
    lblprftper.Caption = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    FRMEMASTER.Visible = False
    Frame2.Visible = False
    FRMECONTROLS.Visible = False
    Frmmain.Enabled = True
    TXTINVOICE.SetFocus
    M_ADD = False
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "PURCHASE ENTRY"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub


Private Sub txtaddlamt_GotFocus()
    txtaddlamt.SelStart = 0
    txtaddlamt.SelLength = Len(txtaddlamt.text)
End Sub

Private Sub txtaddlamt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
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
        MsgBox "Discount Amount More than Bill Amount", , "PURCHASE..."
        txtaddlamt.SelStart = 0
        txtaddlamt.SelLength = Len(txtaddlamt.text)
        txtaddlamt.SetFocus
        Exit Sub
    End If
    txtaddlamt.text = Format(txtaddlamt.text, ".00")
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.text)) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), ".00")
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
        Case Asc("'"), Asc("["), Asc("]")
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
        MsgBox "Credit Note Amount More than Bill Amount", , "PURCHASE..."
        txtcramt.SelStart = 0
        txtcramt.SelLength = Len(txtcramt.text)
        txtcramt.SetFocus
        Exit Sub
    End If
    txtcramt.text = Format(txtcramt.text, ".00")
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.text)) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), ".00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    txtcramt.SetFocus
End Sub

Private Sub OPTTaxMRP_GotFocus()
    lbltaxamount.Caption = Val(txtmrpbt.text) * (Val(TXTQTY.text) + Val(TXTFREE.text)) * Val(TxttaxMRP.text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption), ".000")
End Sub

Private Sub OPTVAT_GotFocus()
    lbltaxamount.Caption = (Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100) * (Val(TXTQTY.text) + Val(TXTFREE.text))
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption), ".000")
End Sub

Private Sub OPTNET_GotFocus()
    lbltaxamount.Caption = ""
    LBLSUBTOTAL.Caption = Format(Val(TXTQTY.text) * Val(TXTPTR.text), ".000")
End Sub

Private Sub txtPD_GotFocus()
    txtPD.SelStart = 0
    txtPD.SelLength = Len(txtPD.text)
End Sub

Private Sub txtPD_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtPD.Enabled = False
            Txtdisccust.Enabled = True
            Txtdisccust.SetFocus
         Case vbKeyEscape
            txtPD.Enabled = False
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
    End Select
End Sub

Private Sub txtPD_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPD_LostFocus()
    Call TxttaxMRP_LostFocus
    txtPD.Tag = ((Val(LBLSUBTOTAL.Caption) - Val(lbltaxamount.Caption)) * Val(txtPD.text) / 100)
    LBLSUBTOTAL.Caption = Format(Val(LBLSUBTOTAL.Caption) - Val(txtPD.Tag), ".000")
    txtPD.text = Format(txtPD.text, "0.00")
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            If OptOthers.Value = True Then
                ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            Else
                ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='411')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            End If
            ACT_FLAG = False
        Else
            ACT_REC.Close
            If OptOthers.Value = True Then
                ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            Else
                ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='411')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            End If
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ACT_NAME
        End If
        Set DataList2.RowSource = ACT_REC
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
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            TXTINVOICE.SetFocus
    End Select
End Sub

Private Sub TXTDEALER_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
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
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Purchase Bill..."
                DataList2.SetFocus
                Exit Sub
            End If
            CmdLoadInv.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
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

Private Sub Txtdisccust_GotFocus()
    Txtdisccust.SelStart = 0
    Txtdisccust.SelLength = Len(Txtdisccust.text)
End Sub

Private Sub Txtdisccust_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txtdisccust.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            Txtdisccust.Enabled = False
            txtPD.Enabled = True
            txtPD.SetFocus
    End Select
End Sub

Private Sub Txtdisccust_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtdisccust_LostFocus()
    Txtdisccust.text = Format(Txtdisccust.text, "0.00")
End Sub

Private Sub TxtItemName_Change()
    On Error GoTo ErrHand
    If flagchange2.Caption <> "1" Then
        If ITEM_FLAG = True Then
            ITEM_REC.Open "Select DISTINCT ITEM_CODE,ITEM_NAME From ITEMMAST  WHERE ITEM_NAME Like '" & TxtItemName.text & "%'ORDER BY ITEM_NAME", db, adOpenForwardOnly
            ITEM_FLAG = False
        Else
            ITEM_REC.Close
            ITEM_REC.Open "Select DISTINCT ITEM_CODE,ITEM_NAME From ITEMMAST  WHERE ITEM_NAME Like '" & TxtItemName.text & "%'ORDER BY ITEM_NAME", db, adOpenForwardOnly
            ITEM_FLAG = False
        End If
        If (ITEM_REC.EOF And ITEM_REC.BOF) Then
            LBLitem.Caption = ""
        Else
            LBLitem.Caption = ITEM_REC!ITEM_NAME
        End If
        Set DataList1.RowSource = ITEM_REC
        DataList1.ListField = "ITEM_NAME"
        DataList1.BoundColumn = "ITEM_CODE"
    End If
    Exit Sub
ErrHand:
    MsgBox err.Description
    'TxtItemName.Text = ""
    If ITEM_FLAG = True Then
        ITEM_FLAG = False
    Else
        ITEM_FLAG = True
    End If
    TxtItemName.text = ""
End Sub

Private Sub TxtItemName_GotFocus()
    TxtItemName.SelStart = Len(TxtItemName.text)
    'TxtItemName.SelLength = Len(TxtItemName.Text)
End Sub

Private Sub TxtItemName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList1.VisibleCount = 0 Then Exit Sub
            DataList1.SetFocus
        Case vbKeyEscape
            TxtItemName.Visible = False
            DataList1.Visible = False
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

Private Sub DataList1_Click()
    TxtItemName.text = DataList1.text
    LBLitem.Caption = TxtItemName.text
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList1.text = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Purchase Bill..."
                DataList1.SetFocus
                Exit Sub
            End If
            grdsales.TextMatrix(grdsales.Row, 1) = DataList1.BoundText
            grdsales.TextMatrix(grdsales.Row, 2) = DataList1.text
            
            Dim rstTRXMAST As ADODB.Recordset
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From ITEMMAST WHERE ITEM_CODE = '" & DataList1.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                grdsales.TextMatrix(grdsales.Row, 20) = IIf(IsNull(rstTRXMAST!CUST_DISC), 0, rstTRXMAST!CUST_DISC)
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
        
            TxtItemName.Visible = False
            DataList1.Visible = False
            grdsales.SetFocus
        Case vbKeyEscape
            TxtItemName.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_GotFocus()
    flagchange2.Caption = 1
    TxtItemName.text = LBLitem.Caption
    DataList1.text = TxtItemName.text
    Call DataList1_Click
End Sub

Private Sub DataList1_LostFocus()
     flagchange2.Caption = ""
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Integer
    If grdsales.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 2
                    TxtItemName.Visible = True
                    DataList1.Visible = True
                    DataList1.Top = grdsales.CellTop + 510
                    DataList1.Left = grdsales.CellLeft '+ 50
                    TxtItemName.Top = grdsales.CellTop + 100
                    TxtItemName.Left = grdsales.CellLeft '+ 50
                    TxtItemName.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TxtItemName.Height = grdsales.CellHeight
                    TxtItemName.SetFocus
                Case 5, 20, 11
                    TXTsample.Visible = True
                    TXTsample.Top = grdsales.CellTop + 110
                    TXTsample.Left = grdsales.CellLeft + 50
                    TXTsample.Width = grdsales.CellWidth
                    TXTsample.Height = grdsales.CellHeight
                    TXTsample.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    TXTsample.SetFocus
                Case 12  ' EXPIRY
                    TXTEXP.Visible = True
                    TXTEXP.Top = grdsales.CellTop + 110
                    TXTEXP.Left = grdsales.CellLeft + 50
                    TXTEXP.Width = grdsales.CellWidth '- 25
                    TXTEXP.text = IIf(IsDate(grdsales.TextMatrix(grdsales.Row, grdsales.Col)), Format(grdsales.TextMatrix(grdsales.Row, grdsales.Col), "MM/YY"), "  /  ")
                    TXTEXP.SetFocus
            End Select
        Case 114
            sitem = UCase(InputBox("Item Name..?", "Purchase"))
            For i = 1 To grdsales.rows - 1
                    If Mid(grdsales.TextMatrix(i, 2), 1, Len(sitem)) = sitem Then
                        grdsales.Row = i
                        grdsales.TopRow = i
                    Exit For
                End If
            Next i
            grdsales.SetFocus
    End Select
End Sub

Private Sub grdsales_Scroll()
    TxtItemName.Visible = False
    DataList1.Visible = False
    TXTsample.Visible = False
    TXTEXP.Visible = False
    grdsales.SetFocus
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim M_STOCK As Integer
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 5   'Pack
                    If Val(TXTsample.text) = 0 Then Exit Sub
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = TXTsample.text
                    grdsales.TextMatrix(grdsales.Row, 8) = (Val(grdsales.TextMatrix(grdsales.Row, 8)) * Val(grdsales.TextMatrix(grdsales.Row, 5))) / Val(TXTsample.text)
                    grdsales.TextMatrix(grdsales.Row, 9) = (Val(grdsales.TextMatrix(grdsales.Row, 9)) * Val(grdsales.TextMatrix(grdsales.Row, 5))) / Val(TXTsample.text)
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                Case 11   'Batch
                    If Val(TXTsample.text) = 0 Then Exit Sub
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = TXTsample.text
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
                Case 20   'Cust Disc
                    'If Val(TXTsample.Text) = 0 Then Exit Sub
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = TXTsample.text
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
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case grdsales.Col
'        Case 5
'             Select Case KeyAscii
'                Case Asc("'"), Asc("["), Asc("]")
'                    KeyAscii = 0
'                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                Case Else
'                    KeyAscii = 0
'            End Select
        Case 5, 20
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Function Import_Bill()
    Dim M_DATE As Date
    Dim D As Integer
    Dim MON As Integer
    Dim Y As Integer
    Dim xlApp As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim var As Variant
    Dim i As Long
    Dim PR_CODE As String
    
    Dim rstTRXMAST, RSTITEMMAST As ADODB.Recordset
    
    Set rstTRXMAST = New ADODB.Recordset
    rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_TYPE='PI' AND PINV = '" & Trim(TXTINVOICE.text) & "' AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
        MsgBox "You have already entered this Invoice number for " & Trim(DataList2.text) & " as Computer Bill No. " & rstTRXMAST!VCH_NO, vbOKOnly, "Purchase Entry"
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
        FRMEMASTER.Enabled = True
        TXTINVOICE.SetFocus
        Exit Function
    End If
    rstTRXMAST.Close
    Set rstTRXMAST = Nothing
    
    On Error GoTo errHandler
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "Excel Files (*.xls*)|*.xls*"
    CommonDialog1.ShowOpen
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    
    Screen.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    
    'Set wb = xlAppr.Workbooks.Open("PATH TO YOUR EXCEL FILE")
    Set wb = xlApp.Workbooks.Open(CommonDialog1.FileName)
    Sleep (5000)
    Set ws = wb.Worksheets("Sheet1") 'Specify your worksheet name
    var = ws.Range("A1").Value
    
    On Error Resume Next
    lbltotalwodiscount.Caption = ""
    LblSale_Val.Caption = ""
    'TXTDISCAMOUNT.Text = ws.Range("T2").Value + ws.Range("T2").Value * ws.Range("P2").Value
    'LBLTOTAL.Caption = ws.Range("B" & T).Value
    TXTDATE.text = Format(Date, "DD/MM/YYYY")
    On Error Resume Next
    TXTINVDATE.text = Format(Date, "DD/MM/YYYY") 'Format(ws.Range("S2").value, "DD/MM/YYYY")
    'TXTINVDATE.Text = Right(ws.Range("R2").Value, 2) & "/" & Mid(ws.Range("R2").Value, 5, 2) & "/" & Mid(ws.Range("R2").Value, 1, 4)
    'TXTDEALER.Text = ws.Range("W2").Value
    TXTREMARKS.text = ""
    Dim n, M As Integer
    n = 0
    grdsales.FixedRows = 0
    grdsales.rows = 1
    For i = 2 To 5000
        If ws.Range("A" & i).Value = "" Then Exit For
        M = 0
        n = n + 1
        grdsales.rows = grdsales.rows + 1
        grdsales.FixedRows = 1
        grdsales.TextMatrix(n, 0) = n
        'grdsales.TextMatrix(n, 1) = "" 'rstTRXMAST!ITEM_CODE
        grdsales.TextMatrix(n, 2) = Trim(ws.Range("B" & i).Value) 'ITEM NAME
        If InStr(Trim(grdsales.TextMatrix(n, 2)), "'") <> 0 Then
            grdsales.TextMatrix(n, 2) = Trim(Mid(grdsales.TextMatrix(n, 2), 1, Len(grdsales.TextMatrix(n, 2)) - 4))
        Else
            grdsales.TextMatrix(n, 2) = Trim(grdsales.TextMatrix(n, 2))
        End If
        grdsales.TextMatrix(n, 2) = Replace(grdsales.TextMatrix(n, 2), "  ", " ")
        grdsales.TextMatrix(n, 2) = Replace(grdsales.TextMatrix(n, 2), "#", "")
        grdsales.TextMatrix(n, 2) = Replace(grdsales.TextMatrix(n, 2), "$", "")
        grdsales.TextMatrix(n, 2) = Replace(grdsales.TextMatrix(n, 2), "'", "")
        
        grdsales.TextMatrix(n, 22) = grdsales.TextMatrix(n, 2) 'ITEM NAME
        
        grdsales.TextMatrix(n, 1) = ""
        On Error Resume Next
        Set rstTRXMAST = New ADODB.Recordset
        rstTRXMAST.Open "Select * From ITEMMAST WHERE ITEM_NAME Like '%" & Trim(grdsales.TextMatrix(n, 2)) & "%' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'            Set RSTITEMMAST = New ADODB.Recordset
'            RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
'            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'                PR_CODE = IIf(IsNull(RSTITEMMAST.Fields(0)), 1, Val(RSTITEMMAST.Fields(0)) + 1)
'            End If
'            RSTITEMMAST.Close
'            Set RSTITEMMAST = Nothing
'
'            Set RSTITEMMAST = New ADODB.Recordset
'            RSTITEMMAST.Open "SELECT * FROM ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
'            RSTITEMMAST.AddNew
'            RSTITEMMAST!ITEM_CODE = PR_CODE
'            RSTITEMMAST!ITEM_NAME = Trim(grdsales.TextMatrix(n, 2))
'            RSTITEMMAST!Category = "GENERAL"
'            RSTITEMMAST!UNIT = 1
'            RSTITEMMAST!MANUFACTURER = ws.Range("U" & i).Value
'            RSTITEMMAST!Remarks = ""
'            RSTITEMMAST!REORDER_QTY = 1
'            RSTITEMMAST!BIN_LOCATION = ""
'            RSTITEMMAST!ITEM_COST = 0
'            RSTITEMMAST!MRP = 0
'            RSTITEMMAST!SALES_TAX = 0
'            RSTITEMMAST!PTR = 0
'            RSTITEMMAST!CST = 0
'            RSTITEMMAST!OPEN_QTY = 0
'            RSTITEMMAST!OPEN_VAL = 0
'            RSTITEMMAST!RCPT_QTY = 0
'            RSTITEMMAST!RCPT_VAL = 0
'            RSTITEMMAST!ISSUE_QTY = 0
'            RSTITEMMAST!ISSUE_VAL = 0
'            RSTITEMMAST!CLOSE_QTY = 0
'            RSTITEMMAST!CLOSE_VAL = 0
'            RSTITEMMAST!DAM_QTY = 0
'            RSTITEMMAST!DAM_VAL = 0
'            RSTITEMMAST!DISC = 0
'            RSTITEMMAST.Update
'            RSTITEMMAST.Close
'            Set RSTITEMMAST = Nothing
'
'            grdsales.TextMatrix(n, 1) = PR_CODE
                
        Else
            grdsales.TextMatrix(n, 1) = rstTRXMAST!ITEM_CODE
            grdsales.TextMatrix(n, 20) = IIf(IsNull(rstTRXMAST!CUST_DISC), 0, rstTRXMAST!CUST_DISC)
            grdsales.TextMatrix(n, 5) = IIf(IsNull(rstTRXMAST!UNIT), 0, rstTRXMAST!UNIT)
        End If
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
        
        'grdsales.TextMatrix(n, 3) = ""
'        grdsales.TextMatrix(N, 5) = 1
        On Error Resume Next
'        If InStr(UCase(ws.Range("D" & i).value), "'S") = 0 Then
'            grdsales.TextMatrix(N, 5) = Int(Replace(ws.Range("D" & i).value, "'", ""))
'            grdsales.TextMatrix(N, 5) = Int(Replace(ws.Range("D" & i).value, "S", ""))
'            grdsales.TextMatrix(N, 5) = Int(Replace(ws.Range("D" & i).value, "s", ""))
'        Else
'            M = InStr(ws.Range("D" & i).value, "'")
'            grdsales.TextMatrix(N, 5) = Mid(ws.Range("D" & i).value, 1, M - 1)
'        End If
        'If Val(grdsales.TextMatrix(N, 5)) <= 0 Then grdsales.TextMatrix(N, 5) = 1  'PACK
        'On Error GoTo errhandler
        
        grdsales.TextMatrix(n, 5) = 1  'PACK
        grdsales.TextMatrix(n, 3) = Val(ws.Range("E" & i).Value) '+ Val(ws.Range("G" & i).value) 'QTY  + FREE
        grdsales.TextMatrix(n, 4) = 1
        grdsales.TextMatrix(n, 6) = Format(Val(ws.Range("D" & i).Value), ".000") 'MRP
        grdsales.TextMatrix(n, 7) = Format(Val(ws.Range("D" & i).Value), ".000") 'MRP / PACK 'Format(Val(ws.Range("J" & i).value) / Val(grdsales.TextMatrix(N, 5)), ".000")
'        If Val(ws.Range("H" & i).value) = 0 Then
'            grdsales.TextMatrix(n, 9) = Format(Val(ws.Range("G" & i).value) / Val(grdsales.TextMatrix(n, 5)), ".000")
'            grdsales.TextMatrix(n, 10) = Val(ws.Range("J" & i).value) 'IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
'            grdsales.TextMatrix(n, 15) = "V" 'IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
'        ElseIf Val(ws.Range("H" & i).value) = 4.77 Then
'            grdsales.TextMatrix(n, 9) = Format(Val(ws.Range("G" & i).value) / Val(grdsales.TextMatrix(n, 5)), ".000")
'            grdsales.TextMatrix(n, 10) = "5"
'            grdsales.TextMatrix(n, 15) = "M" 'IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
'        Else
'            grdsales.TextMatrix(n, 9) = Format(Val(ws.Range("I" & i).value) / Val(grdsales.TextMatrix(n, 5)), ".000")
'            grdsales.TextMatrix(n, 10) = Val(ws.Range("J" & i).value) 'IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
'            grdsales.TextMatrix(n, 15) = "N" 'IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
'        End If
        grdsales.TextMatrix(n, 9) = Format(Val(ws.Range("F" & i).Value), ".000")  'RATE / Pack  Format(Val(ws.Range("H" & i).value) / Val(grdsales.TextMatrix(N, 5)), ".000")
        grdsales.TextMatrix(n, 10) = Val(ws.Range("H" & i).Value) 'TAX * 2 'IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
        grdsales.TextMatrix(n, 15) = "V" 'IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
            
        grdsales.TextMatrix(n, 11) = Trim(ws.Range("C" & i).Value) ' BATCH
        grdsales.TextMatrix(n, 12) = "" ' EXPIRY
'        If Len(Trim(ws.Range("E" & i).value)) = 4 Then
'            MON = Val(Mid(Trim(ws.Range("F" & i).value), 1, 2))
'            Y = Val(Right(Trim(ws.Range("F" & i).value), 2))
'            Y = 2000 + Y
'            M_DATE = "01" & "/" & MON & "/" & Y
'            D = LastDayOfMonth(M_DATE)
'            M_DATE = D & "/" & MON & "/" & Y
'            grdsales.TextMatrix(N, 12) = M_DATE
'        Else
'            If IsDate(Trim(ws.Range("E" & i).value)) Then
'                grdsales.TextMatrix(N, 12) = Format(Trim(ws.Range("E" & i).value), "dd/mm/yyyy")
'            Else
'                grdsales.TextMatrix(N, 12) = ""
'            End If
'        End If
        
        grdsales.TextMatrix(n, 14) = 0 ' FREE Val(ws.Range("G" & i).value)
        grdsales.TextMatrix(n, 17) = Val(ws.Range("G" & i).Value) 'DISC IIf(IsNull(rstTRXMAST!P_DISC), 0, rstTRXMAST!P_DISC)
        grdsales.TextMatrix(n, 18) = Round((100 * Val(grdsales.TextMatrix(n, 6)) / (100 + Val(grdsales.TextMatrix(n, 10))) / Val(grdsales.TextMatrix(n, 5))), 3)
        txtmrpbt.Tag = Round(100 * Val(grdsales.TextMatrix(n, 6)) / (100 + Val(grdsales.TextMatrix(n, 10))), 3)
        Txtdisccust.Tag = (Val(grdsales.TextMatrix(n, 9)) - (Val(grdsales.TextMatrix(n, 9)) * Val(grdsales.TextMatrix(n, 17)) / 100)) * Val(grdsales.TextMatrix(n, 3))
        lbltaxamount.Tag = Val(Txtdisccust.Tag) * Val(grdsales.TextMatrix(n, 10)) / 100
        grdsales.TextMatrix(n, 13) = Round(Val(Txtdisccust.Tag) + Val(lbltaxamount.Tag), 3)
        
'        Txtdisccust.Tag = ((Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 5))) * Val(grdsales.TextMatrix(N, 17)) / 100) * (Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14)))
'        If Val(grdsales.TextMatrix(N, 10)) = 0 Then
'            grdsales.TextMatrix(N, 18) = 0
'            grdsales.TextMatrix(N, 13) = Round((Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))) * (Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 5))) - Val(Txtdisccust.Tag), 3)
'        Else
'            If grdsales.TextMatrix(N, 15) = "M" Then
'                lbltaxamount.Tag = (Val(txtmrpbt.Tag) * Val(grdsales.TextMatrix(N, 3))) * Val(grdsales.TextMatrix(N, 10)) / 100
'                grdsales.TextMatrix(N, 13) = Round(((Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))) * (Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 5)))) + Val(lbltaxamount.Tag) - Val(Txtdisccust.Tag), 3)
'            ElseIf grdsales.TextMatrix(N, 15) = "V" Then
'                'T2 -(T2 * AC2 / 100)
'                Txtdisccust.Tag = (Val(ws.Range("H" & i).value) - (Val(ws.Range("H" & i).value) * Val(grdsales.TextMatrix(N, 17)) / 100)) * Val(Val(ws.Range("M" & i).value))
'                lbltaxamount.Tag = Val(Txtdisccust.Tag) * Val(grdsales.TextMatrix(N, 10)) / 100
'                'lbltaxamount.Tag = (((Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 5)) * Val(grdsales.TextMatrix(N, 10))) - Val(Txtdisccust.Tag)) / 100) * (Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14)))
'                grdsales.TextMatrix(N, 13) = Round(Val(Txtdisccust.Tag) + Val(lbltaxamount.Tag), 3)
'            Else
'                lbltaxamount.Caption = Round((Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 10)) / 100) * (Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))), 3)
'                grdsales.TextMatrix(N, 13) = Format(Round(((Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))) * Val(grdsales.TextMatrix(N, 9))) + Val(lbltaxamount.Caption), 3), "0.000")
'                 'bltaxamount.Tag = Val(txtmrpbt.Tag) * (grdsales.TextMatrix(n, 14)) * 5 / 100
'                'grdsales.TextMatrix(n, 13) = Round(((Val(grdsales.TextMatrix(n, 3)) - Val(grdsales.TextMatrix(n, 14))) * (Val(grdsales.TextMatrix(n, 9)) * Val(grdsales.TextMatrix(n, 5)))) + Val(lbltaxamount.Tag) - Val(Txtdisccust.Tag), 3)
'            End If
'        End If
        'Format(Round((Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.Text)) / Val(TxtPack.Text), 3), ".000")
        If Val(grdsales.TextMatrix(n, 5)) = 0 Then grdsales.TextMatrix(n, 5) = 1
        If Val(grdsales.TextMatrix(n, 3)) - Val(grdsales.TextMatrix(n, 14)) = 0 Then
            grdsales.TextMatrix(n, 8) = "0.00"
        Else
            'grdsales.TextMatrix(N, 8) = Format(Round((Val(grdsales.TextMatrix(N, 13)) / Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))) / Val(grdsales.TextMatrix(N, 5)), 3), ".000")
            grdsales.TextMatrix(n, 8) = Val(Txtdisccust.Tag) / Val(grdsales.TextMatrix(n, 3))
        End If
        'grdsales.TextMatrix(n, 8) = Format(Val(grdsales.TextMatrix(n, 13)) / (Val(grdsales.TextMatrix(n, 3)) - Val(grdsales.TextMatrix(n, 14))) / Val(grdsales.TextMatrix(n, 5)), ".000")
        'if Val(grdsales.TextMatrix(grdsales.Row, 8))= 0 then grdsales.TextMatrix(grdsales.Row, 8)
        'Format(Round((Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.Text)) / Val(TxtPack.Text), 3), ".000")
        
        'grdsales.TextMatrix(n, 8) = Format(Val(grdsales.TextMatrix(n, 13)) / (Val(grdsales.TextMatrix(n, 3)) - Val(grdsales.TextMatrix(n, 14))) / Val(grdsales.TextMatrix(n, 5)), ".000")
        grdsales.TextMatrix(n, 16) = n 'rstTRXMAST!LINE_NO
        grdsales.TextMatrix(n, 23) = "" 'Trim(ws.Range("B" & i).value)
        grdsales.TextMatrix(n, 24) = "" 'Left(Trim(ws.Range("W" & i).value), 20)
        grdsales.TextMatrix(n, 25) = Left(Trim(ws.Range("A" & i).Value), 8) 'HSN
        grdsales.TextMatrix(n, 26) = Val(ws.Range("J" & i).Value) 'R. RATE
        grdsales.TextMatrix(n, 27) = Val(ws.Range("K" & i).Value) 'W. RATE
    Next i
    'or
    var = ws.Cells(1, 1).Value
    wb.Close
    
    xlApp.Quit
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    
'    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
'    LblProfittotal.Caption = Format(Val(LblSale_Val.Caption) - Val(LBLTOTAL.Caption), ".00")
'
    FRMEMASTER.Visible = True
    Frame2.Visible = True
    FRMECONTROLS.Visible = True
    Frmmain.Enabled = False
    
    'TXTSLNO.SetFocus
    Screen.MousePointer = vbNormal
    Exit Function
errHandler:
    Screen.MousePointer = vbNormal
    If err.Number = 9 Then
        MsgBox "NO SUCH INVOICE PRESENT!!", vbOKOnly, "PURCHASE"
        wb.Close
        xlApp.Quit
        Set ws = Nothing
        Set wb = Nothing
        Set xlApp = Nothing
    ElseIf err.Number = 32755 Then
        
    Else
        MsgBox err.Description
    End If
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
End Function

Private Sub TXTEXP_GotFocus()
    TXTEXP.SelStart = 0
    TXTEXP.SelLength = Len(TXTEXP.text)
End Sub

Private Sub TXTEXP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim M_DATE As Date
    Dim D As Integer
    Dim M As Integer
    Dim Y As Integer
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(Mid(TXTEXP.text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXP.text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXP.text, 4, 5)) = 0 Then Exit Sub
            
            M = Val(Mid(TXTEXP.text, 1, 2))
            Y = Val(Right(TXTEXP.text, 2))
            Y = 2000 + Y
            M_DATE = "01" & "/" & M & "/" & Y
            D = LastDayOfMonth(M_DATE)
            M_DATE = D & "/" & M & "/" & Y
            
            TXTEXP.Visible = False
            grdsales.TextMatrix(grdsales.Row, grdsales.Col) = M_DATE
            grdsales.Enabled = True
    
        Case vbKeyEscape
            TXTEXP.Visible = False
            grdsales.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Function Import_Br_Bill()
    Dim M_DATE As Date
    Dim D As Integer
    Dim MON As Integer
    Dim Y As Integer
    Dim xlApp As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim var As Variant
    Dim i As Long
    Dim PR_CODE As String
    
    Dim rstTRXMAST, RSTITEMMAST As ADODB.Recordset
    
    Set rstTRXMAST = New ADODB.Recordset
    rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_TYPE='TF' AND PINV = '" & Trim(TXTINVOICE.text) & "' AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
        MsgBox "You have already entered this Invoice number for " & Trim(DataList2.text) & " as Computer Bill No. " & rstTRXMAST!VCH_NO, vbOKOnly, "Purchase Entry"
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
        FRMEMASTER.Enabled = True
        TXTINVOICE.SetFocus
        Exit Function
    End If
    rstTRXMAST.Close
    Set rstTRXMAST = Nothing
    
    On Error GoTo errHandler
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "Excel Files (*.xls*)|*.xls*"
    CommonDialog1.ShowOpen
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    
    Screen.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    
    'Set wb = xlAppr.Workbooks.Open("PATH TO YOUR EXCEL FILE")
    Set wb = xlApp.Workbooks.Open(CommonDialog1.FileName)
    Sleep (5000)
    Set ws = wb.Worksheets("Sheet1") 'Specify your worksheet name
    var = ws.Range("A1").Value
    
    lbltotalwodiscount.Caption = ""
    LblSale_Val.Caption = ""
    'TXTDISCAMOUNT.Text = ws.Range("T2").Value + ws.Range("T2").Value * ws.Range("P2").Value
    'LBLTOTAL.Caption = ws.Range("B" & T).Value
    TXTDATE.text = Format(Date, "DD/MM/YYYY")
    TXTINVDATE.text = Format(Date, "DD/MM/YYYY") 'Format(ws.Range("S2").value, "DD/MM/YYYY")
    'TXTINVDATE.Text = Right(ws.Range("R2").Value, 2) & "/" & Mid(ws.Range("R2").Value, 5, 2) & "/" & Mid(ws.Range("R2").Value, 1, 4)
    'TXTDEALER.Text = ws.Range("W2").Value
    TXTREMARKS.text = ""
    Dim n, M As Integer
    n = 0
    grdsales.FixedRows = 0
    grdsales.rows = 1
    For i = 4 To 10000
        If ws.Range("A" & i).Value = "" Then Exit For
        M = 0
        n = n + 1
        grdsales.rows = grdsales.rows + 1
        grdsales.FixedRows = 1
        grdsales.TextMatrix(n, 0) = n
        'grdsales.TextMatrix(n, 1) = "" 'rstTRXMAST!ITEM_CODE
        grdsales.TextMatrix(n, 2) = Trim(ws.Range("B" & i).Value) 'ITEM NAME
        'If InStr(Trim(grdsales.TextMatrix(n, 2)), "'") <> 0 Then
        '    grdsales.TextMatrix(n, 2) = Trim(Mid(grdsales.TextMatrix(n, 2), 1, Len(grdsales.TextMatrix(n, 2)) - 4))
        'Else
        '    grdsales.TextMatrix(n, 2) = Trim(grdsales.TextMatrix(n, 2))
        'End If
        'grdsales.TextMatrix(n, 2) = Replace(grdsales.TextMatrix(n, 2), "  ", " ")
        'grdsales.TextMatrix(n, 2) = Replace(grdsales.TextMatrix(n, 2), "#", "")
        'grdsales.TextMatrix(n, 2) = Replace(grdsales.TextMatrix(n, 2), "$", "")
        'grdsales.TextMatrix(n, 2) = Replace(grdsales.TextMatrix(n, 2), "'", "")
        
        grdsales.TextMatrix(n, 22) = grdsales.TextMatrix(n, 2) 'ITEM NAME
        
        grdsales.TextMatrix(n, 1) = ""
        'On Error Resume Next
        Set rstTRXMAST = New ADODB.Recordset
        rstTRXMAST.Open "Select * From ITEMMAST WHERE ITEM_NAME = '" & Trim(grdsales.TextMatrix(n, 2)) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                PR_CODE = IIf(IsNull(RSTITEMMAST.Fields(0)), 1, Val(RSTITEMMAST.Fields(0)) + 1)
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing

            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
            RSTITEMMAST.AddNew
            RSTITEMMAST!ITEM_CODE = PR_CODE
            RSTITEMMAST!ITEM_NAME = Trim(grdsales.TextMatrix(n, 2))
            RSTITEMMAST!Category = "GENERAL"
            RSTITEMMAST!UNIT = 1
            RSTITEMMAST!MANUFACTURER = "GENERAL"
            RSTITEMMAST!REMARKS = ""
            RSTITEMMAST!REORDER_QTY = 1
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
            RSTITEMMAST.Update
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing

            grdsales.TextMatrix(n, 1) = PR_CODE
            grdsales.TextMatrix(n, 20) = 0
            grdsales.TextMatrix(n, 5) = 1
            
        Else
            grdsales.TextMatrix(n, 1) = rstTRXMAST!ITEM_CODE
            grdsales.TextMatrix(n, 20) = IIf(IsNull(rstTRXMAST!CUST_DISC), 0, rstTRXMAST!CUST_DISC)
            grdsales.TextMatrix(n, 5) = IIf(IsNull(rstTRXMAST!UNIT), 1, rstTRXMAST!UNIT)
        End If
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
        
        If grdsales.TextMatrix(n, 1) = "" Then
            MsgBox "Error in loading file... Try upload again.", , "EzBiz"
            Exit For
        End If
        'grdsales.TextMatrix(n, 3) = ""
'        grdsales.TextMatrix(N, 5) = 1
        'On Error Resume Next
'        If InStr(UCase(ws.Range("D" & i).value), "'S") = 0 Then
'            grdsales.TextMatrix(N, 5) = Int(Replace(ws.Range("D" & i).value, "'", ""))
'            grdsales.TextMatrix(N, 5) = Int(Replace(ws.Range("D" & i).value, "S", ""))
'            grdsales.TextMatrix(N, 5) = Int(Replace(ws.Range("D" & i).value, "s", ""))
'        Else
'            M = InStr(ws.Range("D" & i).value, "'")
'            grdsales.TextMatrix(N, 5) = Mid(ws.Range("D" & i).value, 1, M - 1)
'        End If
        'If Val(grdsales.TextMatrix(N, 5)) <= 0 Then grdsales.TextMatrix(N, 5) = 1  'PACK
        'On Error GoTo errhandler
        
        grdsales.TextMatrix(n, 5) = Val(ws.Range("D" & i).Value)  'PACK
        If Val(grdsales.TextMatrix(n, 5)) = 0 Then grdsales.TextMatrix(n, 5) = "1"
        grdsales.TextMatrix(n, 3) = Val(ws.Range("E" & i).Value) '+ Val(ws.Range("G" & i).value) 'QTY  + FREE
        grdsales.TextMatrix(n, 4) = 1
        grdsales.TextMatrix(n, 6) = Format(Val(ws.Range("G" & i).Value), ".000") 'MRP
        grdsales.TextMatrix(n, 7) = Format(Val(ws.Range("I" & i).Value), ".000")   'MRP / PACK 'Format(Val(ws.Range("J" & i).value) / Val(grdsales.TextMatrix(N, 5)), ".000")
'        If Val(ws.Range("H" & i).value) = 0 Then
'            grdsales.TextMatrix(n, 9) = Format(Val(ws.Range("G" & i).value) / Val(grdsales.TextMatrix(n, 5)), ".000")
'            grdsales.TextMatrix(n, 10) = Val(ws.Range("J" & i).value) 'IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
'            grdsales.TextMatrix(n, 15) = "V" 'IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
'        ElseIf Val(ws.Range("H" & i).value) = 4.77 Then
'            grdsales.TextMatrix(n, 9) = Format(Val(ws.Range("G" & i).value) / Val(grdsales.TextMatrix(n, 5)), ".000")
'            grdsales.TextMatrix(n, 10) = "5"
'            grdsales.TextMatrix(n, 15) = "M" 'IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
'        Else
'            grdsales.TextMatrix(n, 9) = Format(Val(ws.Range("I" & i).value) / Val(grdsales.TextMatrix(n, 5)), ".000")
'            grdsales.TextMatrix(n, 10) = Val(ws.Range("J" & i).value) 'IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
'            grdsales.TextMatrix(n, 15) = "N" 'IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
'        End If
        grdsales.TextMatrix(n, 9) = Format(Val(ws.Range("H" & i).Value) / Val(grdsales.TextMatrix(n, 5)), ".000") 'RATE / Pack  Format(Val(ws.Range("H" & i).value) / Val(grdsales.TextMatrix(N, 5)), ".000")
        grdsales.TextMatrix(n, 10) = Val(ws.Range("F" & i).Value) 'TAX  'IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
        grdsales.TextMatrix(n, 15) = "V" 'IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
            
        grdsales.TextMatrix(n, 11) = "" 'Trim(ws.Range("C" & i).Value) ' BATCH
        grdsales.TextMatrix(n, 12) = "" ' EXPIRY
'        If Len(Trim(ws.Range("E" & i).value)) = 4 Then
'            MON = Val(Mid(Trim(ws.Range("F" & i).value), 1, 2))
'            Y = Val(Right(Trim(ws.Range("F" & i).value), 2))
'            Y = 2000 + Y
'            M_DATE = "01" & "/" & MON & "/" & Y
'            D = LastDayOfMonth(M_DATE)
'            M_DATE = D & "/" & MON & "/" & Y
'            grdsales.TextMatrix(N, 12) = M_DATE
'        Else
'            If IsDate(Trim(ws.Range("E" & i).value)) Then
'                grdsales.TextMatrix(N, 12) = Format(Trim(ws.Range("E" & i).value), "dd/mm/yyyy")
'            Else
'                grdsales.TextMatrix(N, 12) = ""
'            End If
'        End If
        
        grdsales.TextMatrix(n, 14) = 0 ' FREE Val(ws.Range("G" & i).value)
        grdsales.TextMatrix(n, 17) = Val(ws.Range("M" & i).Value) 'DISC IIf(IsNull(rstTRXMAST!P_DISC), 0, rstTRXMAST!P_DISC)
        grdsales.TextMatrix(n, 18) = Round((100 * Val(grdsales.TextMatrix(n, 6)) / (100 + Val(grdsales.TextMatrix(n, 10))) / Val(grdsales.TextMatrix(n, 5))), 3)
        txtmrpbt.Tag = Round(100 * Val(grdsales.TextMatrix(n, 6)) / (100 + Val(grdsales.TextMatrix(n, 10))), 3)
        Txtdisccust.Tag = (Val(grdsales.TextMatrix(n, 9)) - (Val(grdsales.TextMatrix(n, 9)) * Val(grdsales.TextMatrix(n, 17)) / 100)) * Val(grdsales.TextMatrix(n, 3))
        lbltaxamount.Tag = Val(Txtdisccust.Tag) * Val(grdsales.TextMatrix(n, 10)) / 100
        grdsales.TextMatrix(n, 13) = Round(Val(Txtdisccust.Tag) + Val(lbltaxamount.Tag), 3)
        
'        Txtdisccust.Tag = ((Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 5))) * Val(grdsales.TextMatrix(N, 17)) / 100) * (Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14)))
'        If Val(grdsales.TextMatrix(N, 10)) = 0 Then
'            grdsales.TextMatrix(N, 18) = 0
'            grdsales.TextMatrix(N, 13) = Round((Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))) * (Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 5))) - Val(Txtdisccust.Tag), 3)
'        Else
'            If grdsales.TextMatrix(N, 15) = "M" Then
'                lbltaxamount.Tag = (Val(txtmrpbt.Tag) * Val(grdsales.TextMatrix(N, 3))) * Val(grdsales.TextMatrix(N, 10)) / 100
'                grdsales.TextMatrix(N, 13) = Round(((Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))) * (Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 5)))) + Val(lbltaxamount.Tag) - Val(Txtdisccust.Tag), 3)
'            ElseIf grdsales.TextMatrix(N, 15) = "V" Then
'                'T2 -(T2 * AC2 / 100)
'                Txtdisccust.Tag = (Val(ws.Range("H" & i).value) - (Val(ws.Range("H" & i).value) * Val(grdsales.TextMatrix(N, 17)) / 100)) * Val(Val(ws.Range("M" & i).value))
'                lbltaxamount.Tag = Val(Txtdisccust.Tag) * Val(grdsales.TextMatrix(N, 10)) / 100
'                'lbltaxamount.Tag = (((Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 5)) * Val(grdsales.TextMatrix(N, 10))) - Val(Txtdisccust.Tag)) / 100) * (Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14)))
'                grdsales.TextMatrix(N, 13) = Round(Val(Txtdisccust.Tag) + Val(lbltaxamount.Tag), 3)
'            Else
'                lbltaxamount.Caption = Round((Val(grdsales.TextMatrix(N, 9)) * Val(grdsales.TextMatrix(N, 10)) / 100) * (Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))), 3)
'                grdsales.TextMatrix(N, 13) = Format(Round(((Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))) * Val(grdsales.TextMatrix(N, 9))) + Val(lbltaxamount.Caption), 3), "0.000")
'                 'bltaxamount.Tag = Val(txtmrpbt.Tag) * (grdsales.TextMatrix(n, 14)) * 5 / 100
'                'grdsales.TextMatrix(n, 13) = Round(((Val(grdsales.TextMatrix(n, 3)) - Val(grdsales.TextMatrix(n, 14))) * (Val(grdsales.TextMatrix(n, 9)) * Val(grdsales.TextMatrix(n, 5)))) + Val(lbltaxamount.Tag) - Val(Txtdisccust.Tag), 3)
'            End If
'        End If
        'Format(Round((Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.Text)) / Val(TxtPack.Text), 3), ".000")
        If Val(grdsales.TextMatrix(n, 5)) = 0 Then grdsales.TextMatrix(n, 5) = 1
        If Val(grdsales.TextMatrix(n, 3)) - Val(grdsales.TextMatrix(n, 14)) = 0 Then
            grdsales.TextMatrix(n, 8) = "0.00"
        Else
            'grdsales.TextMatrix(N, 8) = Format(Round((Val(grdsales.TextMatrix(N, 13)) / Val(grdsales.TextMatrix(N, 3)) - Val(grdsales.TextMatrix(N, 14))) / Val(grdsales.TextMatrix(N, 5)), 3), ".000")
            grdsales.TextMatrix(n, 8) = Val(Txtdisccust.Tag) / Val(grdsales.TextMatrix(n, 3))
        End If
        'grdsales.TextMatrix(n, 8) = Format(Val(grdsales.TextMatrix(n, 13)) / (Val(grdsales.TextMatrix(n, 3)) - Val(grdsales.TextMatrix(n, 14))) / Val(grdsales.TextMatrix(n, 5)), ".000")
        'if Val(grdsales.TextMatrix(grdsales.Row, 8))= 0 then grdsales.TextMatrix(grdsales.Row, 8)
        'Format(Round((Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.Text)) / Val(TxtPack.Text), 3), ".000")
        
        'grdsales.TextMatrix(n, 8) = Format(Val(grdsales.TextMatrix(n, 13)) / (Val(grdsales.TextMatrix(n, 3)) - Val(grdsales.TextMatrix(n, 14))) / Val(grdsales.TextMatrix(n, 5)), ".000")
        grdsales.TextMatrix(n, 16) = n 'rstTRXMAST!LINE_NO
        grdsales.TextMatrix(n, 23) = "" 'Trim(ws.Range("B" & i).value)
        grdsales.TextMatrix(n, 24) = "" 'Left(Trim(ws.Range("W" & i).value), 20)
        grdsales.TextMatrix(n, 25) = Left(Trim(ws.Range("K" & i).Value), 8) 'HSN
        grdsales.TextMatrix(n, 26) = Val(ws.Range("I" & i).Value) 'R. RATE
        grdsales.TextMatrix(n, 27) = Val(ws.Range("I" & i).Value) 'W. RATE
        grdsales.TextMatrix(n, 28) = ws.Range("J" & i).Value 'barcode
    Next i
    'or
    var = ws.Cells(1, 1).Value
    wb.Close
    
    xlApp.Quit
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    
'    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
'    LblProfittotal.Caption = Format(Val(LblSale_Val.Caption) - Val(LBLTOTAL.Caption), ".00")
'
    FRMEMASTER.Visible = True
    Frame2.Visible = True
    FRMECONTROLS.Visible = True
    Frmmain.Enabled = False
    
    'TXTSLNO.SetFocus
    Screen.MousePointer = vbNormal
    Exit Function
errHandler:
    Screen.MousePointer = vbNormal
    If err.Number = 9 Then
        MsgBox "NO SUCH INVOICE PRESENT!!", vbOKOnly, "PURCHASE"
        wb.Close
        xlApp.Quit
        Set ws = Nothing
        Set wb = Nothing
        Set xlApp = Nothing
    ElseIf err.Number = 32755 Then
        
    Else
        MsgBox err.Description
    End If
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
End Function

