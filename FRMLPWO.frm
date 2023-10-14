VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLPWO 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE WO"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11250
   ControlBox      =   0   'False
   Icon            =   "FRMLPWO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   11250
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
      Left            =   4440
      TabIndex        =   43
      Top             =   8010
      Width           =   1200
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
      TabIndex        =   41
      Top             =   150
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
      Left            =   4440
      TabIndex        =   38
      Top             =   8520
      Width           =   1200
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   2145
      TabIndex        =   22
      Top             =   2670
      Visible         =   0   'False
      Width           =   5790
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   2985
         Left            =   90
         TabIndex        =   23
         Top             =   105
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   5265
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
   Begin VB.Frame Fram 
      BackColor       =   &H00C0C000&
      Caption         =   "Frame1"
      Height          =   9105
      Left            =   -120
      TabIndex        =   9
      Top             =   -45
      Width           =   11370
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00C0C000&
         Height          =   1575
         Left            =   150
         TabIndex        =   27
         Top             =   -15
         Width           =   11190
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
            TabIndex        =   39
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
            Left            =   7665
            MaxLength       =   100
            TabIndex        =   35
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
            Left            =   2820
            MaxLength       =   10
            TabIndex        =   34
            Top             =   210
            Width           =   1260
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
            Height          =   300
            Left            =   7665
            MaxLength       =   10
            TabIndex        =   28
            Top             =   540
            Width           =   1005
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   7665
            TabIndex        =   37
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
            TabIndex        =   40
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
            Left            =   6315
            TabIndex        =   36
            Top             =   1260
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
            TabIndex        =   33
            Top             =   210
            Width           =   780
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
            Height          =   300
            Index           =   0
            Left            =   165
            TabIndex        =   32
            Top             =   210
            Width           =   870
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   1
            Left            =   6315
            TabIndex        =   31
            Top             =   510
            Width           =   1215
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE DATE"
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
            Left            =   6315
            TabIndex        =   30
            Top             =   945
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "SUPPLIER"
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
            TabIndex        =   29
            Top             =   600
            Width           =   1005
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdsales 
         Height          =   4995
         Left            =   150
         TabIndex        =   85
         Top             =   1590
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8811
         _Version        =   393216
         Rows            =   1
         Cols            =   26
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
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00C0C000&
         Height          =   2580
         Left            =   150
         TabIndex        =   10
         Top             =   6495
         Width           =   11190
         Begin VB.TextBox txtvanrate 
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
            Left            =   1965
            MaxLength       =   7
            TabIndex        =   91
            Top             =   1035
            Width           =   1005
         End
         Begin VB.TextBox txtcrtnpack 
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
            Left            =   4005
            MaxLength       =   7
            TabIndex        =   89
            Top             =   1035
            Width           =   990
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   60
            TabIndex        =   86
            Top             =   1365
            Width           =   3705
            Begin VB.OptionButton OptComAmt 
               BackColor       =   &H00C0C000&
               Caption         =   "Commision Amt"
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
               Left            =   1770
               TabIndex        =   88
               Top             =   195
               Width           =   1755
            End
            Begin VB.OptionButton OptComper 
               BackColor       =   &H00C0C000&
               Caption         =   "Commision %"
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
               TabIndex        =   87
               Top             =   195
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.TextBox TxtComper 
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
            Left            =   5040
            MaxLength       =   7
            TabIndex        =   82
            Top             =   1035
            Width           =   765
         End
         Begin VB.TextBox TxtComAmt 
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
            Left            =   5820
            MaxLength       =   7
            TabIndex        =   81
            Top             =   1035
            Width           =   945
         End
         Begin VB.TextBox txtcrtn 
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
            Left            =   2985
            MaxLength       =   7
            TabIndex        =   79
            Top             =   1035
            Width           =   1005
         End
         Begin VB.TextBox txtWS 
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
            Height          =   315
            Left            =   960
            MaxLength       =   7
            TabIndex        =   77
            Top             =   1035
            Width           =   975
         End
         Begin VB.TextBox TXTRETAIL 
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
            Left            =   75
            MaxLength       =   7
            TabIndex        =   75
            Top             =   1050
            Width           =   870
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
            Left            =   9210
            MaxLength       =   7
            TabIndex        =   69
            Top             =   420
            Width           =   630
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
            Top             =   2265
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
            Top             =   2430
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
            Left            =   8205
            TabIndex        =   63
            Top             =   1665
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
            Left            =   8205
            TabIndex        =   62
            Top             =   2160
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
            Height          =   330
            Left            =   8205
            TabIndex        =   55
            Top             =   1110
            Width           =   1095
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
            Left            =   5715
            TabIndex        =   52
            Top             =   2115
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1005
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
            Left            =   6600
            MaxLength       =   7
            TabIndex        =   53
            Top             =   420
            Width           =   660
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
            Left            =   5715
            TabIndex        =   50
            Top             =   1485
            Visible         =   0   'False
            Width           =   1740
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
            Left            =   5715
            TabIndex        =   51
            Top             =   1740
            Visible         =   0   'False
            Width           =   1005
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
            Left            =   6585
            MaxLength       =   7
            TabIndex        =   46
            Top             =   2205
            Visible         =   0   'False
            Width           =   645
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
            Left            =   4905
            MaxLength       =   7
            TabIndex        =   44
            Top             =   420
            Width           =   600
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
            Left            =   8220
            MaxLength       =   7
            TabIndex        =   26
            Top             =   420
            Width           =   975
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
            Left            =   7290
            MaxLength       =   7
            TabIndex        =   25
            Top             =   420
            Width           =   900
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
            TabIndex        =   5
            Top             =   2070
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
            Height          =   300
            Left            =   75
            TabIndex        =   0
            Top             =   420
            Width           =   540
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
            Left            =   630
            TabIndex        =   1
            Top             =   420
            Width           =   4260
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
            Left            =   5535
            MaxLength       =   8
            TabIndex        =   2
            Top             =   420
            Width           =   1050
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
            Left            =   2340
            TabIndex        =   7
            Top             =   2070
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
            Left            =   1230
            TabIndex        =   6
            Top             =   2070
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
            Height          =   300
            Left            =   12300
            TabIndex        =   12
            Top             =   2295
            Visible         =   0   'False
            Width           =   690
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
            TabIndex        =   11
            Top             =   2370
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
            Left            =   3405
            TabIndex        =   8
            Top             =   2070
            Width           =   975
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   315
            Left            =   13695
            TabIndex        =   3
            Top             =   2385
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
            Left            =   13740
            TabIndex        =   42
            Top             =   2385
            Visible         =   0   'False
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
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Van Rate"
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
            Left            =   1965
            TabIndex        =   92
            Top             =   780
            Width           =   1005
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
            Left            =   4005
            TabIndex        =   90
            Top             =   780
            Width           =   990
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
            Left            =   5040
            TabIndex        =   84
            Top             =   780
            Width           =   750
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
            Left            =   5820
            TabIndex        =   83
            Top             =   780
            Width           =   945
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
            Left            =   2985
            TabIndex        =   80
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "W.Sale"
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
            Left            =   960
            TabIndex        =   78
            Top             =   780
            Width           =   975
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
            Top             =   2190
            Visible         =   0   'False
            Width           =   780
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
            Left            =   9210
            TabIndex        =   70
            Top             =   165
            Width           =   630
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
            Left            =   75
            TabIndex        =   67
            Top             =   780
            Width           =   870
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
            Left            =   8235
            TabIndex        =   65
            Top             =   1425
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
            Left            =   8220
            TabIndex        =   64
            Top             =   1935
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
            Height          =   210
            Index           =   21
            Left            =   9585
            TabIndex        =   60
            Top             =   1710
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
            Left            =   9360
            TabIndex        =   59
            Top             =   1950
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
            Left            =   8280
            TabIndex        =   58
            Top             =   885
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
            Height          =   570
            Left            =   9405
            TabIndex        =   57
            Top             =   1140
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PURCHASE AMT"
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
            Left            =   9435
            TabIndex        =   56
            Top             =   900
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
            Height          =   255
            Index           =   17
            Left            =   6600
            TabIndex        =   54
            Top             =   165
            Width           =   660
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
            Left            =   7245
            TabIndex        =   49
            Top             =   2205
            Visible         =   0   'False
            Width           =   915
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
            Left            =   7245
            TabIndex        =   48
            Top             =   1950
            Visible         =   0   'False
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
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   12
            Left            =   6585
            TabIndex        =   47
            Top             =   1950
            Visible         =   0   'False
            Width           =   645
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
            Height          =   255
            Index           =   4
            Left            =   4905
            TabIndex        =   45
            Top             =   165
            Width           =   600
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
            Height          =   255
            Index           =   2
            Left            =   8220
            TabIndex        =   24
            Top             =   165
            Width           =   960
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
            Left            =   75
            TabIndex        =   21
            Top             =   165
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
            Height          =   255
            Index           =   9
            Left            =   630
            TabIndex        =   20
            Top             =   165
            Width           =   4245
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
            Left            =   5535
            TabIndex        =   19
            Top             =   165
            Width           =   1050
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
            Height          =   255
            Index           =   11
            Left            =   7290
            TabIndex        =   18
            Top             =   165
            Width           =   900
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
            Left            =   6780
            TabIndex        =   17
            Top             =   780
            Width           =   1335
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
            TabIndex        =   16
            Top             =   2310
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
            Height          =   255
            Index           =   16
            Left            =   13695
            TabIndex        =   15
            Top             =   2160
            Visible         =   0   'False
            Width           =   1215
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
            Left            =   6780
            TabIndex        =   4
            Top             =   1035
            Width           =   1335
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
            TabIndex        =   14
            Top             =   2100
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
            Left            =   10275
            TabIndex        =   13
            Top             =   2385
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   135
         TabIndex        =   72
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   705
         TabIndex        =   71
         Top             =   45
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmLPWO"
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

Private Sub CMDADD_Click()
    Dim i As Integer
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim M_DATA As Long
    
    M_DATA = 0
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text) + Val(TXTFREE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = 1 'Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Val(TxtPack.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Round(Val(TXTRATE.Text) / Val(TxtPack.Text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Format(Round((Val(LBLSUBTOTAL.Caption) / (Val(TxtPack.Text) * TXTQTY.Text)), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Format(Round(Val(TXTPTR.Text) / Val(TxtPack.Text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Format((Val(txtprofit.Text)), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = IIf(Val(TxttaxMRP.Text) = 0, "", Format(Val(TxttaxMRP.Text), ".00")) 'TAX
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = ""
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = "" 'IIf(Trim(TXTEXPDATE.Text) = "/  /", "", TXTEXPDATE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Val(TXTFREE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = Val(txtPD.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = Format(Val(txtretail.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = Format(Val(txtWS.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = Format(Val(txtvanrate.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = Format(Val(txtcrtn.Text), ".000")
    If OptComAmt.Value = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 21) = ""
        grdsales.TextMatrix(Val(TXTSLNO.Text), 22) = Format(Val(TxtComAmt.Text), ".00")
        grdsales.TextMatrix(Val(TXTSLNO.Text), 23) = "A"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.Text), 21) = Format(Val(TxtComper.Text), ".00")
        grdsales.TextMatrix(Val(TXTSLNO.Text), 22) = ""
        grdsales.TextMatrix(Val(TXTSLNO.Text), 23) = "P"
    End If
    grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = Format(Val(txtcrtnpack.Text), ".000")
    If Val(TxttaxMRP.Text) = 0 Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = "N"
    Else
        If OPTTaxMRP.Value = True Then
            grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = "M"
        ElseIf OPTVAT.Value = True Then
            grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = "V"
        End If
    End If
    
    If M_EDIT = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16))
    Else
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Val(TXTSLNO.Text)
    End If
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILEWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "", db2, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!TRX_TYPE = "PI"
        RSTRTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTRTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16))
        RSTRTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1))
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
        RSTRTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                rststock!CATEGORY = IIf(IsNull(rststock!CATEGORY), "OTHERS", rststock!CATEGORY)
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
                
                !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                !RCPT_VAL = !RCPT_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
                
                !P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
                !P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19))
                !P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
                !P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25))
                !SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
                !ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
                !CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23)) = "A" Then
                    !COM_FLAG = "A"
                    !COM_PER = 0
                    !COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22))
                Else
                    !COM_FLAG = "P"
                    !COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21))
                    !COM_AMT = 0
                End If
                
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
        rststock.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                rststock!CATEGORY = IIf(IsNull(rststock!CATEGORY), "OTHERS", rststock!CATEGORY)
                !CLOSE_QTY = !CLOSE_QTY - RSTRTRXFILE!QTY
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
                
                !RCPT_QTY = !RCPT_QTY - RSTRTRXFILE!QTY
                !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                !RCPT_VAL = !RCPT_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
                
                !P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
                !P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19))
                !P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
                !P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25))
                
                !SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
                !ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
                !CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23)) = "A" Then
                    !COM_FLAG = "A"
                    !COM_PER = 0
                    !COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22))
                Else
                    !COM_FLAG = "P"
                    !COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21))
                    !COM_AMT = 0
                End If
    
                RSTRTRXFILE!MFGR = !MANUFACTURER
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
    End If
    
    RSTRTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
    RSTRTRXFILE!VCH_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2))
    RSTRTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
    RSTRTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
    RSTRTRXFILE!P_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
    RSTRTRXFILE!MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
    RSTRTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 9))
    RSTRTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
    RSTRTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
    RSTRTRXFILE!P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19))
    RSTRTRXFILE!P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
    RSTRTRXFILE!CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
    RSTRTRXFILE!P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25))
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
    RSTRTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
    'RSTRTRXFILE!VCH_DESC = "Received From " & DataList2.Text
    RSTRTRXFILE!REF_NO = ""
    'RSTRTRXFILE!ISSUE_QTY = 0
    RSTRTRXFILE!CST = 0
        
    RSTRTRXFILE!SCHEME = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))
    RSTRTRXFILE!EXP_DATE = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = "", Null, Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "dd/mm/yyyy"))
    RSTRTRXFILE!FREE_QTY = 0
    RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!C_USER_ID = "SM"
    RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))
        
    'RSTRTRXFILE!M_USER_ID = DataList2.BoundText
    ''''RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))  'MODE OF TAX
    'RSTRTRXFILE!PINV = Trim(TXTINVOICE.Text)
    RSTRTRXFILE.Update
    RSTRTRXFILE.Close
    
    M_DATA = 0
    Set RSTRTRXFILE = Nothing
    
           
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    For i = 1 To grdsales.Rows - 1
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
    Next i
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TXTPTR.Text = ""
    TXTQTY.Text = ""
    TxtPack.Text = ""
    TXTFREE.Text = ""
    TxttaxMRP.Text = ""
    txtPD.Text = ""
    txtprofit.Text = ""
    txtretail.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    txtcrtn.Text = ""
    txtcrtnpack.Text = ""
    TXTRATE.Text = ""
    TxtComAmt.Text = ""
    TxtComper.Text = ""
    txtmrpbt.Text = ""
    
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    cmdadd.Enabled = False
    cmddelete.Enabled = False
    CMDEXIT.Enabled = False
    optnet.Value = True
    OptComper.Value = True
    M_EDIT = False
    M_ADD = True
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    txtBillNo.Enabled = False
    
    If grdsales.Rows >= 18 Then grdsales.TopRow = grdsales.Rows - 1
        
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            txtcrtnpack.Enabled = True
            txtcrtnpack.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim i As Integer
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
   
    db2.Execute "delete *  From RTRXFILEWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & ""
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
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
    rststock.Close
    Set rststock = Nothing
    
    i = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(Val(LINE_NO)) From RTRXFILEWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db2, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILEWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db2, adOpenStatic, adLockOptimistic, adCmdText
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
    RSTRTRXFILE.Open "SELECT * From RTRXFILEWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db2, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTRTRXFILE.EOF
        RSTRTRXFILE!LINE_NO = i
        i = i + 1
        RSTRTRXFILE.Update
        RSTRTRXFILE.MoveNext
    Loop
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    grdsales.Rows = 1
    i = 0
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    grdsales.Rows = 1
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "Select * From RTRXFILEWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY [LINE_NO]", db2, adOpenStatic, adLockReadOnly
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
        grdsales.TextMatrix(i, 11) = ""
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
        grdsales.TextMatrix(i, 24) = IIf(IsNull(RSTRTRXFILE!CRTN_PACK), 0, RSTRTRXFILE!CRTN_PACK)
        grdsales.TextMatrix(i, 25) = IIf(IsNull(RSTRTRXFILE!P_VAN), 0, RSTRTRXFILE!P_VAN)
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
        'TXTDEALER.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
        
        'TXTINVDATE.Text = Format(RSTRTRXFILE!VCH_DATE, "DD/MM/YYYY")
        'TXTREMARKS.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
        'TXTINVOICE.Text = IIf(IsNull(RSTRTRXFILE!PINV), "", RSTRTRXFILE!PINV)
        RSTRTRXFILE.MoveNext
    Loop
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTQTY.Text = ""
    TxtPack.Text = ""
    TXTFREE.Text = ""
    TxttaxMRP.Text = ""
    txtPD.Text = ""
    txtprofit.Text = ""
    txtretail.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    txtcrtn.Text = ""
    txtcrtnpack.Text = ""
    TXTRATE.Text = ""
    TxtComAmt.Text = ""
    TxtComper.Text = ""
    txtmrpbt.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    cmddelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_ADD = True
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub cmditemcreate_Click()
    frmitemmasterwo.Show
    MDIMAIN.Enabled = False
End Sub

Private Sub cmditemcreate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            'If TXTUNIT.Enabled = True Then TXTUNIT.SetFocus
            If TxtPack.Enabled = True Then TxtPack.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If txtPD.Enabled = True Then txtPD.SetFocus
            If txtprofit.Enabled = True Then txtprofit.SetFocus
            If TxttaxMRP.Enabled = True Then TxttaxMRP.SetFocus
            If txtretail.Enabled = True Then txtretail.SetFocus
            If txtWS.Enabled = True Then txtWS.SetFocus
            If txtcrtn.Enabled = True Then txtcrtn.SetFocus
            If txtcrtnpack.Enabled = True Then txtcrtnpack.SetFocus
            
            If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TXTPTR.Enabled = True Then TXTPTR.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            'If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            'If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    M_EDIT = True
    CMDMODIFY.Enabled = False
    cmddelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
    
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TxtPack.Text = ""
            TXTFREE.Text = ""
            TxttaxMRP.Text = ""
            txtPD.Text = ""
            txtprofit.Text = ""
            txtretail.Text = ""
            txtWS.Text = ""
            txtvanrate.Text = ""
            txtcrtn.Text = ""
            txtcrtnpack.Text = ""
            TXTRATE.Text = ""
            TxtComAmt.Text = ""
            TxtComper.Text = ""
            txtmrpbt.Text = ""
            TXTITEMCODE.Text = ""
            LBLSUBTOTAL.Caption = ""
            lbltaxamount.Caption = ""
            TXTEXPDATE.Text = "  /  /    "
            TXTEXPIRY.Text = "  /  "
            
        
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            TXTEXPIRY.Visible = False
            
            CMDMODIFY.Enabled = False
            cmddelete.Enabled = False
            M_EDIT = False
    End Select
End Sub

Private Sub cmdRefresh_Click()
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Supplier From List", vbOKOnly, "Sale Bil..."
        DataList2.SetFocus
        Exit Sub
    End If
    If TXTINVOICE.Text = "" Then
        MsgBox "Enter Supplier Invoice No.", vbOKOnly, "PURCHASE"
        Exit Sub
    End If
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Enter Supplier Invoice Date", vbOKOnly, "PURCHASE"
        Exit Sub
    End If
    Me.Enabled = False
    MDIMAIN.cmdpurchase.Enabled = False
    frmCREDITWO.Show
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrHand
    txtBillNo.SetFocus
    Exit Sub
ErrHand:
    If Err.Number = 5 Then Exit Sub
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(Val(VCH_NO)) From RTRXFILEWO WHERE TRX_TYPE = 'PI'", db2, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    ACT_FLAG = True
    grdsales.ColWidth(0) = 500
    ''grdsales.ColWidth(1) = 0
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
    grdsales.ColWidth(16) = 0
    grdsales.ColWidth(17) = 800
    grdsales.ColWidth(18) = 800
    grdsales.ColWidth(19) = 800
    grdsales.ColWidth(20) = 800
    grdsales.ColWidth(21) = 800
    grdsales.ColWidth(22) = 1000
    grdsales.ColWidth(23) = 500
    
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
    grdsales.ColAlignment(23) = 4
    
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
    grdsales.TextArray(20) = "Cartn Price"
    grdsales.TextArray(21) = "Comm %"
    grdsales.TextArray(22) = "Comm Amt"
    grdsales.TextArray(23) = "Comm Flag"
    grdsales.TextArray(24) = "Cnt Pck"
    grdsales.TextArray(25) = "Van Rate"
    
    PHYFLAG = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTDATE.Text = Date
    TXTEXPDATE.Enabled = False
    
    cmddelete.Enabled = False
    CMDMODIFY.Enabled = False
    TXTUNIT.Enabled = False
    TXTSLNO.Text = 1
    TXTSLNO.Enabled = True
    FRMECONTROLS.Enabled = False
    FRMEMASTER.Enabled = False
    CLOSEALL = 1
    lblcredit.Caption = "1"
    TXTDEALER.Text = ""
    M_ADD = False
    'Me.Width = 15135
    'Me.Height = 9660
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If MDIMAIN.PCTMENU.Visible = True Then
            MDIMAIN.PCTMENU.Enabled = True
            MDIMAIN.PCTMENU.SetFocus
        Else
            MDIMAIN.pctmenu2.Enabled = True
            MDIMAIN.pctmenu2.SetFocus
        End If
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn

            TXTPRODUCT.Text = grdtmp.Columns(1)
            TXTITEMCODE.Text = grdtmp.Columns(0)
            
            For i = 1 To grdsales.Rows - 1
                If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                End If
            Next i
            
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "Select * From RTRXFILEWO  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ORDER BY CREATE_DATE", db2, adOpenStatic, adLockReadOnly
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                RSTRXFILE.MoveLast
                TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                TxtPack.Text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                TXTEXPIRY.Text = IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  ", Format(RSTRXFILE!EXP_DATE, "MM/YY"))
                TXTRATE.Text = IIf(IsNull(RSTRXFILE!MRP), "", Format(Round(Val(RSTRXFILE!MRP) * Val(TxtPack.Text), 2), ".000"))
                If IsNull(RSTRXFILE!MRP_BT) Then
                    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105
                Else
                    txtmrpbt.Text = Format(Val(RSTRXFILE!MRP_BT), ".000")
                End If
                TXTPTR.Text = IIf(IsNull(RSTRXFILE!PTR), "", Format(Round(Val(RSTRXFILE!PTR), 2), ".000"))
                txtretail.Text = IIf(IsNull(RSTRXFILE!P_RETAIL), "", Format(Round(Val(RSTRXFILE!P_RETAIL) * Val(TxtPack.Text), 2), ".000"))
                txtWS.Text = IIf(IsNull(RSTRXFILE!P_WS), "", Format(Round(Val(RSTRXFILE!P_WS), 2), ".000"))
                txtvanrate.Text = IIf(IsNull(RSTRXFILE!P_VAN), "", Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000"))
                txtcrtn.Text = IIf(IsNull(RSTRXFILE!P_CRTN), "", Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000"))
                txtcrtnpack.Text = IIf(IsNull(RSTRXFILE!CRTN_PACK), "", Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000"))
                txtprofit.Text = IIf(IsNull(RSTRXFILE!SALES_PRICE), "", Format(Round(Val(RSTRXFILE!SALES_PRICE), 2), ".000"))
                If IsNull(RSTRXFILE!SALES_TAX) Then
                    TxttaxMRP.Text = ""
                Else
                    TxttaxMRP.Text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                End If
                ''TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                If RSTRXFILE!CHECK_FLAG = "M" Then
                    OPTTaxMRP.Value = True
                ElseIf RSTRXFILE!CHECK_FLAG = "V" Then
                    OPTVAT.Value = True
                Else
                    optnet.Value = True
                End If
            End If
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing
                
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Fram.Enabled = True
            TXTPRODUCT.Enabled = False
            TxtPack.Enabled = True
            TxtPack.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTFREE.Text = ""
            Fram.Enabled = True
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub OPTNET_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxttaxMRP.Text) <> 0 Then
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
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "PURCHASE"
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
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "PURCHASE"
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

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTDIST As ADODB.Recordset
    Dim RSTTRNSMAST As ADODB.Recordset
    Dim i As Integer

    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            grdsales.Rows = 1
            i = 0
            LBLTOTAL.Caption = ""
            lbltotalwodiscount = ""
            grdsales.Rows = 1
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From RTRXFILEWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY [LINE_NO]", db2, adOpenStatic, adLockReadOnly
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
                grdsales.TextMatrix(i, 11) = ""
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
                lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
                TXTDEALER.Text = Mid(rstTRXMAST!VCH_DESC, 15)
                TXTINVDATE.Text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                txtremarks.Text = Mid(rstTRXMAST!VCH_DESC, 15)
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRANSMASTWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db2, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                TXTDISCAMOUNT.Text = IIf(IsNull(rstTRXMAST!DISCOUNT), "", Format(rstTRXMAST!DISCOUNT, ".00"))
                txtaddlamt.Text = IIf(IsNull(rstTRXMAST!ADD_AMOUNT), "", Format(rstTRXMAST!ADD_AMOUNT, ".00"))
                txtcramt.Text = IIf(IsNull(rstTRXMAST!DISC_PERS), "", Format(rstTRXMAST!DISC_PERS, ".00"))
                If rstTRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                TXTINVOICE.Text = IIf(IsNull(rstTRXMAST!PINV), "", rstTRXMAST!PINV)
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            ''''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
            LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
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
                TXTDEALER.SetFocus
            End If
            
            Set RSTTRNSMAST = New ADODB.Recordset
            RSTTRNSMAST.Open "Select CHECK_FLAG From TRANSMASTWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db2, adOpenStatic, adLockReadOnly
            If Not (RSTTRNSMAST.EOF Or RSTTRNSMAST.BOF) Then
                If RSTTRNSMAST!CHECK_FLAG = "Y" Then FRMEMASTER.Enabled = False
            End If
            RSTTRNSMAST.Close
            Set RSTTRNSMAST = Nothing
    End Select
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    Exit Sub
ErrHand:
    MsgBox Err.Description
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

Private Sub txtBillNo_LostFocus()
    If Val(txtBillNo.Text) = 0 Or Val(txtBillNo.Text) > Val(TXTLASTBILL.Text) Then txtBillNo.Text = TXTLASTBILL.Text
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
                MsgBox "Item Expired....", vbOKOnly, "PURCHASE.."
                TXTEXPDATE.SelStart = 0
                TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 60 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days", vbOKOnly, "PURCHASE.."
                TXTEXPDATE.SelStart = 0
                TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 180 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "PURCHASE..") = vbNo Then
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
            TXTEXPDATE.Enabled = False
            TXTEXPIRY.Visible = False
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

Private Sub TXTEXPDATE_LostFocus()
    TXTEXPDATE.Text = Format(TXTEXPDATE.Text, "DD/MM/YYYY")
    If TXTEXPDATE.Text <> "  /  /    " Then TXTEXPIRY.Text = Format(TXTEXPDATE.Text, "MM/YY")
End Sub

Private Sub TxtFree_GotFocus()
    TXTFREE.SelStart = 0
    TXTFREE.SelLength = Len(TXTFREE.Text)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTFREE.Enabled = False
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
        Case vbKeyEscape
            TXTQTY.Enabled = True
            TXTFREE.Enabled = False
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub TxtFree_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtFree_LostFocus()
    If Val(TXTFREE.Text) = 0 Then TXTFREE.Text = 0
    TXTFREE.Text = Format(TXTFREE.Text, "0.00")
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
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                txtremarks.SetFocus
            End If
        Case vbKeyEscape
            TXTINVOICE.SetFocus
    End Select
End Sub

Private Sub TXTINVDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTINVOICE_GotFocus()
    TXTINVOICE.SelStart = 0
    TXTINVOICE.SelLength = Len(TXTINVOICE.Text)
End Sub

Private Sub TXTINVOICE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If TXTINVOICE.Text = "" Then Exit Sub
            TXTINVDATE.SetFocus
        Case vbKeyEscape
            DataList2.SetFocus
    End Select
End Sub

Private Sub TXTINVOICE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("-")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtPack_GotFocus()
    TxtPack.SelStart = 0
    TxtPack.SelLength = Len(TxtPack.Text)
End Sub

Private Sub TxtPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtPack.Text) = 0 Then Exit Sub
            TxtPack.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            TxtPack.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TxtPack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
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
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Integer
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
        
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            cmddelete.Enabled = False
                
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select * From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select * From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From RTRXFILEWO  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ORDER BY CREATE_DATE", db2, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    RSTRXFILE.MoveLast
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    TxtPack.Text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    TXTEXPDATE.Text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    TXTEXPIRY.Text = IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  ", Format(RSTRXFILE!EXP_DATE, "MM/YY"))
                    TXTRATE.Text = IIf(IsNull(RSTRXFILE!MRP), "", Format(Round(Val(RSTRXFILE!MRP) * Val(TxtPack.Text), 2), ".000"))
                    If (IsNull(RSTRXFILE!MRP_BT)) Then
                        txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105
                    Else
                        txtmrpbt.Text = Val(TXTRATE.Text)
                    End If
                    TXTPTR.Text = IIf(IsNull(RSTRXFILE!PTR), "", Format(Round(Val(RSTRXFILE!PTR) * Val(TxtPack.Text), 2), ".000"))
                    txtretail.Text = IIf(IsNull(RSTRXFILE!P_RETAIL), "", Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000"))
                    txtWS.Text = IIf(IsNull(RSTRXFILE!P_WS), "", Format(Round(Val(RSTRXFILE!P_WS), 2), ".000"))
                    txtvanrate.Text = IIf(IsNull(RSTRXFILE!P_VAN), "", Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000"))
                    txtcrtn.Text = IIf(IsNull(RSTRXFILE!P_CRTN), "", Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000"))
                    txtcrtnpack.Text = IIf(IsNull(RSTRXFILE!CRTN_PACK), "", Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000"))
                    txtprofit.Text = IIf(IsNull(RSTRXFILE!SALES_PRICE), "", Format(Round(Val(RSTRXFILE!SALES_PRICE), 2), ".000"))
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.Text = ""
                    Else
                        TxttaxMRP.Text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                    End If
                    'TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                    If RSTRXFILE!CHECK_FLAG = "M" Then
                        OPTTaxMRP.Value = True
                    ElseIf RSTRXFILE!CHECK_FLAG = "V" Then
                        OPTVAT.Value = True
                    Else
                        optnet.Value = True
                    End If
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                
                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    TxtPack.Enabled = True
                    TxtPack.SetFocus
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
            
            TXTSLNO.SetFocus
            cmddelete.Enabled = False
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTPTR_GotFocus()
    TXTPTR.SelStart = 0
    TXTPTR.SelLength = Len(TXTPTR.Text)
End Sub

Private Sub TXTPTR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            txtPD.Enabled = True
            TXTPTR.Enabled = False
            txtPD.SetFocus
        Case vbKeyEscape
            TXTPTR.Enabled = False
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
    End Select
End Sub

Private Sub TXTPTR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPTR_LostFocus()
    'tXTptrdummy.Text = Format(Val(TXTPTR.Text) / Val(TXTUNIT.Text), ".000")
    TXTPTR.Text = Format(TXTPTR.Text, ".000")
    'TXTRETAIL.Text = Round(Val(txtmrpbt.Text) * 0.8, 2)
'    txtretail.Text = Format(Round(Val(TXTRATE.Text) - (Val(txtmrpbt.Text) * 20 / 100), 3), ".000")
'    txtprofit.Text = Format(Round(Val(txtretail.Text) - Val(txtretail.Text) * 10 / 100, 3), ".000")
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            TXTQTY.Enabled = False
            TXTFREE.Enabled = True
            TXTFREE.SetFocus
        Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTQTY.Text = ""
            TXTFREE.Text = ""
            TxttaxMRP.Text = ""
            txtprofit.Text = ""
            txtretail.Text = ""
            txtWS.Text = ""
            txtvanrate.Text = ""
            txtcrtn.Text = ""
            txtcrtnpack.Text = ""
            txtPD.Text = ""
            TXTRATE.Text = ""
            txtmrpbt.Text = ""
            TXTPTR.Text = ""
            LBLSUBTOTAL.Caption = ""
            lbltaxamount.Caption = ""
            TxtPack.Enabled = True
            TxtPack.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
End Sub

Private Sub TXTRATE_GotFocus()
    TXTRATE.SelStart = 0
    TXTRATE.SelLength = Len(TXTRATE.Text)
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
            TXTFREE.Enabled = True
            TXTFREE.SetFocus
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
    TXTRATE.Text = Format(TXTRATE.Text, ".000")
    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105 '(100 + Val(TxttaxMRP.Text))
End Sub

Private Sub txtremarks_GotFocus()
    txtremarks.SelStart = 0
    txtremarks.SelLength = Len(txtremarks.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtBillNo.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then Exit Sub
            'If TXTINVOICE.Text = "" Then Exit Sub
            If Not IsDate(TXTINVDATE.Text) Then Exit Sub
            
            FRMECONTROLS.Enabled = True
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTINVDATE.SetFocus
    End Select
End Sub

Private Sub TXTREMARKS_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
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
                TXTQTY.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))
                TXTUNIT.Text = 1 'grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                TxtPack.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTRATE.Text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)), 2), "0.000")
                TXTPTR.Text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 9)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)), 2), "0.000")
                txtprofit.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7)), "0.00")
                txtretail.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)), "0.00")
                txtWS.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)), "0.00")
                txtvanrate.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)), "0.00")
                txtcrtn.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)), "0.00")
                txtcrtnpack.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)), "0.00")
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23)) = "A" Then
                    OptComAmt.Value = True
                    TxtComper.Text = ""
                    TxtComAmt.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22)), "0.00")
                Else
                    OptComper.Value = True
                    TxtComper.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21)), "0.00")
                    TxtComAmt.Text = ""
                End If
                
                'TXTPTR.Text = Format((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))) * Val(Txtpack.Text), "0.000")

                TXTEXPDATE.Text = "  /  /    " 'IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = "", "  /  /    ", grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                TXTEXPIRY.Text = "  /  " 'IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = "", "  /  ", Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "mm/yy"))
                LBLSUBTOTAL.Caption = Format(Val(TXTQTY.Text) * (Val(TXTPTR.Text) + Val(lbltaxamount.Caption)), ".000")
                TXTFREE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TxttaxMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105 '(100 + Val(TxttaxMRP.Text))
                txtPD.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) = "V" Then
                    OPTVAT.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) = "M" Then
                    OPTTaxMRP.Value = True
                Else
                    optnet.Value = True
                End If
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                TXTRATE.Enabled = False
                TXTEXPDATE.Enabled = False
                
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                cmddelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If cmddelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TXTQTY.Text = ""
                TxtPack.Text = ""
                TXTFREE.Text = ""
                TxttaxMRP.Text = ""
                txtPD.Text = ""
                txtprofit.Text = ""
                txtretail.Text = ""
                txtWS.Text = ""
                txtvanrate.Text = ""
                txtcrtn.Text = ""
                txtcrtnpack.Text = ""
                OptComper.Value = True
                TXTRATE.Text = ""
                TxtComAmt.Text = ""
                TxtComper.Text = ""
                txtmrpbt.Text = ""
                LBLSUBTOTAL.Caption = ""
                lbltaxamount.Caption = ""
                TXTEXPDATE.Text = "  /  /    "
                TXTEXPIRY.Text = "  /  "
                cmdadd.Enabled = False
                cmddelete.Enabled = False
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
                MsgBox "Item Expired....", vbOKOnly, "PURCHASE.."
                TXTEXPDATE.Text = "  /  /    "
                TXTEXPIRY.SelStart = 0
                TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
                TXTEXPIRY.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 60 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days", vbOKOnly, "PURCHASE.."
                TXTEXPDATE.Text = "  /  /    "
                TXTEXPIRY.SelStart = 0
                TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
                TXTEXPIRY.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.Text) < 180 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.Text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "PURCHASE..") = vbNo Then
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
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
                        
    End Select
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
            If Val(TxttaxMRP.Text) <> 0 And optnet.Value = True Then
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxttaxMRP_LostFocus()
    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / (100 + Val(TxttaxMRP.Text))
    If Val(TxttaxMRP.Text) = 0 Then
        TxttaxMRP.Text = 0
        lbltaxamount.Caption = 0
        lbltaxamount.Caption = ""
        LBLSUBTOTAL.Caption = Format(Val(TXTQTY.Text) * Val(TXTPTR.Text), ".000")

    Else
        If OPTTaxMRP.Value = True Then
            'lbltaxamount.Caption = Val(txtmrpbt.Text) * ((Val(TXTQTY.Text) + Val(TxtFree.Text))) * Val(TxttaxMRP.Text) / 100
            lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
            LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
            'lbltaxamount.Caption = (((Val(TXTRATE.Text) / Val(Txtpack.Text)))*VAL(TXTQTY.Text)+VAL(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100) * Val(Txtpack.Text)
        ElseIf OPTVAT.Value = True Then
            'lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text) + Val(TxtFree.Text))
            lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text))
            LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
        Else
            lbltaxamount.Caption = ""
            LBLSUBTOTAL.Caption = Format(Val(TXTQTY.Text) * Val(TXTPTR.Text), ".000")
        End If
    End If

    TxttaxMRP.Text = Format(TxttaxMRP.Text, "0.00")
    lbltaxamount.Caption = Format(lbltaxamount.Caption, "0.00")
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
            TxtPack.Enabled = True
            TxtPack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTQTY.Text = ""
            TXTFREE.Text = ""
            TxttaxMRP.Text = ""
            txtprofit.Text = ""
            txtretail.Text = ""
            txtWS.Text = ""
            txtvanrate.Text = ""
            txtcrtn.Text = ""
            txtcrtnpack.Text = ""
            txtPD.Text = ""
            TXTRATE.Text = ""
            txtmrpbt.Text = ""
            TXTPTR.Text = ""
            TXTEXPDATE.Text = "  /  /    "
            TXTEXPIRY.Text = "  /  "
            LBLSUBTOTAL.Caption = ""
            lbltaxamount.Caption = ""
            TXTPRODUCT.Enabled = True
            TXTUNIT.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTUNIT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
    If (TXTDISCAMOUNT.Text = "") Then
        DISC = 0
    Else
        DISC = TXTDISCAMOUNT.Text
    End If
    If grdsales.Rows = 1 Then
        TXTDISCAMOUNT.Text = "0"
    ElseIf Val(TXTDISCAMOUNT.Text) > Val(lbltotalwodiscount.Caption) Then
        MsgBox "Discount Amount More than Bill Amount", , "SALES..."
        TXTDISCAMOUNT.SelStart = 0
        TXTDISCAMOUNT.SelLength = Len(TXTDISCAMOUNT.Text)
        TXTDISCAMOUNT.SetFocus
        Exit Sub
    End If
    TXTDISCAMOUNT.Text = Format(TXTDISCAMOUNT.Text, ".00")
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    ''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    TXTDISCAMOUNT.SetFocus
End Sub

Private Sub TXTDISCAMOUNT_GotFocus()
    TXTDISCAMOUNT.SelStart = 0
    TXTDISCAMOUNT.SelLength = Len(TXTDISCAMOUNT.Text)
End Sub

Private Sub TXTDISCAMOUNT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Public Sub appendpurchase()
    
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    Dim rstMaxRec As New ADODB.Recordset
    
    Dim M_DATA As Double
    Dim i As Integer
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    
    db2.Execute "delete * From TRANSMASTWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    
    If grdsales.Rows = 1 Then
        db2.Execute "delete * FROM CRDTPYMT WHERE INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'CR'"
        db2.Execute "delete * FROM CASHATRXFILE WHERE INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'PY'"
        GoTo SKIP
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRANSMASTWO WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db2, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "PI"
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = Trim(DataList2.Text)
        RSTTRXFILE!VCH_AMOUNT = Val(lbltotalwodiscount.Caption)
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
        RSTTRXFILE!REMARKS = Trim(DataList2.Text)
        RSTTRXFILE!DISC_PERS = Val(txtcramt.Text)
        RSTTRXFILE!LETTER_NO = 0
        RSTTRXFILE!LETTER_DATE = Date
        RSTTRXFILE!INV_MSGS = ""
        RSTTRXFILE!CREATE_DATE = Format(TXTDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!MODIFY_DATE = Format(TXTDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!PINV = Trim(TXTINVOICE.Text)
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
            
    i = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(Val(CR_NO)) From CRDTPYMT", db2, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CRDTPYMT WHERE INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'CR'", db2, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!TRX_TYPE = "CR"
        RSTITEMMAST!CR_NO = i
        RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
        RSTITEMMAST!RCPT_AMOUNT = 0
    End If
    RSTITEMMAST!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTITEMMAST!INV_AMT = Val(LBLTOTAL.Caption)
    RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption) - RSTITEMMAST!RCPT_AMOUNT
    If lblcredit.Caption = "0" Then RSTITEMMAST!CHECK_FLAG = "Y" Else RSTITEMMAST!CHECK_FLAG = "N"
    RSTITEMMAST!PINV = Trim(TXTINVOICE.Text)
    RSTITEMMAST!ACT_CODE = DataList2.BoundText
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from [RTRXFILEWO] WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db2, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(Trim(TXTINVDATE.Text), "dd/mm/yyyy")
        RSTTRXFILE!VCH_DESC = "Received From " & DataList2.Text
        RSTTRXFILE!PINV = Trim(TXTINVOICE.Text)
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    i = 0
    Set rstMaxRec = New ADODB.Recordset
    rstMaxRec.Open "Select MAX(Val(REC_NO)) From CASHATRXFILE ", db2, adOpenForwardOnly
    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
        i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
    End If
    rstMaxRec.Close
    Set rstMaxRec = Nothing

    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'PY'", db2, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!REC_NO = i + 1
        RSTITEMMAST!INV_TYPE = "PY"
        RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
    End If
    If lblcredit.Caption = "0" Then
        RSTITEMMAST!TRX_TYPE = "DR"
    Else
        RSTITEMMAST!TRX_TYPE = "CR"
    End If
    RSTITEMMAST!ACT_CODE = DataList2.BoundText
    RSTITEMMAST!ACT_NAME = Trim(DataList2.Text)
    RSTITEMMAST!AMOUNT = Val(LBLTOTAL.Caption)
    RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTITEMMAST!CHECK_FLAG = "P"
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing

SKIP:
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(Val(VCH_NO)) From RTRXFILEWO WHERE TRX_TYPE = 'PI'", db2, adOpenStatic, adLockReadOnly
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
    TXTINVOICE.Text = ""
    txtremarks.Text = ""
    TXTSLNO.Text = ""
    TXTITEMCODE.Text = ""
    TXTPRODUCT.Text = ""
    TXTQTY.Text = ""
    TxtPack.Text = ""
    TXTFREE.Text = ""
    TxttaxMRP.Text = ""
    txtPD.Text = ""
    txtprofit.Text = ""
    txtretail.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    txtcrtn.Text = ""
    txtcrtnpack.Text = ""
    
    TXTRATE.Text = ""
    txtmrpbt.Text = ""
    TXTPTR.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    txtaddlamt.Text = ""
    txtcramt.Text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    TXTDISCAMOUNT.Text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    grdsales.Rows = 1
    CMDEXIT.Enabled = True
    OptComper.Value = True
    txtBillNo.SetFocus
    M_ADD = False
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "PURCHASE ENTRY"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number = 7 Then
        MsgBox "Select Supplier from the list", vbOKOnly, "Purchase Entry"
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
        Case Asc("'")
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
            
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub txtaddlamt_LostFocus()
    Dim DISC As Currency
    
    On Error GoTo ErrHand
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
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    txtaddlamt.SetFocus
End Sub

Private Sub txtcramt_GotFocus()
    txtcramt.SelStart = 0
    txtcramt.SelLength = Len(txtcramt.Text)
End Sub

Private Sub txtcramt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
            
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub txtcramt_LostFocus()
    Dim DISC As Currency
    
    On Error GoTo ErrHand
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
    LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    txtcramt.SetFocus
End Sub

Private Sub OPTTaxMRP_GotFocus()
    'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100
    lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
End Sub

Private Sub OPTVAT_GotFocus()
    'lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text) + Val(TxtFree.Text))
    lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * Val(TXTQTY.Text)
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
End Sub

Private Sub OPTNET_GotFocus()
    lbltaxamount.Caption = ""
    LBLSUBTOTAL.Caption = Format(Val(TXTQTY.Text) * Val(TXTPTR.Text), ".000")
End Sub

Private Sub txtprofit_GotFocus()
    txtprofit.SelStart = 0
    txtprofit.SelLength = Len(txtprofit.Text)
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
        Case Asc("'")
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
            txtPD.Enabled = False
            txtretail.Enabled = True
            txtretail.SetFocus
         Case vbKeyEscape
            txtPD.Enabled = False
            TXTPTR.Enabled = True
            TXTPTR.SetFocus
    End Select
End Sub

Private Sub txtPD_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPD_LostFocus()
    Call TxttaxMRP_LostFocus
    txtPD.Tag = ((Val(LBLSUBTOTAL.Caption) - Val(lbltaxamount.Caption)) * Val(txtPD.Text) / 100)
    LBLSUBTOTAL.Caption = Format(Val(LBLSUBTOTAL.Caption) - Val(txtPD.Tag), ".000")
    txtPD.Text = Format(txtPD.Text, "0.00")
End Sub


Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
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
    MsgBox Err.Description
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_Click()
    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Purchase Bill..."
                DataList2.SetFocus
                Exit Sub
            End If
            TXTINVOICE.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub TXTRETAIL_GotFocus()
    txtretail.SelStart = 0
    txtretail.SelLength = Len(txtretail.Text)
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtretail.Enabled = False
            txtWS.Enabled = True
            txtWS.SetFocus
         Case vbKeyEscape
            txtretail.Enabled = False
            txtPD.Enabled = True
            txtPD.SetFocus
    End Select
End Sub

Private Sub TXTRETAIL_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRETAIL_LostFocus()
    txtretail.Text = Format(txtretail.Text, "0.00")
End Sub

Private Sub txtws_GotFocus()
    txtWS.SelStart = 0
    txtWS.SelLength = Len(txtWS.Text)
End Sub

Private Sub txtws_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtWS.Enabled = False
            txtvanrate.Enabled = True
            txtvanrate.SetFocus
         Case vbKeyEscape
            txtWS.Enabled = False
            txtretail.Enabled = True
            txtretail.SetFocus
    End Select
End Sub

Private Sub txtws_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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

Private Sub txtcrtn_GotFocus()
    txtcrtn.SelStart = 0
    txtcrtn.SelLength = Len(txtcrtn.Text)
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
            txtWS.SetFocus
    End Select
End Sub

Private Sub txtcrtn_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
    OptComper.Value = True
End Sub

Private Sub TxtComper_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtComper.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            TxtComper.Enabled = False
            txtcrtnpack.Enabled = True
            txtcrtnpack.SetFocus
    End Select
End Sub

Private Sub TxtComper_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
    OptComAmt.Value = True
End Sub

Private Sub TxtComAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtComAmt.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            TxtComAmt.Enabled = False
            txtcrtn.Enabled = True
            txtcrtn.SetFocus
    End Select
End Sub

Private Sub TxtComAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
            Call cmditemcreate_KeyDown(27, 0)
'            TxtComAmt.Enabled = True
'            TxtComAmt.SetFocus
    End Select
End Sub

Private Sub OptComAmt_GotFocus()
    TxtComper.Text = ""
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
            Call cmditemcreate_KeyDown(27, 0)
'            TxtComper.Enabled = True
'            TxtComper.SetFocus
    End Select
End Sub

Private Sub OptComper_GotFocus()
    TxtComAmt.Text = ""
End Sub

Private Sub txtcrtnpack_GotFocus()
    If Val(txtcrtn.Text) = 0 Then txtcrtnpack.Text = 0
    txtcrtnpack.SelStart = 0
    txtcrtnpack.SelLength = Len(txtcrtnpack.Text)
End Sub

Private Sub txtcrtnpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtcrtn.Text) <> 0 And Val(txtcrtnpack.Text) = 0 Then
                MsgBox "Please enter the Pack Qty for Carton", vbOKOnly, "Purchase"
                txtcrtnpack.SetFocus
                Exit Sub
            End If
            If Val(txtcrtn.Text) = 0 And Val(txtcrtnpack.Text) <> 0 Then
                MsgBox "Please enter the Rate for Carton", vbOKOnly, "Purchase"
                txtcrtnpack.Enabled = False
                txtcrtn.Enabled = True
                txtcrtn.SetFocus
                Exit Sub
            End If
            txtcrtnpack.Enabled = False
            If OptComper.Value = True Then
                TxtComper.Enabled = True
                TxtComper.SetFocus
            Else
                TxtComAmt.Enabled = True
                TxtComAmt.SetFocus
            End If
         Case vbKeyEscape
            txtcrtnpack.Enabled = False
            txtcrtn.Enabled = True
            txtcrtn.SetFocus
    End Select
End Sub

Private Sub txtcrtnpack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
            txtvanrate.Enabled = False
            txtcrtn.Enabled = True
            txtcrtn.SetFocus
         Case vbKeyEscape
            txtvanrate.Enabled = False
            txtWS.Enabled = True
            txtWS.SetFocus
    End Select
End Sub

Private Sub txtvanrate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
