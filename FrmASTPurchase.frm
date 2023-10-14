VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAstPurchase 
   BackColor       =   &H0098DAA9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASSETS PURCHASE"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18645
   ControlBox      =   0   'False
   Icon            =   "FrmASTPurchase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   18645
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   45
      TabIndex        =   23
      Top             =   1995
      Visible         =   0   'False
      Width           =   14835
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   3675
         Left            =   30
         TabIndex        =   24
         Top             =   405
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6482
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
         Left            =   3795
         TabIndex        =   26
         Top             =   30
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
         TabIndex        =   25
         Top             =   30
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
      TabIndex        =   17
      Top             =   90
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
      Left            =   4185
      TabIndex        =   14
      Top             =   6840
      Width           =   990
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   2040
      TabIndex        =   1
      Top             =   1965
      Visible         =   0   'False
      Width           =   8700
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   4080
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   8670
         _ExtentX        =   15293
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
      BackColor       =   &H00CBDED7&
      Caption         =   "Frame1"
      Height          =   10275
      Left            =   -120
      TabIndex        =   0
      Top             =   -90
      Width           =   18690
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00CBDED7&
         Height          =   1575
         Left            =   135
         TabIndex        =   3
         Top             =   0
         Width           =   11490
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
            Left            =   5940
            MaxLength       =   20
            TabIndex        =   109
            Top             =   900
            Width           =   2355
         End
         Begin VB.OptionButton OptCredit 
            BackColor       =   &H00CBDED7&
            Caption         =   "Saved Suppliers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   6270
            TabIndex        =   31
            Top             =   180
            Value           =   -1  'True
            Width           =   2205
         End
         Begin VB.OptionButton OptCash 
            BackColor       =   &H00CBDED7&
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   5280
            TabIndex        =   30
            Top             =   180
            Width           =   1155
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
            TabIndex        =   21
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
            Left            =   11610
            TabIndex        =   15
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
            Left            =   7395
            MaxLength       =   100
            TabIndex        =   11
            Top             =   1230
            Width           =   4050
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
            TabIndex        =   10
            Top             =   210
            Width           =   1260
         End
         Begin VB.TextBox TXTphone 
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
            Left            =   5955
            MaxLength       =   20
            TabIndex        =   4
            Top             =   540
            Width           =   2355
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   9735
            TabIndex        =   13
            Top             =   885
            Width           =   1695
            _ExtentX        =   2990
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
            TabIndex        =   29
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bill No."
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
            Index           =   24
            Left            =   5265
            TabIndex        =   110
            Top             =   930
            Width           =   1215
         End
         Begin VB.Label lblcredit 
            Height          =   525
            Left            =   11565
            TabIndex        =   18
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
            Left            =   11505
            TabIndex        =   16
            Top             =   150
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Del. Order No. && Date"
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
            Left            =   5280
            TabIndex        =   12
            Top             =   1260
            Width           =   2085
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
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   210
            Width           =   870
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "PHONE"
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
            Left            =   5280
            TabIndex        =   7
            Top             =   570
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
            Left            =   8385
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   600
            Width           =   1005
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdsales 
         Height          =   4500
         Left            =   150
         TabIndex        =   22
         Top             =   1590
         Width           =   18495
         _ExtentX        =   32623
         _ExtentY        =   7938
         _Version        =   393216
         Rows            =   1
         Cols            =   37
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
      End
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00CBDED7&
         Height          =   4275
         Left            =   120
         TabIndex        =   32
         Top             =   6015
         Width           =   16560
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
            Left            =   5220
            TabIndex        =   108
            Top             =   915
            Width           =   1050
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CBDED7&
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   11070
            TabIndex        =   72
            Top             =   840
            Width           =   2280
            Begin VB.OptionButton Optdiscamt 
               BackColor       =   &H00CBDED7&
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
               Left            =   975
               TabIndex        =   74
               Top             =   135
               Width           =   1125
            End
            Begin VB.OptionButton optdiscper 
               BackColor       =   &H00CBDED7&
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
               Height          =   240
               Left            =   30
               TabIndex        =   73
               Top             =   135
               Value           =   -1  'True
               Width           =   945
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00CBDED7&
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   9975
            TabIndex        =   67
            Top             =   2295
            Width           =   2565
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
               TabIndex        =   69
               Top             =   150
               Width           =   945
            End
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
               TabIndex        =   68
               Top             =   510
               Width           =   945
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
               TabIndex        =   71
               Top             =   195
               Width           =   1050
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
               TabIndex        =   70
               Top             =   525
               Width           =   1470
            End
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
            Left            =   3240
            TabIndex        =   64
            Top             =   915
            Width           =   915
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
            TabIndex        =   63
            Top             =   3945
            Visible         =   0   'False
            Width           =   765
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
            Left            =   14010
            MaxLength       =   15
            TabIndex        =   62
            Top             =   4170
            Visible         =   0   'False
            Width           =   2190
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
            TabIndex        =   61
            Top             =   2895
            Visible         =   0   'False
            Width           =   3300
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
            TabIndex        =   60
            Top             =   915
            Width           =   1065
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
            Left            =   2265
            TabIndex        =   59
            Top             =   915
            Width           =   960
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
            Left            =   7905
            MaxLength       =   8
            TabIndex        =   58
            Top             =   480
            Width           =   915
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
            Left            =   2010
            TabIndex        =   57
            Top             =   480
            Width           =   4680
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
            TabIndex        =   56
            Top             =   480
            Width           =   540
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
            TabIndex        =   55
            Top             =   915
            Width           =   1095
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
            Left            =   13620
            MaxLength       =   7
            TabIndex        =   54
            Top             =   2355
            Visible         =   0   'False
            Width           =   1020
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
            Left            =   8835
            MaxLength       =   7
            TabIndex        =   53
            Top             =   465
            Width           =   1170
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
            TabIndex        =   52
            Top             =   2895
            Visible         =   0   'False
            Width           =   780
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
            Left            =   11520
            MaxLength       =   7
            TabIndex        =   51
            Top             =   465
            Width           =   780
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H0098DAA9&
            Caption         =   "GST%"
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
            Left            =   14835
            TabIndex        =   50
            Top             =   960
            Width           =   945
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H0098DAA9&
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
            Left            =   13410
            TabIndex        =   49
            Top             =   960
            Width           =   1410
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
            Left            =   12795
            MaxLength       =   7
            TabIndex        =   48
            Top             =   2370
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.OptionButton OPTNET 
            BackColor       =   &H0098DAA9&
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
            Left            =   15810
            TabIndex        =   47
            Top             =   960
            Value           =   -1  'True
            Width           =   690
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
            Left            =   6345
            TabIndex        =   46
            Top             =   1215
            Width           =   945
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
            Left            =   8370
            TabIndex        =   45
            Top             =   1215
            Width           =   1050
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
            Left            =   7305
            TabIndex        =   44
            Top             =   1215
            Width           =   1050
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
            Left            =   14910
            MaxLength       =   7
            TabIndex        =   43
            Top             =   2250
            Visible         =   0   'False
            Width           =   780
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
            TabIndex        =   42
            Top             =   3840
            Visible         =   0   'False
            Width           =   765
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
            Left            =   12315
            MaxLength       =   7
            TabIndex        =   41
            Top             =   450
            Width           =   825
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
            Left            =   10020
            MaxLength       =   7
            TabIndex        =   40
            Top             =   465
            Width           =   1485
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
            Left            =   13170
            MaxLength       =   7
            TabIndex        =   39
            Top             =   1665
            Visible         =   0   'False
            Width           =   870
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
            ItemData        =   "FrmASTPurchase.frx":030A
            Left            =   6705
            List            =   "FrmASTPurchase.frx":0332
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   480
            Width           =   1200
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
            ItemData        =   "FrmASTPurchase.frx":0377
            Left            =   10590
            List            =   "FrmASTPurchase.frx":0381
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   4830
            Visible         =   0   'False
            Width           =   1350
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
            Left            =   9795
            MaxLength       =   4
            TabIndex        =   36
            Top             =   4080
            Visible         =   0   'False
            Width           =   795
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
            Left            =   600
            TabIndex        =   35
            Top             =   480
            Width           =   1395
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
            Left            =   15075
            TabIndex        =   34
            Top             =   1515
            Width           =   1335
         End
         Begin VB.CheckBox Chkcancel 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "Cancel Bill"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   15075
            TabIndex        =   33
            Top             =   1290
            Width           =   1320
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   315
            Left            =   13695
            TabIndex        =   65
            Top             =   3960
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
            TabIndex        =   66
            Top             =   3960
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
            TabIndex        =   107
            Top             =   3615
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
            Left            =   12945
            TabIndex        =   106
            Top             =   3660
            Visible         =   0   'False
            Width           =   765
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
            Left            =   14265
            TabIndex        =   105
            Top             =   450
            Width           =   1635
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
            Height          =   285
            Index           =   7
            Left            =   14010
            TabIndex        =   104
            Top             =   3885
            Visible         =   0   'False
            Width           =   2190
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
            TabIndex        =   103
            Top             =   3735
            Visible         =   0   'False
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
            Left            =   11085
            TabIndex        =   102
            Top             =   3885
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
            Height          =   255
            Index           =   14
            Left            =   14265
            TabIndex        =   101
            Top             =   195
            Width           =   1635
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
            Left            =   13620
            TabIndex        =   100
            Top             =   2085
            Visible         =   0   'False
            Width           =   1020
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
            Left            =   7905
            TabIndex        =   99
            Top             =   180
            Width           =   915
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
            Left            =   2010
            TabIndex        =   98
            Top             =   195
            Width           =   4680
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
            TabIndex        =   97
            Top             =   195
            Width           =   540
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
            Left            =   8835
            TabIndex        =   96
            Top             =   195
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
            Height          =   285
            Index           =   4
            Left            =   3915
            TabIndex        =   95
            Top             =   2610
            Visible         =   0   'False
            Width           =   780
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
            Left            =   11520
            TabIndex        =   94
            Top             =   195
            Width           =   780
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
            Left            =   13155
            TabIndex        =   93
            Top             =   195
            Width           =   1095
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
            Left            =   12795
            TabIndex        =   92
            Top             =   2070
            Visible         =   0   'False
            Width           =   810
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
            Left            =   9420
            TabIndex        =   91
            Top             =   1500
            Width           =   1605
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
            Left            =   9435
            TabIndex        =   90
            Top             =   1725
            Width           =   1590
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
            Left            =   6345
            TabIndex        =   89
            Top             =   990
            Width           =   915
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
            Left            =   11040
            TabIndex        =   88
            Top             =   1725
            Width           =   1500
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
            Left            =   11160
            TabIndex        =   87
            Top             =   1500
            Width           =   1185
            WordWrap        =   -1  'True
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
            Left            =   8355
            TabIndex        =   86
            Top             =   975
            Width           =   1140
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
            Left            =   7305
            TabIndex        =   85
            Top             =   975
            Width           =   1020
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
            Left            =   12315
            TabIndex        =   84
            Top             =   195
            Width           =   825
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
            TabIndex        =   83
            Top             =   3765
            Visible         =   0   'False
            Width           =   780
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
            Left            =   10020
            TabIndex        =   82
            Top             =   195
            Width           =   1485
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
            Left            =   6705
            TabIndex        =   81
            Top             =   180
            Width           =   1185
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
            TabIndex        =   80
            Top             =   2610
            Visible         =   0   'False
            Width           =   3300
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
            Height          =   285
            Index           =   39
            Left            =   9795
            TabIndex        =   79
            Top             =   4575
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Item Part"
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
            Left            =   600
            TabIndex        =   78
            Top             =   195
            Width           =   1395
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
            Left            =   13155
            TabIndex        =   77
            Top             =   450
            Width           =   1095
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
            TabIndex        =   76
            Top             =   2610
            Visible         =   0   'False
            Width           =   1635
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
            TabIndex        =   75
            Top             =   2865
            Visible         =   0   'False
            Width           =   1635
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase for this month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   570
         Index           =   50
         Left            =   11580
         TabIndex        =   28
         Top             =   990
         Width           =   1875
         WordWrap        =   -1  'True
      End
      Begin VB.Label LBLmonth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   13380
         TabIndex        =   27
         Top             =   1050
         Width           =   1980
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   135
         TabIndex        =   20
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   705
         TabIndex        =   19
         Top             =   45
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmAstPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytData() As Byte
Dim PHY As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim ACT_FLAG As Boolean
Dim PHY_CODE As New ADODB.Recordset
Dim PHYCODE_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT, M_ADD, OLD_BILL As Boolean
Dim PHY_PRERATE As New ADODB.Recordset
Dim PRERATE_FLAG As Boolean

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = "1"
            TXTQTY.SetFocus
         Case vbKeyEscape
            'TXTUNIT.Text = ""
            If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            CmbPack.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub CMDADD_Click()
        
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
    If Val(TXTQTY.text) = 0 Then
        MsgBox "Please enter the Qty", vbOKOnly, "EzBiz"
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
        Exit Sub
    End If
    If Val(TXTPTR.text) = 0 Then
        MsgBox "Please enter the Price", vbOKOnly, "EzBiz"
        TXTPTR.SetFocus
        Exit Sub
    End If
    'Call TXTPTR_LostFocus
    'Call Txtgrossamt_LostFocus
    Call txtPD_LostFocus
    
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim M_DATA As Double
    
    M_DATA = 0
    Txtpack.text = 1
    If grdsales.rows <= Val(TXTSLNO.text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.text), 0) = Val(TXTSLNO.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 1) = Trim(TXTITEMCODE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 2) = Trim(TXTPRODUCT.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 3) = Val(TXTQTY.text) + Val(TXTFREE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 4) = 1 'Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 5) = Val(Los_Pack.text) ' 1 'Val(TxtPack.Text)
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Round(Val(TXTRATE.Text) / Val(Los_Pack.Text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 6) = Format(Val(TXTRATE.text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 8) = Format(Round(((Val(LblGross.Caption) / (Val(Los_Pack.text) * TXTQTY.text))), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 9) = Format(Round(Val(TXTPTR.text) / Val(Los_Pack.text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 7) = Format((Val(txtprofit.text)), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 10) = IIf(Val(TxttaxMRP.text) = 0, "", Format(Val(TxttaxMRP.text), ".00")) 'TAX
    grdsales.TextMatrix(Val(TXTSLNO.text), 11) = Trim(txtBatch.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 12) = "" 'IIf(Trim(TXTEXPDATE.Text) = "/  /", "", TXTEXPDATE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 13) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 14) = Val(TXTFREE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 17) = Val(txtPD.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 26) = Format(Val(Txtgrossamt.text), ".00")
    If optdiscper.Value = True Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 27) = "P"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.text), 27) = "A"
    End If
    grdsales.TextMatrix(Val(TXTSLNO.text), 28) = Format(Val(Los_Pack.text), ".00")
    grdsales.TextMatrix(Val(TXTSLNO.text), 29) = Trim(CmbPack.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 30) = Val(TxtWarranty.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 36) = Val(LblGross.Caption)
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
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From ASTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!TRX_TYPE = "AP"
        RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTRTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTRTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16))
        RSTRTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1))
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
        RSTRTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))

        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ASTMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                RSTRTRXFILE!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
'                If UCase(rststock!CATEGORY) = "CUTSHEET" Then
'                Else
                !ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 8))
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                '!CLOSE_VAL = !CLOSE_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !CLOSE_VAL = Round(!ITEM_COST * !CLOSE_QTY, 3)
                !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                '!RCPT_VAL = !RCPT_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !RCPT_VAL = Round(!ITEM_COST * !RCPT_QTY, 3)
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))) >= 5 Then !SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))
                '!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
                !CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15))
                !PACK_TYPE = Trim(CmbPack.text)
                RSTRTRXFILE!MFGR = !MANUFACTURER
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        
    Else
        M_DATA = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
        M_DATA = M_DATA - (RSTRTRXFILE!QTY - RSTRTRXFILE!BAL_QTY)
        RSTRTRXFILE!BAL_QTY = M_DATA
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ASTMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                RSTRTRXFILE!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
                '!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
                !ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 8))
                !CLOSE_QTY = !CLOSE_QTY - RSTRTRXFILE!QTY
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                '!CLOSE_VAL = !CLOSE_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !CLOSE_VAL = Round(!ITEM_COST * !CLOSE_QTY, 3)
                
                !RCPT_QTY = !RCPT_QTY - RSTRTRXFILE!QTY
                !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                '!RCPT_VAL =  !RCPT_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !RCPT_VAL = Round(!ITEM_COST * !RCPT_QTY, 3)
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))) >= 5 Then !SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))
                '!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
                !CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15))
                !PACK_TYPE = Trim(CmbPack.text)
                RSTRTRXFILE!MFGR = !MANUFACTURER
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
    End If
    RSTRTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 13))
    RSTRTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "dd/mm/yyyy")
    RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 2))
    RSTRTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 8))
    RSTRTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
    RSTRTRXFILE!P_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 17))
    RSTRTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 9))
    RSTRTRXFILE!gross_amt = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 26))
    RSTRTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))
    RSTRTRXFILE!PACK_TYPE = Trim(CmbPack.text)
    RSTRTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
    'RSTRTRXFILE!VCH_DESC = "Received From " & DataList2.Text
    RSTRTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 11))
    'RSTRTRXFILE!ISSUE_QTY = 0
    RSTRTRXFILE!CST = 0
    If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 27)) = "P" Then
        RSTRTRXFILE!DISC_FLAG = "P"
    Else
        RSTRTRXFILE!DISC_FLAG = "A"
    End If
    RSTRTRXFILE!FREE_QTY = 0
    RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!C_USER_ID = "SM"
    RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15))
    
    'RSTRTRXFILE!M_USER_ID = DataList2.BoundText
    ''''RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))  'MODE OF TAX
    'RSTRTRXFILE!PINV = Trim(TXTINVOICE.Text)
    RSTRTRXFILE.Update
    RSTRTRXFILE.Close
    
    M_DATA = 0
    Set RSTRTRXFILE = Nothing
    
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From ASTRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!TRX_TYPE = "AP"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
           
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    For i = 1 To grdsales.rows - 1
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
    Next i
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    TXTSLNO.text = grdsales.rows
    TXTPRODUCT.text = ""
    
    TXTITEMCODE.text = ""
    TXTPTR.text = ""
    Txtgrossamt.text = ""
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    txtPD.text = ""
    txtprofit.text = ""
    Txtgrossamt.text = ""
    TXTRATE.text = ""
    txtmrpbt.text = ""
    txtBatch.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    optnet.Value = True
    M_ADD = True
    Chkcancel.Value = 0
    OLD_BILL = True
    txtcategory.Enabled = True
    txtBillNo.Enabled = False
    FRMEGRDTMP.Visible = False
    cmdRefresh.Enabled = True
    Los_Pack.Enabled = False
    CmbPack.Enabled = False
    TXTQTY.Enabled = False
    TXTFREE.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxttaxMRP.Enabled = False
    txtPD.Enabled = False
    cmdadd.Enabled = False
    Txtgrossamt.Enabled = False
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
    If M_EDIT = True Then
        TXTSLNO.Enabled = True
        txtcategory.Enabled = False
        grdsales.SetFocus
    Else
        If grdsales.rows >= 11 Then grdsales.TopRow = grdsales.rows - 1
        txtcategory.SetFocus
    End If
    M_EDIT = False
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtPD.SetFocus
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
   
    On Error GoTo ErrHand
    db.Execute "delete  From ASTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & ""
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM ASTMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With rststock
        If Not (.EOF And .BOF) Then
            !RCPT_QTY = !RCPT_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
            If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
            !RCPT_VAL = !RCPT_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 13))
            
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
    rstMaxNo.Open "Select MAX(LINE_NO) From ASTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From ASTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
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
    RSTRTRXFILE.Open "SELECT * From ASTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
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
    grdsales.rows = 1
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "Select * From ASTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTRTRXFILE.EOF
        grdsales.rows = grdsales.rows + 1
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
        If RSTRTRXFILE!DISC_FLAG = "P" Then
            grdsales.TextMatrix(i, 27) = "P"
        Else
            grdsales.TextMatrix(i, 27) = "A"
        End If
        grdsales.TextMatrix(i, 24) = IIf(IsNull(RSTRTRXFILE!CRTN_PACK), 0, RSTRTRXFILE!CRTN_PACK)
        grdsales.TextMatrix(i, 25) = IIf(IsNull(RSTRTRXFILE!P_VAN), 0, RSTRTRXFILE!P_VAN)
        grdsales.TextMatrix(i, 26) = IIf(IsNull(RSTRTRXFILE!gross_amt), 0, RSTRTRXFILE!gross_amt)
        grdsales.TextMatrix(i, 28) = IIf(IsNull(RSTRTRXFILE!LOOSE_PACK), 1, RSTRTRXFILE!LOOSE_PACK)
        grdsales.TextMatrix(i, 29) = IIf(IsNull(RSTRTRXFILE!PACK_TYPE), "Nos", RSTRTRXFILE!PACK_TYPE)
        grdsales.TextMatrix(i, 30) = IIf(IsNull(RSTRTRXFILE!WARRANTY), "", RSTRTRXFILE!WARRANTY)
        grdsales.TextMatrix(i, 31) = IIf(IsNull(RSTRTRXFILE!WARRANTY_TYPE), "", RSTRTRXFILE!WARRANTY_TYPE)
        grdsales.TextMatrix(i, 32) = IIf(IsNull(RSTRTRXFILE!EXPENSE), "", RSTRTRXFILE!EXPENSE)
        grdsales.TextMatrix(i, 33) = IIf(IsNull(RSTRTRXFILE!EXDUTY), "", RSTRTRXFILE!EXDUTY)
        grdsales.TextMatrix(i, 34) = IIf(IsNull(RSTRTRXFILE!CSTPER), "", RSTRTRXFILE!CSTPER)
        grdsales.TextMatrix(i, 35) = IIf(IsNull(RSTRTRXFILE!TR_DISC), "", RSTRTRXFILE!TR_DISC)
        grdsales.TextMatrix(i, 36) = IIf(IsNull(RSTRTRXFILE!GROSS_AMOUNT), "", RSTRTRXFILE!GROSS_AMOUNT)
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    
    TXTSLNO.text = Val(grdsales.rows)
    TXTPRODUCT.text = ""
    TXTITEMCODE.text = ""
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    txtPD.text = ""
    txtprofit.text = ""
    
    Txtgrossamt.text = ""
    
    
    TXTRATE.text = ""
    
    
    txtmrpbt.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    txtBatch.text = ""
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
ErrHand:
    MsgBox err.Description
End Sub

Private Sub CmdDeleteAll_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    On Error GoTo ErrHand
    If Chkcancel.Value = 0 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE ALL", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
   
    For i = 1 To grdsales.rows - 1
        db.Execute "delete  From ASTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & ""
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ASTMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
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
        rststock.Close
        Set rststock = Nothing
    Next i
    
    grdsales.FixedRows = 0
    grdsales.rows = 1
    Call appendpurchase
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.text) >= grdsales.rows Then Exit Sub
    
    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    Los_Pack.Enabled = True
    CmbPack.Enabled = True
    TXTQTY.Enabled = True
    TXTFREE.Enabled = True
    TXTRATE.Enabled = True
    TXTPTR.Enabled = True
    TxttaxMRP.Enabled = True
    txtPD.Enabled = True
    cmdadd.Enabled = True
    Txtgrossamt.Enabled = True
    TXTQTY.SetFocus
    
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            TXTQTY.text = ""
            Txtpack.text = 1 '""
            Los_Pack.text = ""
            CmbPack.ListIndex = -1
            TxtWarranty.text = ""
            CmbWrnty.ListIndex = -1
            TXTFREE.text = ""
            TxttaxMRP.text = ""
            
            
            
            txtPD.text = ""
            
            txtprofit.text = ""
            
            
            
            
           
            
            Txtgrossamt.text = ""
            
            
            TXTRATE.text = ""
            
            
            txtmrpbt.text = ""
            TXTITEMCODE.text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            lbltaxamount.Caption = ""
            TXTEXPDATE.text = "  /  /    "
            TXTEXPIRY.text = "  /  "
            txtBatch.text = ""
        
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
    If grdsales.rows <= 1 Then
        lblcredit.Caption = "0"
        Call appendpurchase
    Else
        If optCash.Value = True Then lblcredit.Caption = "0"
        If OptCredit.Value = True And IsNull(DataList2.SelectedItem) Then
            MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
            DataList2.SetFocus
            Exit Sub
        End If
'        If TXTINVOICE.Text = "" Then
'            MsgBox "Enter Supplier Invoice No.", vbOKOnly, "EzBiz"
'            Exit Sub
'        End If
        If Not IsDate(TXTINVDATE.text) Then
            MsgBox "Enter Supplier Invoice Date", vbOKOnly, "EzBiz"
            Exit Sub
        End If
        If OptCredit.Value = True Then
            Me.Enabled = False
            MDIMAIN.cmdpurchase.Enabled = False
            Set creditbill = Me
            frmCREDIT.Show
        Else
            Call appendpurchase
        End If
    End If
    
End Sub

Private Sub cmdRefresh_GotFocus()
    FRMEGRDTMP.Visible = False
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrHand
    txtBillNo.SetFocus
    Exit Sub
ErrHand:
    If err.Number = 5 Then Exit Sub
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From ASTRXFILE WHERE TRX_TYPE = 'AP'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    ACT_FLAG = True
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
    grdsales.TextArray(20) = "Loose Price"
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
    
    PHYFLAG = True
    PHYCODE_FLAG = True
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    'TXTDATE.Text = Date
    TXTEXPDATE.Enabled = False
    txtBatch.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    TXTUNIT.Enabled = False
    TXTSLNO.text = 1
    TXTSLNO.Enabled = True
    FRMECONTROLS.Enabled = False
    FRMEMASTER.Enabled = False
    CLOSEALL = 1
    lblcredit.Caption = "1"
    TXTDEALER.text = ""
    M_ADD = False
    'Me.Width = 15135
    'Me.Height = 9660
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If PHYCODE_FLAG = False Then PHY_CODE.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If PRERATE_FLAG = False Then PHY_PRERATE.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            
            On Error Resume Next
            TXTITEMCODE.text = grdtmp.Columns(0)
            TXTPRODUCT.text = grdtmp.Columns(1)
'            On Error Resume Next
'            Set Image1.DataSource = PHY
'            If IsNull(PHY!PHOTO) Then
'                Frame6.Visible = False
'                Set Image1.DataSource = Nothing
'                bytData = ""
'            Else
'                If Err.Number = 545 Then
'                    Frame6.Visible = False
'                    Set Image1.DataSource = Nothing
'                    bytData = ""
'                Else
'                    Frame6.Visible = True
'                    Set Image1.DataSource = PHY 'setting image1’s datasource
'                    Image1.DataField = "PHOTO"
'                    bytData = PHY!PHOTO
'                End If
'            End If
            On Error GoTo ErrHand
            For i = 1 To grdsales.rows - 1
                If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                    If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                    Exit For
                End If
            Next i
            
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "Select * From ASTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                'RSTRXFILE.MoveLast
                TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                Los_Pack.text = 1 'IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                If IsNull(RSTRXFILE!LINE_DISC) Then
                    Txtpack.text = ""
                Else
                    Txtpack.text = RSTRXFILE!LINE_DISC
                End If
                Txtpack.text = 1
                TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                If IsNull(RSTRXFILE!REF_NO) Then
                    txtBatch.text = ""
                Else
                    txtBatch.text = RSTRXFILE!REF_NO
                End If
                If IsNull(RSTRXFILE!PTR) Then
                    TXTPTR.text = ""
                Else
                    TXTPTR.text = Format(Round(Val(RSTRXFILE!PTR), 2), ".000")
                End If
                If IsNull(RSTRXFILE!SALES_TAX) Then
                    TxttaxMRP.text = ""
                Else
                    TxttaxMRP.text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                End If
                On Error Resume Next
                CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                On Error GoTo ErrHand
                
                ''TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                If RSTRXFILE!CHECK_FLAG = "M" Then
                    OPTTaxMRP.Value = True
                ElseIf RSTRXFILE!CHECK_FLAG = "V" Then
                    OPTVAT.Value = True
                Else
                    optnet.Value = True
                End If
            Else
                TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = 1
                    Los_Pack.text = 1
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.text = ""
                    TXTEXPIRY.text = "  /  "
                    TXTRATE.text = ""
                    txtmrpbt.text = ""
                    TXTPTR.text = ""
                    
                   
                    
                    
                    
                    txtprofit.text = ""
                    TxttaxMRP.text = "5"
                    Los_Pack.text = "1"
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    OPTVAT.Value = True
            End If
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing
            
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "SELECT *  FROM ASTMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With RSTRXFILE
                If Not (.EOF And .BOF) Then
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.text = ""
                    Else
                        TxttaxMRP.text = Format(Round(Val(RSTRXFILE!SALES_TAX), 2), ".000")
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
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
            'TxtPack.Enabled = True
            'TxtPack.SetFocus
        Case vbKeyEscape
            TXTQTY.text = ""
            TXTFREE.text = ""
            Fram.Enabled = True
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
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
            If Val(TxttaxMRP.text) <> 0 Then
                If OPTTaxMRP.Value = False And OPTVAT.Value = False Then
                'If OPTVAT.Value = False Then
                    MsgBox "Tax should be Zero ....", vbOKOnly, "EzBiz"
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
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "EzBiz"
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
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "EzBiz"
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

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(txtBatch.Text) = "" Then Exit Sub
            txtBatch.Enabled = False
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
        Case vbKeyEscape
            TXTFREE.Enabled = True
            txtBatch.Enabled = False
            TXTFREE.SetFocus
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
    txtBillNo.SelLength = Len(txtBillNo.text)
End Sub

Public Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTDIST As ADODB.Recordset
    Dim RSTTRNSMAST As ADODB.Recordset
    Dim i As Long

    On Error GoTo ErrHand
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
            Chkcancel.Value = 0
            grdsales.rows = 1
            i = 0
            LBLTOTAL.Caption = ""
            lbltotalwodiscount = ""
            grdsales.rows = 1
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From ASTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until rstTRXMAST.EOF
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                i = i + 1
                
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = rstTRXMAST!ITEM_CODE
                grdsales.TextMatrix(i, 2) = rstTRXMAST!ITEM_NAME
                grdsales.TextMatrix(i, 3) = Val(rstTRXMAST!QTY) / Val(rstTRXMAST!LINE_DISC)
                grdsales.TextMatrix(i, 4) = rstTRXMAST!UNIT
                grdsales.TextMatrix(i, 5) = rstTRXMAST!LINE_DISC
                grdsales.TextMatrix(i, 6) = 0 'Format(rstTRXMAST!MRP, ".000")
                grdsales.TextMatrix(i, 7) = 0 'Format(rstTRXMAST!SALES_PRICE, ".000")
                grdsales.TextMatrix(i, 8) = Format(rstTRXMAST!ITEM_COST, ".000")
                grdsales.TextMatrix(i, 9) = Format(rstTRXMAST!PTR, ".000")
                grdsales.TextMatrix(i, 10) = IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
                grdsales.TextMatrix(i, 11) = IIf(IsNull(rstTRXMAST!REF_NO), "", rstTRXMAST!REF_NO)
                grdsales.TextMatrix(i, 12) = "" ' IIf(IsNull(rstTRXMAST!EXP_DATE), "", Format(rstTRXMAST!EXP_DATE, "DD/MM/YYYY"))
                grdsales.TextMatrix(i, 13) = Format(rstTRXMAST!TRX_TOTAL, ".000")
                grdsales.TextMatrix(i, 14) = IIf(IsNull(rstTRXMAST!SCHEME), "", rstTRXMAST!SCHEME)
                grdsales.TextMatrix(i, 15) = IIf(IsNull(rstTRXMAST!CHECK_FLAG), "N", rstTRXMAST!CHECK_FLAG)
                grdsales.TextMatrix(i, 16) = rstTRXMAST!LINE_NO
                grdsales.TextMatrix(i, 17) = IIf(IsNull(rstTRXMAST!P_DISC), 0, rstTRXMAST!P_DISC)
                grdsales.TextMatrix(i, 26) = IIf(IsNull(rstTRXMAST!gross_amt), 0, Format(rstTRXMAST!gross_amt, "0.00"))
                If rstTRXMAST!DISC_FLAG = "P" Then
                    grdsales.TextMatrix(i, 27) = "P"
                Else
                    grdsales.TextMatrix(i, 27) = "A"
                End If
                grdsales.TextMatrix(i, 29) = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)
                lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
                'TXTDEALER.Text = IIf(IsNull(rstTRXMAST!VCH_DESC), "", Mid(rstTRXMAST!VCH_DESC, 15))
                On Error Resume Next
                TXTINVDATE.text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                On Error GoTo ErrHand
                TXTREMARKS.text = IIf(IsNull(rstTRXMAST!VCH_DESC), "", Mid(rstTRXMAST!VCH_DESC, 15))
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From ASTRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                TXTDISCAMOUNT.text = IIf(IsNull(rstTRXMAST!DISCOUNT), "", Format(rstTRXMAST!DISCOUNT, ".00"))
                txtaddlamt.text = IIf(IsNull(rstTRXMAST!ADD_AMOUNT), "", Format(rstTRXMAST!ADD_AMOUNT, ".00"))
                txtcramt.text = IIf(IsNull(rstTRXMAST!DISC_PERS), "", Format(rstTRXMAST!DISC_PERS, ".00"))
                TxtCST.text = IIf(IsNull(rstTRXMAST!CST_PER), "", Format(rstTRXMAST!CST_PER, ".00"))
                TxtInsurance.text = IIf(IsNull(rstTRXMAST!INS_PER), "", Format(rstTRXMAST!INS_PER, ".00"))
                If rstTRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                If rstTRXMAST!Cash_Flag = "Y" Then optCash.Value = True Else OptCredit.Value = True
                On Error Resume Next
                TXTINVDATE.text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                TXTDATE.text = Format(rstTRXMAST!CREATE_DATE, "DD/MM/YYYY")
                On Error GoTo ErrHand
                TXTINVOICE.text = IIf(IsNull(rstTRXMAST!PINV), "", rstTRXMAST!PINV)
                TXTDEALER.text = IIf(IsNull(rstTRXMAST!ACT_NAME), "", rstTRXMAST!ACT_NAME)
                TxtPhone.text = IIf(IsNull(rstTRXMAST!REF_NO), "", rstTRXMAST!REF_NO)
                
                OLD_BILL = True
            Else
                TXTDATE.text = Format(Date, "DD/MM/YYYY")
                OLD_BILL = False
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            ''''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
            'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
            LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
            
            TXTSLNO.text = grdsales.rows
            TXTSLNO.Enabled = True
            txtBillNo.Enabled = False
            FRMEMASTER.Enabled = True
            If i > 0 Or (Val(txtBillNo.text) < Val(TXTLASTBILL.text)) Then
                FRMEMASTER.Enabled = True
                FRMECONTROLS.Enabled = True
                cmdRefresh.Enabled = True
                cmdRefresh.SetFocus
            Else
                TXTDEALER.SetFocus
            End If
            
'            Set RSTTRNSMAST = New ADODB.Recordset
'            RSTTRNSMAST.Open "Select CHECK_FLAG From ASTRXMAST WHERE TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
'            If Not (RSTTRNSMAST.EOF Or RSTTRNSMAST.BOF) Then
'                If RSTTRNSMAST!CHECK_FLAG = "Y" Then FRMEMASTER.Enabled = False
'            End If
'            RSTTRNSMAST.Close
'            Set RSTTRNSMAST = Nothing
    
    End Select
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
    Exit Sub
ErrHand:
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
    If Val(txtBillNo.text) = 0 Or Val(txtBillNo.text) > Val(TXTLASTBILL.text) Then txtBillNo.text = TXTLASTBILL.text
End Sub

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.text)
    FRMEGRDTMP.Visible = False
    Call TXTSLNO_KeyDown(13, 0)
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

Private Sub TxtFree_GotFocus()
    TXTFREE.SelStart = 0
    TXTFREE.SelLength = Len(TXTFREE.text)
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
    If Val(TXTFREE.text) = 0 Then TXTFREE.text = 0
    TXTFREE.text = Format(TXTFREE.text, "0.00")
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
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
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTINVOICE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVOICE.text = "" Then Exit Sub
'            If OptCredit.value = True Then
'                Set rstTRXMAST = New ADODB.Recordset
'                rstTRXMAST.Open "Select * From ASTRXMAST WHERE TRX_TYPE='AP' AND PINV = '" & Trim(TXTINVOICE.Text) & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
'                If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
'                    MsgBox "You have already entered this Invoice number for " & Trim(DataList2.Text) & " as Computer Bill No. " & rstTRXMAST!VCH_NO, vbOKOnly, "EzBiz"
'                    TXTINVOICE.SetFocus
'                Else
'                    TXTINVDATE.SetFocus
'                End If
'                rstTRXMAST.Close
'                Set rstTRXMAST = Nothing
'            End If
            TXTINVDATE.SetFocus
        Case vbKeyEscape
            DataList2.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTINVOICE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtpack_GotFocus()
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.text)
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtpack.text) = 0 Then Exit Sub
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

Private Sub TxtPhone_GotFocus()
    TxtPhone.SelStart = 0
    TxtPhone.SelLength = Len(TxtPhone.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TxtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            'If TXTINVOICE.Text = "" Then Exit Sub
            TXTINVOICE.SetFocus
        Case vbKeyEscape
            DataList2.SetFocus
    End Select
    
End Sub

Private Sub TXTPRODUCT_Change()
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHand
        
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
        
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            PHY.Open "Select * From ASTMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ITEM_NAME Like '%" & Me.txtcategory.text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         Else
             PHY.Close
             PHY.Open "Select * From ASTMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ITEM_NAME Like '%" & Me.txtcategory.text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
        Exit Sub
ErrHand:
        MsgBox err.Description
                
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.text)
    Call TXTPRODUCT_Change
    CmbPack.Enabled = False
    TXTQTY.Enabled = False
    TXTFREE.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxttaxMRP.Enabled = False
    txtPD.Enabled = False
    cmdadd.Enabled = False
    Txtgrossamt.Enabled = False
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE, RSTITEMMAST  As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
        Case vbKeyReturn
            
            On Error Resume Next
            TXTITEMCODE.text = ""
            TXTITEMCODE.text = grdtmp.Columns(0)
            If Trim(TXTITEMCODE.text) = "" Then
                If MsgBox("Item not exists!!! Do You want to add this item?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
                TXTPRODUCT.Tag = ""
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ASTMAST ", db, adOpenStatic, adLockReadOnly
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
                RSTITEMMAST.Open "SELECT * FROM ASTMAST WHERE ITEM_CODE = '" & TXTPRODUCT.Tag & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                RSTITEMMAST.AddNew
                'RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
                RSTITEMMAST!ITEM_CODE = TXTPRODUCT.Tag
                RSTITEMMAST!ITEM_NAME = Trim(TXTPRODUCT.text)
                RSTITEMMAST!Category = "GENERAL"
                RSTITEMMAST!UNIT = 1
                RSTITEMMAST!MANUFACTURER = "GENERAL"
                RSTITEMMAST!DEAD_STOCK = "N"
                RSTITEMMAST!REMARKS = ""
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
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                TXTITEMCODE.text = TXTPRODUCT.Tag
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
            If Trim(TXTPRODUCT.text) = "" Then
                txtcategory.Enabled = True
                txtcategory.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False
                
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select * From ASTMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select * From ASTMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' and ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            
            If PHY.RecordCount = 0 Then
                If MsgBox("Item not exists!!! Do You want to add this item?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
                frmitemmaster.Show
                frmitemmaster.TXTITEM.text = Trim(TXTPRODUCT.text)
                'MsgBox "Item not found!!!!", , "EzBiz"
                Exit Sub
            End If
            
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.text = grdtmp.Columns(0)
                TXTPRODUCT.text = grdtmp.Columns(1)
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                        Exit For
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From ASTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    'RSTRXFILE.MoveLast
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.text = 1
                    Los_Pack.text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    TxtWarranty.text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo ErrHand
                    
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.text = IIf(IsNull(RSTRXFILE!REF_NO), "", RSTRXFILE!REF_NO)
                    TXTEXPIRY.text = IIf(IsDate(RSTRXFILE!EXP_DATE), Format(RSTRXFILE!EXP_DATE, "MM/YY"), "  /  ")
                    If (IsNull(RSTRXFILE!MRP)) Then
                        TXTRATE.text = ""
                    Else
                        TXTRATE.text = Format(Round(Val(RSTRXFILE!MRP) * Val(Los_Pack.text), 2), ".000")
                    End If
                    If (IsNull(RSTRXFILE!MRP_BT)) Then
                        txtmrpbt.text = 100 * Val(TXTRATE.text) / 105
                    Else
                        txtmrpbt.text = Val(TXTRATE.text)
                    End If
                    If IsNull(RSTRXFILE!PTR) Then
                        TXTPTR.text = ""
                    Else
                        TXTPTR.text = Format(Round(Val(RSTRXFILE!PTR), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.text = ""
                    Else
                        TxttaxMRP.text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                    End If
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    On Error GoTo ErrHand
                
                    'TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                    If RSTRXFILE!CHECK_FLAG = "M" Then
                        OPTTaxMRP.Value = True
                    ElseIf RSTRXFILE!CHECK_FLAG = "V" Then
                        OPTVAT.Value = True
                    Else
                        optnet.Value = True
                    End If
                Else
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = 1
                    Los_Pack.text = 1
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.text = ""
                    TXTEXPIRY.text = "  /  "
                    TXTRATE.text = ""
                    txtmrpbt.text = ""
                    TXTPTR.text = ""
                    
                   
                    
                    
                    
                    txtprofit.text = ""
                    TxttaxMRP.text = "5"
                    
                    
                    
                    Los_Pack.text = "1"
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    OPTVAT.Value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                
                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    TXTITEMCODE.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
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
                grdtmp.SetFocus
            End If
            
        Case vbKeyEscape
            txtcategory.Enabled = True
            TXTPRODUCT.Enabled = False
            txtcategory.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTPTR_GotFocus()
    TXTPTR.SelStart = 0
    TXTPTR.SelLength = Len(TXTPTR.text)
    Call FILL_PREVIIOUSRATE
End Sub

Private Sub TXTPTR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTPTR.text) = 0 Then Exit Sub
            TxttaxMRP.SetFocus
        Case vbKeyEscape
            TXTQTY.SetFocus
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
    Txtgrossamt.text = Val(TXTPTR.text) * Val(TXTQTY.text)
    TXTPTR.text = Format(TXTPTR.text, ".000")
    'TXTRETAIL.Text = Round(Val(txtmrpbt.Text) * 0.8, 2)
'    txtretail.Text = Format(Round(Val(TXTRATE.Text) - (Val(txtmrpbt.Text) * 20 / 100), 3), ".000")
'    txtprofit.Text = Format(Round(Val(txtretail.Text) - Val(txtretail.Text) * 10 / 100, 3), ".000")
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.text)
    FRMEGRDTMP.Visible = False
    CmbPack.Enabled = True
    Los_Pack.Enabled = True
    TXTQTY.Enabled = True
    TXTPTR.Enabled = True
    TxttaxMRP.Enabled = True
    txtPD.Enabled = True
    cmdadd.Enabled = True
    Txtgrossamt.Enabled = True
    txtBatch.Enabled = True
    txtcategory.Enabled = False
    TXTPRODUCT.Enabled = False
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.text)
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.text) = 0 Then Exit Sub
            TXTPTR.SetFocus
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
    TXTQTY.text = Format(TXTQTY.text, ".00")
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Round(Val(TXTPTR.text), 2)), ".000")
    LblGross.Caption = Format((Val(TXTQTY.text) * Round(Val(TXTPTR.text), 2)), ".000")
End Sub

Private Sub TXTRATE_GotFocus()
    TXTRATE.SelStart = 0
    TXTRATE.SelLength = Len(TXTRATE.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TXTRATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTRATE.Text) = 0 Then Exit Sub
            TXTPTR.SetFocus
         Case vbKeyEscape
            TXTFREE.SetFocus
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
    TXTRATE.text = Format(TXTRATE.text, ".000")
    txtmrpbt.text = 100 * Val(TXTRATE.text) / 105 '(100 + Val(TxttaxMRP.Text))
End Sub

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If txtBillNo.text = "" Then Exit Sub
            If OptCredit.Value = True And IsNull(DataList2.SelectedItem) Then Exit Sub
            'If TXTINVOICE.Text = "" Then Exit Sub
            If Not IsDate(TXTINVDATE.text) Then Exit Sub
'            If OptCredit.value = True Then
'                Set rstTRXMAST = New ADODB.Recordset
'                rstTRXMAST.Open "Select * From ASTRXMAST WHERE TRX_TYPE='AP' AND PINV = '" & Trim(TXTINVOICE.Text) & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
'                If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
'                    MsgBox "You have already entered this Invoice number for " & Trim(DataList2.Text) & " as Computer Bill No. " & rstTRXMAST!VCH_NO, vbOKOnly, "EzBiz"
'                    rstTRXMAST.Close
'                    Set rstTRXMAST = Nothing
'                    TXTINVOICE.SetFocus
'                    Exit Sub
'                End If
'                rstTRXMAST.Close
'                Set rstTRXMAST = Nothing
'            End If
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
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
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
                TXTPRODUCT.text = grdsales.TextMatrix(Val(TXTSLNO.text), 2)
                TXTQTY.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 14))
                TXTUNIT.text = 1 'grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                Txtpack.text = 1 'grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                'TXTRATE.Text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)), 2), "0.000")
                TXTRATE.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6)), "0.000")
                TXTPTR.text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 9)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5)), 2), "0.000")
                txtprofit.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 7)), "0.00")
                Txtgrossamt.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 26)), "0.00")
                'TXTPTR.Text = Format((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))) * Val(Los_Pack.Text), "0.000")

                txtBatch.text = grdsales.TextMatrix(Val(TXTSLNO.text), 11)
                TXTEXPDATE.text = "  /  /    " 'IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = "", "  /  /    ", grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
                TXTEXPIRY.text = "  /  " 'IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = "", "  /  ", Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "mm/yy"))
                'LBLSUBTOTAL.Caption = Format(Val(TXTQTY.Text) * (Val(TXTPTR.Text) + Val(lbltaxamount.Caption)), ".000")
                If OptDiscAmt.Value = True Then
                    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text) + Val(lbltaxamount.Caption) - Val(txtPD.text), ".000")
                    LblGross.Caption = Format(Val(Txtgrossamt.text) - Val(txtPD.text), ".000")
                Else
                    LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100), ".000")
                    LblGross.Caption = Format(Val(Txtgrossamt.text) - (Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100)), ".000")
                End If
                TXTFREE.text = grdsales.TextMatrix(Val(TXTSLNO.text), 14)
                TxttaxMRP.text = grdsales.TextMatrix(Val(TXTSLNO.text), 10)
                txtmrpbt.text = 100 * Val(TXTRATE.text) / 105 '(100 + Val(TxttaxMRP.Text))
                txtPD.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 17))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) = "V" Then
                    OPTVAT.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) = "M" Then
                    OPTTaxMRP.Value = True
                Else
                    optnet.Value = True
                End If
                
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 27)) = "P" Then
                    optdiscper.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 27)) = "A" Then
                    OptDiscAmt.Value = True
                End If
                On Error Resume Next
                Los_Pack.text = grdsales.TextMatrix(Val(TXTSLNO.text), 28)
                CmbPack.text = grdsales.TextMatrix(Val(TXTSLNO.text), 29)
                LblGross.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 36))
                
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
            TXTPRODUCT.Enabled = False
            txtcategory.Enabled = True
            txtcategory.SetFocus
            'TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.text = Val(grdsales.rows)
                TXTPRODUCT.text = ""
                TXTITEMCODE.text = ""
                TXTQTY.text = ""
                Txtpack.text = 1 '""
                Los_Pack.text = ""
                CmbPack.ListIndex = -1
                TxtWarranty.text = ""
                CmbWrnty.ListIndex = -1
                TXTFREE.text = ""
                TxttaxMRP.text = ""
                
                
                
                txtPD.text = ""
                
                txtprofit.text = ""
                
                
                
                
               
                
                Txtgrossamt.text = ""
                
                
                
                TXTRATE.text = ""
                
                
                txtmrpbt.text = ""
                LBLSUBTOTAL.Caption = ""
                LblGross.Caption = ""
                lbltaxamount.Caption = ""
                TXTEXPDATE.text = "  /  /    "
                TXTEXPIRY.text = "  /  "
                txtBatch.text = ""
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
            txtPD.SetFocus
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
    Txtgrossamt.Tag = Val(Txtgrossamt.text)
    If Val(TxttaxMRP.text) = 0 Then
        
        TxttaxMRP.text = 0
        lbltaxamount.Caption = 0
        lbltaxamount.Caption = ""
        If optdiscper.Value = True Then
            LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag)) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100)
            LblGross.Caption = (Val(Txtgrossamt.Tag)) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100)
        Else
            LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) - Val(txtPD.text))
            LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(txtPD.text))
        End If

    Else
        If OPTTaxMRP.Value = True Then
            lbltaxamount.Caption = Val(txtmrpbt.text) * (Val(TXTQTY.text) + Val(TXTFREE.text)) * Val(TxttaxMRP.text) / 100
            If optdiscper.Value = True Then
                LBLSUBTOTAL.Caption = (Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption)
                LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - (Val(LBLSUBTOTAL.Caption) * Val(txtPD.text) / 100)
            Else
                LBLSUBTOTAL.Caption = (Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption)
                LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - Val(txtPD.text)
            End If
            LblGross.Caption = LBLSUBTOTAL.Caption
        ElseIf OPTVAT.Value = True Then
           If optdiscper.Value = True Then
                lbltaxamount.Caption = Round((Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100)) * Val(TxttaxMRP.text) / 100, 2)
                LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100)
                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100)
            Else
                lbltaxamount.Caption = Round((Val(Txtgrossamt.Tag) - Val(txtPD.text)) * Val(TxttaxMRP.text) / 100, 2)
                LBLSUBTOTAL.Caption = Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption) - Val(txtPD.text)
                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
            End If
        Else
            TxttaxMRP.text = 0
            lbltaxamount.Caption = 0
            lbltaxamount.Caption = ""
            If optdiscper.Value = True Then
                LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag)) - Val(txtPD.text)
            Else
                LBLSUBTOTAL.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
            End If
            LblGross.Caption = LBLSUBTOTAL.Caption
        End If
    End If
    
    LBLSUBTOTAL.Caption = Format(LBLSUBTOTAL.Caption, "0.00")
    LblGross.Caption = Format(LblGross.Caption, "0.00")
    TxttaxMRP.text = Format(TxttaxMRP.text, "0.00")
    lbltaxamount.Caption = Format(lbltaxamount.Caption, "0.00")
End Sub

Private Sub TXTUNIT_GotFocus()
    TXTUNIT.SelStart = 0
    TXTUNIT.SelLength = Len(TXTUNIT.text)
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.text) = 0 Then Exit Sub
            
            TXTUNIT.Enabled = False
            Txtpack.Enabled = True
            Txtpack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTQTY.text = ""
            TXTFREE.text = ""
            TxttaxMRP.text = ""
            txtprofit.text = ""
            Txtgrossamt.text = ""
            txtPD.text = ""
            txtBatch.text = ""
            TXTRATE.text = ""
            txtmrpbt.text = ""
            TXTPTR.text = ""
            Txtgrossamt.text = ""
            TXTEXPDATE.text = "  /  /    "
            TXTEXPIRY.text = "  /  "
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
    
    On Error GoTo ErrHand
    If (TXTDISCAMOUNT.text = "") Then
        DISC = 0
    Else
        DISC = TXTDISCAMOUNT.text
    End If
    If grdsales.rows = 1 Then
        TXTDISCAMOUNT.text = "0"
    ElseIf Val(TXTDISCAMOUNT.text) > Val(lbltotalwodiscount.Caption) Then
'        MsgBox "Discount Amount More than Bill Amount", , "PURCHASE..."
'        TXTDISCAMOUNT.SelStart = 0
'        TXTDISCAMOUNT.SelLength = Len(TXTDISCAMOUNT.Text)
'        TXTDISCAMOUNT.SetFocus
'        Exit Sub
    End If
    TXTDISCAMOUNT.text = Format(TXTDISCAMOUNT.text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
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
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    
    If OLD_BILL = False Then Call checklastbill
    db.Execute "delete From ASTRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'CR' AND INV_TRX_TYPE = 'AP'"
    db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'AP'"
    If grdsales.rows = 1 Then GoTo SKIP
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From ASTRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!TRX_TYPE = "AP"
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
        RSTTRXFILE!REF_NO = Trim(TxtPhone.text)
        RSTTRXFILE!SLSM_CODE = "CS"
        RSTTRXFILE!CHECK_FLAG = "N"
        If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
        RSTTRXFILE!CFORM_NO = ""
        RSTTRXFILE!CFORM_DATE = Date
        RSTTRXFILE!REMARKS = Trim(DataList2.text)
        RSTTRXFILE!DISC_PERS = Val(txtcramt.text)
        RSTTRXFILE!CST_PER = Val(TxtCST.text)
        RSTTRXFILE!INS_PER = Val(TxtInsurance.text)
        RSTTRXFILE!LETTER_NO = 0
        RSTTRXFILE!LETTER_DATE = Date
        RSTTRXFILE!INV_MSGS = ""
        If Not IsDate(TXTDATE.text) Then TXTDATE.text = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!CREATE_DATE = Format(TXTDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!PINV = Trim(TXTINVOICE.text)
        If optCash.Value = True Then
            RSTTRXFILE!Cash_Flag = "Y"
        Else
            RSTTRXFILE!Cash_Flag = "N"
        End If
        RSTTRXFILE.Update
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
    
    If lblcredit.Caption = "1" Then
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'CR' AND INV_TRX_TYPE = 'AP'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!TRX_TYPE = "CR"
            RSTITEMMAST!INV_TRX_TYPE = "AP"
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTITEMMAST!CR_NO = i
            RSTITEMMAST!INV_NO = Val(txtBillNo.text)
            RSTITEMMAST!RCPT_AMOUNT = 0
        End If
        RSTITEMMAST!INV_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTITEMMAST!INV_AMT = Val(LBLTOTAL.Caption)
        If lblcredit.Caption = "0" Then
            RSTITEMMAST!CHECK_FLAG = "Y"
            RSTITEMMAST!BAL_AMT = 0
        Else
            RSTITEMMAST!CHECK_FLAG = "N"
            RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption) - RSTITEMMAST!RCPT_AMOUNT
        End If
        RSTITEMMAST!PINV = Trim(TXTINVOICE.text)
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = DataList2.text
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    End If
        
'    For i = 1 To grdsales.Rows - 1
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
'            RSTLINK!UNIT = grdsales.TextMatrix(i, 4)
'            RSTLINK!Remarks = grdsales.TextMatrix(i, 4)
'            RSTLINK!ORD_QTY = 0
'            RSTLINK!CST = 0
'            RSTLINK!ACT_CODE = DataList2.BoundText
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
'
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from ASTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(Trim(TXTINVDATE.text), "dd/mm/yyyy")
        If OptCredit.Value = True Then
            RSTTRXFILE!VCH_DESC = "Received From " & Left(DataList2.text, 85)
            RSTTRXFILE!M_USER_ID = DataList2.BoundText
        Else
            RSTTRXFILE!VCH_DESC = "Received From " & Left(Trim(TXTDEALER.text), 85)
            RSTTRXFILE!M_USER_ID = "11111"
        End If
        RSTTRXFILE!PINV = Trim(TXTINVOICE.text)
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    If lblcredit.Caption = "0" Then
        i = 0
        Set rstMaxRec = New ADODB.Recordset
        rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
            i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
        End If
        rstMaxRec.Close
        Set rstMaxRec = Nothing
    
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'AP'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!REC_NO = i + 1
            RSTITEMMAST!INV_TYPE = "PY"
            RSTITEMMAST!INV_TRX_TYPE = "AP"
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTITEMMAST!INV_NO = Val(txtBillNo.text)
        End If
        If lblcredit.Caption = "0" Then
            RSTITEMMAST!TRX_TYPE = "DR"
        Else
            RSTITEMMAST!TRX_TYPE = "CR"
        End If
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = Trim(DataList2.text)
        RSTITEMMAST!AMOUNT = Val(LBLTOTAL.Caption)
        RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
        RSTITEMMAST!CHECK_FLAG = "P"
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    End If

SKIP:
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From ASTRXFILE WHERE TRX_TYPE = 'AP'", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    cmdRefresh.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.text = TXTLASTBILL.text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.text = "  /  /    "
    TXTINVOICE.text = ""
    TxtPhone.text = ""
    TXTREMARKS.text = ""
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TXTPRODUCT.text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.text = ""
    TxttaxMRP.text = ""
    
    
    
    txtPD.text = ""
    
    txtprofit.text = ""
    
    
    
    
   
    
    Txtgrossamt.text = ""
    
    
    txtBatch.text = ""
    TXTRATE.text = ""
    txtmrpbt.text = ""
    TXTPTR.text = ""
    Txtgrossamt.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    txtaddlamt.text = ""
    txtcramt.text = ""
    TxtInsurance.text = ""
    TxtCST.text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    TXTDISCAMOUNT.text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.text = ""
    lbldealer.Caption = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    
    txtBillNo.SetFocus
    M_ADD = False
    OLD_BILL = False
    LBLmonth.Caption = "0.00"
    Chkcancel.Value = 0
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "EzBiz"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If err.Number = 7 Then
        MsgBox "Select Supplier from the list", vbOKOnly, "EzBiz"
    Else
        MsgBox err.Description
    End If
End Sub


Private Sub txtaddlamt_GotFocus()
    txtaddlamt.SelStart = 0
    txtaddlamt.SelLength = Len(txtaddlamt.text)
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
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
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
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    txtcramt.SetFocus
End Sub

Private Sub OPTTaxMRP_GotFocus()
    'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100
    'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
    lbltaxamount.Caption = ((Val(TXTRATE.text) * (Val(TXTQTY.text) + Val(TXTFREE.text)) * 55 / 100)) * Val(TxttaxMRP.text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)) + Val(lbltaxamount.Caption), ".000")
    LblGross.Caption = Format((Val(TXTQTY.text) * Val(TXTPTR.text)), ".000")
            
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
    If optdiscper.Value = True Then
        lbltaxamount.Caption = Round((Val(Txtgrossamt.text) - (Val(Txtgrossamt.text) * Val(txtPD.text) / 100)) * Val(TxttaxMRP.text) / 100, 2)
        LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100), ".000")
        LblGross.Caption = Format(Val(Txtgrossamt.text) - Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100), ".000")
    Else
        lbltaxamount.Caption = Round((Val(Txtgrossamt.text) - Val(txtPD.text)) * Val(TxttaxMRP.text) / 100, 2)
        LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text) + Val(lbltaxamount.Caption) - Val(txtPD.text), ".000")
        LblGross.Caption = Format(Val(Txtgrossamt.text) - Val(txtPD.text), ".000")
    End If
End Sub

Private Sub OPTNET_GotFocus()
    lbltaxamount.Caption = ""
    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text), ".000")
    LblGross.Caption = Format(Val(Txtgrossamt.text), ".000")
End Sub

Private Sub txtPD_GotFocus()
    txtPD.SelStart = 0
    txtPD.SelLength = Len(txtPD.text)
End Sub

Private Sub txtPD_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
'            txtPD.Enabled = False
'            cmdadd.Enabled = True
'            cmdadd.SetFocus
'            Exit Sub
            cmdadd.Enabled = True
            cmdadd.SetFocus
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


Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If optCash.Value = True Then Exit Sub
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
    MsgBox err.Description
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If optCash.Value = True Then
                TXTINVOICE.SetFocus
            Else
                If DataList2.VisibleCount = 0 Then Exit Sub
                DataList2.SetFocus
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

Private Sub DataList2_Click()
    If optCash.Value = True Then Exit Sub
    TXTDEALER.text = DataList2.text
    lbldealer.Caption = TXTDEALER.text
    On Error GoTo ErrHand
    Dim rstCustomer As ADODB.Recordset
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from ACTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        TxtPhone.text = IIf(IsNull(rstCustomer!TELNO), "", rstCustomer!TELNO)
    Else
        TxtPhone.text = ""
    End If
    Call Monthly_purchase
    Exit Sub
ErrHand:
    MsgBox err.Description
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
            TXTINVOICE.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub Txtgrossamt_GotFocus()
    Txtgrossamt.SelStart = 0
    Txtgrossamt.SelLength = Len(Txtgrossamt.text)
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
    If Val(Txtgrossamt.text) <> 0 Then
        Txtgrossamt.text = Format(Txtgrossamt.text, ".000")
        If Val(TXTQTY.text) <> 0 Then
            TXTPTR.text = Format(Round(Val(Txtgrossamt.text) / Val(TXTQTY.text), 1), "0.00")
        ElseIf Val(TXTPTR.text) <> 0 Then
            TXTQTY.text = Format(Round(Val(Txtgrossamt.text) / Val(TXTPTR.text), 1), "0.00")
        End If
    End If
    Call TxttaxMRP_LostFocus
End Sub

Function FILL_PREVIIOUSRATE()
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, ITEM_COST, PACK_TYPE, VCH_NO, VCH_DATE, VCH_DESC  From ASTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, ITEM_COST, PACK_TYPE, VCH_NO, VCH_DATE, VCH_DESC  From ASTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
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
        GRDPRERATE.Columns(0).Caption = "TYPE"
        GRDPRERATE.Columns(1).Caption = "ITEM CODE"
        GRDPRERATE.Columns(2).Caption = "ITEM NAME"
        GRDPRERATE.Columns(3).Caption = "COST"
        GRDPRERATE.Columns(4).Caption = "PACK"
        GRDPRERATE.Columns(5).Caption = "BILL NO."
        GRDPRERATE.Columns(6).Caption = "BILL DATE"
        GRDPRERATE.Columns(7).Caption = "RECEIVED FROM"
    
        GRDPRERATE.Columns(0).Visible = False
        GRDPRERATE.Columns(1).Visible = False
        GRDPRERATE.Columns(2).Width = 0
        GRDPRERATE.Columns(3).Width = 600
        GRDPRERATE.Columns(4).Width = 600
        GRDPRERATE.Columns(5).Width = 1200
        GRDPRERATE.Columns(6).Width = 1200
        GRDPRERATE.Columns(7).Width = 1100
        
        
        'GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(2).text
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
    Los_Pack.SelLength = Len(Los_Pack.text)
    FRMEGRDTMP.Visible = False
    CmbPack.Enabled = True
    TXTQTY.Enabled = True
    TXTFREE.Enabled = True
    TXTRATE.Enabled = True
    TXTPTR.Enabled = True
    TxttaxMRP.Enabled = True
    txtPD.Enabled = True
    
    
   
   
    
    
   
  
   
   
    
    cmdadd.Enabled = True
    Txtgrossamt.Enabled = True
End Sub

Private Sub Los_Pack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
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
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
        
            If Trim(TXTITEMCODE.text) = "" Then
                TXTPRODUCT.Enabled = True
                TXTPRODUCT.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False
            
            Set grdtmp.DataSource = Nothing
            If PHYCODE_FLAG = True Then
                PHY_CODE.Open "Select * From ASTMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            Else
                PHY_CODE.Close
                PHY_CODE.Open "Select * From ASTMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' and ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            End If
            
            Set grdtmp.DataSource = PHY_CODE
            
            If PHY_CODE.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "EzBiz"
                Exit Sub
            End If
            
            If PHY_CODE.RecordCount = 1 Then
                TXTITEMCODE.text = grdtmp.Columns(0)
                TXTPRODUCT.text = grdtmp.Columns(1)
'                On Error Resume Next
'                Set Image1.DataSource = PHY
'                If IsNull(PHY!PHOTO) Then
'                    Frame6.Visible = False
'                    Set Image1.DataSource = Nothing
'                    bytData = ""
'                Else
'                    If Err.Number = 545 Then
'                        Frame6.Visible = False
'                        Set Image1.DataSource = Nothing
'                        bytData = ""
'                    Else
'                        Frame6.Visible = True
'                        Set Image1.DataSource = PHY 'setting image1’s datasource
'                        Image1.DataField = "PHOTO"
'                        bytData = PHY!PHOTO
'                    End If
'                End If
                On Error GoTo ErrHand
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                        Exit For
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From ASTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    'RSTRXFILE.MoveLast
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.text = 1
                    txtBatch.text = IIf(IsNull(RSTRXFILE!REF_NO), "", RSTRXFILE!REF_NO)
                    If IsNull(RSTRXFILE!PTR) Then
                        TXTPTR.text = ""
                    Else
                        TXTPTR.text = Format(Round(Val(RSTRXFILE!PTR) * Val(Los_Pack.text), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.text = ""
                    Else
                        TxttaxMRP.text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                    End If
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    On Error GoTo ErrHand
                    
                    'TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                    If RSTRXFILE!CHECK_FLAG = "M" Then
                        OPTTaxMRP.Value = True
                    ElseIf RSTRXFILE!CHECK_FLAG = "V" Then
                        OPTVAT.Value = True
                    Else
                        optnet.Value = True
                    End If
                Else
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = 1
                    Los_Pack.text = 1
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.text = ""
                    TXTEXPIRY.text = "  /  "
                    TXTRATE.text = ""
                    txtmrpbt.text = ""
                    TXTPTR.text = ""
                    
                   
                    
                    
                    
                    txtprofit.text = ""
                    TxttaxMRP.text = "5"
                    
                    
                    
                    Los_Pack.text = "1"
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ErrHand
                    OPTVAT.Value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                If PHY_CODE.RecordCount = 1 Then
                    TXTITEMCODE.Enabled = False
                    TXTPRODUCT.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
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
ErrHand:
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

Private Sub txtcst_GotFocus()
    TxtCST.SelStart = 0
    TxtCST.SelLength = Len(TxtCST.text)
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
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub TxtCST_LostFocus()
    Dim DISC As Currency
    
    On Error GoTo ErrHand
    If (TxtCST.text = "") Then
        DISC = 0
    Else
        DISC = TxtCST.text
    End If
    If grdsales.rows = 1 Then
        TxtCST.text = "0"
        Exit Sub
    End If
    TxtCST.text = Format(TxtCST.text, ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtCST.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtCST.SetFocus
End Sub

Private Sub TxtInsurance_GotFocus()
    TxtInsurance.SelStart = 0
    TxtInsurance.SelLength = Len(TxtInsurance.text)
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
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub TxtInsurance_LostFocus()
    Dim DISC As Currency
    
    On Error GoTo ErrHand
    If (TxtInsurance.text = "") Then
        DISC = 0
    Else
        DISC = TxtInsurance.text
    End If
    If grdsales.rows = 1 Then
        TxtInsurance.text = "0"
        Exit Sub
    End If
    TxtInsurance.text = Format(TxtInsurance.text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + (Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(txtcst.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtInsurance.Text)), 0), ".00")
    Exit Sub
ErrHand:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtInsurance.SetFocus
End Sub

Private Function checklastbill()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From ASTRXMAST WHERE TRX_TYPE = 'AP'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
Exit Function
ErrHand:
    MsgBox err.Description
End Function

Private Function Monthly_purchase()
    Dim rstTRANX As ADODB.Recordset
    Dim TOT_SALE As Long
    Dim FROM_DATE As Date
    
    FROM_DATE = "01/" & Month(Date) & "/" & Year(Date)
    On Error GoTo ErrHand
    TOT_SALE = 0
    LBLmonth.Caption = "0.00"
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From ASTRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='AP' AND VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(Date, "yyyy/mm/dd") & "' AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        TOT_SALE = TOT_SALE + (rstTRANX!VCH_AMOUNT + IIf(IsNull(rstTRANX!ADD_AMOUNT), 0, rstTRANX!ADD_AMOUNT) - IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT))
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    LBLmonth.Caption = Format(TOT_SALE, "0.00")
    'LBLRETURNED.Caption = Format(TOT_RET, "0.00")
    
    Exit Function
ErrHand:
    MsgBox err.Description
End Function

Private Sub optCash_Click()
    DataList2.Visible = False
End Sub

Private Sub OptCredit_Click()
    DataList2.Visible = True
End Sub

Private Sub CmdPrint_Click()
    Dim i As Long
    
    On Error GoTo ErrHand
     
    db.Execute "delete from TEMPTRXFILE"
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "AP"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        'RSTTRXFILE!Category = "" 'grdsales.TextMatrix(i, 25)
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        
        
        RSTTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(i, 13))
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "dd/mm/yyyy")
        RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
        RSTTRXFILE!ITEM_COST = Format(Round(Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 5)), 4), "0.0000") 'Val(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 17))
        If Trim(grdsales.TextMatrix(i, 27)) = "P" Then
            RSTTRXFILE!Category = "P"
        ElseIf Trim(grdsales.TextMatrix(i, 27)) = "A" Then
            RSTTRXFILE!Category = "A"
        End If
                
        
        'RSTTRXFILE!P_DISC = Val(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!PTR = Format(Round(Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 5)), 4), "0.0000")
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 18))
        RSTTRXFILE!P_RETAILWOTAX = Round(Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 5)), 4)
        'RSTTRXFILE!P_WS = Val(grdsales.TextMatrix(i, 19))
        'RSTTRXFILE!P_CRTN = Val(grdsales.TextMatrix(i, 20))
        'RSTTRXFILE!CRTN_PACK = Val(grdsales.TextMatrix(i, 24))
        'RSTTRXFILE!P_VAN = Val(grdsales.TextMatrix(i, 25))
        'RSTTRXFILE!GROSS_AMT = Val(grdsales.TextMatrix(i, 26))
        RSTTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 28))
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 29)
        RSTTRXFILE!WARRANTY = Val(grdsales.TextMatrix(i, 30))
        RSTTRXFILE!WARRANTY_TYPE = Trim(grdsales.TextMatrix(i, 31))
        RSTTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(I, 4))
        If MDIMAIN.StatusBar.Panels(14).text = "Y" Then
            RSTTRXFILE!VCH_DESC = "Y"
        Else
            RSTTRXFILE!VCH_DESC = "N"
        End If
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 11))
        'RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!SCHEME = Val(grdsales.TextMatrix(i, 14))
        If IsDate(grdsales.TextMatrix(i, 12)) Then
            RSTTRXFILE!EXP_DATE = Format(grdsales.TextMatrix(i, 12), "dd/mm/yyyy")
        End If
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(i, 15))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
        RSTTRXFILE.Update
    Next i
    
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompAddress4, CompAddress5, CompTin, CompCST, DL, ML, DL1, DL2, INV_TERMS, BANK_DET, PAN_NO As String
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
            
    ReportNameVar = Rptpath & "rptLPA"
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
        Report.OpenSubreport("RPTBILL" & i & ".rpt").RecordSelectionFormula = "({TRXFILE.VCH_NO}= " & Val(txtBillNo.text) & ")"
        Report.OpenSubreport("RPTBILL" & i & ".rpt").DiscardSavedData
        Report.OpenSubreport("RPTBILL" & i & ".rpt").VerifyOnEveryPrint = True
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
            If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.text = "'" & INV_TERMS & "'"
            If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.text = "'" & BANK_DET & "'"
            If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.text = "'" & PAN_NO & "'"
            If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.text = "'" & Trim(TXTDEALER.text) & "'"
            If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.text = "'" & Trim(TXTDEALER.text) & "'"
'            If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.Text = "'" & Trim(lbladdress.Caption) & "'"
            If CRXFormulaField.Name = "{DLNO2}" Then CRXFormulaField.text = "'" & DL1 & "'"
            If CRXFormulaField.Name = "{DLNO}" Then CRXFormulaField.text = "'" & DL2 & "'"
            'If CRXFormulaField.Name = "{@Area}" Then CRXFormulaField.Text = "'" & Trim(TXTAREA.Text) & "'"
            'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Tag), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
            If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
    '        If Tax_Print = False Then
    '            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
    '        End If
            'If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
            If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.text = "'" & TxtPhone.text & "'"
            If CRXFormulaField.Name = "{@VCH_NO}" Then
                Me.Tag = Format(Trim(txtBillNo.text), bill_for)
                CRXFormulaField.text = "'" & Me.Tag & "' "
            End If
            'If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
            'If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.Text = "'" & Trim(TxtOrder.Text) & "'"
    '            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLTOTAL.Caption), 0), "0.00") & "'"
        Next
    Next i
        
    'Preview
    frmreport.Caption = "BILL"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub


