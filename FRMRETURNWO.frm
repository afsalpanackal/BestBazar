VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMSALERETURNWO 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALESRETURN"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13560
   ControlBox      =   0   'False
   Icon            =   "FRMRETURNWO.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   13560
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   1440
      TabIndex        =   47
      Top             =   3120
      Visible         =   0   'False
      Width           =   6030
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   2835
         Left            =   75
         TabIndex        =   48
         Top             =   120
         Width           =   5880
         _ExtentX        =   10372
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
   Begin VB.Frame FRMEGRDBILL 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   840
      TabIndex        =   70
      Top             =   3120
      Visible         =   0   'False
      Width           =   7395
      Begin MSDataGridLib.DataGrid GRDPOPUPBILL 
         Height          =   2535
         Left            =   90
         TabIndex        =   71
         Top             =   360
         Width           =   7200
         _ExtentX        =   12700
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
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   "BILL DETAILS FOR THE ITEM"
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
         TabIndex        =   73
         Top             =   105
         Visible         =   0   'False
         Width           =   2685
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
         Left            =   2775
         TabIndex        =   72
         Top             =   105
         Visible         =   0   'False
         Width           =   4515
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   -105
      TabIndex        =   17
      Top             =   90
      Width           =   11160
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
         TabIndex        =   51
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00FF8080&
         Height          =   1725
         Left            =   210
         TabIndex        =   18
         Top             =   15
         Width           =   10845
         Begin MSDataListLib.DataCombo cmbinv 
            Height          =   330
            Left            =   1545
            TabIndex        =   68
            Top             =   1350
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Left            =   1545
            TabIndex        =   0
            Top             =   600
            Width           =   3735
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
            Left            =   1560
            TabIndex        =   15
            Top             =   225
            Width           =   885
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   345
            Left            =   3840
            TabIndex        =   52
            Top             =   225
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
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1545
            TabIndex        =   65
            Top             =   945
            Width           =   3720
            _ExtentX        =   6562
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
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Index           =   22
            Left            =   5325
            TabIndex        =   61
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "TIN"
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
            Left            =   8295
            TabIndex        =   60
            Top             =   1365
            Width           =   360
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DL No."
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
            Index           =   3
            Left            =   5370
            TabIndex        =   59
            Top             =   1365
            Width           =   645
         End
         Begin VB.Label lbltin 
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   8640
            TabIndex        =   58
            Top             =   1335
            Width           =   2145
         End
         Begin VB.Label lbldlno 
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6150
            TabIndex        =   57
            Top             =   1335
            Width           =   2130
         End
         Begin VB.Label lbladdress 
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
            ForeColor       =   &H80000008&
            Height          =   660
            Left            =   6150
            TabIndex        =   56
            Top             =   630
            Width           =   4635
         End
         Begin VB.Label lblcust 
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
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Index           =   2
            Left            =   90
            TabIndex        =   54
            Top             =   675
            Width           =   1230
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
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Index           =   8
            Left            =   2490
            TabIndex        =   53
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label LblInvoice 
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
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Index           =   0
            Left            =   105
            TabIndex        =   22
            Top             =   240
            Width           =   1230
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
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Index           =   1
            Left            =   5505
            TabIndex        =   21
            Top             =   255
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
            Left            =   6120
            TabIndex        =   20
            Top             =   225
            Width           =   1215
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
            Left            =   7335
            TabIndex        =   19
            Top             =   225
            Width           =   1110
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Height          =   6465
         Left            =   210
         TabIndex        =   23
         Top             =   1695
         Width           =   10890
         Begin VB.Frame Frame3 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   6240
            Left            =   8985
            TabIndex        =   24
            Top             =   165
            Width           =   1815
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
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   23
               Left            =   150
               TabIndex        =   62
               Top             =   1290
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
               Height          =   570
               Left            =   180
               TabIndex        =   55
               Top             =   1590
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
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   6
               Left            =   165
               TabIndex        =   26
               Top             =   255
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
               Height          =   570
               Left            =   195
               TabIndex        =   25
               Top             =   645
               Width           =   1440
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   5730
            Left            =   90
            TabIndex        =   16
            Top             =   270
            Width           =   8730
            _ExtentX        =   15399
            _ExtentY        =   10107
            _Version        =   393216
            Rows            =   1
            Cols            =   19
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
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   465
         Left            =   11100
         TabIndex        =   46
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
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00FF8080&
         Height          =   1365
         Left            =   210
         TabIndex        =   27
         Top             =   8070
         Width           =   10890
         Begin VB.TextBox TxtPack 
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
            Left            =   3975
            MaxLength       =   6
            TabIndex        =   74
            Top             =   450
            Width           =   525
         End
         Begin VB.CommandButton cmdview 
            Caption         =   "&VIEW"
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
            Left            =   7755
            TabIndex        =   69
            Top             =   810
            Width           =   1125
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
            Height          =   465
            Left            =   10335
            TabIndex        =   63
            Top             =   810
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
            Left            =   5295
            MaxLength       =   6
            TabIndex        =   49
            Top             =   450
            Width           =   765
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
            Height          =   465
            Left            =   1815
            TabIndex        =   9
            Top             =   780
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
            Left            =   45
            TabIndex        =   1
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
            Left            =   630
            TabIndex        =   2
            Top             =   450
            Width           =   3330
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
            Left            =   4515
            MaxLength       =   7
            TabIndex        =   3
            Top             =   450
            Width           =   765
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
            Left            =   6075
            MaxLength       =   6
            TabIndex        =   4
            Top             =   450
            Width           =   780
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
            Left            =   6870
            MaxLength       =   4
            TabIndex        =   5
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
            Left            =   9255
            MaxLength       =   4
            TabIndex        =   7
            Top             =   465
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
            Left            =   5400
            TabIndex        =   12
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
            Left            =   8955
            TabIndex        =   14
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
            Left            =   4215
            TabIndex        =   11
            Top             =   795
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
            Height          =   465
            Left            =   3000
            TabIndex        =   10
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
            Left            =   2070
            TabIndex        =   32
            Top             =   1245
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
            Left            =   7500
            MaxLength       =   15
            TabIndex        =   6
            Top             =   465
            Width           =   1740
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
            TabIndex        =   31
            Top             =   1260
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
            TabIndex        =   30
            Top             =   1260
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
            TabIndex        =   29
            Top             =   1290
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
            TabIndex        =   28
            Top             =   1335
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
            Height          =   465
            Left            =   6585
            TabIndex        =   13
            Top             =   810
            Width           =   1125
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
            Index           =   0
            Left            =   3975
            TabIndex        =   75
            Top             =   225
            Width           =   525
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
            Left            =   5295
            TabIndex        =   50
            Top             =   225
            Width           =   750
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
            Left            =   45
            TabIndex        =   45
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
            Left            =   630
            TabIndex        =   44
            Top             =   225
            Width           =   3330
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
            Left            =   4515
            TabIndex        =   43
            Top             =   225
            Width           =   765
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
            Left            =   6090
            TabIndex        =   42
            Top             =   225
            Width           =   750
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
            Left            =   6870
            TabIndex        =   41
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
            Left            =   9255
            TabIndex        =   40
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
            Left            =   9930
            TabIndex        =   39
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
            Left            =   930
            TabIndex        =   38
            Top             =   1260
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Batch No."
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
            Left            =   7515
            TabIndex        =   37
            Top             =   240
            Width           =   1710
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
            Left            =   9930
            TabIndex        =   8
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
            Left            =   2895
            TabIndex        =   36
            Top             =   1275
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
            TabIndex        =   35
            Top             =   1305
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
            TabIndex        =   34
            Top             =   1350
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
            TabIndex        =   33
            Top             =   1275
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
   Begin MSDataListLib.DataCombo CMBDISTI 
      Height          =   1020
      Left            =   9825
      TabIndex        =   64
      Top             =   1275
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   1773
      _Version        =   393216
      Appearance      =   0
      Style           =   1
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
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11655
      TabIndex        =   67
      Top             =   2610
      Width           =   495
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11445
      TabIndex        =   66
      Top             =   3255
      Width           =   1620
   End
End
Attribute VB_Name = "FRMSALERETURNWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim TMPREC As New ADODB.Recordset
Dim TMPFLAG As Boolean
Dim ACT_REC As New ADODB.Recordset

Dim ACT_FLAG As Boolean
Dim PHY_BILL As New ADODB.Recordset
Dim BILL_FLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean
Dim INV_FLAG As Boolean
Dim INV_REC As New ADODB.Recordset

Dim CLOSEALL As Integer
Dim M_STOCK As Double
Dim EDIT_BILL As Boolean
Dim M_EDIT As Boolean
Dim B_FLAG As Boolean
Dim CN_FLAG As Boolean
Dim M_DELETE As Boolean

Private Sub cmbinv_Change()
    txtBillNo.Text = cmbinv.Text
    Call VIEWGRID
    FRMEMASTER.Enabled = True
End Sub

Private Sub CMDADD_Click()
    Dim rststock As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHand
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Val(TxtPack.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Format(Val(TxtMRP.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTRATE.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Val(TXTDISC.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(TXTTAX.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = ""
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTTRXTYPE.Text)
    
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMASTWO WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db2, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = IIf(IsNull(RSTTRXFILE!MANUFACTURER), "", Trim(RSTTRXFILE!MANUFACTURER))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "N"
    grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = Val(TXTQTY.Tag)
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 12) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
          
    CN_FLAG = False
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILEWO WHERE TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & "", db2, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            RSTTRXFILE.Update
        Else
            CN_FLAG = True
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    db2.Execute "delete * From RTRXFILEWO WHERE TRX_TYPE = 'CN' AND VCH_NO = " & Val(txtBillNo.Text) & " "
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILEWO", db2, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "CN"
        RSTTRXFILE!VCH_NO = Trim(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(Trim(LBLDATE.Caption), "dd/mm/yyyy")
        RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 0))
        RSTTRXFILE!CATEGORY = "MEDICINE"
        RSTTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1))
        RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2))
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
        RSTTRXFILE!UNIT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
        RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11)) / Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
        RSTTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
        RSTTRXFILE!VCH_DESC = "C/Note from   " & DataList2.Text
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 9))
        'RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        If CN_FLAG = True Then
            RSTTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) '* Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
        Else
            RSTTRXFILE!BAL_QTY = 0
        End If
        RSTTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Date
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "311000"
        RSTTRXFILE!CHECK_FLAG = ""
        RSTTRXFILE!PINV = "CN# " & Trim(txtBillNo.Text)
        RSTTRXFILE.Update
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
SKIP:
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 11)), 2)
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption), 2)
    
    'Call STOCKADJUST
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TxtPack.Text = ""
    
    TXTQTY.Text = ""
    TxtMRP.Text = ""
    TXTRATE.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtBatch.Text = ""
    
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    cmddelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    M_EDIT = False
    TXTSLNO.SetFocus
    'grdsales.TopRow = grdsales.Rows - 1
Exit Sub
ErrHand:
    MsgBox Err.Description
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
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 12) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILEWO WHERE TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & "", db2, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
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
        grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(i + 1, 6)
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
    TxtPack.Text = ""
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    TxtMRP.Text = ""
    TXTTAX.Text = ""
    
    TXTDISC.Text = ""
    txtBatch.Text = ""
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    cmddelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    If grdsales.Rows = 1 Then
'        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        cmdRefresh.Enabled = True
        cmdRefresh.SetFocus
    End If
    M_DELETE = True
End Sub

Private Sub CMDEXIT_Click()
    If CMDEXIT.Caption = "E&XIT" Then
        CLOSEALL = 0
        Unload Me
    Else
        FRMEMASTER.Enabled = True
        txtBillNo.Enabled = True
        txtBillNo.SetFocus
        CMDEXIT.Caption = "E&XIT"
        TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
        DataList2.Text = ""
        lbladdress.Caption = ""
        lbldlno.Caption = ""
        lbltin.Caption = ""
        lblnetamount.Caption = ""
        LBLDATE.Caption = Date
        LBLTIME.Caption = Time
        LBLTOTAL.Caption = ""
        grdsales.Rows = 1
        TXTSLNO.Text = 1
        M_EDIT = False
        cmdRefresh.Enabled = False
        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
        CMDEXIT.Enabled = True
        TXTQTY.Tag = ""
        cmdview.Enabled = True
        LblInvoice(0).Top = 240
        TXTDEALER.Top = 600
        DataList2.Top = 945
        lblcust(2).Top = 675
        cmbinv.Visible = False
        TXTDEALER.Text = ""
        TXTDEALER.Enabled = False
        DataList2.Enabled = False
        TXTINVDATE.Enabled = False
    End If
End Sub

Private Sub CMDMODIFY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    On Error GoTo ErrHand
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 12) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILEWO WHERE TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "' AND VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & "", db2, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    TXTQTY.Tag = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
    CMDMODIFY.Enabled = False
    cmddelete.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = True
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TxtPack.Text = ""
            LBLSUBTOTAL.Caption = ""
            
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            cmddelete.Enabled = False
    End Select
End Sub

Private Sub cmdPrint_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Integer
    If grdsales.Rows = 1 Then Exit Sub
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
        DataList2.SetFocus
        Exit Sub
    End If
    
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
    
    Exit Sub
    db2.Execute "delete * From TRXFILEWO"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFILEWO", db2, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        
        Set TRXMAST = New ADODB.Recordset
        TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMASTWO WHERE ITEMMAST.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 12)) & "'", db2, adOpenStatic, adLockReadOnly
        If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
            RSTTRXFILE!MFGR = IIf(IsNull(TRXMAST!MANUFACTURER), "", TRXMAST!MANUFACTURER)
        End If
        TRXMAST.Close
        Set TRXMAST = Nothing
        
        RSTTRXFILE!TRX_TYPE = "DN"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!CATEGORY = "MEDICINE"
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 12)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = grdsales.TextMatrix(i, 3)
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = grdsales.TextMatrix(i, 5)
        RSTTRXFILE!PTR = grdsales.TextMatrix(i, 6)
        RSTTRXFILE!SALES_PRICE = grdsales.TextMatrix(i, 6)
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 8)
        RSTTRXFILE!UNIT = 1
        RSTTRXFILE!VCH_DESC = "Issued to     " & DataList2.Text
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 11)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
GOSKIP:
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    rptPRINT.ReportFileName = App.Path & "\RPTBILL.RPT"
    rptprintsmall.ReportFileName = App.Path & "\RPTBILLsmall.RPT"
    
    cmdRefresh.SetFocus
    
    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption), 2)) - Val(Round(Val(LBLTOTAL.Caption), 0)), 2)
    rptPRINT.Formulas(2) = "Company = '" & DataList2.Text & "'"
    rptPRINT.Formulas(11) = "Total = '" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
    'rptprint.Formulas(3) = "Disc = '" & Val(TXTAMOUNT.Text) & "'"
    'rptprint.Formulas(6) = "net1 = '" & Val(lblnetamount.Caption) & "'"
    'rptprint.Formulas(6) = "net1 = '" & Val(LBLTOTAL.Caption) & "'"
    rptPRINT.Formulas(8) = "Round1 = '" & Format(Val(lblnetamount.Tag), "0.00") & "'"
    rptPRINT.Formulas(9) = "Round2 = '" & Format(Round(Val(LBLTOTAL.Caption), 0), "0.00") & "'"
    rptPRINT.Formulas(1) = "Address = '" & lbladdress.Caption & "'"
    rptPRINT.Formulas(4) = "DLNO = '" & lbldlno.Caption & "'"
    rptPRINT.Formulas(10) = "TIN = '" & lbltin.Caption & "'"
    
    rptprintsmall.Formulas(2) = "Company = '" & DataList2.Text & "'"
    rptprintsmall.Formulas(11) = "Total = '" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
    'rptprint.Formulas(3) = "Disc = '" & Val(TXTAMOUNT.Text) & "'"
    'rptprint.Formulas(6) = "net1 = '" & Val(lblnetamount.Caption) & "'"
    'rptprintsmall.Formulas(6) = "net1 = '" & Val(LBLTOTAL.Caption) & "'"
    rptprintsmall.Formulas(8) = "Round1 = '" & Format(Val(lblnetamount.Tag), "0.00") & "'"
    rptprintsmall.Formulas(9) = "Round2 = '" & Format(Round(Val(LBLTOTAL.Caption), 0), "0.00") & "'"
    rptprintsmall.Formulas(1) = "Address = '" & lbladdress.Caption & "'"
    rptprintsmall.Formulas(4) = "DLNO = '" & lbldlno.Caption & "'"
    rptprintsmall.Formulas(10) = "TIN = '" & lbltin.Caption & "'"
    
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTTAX.Enabled = False
    
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    
    If grdsales.Rows < 10 Then
        rptprintsmall.Action = 1
    Else
        rptPRINT.Action = 1
    End If
    
End Sub

Private Sub cmdRefresh_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Double
    
    Dim DAY_DATE As String
    Dim MONTH_DATE As String
    Dim YEAR_DATE As String
    Dim E_DATE As Date
    
    'If grdsales.Rows = 1 Then GoTo SKIP
    
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
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
        DataList2.SetFocus
        Exit Sub
    End If
    
    i = 0
    On Error GoTo ErrHand
    
    db2.Execute "delete * From SALERETURNWO WHERE VCH_NO = " & Val(txtBillNo.Text) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From SALERETURNWO", db2, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!TRX_TYPE = "CN"
        RSTTRXFILE!CATEGORY = "MEDICINE"
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = grdsales.TextMatrix(i, 3)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = grdsales.TextMatrix(i, 5)
        RSTTRXFILE!PTR = grdsales.TextMatrix(i, 6)
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 11)) / Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 8)
        RSTTRXFILE!VCH_DESC = "Received from " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 11)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!CHECK_FLAG = "N"

        RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(i, 13) = "", 0, grdsales.TextMatrix(i, 13))
        RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(i, 14) = "", 0, grdsales.TextMatrix(i, 14))
        RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(i, 15) = "", "CN", grdsales.TextMatrix(i, 15))
        RSTTRXFILE!ISSUEQTY = Val(grdsales.TextMatrix(i, 18))
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    txtBillNo.Text = 1
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select MAX(Val(VCH_NO)) From SALERETURNWO", db2, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        txtBillNo.Text = IIf(IsNull(RSTTRXFILE.Fields(0)), 1, RSTTRXFILE.Fields(0) + 1)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
SKIP:
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    lbladdress.Caption = ""
    lbldlno.Caption = ""
    lbltin.Caption = ""
    lblnetamount.Caption = ""
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    LBLTOTAL.Caption = ""
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    M_EDIT = False
    cmdRefresh.Enabled = False
    CMDEXIT.Enabled = True
    CMDPRINT.Enabled = False
    CMDEXIT.Enabled = True
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    TXTQTY.Tag = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    cmdview.Enabled = True
    M_DELETE = False
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub cmdview_Click()
    LblInvoice(0).Top = 1400
    TXTDEALER.Top = 225
    DataList2.Top = 570
    lblcust(2).Top = 240
    cmbinv.Visible = True
    TXTDEALER.Text = ""
    CMDEXIT.Caption = "CANCEL"
    TXTDEALER.Enabled = True
    DataList2.Enabled = True
    cmdRefresh.Enabled = False
    TXTDEALER.SetFocus
    cmdview.Enabled = False
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(VCH_NO)) From SALERETURNWO", db2, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    ACT_FLAG = True
    Call FILLCOMBO
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2000
    grdsales.ColWidth(3) = 500
    grdsales.ColWidth(4) = 500
    grdsales.ColWidth(5) = 700
    grdsales.ColWidth(6) = 600
    grdsales.ColWidth(7) = 500
    grdsales.ColWidth(8) = 600
    grdsales.ColWidth(9) = 800
    grdsales.ColWidth(10) = 0
    grdsales.ColWidth(11) = 1000
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "UNIT"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "RATE"
    grdsales.TextArray(7) = "DISC %"
    grdsales.TextArray(8) = "TAX %"
    grdsales.TextArray(9) = "Serial No"
    grdsales.TextArray(10) = "EXPIRY"
    grdsales.TextArray(11) = "SUB TOTAL"
    grdsales.TextArray(12) = "ITEM CODE"
    grdsales.TextArray(13) = "Vch No"
    grdsales.TextArray(14) = "Line No"
    grdsales.TextArray(15) = "Trx Type"
    grdsales.TextArray(16) = "MFGR"
    grdsales.TextArray(17) = "FLAG"
    grdsales.TextArray(18) = "ISSUE QTY"
    
    grdsales.ColWidth(12) = 0
    grdsales.ColWidth(13) = 0
    grdsales.ColWidth(14) = 0
    grdsales.ColWidth(15) = 0
    
    grdsales.ColWidth(17) = 0
    grdsales.ColWidth(18) = 0
    LBLTOTAL.Caption = 0
    
    PHYFLAG = True
    TMPFLAG = True
    BILL_FLAG = True
    ITEM_FLAG = True
    Me.Top = 0
    INV_FLAG = True
    M_DELETE = False
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTTAX.Enabled = False
    
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    cmddelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    TXTSLNO.Text = 1
    TXTSLNO.Enabled = True
    txtBillNo.Enabled = False
    CLOSEALL = 1
    M_EDIT = False
    Me.Width = 11100
    Me.Height = 10000
    Me.Left = 0

    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If TMPFLAG = False Then TMPREC.Close
        If BILL_FLAG = False Then PHY_BILL.Close
        If ITEM_FLAG = False Then PHY_ITEM.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If INV_FLAG = False Then INV_REC.Close
        
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

Private Sub GRDPOPUPBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
        
            '0 "BILL NO."
            '1 "BILL DATE"
            '2 "UNIT"
            '3 "QTY"
            '4 "MRP"
            '5 "SOLD PRICE"
            '6 "TAX"
            '7 "ITEM"
            '8- EXP DATE
            '9- BATCH
            '10- R_VCH NO
            '11- R_TYPE
            '12 - R_LINE NO
            
            FRMSALERETURNWO.TXTQTY.Text = ""
            FRMSALERETURNWO.TxtPack.Text = GRDPOPUPBILL.Columns(2)
            FRMSALERETURNWO.TXTQTY.Text = GRDPOPUPBILL.Columns(3)
            FRMSALERETURNWO.TXTQTY.Tag = Val(FRMSALERETURNWO.TXTQTY.Text)
            FRMSALERETURNWO.TxtMRP.Text = GRDPOPUPBILL.Columns(4)
            FRMSALERETURNWO.TXTRATE.Text = GRDPOPUPBILL.Columns(5)
            FRMSALERETURNWO.TXTTAX.Text = GRDPOPUPBILL.Columns(6)
            FRMSALERETURNWO.txtBatch.Text = GRDPOPUPBILL.Columns(9)
            FRMSALERETURNWO.TXTVCHNO.Text = GRDPOPUPBILL.Columns(10)
            FRMSALERETURNWO.TXTLINENO.Text = GRDPOPUPBILL.Columns(12)
            FRMSALERETURNWO.TXTTRXTYPE.Text = GRDPOPUPBILL.Columns(11)
            
            Set GRDPOPUPBILL.DataSource = Nothing
            
            FRMEGRDBILL.Visible = False
            FRMEMAIN.Enabled = True
            FRMSALERETURNWO.TXTPRODUCT.Enabled = False
            FRMSALERETURNWO.TxtPack.Enabled = True
            FRMSALERETURNWO.TxtPack.SetFocus
        Case vbKeyEscape
            FRMSALERETURNWO.TXTQTY.Text = ""
            FRMSALERETURNWO.TXTVCHNO.Text = ""
            FRMSALERETURNWO.TXTLINENO.Text = ""
            FRMSALERETURNWO.TXTTRXTYPE.Text = ""
            FRMSALERETURNWO.TxtPack.Text = ""
            
            Set GRDPOPUPBILL.DataSource = Nothing
            FRMEGRDBILL.Visible = False
            FRMEMAIN.Enabled = True
            FRMSALERETURNWO.TXTPRODUCT.Enabled = True
            FRMSALERETURNWO.TxtPack.Enabled = False
            FRMSALERETURNWO.TXTPRODUCT.SetFocus
        
    End Select
End Sub

Private Sub GRDPOPUPITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            FRMSALERETURNWO.TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            FRMSALERETURNWO.TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            For i = 1 To FRMSALERETURNWO.grdsales.Rows - 1
                If Trim(FRMSALERETURNWO.grdsales.TextMatrix(i, 12)) = Trim(FRMSALERETURNWO.TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists.... Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then
                        Set GRDPOPUPITEM.DataSource = Nothing
                        FRMEITEM.Visible = False
                        FRMEMAIN.Enabled = True
                        FRMSALERETURNWO.TXTPRODUCT.Enabled = True
                        FRMSALERETURNWO.TXTQTY.Enabled = False
                        FRMSALERETURNWO.TXTPRODUCT.SetFocus
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next i
          
            Call FILLBILLDB
            If B_FLAG = True Then
                Call FILL_BILLGRID
            Else
                FRMSALERETURNWO.FRMEITEM.Visible = False
                FRMSALERETURNWO.FRMEMAIN.Enabled = True
                If MsgBox("This Item has not been sold to " & DataList2.Text & " this Year... Do You Want to Continue...?", vbYesNo, "SALES RETURN..") = vbYes Then
                    FRMSALERETURNWO.TXTPRODUCT.Enabled = False
                    FRMSALERETURNWO.TxtPack.Enabled = True
                    FRMSALERETURNWO.TxtPack.SetFocus
                Else
                    FRMSALERETURNWO.TXTPRODUCT.Enabled = True
                    FRMSALERETURNWO.TXTPRODUCT.SetFocus
                End If
            End If
        Case vbKeyEscape
            FRMSALERETURNWO.TXTQTY.Text = ""
            FRMSALERETURNWO.TXTVCHNO.Text = ""
            FRMSALERETURNWO.TXTLINENO.Text = ""
            FRMSALERETURNWO.TXTTRXTYPE.Text = ""
            FRMSALERETURNWO.TxtPack.Text = ""
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            FRMSALERETURNWO.TXTPRODUCT.Enabled = True
            FRMSALERETURNWO.TXTQTY.Enabled = False
            FRMSALERETURNWO.TXTPRODUCT.SetFocus
            
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TXTAMOUNT_Change()

End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(txtBatch.Text) = "" Then Exit Sub
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyEscape
            txtBatch.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
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
Dim TRXMAST As ADODB.Recordset
Dim RSTDN As ADODB.Recordset

Dim E_Bill As String
Dim i As Integer
On Error GoTo ErrHand
Select Case KeyCode
    Case vbKeyReturn
        If Val(txtBillNo.Text) = 0 Then Exit Sub
        grdsales.Rows = 1
        i = 0
        EDIT_BILL = False
        Set RSTDN = New ADODB.Recordset
        RSTDN.Open "Select * From SALERETURNWO WHERE VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db2, adOpenStatic, adLockReadOnly
        Do Until RSTDN.EOF
            i = i + 1
            LBLDATE.Caption = Format(RSTDN!VCH_DATE, "DD/MM/YYYY")
            LBLTIME.Caption = Time
            grdsales.Rows = grdsales.Rows + 1
            grdsales.FixedRows = 1
            grdsales.TextMatrix(i, 0) = i
            grdsales.TextMatrix(i, 1) = RSTDN!ITEM_CODE
            grdsales.TextMatrix(i, 2) = RSTDN!ITEM_NAME
            grdsales.TextMatrix(i, 3) = RSTDN!QTY
            grdsales.TextMatrix(i, 4) = Val(RSTDN!UNIT)
            grdsales.TextMatrix(i, 5) = Format(RSTDN!MRP, ".000")
            grdsales.TextMatrix(i, 6) = Format(RSTDN!SALES_PRICE, ".000")
            grdsales.TextMatrix(i, 7) = 0 'DISC
            grdsales.TextMatrix(i, 8) = Val(RSTDN!SALES_TAX)
            grdsales.TextMatrix(i, 9) = RSTDN!REF_NO
            grdsales.TextMatrix(i, 10) = ""
            grdsales.TextMatrix(i, 11) = Format(Val(RSTDN!TRX_TOTAL), ".000")
            
            grdsales.TextMatrix(i, 12) = RSTDN!ITEM_CODE
            grdsales.TextMatrix(i, 13) = RSTDN!R_VCH_NO
            grdsales.TextMatrix(i, 14) = RSTDN!R_LINE_NO
            grdsales.TextMatrix(i, 15) = RSTDN!R_TRX_TYPE
            TXTDEALER.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
            'DataList2.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
            TXTINVDATE.Text = IIf(IsNull(RSTDN!VCH_DATE), Date, RSTDN!VCH_DATE)
            
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMASTWO WHERE ITEMMAST.ITEM_CODE = '" & Trim(RSTDN!ITEM_CODE) & "'", db2, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                grdsales.TextMatrix(i, 16) = IIf(IsNull(TRXMAST!MANUFACTURER), "", Trim(TRXMAST!MANUFACTURER))
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            grdsales.TextMatrix(i, 17) = RSTDN!CHECK_FLAG
            grdsales.TextMatrix(i, 18) = RSTDN!ISSUEQTY
            If RSTDN!CHECK_FLAG = "Y" Then EDIT_BILL = True
            RSTDN.MoveNext
        Loop
        RSTDN.Close
        Set RSTDN = Nothing
        
        LBLTOTAL.Caption = ""
        For i = 1 To grdsales.Rows - 1
            grdsales.TextMatrix(i, 0) = i
            LBLTOTAL.Caption = Format(Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 11)), 2), "0.00")
        Next i
        LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
        lblnetamount.Caption = Format(Round(Val(LBLTOTAL.Caption), 2), "0.00")
        
        TXTSLNO.Text = grdsales.Rows
        txtBillNo.Enabled = False
        TXTSLNO.Enabled = True
        
        If EDIT_BILL = True Then
            CMDEXIT.Caption = "CANCEL"
            FRMEMASTER.Enabled = False
            TXTSLNO.Enabled = False
            cmdview.Enabled = False
            CMDEXIT.SetFocus
            'TXTSLNO.SetFocus
        Else
            cmdview.Enabled = True
            CMDEXIT.Caption = "E&XIT"
            TXTINVDATE.Enabled = False
            DataList2.Enabled = True
            TXTDEALER.Enabled = True
            TXTDEALER.SetFocus
        End If
    
End Select
    
    'DataList2.BoundText = DataList2.TextMatrix(grdSTOCKLESS.Row, 1)
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
    Dim TRXDN As ADODB.Recordset
    Dim i As Integer
    Dim n As Integer
    
    i = 1
    n = 1
    Set TRXDN = New ADODB.Recordset
    TRXDN.Open "Select MAX(Val(VCH_NO)) From SALERETURNWO", db2, adOpenStatic, adLockReadOnly
    If Not (TRXDN.EOF And TRXDN.BOF) Then
        i = IIf(IsNull(TRXDN.Fields(0)), 1, TRXDN.Fields(0) + 1)
        If Val(txtBillNo.Text) > i Then txtBillNo.Text = i
    End If
    TRXDN.Close
    Set TRXDN = Nothing
    
    txtBillNo.Enabled = False
    'Call TXTBILLNO_KeyDown(13, 0)
End Sub

Private Sub TXTDISC_GotFocus()
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.Text)
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            cmdadd.Enabled = True
            TXTDISC.Enabled = False
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
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
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
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
                TXTINVDATE.Enabled = False
                TXTDEALER.Enabled = True
                DataList2.Enabled = True
                TXTDEALER.SetFocus
            End If
        Case vbKeyEscape
            TXTINVDATE.Enabled = False
            txtBillNo.Enabled = True
            txtBillNo.SetFocus
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

Private Sub TXTMRP_GotFocus()
    TxtMRP.SelStart = 0
    TxtMRP.SelLength = Len(TxtMRP.Text)
End Sub

Private Sub TXTMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TxtMRP.Text) = 0 Then Exit Sub
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TxtMRP.Enabled = False
            TXTRATE.Enabled = True
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTRATE.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TxtMRP.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub TXTMRP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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

Private Sub TxtPack_GotFocus()
    TxtPack.SelStart = 0
    TxtPack.SelLength = Len(TxtPack.Text)
End Sub

Private Sub TxtPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TxtPack.Text) = 0 Then Exit Sub
        
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TxtMRP.Enabled = False
            TXTRATE.Enabled = False
            TxtPack.Enabled = False
            TXTQTY.Enabled = True
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTQTY.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TxtPack.Text = ""
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
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

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTP_RATE As ADODB.Recordset

'    On Error GoTo eRRhAND
    Select Case KeyCode
        Case 106
            If TXTQTY.Tag <> "" Then
                TXTPRODUCT.Text = Trim(TXTQTY.Tag)
                TXTPRODUCT.SelStart = 0
                TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
            End If
        Case vbKeyReturn
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            cmddelete.Enabled = False
            'If NONSTOCK = True Then GoTo SKIP
            TxtPack.Text = ""
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            
            txtBatch.Text = ""
            LBLSUBTOTAL.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
           
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] FROM ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] FROM ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 12)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                Next i
                
                Call FILLBILLDB
                
                If B_FLAG = True Then
                    Call FILL_BILLGRID
                Else
                    FRMSALERETURNWO.FRMEITEM.Visible = False
                    FRMSALERETURNWO.FRMEMAIN.Enabled = True
                    If MsgBox("This Item has not been sold to " & DataList2.Text & " this Year... Do You Want to Continue...?", vbYesNo, "SALES RETURN..") = vbYes Then
                        FRMSALERETURNWO.TXTPRODUCT.Enabled = False
                        FRMSALERETURNWO.TxtPack.Enabled = True
                        FRMSALERETURNWO.TxtPack.SetFocus
                    Else
                        FRMSALERETURNWO.TXTPRODUCT.Enabled = True
                        FRMSALERETURNWO.TXTPRODUCT.SetFocus
                    End If
                End If

                Exit Sub
            ElseIf PHY.RecordCount > 1 Then
                'FRMSUB.grdsub.Columns(0).Visible = True
                'FRMSUB.grdsub.Columns(1).Caption = "ITEM NAME"
                'FRMSUB.grdsub.Columns(1).Width = 3200
                'FRMSUB.grdsub.Columns(2).Caption = "QTY"
                'FRMSUB.grdsub.Columns(2).Width = 1300
                Call FILL_ITEMGRID
            End If

            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            cmddelete.Enabled = False
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
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

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            
            If Val(TXTQTY.Text) > Val(TXTQTY.Tag) Then
            
                If (MsgBox("Sold Qty is only .. " & Val(TXTQTY.Tag) & "...Do you want to Continue", vbYesNo, "SALES RETURN") = vbNo) Then
                    TXTQTY.SelStart = 0
                    TXTQTY.SelLength = Len(TXTQTY.Text)
                    Exit Sub
                End If
            End If
        
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TxtMRP.Enabled = True
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TxtMRP.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TxtPack.Text = ""
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TxtPack.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TxtPack.SetFocus
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
    TXTQTY.Text = Format(TXTQTY.Text, ".000")
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
End Sub

Private Sub TXTRATE_GotFocus()
    TXTRATE.SelStart = 0
    TXTRATE.SelLength = Len(TXTRATE.Text)
End Sub

Private Sub TXTRATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTRATE.Text) = 0 Then Exit Sub
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = True
            TxtMRP.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTTAX.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TxtMRP.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TxtMRP.SetFocus
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
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
    cmdview.Enabled = False
    DataList2.Enabled = False
    TXTDEALER.Enabled = False
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.Text) = 0 Then
                TXTSLNO.Text = ""
                TXTPRODUCT.Text = ""
                TXTQTY.Text = ""
                TXTRATE.Text = ""
                TxtMRP.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TxtPack.Text = ""
                LBLSUBTOTAL.Caption = ""
                
                txtBatch.Text = ""
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
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 3)
                TxtMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTRATE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 11), ".000")
                
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 12)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 15)
                TxtPack.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                LBLSUBTOTAL.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 11)
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                TXTRATE.Enabled = False
                TXTTAX.Enabled = False
                
                txtBatch.Enabled = False
                TXTDISC.Enabled = False
                TxtMRP.Enabled = False
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
            TXTTAX.Enabled = False
            
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If cmddelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TxtPack.Text = ""
                TXTQTY.Text = ""
                TXTRATE.Text = ""
                TxtMRP.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                txtBatch.Text = ""
                cmdadd.Enabled = False
                cmddelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.Rows > 1 Then
                CMDPRINT.Enabled = True
                cmdRefresh.Enabled = True
                CMDPRINT.SetFocus
            Else
                FRMEMASTER.Enabled = True
                TXTDEALER.Enabled = True
                DataList2.Enabled = True
                TXTDEALER.SetFocus
            End If
            If M_DELETE = True Then
                cmdRefresh.Enabled = True
                cmdRefresh.SetFocus
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
            TXTTAX.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = True
            TXTTAX.Enabled = False
            
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
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
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
    'Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    'FRMEGRDTMP.Visible = False
    
    
    If ITEM_FLAG = True Then
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY] FROM ITEMMASTWO  WHERE ITEM_NAME Like '" & FRMSALERETURNWO.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME], [CLOSE_QTY] FROM ITEMMASTWO  WHERE ITEM_NAME Like '" & FRMSALERETURNWO.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
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

Function FILL_BILLGRID()
                    
    FRMEMAIN.Enabled = False
    FRMEITEM.Visible = False
    FRMEGRDBILL.Visible = True
    Set GRDPOPUPBILL.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    
    If BILL_FLAG = True Then
        PHY_BILL.Open "Select * From BILLDETAILS ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
        BILL_FLAG = False
    Else
        PHY_BILL.Close
        PHY_BILL.Open "Select * From BILLDETAILS ORDER BY [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
        BILL_FLAG = False
    End If
    
    Set GRDPOPUPBILL.DataSource = PHY_BILL
    
    GRDPOPUPBILL.Columns(0).Caption = "BILL NO."
    GRDPOPUPBILL.Columns(1).Caption = "BILL DATE"
    GRDPOPUPBILL.Columns(2).Caption = "UNIT"
    GRDPOPUPBILL.Columns(3).Caption = "QTY"
    GRDPOPUPBILL.Columns(4).Caption = "MRP"
    GRDPOPUPBILL.Columns(5).Caption = "SOLD PRICE"
    GRDPOPUPBILL.Columns(6).Caption = "TAX"
    GRDPOPUPBILL.Columns(7).Caption = "ITEM"
    GRDPOPUPBILL.Columns(8).Caption = "EXP DATE"
    GRDPOPUPBILL.Columns(9).Caption = "Serial No"
    '10- R_VCH NO
    '11- R_TYPE
    '12 - R_LINE NO
    
    GRDPOPUPBILL.Columns(0).Width = 900
    GRDPOPUPBILL.Columns(1).Width = 1150
    GRDPOPUPBILL.Columns(2).Width = 700
    GRDPOPUPBILL.Columns(3).Width = 700
    GRDPOPUPBILL.Columns(4).Width = 1000
    GRDPOPUPBILL.Columns(5).Width = 1000
    GRDPOPUPBILL.Columns(6).Width = 700
    GRDPOPUPBILL.Columns(7).Width = 0
    GRDPOPUPBILL.Columns(8).Width = 0
    GRDPOPUPBILL.Columns(9).Width = 1200
    GRDPOPUPBILL.Columns(10).Width = 0
    GRDPOPUPBILL.Columns(11).Width = 0
    GRDPOPUPBILL.Columns(12).Width = 0
    
    GRDPOPUPBILL.SetFocus
    LBLHEAD(0).Caption = GRDPOPUPBILL.Columns(7).Text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
    
End Function

Private Function STOCKADJUST()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    M_STOCK = 0
    On Error GoTo ErrHand
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from [RTRXFILEWO] where ITEM_CODE = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rststock.EOF
        M_STOCK = M_STOCK + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
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
    
ErrHand:
    MsgBox Err.Description
End Function

Private Sub FILLCOMBO()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If ACT_FLAG = True Then
        ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    Else
        ACT_REC.Close
        ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    End If
    
    Set Me.CMBDISTI.RowSource = ACT_REC
    CMBDISTI.ListField = "ACT_NAME"
    CMBDISTI.BoundColumn = "ACT_CODE"
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub


Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
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
            lbladdress.Caption = ""
            lbldlno.Caption = ""
            lbltin.Caption = ""
            DataList2.Enabled = True
            DataList2.SetFocus
        Case vbKeyEscape
            TXTDEALER.Enabled = False
            DataList2.Enabled = False
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
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

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
                DataList2.SetFocus
                Exit Sub
            End If
            'FRMEMASTER.Enabled = False
            DataList2.Enabled = False
            TXTDEALER.Enabled = False
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTDEALER.Enabled = True
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

Private Sub DataList2_Click()
    Dim rstCustomer As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    On Error GoTo ErrHand

    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select ADDRESS, DL_NO, KGST from [ACTMAST]  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db2, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = Trim(rstCustomer!ADDRESS)
        lbldlno.Caption = Trim(rstCustomer!DL_NO)
        lbltin.Caption = Trim(rstCustomer!KGST)
    Else
        lbladdress.Caption = ""
        lbldlno.Caption = ""
        lbltin.Caption = ""
    End If
    Call FILLINVOICE
    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text
    cmbinv.Text = ""
    If TXTDEALER.Top = 225 Then
        grdsales.FixedRows = 0
        grdsales.Rows = 1
    End If
    Exit Sub
    
ErrHand:
    MsgBox Err.Description
End Sub


Private Function FILLINVOICE()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set cmbinv.DataSource = Nothing
    If INV_FLAG = True Then
        INV_REC.Open "Select DISTINCT VCH_NO From SALERETURNWO WHERE ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_NO", db2, adOpenStatic, adLockReadOnly
        INV_FLAG = False
    Else
        INV_REC.Close
        INV_REC.Open "Select DISTINCT VCH_NO From SALERETURNWO WHERE ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_NO", db2, adOpenStatic, adLockReadOnly
        INV_FLAG = False
    End If
    
    Set Me.cmbinv.RowSource = INV_REC
    cmbinv.ListField = "VCH_NO"
    cmbinv.BoundColumn = "VCH_NO"
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
End Sub

Private Function VIEWGRID()
    Dim TRXMAST As ADODB.Recordset
    Dim RSTDN As ADODB.Recordset
    
    Dim E_Bill As String
    Dim i As Integer
    On Error GoTo ErrHand
        If Val(txtBillNo.Text) = 0 Then Exit Function
        grdsales.Rows = 1
        i = 0
        Set RSTDN = New ADODB.Recordset
        RSTDN.Open "Select * From SALERETURNWO WHERE VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db2, adOpenStatic, adLockReadOnly
        Do Until RSTDN.EOF
            i = i + 1
            LBLDATE.Caption = Format(RSTDN!VCH_DATE, "DD/MM/YYYY")
            LBLTIME.Caption = Time
            grdsales.Rows = grdsales.Rows + 1
            grdsales.FixedRows = 1
            grdsales.TextMatrix(i, 0) = i
            grdsales.TextMatrix(i, 1) = RSTDN!ITEM_CODE
            grdsales.TextMatrix(i, 2) = RSTDN!ITEM_NAME
            grdsales.TextMatrix(i, 3) = RSTDN!QTY
            grdsales.TextMatrix(i, 4) = Val(RSTDN!UNIT)
            
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMASTWO WHERE ITEMMAST.ITEM_CODE = '" & Trim(RSTDN!ITEM_CODE) & "'", db2, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                grdsales.TextMatrix(i, 16) = IIf(IsNull(TRXMAST!MANUFACTURER), "", Trim(TRXMAST!MANUFACTURER))
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            grdsales.TextMatrix(i, 5) = Format(RSTDN!MRP, ".000")
            grdsales.TextMatrix(i, 6) = Format(RSTDN!SALES_PRICE, ".000")
            grdsales.TextMatrix(i, 7) = 0 'DISC
            grdsales.TextMatrix(i, 8) = Val(RSTDN!SALES_TAX)
            grdsales.TextMatrix(i, 9) = RSTDN!REF_NO
            grdsales.TextMatrix(i, 10) = ""
            grdsales.TextMatrix(i, 11) = Format(Val(RSTDN!TRX_TOTAL), ".000")
            
            grdsales.TextMatrix(i, 12) = RSTDN!ITEM_CODE
            grdsales.TextMatrix(i, 13) = RSTDN!R_VCH_NO
            grdsales.TextMatrix(i, 14) = RSTDN!R_LINE_NO
            grdsales.TextMatrix(i, 15) = RSTDN!R_TRX_TYPE
            TXTDEALER.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
            'DataList2.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
            TXTINVDATE.Text = IIf(IsNull(RSTDN!VCH_DATE), Date, RSTDN!VCH_DATE)
            
            RSTDN.MoveNext
        Loop
        RSTDN.Close
        Set RSTDN = Nothing
        
        LBLTOTAL.Caption = ""
        For i = 1 To grdsales.Rows - 1
            grdsales.TextMatrix(i, 0) = i
            LBLTOTAL.Caption = Format(Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 11)), 2), "0.00")
        Next i
        LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
        lblnetamount.Caption = Format(Round(Val(LBLTOTAL.Caption), 2), "0.00")
        
        TXTSLNO.Text = grdsales.Rows
        Exit Function
ErrHand:
    MsgBox Err.Description

End Function

Private Function FILLBILLDB()
    Dim TRXFILE As ADODB.Recordset
    Dim TRXFILESUB As ADODB.Recordset
    Dim TRXBILL As ADODB.Recordset
    
    Dim n As Integer
    Dim M As Integer
    
    B_FLAG = False
    db2.Execute "delete * From BILLDETAILS"
    Set TRXFILE = New ADODB.Recordset
    TRXFILE.Open "Select * From TRXFILEWO WHERE (TRX_TYPE='RI' OR TRX_TYPE='SI') AND CST <>2 AND ITEM_CODE = '" & FRMSALERETURNWO.TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "'", db2, adOpenStatic, adLockReadOnly
    Do Until TRXFILE.EOF
        Set TRXFILESUB = New ADODB.Recordset
        TRXFILESUB.Open "Select * From TRXSUBWO WHERE VCH_NO = " & TRXFILE!VCH_NO & " AND LINE_NO = " & TRXFILE!LINE_NO & "", db2, adOpenStatic, adLockReadOnly

        If Not (TRXFILESUB.EOF And TRXFILESUB.BOF) Then
            Set TRXBILL = New ADODB.Recordset
            TRXBILL.Open "SELECT *  FROM BILLDETAILS", db2, adOpenStatic, adLockOptimistic, adCmdText
            B_FLAG = True
            TRXBILL.AddNew
            TRXBILL!VCH_NO = TRXFILESUB!VCH_NO
            TRXBILL!TRX_TYPE = TRXFILESUB!TRX_TYPE
            TRXBILL!LINE_NO = TRXFILESUB!LINE_NO
            TRXBILL!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
            TRXBILL!MRP = TRXFILE!MRP
            TRXBILL!UNIT = TRXFILE!UNIT
            TRXBILL!QTY = TRXFILE!QTY
            TRXBILL!PTR = TRXFILE!PTR
            TRXBILL!SALES_TAX = IIf(IsNull(TRXFILE!SALES_TAX), 0, TRXFILE!SALES_TAX)
            TRXBILL!ITEM_NAME = TRXFILE!ITEM_NAME
            TRXBILL!EXP_DATE = Null
            TRXBILL!REF_NO = TRXFILE!REF_NO
            TRXBILL!R_VCH_NO = TRXFILESUB!R_VCH_NO
            TRXBILL!R_TRX_TYPE = TRXFILESUB!R_TRX_TYPE
            TRXBILL!R_LINE_NO = TRXFILESUB!R_LINE_NO
            
            TRXBILL.Update
            TRXBILL.Close
            Set TRXBILL = Nothing
        End If
        TRXFILESUB.Close
        Set TRXFILESUB = Nothing
        TRXFILE.MoveNext
    Loop
    TRXFILE.Close
    Set TRXFILE = Nothing

    Set GRDPOPUPITEM.DataSource = Nothing
End Function
