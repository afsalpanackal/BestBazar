VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FRMCN 
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DELIVERY NOTE"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13560
   ControlBox      =   0   'False
   Icon            =   "FRMCN.frx":0000
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
      Left            =   1785
      TabIndex        =   58
      Top             =   3615
      Visible         =   0   'False
      Width           =   6030
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   2835
         Left            =   75
         TabIndex        =   59
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
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   1860
      TabIndex        =   54
      Top             =   3630
      Visible         =   0   'False
      Width           =   5835
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2535
         Left            =   90
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
         Top             =   105
         Visible         =   0   'False
         Width           =   2610
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   -105
      TabIndex        =   18
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
         TabIndex        =   62
         Top             =   8685
         Visible         =   0   'False
         Width           =   930
      End
      Begin Crystal.CrystalReport rptprint 
         Left            =   0
         Top             =   8460
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00800080&
         Height          =   1725
         Left            =   210
         TabIndex        =   19
         Top             =   15
         Width           =   10845
         Begin MSDataListLib.DataCombo cmbinv 
            Height          =   330
            Left            =   1545
            TabIndex        =   83
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
            TabIndex        =   16
            Top             =   225
            Width           =   885
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   345
            Left            =   3840
            TabIndex        =   63
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
            TabIndex        =   80
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
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   65
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
            TabIndex        =   64
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   225
            Width           =   1110
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00800080&
         Height          =   6465
         Left            =   210
         TabIndex        =   24
         Top             =   1695
         Width           =   10830
         Begin VB.Frame Frame3 
            BackColor       =   &H00800080&
            ForeColor       =   &H00FFFFFF&
            Height          =   6240
            Left            =   8880
            TabIndex        =   25
            Top             =   165
            Width           =   1815
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
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   28
               Left            =   210
               TabIndex        =   88
               Top             =   5460
               Width           =   1395
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
               Height          =   360
               Left            =   210
               TabIndex        =   87
               Top             =   5715
               Width           =   1425
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
               Height          =   345
               Left            =   210
               TabIndex        =   86
               Top             =   4995
               Width           =   1425
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
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   27
               Left            =   240
               TabIndex        =   85
               Top             =   4710
               Width           =   1425
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
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   26
               Left            =   150
               TabIndex        =   77
               Top             =   3135
               Width           =   1515
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
               Height          =   570
               Left            =   180
               TabIndex        =   76
               Top             =   3420
               Width           =   1440
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
               Height          =   570
               Left            =   180
               TabIndex        =   75
               Top             =   2505
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "COST PRICE"
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
               Index           =   25
               Left            =   150
               TabIndex        =   74
               Top             =   2205
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
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   23
               Left            =   150
               TabIndex        =   73
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
               TabIndex        =   66
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
               TabIndex        =   27
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
               TabIndex        =   26
               Top             =   645
               Width           =   1440
            End
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
            Left            =   6240
            TabIndex        =   29
            Top             =   6075
            Width           =   600
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
            Left            =   7755
            TabIndex        =   28
            Top             =   6060
            Width           =   1080
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   5730
            Left            =   90
            TabIndex        =   17
            Top             =   270
            Width           =   8730
            _ExtentX        =   15399
            _ExtentY        =   10107
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
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
            Index           =   4
            Left            =   5325
            TabIndex        =   31
            Top             =   6105
            Width           =   870
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Left            =   6900
            TabIndex        =   30
            Top             =   6105
            Width           =   780
         End
      End
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   465
         Left            =   11100
         TabIndex        =   53
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
         BackColor       =   &H00800080&
         Height          =   1365
         Left            =   210
         TabIndex        =   32
         Top             =   8070
         Width           =   10830
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
            TabIndex        =   84
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
            TabIndex        =   78
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
            Left            =   5010
            MaxLength       =   6
            TabIndex        =   60
            Top             =   450
            Width           =   630
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
            TabIndex        =   10
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
            Left            =   165
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
            Left            =   750
            TabIndex        =   2
            Top             =   450
            Width           =   3420
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
            Left            =   4215
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
            Left            =   5655
            MaxLength       =   6
            TabIndex        =   4
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
            Left            =   6315
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
            Left            =   9060
            MaxLength       =   4
            TabIndex        =   8
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
            Left            =   5400
            TabIndex        =   13
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
            TabIndex        =   15
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   37
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
            Left            =   8085
            MaxLength       =   15
            TabIndex        =   7
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
            Left            =   4035
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   14
            Top             =   810
            Width           =   1125
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   285
            Left            =   6945
            TabIndex        =   6
            Top             =   465
            Width           =   1095
            _ExtentX        =   1931
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
            Left            =   5010
            TabIndex        =   61
            Top             =   225
            Width           =   630
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
            Left            =   165
            TabIndex        =   52
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
            Left            =   750
            TabIndex        =   51
            Top             =   225
            Width           =   3420
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
            Left            =   4215
            TabIndex        =   50
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
            Left            =   5655
            TabIndex        =   49
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
            Left            =   6315
            TabIndex        =   48
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
            Left            =   9060
            TabIndex        =   47
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
            Left            =   9765
            TabIndex        =   46
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
            TabIndex        =   45
            Top             =   1260
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
            Left            =   6945
            TabIndex        =   44
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
            Left            =   8085
            TabIndex        =   43
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
            Left            =   9765
            TabIndex        =   9
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   1275
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lblnonstockdisplay 
            Caption         =   "Enter Non Stock Items"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   165
            TabIndex        =   38
            Top             =   735
            Visible         =   0   'False
            Width           =   3270
         End
      End
   End
   Begin Crystal.CrystalReport rptprintsmall 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSDataListLib.DataCombo CMBDISTI 
      Height          =   1020
      Left            =   9825
      TabIndex        =   79
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
      TabIndex        =   82
      Top             =   2610
      Width           =   495
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11445
      TabIndex        =   81
      Top             =   3255
      Width           =   1620
   End
End
Attribute VB_Name = "FRMCN"
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
Dim PHY_BATCH As New ADODB.Recordset
Dim BATCH_FLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean
Dim INV_FLAG As Boolean
Dim INV_REC As New ADODB.Recordset

Dim CLOSEALL As Integer
Dim M_STOCK As Integer
Dim M_EDIT As Boolean
Dim NONSTOCK As Boolean
Dim M_DELETE As Boolean
Dim EDIT_BILL As Boolean

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
    
    On Error GoTo eRRhAND
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Format(Val(TxtMRP.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTRATE.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Val(TXTDISC.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(TXTTAX.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = IIf(TXTEXPIRY.Text = "  /  ", "", Trim(TXTEXPIRY.Text))
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTTRXTYPE.Text)
    
    If NONSTOCK = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = "Y"
        lblnonstockdisplay.Visible = False
        NONSTOCK = False
        GoTo SKIP
    Else
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = "N"
    End If
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = Trim(RSTTRXFILE!MANUFACTURER)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            If Val(!ISSUE_VAL) > 0 Then !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            If Val(!CLOSE_VAL) > 0 Then !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
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
    
SKIP:
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 11)), 2)
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = Round((Val(LBLTOTAL.Caption) * Val(TXTTOTALDISC.Text) / 100), 2)
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2)
    
    'Call STOCKADJUST
    
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Text = ""
    TxtMRP.Text = ""
    TXTRATE.Text = ""
    TXTTAX.Text = ""
    TXTDISC.Text = ""
    txtBatch.Text = ""
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTSLNO.Enabled = True
    M_EDIT = False
    Call COSTCALCULATION
    TXTSLNO.SetFocus
    grdsales.TopRow = grdsales.Rows - 1
Exit Sub
eRRhAND:
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
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            If Val(!ISSUE_VAL) > 0 Then !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            If Val(!CLOSE_VAL) > 0 Then !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
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
    Next i
    grdsales.Rows = grdsales.Rows - 1
    LBLTOTAL.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 11))
    Next i
    
    Call COSTCALCULATION
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    TxtMRP.Text = ""
    TXTTAX.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTDISC.Text = ""
    txtBatch.Text = ""
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    M_DELETE = True
    If grdsales.Rows = 1 Then
        CMDEXIT.Enabled = True
        CMDPRINT.Enabled = False
    End If
    
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
        LBLPROFIT.Caption = ""
        LBLDATE.Caption = Date
        LBLTIME.Caption = Time
        LBLTOTAL.Caption = ""
        TXTTOTALDISC.Text = ""
        LBLTOTALCOST.Caption = ""
        TXTAMOUNT.Text = ""
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
    
    On Error GoTo eRRhAND
    
    If grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = "Y" Then NONSTOCK = True
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            If Val(!ISSUE_VAL) > 0 Then !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            If Val(!CLOSE_VAL) > 0 Then !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = True
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
    Exit Sub
eRRhAND:
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
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub cmdprint_Click()
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
    
    db2.Execute "delete * From TRXFILE"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRXFILE", db2, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        If grdsales.TextMatrix(i, 16) = "Y" Then GoTo GOSKIP
        RSTTRXFILE.AddNew
        
        Set TRXMAST = New ADODB.Recordset
        TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMAST WHERE ITEMMAST.ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 12)) & "'", db, adOpenStatic, adLockReadOnly
        If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
            RSTTRXFILE!MFGR = TRXMAST!MANUFACTURER
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
        RSTTRXFILE!EXP_DATE = grdsales.TextMatrix(i, 10)
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
    TXTEXPIRY.Enabled = False
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
    
    If grdsales.Rows = 1 Then GoTo SKIP
    
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
    On Error GoTo eRRhAND
    
    db2.Execute "delete * From TEMPCN WHERE VCH_NO = " & Val(txtBillNo.Text) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPCN", db2, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!TRX_TYPE = "DN"
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
        RSTTRXFILE!VCH_DESC = "Issued to     " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 11)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        If grdsales.TextMatrix(i, 10) = "" Then
            RSTTRXFILE!EXP_DATE = Null
        Else
            RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(i, 10))) & "/" & Trim(grdsales.TextMatrix(i, 10))
        End If
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!VCH_DESC = "Received From " & DataList2.Text
        RSTTRXFILE!CHECK_FLAG = "N"

        RSTTRXFILE!R_VCH_NO = grdsales.TextMatrix(i, 13)
        RSTTRXFILE!R_LINE_NO = grdsales.TextMatrix(i, 14)
        RSTTRXFILE!R_TRX_TYPE = grdsales.TextMatrix(i, 15)
         
        RSTTRXFILE.Update
    Next i

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    txtBillNo.Text = 1
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select MAX(Val(VCH_NO)) From TEMPCN", db2, adOpenStatic, adLockReadOnly
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
    LBLPROFIT.Caption = ""
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    LBLTOTAL.Caption = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    TXTAMOUNT.Text = ""
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    M_EDIT = False
    M_DELETE = False
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
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    cmdview.Enabled = True
    
    Exit Sub
eRRhAND:
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
    TXTDEALER.SetFocus
    cmdview.Enabled = False
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo eRRhAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(VCH_NO)) From TEMPCN", db2, adOpenStatic, adLockReadOnly
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
    grdsales.ColWidth(10) = 1100
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
    Me.Top = 0
    INV_FLAG = True
    M_DELETE = False
    
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTTAX.Enabled = False
    TXTEXPIRY.Enabled = False
    txtBatch.Enabled = False
    TXTDISC.Enabled = False
    CmdDelete.Enabled = False
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
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If TMPFLAG = False Then TMPREC.Close
        If BATCH_FLAG = False Then PHY_BATCH.Close
        If ITEM_FLAG = False Then PHY_ITEM.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If INV_FLAG = False Then INV_REC.Close
        
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDPOPUP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTP_RATE As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            'FRMCN.TXTQTY.Text = GRDPOPUP.Columns(1)
            FRMCN.TxtMRP.Text = GRDPOPUP.Columns(3)
            'FRMCN.TXTRATE.Text = GRDPOPUP.Columns(4)
            FRMCN.TXTTAX.Text = 0  'GRDPOPUP.Columns(5)
            FRMCN.TXTEXPIRY.Text = IIf(GRDPOPUP.Columns(2) = "", "  /  ", Format(GRDPOPUP.Columns(2), "mm/yy"))
            FRMCN.txtBatch.Text = GRDPOPUP.Columns(0)
            
            FRMCN.TXTVCHNO.Text = GRDPOPUP.Columns(8)
            FRMCN.TXTLINENO.Text = GRDPOPUP.Columns(9)
            FRMCN.TXTTRXTYPE.Text = GRDPOPUP.Columns(10)
            FRMCN.TXTUNIT.Text = GRDPOPUP.Columns(11)
            
            Set RSTP_RATE = New ADODB.Recordset
            RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & GRDPOPUP.Columns(6) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
                FRMCN.TXTRATE.Text = RSTP_RATE!SALES_PRICE
            End If
            RSTP_RATE.Close
            Set RSTP_RATE = Nothing
            
            Set GRDPOPUP.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            FRMCN.TXTPRODUCT.Enabled = False
            FRMCN.TXTQTY.Enabled = True
            FRMCN.TXTQTY.SetFocus
        Case vbKeyEscape
            FRMCN.TXTQTY.Text = ""
            FRMCN.TXTVCHNO.Text = ""
            FRMCN.TXTLINENO.Text = ""
            FRMCN.TXTTRXTYPE.Text = ""
            FRMCN.TXTUNIT.Text = ""
            
            Set GRDPOPUP.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            FRMCN.TXTPRODUCT.Enabled = True
            FRMCN.TXTQTY.Enabled = False
            FRMCN.TXTPRODUCT.SetFocus
        
    End Select
End Sub

Private Sub GRDPOPUPITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset

    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            M_STOCK = Val(GRDPOPUPITEM.Columns(2))
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            FRMCN.TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            FRMCN.TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            If M_STOCK = 0 Then
                If (MsgBox("NO STOCK AVAILABLE..Do you want to add to Stockless", vbYesNo, "SALES") = vbYes) Then
                    Set RSTNONSTOCK = New ADODB.Recordset
                    RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & FRMCN.TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                    If (RSTNONSTOCK.EOF And RSTNONSTOCK.BOF) Then
                        RSTNONSTOCK.AddNew
                        RSTNONSTOCK!ITEM_NAME = TXTPRODUCT.Text
                        RSTNONSTOCK!ITEM_CODE = TXTITEMCODE.Text
                        RSTNONSTOCK!Date = Date & " " & Time
                        RSTNONSTOCK.Update
                    End If
                    RSTNONSTOCK.Close
                    Set RSTNONSTOCK = Nothing
                End If
                Exit Sub
            End If
            For i = 1 To FRMCN.grdsales.Rows - 1
                If Trim(FRMCN.grdsales.TextMatrix(i, 12)) = Trim(FRMCN.TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists.... Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then
                        Set GRDPOPUPITEM.DataSource = Nothing
                        FRMEITEM.Visible = False
                        FRMEMAIN.Enabled = True
                        FRMCN.TXTPRODUCT.Enabled = True
                        FRMCN.TXTQTY.Enabled = False
                        FRMCN.TXTPRODUCT.SetFocus
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next i
            Set GRDPOPUPITEM.DataSource = Nothing
            If ITEM_FLAG = True Then
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP From RTRXFILE  WHERE ITEM_CODE = '" & FRMCN.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, MRP From RTRXFILE  WHERE ITEM_CODE = '" & FRMCN.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            End If
            Set GRDPOPUPITEM.DataSource = PHY_ITEM
            If PHY_ITEM.RecordCount = 1 Then
                'FRMCN.TXTQTY.Text = GRDPOPUPITEM.Columns(2)
                'FRMCN.TXTRATE.Text = GRDPOPUPITEM.Columns(3)
                FRMCN.TxtMRP.Text = GRDPOPUPITEM.Columns(11)
                FRMCN.TXTTAX.Text = 0 'GRDPOPUPITEM.Columns(4)
                FRMCN.TXTEXPIRY.Text = IIf(GRDPOPUPITEM.Columns(7) = "", "  /  ", Format(GRDPOPUPITEM.Columns(7), "MM/YY"))
                FRMCN.txtBatch.Text = GRDPOPUPITEM.Columns(6)
                
                FRMCN.TXTVCHNO.Text = GRDPOPUPITEM.Columns(8)
                FRMCN.TXTLINENO.Text = GRDPOPUPITEM.Columns(9)
                FRMCN.TXTTRXTYPE.Text = GRDPOPUPITEM.Columns(10)
                FRMCN.TXTUNIT.Text = GRDPOPUPITEM.Columns(5)
                
                Set RSTP_RATE = New ADODB.Recordset
                RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & GRDPOPUPITEM.Columns(0) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
                    FRMCN.TXTRATE.Text = RSTP_RATE!SALES_PRICE
                End If
                RSTP_RATE.Close
                Set RSTP_RATE = Nothing
                
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                FRMCN.TXTPRODUCT.Enabled = False
                FRMCN.TXTQTY.Enabled = True
                FRMCN.TXTQTY.SetFocus
                Exit Sub
            ElseIf PHY.RecordCount > 1 Then
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEGRDTMP.Visible = False
                Call FILL_BATCHGRID
            End If
        Case vbKeyEscape
            FRMCN.TXTQTY.Text = ""
            FRMCN.TXTVCHNO.Text = ""
            FRMCN.TXTLINENO.Text = ""
            FRMCN.TXTTRXTYPE.Text = ""
            FRMCN.TXTUNIT.Text = ""
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            FRMCN.TXTPRODUCT.Enabled = True
            FRMCN.TXTQTY.Enabled = False
            FRMCN.TXTPRODUCT.SetFocus
            
    End Select
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = True
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTEXPIRY.SetFocus
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

Dim E_BILL As String
Dim i As Integer
On Error GoTo eRRhAND
Select Case KeyCode
    Case vbKeyReturn
        If Val(txtBillNo.Text) = 0 Then Exit Sub
        grdsales.Rows = 1
        i = 0
        EDIT_BILL = False
        Set RSTDN = New ADODB.Recordset
        RSTDN.Open "Select * From TEMPCN WHERE VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db2, adOpenStatic, adLockReadOnly
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
            TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMAST WHERE ITEMMAST.ITEM_CODE = '" & Trim(RSTDN!ITEM_CODE) & "'", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                grdsales.TextMatrix(i, 17) = Trim(TRXMAST!MANUFACTURER)
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            grdsales.TextMatrix(i, 5) = Format(RSTDN!MRP, ".000")
            grdsales.TextMatrix(i, 6) = Format(RSTDN!SALES_PRICE, ".000")
            grdsales.TextMatrix(i, 7) = 0 'DISC
            grdsales.TextMatrix(i, 8) = Val(RSTDN!SALES_TAX)
            grdsales.TextMatrix(i, 9) = RSTDN!REF_NO
            grdsales.TextMatrix(i, 10) = Format(RSTDN!EXP_DATE, "MM/YY")
            grdsales.TextMatrix(i, 11) = Format(Val(RSTDN!TRX_TOTAL), ".000")
            
            grdsales.TextMatrix(i, 12) = RSTDN!ITEM_CODE
            grdsales.TextMatrix(i, 13) = RSTDN!R_VCH_NO
            grdsales.TextMatrix(i, 14) = RSTDN!R_LINE_NO
            grdsales.TextMatrix(i, 15) = RSTDN!R_TRX_TYPE
            TXTDEALER.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
            'DataList2.Text = IIf(IsNull(RSTDN!VCH_DESC), "", Mid(RSTDN!VCH_DESC, 15))
            TXTINVDATE.Text = IIf(IsNull(RSTDN!VCH_DATE), Date, RSTDN!VCH_DATE)
            
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
        TXTAMOUNT.Text = Round((Val(LBLTOTAL.Caption) * Val(TXTTOTALDISC.Text) / 100), 2)
        lblnetamount.Caption = Format(Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2), "0.00")
        Call COSTCALCULATION
        
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
       
    Case vbKeyF3
        E_BILL = Val(InputBox("Enter the Bill No"))

        Set TRXFILE = New ADODB.Recordset
        TRXFILE.Open "Select VCH_NO From TEMPCN WHERE TRX_TYPE='DN' AND VCH_NO = " & Val(E_BILL) & "", db2, adOpenStatic, adLockReadOnly
        If Not (TRXFILE.EOF Or TRXFILE.BOF) Then
            txtBillNo.Text = TRXFILE!T_VCH_NO
        End If
        TRXFILE.Close
        Set TRXFILE = Nothing
        txtBillNo.SetFocus
End Select
    
    'DataList2.BoundText = DataList2.TextMatrix(grdSTOCKLESS.Row, 1)
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    
    Exit Sub
eRRhAND:
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
    Dim N As Integer
    
    i = 1
    N = 1
    Set TRXDN = New ADODB.Recordset
    TRXDN.Open "Select MAX(Val(VCH_NO)) From TEMPCN", db2, adOpenStatic, adLockReadOnly
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
            TXTEXPIRY.Enabled = False
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
            TXTEXPIRY.Enabled = False
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
            If Val(TxtMRP.Text) = 0 Then Exit Sub
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TxtMRP.Enabled = False
            TXTRATE.Enabled = True
            TXTEXPIRY.Enabled = False
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
            TXTEXPIRY.Enabled = False
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

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTP_RATE As ADODB.Recordset

    On Error GoTo eRRhAND
    Select Case KeyCode
        Case 106
            If TXTQTY.Tag <> "" Then
                TXTPRODUCT.Text = Trim(TXTQTY.Tag)
                TXTPRODUCT.SelStart = 0
                TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
            End If
        Case vbKeyReturn
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            CmdDelete.Enabled = False
            If NONSTOCK = True Then GoTo SKIP
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            TxtMRP.Text = ""
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            TXTEXPIRY.Text = "  /  "
            txtBatch.Text = ""
            LBLSUBTOTAL.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
           
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME],[CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME],[CLOSE_QTY] From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
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
                Set grdtmp.DataSource = Nothing
                If TMPFLAG = True Then
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE From RTRXFILE  WHERE ITEM_CODE = '" & FRMCN.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                Else
                    TMPREC.Close
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE From RTRXFILE  WHERE ITEM_CODE = '" & FRMCN.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                    TMPFLAG = False
                End If
                
                Set grdtmp.DataSource = TMPREC
                If TMPREC.RecordCount = 1 Then
                    'FRMCN.TXTQTY.Text = grdtmp.Columns(2)
                    FRMCN.TxtMRP.Text = grdtmp.Columns(3)
                    'FRMCN.TXTRATE.Text = grdtmp.Columns(4)
                    FRMCN.TXTTAX.Text = 0 'grdtmp.Columns(5)
                    FRMCN.TXTEXPIRY.Text = IIf(grdtmp.Columns(8) = "", "  /  ", Format(grdtmp.Columns(8), "MM/YY"))
                    FRMCN.txtBatch.Text = grdtmp.Columns(7)
                    
                    FRMCN.TXTVCHNO.Text = grdtmp.Columns(9)
                    FRMCN.TXTLINENO.Text = grdtmp.Columns(10)
                    FRMCN.TXTTRXTYPE.Text = grdtmp.Columns(11)
                    FRMCN.TXTUNIT.Text = grdtmp.Columns(6)
                    
                    Set RSTP_RATE = New ADODB.Recordset
                    RSTP_RATE.Open "Select * From P_Rate Where CUST_CODE='" & Trim(DataList2.BoundText) & "' And ITEM_CODE='" & grdtmp.Columns(0) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
                    If Not (RSTP_RATE.EOF And RSTP_RATE.BOF) Then
                        FRMCN.TXTRATE.Text = RSTP_RATE!SALES_PRICE
                    End If
                    RSTP_RATE.Close
                    Set RSTP_RATE = Nothing
                    
                    FRMCN.TXTPRODUCT.Enabled = False
                    FRMCN.TXTQTY.Enabled = True
                    FRMCN.TXTQTY.SetFocus
                    Exit Sub
                ElseIf TMPREC.RecordCount = 0 Then
                    FRMCN.TXTQTY.Text = 0
                    If (MsgBox("NO STOCK AVAILABLE..Do you want to add to Stockless", vbYesNo, "SALES") = vbYes) Then
                        Set RSTNONSTOCK = New ADODB.Recordset
                        RSTNONSTOCK.Open "Select * From NONRCVD WHERE Item_Code = '" & FRMCN.TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                        If (RSTNONSTOCK.EOF And RSTNONSTOCK.BOF) Then
                            RSTNONSTOCK.AddNew
                            RSTNONSTOCK!ITEM_NAME = TXTPRODUCT.Text
                            RSTNONSTOCK!ITEM_CODE = TXTITEMCODE.Text
                            RSTNONSTOCK!Date = Date & " " & Time
                            RSTNONSTOCK.Update
                        End If
                        RSTNONSTOCK.Close
                        Set RSTNONSTOCK = Nothing
                    End If
                    'MsgBox "NO STOCK...", vbOKOnly, "BILL.."
                    TXTPRODUCT.Enabled = True
                    TXTQTY.Enabled = False
                    TXTPRODUCT.SelStart = 0
                    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
                    TXTPRODUCT.SetFocus
                    Exit Sub
                ElseIf TMPREC.RecordCount > 1 Then
                    Call FILL_BATCHGRID
                    Exit Sub
                End If
SKIP:
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                TXTRATE.Enabled = False
                TXTTAX.Enabled = False
                TXTEXPIRY.Enabled = False
                txtBatch.Enabled = False
                TXTDISC.Enabled = False
                TXTQTY.SetFocus
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
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            CmdDelete.Enabled = False
        Case vbKeyEscape
            If NONSTOCK = True Then
                NONSTOCK = False
                TXTPRODUCT.Text = ""
                TXTQTY.Text = ""
                TXTRATE.Text = ""
                TxtMRP.Text = ""
                TXTTAX.Text = ""
                TXTEXPIRY.Text = "  /  "
                txtBatch.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                lblnonstockdisplay.Visible = False
            End If
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
eRRhAND:
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

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If NONSTOCK = True Then GoTo SKIP
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT BAL_QTY  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockReadOnly
             If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                i = RSTTRXFILE!BAL_QTY
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            'TXTEXPIRY.Text = Date
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT EXP_DATE  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!EXP_DATE)) Then
                    TXTEXPIRY.Text = "  /  "
                    txtexpirydate.Text = ""
                Else
                    TXTEXPIRY.Text = Format(RSTTRXFILE!EXP_DATE, "MM/YY")
                    txtexpirydate.Text = RSTTRXFILE!EXP_DATE
                End If
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            If Val(TXTQTY.Text) > i Then
                MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                TXTQTY.SelStart = 0
                TXTQTY.SelLength = Len(TXTQTY.Text)
                Exit Sub
            End If
            If TXTEXPIRY.Text = "  /  " Then GoTo SKIP
            If DateDiff("d", Date, txtexpirydate.Text) < 0 Then
                MsgBox "Item Expired....", vbOKOnly, "BILL.."
                TXTQTY.SelStart = 0
                TXTQTY.SelLength = Len(TXTQTY.Text)
                Exit Sub
            End If
            
            If DateDiff("d", Date, txtexpirydate.Text) < 90 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, txtexpirydate.Text)) & "Days", vbOKOnly, "BILL.."
                TXTQTY.SelStart = 0
                TXTQTY.SelLength = Len(TXTQTY.Text)
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TxtMRP.Enabled = True
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            TxtMRP.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TXTUNIT.Text = ""
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
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
            If Val(TXTRATE.Text) = 0 Then Exit Sub
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = True
            TxtMRP.Enabled = False
            TXTEXPIRY.Enabled = False
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
            TXTEXPIRY.Enabled = False
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
                TXTUNIT.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTEXPIRY.Text = "  /  "
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
                TxtMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTRATE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 11), ".000")
                
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 12)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 15)
                TXTUNIT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                LBLSUBTOTAL.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 11)
                If grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = "" Then
                    TXTEXPIRY.Text = "  /  "
                Else
                    TXTEXPIRY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                End If
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                TXTRATE.Enabled = False
                TXTTAX.Enabled = False
                TXTEXPIRY.Enabled = False
                txtBatch.Enabled = False
                TXTDISC.Enabled = False
                TxtMRP.Enabled = False
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
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
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
                TxtMRP.Text = ""
                TXTTAX.Text = ""
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTEXPIRY.Text = "  /  "
                txtBatch.Text = ""
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
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
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTTAX.Enabled = False
            If NONSTOCK = True Then
                TXTEXPIRY.Visible = True
                TXTEXPIRY.SetFocus
            Else
                TXTEXPIRY.Enabled = True
                TXTEXPIRY.SetFocus
            End If
            txtBatch.Enabled = False
            TXTDISC.Enabled = False
            
        Case vbKeyEscape
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = True
            TXTTAX.Enabled = False
            TXTEXPIRY.Enabled = False
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

Private Sub TXTTOTALDISC_GotFocus()
    TXTTOTALDISC.SelStart = 0
    TXTTOTALDISC.SelLength = Len(TXTTOTALDISC.Text)
End Sub

Private Sub TXTTOTALDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
        Case vbKeyReturn
            If TXTSLNO.Enabled = True Then TXTSLNO.SetFocus
        End Select
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
    Dim RSTDOCTOR As ADODB.Recordset
    Dim RSTPATIENT As ADODB.Recordset
    Dim vlineCount As Integer
    Dim vpageCount As Integer
    Dim SN As Integer
    Dim i As Integer
    Dim HN As Integer
    Dim LN As Integer
    vlineCount = 0
    vpageCount = 1
    SN = 0

    On Error GoTo eRRhAND

    'Set Rs = New ADODB.Recordset
    'strSQL = "Select * from products"
    'Rs.Open strSQL, cnn
    
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
              
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(7) & " Kaichoondy Junction" & Space(19) & "Phone: 0477-3290525"
      
    Print #1, Space(7) & "Alappuzha 688006" & Space(15) & "DL No. 6-176/20/2003 Dtd. 31.10.2003"
    Print #1, Space(45) & "6-177/20/2003 Dtd. 31.10.2003"
              
    Print #1, Space(7) & "TIN No.32041339615"
              
    Print #1, Chr(27) & Chr(71) & Chr(10) & Space(7) & AlignRight("INVOICE FORM 8BF", 38)
        
    If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, Date)
    Print #1, Space(7) & "Bill No. " & Trim(LBLBILLNO.Caption) & Chr(27) & Chr(72) & Space(28) & "Date:" & LBLDATE.Caption '& Space(2) & LBLTIME.Caption
    LBLDATE.Caption = Date
    'Print #1, Chr(27) & Chr(72) & Space(7) & "Patient: " & Trim(TXTPATIENT.Text) & Space(27); "Doctor: " & TXTDOCTOR.Text
              
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
            AlignLeft(Format(grdsales.TextMatrix(i, 10), "mm/yy"), 6) & _
            AlignRight(grdsales.TextMatrix(i, 3), 4) & Space(1) & _
            AlignRight(Format(grdsales.TextMatrix(i, 11), ".00"), 6) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Next i
    
    Print #1, Space(7) & AlignRight("-------------", 65)
    Print #1, Chr(27) & Chr(71) & Space(9) & AlignRight("NET AMOUNT", 52) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 10)
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
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
End Sub

Public Function AlignLeft(vStr As String, vSpace As Integer) As String
    If Len(Trim(vStr)) > vSpace Then '//if the string length is greater than the space you mention
        AlignLeft = Left(vStr, vSpace)  '&"..."
        Exit Function
    End If
    
    AlignLeft = vStr & Space(vSpace - Len(Trim(vStr)))
End Function

Public Function AlignRight(vNumber As String, vSpace As Integer) As String
    AlignRight = Space(vSpace - Len(Trim(vNumber))) & vNumber
End Function

Public Function RepeatString(vStr As String, vSpace As Integer) As String

    Dim X As Integer
    
    For X = 1 To vSpace
        RepeatString = RepeatString & vStr
    Next X
End Function

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
    TempDate = Year(DateIn) & "-" & Format(Month(DateIn), "00") & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Function FILL_BATCHGRID()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT From RTRXFILE  WHERE ITEM_CODE = '" & FRMCN.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT From RTRXFILE  WHERE ITEM_CODE = '" & FRMCN.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "BATCH NO."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "EXP DATE"
    GRDPOPUP.Columns(3).Caption = "MRP"
    GRDPOPUP.Columns(4).Caption = "RATE"
    GRDPOPUP.Columns(5).Caption = "TAX"
    GRDPOPUP.Columns(6).Caption = "VCH No"
    GRDPOPUP.Columns(7).Caption = "Line No"
    GRDPOPUP.Columns(8).Caption = "Trx Type"
    
    GRDPOPUP.Columns(0).Width = 1400
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 1400
    GRDPOPUP.Columns(3).Width = 1000
    GRDPOPUP.Columns(4).Width = 1000
    GRDPOPUP.Columns(5).Width = 900
    
    GRDPOPUP.Columns(6).Visible = False
    GRDPOPUP.Columns(7).Visible = False
    GRDPOPUP.Columns(8).Visible = False
    
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).Text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
End Function

Function FILL_ITEMGRID()
    FRMEMAIN.Enabled = False
    FRMEITEM.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEGRDTMP.Visible = False
    
    
    If ITEM_FLAG = True Then
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & FRMCN.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & FRMCN.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
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

Private Function STOCKADJUST()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    M_STOCK = 0
    On Error GoTo eRRhAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rststock.EOF
        M_STOCK = M_STOCK + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTITEMMAST
        If Not (.EOF And .BOF) Then
            !OPEN_QTY = M_STOCK
            !OPEN_VAL = 0
            !RCPT_QTY = 0
            !RCPT_VAL = 0
            !ISSUE_QTY = 0
            !ISSUE_VAL = 0
            !CLOSE_QTY = M_STOCK
            !CLOSE_VAL = 0
            !DAM_QTY = 0
            !DAM_VAL = 0
            RSTITEMMAST.Update
        End If
    End With
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Exit Function
    
eRRhAND:
    MsgBox Err.Description
End Function

Private Sub FILLCOMBO()
    On Error GoTo eRRhAND
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If ACT_FLAG = True Then
        ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='131')And (len(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    Else
        ACT_REC.Close
        ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='131')And (len(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    End If
    
    Set Me.CMBDISTI.RowSource = ACT_REC
    CMBDISTI.ListField = "ACT_NAME"
    CMBDISTI.BoundColumn = "ACT_CODE"
    Screen.MousePointer = vbNormal
    Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub TXTTOTALDISC_LostFocus()
    lblnetamount.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        lblnetamount.Caption = Val(lblnetamount.Caption) + Val(grdsales.TextMatrix(i, 11))
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = Round((Val(LBLTOTAL.Caption) * Val(TXTTOTALDISC.Text) / 100), 2)
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2)
    LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - Val(LBLTOTALCOST.Caption), 2)
    
End Sub

Private Function COSTCALCULATION()
    Dim RSTCOST As ADODB.Recordset
    Dim COST As Double
    Dim N As Integer
    'Dim RSTITEMMAST As ADODB.Recordset
    
     LBLTOTALCOST.Caption = ""
     LBLPROFIT.Caption = ""
    COST = 0
    On Error GoTo eRRhAND
    For N = 1 To grdsales.Rows - 1
        Set RSTCOST = New ADODB.Recordset
        RSTCOST.Open "SELECT [ITEM_COST] FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(N, 15)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(N, 13)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(N, 14)) & "", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTCOST.EOF
            COST = COST + (RSTCOST!ITEM_COST) * Val(grdsales.TextMatrix(N, 3))
            RSTCOST.MoveNext
        Loop
        RSTCOST.Close
        Set RSTCOST = Nothing
    Next N
    
    LBLTOTALCOST.Caption = Round(COST, 2)
    LBLPROFIT.Caption = Format(Round(Val(lblnetamount.Caption) - COST, 2), "0.00")

    Exit Function
    
eRRhAND:
    MsgBox Err.Description
End Function

Private Sub TXTDEALER_Change()
    On Error GoTo eRRhAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='131')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='131')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
eRRhAND:
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
    
    On Error GoTo eRRhAND

    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select ADDRESS, DL_NO, KGST from [ACTMAST]  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
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
    cmbinv.Text = ""
    If TXTDEALER.Top = 225 Then
        grdsales.FixedRows = 0
        grdsales.Rows = 1
    End If
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
End Sub


Private Function FILLINVOICE()
    On Error GoTo eRRhAND
    
    Screen.MousePointer = vbHourglass
    Set cmbinv.DataSource = Nothing
    If INV_FLAG = True Then
        INV_REC.Open "Select DISTINCT VCH_NO From TEMPCN WHERE ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_NO", db2, adOpenStatic, adLockReadOnly
        INV_FLAG = False
    Else
        INV_REC.Close
        INV_REC.Open "Select DISTINCT VCH_NO From TEMPCN WHERE ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_NO", db2, adOpenStatic, adLockReadOnly
        INV_FLAG = False
    End If
    
    Set Me.cmbinv.RowSource = INV_REC
    cmbinv.ListField = "VCH_NO"
    cmbinv.BoundColumn = "VCH_NO"
    Screen.MousePointer = vbNormal
    Exit Function

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Sub TXTPRODUCT_GotFocus()
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TXTQTY_GotFocus()

    Dim RSTITEMCOST As ADODB.Recordset
    
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
    TXTQTY.Tag = Trim(TXTPRODUCT.Text)
    On Error GoTo eRRhAND
    
    Set RSTITEMCOST = New ADODB.Recordset
    RSTITEMCOST.Open "SELECT ITEM_COST, SALES_PRICE FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMCOST.EOF Or RSTITEMCOST.BOF) Then
        LBLITEMCOST.Caption = RSTITEMCOST!ITEM_COST
        LBLSELPRICE.Caption = RSTITEMCOST!SALES_PRICE
    End If
    RSTITEMCOST.Close
    Set RSTITEMCOST = Nothing
    
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub


Private Sub CMDHIDE_Click()
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

Private Function VIEWGRID()
    Dim TRXMAST As ADODB.Recordset
    Dim RSTDN As ADODB.Recordset
    
    Dim E_BILL As String
    Dim i As Integer
    'On Error GoTo eRRhAND
            If Val(txtBillNo.Text) = 0 Then Exit Function
            grdsales.Rows = 1
            i = 0
            Set RSTDN = New ADODB.Recordset
            RSTDN.Open "Select * From TEMPCN WHERE VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db2, adOpenStatic, adLockReadOnly
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
                TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMAST WHERE ITEMMAST.ITEM_CODE = '" & Trim(RSTDN!ITEM_CODE) & "'", db, adOpenStatic, adLockReadOnly
                If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                    grdsales.TextMatrix(i, 17) = Trim(TRXMAST!MANUFACTURER)
                End If
                TRXMAST.Close
                Set TRXMAST = Nothing
                
                grdsales.TextMatrix(i, 5) = Format(RSTDN!MRP, ".000")
                grdsales.TextMatrix(i, 6) = Format(RSTDN!SALES_PRICE, ".000")
                grdsales.TextMatrix(i, 7) = 0 'DISC
                grdsales.TextMatrix(i, 8) = Val(RSTDN!SALES_TAX)
                grdsales.TextMatrix(i, 9) = RSTDN!REF_NO
                grdsales.TextMatrix(i, 10) = Format(RSTDN!EXP_DATE, "MM/YY")
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
            TXTAMOUNT.Text = Round((Val(LBLTOTAL.Caption) * Val(TXTTOTALDISC.Text) / 100), 2)
            lblnetamount.Caption = Format(Round(Val(LBLTOTAL.Caption) - Val(TXTAMOUNT.Text), 2), "0.00")
            Call COSTCALCULATION
            
            TXTSLNO.Text = grdsales.Rows

End Function
