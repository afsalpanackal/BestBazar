VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMPURCHASERET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE RETURN (DEBIT NOTE)"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18195
   Icon            =   "FrmPurchRet.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   18195
   Visible         =   0   'False
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   3930
      TabIndex        =   38
      Top             =   3255
      Visible         =   0   'False
      Width           =   10950
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   3165
         Left            =   15
         TabIndex        =   39
         Top             =   390
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5583
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
         Left            =   3825
         TabIndex        =   41
         Top             =   15
         Width           =   7125
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   "PURCHASE HISTORY FOR"
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
         TabIndex        =   40
         Top             =   15
         Width           =   3810
      End
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   1725
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   8340
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2835
         Left            =   30
         TabIndex        =   16
         Top             =   360
         Width           =   8250
         _ExtentX        =   14552
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   30
         Visible         =   0   'False
         Width           =   5205
      End
   End
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   1725
      TabIndex        =   17
      Top             =   3570
      Visible         =   0   'False
      Width           =   10965
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   3165
         Left            =   45
         TabIndex        =   18
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
      Height          =   4140
      Left            =   1710
      TabIndex        =   71
      Top             =   2670
      Visible         =   0   'False
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   7303
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
      Left            =   1020
      TabIndex        =   70
      Top             =   120
      Width           =   885
   End
   Begin VB.Frame FRMEMAIN 
      BorderStyle     =   0  'None
      Height          =   10200
      Left            =   -135
      TabIndex        =   5
      Top             =   0
      Width           =   18330
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
         TabIndex        =   19
         Top             =   10635
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame FRMEHEAD 
         BackColor       =   &H00DDDCF1&
         ForeColor       =   &H008080FF&
         Height          =   2175
         Left            =   210
         TabIndex        =   7
         Top             =   -30
         Width           =   18120
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
            Left            =   10980
            TabIndex        =   175
            Top             =   180
            Width           =   1410
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
            Left            =   12450
            TabIndex        =   174
            Top             =   180
            Width           =   1410
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   12495
            MaxLength       =   300
            TabIndex        =   168
            Top             =   1680
            Visible         =   0   'False
            Width           =   1380
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5850
            Locked          =   -1  'True
            MaxLength       =   200
            TabIndex        =   167
            Top             =   1680
            Width           =   6045
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00DDDCF1&
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
            ForeColor       =   &H000000C0&
            Height          =   1515
            Left            =   4995
            TabIndex        =   61
            Top             =   90
            Width           =   4890
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
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   960
               MaxLength       =   35
               TabIndex        =   62
               Top             =   1125
               Width           =   3180
            End
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
               ForeColor       =   &H80000008&
               Height          =   885
               Left            =   45
               TabIndex        =   64
               Top             =   210
               Width           =   4800
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
               Left            =   60
               TabIndex        =   63
               Top             =   1140
               Width           =   885
            End
         End
         Begin VB.TextBox TXTDEALER 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   945
            TabIndex        =   1
            Top             =   570
            Width           =   3990
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   300
            Left            =   2445
            TabIndex        =   0
            Top             =   165
            Width           =   1395
            _ExtentX        =   2461
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
            Height          =   1140
            Left            =   945
            TabIndex        =   2
            Top             =   975
            Width           =   3990
            _ExtentX        =   7038
            _ExtentY        =   2011
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
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
            Index           =   32
            Left            =   11970
            TabIndex        =   169
            Top             =   1710
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Details"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   5040
            TabIndex        =   166
            Top             =   1620
            Width           =   765
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "SUPLIER"
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
            Left            =   105
            TabIndex        =   21
            Top             =   630
            Width           =   855
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
            Left            =   1890
            TabIndex        =   20
            Top             =   180
            Width           =   630
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "PR. NO."
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
            TabIndex        =   11
            Top             =   180
            Width           =   780
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
            TabIndex        =   10
            Top             =   1485
            Visible         =   0   'False
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
            Height          =   315
            Left            =   1305
            TabIndex        =   9
            Top             =   930
            Visible         =   0   'False
            Width           =   1215
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
            Height          =   315
            Left            =   945
            TabIndex        =   8
            Top             =   150
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DDDCF1&
         ForeColor       =   &H008080FF&
         Height          =   4710
         Left            =   210
         TabIndex        =   12
         Top             =   2055
         Width           =   18120
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
            TabIndex        =   163
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
            TabIndex        =   29
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
            Height          =   345
            Left            =   12210
            TabIndex        =   28
            Top             =   4260
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
            Left            =   10125
            TabIndex        =   27
            Top             =   4260
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
            Left            =   11070
            TabIndex        =   26
            Top             =   4260
            Width           =   1125
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4140
            Left            =   30
            TabIndex        =   6
            Top             =   120
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   7303
            _Version        =   393216
            Rows            =   1
            Cols            =   36
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   450
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00DDDCF1&
            Height          =   4695
            Left            =   14820
            TabIndex        =   42
            Top             =   0
            Width           =   3285
            Begin VB.TextBox Txthandle 
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
               Left            =   1980
               TabIndex        =   75
               Top             =   3465
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
               TabIndex        =   74
               Text            =   "Handling Charge"
               Top             =   3510
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
               TabIndex        =   73
               Text            =   "Frieght Charge"
               Top             =   3885
               Width           =   1875
            End
            Begin VB.TextBox TxtFrieght 
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
               Left            =   1980
               TabIndex        =   72
               Top             =   3855
               Width           =   1230
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "COMM AMT"
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
               Index           =   53
               Left            =   120
               TabIndex        =   162
               Top             =   2775
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
               Left            =   195
               TabIndex        =   161
               Top             =   3675
               Visible         =   0   'False
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
               Height          =   450
               Left            =   195
               TabIndex        =   151
               Top             =   2340
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
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
               Left            =   120
               TabIndex        =   150
               Top             =   2115
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00008000&
               Height          =   450
               Index           =   45
               Left            =   1770
               TabIndex        =   69
               Top             =   2115
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
               Height          =   450
               Left            =   1770
               TabIndex        =   68
               Top             =   2340
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   44
               Left            =   1770
               TabIndex        =   67
               Top             =   2775
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
               Height          =   450
               Left            =   1770
               TabIndex        =   66
               Top             =   3000
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
               Height          =   450
               Left            =   195
               TabIndex        =   60
               Top             =   330
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   6
               Left            =   180
               TabIndex        =   59
               Top             =   105
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
               Height          =   450
               Left            =   195
               TabIndex        =   58
               Top             =   990
               Width           =   1440
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   23
               Left            =   195
               TabIndex        =   57
               Top             =   780
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
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
               Left            =   1785
               TabIndex        =   56
               Top             =   105
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
               Height          =   450
               Left            =   1785
               TabIndex        =   55
               Top             =   990
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
               Height          =   450
               Left            =   1785
               TabIndex        =   54
               Top             =   1680
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
               TabIndex        =   53
               Top             =   4605
               Visible         =   0   'False
               Width           =   1440
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   27
               Left            =   1785
               TabIndex        =   52
               Top             =   1455
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
               TabIndex        =   51
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
               TabIndex        =   50
               Top             =   4050
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
               TabIndex        =   49
               Top             =   4350
               Visible         =   0   'False
               Width           =   1455
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
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   4
               Left            =   195
               TabIndex        =   48
               Top             =   1455
               Width           =   1440
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
               Height          =   450
               Left            =   195
               TabIndex        =   47
               Top             =   1680
               Width           =   1440
            End
            Begin VB.Label lblcomamt 
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
               Height          =   495
               Left            =   195
               TabIndex        =   46
               Top             =   2985
               Width           =   1440
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
               TabIndex        =   45
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
               Height          =   450
               Left            =   1785
               TabIndex        =   44
               Top             =   330
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
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
               Left            =   1785
               TabIndex        =   43
               Top             =   780
               Visible         =   0   'False
               Width           =   1425
            End
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
            Height          =   300
            Index           =   52
            Left            =   6750
            TabIndex        =   155
            Top             =   4305
            Width           =   405
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
            TabIndex        =   154
            Top             =   4335
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblcase 
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
            Left            =   7155
            TabIndex        =   153
            Top             =   4305
            Width           =   840
         End
         Begin VB.Label lblcrtnpack 
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
            Left            =   6120
            TabIndex        =   152
            Top             =   4305
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L.Price"
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
            Index           =   42
            Left            =   5280
            TabIndex        =   65
            Top             =   4305
            Width           =   840
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
            TabIndex        =   37
            Top             =   4920
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label lblvan 
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
            Left            =   4050
            TabIndex        =   36
            Top             =   4305
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
            Left            =   2325
            TabIndex        =   35
            Top             =   4305
            Width           =   855
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
            Left            =   720
            TabIndex        =   34
            Top             =   4305
            Width           =   855
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
            TabIndex        =   33
            Top             =   4920
            Visible         =   0   'False
            Width           =   600
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
            Height          =   300
            Index           =   30
            Left            =   3240
            TabIndex        =   32
            Top             =   4305
            Width           =   795
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
            Left            =   1635
            TabIndex        =   31
            Top             =   4305
            Width           =   690
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
            Left            =   90
            TabIndex        =   30
            Top             =   4305
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DDDCF1&
         ForeColor       =   &H008080FF&
         Height          =   3540
         Left            =   210
         TabIndex        =   76
         Top             =   6675
         Width           =   18090
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
            Left            =   6000
            TabIndex        =   173
            Top             =   1830
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox txtinvnodate 
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
            Height          =   450
            Left            =   9990
            MaxLength       =   60
            TabIndex        =   113
            Top             =   375
            Width           =   1290
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   450
            Left            =   8475
            TabIndex        =   111
            Top             =   375
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   794
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox txtcommi 
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
            Left            =   13740
            MaxLength       =   6
            TabIndex        =   159
            Top             =   1680
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox TxtName1 
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
            Left            =   1845
            MaxLength       =   15
            TabIndex        =   106
            Top             =   390
            Width           =   1080
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
            TabIndex        =   148
            Top             =   3315
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
            Left            =   5745
            TabIndex        =   100
            Top             =   930
            Width           =   1095
         End
         Begin VB.CommandButton CmdPrintA5 
            Caption         =   "PRINT -&A5"
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
            Left            =   6885
            TabIndex        =   101
            Top             =   930
            Width           =   1110
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
            Height          =   450
            Left            =   12285
            MaxLength       =   30
            TabIndex        =   124
            Top             =   3675
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox TxtWarranty_type 
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
            Height          =   450
            Left            =   14535
            MaxLength       =   30
            TabIndex        =   123
            Top             =   4080
            Visible         =   0   'False
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
            Left            =   9150
            TabIndex        =   102
            Top             =   930
            Width           =   990
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
            TabIndex        =   121
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
            TabIndex        =   119
            Top             =   3405
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
            TabIndex        =   117
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
            TabIndex        =   115
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   450
            Left            =   7605
            MaxLength       =   30
            TabIndex        =   110
            Top             =   375
            Width           =   855
         End
         Begin VB.TextBox TXTITEMCODE 
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
            Height          =   435
            Left            =   465
            TabIndex        =   105
            Top             =   390
            Width           =   1365
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
            Left            =   3450
            TabIndex        =   98
            Top             =   930
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
            Left            =   4605
            TabIndex        =   99
            Top             =   930
            Width           =   1095
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
            Left            =   10170
            TabIndex        =   103
            Top             =   930
            Width           =   1020
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
            Height          =   450
            Left            =   15960
            MaxLength       =   5
            TabIndex        =   122
            Top             =   375
            Width           =   585
         End
         Begin VB.TextBox TXTTAX 
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
            Left            =   13245
            MaxLength       =   4
            TabIndex        =   116
            Top             =   375
            Width           =   540
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
            Height          =   450
            Left            =   9375
            MaxLength       =   8
            TabIndex        =   112
            Top             =   375
            Width           =   600
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
            Height          =   435
            Left            =   2940
            TabIndex        =   107
            Top             =   390
            Width           =   3735
         End
         Begin VB.TextBox TXTSLNO 
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
            Height          =   450
            Left            =   30
            TabIndex        =   104
            Top             =   375
            Width           =   420
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
            Left            =   2265
            TabIndex        =   97
            Top             =   915
            Width           =   1125
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
            Height          =   450
            Left            =   12525
            MaxLength       =   6
            TabIndex        =   114
            Top             =   375
            Width           =   705
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
            Left            =   11220
            TabIndex        =   96
            Top             =   930
            Width           =   375
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
            TabIndex        =   95
            Top             =   1305
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
            TabIndex        =   94
            Top             =   1305
            Width           =   1380
         End
         Begin VB.TextBox TXTFREE 
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
            Left            =   9675
            MaxLength       =   7
            TabIndex        =   93
            Top             =   3555
            Visible         =   0   'False
            Width           =   570
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
            TabIndex        =   92
            Top             =   4290
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00DDDCF1&
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
            Left            =   11625
            TabIndex        =   91
            Top             =   1140
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00DDDCF1&
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
            Left            =   11610
            TabIndex        =   90
            Top             =   855
            Width           =   975
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
            Left            =   2115
            MaxLength       =   6
            TabIndex        =   89
            Top             =   1755
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.OptionButton optnet 
            BackColor       =   &H00DDDCF1&
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
            Left            =   12585
            TabIndex        =   88
            Top             =   870
            Width           =   975
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
            Height          =   450
            Left            =   13800
            MaxLength       =   9
            TabIndex        =   118
            Top             =   375
            Width           =   1065
         End
         Begin VB.TextBox txtretail 
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
            Left            =   14880
            MaxLength       =   9
            TabIndex        =   120
            Top             =   375
            Width           =   1065
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
            TabIndex        =   87
            Top             =   4035
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtretaildummy 
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
            Left            =   4065
            MaxLength       =   6
            TabIndex        =   86
            Top             =   3285
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
            TabIndex        =   85
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
            Left            =   4005
            MaxLength       =   15
            TabIndex        =   84
            Top             =   1800
            Visible         =   0   'False
            Width           =   1425
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
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
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
            Left            =   14850
            TabIndex        =   79
            Top             =   930
            Width           =   1100
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
            TabIndex        =   78
            Top             =   1740
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.CheckBox Chkcancel 
            Appearance      =   0  'Flat
            BackColor       =   &H00DDDCF1&
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
            Height          =   405
            Left            =   13590
            TabIndex        =   77
            Top             =   945
            Width           =   1230
         End
         Begin VB.TextBox LblPack 
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
            Height          =   450
            Left            =   6690
            MaxLength       =   8
            TabIndex        =   108
            Top             =   375
            Width           =   480
         End
         Begin VB.Frame FrmeType 
            BackColor       =   &H0080C0FF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   720
            Left            =   8010
            TabIndex        =   156
            Top             =   750
            Width           =   1140
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
               Left            =   45
               TabIndex        =   158
               Top             =   135
               Value           =   -1  'True
               Width           =   1050
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
               TabIndex        =   157
               Top             =   405
               Width           =   1050
            End
         End
         Begin MSMask.MaskEdBox TxtInvoiceDate 
            Height          =   450
            Left            =   11295
            TabIndex        =   171
            Top             =   375
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   794
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
            Caption         =   "Inv. Date"
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
            Index           =   39
            Left            =   11295
            TabIndex        =   172
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Invoice No."
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
            Index           =   35
            Left            =   9990
            TabIndex        =   170
            Top             =   150
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Expiry"
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
            Height          =   210
            Index           =   61
            Left            =   8475
            TabIndex        =   165
            Top             =   165
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Free"
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
            Index           =   55
            Left            =   9675
            TabIndex        =   164
            Top             =   3330
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Comm"
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
            Index           =   46
            Left            =   13740
            TabIndex        =   160
            Top             =   1455
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label LBLTYPE 
            Caption         =   "RI"
            Height          =   330
            Left            =   2805
            TabIndex        =   149
            Top             =   915
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
            TabIndex        =   3
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
            TabIndex        =   4
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
            TabIndex        =   147
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
            TabIndex        =   146
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
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   16560
            TabIndex        =   145
            Top             =   375
            Width           =   1485
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
            Left            =   7605
            TabIndex        =   144
            Top             =   150
            Width           =   855
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
            TabIndex        =   143
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   14
            Left            =   16560
            TabIndex        =   142
            Top             =   150
            Width           =   1485
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
            Height          =   225
            Index           =   13
            Left            =   15960
            TabIndex        =   141
            Top             =   150
            Width           =   585
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
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   12
            Left            =   13245
            TabIndex        =   140
            Top             =   150
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Net Rate"
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
            Left            =   14880
            TabIndex        =   139
            Top             =   150
            Width           =   1065
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
            Left            =   9375
            TabIndex        =   138
            Top             =   150
            Width           =   600
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
            Left            =   1845
            TabIndex        =   137
            Top             =   150
            Width           =   4830
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
            Height          =   225
            Index           =   8
            Left            =   30
            TabIndex        =   136
            Top             =   150
            Width           =   420
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
            Left            =   12525
            TabIndex        =   135
            Top             =   150
            Width           =   705
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
            TabIndex        =   134
            Top             =   1380
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Free"
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
            Index           =   29
            Left            =   9645
            TabIndex        =   133
            Top             =   3735
            Visible         =   0   'False
            Width           =   600
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
            Height          =   225
            Index           =   30
            Left            =   13800
            TabIndex        =   132
            Top             =   150
            Width           =   1065
         End
         Begin VB.Label lblP_Rate 
            Caption         =   "0"
            Height          =   390
            Left            =   13200
            TabIndex        =   131
            Top             =   3690
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Commi"
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
            Index           =   5
            Left            =   10965
            TabIndex        =   130
            Top             =   3540
            Visible         =   0   'False
            Width           =   915
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
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   43
            Left            =   510
            TabIndex        =   129
            Top             =   915
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Warranty"
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
            Index           =   48
            Left            =   14250
            TabIndex        =   128
            Top             =   3180
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label lblunit 
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
            Left            =   7185
            TabIndex        =   109
            Top             =   375
            Width           =   405
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Unit"
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
            Left            =   7185
            TabIndex        =   127
            Top             =   150
            Width           =   405
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
            Index           =   37
            Left            =   6690
            TabIndex        =   126
            Top             =   150
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Item Code"
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
            Index           =   38
            Left            =   465
            TabIndex        =   125
            Top             =   150
            Width           =   1365
         End
      End
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   13155
      TabIndex        =   22
      Top             =   3090
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1482
      _Version        =   393216
   End
   Begin VB.Label lblcredit 
      Height          =   690
      Left            =   -15
      TabIndex        =   25
      Top             =   -225
      Width           =   915
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   11340
      TabIndex        =   24
      Top             =   1065
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   11565
      TabIndex        =   23
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "FRMPURCHASERET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim PHYCODE As New ADODB.Recordset
Dim PHYCODEFLAG As Boolean
Dim TMPREC As New ADODB.Recordset
Dim TMPFLAG As Boolean
Dim ACT_REC As New ADODB.Recordset
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
Dim CHANGE_NAME As Boolean

Private Sub cmdadd_GotFocus()
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    Call TXTDISC_LostFocus
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
            MsgBox "Permission Denied", vbOKOnly, "Sales"
            Exit Sub
        End If
        If MsgBox("ARE YOU SURE YOU WANT TO CANCEL THE BILL!!!!!", vbYesNo + vbDefaultButton2, "DELETE!!!") = vbNo Then
            Chkcancel.Value = 0
            Exit Sub
        End If
    End If
    
    'db.Execute "delete From RTRXFILE WHERE TRX_TYPE='CN' AND VCH_NO = " & Val(TxtCN.Text) & ""
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM TEMPCN WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "' AND BILL_NO = " & Val(txtBillNo.Text) & " AND BILL_TRX_TYPE = 'PR' AND TRX_TYPE = 'SI'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!check_flag = "N"
        RSTTRXFILE!BILL_NO = 0 'Null
        RSTTRXFILE!BILL_TRX_TYPE = Null
        'RSTTRXFILE!BILL_DATE = Null
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & " "
    db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & " "
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE TRX_TYPE='PR' AND VCH_NO = " & Val(TxtCN.Text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
    With RSTTRXFILE
        Do Until .EOF
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With rststock
                If Not (.EOF And .BOF) Then
                    '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20)))
                    !ISSUE_QTY = !ISSUE_QTY + RSTTRXFILE!QTY
                    If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                    !ISSUE_VAL = !ISSUE_VAL + RSTTRXFILE!TRX_TOTAL
                    !CLOSE_QTY = !CLOSE_QTY - RSTTRXFILE!QTY
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    !CLOSE_VAL = !CLOSE_VAL - RSTTRXFILE!TRX_TOTAL
                    rststock.Update
                End If
            End With
            rststock.Close
            Set rststock = Nothing
        RSTTRXFILE.MoveNext
        Loop
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    db.Execute "delete From RTRXFILE WHERE TRX_TYPE='PR' AND VCH_NO = " & Val(TxtCN.Text) & ""
    For n = 1 To grdsales.rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(n, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(N, 3)) + Val(grdsales.TextMatrix(N, 20)))
                !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(n, 3)) * Val(grdsales.TextMatrix(n, 27)), 3)
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(n, 20)) * Val(grdsales.TextMatrix(n, 27)), 3)
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(n, 12))
                !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20))) * Val(grdsales.TextMatrix(n, 27)), 3)
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(n, 12))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
           
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(n, 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(n, 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(n, 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(n, 35)) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20))) * Val(grdsales.TextMatrix(n, 27)), 3)
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(n, 3)) + Val(grdsales.TextMatrix(n, 20))) * Val(grdsales.TextMatrix(n, 27)), 3)
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    Next n
    grdsales.FixedRows = 0
    grdsales.rows = 1
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    Call AppendSale
    Chkcancel.Value = 0
End Sub

Private Sub CMDPre_Click()
    Dim currbillno As Long
    currbillno = Val(txtBillNo.Text)
    On Error GoTo ERRHAND
    If CMDEXIT.Enabled = False Then cmdRefresh_Click
    If currbillno <= 1 Then
        txtBillNo.Text = 1
    Else
        txtBillNo.Text = currbillno - 1
    End If
    LBLBILLNO.Caption = Val(txtBillNo.Text)
    TXTDEALER.Text = ""
    DataList2.BoundText = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    LBLRETAMT.Caption = ""
    TXTTIN.Text = ""
    TxtCN.Text = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    Chkcancel.Value = 0
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    OLD_BILL = False
    Call txtBillNo_KeyDown(13, 0)
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdnxt_Click()
        
    Dim currbillno As Long
    currbillno = Val(txtBillNo.Text)
    If CMDEXIT.Enabled = False Then cmdRefresh_Click
    
    Dim rstBILL As ADODB.Recordset
    Dim lastbillno As Long
    
    On Error GoTo ERRHAND
    lbldealer.Caption = ""
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    DataList2.BoundText = ""
    
    LBLRETAMT.Caption = ""
    TXTTIN.Text = ""
    
    TxtCN.Text = ""
    TXTTOTALDISC.Text = ""
    
    LBLTOTALCOST.Caption = ""
    
    
    Chkcancel.Value = 0
    OLD_BILL = False
    
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PR'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        lastbillno = IIf(IsNull(rstBILL.Fields(0)), 0, rstBILL.Fields(0))
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    If currbillno > lastbillno Then Exit Sub
    txtBillNo.Text = currbillno + 1
    LBLBILLNO.Caption = Val(txtBillNo.Text)
    
    Call txtBillNo_KeyDown(13, 0)
    
    
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CmdPrintA5_Click()
    
    Chkcancel.Value = 0
    If grdsales.rows = 1 Then Exit Sub
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    'If Month(Date) >= 4 And Year(Date) >= 2021 Then Exit Sub
'    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then
'        'db.Execute "delete From USERS "
'        Exit Sub
'    End If
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) FROM PURCAHSERETURN", db, adOpenForwardOnly
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
            
'    If Trim(TXTAREA.Text) = "" Then
'        MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTAREA.SetFocus
'        Exit Sub
'    End If
    
    'If Val(txtcrdays.Text) = 0 Or DataList2.BoundText = "130000" Then
    Small_Print = True
    Dos_Print = False
    Me.lblcredit.Caption = "1"
    Me.Generateprint
End Sub

Private Sub CmdPrintA5_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.rows
            'TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.Value = True
            TxtMRP.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TrxRYear.Text = ""
            TXTUNIT.Text = ""
            
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
            TXTPRODUCT.Enabled = False
            TxtName1.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub cmdreturn_Click()
    If DataList2.BoundText = "" Then Exit Sub
    Set creditbill = Me
    Enabled = False
    M_ADD = True
    MDIMAIN.Enabled = False
    FRMCRDTNOTE.LBLCUSTOMER.Caption = DataList2.BoundText
    FRMCRDTNOTE.Tag = "Y"
    FRMCRDTNOTE.Show
End Sub

Private Sub DataList2_Click()
    Dim rstCustomer As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    On Error GoTo ERRHAND

    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from ACTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = Trim(rstCustomer!Address) & IIf(IsNull(rstCustomer!TELNO) Or rstCustomer!TELNO = "", "", ". Ph: " & Trim(rstCustomer!TELNO))
        'If TxtBillName.Text = "" Then TxtBillName.Text = DataList2.Text
        TXTTIN.Text = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
        'TxtPhone.Text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO))
    Else
        'TxtPhone.Text = ""
        TXTTIN.Text = ""
        lbladdress.Caption = ""
        'lblcusttype.Caption = "R"
    End If

    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text
    Exit Sub
    
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
            'FRMEHEAD.Enabled = False
            'TxtName1.Enabled = True
            'TxtName1.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMDADD_Click()
    Dim rststock As ADODB.Recordset
    'Dim RSTMINQTY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim i As Long
    
    Chkcancel.Value = 0
    On Error GoTo ERRHAND
'   ' 'If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then Exit Sub
    
    If OLD_BILL = False Then Call checklastbill
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "PR"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.Text
        RSTTRXFILE!BILL_NAME = Trim(TXTDEALER.Text)
        RSTTRXFILE!BILL_ADDRESS = Trim(lbladdress.Caption)
        RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
        RSTTRXFILE!ADD_AMOUNT = 0
        RSTTRXFILE!ROUNDED_OFF = 0
        RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    If grdsales.rows <= Val(TXTSLNO.Text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text) + Val(TXTAPPENDQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Format(Val(TxtMRP.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTRETAILNOTAX.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Format(Val(txtretail.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(TXTDISC.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Val(TXTTAX.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Val(LBLITEMCOST.Caption)
    
    TXTDISC.Tag = 0
    If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
        TXTAPPENDTOTAL.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
    Else
        TXTDISC.Tag = Val(TXTAPPENDQTY.Text) * Val(txtretail.Text) * Val(TXTDISC.Text) / 100
        TXTAPPENDTOTAL.Text = Format((Val(TXTAPPENDQTY.Text) * Round(Val(txtretail.Text), 3)) - Val(TXTDISC.Tag), ".000")
    End If
    
    TXTAPPENDTOTAL.Text = ""
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Format(Val(LBLSUBTOTAL.Caption) + Val(TXTAPPENDTOTAL.Text), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Trim(TXTVCHNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Trim(TXTLINENO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Trim(TXTTRXTYPE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 35) = Trim(TrxRYear.Text)
    
    If OPTVAT.Value = True And Val(TXTTAX.Text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "V"
    If OPTTaxMRP.Value = True And Val(TXTTAX.Text) > 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "M"
    If Val(TXTTAX.Text) <= 0 Or optnet.Value = True Then grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "N"
    
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = "N"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = IIf(IsNull(RSTTRXFILE!MANUFACTURER), "", Trim(RSTTRXFILE!MANUFACTURER))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Select Case LBLDNORCN.Caption
        Case "DN"
            grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = "DN"
        Case "CN"
            grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = "CN"
        Case Else
            grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = "B"
    End Select
    grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = Val(TXTFREE.Text) + Val(TXTFREEAPPEND.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 21) = Format(Val(txtretail.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 22) = Format(Val(TXTRETAILNOTAX.Text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 23) = Trim(TXTSALETYPE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = Val(txtcommi.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = Trim(txtcategory.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 26) = "L"
    grdsales.TextMatrix(Val(TXTSLNO.Text), 27) = IIf(Val(LblPack.Text) = 0, "1", Val(LblPack.Text))
    grdsales.TextMatrix(Val(TXTSLNO.Text), 28) = Val(TxtWarranty.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 29) = Trim(TxtWarranty_type.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 30) = Trim(lblunit.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 31) = IIf(TXTEXPIRY.Text = "  /  ", "", Trim(TXTEXPIRY.Text))
    grdsales.TextMatrix(Val(TXTSLNO.Text), 33) = Trim(txtinvnodate.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 34) = IIf(IsDate(TxtInvoiceDate.Text), TxtInvoiceDate.Text, "")
    If M_EDIT = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32))
    Else
        i = 0
        Dim rstMaxNo As ADODB.Recordset
        Set rstMaxNo = New ADODB.Recordset
        rstMaxNo.Open "Select MAX(LINE_NO) From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
        If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
            grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        Else
            grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = Val(TXTSLNO.Text)
        End If
        rstMaxNo.Close
        Set rstMaxNo = Nothing
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "PR"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32))
    End If
    RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = "", 0, grdsales.TextMatrix(Val(TXTSLNO.Text), 14))
    RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = "", 0, grdsales.TextMatrix(Val(TXTSLNO.Text), 15))
    RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = "", "MI", grdsales.TextMatrix(Val(TXTSLNO.Text), 16))
    RSTTRXFILE!R_TRX_YEAR = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 35) = "", "", grdsales.TextMatrix(Val(TXTSLNO.Text), 35))
    RSTTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "PR"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32))
    End If
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    If UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICE CHARGE" Then
        RSTTRXFILE!Category = "SERVICE CHARGE"
    Else
        RSTTRXFILE!Category = "General"
    End If
    RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
    RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
    RSTTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
    RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
    RSTTRXFILE!MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
    RSTTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
    RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
    RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21))
    RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22))
    RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
    RSTTRXFILE!Category = grdsales.TextMatrix(Val(TXTSLNO.Text), 25)
    RSTTRXFILE!COM_FLAG = "N"
    RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(Val(TXTSLNO.Text), 26)
    RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 27))
    RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
    RSTTRXFILE!UNIT = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
    RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.Text), 1, 30)
    RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
    RSTTRXFILE!ISSUE_QTY = 0
    RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
    RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
    Select Case grdsales.TextMatrix(Val(TXTSLNO.Text), 19)
        Case "DN"
            RSTTRXFILE!CST = 1
        Case "CN"
            RSTTRXFILE!CST = 2
        Case Else
            RSTTRXFILE!CST = 0
    End Select
    RSTTRXFILE!BAL_QTY = 0
    RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(Val(TXTSLNO.Text), 12)
    RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
    RSTTRXFILE!SCHEME = (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
    'RSTTRXFILE!EXP_DATE = Null
    RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!M_USER_ID = DataList2.BoundText
    RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23))
    RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 28) = "", Null, grdsales.TextMatrix(Val(TXTSLNO.Text), 28))
    RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(Val(TXTSLNO.Text), 29)
    RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(Val(TXTSLNO.Text), 30)
    If grdsales.TextMatrix(Val(TXTSLNO.Text), 31) = "" Then
        'RSTTRXFILE!EXP_DATE = Null
    Else
        RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 31))) & "/" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 31))
    End If
    RSTTRXFILE!INV_DETAILS = grdsales.TextMatrix(Val(TXTSLNO.Text), 33)
    RSTTRXFILE!INV_DATE = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 34)), Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 34), "dd/mm/yyyy"), "")
    RSTTRXFILE.Update

    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 35)) & "' AND BAL_QTY > 0", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text), 3)
            !BAL_QTY = !BAL_QTY - Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text), 3)
            grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(LblPack.Text))
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        Else
            'BALQTY = 0
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "' AND BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!ISSUE_QTY)) Then RSTTRXFILE!ISSUE_QTY = 0
                If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                'BALQTY = RSTTRXFILE!BAL_QTY
                RSTTRXFILE!ISSUE_QTY = RSTTRXFILE!ISSUE_QTY + Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text), 3)
                RSTTRXFILE!BAL_QTY = RSTTRXFILE!BAL_QTY - Round((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text), 3)
                
                grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = RSTTRXFILE!VCH_NO
                grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = RSTTRXFILE!LINE_NO
                grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = RSTTRXFILE!TRX_TYPE
                grdsales.TextMatrix(Val(TXTSLNO.Text), 35) = RSTTRXFILE!TRX_YEAR
                grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(LblPack.Text))
                RSTTRXFILE.Update
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            Else
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = IIf(IsNull(RSTTRXFILE!ITEM_COST), "", RSTTRXFILE!ITEM_COST * Val(LblPack.Text))
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            End If
        End If
    End With
    
'    Dim RET_PRICE, LOOS_PRICE, LOOSE_PCK, ITEM_CST As Double
'    Set RSTTRXFILE = Nothing
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "' AND  BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockReadOnly
'    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'        If Not (IsNull(RSTTRXFILE!P_RETAIL) Or RSTTRXFILE!P_RETAIL = 0) Then
'            RET_PRICE = IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL)
'            LOOS_PRICE = IIf(IsNull(RSTTRXFILE!P_CRTN), 0, RSTTRXFILE!P_CRTN)
'            LOOSE_PCK = IIf(IsNull(RSTTRXFILE!LOOSE_PACK), 0, RSTTRXFILE!LOOSE_PACK)
'            ITEM_CST = IIf(IsNull(RSTTRXFILE!ITEM_COST), 0, RSTTRXFILE!ITEM_COST)
'        End If
'    End If
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
'            If RET_PRICE > 0 Then
'                !P_RETAIL = RET_PRICE
'                !P_CRTN = LOOS_PRICE
'                !LOOSE_PACK = LOOSE_PCK
'                If ITEM_CST > 0 Then !ITEM_COST = ITEM_CST
'            End If
            '!ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))
            If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY + Round((Val(TXTQTY.Text) * Val(LblPack.Text)), 3)
            !FREE_QTY = !FREE_QTY + Round((Val(TXTFREE.Text) * Val(LblPack.Text)), 3)
            !CLOSE_QTY = !CLOSE_QTY - Round(((Val(TXTQTY.Text) + Val(TXTFREE.Text)) * Val(LblPack.Text)), 3)

            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.Caption = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
        Else
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
        End If
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    
    
    TXTSLNO.Text = grdsales.rows
    TXTPRODUCT.Text = ""
    txtcategory.Text = ""
    TxtName1.Text = ""
    TXTITEMCODE.Text = ""
    optnet.Value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TrxRYear.Text = ""
    TXTUNIT.Text = ""
    
    lblretail.Caption = ""
    lblwsale.Caption = ""
    lblvan.Caption = ""
    lblunit.Caption = ""
    LblPack.Text = ""
    lblOr_Pack.Caption = ""
    lblcase.Caption = ""
    lblcrtnpack.Caption = ""
    LBLITEMCOST.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTAPPENDQTY.Text = ""
    TXTFREEAPPEND.Text = ""
    txtappendcomm.Text = ""
    TXTAPPENDTOTAL.Text = ""
    TxtMRP.Text = ""
    txtinvnodate.Text = ""
    TxtInvoiceDate.Text = "  /  /    "
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    txtcommi.Text = ""
    TxtRetailmode.Text = ""
    txtretail.Text = ""
    txtBatch.Text = ""
    TXTTAX.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    TXTDISC.Text = ""
    LBLSUBTOTAL.Caption = ""
    TxtWarranty.Text = ""
    TxtWarranty_type.Text = ""
    lblP_Rate.Caption = "0"
    cmdadd.Enabled = False
    'cmddelete.Enabled = False
    CMDEXIT.Enabled = False
    
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
    TXTITEMCODE.Enabled = True
    TXTITEMCODE.SetFocus
    'grdsales.TopRow = grdsales.Rows - 1
Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
            'TxtWarranty.Enabled = True
            'TxtWarranty.SetFocus
    End Select

End Sub

Private Sub CmdDelete_Click()
    
    If grdsales.rows < 1 Then Exit Sub
    If M_EDIT = True Then Exit Sub
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
        MsgBox "Permission Denied", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    TXTSLNO.Text = grdsales.TextMatrix(grdsales.Row, 0)
    Call TXTSLNO_KeyDown(13, 0)


    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & ""
    db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 32)) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            '!ISSUE_QTY = !ISSUE_QTY - (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)))
            If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(LblPack.Text), 3)
            !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) * Val(LblPack.Text), 3)
            !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 35)) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
            !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
      
    For i = Val(TXTSLNO.Text) To grdsales.rows - 2
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
    Next i
    grdsales.rows = grdsales.rows - 1
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.Caption = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
        Else
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
        End If
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    
    Call COSTCALCULATION
    Call Addcommission
    
    TXTSLNO.Text = Val(grdsales.rows)
    TXTPRODUCT.Text = ""
    txtcategory.Text = ""
    TxtName1.Text = ""
    TXTITEMCODE.Text = ""
    optnet.Value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TrxRYear.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTAPPENDQTY.Text = ""
    TXTFREEAPPEND.Text = ""
    txtappendcomm.Text = ""
    TXTAPPENDTOTAL.Text = ""
    txtretail.Text = ""
    txtBatch.Text = ""
    TxtWarranty.Text = ""
    TxtWarranty_type.Text = ""
    TXTTAX.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    TxtMRP.Text = ""
    txtinvnodate.Text = ""
    TxtInvoiceDate.Text = "  /  /    "
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    txtcommi.Text = ""
    TxtRetailmode.Text = ""
    
    TXTDISC.Text = ""
    LBLSUBTOTAL.Caption = ""
    LBLDNORCN.Caption = ""
    cmdadd.Enabled = False
    TXTITEMCODE.Enabled = True
    TXTITEMCODE.SetFocus
    'cmddelete.Enabled = False
    'CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
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
End Sub

Private Sub CmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.Value = True
            TxtMRP.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TrxRYear.Text = ""
            TXTUNIT.Text = ""
            
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
            TXTPRODUCT.Enabled = False
            TxtName1.Enabled = False
            TXTQTY.Enabled = False
            
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
    End Select
End Sub

Private Sub CmdExit_Click()
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
        LblProfitAmt.Visible = True
        LBLITEMCOST.Visible = True
        'LBLSELPRICE.Visible = True
    End If
End Sub

Private Sub CMDMODIFY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    
    If grdsales.rows < 1 Then Exit Sub
    'If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    If M_EDIT = True Then Exit Sub
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = False Then
        MsgBox "Permission Denied", vbOKOnly, "Sales"
        Exit Sub
    End If
    
    TXTSLNO.Text = grdsales.TextMatrix(grdsales.Row, 0)
    Call TXTSLNO_KeyDown(13, 0)

    If UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICE CHARGE" Then
        CMDMODIFY.Enabled = False
        CmdDelete.Enabled = False
        CMDEXIT.Enabled = False
        M_EDIT = True
        TXTTAX.Enabled = True
        TXTTAX.SetFocus
        Exit Sub
    End If
    
    M_ADD = True
    On Error GoTo ERRHAND
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_QTY = !ISSUE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(LblPack.Text), 3)
            !FREE_QTY = !FREE_QTY - Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) * Val(LblPack.Text), 3)
            !CLOSE_QTY = !CLOSE_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 35)) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY - Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
            !BAL_QTY = !BAL_QTY + Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20))) * Val(LblPack.Text), 3)
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
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            
            optnet.Value = True
            TxtMRP.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TrxRYear.Text = ""
            TXTUNIT.Text = ""
            
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
            TXTPRODUCT.Enabled = False
            TxtName1.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub CmdPrint_Click()
        
    Chkcancel.Value = 0
    If grdsales.rows = 1 Then Exit Sub
    
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    'If Val(txtBillNo.Text) > 100 Then Exit Sub
    'If Month(Date) >= 4 And Year(Date) >= 2021 Then Exit Sub
'    If Month(TXTINVDATE.Text) >= 5 And Year(TXTINVDATE.Text) >= 2021 Then
'        'db.Execute "delete From USERS "
'        Exit Sub
'    End If
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) FROM PURCAHSERETURN", db, adOpenForwardOnly
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
    Me.lblcredit.Caption = "1"
    Me.Generateprint
    
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim DN_CN_FLag As Boolean
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
    
    If OLD_BILL = False Then Call checklastbill
    OLD_BILL = True
    db.Execute "delete FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From TEMPTRXFILE "
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM PURCAHSERETURN ", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
    RSTTRXFILE!TRX_TYPE = "PR"
    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
    RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.Text
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
    RSTTRXFILE!BILL_NAME = Trim(TXTDEALER.Text)
    RSTTRXFILE!BILL_ADDRESS = Trim(lbladdress.Caption)
    RSTTRXFILE!Area = ""
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "PR"
        'RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 13)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 5))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 9)
        RSTTRXFILE!UNIT = grdsales.TextMatrix(i, 4)
        'RSTTRXFILE!VCH_DESC = "" '"Issued to     " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 10))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 18))
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
        If grdsales.TextMatrix(i, 31) = "" Then
            'RSTTRXFILE!EXP_DATE = Null
        Else
            RSTTRXFILE!EXP_DATE = LastDayOfMonth("01/" & Trim(grdsales.TextMatrix(i, 31))) & "/" & Trim(grdsales.TextMatrix(i, 31))
        End If
        RSTTRXFILE!FREE_QTY = Val(grdsales.TextMatrix(i, 20))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 21))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 22))
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!SALE_1_FLAG = Trim(grdsales.TextMatrix(i, 23))
        RSTTRXFILE!COM_AMT = Val(grdsales.TextMatrix(i, 24))
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 27))
        RSTTRXFILE!WARRANTY = IIf(grdsales.TextMatrix(i, 28) = "", Null, grdsales.TextMatrix(i, 28))
        RSTTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(i, 29)
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 30)
        RSTTRXFILE!LOOSE_FLAG = grdsales.TextMatrix(i, 26)
        RSTTRXFILE!VCH_DESC = grdsales.TextMatrix(i, 33) & IIf(IsDate(grdsales.TextMatrix(i, 34)), " DTD " & Format(grdsales.TextMatrix(i, 34), "dd/mm/yyyy"), "")
        RSTTRXFILE!CESS_PER = 0
        RSTTRXFILE!cess_amt = 0
        RSTTRXFILE!kfc_tax = 0
        
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!C_USER_ID = IIf(IsNull(RSTITEMMAST!REMARKS), "", Left(RSTITEMMAST!REMARKS, 8))
            RSTTRXFILE!MFGR = IIf(IsNull(RSTITEMMAST!ITEM_MAL), "", RSTITEMMAST!ITEM_MAL)
            RSTTRXFILE!M_USER_ID = IIf(IsNull(RSTITEMMAST!FULL_PACK), "", RSTITEMMAST!FULL_PACK)
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
'        Dim RSTITEMMAST As ADODB.Recordset
'        Set RSTITEMMAST = New ADODB.Recordset
'        RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockReadOnly
'        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'            RSTTRXFILE!Area = RSTITEMMAST!Area
'        End If
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.Update
    Next i

    'Call ReportGeneratION_vpestimate
    LBLFOT.Tag = ""
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") And NEW_BILL = True Then
        If MsgBox("You do not have any permission to modify this further. Are you sure to print?", vbYesNo, "BILL..") = vbNo Then Exit Function
    Else
        Screen.MousePointer = vbHourglass
        Sleep (300)
    End If
    
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
        If Trim(TxtVehicle.Text) = "" Then TxtVehicle.Text = IIf(IsNull(RSTCOMPANY!VEHICLE), "", RSTCOMPANY!VEHICLE)
        INV_TERMS = IIf(IsNull(RSTCOMPANY!INV_TERMS) Or RSTCOMPANY!INV_TERMS = "", "", RSTCOMPANY!INV_TERMS)
        BANK_DET = IIf(IsNull(RSTCOMPANY!bank_details) Or RSTCOMPANY!bank_details = "", "", RSTCOMPANY!bank_details)
        PAN_NO = IIf(IsNull(RSTCOMPANY!PAN_NO) Or RSTCOMPANY!PAN_NO = "", "", RSTCOMPANY!PAN_NO)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim IGST, Sup_Address, Sup_State, Sup_Phone As String
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "select * from ACTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Sup_Address = IIf(IsNull(RSTCOMPANY!Address), "", Trim(RSTCOMPANY!Address))
        IGST = IIf(IsNull(RSTCOMPANY!CUST_IGST), "", Trim(RSTCOMPANY!CUST_IGST))
        Sup_State = IIf(IsNull(RSTCOMPANY!Area), "", "State: " & Trim(RSTCOMPANY!Area))
        Sup_Phone = IIf(IsNull(RSTCOMPANY!TELNO) Or RSTCOMPANY!TELNO = "", "", ". Ph: " & Trim(RSTCOMPANY!TELNO))
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    NEW_BILL = False
    lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0)) - Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 2)), 2)
    Figre = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text) - (Val(LBLDISCAMT.Caption) + Val(LBLRETAMT.Caption)), 0))
    Num = Abs(Figre)
    If Figre < 0 Then
        LBLFOT.Tag = "(-)Rupees " & Words_1_all(Num) & " Only"
    ElseIf Figre > 0 Then
        LBLFOT.Tag = "(Rupees " & Words_1_all(Num) & " Only)"
    End If
    
    If IGST = "Y" Then
        If Small_Print = True Then
            ReportNameVar = Rptpath & "RptPRIGSTA5"
        Else
            ReportNameVar = Rptpath & "RptPRIGST"
        End If
    Else
        If Small_Print = True Then
            ReportNameVar = Rptpath & "RptPurchRetA5"
        Else
            ReportNameVar = Rptpath & "RptPurchRet"
        End If
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields
'    For Each CRXFormulaField In CRXFormulaFields
'        If Small_Print = True Then
'            If CRXFormulaField.Name = "{@ORG}" Then CRXFormulaField.Text = "'Y'"
'            If CRXFormulaField.Name = "{@DUP}" Then CRXFormulaField.Text = "'N'"
'            If CRXFormulaField.Name = "{@TRI}" Then CRXFormulaField.Text = "'Y'"
'        Else
'            If CRXFormulaField.Name = "{@ORG}" Then CRXFormulaField.Text = "'N'"
'            If CRXFormulaField.Name = "{@DUP}" Then CRXFormulaField.Text = "'Y'"
'            If CRXFormulaField.Name = "{@TRI}" Then CRXFormulaField.Text = "'N'"
'        End If
'    Next
    'Report.RecordSelectionFormula = "( {TRXFILE.TRX_TYPE}='PR' AND {TRXFILE.VCH_NO}= " & Val(txtBillNo.Text) & ")"
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
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
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.Text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
        If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.Text = "'" & CompName & "'"
        If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.Text = "'" & CompAddress1 & "'"
        If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.Text = "'" & CompAddress2 & "'"
        If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.Text = "'" & CompAddress3 & "'"
        If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.Text = "'" & CompAddress4 & "'"
        If CRXFormulaField.Name = "{@Comp_Address5}" Then CRXFormulaField.Text = "'" & CompAddress5 & "'"
        If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.Text = "'" & CompTin & "'"
        If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.Text = "'" & CompCST & "'"
        If CRXFormulaField.Name = "{@DL}" Then CRXFormulaField.Text = "'" & DL & "'"
        If CRXFormulaField.Name = "{@ML}" Then CRXFormulaField.Text = "'" & ML & "'"
        If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.Text = "'" & INV_TERMS & "'"
        If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.Text = "'" & BANK_DET & "'"
        If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.Text = "'" & PAN_NO & "'"
        If CRXFormulaField.Name = "{@DL1}" Then CRXFormulaField.Text = "'" & DL1 & "'"
        If CRXFormulaField.Name = "{@DL2}" Then CRXFormulaField.Text = "'" & DL2 & "'"
            
        If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & TXTDEALER.Text & "'"
        If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Sup_Address & "'"
        If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & Sup_Phone & "'"
        If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = "'" & Sup_State & "'"
        If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
        If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.Text = "'" & Trim(TxtOrder.Text) & "'"
        'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
'            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
'            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
        If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
'        If Tax_Print = False Then
'            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
'        End If
        If Trim(TXTTIN.Text) <> "" Then If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'GSTIN: ' & '" & Trim(TXTTIN.Text) & "'"
        'If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
        If CRXFormulaField.Name = "{@VCH_NO}" Then
                Me.Tag = Format(Trim(txtBillNo.Text), bill_for)
                CRXFormulaField.Text = "'" & Me.Tag & "' "
            End If
        'If CRXFormulaField.Name = "{@Order}" Then CRXFormulaField.Text = "'" & Trim(TxtOrder.Text) & "'"
        If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = " " & Val(LBLDISCAMT.Caption) & " "
'            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
        If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.Text = " " & Val(txtOutstanding.Text) & " "
        If CRXFormulaField.Name = "{@Frieght}" Then CRXFormulaField.Text = "'" & Trim(lblFrieght.Text) & "'"
        If CRXFormulaField.Name = "{@FC}" Then CRXFormulaField.Text = " " & Val(TxtFrieght.Text) & " "
        If CRXFormulaField.Name = "{@HANDLE}" Then CRXFormulaField.Text = " '" & Trim(lblhandle.Text) & "' "
        If CRXFormulaField.Name = "{@HC}" Then CRXFormulaField.Text = " " & Val(Txthandle.Text) & " "
'        If Val(LBLRETAMT.Caption) = 0 Then
'            If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'N' "
'        Else
'            If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.Text = " 'Y' "
'        End If
    Next

    frmreport.Caption = "BILL"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
SKIP:
    CMDEXIT.Enabled = False
    TxtName1.Enabled = True
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    TXTQTY.Enabled = False
    
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    txtretail.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTDISC.Enabled = False
    txtcommi.Enabled = False
    OLD_BILL = True
    ST_PRINT = False
    ''rptPRINT.Action = 1
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
    ST_PRINT = False
End Function

Private Sub CMDPRINT_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.rows
            'TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            
            optnet.Value = True
            TxtMRP.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TrxRYear.Text = ""
            TXTUNIT.Text = ""
            
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
            TXTPRODUCT.Enabled = False
            TxtName1.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTDISC.Enabled = False
            txtcommi.Enabled = False
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
    
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
    If (DateValue(TXTINVDATE.Text) < DateValue(MDIMAIN.DTFROM.Value)) Or (DateValue(TXTINVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.Value))) Then
        'db.Execute "delete from Users"
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
            
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        FRMEHEAD.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
            
'    If Trim(TXTAREA.Text) = "" Then
'        MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'        FRMEHEAD.Enabled = True
'        TXTAREA.SetFocus
'        Exit Sub
'    End If

    lblcredit.Caption = "1"
    Call AppendSale
    TxtCN.Text = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    Chkcancel.Value = 0
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    
    'Me.Enabled = False
    'FRMDEBITRT.Show
    
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.rows
            'TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTTAX.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            
            optnet.Value = True
            TxtMRP.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            txtmrpbt.Text = ""
            txtretaildummy.Text = ""
            txtcommi.Text = ""
            TxtRetailmode.Text = ""
            
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TrxRYear.Text = ""
            TXTUNIT.Text = ""
            
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
            TXTPRODUCT.Enabled = False
            TxtName1.Enabled = False
            TXTQTY.Enabled = False
            
            txtretail.Enabled = False
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = False
            TXTFREE.Enabled = False
            TXTDISC.Enabled = False
            'CMDMODIFY.Enabled = False
            'cmddelete.Enabled = False
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub Form_Activate()
    If txtBillNo.Visible = True Then txtBillNo.SetFocus
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    'If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
    If txtretail.Enabled = True Then txtretail.SetFocus
    If TXTRETAILNOTAX.Enabled = True Then TXTRETAILNOTAX.SetFocus
    If TXTTAX.Enabled = True Then TXTTAX.SetFocus
    If TXTDISC.Enabled = True Then TXTDISC.SetFocus
    '''If txtcommi.Enabled = True Then txtcommi.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
    If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
    'If CmdPrintA5.Enabled = True Then CmdPrintA5.SetFocus
    'If  Then CMDDOS.SetFocus
    If TxtName1.Enabled = True Then TxtName1.SetFocus
    'If TXTDEALER.Enabled = True Then TXTDEALER.SetFocus
    If cmdRefresh.Enabled = True Then cmdRefresh.SetFocus
    If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
    If TXTDEALER.Enabled = True Then TXTDEALER.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
'        Select Case KeyCode
'            Case vbKeyF1
'                ST_PRINT = True
'                Call Generateprint
'                ST_PRINT = False
'        End Select
    End If
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ERRHAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PR'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing

    
'    If Val(txtBillNo.Text) > 20 Then
'        Open "C:\WINDOWS\system32\mwp.lp1" For Output As #1 '//Report file Creation
'        Print #1, ""
'        Close #1
'        Exit Sub
'    End If
    
    NEW_BILL = True
    SERIAL_FLAG = False
    ACT_FLAG = True
    AGNT_FLAG = True
    M_ADD = False
    lblcredit.Caption = "1"
    
    lblP_Rate.Caption = "0"
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    grdsales.ColWidth(0) = 600
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 4000
    grdsales.ColWidth(3) = 900
    grdsales.ColWidth(4) = 0
    grdsales.ColWidth(5) = 1200
    grdsales.ColWidth(6) = 1300
    grdsales.ColWidth(7) = 1300
    grdsales.ColWidth(8) = 900
    grdsales.ColWidth(9) = 900
    grdsales.ColWidth(10) = 0
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 1900
    grdsales.ColWidth(13) = 0
    grdsales.ColWidth(14) = 0
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(16) = 0
    grdsales.ColWidth(17) = 0
    grdsales.ColWidth(18) = 0
    grdsales.ColWidth(19) = 0
    grdsales.ColWidth(20) = 1100
    grdsales.ColWidth(21) = 0
    grdsales.ColWidth(22) = 0
    grdsales.ColWidth(23) = 0
    grdsales.ColWidth(24) = 1000
    grdsales.ColWidth(25) = 0
    grdsales.ColWidth(26) = 0
    grdsales.ColWidth(27) = 0
    grdsales.ColWidth(28) = 0
    grdsales.ColWidth(29) = 0
    grdsales.ColWidth(30) = 0
    grdsales.ColWidth(32) = 0
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "UNIT"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "RATE"
    grdsales.TextArray(7) = "NET RATE"
    grdsales.TextArray(8) = "DISC %"
    grdsales.TextArray(9) = "TAX %"
    grdsales.TextArray(10) = "Serial No"
    grdsales.TextArray(11) = "COST"
    grdsales.TextArray(12) = "SUB TOTAL"
    grdsales.TextArray(13) = "ITEM CODE"
    grdsales.TextArray(14) = "Vch No"
    grdsales.TextArray(15) = "Line No"
    grdsales.TextArray(16) = "Trx Type"
    grdsales.TextArray(17) = "Tax Mode"
    grdsales.TextArray(18) = "MFGR"
    grdsales.TextArray(19) = "CN/DN"
    grdsales.TextArray(20) = "FREE"
    grdsales.TextArray(21) = "PTR"
    grdsales.TextArray(22) = "PTRWOTAX"
    grdsales.TextArray(24) = "Commi"
    grdsales.TextArray(31) = "Expiry"
    'grdsales.ColWidth(12) = 0
    'grdsales.ColWidth(13) = 0
    'grdsales.ColWidth(14) = 0
   'grdsales.ColWidth(15) = 0
    'grdsales.ColWidth(16) = 0
    
    grdsales.ColAlignment(0) = 4
    grdsales.ColAlignment(2) = 1
    grdsales.ColAlignment(3) = 4
    grdsales.ColAlignment(5) = 7
    grdsales.ColAlignment(7) = 7
    grdsales.ColAlignment(8) = 4
    grdsales.ColAlignment(12) = 7
    grdsales.ColAlignment(20) = 4
    
    If frmLogin.rs!Level <> "0" Then
'        Label1(21).Visible = False
'        lblretail.Visible = False
        grdsales.ColWidth(31) = 0
    Else
        grdsales.ColWidth(31) = 1100
        Label1(21).Visible = True
        lblretail.Visible = True
    End If
    
    LBLTOTAL.Caption = 0
    lblcomamt.Caption = 0
    LBLRETAMT.Caption = 0
    
    PHYFLAG = True
    PHYCODEFLAG = True
    TMPFLAG = True
    BATCH_FLAG = True
    ITEM_FLAG = True
    PRERATE_FLAG = True
    cr_days = False
    
    TXTPRODUCT.Enabled = False
    TXTITEMCODE.Enabled = False
    TXTQTY.Enabled = False
    
    TxtMRP.Enabled = False
    txtinvnodate.Enabled = False
    TxtInvoiceDate.Enabled = False
    TxtInvoiceDate.Enabled = False
    txtretail.Enabled = False
    txtcommi.Enabled = False
    TXTRETAILNOTAX.Enabled = False
    TXTTAX.Enabled = False
    TXTFREE.Enabled = False
    TXTDISC.Enabled = False
    txtcommi.Enabled = False
    'cmddelete.Enabled = False
    'CMDMODIFY.Enabled = False
    CMDPRINT.Enabled = False
    
    CmdPrintA5.Enabled = False
    
    TXTSLNO.Text = 1
    TxtName1.Enabled = True
    CLOSEALL = 1
    TxtCN.Text = ""
    M_EDIT = False
    
    TXTSLNO.Text = grdsales.rows
    txtBillNo.Visible = False
    TXTDEALER.Text = ""
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    
    
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
        If PHYFLAG = False Then PHY.Close
        If PHYCODEFLAG = False Then PHYCODE.Close
        If TMPFLAG = False Then TMPREC.Close
        If BATCH_FLAG = False Then PHY_BATCH.Close
        If ITEM_FLAG = False Then PHY_ITEM.Close
        If PRERATE_FLAG = False Then PHY_PRERATE.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If AGNT_FLAG = False Then ACT_AGNT.Close
    End If
    Cancel = CLOSEALL
    
End Sub

Private Sub GRDPOPUP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTtax As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            SERIAL_FLAG = True
            txtBatch.Text = GRDPOPUP.Columns(0)
            TXTVCHNO.Text = GRDPOPUP.Columns(2)
            TXTLINENO.Text = GRDPOPUP.Columns(3)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(4)
            TrxRYear.Text = GRDPOPUP.Columns(26)
            TxtMRP.Text = IIf(IsNull(GRDPOPUP.Columns(21)), "", GRDPOPUP.Columns(21))
            TXTEXPIRY.Text = IIf(IsDate(GRDPOPUP.Columns(23)), Format(GRDPOPUP.Columns(23), "MM/YY"), "  /  ")
            'TXTUNIT.Text = GRDPOPUP.Columns(5)
            
            FRMEGRDTMP.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = True
            
            TXTQTY.SetFocus
            
            Call CONTINUE_BATCH
            TxtWarranty.Text = GRDPOPUP.Columns(7)
            TxtWarranty_type.Text = GRDPOPUP.Columns(8)
            Set GRDPOPUP.DataSource = Nothing
            Exit Sub
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtappendcomm.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TrxRYear.Text = ""
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
    Dim RSTMINQTY As ADODB.Recordset
    Dim RSTNONSTOCK As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim NONSTOCKFLAG As Boolean
    Dim MINUSFLAG As Boolean
    Dim i As Long
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            NONSTOCKFLAG = False
            MINUSFLAG = False
            M_STOCK = Val(GRDPOPUPITEM.Columns(2))
            'If Trim(GRDPOPUPITEM.Columns(2)) = "" Then Call STOCKADJUST
            txtcommi.Text = ""
            TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            TxtMRP.Text = IIf(IsNull(GRDPOPUPITEM.Columns(20)), "", GRDPOPUPITEM.Columns(20))
            txtcategory.Text = IIf(IsNull(GRDPOPUPITEM.Columns(7)), "", GRDPOPUPITEM.Columns(7))
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TXTPRODUCT.Enabled = False
                TXTITEMCODE.Enabled = False
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
                Exit Sub
            End If
            i = 0
            If M_STOCK <= 0 Then
                If SERIAL_FLAG = True Then
                    MsgBox "AVAILABLE STOCK IS  " & M_STOCK & " ", , "SALES"
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
                If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
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
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, ITEM_COST, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE, CATEGORY, WARRANTY, WARRANTY_TYPE, MRP  From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, ITEM_COST, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, LOOSE_PACK, PACK_TYPE, CATEGORY, WARRANTY, WARRANTY_TYPE, MRP  From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ITEM_FLAG = False
            End If
            Set GRDPOPUPITEM.DataSource = PHY_ITEM
            'GRDPOPUPITEM.RowHeight = 350
            If PHY_ITEM.RecordCount = 0 Then
                FRMEITEM.Visible = False
                FRMEMAIN.Enabled = True
                TXTPRODUCT.Enabled = False
                TXTITEMCODE.Enabled = False
                If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
                TXTQTY.Enabled = True
                
                'TXTQTY.SetFocus
                Exit Sub
            End If
            
'            Dim RSTBATCH As ADODB.Recordset
'            Set RSTBATCH = New ADODB.Recordset
'            RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0", db, adOpenStatic, adLockReadOnly
'            If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
'                'If RSTBATCH.RecordCount > 1 Then
'                    Call FILL_BATCHGRID
'                    RSTBATCH.Close
'                    Set RSTBATCH = Nothing
'                    Exit Sub
'                'End If
'            End If
'            RSTBATCH.Close
'            Set RSTBATCH = Nothing
            LblPack.Text = IIf(IsNull(GRDPOPUPITEM.Columns(15)) Or Val(GRDPOPUPITEM.Columns(15)) = 0, "1", GRDPOPUPITEM.Columns(15))
            lblOr_Pack.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(15)) Or Val(GRDPOPUPITEM.Columns(15)) = 0, "1", GRDPOPUPITEM.Columns(15))
            'txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", Val(GRDPOPUPITEM.Columns(12)) * Val(LblPack.Text))
            txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", Val(GRDPOPUPITEM.Columns(6)))
            lblretail.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(6)), "", GRDPOPUPITEM.Columns(6))
            lblwsale.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(12)), "", GRDPOPUPITEM.Columns(12))
            lblvan.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(13)), "", GRDPOPUPITEM.Columns(13))
            lblcase.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(11)), "", GRDPOPUPITEM.Columns(11))
            lblcrtnpack.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(10)), "", GRDPOPUPITEM.Columns(10))
            lblunit.Caption = IIf(IsNull(GRDPOPUPITEM.Columns(16)), "Nos", GRDPOPUPITEM.Columns(16))
            TxtWarranty.Text = IIf(IsNull(GRDPOPUPITEM.Columns(18)), "", GRDPOPUPITEM.Columns(18))
            TxtWarranty_type.Text = IIf(IsNull(GRDPOPUPITEM.Columns(19)), "", GRDPOPUPITEM.Columns(19))
        
            LblPack.Text = IIf(IsNull(GRDPOPUPITEM.Columns(10)), "", GRDPOPUPITEM.Columns(10))
            If Val(LblPack.Text) = 0 Then LblPack.Text = "1"
            txtretail.Text = IIf(IsNull(GRDPOPUPITEM.Columns(11)), "", GRDPOPUPITEM.Columns(11))
            
            If GRDPOPUPITEM.Columns(7) = "A" Then
                txtretaildummy.Text = IIf(IsNull(GRDPOPUPITEM.Columns(9)), "P", GRDPOPUPITEM.Columns(9))
                TxtRetailmode.Text = "A"
            Else
                txtretaildummy.Text = IIf(IsNull(GRDPOPUPITEM.Columns(8)), "P", GRDPOPUPITEM.Columns(8))
                TxtRetailmode.Text = "P"
            End If
            Select Case PHY_ITEM!check_flag
                Case "M"
                    OPTTaxMRP.Value = True
                    TXTTAX.Text = GRDPOPUPITEM.Columns(4)
                    TXTSALETYPE.Text = "2"
                Case "V"
                    OPTVAT.Value = True
                    TXTSALETYPE.Text = "2"
                    TXTTAX.Text = GRDPOPUPITEM.Columns(4)
                Case Else
                    TXTSALETYPE.Text = "2"
                    optnet.Value = True
                    TXTTAX.Text = "0"
            End Select
            
'            OPTVAT.value = True
'            TXTTAX.Text = "14.5"
'            TXTSALETYPE.Text = "2"
            
'            optnet.Value = True
            TXTUNIT.Text = GRDPOPUPITEM.Columns(5)
                        
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            FRMEMAIN.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
            TXTQTY.Enabled = True
            
            'TXTQTY.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TrxRYear.Text = ""
            TXTUNIT.Text = ""
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
        Case vbKeyReturn
            txtBatch.Text = GRDPRERATE.Columns(6)
            TXTVCHNO.Text = GRDPRERATE.Columns(15)
            TXTLINENO.Text = GRDPRERATE.Columns(16)
            TXTTRXTYPE.Text = GRDPRERATE.Columns(17)
            TrxRYear.Text = GRDPRERATE.Columns(21)
            TxtMRP.Text = IIf(IsNull(GRDPRERATE.Columns(8)), "", GRDPRERATE.Columns(8))
            TXTEXPIRY.Text = IIf(IsDate(GRDPRERATE.Columns(7)), Format(GRDPRERATE.Columns(7), "MM/YY"), "  /  ")
            
            Dim i As Double
            For i = 1 To grdsales.rows - 1
                If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
                        Exit Sub
                    Else
                        Select Case grdsales.TextMatrix(i, 19)
                            Case "CN", "DN"
                                Exit For
                        End Select
                    End If
                    Exit For
                End If
            Next i
            txtcategory.Text = IIf(IsNull(GRDPRERATE.Columns(18)), "", GRDPRERATE.Columns(18))
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
                Exit Sub
            End If
            LblPack.Text = IIf(IsNull(GRDPRERATE.Columns(12)) Or GRDPRERATE.Columns(12) = "", "1", GRDPRERATE.Columns(12))
            lblOr_Pack.Caption = IIf(IsNull(GRDPRERATE.Columns(12)) Or GRDPRERATE.Columns(12) = "", "1", GRDPRERATE.Columns(12))
            If Val(LblPack.Text) = 0 Then LblPack.Text = 1
            TXTRETAILNOTAX.Text = IIf(IsNull(GRDPRERATE.Columns(9)), "", GRDPRERATE.Columns(9))
            lblretail.Caption = IIf(IsNull(GRDPRERATE.Columns(9)), "", GRDPRERATE.Columns(9))
            lblunit.Caption = IIf(IsNull(GRDPRERATE.Columns(19)), "Nos", GRDPRERATE.Columns(19))
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            'txtinvnodate.Text = Trim(IIf(IsNull(GRDPRERATE.Columns(24)), "", GRDPRERATE.Columns(24))) '& " DTD " & IIf(IsDate(GRDPRERATE.Columns(25)), GRDPRERATE.Columns(25), ""))
            
            txtinvnodate.Text = Trim(IIf(IsNull(GRDPRERATE.Columns(3)), "", GRDPRERATE.Columns(3)))
            TxtInvoiceDate.Text = IIf(IsDate(GRDPRERATE.Columns(2)), Format(GRDPRERATE.Columns(2), "DD/MM/YYYY"), "  /  /    ")
            OPTVAT.Value = True
            TXTSALETYPE.Text = "2"
            TXTTAX.Text = IIf(IsNull(GRDPRERATE.Columns(13)), "", GRDPRERATE.Columns(13))
            TXTUNIT.Text = IIf(IsNull(GRDPRERATE.Columns(20)), "", GRDPRERATE.Columns(20))
            
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            FRMEMAIN.Enabled = True
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
            
        Case vbKeyEscape
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            FRMEMAIN.Enabled = True
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            On Error Resume Next
            'TXTPRODUCT.Text = grdtmp.Columns(1)
            'TXTITEMCODE.Text = grdtmp.Columns(0)
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyReturn
            On Error Resume Next
            TXTITEMCODE.Text = grdtmp.Columns(0)
            TXTPRODUCT.Text = grdtmp.Columns(1)
            Call TxtItemcode_KeyDown(13, 0)
            
            Set grdtmp.DataSource = Nothing
            grdtmp.Visible = False
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
            Else
                If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
                TXTQTY.Enabled = True
                
                'TXTQTY.SetFocus
            End If
    End Select
End Sub

Private Sub LblPack_GotFocus()
    LblPack.SelStart = 0
    LblPack.SelLength = Len(LblPack.Text)
End Sub

Private Sub LblPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(LblPack.Text) = 0 Then Exit Sub
            LblPack.Enabled = False
            If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
            TXTQTY.Enabled = True
            'TXTQTY.SetFocus
        Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            TrxRYear.Text = ""
            TXTUNIT.Text = ""
            TXTPRODUCT.Enabled = True
            LblPack.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub LblPack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub LblPack_LostFocus()
    On Error Resume Next
    If Val(lblOr_Pack.Caption) = 0 Then lblOr_Pack.Caption = 1
    If Val(LblPack.Text) <> Val(lblOr_Pack.Caption) Then
        'txtretail.Text = Val(lblcase.Caption) * Val(LblPack.Text)
        TXTRETAILNOTAX.Text = Val(lblretail.Caption) / Val(lblOr_Pack.Caption) '(Val(lblcase.Caption) / Val(lblcrtnpack.Caption)) * Val(LblPack.Text)
    Else
        TXTRETAILNOTAX.Text = Val(lblretail.Caption)
    End If
    Call TXTRETAILNOTAX_LostFocus
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

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtBatch.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            If M_EDIT = True Then
                'If MsgBox("THIS WILL REMOVE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
                'Call REMOVE_ITEM
                Exit Sub
            End If
            txtBatch.Enabled = False
            LblPack.Enabled = True
            LblPack.SetFocus
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

Private Sub txtBillNo_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
    cr_days = False
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
'    MDIMAIN.MNUENTRY.Visible = False
'    MDIMAIN.MNUREPORT.Visible = False
'    MDIMAIN.mnugud_rep.Visible = False
'    MDIMAIN.MNUTOOLS.Visible = False
End Sub

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TRXMAST As ADODB.Recordset
    Dim TRXSUB As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    
    Dim i As Long
    Dim n As Integer
    Dim M As Integer

    On Error GoTo ERRHAND
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click

    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtBillNo.Text) = 0 Then Exit Sub
            'If Val(txtBillNo.Text) > 800 Then Exit Sub
            grdsales.rows = 1
            i = 0
            Set TRXFILE = New ADODB.Recordset
            TRXFILE.Open "Select * From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until TRXFILE.EOF
                i = i + 1
                TXTINVDATE.Text = Format(TRXFILE!VCH_DATE, "DD/MM/YYYY")
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = TRXFILE!ITEM_CODE
                grdsales.TextMatrix(i, 2) = TRXFILE!ITEM_NAME
                grdsales.TextMatrix(i, 3) = TRXFILE!QTY
'                    Set TRXMAST = New ADODB.Recordset
'                    TRXMAST.Open "SELECT UNIT FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(TRXSUB!R_TRX_TYPE) & "' AND RTRXFILE.VCH_NO = " & Val(TRXSUB!R_VCH_NO) & " AND RTRXFILE.LINE_NO = " & Val(TRXSUB!R_LINE_NO) & "", db, adOpenStatic, adLockReadOnly
'                    If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
'
'                    End If
'                    TRXMAST.Close
'                    Set TRXMAST = Nothing
                
                Set TRXMAST = New ADODB.Recordset
                TRXMAST.Open "SELECT MANUFACTURER FROM ITEMMAST WHERE ITEMMAST.ITEM_CODE = '" & Trim(TRXFILE!ITEM_CODE) & "'", db, adOpenStatic, adLockReadOnly
                If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                    grdsales.TextMatrix(i, 18) = IIf(IsNull(TRXMAST!MANUFACTURER), "", Trim(TRXMAST!MANUFACTURER))
                End If
                TRXMAST.Close
                Set TRXMAST = Nothing
                
                grdsales.TextMatrix(i, 4) = 1
                grdsales.TextMatrix(i, 5) = Format(TRXFILE!MRP, ".000")
                grdsales.TextMatrix(i, 6) = Format(TRXFILE!PTR, ".0000")
                grdsales.TextMatrix(i, 7) = Format(TRXFILE!SALES_PRICE, ".0000")
                grdsales.TextMatrix(i, 8) = IIf(IsNull(TRXFILE!LINE_DISC), 0, TRXFILE!LINE_DISC) 'DISC
                grdsales.TextMatrix(i, 9) = Val(TRXFILE!SALES_TAX)
        
                grdsales.TextMatrix(i, 10) = IIf(IsNull(TRXFILE!REF_NO), "", TRXFILE!REF_NO) 'SERIAL
                grdsales.TextMatrix(i, 11) = IIf(IsNull(TRXFILE!ITEM_COST), 0, TRXFILE!ITEM_COST)
                grdsales.TextMatrix(i, 12) = Format(Val(TRXFILE!TRX_TOTAL), ".000")
                
                grdsales.TextMatrix(i, 13) = TRXFILE!ITEM_CODE
                
                Set TRXSUB = New ADODB.Recordset
                TRXSUB.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & " AND LINE_NO = " & Val(TRXFILE!LINE_NO) & "", db, adOpenStatic, adLockReadOnly
                If Not (TRXSUB.EOF Or TRXSUB.BOF) Then
                    grdsales.TextMatrix(i, 14) = Val(TRXSUB!R_VCH_NO)
                    grdsales.TextMatrix(i, 15) = Val(TRXSUB!R_LINE_NO)
                    grdsales.TextMatrix(i, 16) = Trim(TRXSUB!R_TRX_TYPE)
                    grdsales.TextMatrix(i, 35) = Trim(TRXSUB!R_TRX_YEAR)
                    'grdsales.TextMatrix(i, 43) = IIf(IsNull(TRXSUB!R_TRX_YEAR), "", Trim(TRXSUB!R_TRX_YEAR))
                End If
                TRXSUB.Close
                Set TRXSUB = Nothing
                
                grdsales.TextMatrix(i, 17) = IIf(IsNull(TRXFILE!check_flag), "", Trim(TRXFILE!check_flag))
                TXTDEALER.Text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                'TxtCode.Text = TRXFILE!ACT_CODE
                'DataList2.Text = IIf(IsNull(TRXFILE!VCH_DESC), "", Mid(TRXFILE!VCH_DESC, 15))
                LBLDATE.Caption = IIf(IsNull(TRXFILE!CREATE_DATE), Date, TRXFILE!CREATE_DATE)
                Select Case TRXFILE!CST
                    Case 0
                        grdsales.TextMatrix(i, 19) = "B"
                    Case 1
                        grdsales.TextMatrix(i, 19) = "DN"
                    Case 2
                        grdsales.TextMatrix(i, 19) = "CN"
                End Select
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
                grdsales.TextMatrix(i, 31) = IIf(IsDate(TRXFILE!EXP_DATE), Format(TRXFILE!EXP_DATE, "MM/YY"), "")
                grdsales.TextMatrix(i, 32) = TRXFILE!LINE_NO
                grdsales.TextMatrix(i, 33) = IIf(IsNull(TRXFILE!INV_DETAILS), "", TRXFILE!INV_DETAILS)
                grdsales.TextMatrix(i, 34) = IIf(IsDate(TRXFILE!INV_DATE), Format(TRXFILE!INV_DATE, "DD/MM/YYYY"), "")
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
            
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "Select * FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF Or TRXMAST.BOF) Then
                If TRXMAST!SLSM_CODE = "A" Then
                    TXTTOTALDISC.Text = IIf(IsNull(TRXMAST!DISCOUNT), "", TRXMAST!DISCOUNT)
                    OptDiscAmt.Value = True
                ElseIf TRXMAST!SLSM_CODE = "P" Then
                    If IsNull(TRXMAST!VCH_AMOUNT) Or TRXMAST!VCH_AMOUNT = 0 Then
                        TXTTOTALDISC.Text = 0
                    Else
                        TXTTOTALDISC.Text = IIf(IsNull(TRXMAST!DISCOUNT), "", Round((TRXMAST!DISCOUNT * 100 / TRXMAST!VCH_AMOUNT), 2))
                    End If
                    OPTDISCPERCENT.Value = True
                End If
                LBLRETAMT.Caption = IIf(IsNull(TRXMAST!ADD_AMOUNT), "", Format(TRXMAST!ADD_AMOUNT, "0.00"))
                If TRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                TxtFrieght.Text = IIf(IsNull(TRXMAST!FRIEGHT), "", TRXMAST!FRIEGHT)
                Txthandle.Text = IIf(IsNull(TRXMAST!HANDLE), "", TRXMAST!HANDLE)
                TXTTIN.Text = IIf(IsNull(TRXMAST!TIN), "", TRXMAST!TIN)
                TXTDEALER.Text = IIf(IsNull(TRXMAST!ACT_NAME), "", TRXMAST!ACT_NAME)
                lbladdress.Caption = IIf(IsNull(TRXMAST!BILL_ADDRESS), "", TRXMAST!BILL_ADDRESS)
                TxtVehicle.Text = IIf(IsNull(TRXMAST!VEHICLE), "", TRXMAST!VEHICLE)
                TxtOrder.Text = IIf(IsNull(TRXMAST!D_ORDER), "", TRXMAST!D_ORDER)
                'DataList2.BoundText = IIf(IsNull(TRXMAST!ACT_CODE), "", TRXMAST!ACT_CODE)
                TxtCN.Text = IIf(IsNull(TRXMAST!CN_REF), "", TRXMAST!CN_REF)
                NEW_BILL = False
                OLD_BILL = True
            Else
                NEW_BILL = True
                OLD_BILL = False
                'TXTTYPE.Text = ""
                'cmbtype.ListIndex = -1
                
                
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
            
            If grdsales.rows > 1 Then
                NEW_BILL = False
                OLD_BILL = True
            End If
            LBLBILLNO.Caption = Val(txtBillNo.Text)
            LBLTOTAL.Caption = ""
            lblnetamount.Caption = ""
            LBLFOT.Caption = ""
            lblcomamt.Caption = ""
            For i = 1 To grdsales.rows - 1
                grdsales.TextMatrix(i, 0) = i
                Select Case grdsales.TextMatrix(i, 19)
                    Case "CN"
                        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                        LBLFOT.Caption = ""
                    Case Else
                        LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                        If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                        LBLFOT.Caption = ""
                End Select
                
                If Val(grdsales.TextMatrix(i, 3)) = 0 Then
                    lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
                Else
                    lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
                End If
            Next i
            LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
            TXTAMOUNT.Text = ""
            If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
                TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
            ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
                TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
            End If
            LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
            lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
            Call COSTCALCULATION
            Call Addcommission
            
            TXTSLNO.Text = grdsales.rows
            txtBillNo.Visible = False
            TxtName1.Enabled = True
            
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
    End Select
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    Exit Sub
ERRHAND:
    MsgBox err.Description
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
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ERRHAND
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'PR'", db, adOpenStatic, adLockReadOnly
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
'    TRXMAST.Open "Select MIN(VCH_NO) FROM TRXFILE WHERE TRX_TYPE = 'PR'", db, adOpenStatic, adLockReadOnly
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
    txtcategory.SelLength = Len(txtcategory.Text)
    SERIAL_FLAG = False
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
            TxtName1.Enabled = True
            txtcategory.Enabled = False
            TxtName1.SetFocus
    End Select
End Sub

Private Sub txtcategory_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
ERRHAND:
    MsgBox err.Description
    
    
End Sub

Private Sub TXTDISC_GotFocus()
    TXTDISC.SelStart = 0
    TXTDISC.SelLength = Len(TXTDISC.Text)
End Sub

Private Sub TXTDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTDISC.Enabled = False
'            txtcommi.Enabled = True
'            txtcommi.SetFocus
'            Exit Sub
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Call TXTDISC_LostFocus
            Call CMDADD_Click
'            TXTDISC.Enabled = False
'            cmdadd.Enabled = True
'            cmdadd.SetFocus
'            'TxtWarranty.Enabled = True
'            'TxtWarranty.SetFocus
        Case vbKeyEscape
            txtretail.Enabled = True
            TXTDISC.Enabled = False
            txtretail.SetFocus
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
    If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
        TXTDISC.Tag = Val(txtretail.Text) * Val(TXTDISC.Text) / 100
        LBLSUBTOTAL.Caption = Format(Round(Val(txtretail.Text) - Val(TXTDISC.Tag), 2), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.Text) * Val(txtretail.Text) * Val(TXTDISC.Text) / 100
        LBLSUBTOTAL.Caption = Format(Round((Val(TXTQTY.Text) * Val(txtretail.Text)) - Val(TXTDISC.Tag), 2), ".000")
    End If
    
    ''TXTDISC.Text = Format(TXTDISC.Text, ".000")

End Sub

Private Sub TxtFrieght_GotFocus()
    TxtFrieght.SelStart = 0
    TxtFrieght.SelLength = Len(TxtFrieght.Text)
End Sub

Private Sub TxtFrieght_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            'If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            '''If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub TxtFrieght_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtFrieght_LostFocus()
    Call TXTTOTALDISC_LostFocus
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
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTINVDATE.SetFocus
                Exit Sub
            End If
            
            If (DateValue(TXTINVDATE.Text) < DateValue(MDIMAIN.DTFROM.Value)) Or (DateValue(TXTINVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.Value))) Then
                'db.Execute "delete from Users"
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTINVDATE.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtInvoiceDate_GotFocus()
    TxtInvoiceDate.SelStart = 0
    TxtInvoiceDate.SelLength = Len(TxtInvoiceDate.Text)
End Sub

Private Sub TxtInvoiceDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Len(Trim(TxtInvoiceDate.Text)) = 4 Then GoTo SKID
            If Not IsDate(TxtInvoiceDate.Text) Then Exit Sub
SKID:
            txtinvnodate.Enabled = False
            TxtInvoiceDate.Enabled = False
            TxtMRP.Enabled = True
            TxtMRP.SetFocus
        Case vbKeyEscape
            If Trim(txtinvnodate.Text) = "" Then
                txtinvnodate.Enabled = True
                TxtInvoiceDate.Enabled = True
                txtinvnodate.SetFocus
            Else
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
            End If
    End Select
End Sub

Private Sub TxtInvoiceDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
            TxtMRP.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
        Case vbKeyEscape
            TxtMRP.Enabled = False
            TXTQTY.Enabled = True
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

Private Sub TxtName1_GotFocus()
    TxtName1.SelStart = 0
    TxtName1.SelLength = Len(TxtName1.Text)
    SERIAL_FLAG = False
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TxtName1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtBillNo.Text) > 100 Then Exit Sub
            If UCase(TxtName1.Text) = "OT" Then TXTITEMCODE.Text = "OT"
                
            TxtName1.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If grdsales.rows > 1 Then
                TxtName1.Enabled = False
                
                CMDPRINT.Enabled = True
                CmdPrintA5.Enabled = True
                cmdRefresh.Enabled = True
                CmdPrintA5.SetFocus
            Else
                TxtName1.Enabled = False
                FRMEHEAD.Enabled = True
                TXTDEALER.Enabled = True
                TXTDEALER.SetFocus
            End If
            LBLDNORCN.Caption = ""
    End Select
End Sub

Private Sub TxtName1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
        Case vbKeyEscape
            TxtVehicle.SetFocus
    End Select
End Sub

Private Sub TXTPRODUCT_Change()
    
        If CHANGE_NAME = False Then Exit Sub
        Dim i As Long
        Dim RSTBATCH As ADODB.Recordset
    
        M_STOCK = 0
        Set grdtmp.DataSource = Nothing
        If PHYFLAG = True Then
            'PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY CATEGORY, ITEM_SLNO", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, BIN_LOCATION  From ITEMMAST  WHERE ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
            PHYFLAG = False
        Else
            PHY.Close
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, BIN_LOCATION  From ITEMMAST  WHERE ITEM_NAME Like '%" & Trim(Me.TXTPRODUCT.Text) & "%' AND ITEM_NAME Like '%" & Trim(Me.TxtName1.Text) & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
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
        grdtmp.Columns(1).Width = 4600
        grdtmp.Columns(2).Caption = "QTY"
        grdtmp.Columns(2).Width = 900
        grdtmp.Columns(6).Caption = "RT"
        grdtmp.Columns(6).Width = 800
        grdtmp.Columns(4).Width = 0
        grdtmp.Columns(4).Width = 0
        grdtmp.Columns(5).Width = 0
        grdtmp.Columns(3).Width = 0
        grdtmp.Columns(7).Width = 0
        grdtmp.Columns(8).Width = 0
        grdtmp.Columns(9).Width = 0
        grdtmp.Columns(10).Width = 800
        grdtmp.Columns(10).Caption = "L/Pack"
        grdtmp.Columns(11).Caption = "LP"
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
        grdtmp.Columns(20).Width = 0
        grdtmp.Columns(21).Width = 3300
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTPRODUCT_GotFocus()
    LBLITEMCOST.Caption = ""
    LBLSELPRICE.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    TXTPRODUCT.Tag = TXTPRODUCT.Text
    TXTPRODUCT.Text = ""
    TXTPRODUCT.Text = TXTPRODUCT.Tag
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
    CHANGE_NAME = True
    'If Trim(TXTPRODUCT.Text) <> "" Or Trim(TxtName1.Text) <> "" Then
    Call TXTPRODUCT_Change
    'grdtmp.Visible = True
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
    
        Case vbKeyReturn
            M_STOCK = 0
            On Error Resume Next
            TXTPRODUCT.Text = grdtmp.Columns(1)
            TxtMRP.Text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
            'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If Trim(TXTPRODUCT.Text) = "" Then
                TXTPRODUCT.Enabled = False
                TxtName1.Enabled = True
                TxtName1.SetFocus
                Exit Sub
            End If
            'cmddelete.Enabled = False
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.Value = True
            TxtMRP.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            LBLSUBTOTAL.Caption = ""
            LblPack.Text = "1"
            lblunit.Caption = "Nos"
            txtcommi.Text = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
            If UCase(TxtName1.Text) = "OT" Then TXTITEMCODE.Text = "OT"
            If UCase(TXTITEMCODE.Text) <> "OT" Then
                Set grdtmp.DataSource = Nothing
                If PHYFLAG = True Then
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, ITEM_COST, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP  From ITEMMAST  WHERE ITEM_NAME = '" & Me.TXTPRODUCT.Text & "'ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                Else
                    PHY.Close
                    PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, ITEM_COST, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP   From ITEMMAST  WHERE ITEM_NAME = '" & Me.TXTPRODUCT.Text & "'ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
                    PHYFLAG = False
                End If
                Set grdtmp.DataSource = PHY
                
                If PHY.RecordCount = 0 Then
                    MsgBox "Item not found!!!!", , "Sales"
                    Exit Sub
                End If
                If PHY.RecordCount = 1 Then
                    TXTITEMCODE.Text = grdtmp.Columns(0)
                    TXTPRODUCT.Text = grdtmp.Columns(1)
'                    Set RSTBATCH = New ADODB.Recordset
'                    RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
'                    If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
'                        'If RSTBATCH.RecordCount > 1 Then
'                            Set grdtmp.DataSource = Nothing
'                            grdtmp.Visible = False
'                            Call FILL_BATCHGRID
'                            RSTBATCH.Close
'                            Set RSTBATCH = Nothing
'                            Exit Sub
'                        'End If
'                    End If
'                    RSTBATCH.Close
'                    Set RSTBATCH = Nothing
                    
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
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
            Else
                If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
                TXTQTY.Enabled = True
                
                'TXTQTY.SetFocus
            End If
        Case vbKeyEscape
            TxtName1.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTDISC.Enabled = False
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
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
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
                txtcategory.Text = IIf(IsNull(PHY!Category), "", PHY!Category)
                If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                    TXTTAX.Enabled = True
                    TXTTAX.SetFocus
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
                LblPack.Text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
                'txtretail.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)) * Val(LblPack.Text))
                txtretail.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                lblretail.Caption = IIf(IsNull(grdtmp.Columns(6)), "", grdtmp.Columns(6))
                lblwsale.Caption = IIf(IsNull(grdtmp.Columns(12)), "", grdtmp.Columns(12))
                lblvan.Caption = IIf(IsNull(grdtmp.Columns(13)), "", grdtmp.Columns(13))
                lblcase.Caption = IIf(IsNull(grdtmp.Columns(11)), "", grdtmp.Columns(11))
                lblcrtnpack.Caption = IIf(IsNull(grdtmp.Columns(10)), "", grdtmp.Columns(10))

                lblunit.Caption = IIf(IsNull(grdtmp.Columns(17)), "Nos", grdtmp.Columns(17))
                TxtWarranty.Text = IIf(IsNull(grdtmp.Columns(18)), "", grdtmp.Columns(18))
                TxtWarranty_type.Text = IIf(IsNull(grdtmp.Columns(19)), "", grdtmp.Columns(19))
                TxtMRP.Text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
                'txtinvnodate.Text = IIf(IsNull(grdtmp.Columns(10)), "", grdtmp.Columns(10))
                'LblPack.Text = IIf(IsNull(grdtmp.Columns(10)), "", grdtmp.Columns(10))
                'If Val(LblPack.Text) = 0 Then LblPack.Text = "1"
                'txtretail.Text = IIf(IsNull(grdtmp.Columns(11)), "", grdtmp.Columns(11))
            
                If grdtmp.Columns(7) = "A" Then
                    txtretaildummy.Text = IIf(IsNull(grdtmp.Columns(9)), "P", grdtmp.Columns(9))
                    TxtRetailmode.Text = "A"
                Else
                    txtretaildummy.Text = IIf(IsNull(grdtmp.Columns(8)), "P", grdtmp.Columns(8))
                    TxtRetailmode.Text = "P"
                End If
                Select Case PHY!check_flag
                    Case "M"
                        OPTTaxMRP.Value = True
                        TXTTAX.Text = grdtmp.Columns(4)
                        TXTSALETYPE.Text = "2"
                    Case "V"
                        OPTVAT.Value = True
                        TXTSALETYPE.Text = "2"
                        TXTTAX.Text = grdtmp.Columns(4)
                    Case Else
                        TXTSALETYPE.Text = "2"
                        optnet.Value = True
                        TXTTAX.Text = "0"
                End Select
                
'                OPTVAT.value = True
'                TXTTAX.Text = "14.5"
'                TXTSALETYPE.Text = "2"
                
                TXTUNIT.Text = grdtmp.Columns(5)
                
                If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
                'TXTPRODUCT.Enabled = False
                
                TXTQTY.Enabled = True
                '
                'TXTQTY.SetFocus
                Exit Function
End Function

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTPRODUCT_LostFocus()
    CHANGE_NAME = False
End Sub

Private Sub TXTQTY_GotFocus()
    If Val(LblPack.Text) = 0 Then LblPack.Text = 1
    If M_EDIT = False Then
        If LblPack.Text = 1 Then
            FrmeType.Visible = False
        Else
            FrmeType.Visible = True
        End If
        If Val(LblPack.Text) = Val(lblOr_Pack.Caption) Then
            OptNormal.Value = True
        Else
            OptLoose.Value = True
        End If
    Else
        FrmeType.Visible = False
    End If
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
    TXTQTY.Tag = Trim(TXTPRODUCT.Text)
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Double
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            i = 0
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT CLOSE_QTY  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!CLOSE_QTY)) Then RSTTRXFILE!CLOSE_QTY = 0
                i = RSTTRXFILE!CLOSE_QTY / Val(LblPack.Text)
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If i <> 0 Then
                If Val(TXTQTY.Text) > i Then
                    If SERIAL_FLAG = True Then
                        MsgBox "AVAILABLE STOCK IS  " & i & " ", , "SALES"
                        TXTQTY.SelStart = 0
                        TXTQTY.SelLength = Len(TXTQTY.Text)
                        Exit Sub
                    End If
                    If (MsgBox("AVAILABLE STOCK IS  " & i & "  Do you want to CONTINUE", vbYesNo, "SALES") = vbNo) Then
                        'MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                        TXTQTY.SelStart = 0
                        TXTQTY.SelLength = Len(TXTQTY.Text)
                        Exit Sub
                    End If
                End If
            End If
SKIP:
            TXTQTY.Enabled = False
            txtinvnodate.Enabled = True
            TxtInvoiceDate.Enabled = True
            TxtInvoiceDate.Enabled = True
            txtinvnodate.SetFocus
         Case vbKeyEscape
            txtBatch.Enabled = True
            TXTQTY.Enabled = False
            txtBatch.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case 102, 70
            If FrmeType.Visible = False Then
                KeyAscii = 0
                Exit Sub
            End If
            If M_EDIT = False Then OptNormal.Value = True
            LblPack.Text = Val(lblOr_Pack.Caption)
            Call LblPack_LostFocus
            KeyAscii = 0
        Case 76, 108
            If FrmeType.Visible = False Then
                KeyAscii = 0
                Exit Sub
            End If
            If M_EDIT = False Then OptLoose.Value = True
            If Val(lblcrtnpack.Caption) = 0 Then lblcrtnpack.Caption = 1
            LblPack.Text = Val(lblcrtnpack.Caption)
            Call LblPack_LostFocus
            KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    
    Dim RSTITEMCOST As ADODB.Recordset
    
    TXTQTY.Text = Format(TXTQTY.Text, ".000")
    TXTDISC.Tag = 0
    TXTTAX.Tag = 0
    If Val(TXTRETAILNOTAX.Text) = 0 Then
        TXTDISC.Tag = Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(txtretail.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    Else
        TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
        TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    End If
    
    On Error GoTo ERRHAND
    Set RSTITEMCOST = New ADODB.Recordset
    RSTITEMCOST.Open "SELECT ITEM_COST, SALES_PRICE FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
        
    If Not (RSTITEMCOST.EOF Or RSTITEMCOST.BOF) Then
        LBLITEMCOST.Caption = IIf(IsNull(RSTITEMCOST!ITEM_COST), "", RSTITEMCOST!ITEM_COST * Val(LblPack.Text))
        LBLSELPRICE.Caption = IIf(IsNull(RSTITEMCOST!SALES_PRICE), "", RSTITEMCOST!SALES_PRICE * Val(LblPack.Text))
    End If
    RSTITEMCOST.Close
    Set RSTITEMCOST = Nothing
    
    Exit Sub
ERRHAND:
    MsgBox err.Description

End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
    Chkcancel.Value = 0
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            'If Val(TXTSLNO.Text) < grdsales.Rows Then Exit Sub
            If Val(TXTSLNO.Text) = 0 Then
                SERIAL_FLAG = False
                TXTSLNO.Text = ""
                TXTPRODUCT.Text = ""
                TXTQTY.Text = ""
                TXTEXPIRY.Text = "  /  "
                TXTAPPENDQTY.Text = ""
                TXTFREEAPPEND.Text = ""
                txtappendcomm.Text = ""
                TXTAPPENDTOTAL.Text = ""
                TXTFREE.Text = ""
                optnet.Value = True
                TxtMRP.Text = ""
                txtinvnodate.Text = ""
                TxtInvoiceDate.Text = "  /  /    "
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TrxRYear.Text = ""
                TXTUNIT.Text = ""
                TXTSLNO.Text = grdsales.rows
                'cmddelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.Text) >= grdsales.rows Then
                TXTSLNO.Text = grdsales.rows
                CmdDelete.Enabled = False
                CMDMODIFY.Enabled = False
            End If
            If Val(TXTSLNO.Text) < grdsales.rows Then
                lblP_Rate.Caption = "1"
                TXTSLNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 0)
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 3)
                TXTFREE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 20)
                TxtMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TXTDISC.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                LBLSUBTOTAL.Caption = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), ".000")
                
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 13)
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 15)
                TXTTRXTYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 16)
                TrxRYear.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 35)
                
                TXTUNIT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                'TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 22)
                TxtRetailmode.Text = "A"
                
                Select Case grdsales.TextMatrix(Val(TXTSLNO.Text), 17)
                    Case "M"
                        OPTTaxMRP.Value = True
                        TXTSALETYPE.Text = "2"
                    Case "V"
                        OPTVAT.Value = True
                        TXTSALETYPE.Text = "2"
                    Case Else
                        TXTSALETYPE.Text = "2"
                        optnet.Value = True
                        TXTTAX.Text = "0"
                End Select
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                TXTRETAILNOTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                txtretail.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                TXTSALETYPE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 23)
                txtcategory.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 25)
'                If UCase(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = "SERVICE CHARGE" Then
'                    txtretaildummy.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)), 2)
'                    'txtcommi.Text = 0 'Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) / Val(TXTQTY.Text), 2)
'                Else
'                    txtretaildummy.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)), 2)
'                    'txtcommi.Text = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) / Val(TXTQTY.Text), 2)
'                End If
                txtretaildummy.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
                txtcommi.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
                txtinvnodate.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 33))
                TxtInvoiceDate.Text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 34)), grdsales.TextMatrix(Val(TXTSLNO.Text), 34), "  /  /    ")
                
                TxtName1.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTITEMCODE.Enabled = False
                TXTQTY.Enabled = False
                
                TXTTAX.Enabled = False
                TXTFREE.Enabled = False
                txtretail.Enabled = False
                TXTRETAILNOTAX.Enabled = False
                TXTDISC.Enabled = False
                TxtMRP.Enabled = False
                txtinvnodate.Enabled = False
                TxtInvoiceDate.Enabled = False
                Select Case grdsales.TextMatrix(Val(TXTSLNO.Text), 19)
                    Case "CN", "DN"
                        CmdDelete.Enabled = True
                        CmdDelete.SetFocus
                        
                    Case Else
                        CMDMODIFY.Enabled = True
                        CMDMODIFY.SetFocus
                        CmdDelete.Enabled = True
                End Select
                LBLDNORCN.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 19)
                LblPack.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 27))
                lblOr_Pack.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 27))
                TxtWarranty.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 28)
                TxtWarranty_type.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 29))
                lblunit.Caption = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 30))
                
                If grdsales.TextMatrix(Val(TXTSLNO.Text), 31) = "" Then
                    TXTEXPIRY.Text = "  /  "
                Else
                    TXTEXPIRY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 31)
                End If
                Set grdtmp.DataSource = Nothing
                grdtmp.Visible = False
                Exit Sub
            End If
SKIP:
            lblP_Rate.Caption = "0"
            TxtName1.Enabled = False
            TXTITEMCODE.Enabled = True
            TXTQTY.Enabled = False
            
            TXTTAX.Enabled = False
            TXTDISC.Enabled = False
            Set grdtmp.DataSource = Nothing
            grdtmp.Visible = False
            TXTITEMCODE.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                optnet.Value = True
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTTRXTYPE.Text = ""
                TrxRYear.Text = ""
                TXTUNIT.Text = ""
                TXTQTY.Text = ""
                TXTEXPIRY.Text = "  /  "
                TXTAPPENDQTY.Text = ""
                TXTAPPENDTOTAL.Text = ""
                TXTFREEAPPEND.Text = ""
                txtappendcomm.Text = ""
                txtretail.Text = ""
                txtBatch.Text = ""
                TxtWarranty.Text = ""
                TxtWarranty_type.Text = ""
                TXTTAX.Text = ""
                TXTRETAILNOTAX.Text = ""
                TXTSALETYPE.Text = ""
                TXTFREE.Text = ""
                TxtMRP.Text = ""
                txtinvnodate.Text = ""
                TxtInvoiceDate.Text = "  /  /    "
                TXTDISC.Text = ""
                LBLSUBTOTAL.Caption = ""
                lblP_Rate.Caption = "0"
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTITEMCODE.Enabled = True
                TXTITEMCODE.SetFocus
            ElseIf grdsales.rows > 1 Then
                TxtName1.Enabled = False
                
                CMDPRINT.Enabled = True
                CmdPrintA5.Enabled = True
                cmdRefresh.Enabled = True
                CmdPrintA5.SetFocus
            Else
                TxtName1.Enabled = False
                FRMEHEAD.Enabled = True
                TXTDEALER.Enabled = True
                TXTDEALER.SetFocus
            End If
            LBLDNORCN.Caption = ""
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

Private Sub TxtTax_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TxtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRETAILNOTAX.Enabled = True
            TXTTAX.Enabled = False
            TXTRETAILNOTAX.SetFocus
        Case vbKeyEscape
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                If M_EDIT = True Then Exit Sub
                txtretail.Enabled = False
                TXTPRODUCT.Enabled = True
                TXTPRODUCT.SetFocus
            Else
                txtretail.Enabled = False
                TXTRETAILNOTAX.Enabled = True
                TXTRETAILNOTAX.SetFocus
            End If
    End Select
End Sub

Private Sub TxtTax_KeyPress(KeyAscii As Integer)
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
'    TXTDISC.Tag = 0
'    TXTTAX.Tag = 0
'    TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTDISC.Text) / 100
'    TXTTAX.Tag = Val(TXTQTY.Text) * Val(TXTRATE.Text) * Val(TXTTAX.Text) / 100
'    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 3)) - Val(TXTDISC.Tag) + Val(TXTTAX.Tag), ".000")
    If optnet.Value = True And Val(TXTTAX.Text) > 0 Then
        OPTVAT.Value = True
        TXTRETAILNOTAX_LostFocus
    End If
    txtmrpbt.Text = 100 * Val(TxtMRP.Text) / (100 + Val(TXTTAX.Text))
    
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
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, P_CRTN, CATEGORY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
    For n = 1 To grdsales.rows - 1
        COST = COST + (Val(grdsales.TextMatrix(n, 11)) * Val(grdsales.TextMatrix(n, 3)))
    Next n
    
    LBLTOTALCOST.Caption = Round(COST, 2)
    LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - COST, 2)

    Exit Function
    
ERRHAND:
    MsgBox err.Description
End Function

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
    
    If OLD_BILL = False Then Call checklastbill
    db.Execute "delete FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete From CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TRX_TYPE = 'CN'"
    
    
    TxtVehicle.Text = ""
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT DISTINCT INV_DETAILS FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'PR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTITEMMAST.EOF
        TxtVehicle.Text = TxtVehicle.Text & IIf(IsNull(RSTITEMMAST!INV_DETAILS) Or RSTITEMMAST!INV_DETAILS = "", "", RSTITEMMAST!INV_DETAILS) & ", "
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    TxtVehicle.Text = Left(Trim(TxtVehicle.Text), 200)
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(DataList2.Text), 1, 30)
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    If grdsales.rows = 1 Then
        If OLD_BILL = True Then
            Dim LASTBILL As Long
            LASTBILL = 1
            Set rstBILL = New ADODB.Recordset
            rstBILL.Open "Select MAX(VCH_NO) FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PR'", db, adOpenForwardOnly
            If Not (rstBILL.EOF And rstBILL.BOF) Then
                LASTBILL = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0))
            End If
            rstBILL.Close
            Set rstBILL = Nothing
            
            If LASTBILL = 1 Then GoTo SKIP
            If LASTBILL <> Val(txtBillNo.Text) - 1 Then
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "Select * FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    RSTTRXFILE.AddNew
                    RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
                    RSTTRXFILE!TRX_TYPE = "PR"
                    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                End If
                RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
                RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
                RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                RSTTRXFILE!ACT_CODE = DataList2.BoundText
                RSTTRXFILE!ACT_NAME = DataList2.Text
                RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
                RSTTRXFILE!ADD_AMOUNT = 0
                RSTTRXFILE!ROUNDED_OFF = 0
                RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
                RSTTRXFILE.Update
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            End If
        End If
        GoTo SKIP
    End If
    
    i = 0
   
    'If lblcredit.Caption <> "0" Then
    
    db.Execute "delete From CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'PR' AND INV_TRX_TYPE = 'CN' "
    Set rstMaxRec = New ADODB.Recordset
    rstMaxRec.Open "Select MAX(CR_NO) From CRDTPYMT", db, adOpenForwardOnly
    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
        i = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
    End If
    rstMaxRec.Close
    Set rstMaxRec = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'PR' AND INV_TRX_TYPE = 'CN'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!TRX_TYPE = "PR"
        RSTITEMMAST!INV_TRX_TYPE = "CN"
        RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTITEMMAST!CR_NO = i
        RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
        RSTITEMMAST!RCPT_AMOUNT = Val(lblnetamount.Caption)
    End If
    RSTITEMMAST!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTITEMMAST!INV_AMT = 0
    RSTITEMMAST!check_flag = "N"
    RSTITEMMAST!BAL_AMT = Val(lblnetamount.Caption) - RSTITEMMAST!RCPT_AMOUNT
    RSTITEMMAST!PINV = ""
    RSTITEMMAST!ACT_CODE = DataList2.BoundText
    RSTITEMMAST!ACT_NAME = DataList2.Text
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "PR"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    Else
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    End If
    
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.Text
    RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
    RSTTRXFILE!ADD_AMOUNT = Val(LBLRETAMT.Caption)
    RSTTRXFILE!REF_NO = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "A"
        RSTTRXFILE!DISCOUNT = Val(TXTTOTALDISC.Text)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        RSTTRXFILE!SLSM_CODE = "P"
        RSTTRXFILE!DISCOUNT = Round(RSTTRXFILE!VCH_AMOUNT * Val(TXTTOTALDISC.Text) / 100, 2)
    End If
    RSTTRXFILE!check_flag = "I"
    If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
    RSTTRXFILE!CFORM_NO = Time
    RSTTRXFILE!REMARKS = DataList2.Text
    RSTTRXFILE!DISC_PERS = 0
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!VEHICLE = ""
    RSTTRXFILE!PHONE = ""
    RSTTRXFILE!TIN = Trim(TXTTIN.Text)
    RSTTRXFILE!FRIEGHT = Val(TxtFrieght.Text)
    RSTTRXFILE!HANDLE = Val(Txthandle.Text)
    RSTTRXFILE!Area = ""
    RSTTRXFILE!VEHICLE = Trim(TxtVehicle.Text)
    RSTTRXFILE!D_ORDER = Trim(TxtOrder.Text)
    RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!cr_days = 0
    RSTTRXFILE!BILL_NAME = Trim(TXTDEALER.Text)
    RSTTRXFILE!BILL_ADDRESS = Trim(lbladdress.Caption)
    txtcommi.Tag = ""
    RSTTRXFILE!AGENT_CODE = ""
    RSTTRXFILE!AGENT_NAME = ""
    RSTTRXFILE!BILL_TYPE = ""

    If Val(TxtCN.Text) <> 0 Then RSTTRXFILE!CN_REF = Val(TxtCN.Text)
    RSTTRXFILE!RCPT_AMOUNT = 0
    RSTTRXFILE!RCPT_REFNO = ""
    RSTTRXFILE!BANK_FLAG = "N"
    RSTTRXFILE!CHQ_NO = Null
    RSTTRXFILE!BANK_CODE = Null
    RSTTRXFILE!BANK_NAME = Null
    'RSTTRXFILE!CHQ_DATE = Null
    RSTTRXFILE!CHQ_STATUS = "N"
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
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
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PR'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    OLD_BILL = False
    
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    lbladdress.Caption = ""
    LBLDNORCN.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    LBLRETAMT.Caption = ""
    LBLPROFIT.Caption = ""
    LBLDATE.Caption = Date
    LBLTOTAL.Caption = ""
    lblcomamt.Caption = ""
    TXTTOTALDISC.Text = ""
    LBLTOTALCOST.Caption = ""
    TXTAMOUNT.Text = ""
    LBLDISCAMT.Caption = ""
    grdsales.rows = 1
    TXTSLNO.Text = 1
    M_EDIT = False
    NEW_BILL = True
    cmdRefresh.Enabled = False
    CMDEXIT.Enabled = True
    CMDPRINT.Enabled = False
    
    CmdPrintA5.Enabled = False
    CMDEXIT.Enabled = True
    TxtName1.Enabled = False
    FRMEHEAD.Enabled = True
    TXTDEALER.Enabled = True
    '
    'TXTDEALER.SetFocus
    LBLITEMCOST.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTQTY.Tag = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblcredit.Caption = "1"
    TxtFrieght.Text = ""
    Txthandle.Text = ""
    TxtOrder.Text = ""
    TxtVehicle.Text = ""
    txtOutstanding.Text = ""
    TXTTIN.Text = ""
    lblActAmt.Caption = ""
    cr_days = False
    M_ADD = False
    TXTDEALER.Text = ""
    'TXTTYPE.Text = ""
    'cmbtype.ListIndex = -1
    
    
    TXTDEALER.SetFocus
    'TxtBillName.SetFocus
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub TxtFree_GotFocus()
    TXTFREE.SelStart = 0
    TXTFREE.SelLength = Len(TXTFREE.Text)
    TXTFREE.Tag = Trim(TXTPRODUCT.Text)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(TXTFREE.Text) = 0 Then GoTo SKIP
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT CLOSE_QTY  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!CLOSE_QTY)) Then RSTTRXFILE!CLOSE_QTY = 0
                i = RSTTRXFILE!CLOSE_QTY / Val(LblPack.Text)
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            If i > 0 Then
                If Val(TXTFREE.Text) + Val(TXTQTY.Text) > i Then
                    MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                    TXTFREE.SelStart = 0
                    TXTFREE.SelLength = Len(TXTFREE.Text)
                    Exit Sub
                End If
            End If
SKIP:
            TXTFREE.Enabled = False
            If UCase(TXTITEMCODE.Text) = "OT" Then
                TxtMRP.Enabled = True
                TxtMRP.SetFocus
            Else
                txtretail.Enabled = True
                txtretail.SetFocus
            End If
'            TXTFREE.Enabled = False
'            TXTTAX.Enabled = True
'            TXTTAX.SetFocus
         Case vbKeyEscape
            TXTFREE.Enabled = False
            If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
            TXTQTY.Enabled = True
            
            'TXTQTY.SetFocus
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
End Sub

Private Sub TXTTOTALDISC_GotFocus()
    TXTTOTALDISC.SelStart = 0
    TXTTOTALDISC.SelLength = Len(TXTTOTALDISC.Text)
End Sub

Private Sub TXTTOTALDISC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            'If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            '''If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
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
    TXTAMOUNT.Text = 0
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    LBLPROFIT.Caption = Round(Val(lblnetamount.Caption) - Val(LBLTOTALCOST.Caption), 2)
    
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
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
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
        If Trim(TXTTIN.Text) <> "" Then
            Print #1, Space(20) & "FORM NO. 8 See rule 58(10), TAX INVOICE"
        Else
            Print #1, Space(20) & "FORM NO. 8B See rule 58(10), RETAIL INVOICE"
        End If
        Print #1, Space(32) & AlignLeft("CASH / CREDIT SALE", 25)
        Print #1, RepeatString("-", 80)
        Print #1, "D.N. NO & Date" & Space(5) & "P.O. NO. & Date" & Space(5) & "D.Doc.NO & Date" & Space(5) & "Del Terms" & Space(5) & "Veh. No"
        Print #1,
        Print #1, RepeatString("-", 80)
        'Print #1, Chr(27) & Chr(71) & Chr(10) & Space(41) & AlignLeft("INVOICE FORM 8H", 16)
    
        'If Weekday(Date) = 1 Then LBLDATE.Caption = DateAdd("d", 1, LBLDATE.Caption)
        Print #1, "Bill No. " & Trim(LBLBILLNO.Caption) & Space(2) & AlignRight("Date:" & TXTINVDATE.Text, 67) '& Space(2) & LBLTIME.Caption
        Print #1, "TO: " & TXTDEALER.Text
        If Trim(lbladdress.Caption) <> "" Then Print #1, lbladdress.Caption
        If Trim(TXTTIN.Text) <> "" Then Print #1, "TIN: " & TXTTIN.Text
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
                Chr(27) & Chr(72)  '//Bold Ends
    
        Print #1, RepeatString("-", 80)
    
        For i = 1 To grdsales.rows - 1
            Print #1, AlignLeft(grdsales.TextMatrix(i, 2), 22) & Space(9) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 4) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)), 2), "0.00"), 7) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 9)), 2), "0.00"), 7) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 6)) * Val(grdsales.TextMatrix(i, 9)) / 100, 2), "0.00"), 7) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 10) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 12) '& _
                Chr(27) & Chr(72)  '//Bold Ends
        Next i
    
        Print #1, AlignRight("-------------", 80)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
            Print #1, AlignRight("DISC AMOUNT ", 65) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
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
        If Trim(TXTTIN.Text) <> "" Then
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

Private Sub TXTRETAIL_GotFocus()
'    If M_EDIT = False Then
'        If Val(LBLITEMCOST.Caption) <> 0 Then txtretail.Text = Round(Val(LBLITEMCOST.Caption) + (Val(LBLITEMCOST.Caption) * 10 / 100), 3)
'    End If
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    txtretail.SelStart = 0
    txtretail.SelLength = Len(txtretail.Text)
    'If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtretail.Text) = 0 Then Exit Sub
'            If Val(TXTTAX.Text) = 0 Then
'                MsgBox "Please enter the Tax", vbOKOnly, "Sales"
'                Exit Sub
'            End If
            txtretail.Enabled = False
            TXTDISC.Enabled = True
            TXTDISC.SetFocus
            'cmdadd.Enabled = True
            'cmdadd.SetFocus
        Case vbKeyEscape
'            If UCase(txtcategory.text) = "SERVICE CHARGE" Then
'                If M_EDIT = True Then Exit Sub
'                txtretail.Enabled = False
'                TXTPRODUCT.Enabled = True
'                TXTPRODUCT.SetFocus
'            Else
                txtretail.Enabled = False
                TXTRETAILNOTAX.Enabled = True
                TXTRETAILNOTAX.SetFocus
'            End If
'        Case 116
'            Call FILL_PREVIIOUSRATE
        Case 117
            If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
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

Private Sub TXTRETAILNOTAX_LostFocus()
    TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.0000")
    ''If lblP_Rate.Caption = "0" Then
    If Val(TXTRETAILNOTAX.Text) <> 0 Then
        If OPTTaxMRP.Value = True Then
            txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + Val(txtmrpbt.Text) * Val(TXTTAX.Text) / 100, 4)
        End If
        If OPTVAT.Value = True Then
            txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100, 4)
        End If
        If optnet.Value = True Then
            txtretail.Text = TXTRETAILNOTAX.Text
        End If
        TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.0000")
        If TxtRetailmode.Text = "A" Then
            txtcommi.Text = Format(Round(Val(txtretaildummy.Text), 2), "0.00")
        Else
            txtcommi.Text = Format(Round((Val(TXTRETAILNOTAX.Text) * Val(txtretaildummy.Text) / 100), 2), "0.00")
        End If
    End If
End Sub

Private Sub TXTRETAILNOTAX_GotFocus()
    TXTRETAILNOTAX.SelStart = 0
    TXTRETAILNOTAX.SelLength = Len(TXTRETAILNOTAX.Text)
    'If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
End Sub

Private Sub TXTRETAILNOTAX_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTRETAILNOTAX.Text) = 0 Then Exit Sub
            TXTRETAILNOTAX.Enabled = False
            txtretail.Enabled = True
            txtretail.SetFocus
        Case vbKeyEscape
            TXTRETAILNOTAX.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
'        Case 116
'            Call FILL_PREVIIOUSRATE
        Case 117
            If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
    End Select
End Sub

Private Sub TXTRETAILNOTAX_KeyPress(KeyAscii As Integer)
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
'    If Val(txtretail.Text) = 0 Then
'        optnet.value = True
'        TXTTAX.Text = 0
'    End If
    If OPTVAT.Value = False Then TXTTAX.Text = 0
    TXTRETAILNOTAX.Text = Round(Val(txtretail.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
    TXTRETAILNOTAX.Text = Format(Val(TXTRETAILNOTAX.Text), "0.0000")
    txtretail.Text = Format(Val(txtretail.Text), "0.0000")
    
    If Val(LBLITEMCOST.Caption) <> 0 Then
        LblProfitPerc.Caption = Round(((Val(txtretail.Text) - Val(LBLITEMCOST.Caption)) * 100) / Val(LBLITEMCOST.Caption), 2)
        LblProfitPerc.Caption = Format(Val(LblProfitPerc.Caption), "0.00")
    End If
    
    LblProfitAmt.Caption = Round((Val(txtretail.Text) - Val(LBLITEMCOST.Caption)) * Val(TXTQTY.Text), 2)
    LblProfitAmt.Caption = Format(Val(LblProfitAmt.Caption), "0.00")
    
    If TxtRetailmode.Text = "A" Then
        txtcommi.Text = Format(Round(Val(txtretaildummy.Text), 2), "0.00")
    Else
        txtcommi.Text = Format(Round((Val(TXTRETAILNOTAX.Text) * Val(txtretaildummy.Text) / 100), 2), "0.00")
    End If
    'TXTDISC.Tag = 0
    'TXTDISC.Tag = Val(TXTQTY.Text) * Val(TXTRETAILNOTAX.Text) * Val(TXTDISC.Text) / 100
    'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRETAILNOTAX.Text), 3)) - Val(TXTDISC.Tag), ".000")
End Sub

Private Sub txtcommi_GotFocus()
    If Val(txtcommi.Text) = 0 Then txtcommi.Text = ""
    txtcommi.SelStart = 0
    txtcommi.SelLength = Len(txtcommi.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub txtcommi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If txtcommi.Text = "" Then Exit Sub
            If Val(txtcommi.Text) > Val(txtretail.Text) Then
                MsgBox "Commission Rate greater than actual Rate", vbOKOnly, "Sales"
                txtcommi.SetFocus
                Exit Sub
            End If
            txtcommi.Enabled = False
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Call TXTDISC_LostFocus
            Call CMDADD_Click
'            TXTDISC.Enabled = False
'            cmdadd.Enabled = True
'            cmdadd.SetFocus
'            'TxtWarranty.Enabled = True
'            'TxtWarranty.SetFocus
        Case vbKeyEscape
            TXTDISC.Enabled = True
            txtcommi.Enabled = False
            TXTDISC.SetFocus
    End Select
End Sub

Private Sub txtcommi_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcommi_LostFocus()
    txtcommi.Text = Format(txtcommi.Text, ".000")
End Sub

'Function FILL_PREVIIOUSRATE()
'    Set GRDPRERATE.DataSource = Nothing
'
'    If PRERATE_FLAG = True Then
'        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILE  WHERE TRX_TYPE ='PR' AND ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
'        PRERATE_FLAG = False
'    Else
'        PHY_PRERATE.Close
'        PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILE  WHERE TRX_TYPE ='PR' AND ITEM_CODE = '" & TXTITEMCODE.text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
'        PRERATE_FLAG = False
'    End If
'
'    If PHY_PRERATE.RecordCount > 0 Then
'        FRMEMAIN.Enabled = False
'        fRMEPRERATE.Visible = True
'        Set GRDPRERATE.DataSource = PHY_PRERATE
'        GRDPRERATE.Columns(0).Caption = "ITEM CODE"
'        GRDPRERATE.Columns(1).Caption = "OUTWARD"
'        GRDPRERATE.Columns(2).Caption = "DATE"
'        GRDPRERATE.Columns(3).Caption = "SOLD QTY"
'        GRDPRERATE.Columns(4).Caption = "RATE"
'        GRDPRERATE.Columns(5).Caption = "CUSTOMER"
'        GRDPRERATE.Columns(6).Caption = "INV NO"
'
'        GRDPRERATE.Columns(0).Visible = False
'        GRDPRERATE.Columns(1).Width = 3500
'        GRDPRERATE.Columns(2).Width = 1300
'        GRDPRERATE.Columns(3).Width = 1200
'        GRDPRERATE.Columns(4).Width = 1500
'        GRDPRERATE.Columns(5).Visible = False
'        GRDPRERATE.Columns(6).Width = 1400
'
'
'        GRDPRERATE.SetFocus
'        LBLHEAD(2).Caption = GRDPRERATE.Columns(7).text
'    End If
'End Function

Private Sub TxtItemcode_GotFocus()
    LBLITEMCOST.Caption = ""
    LblProfitPerc.Caption = ""
    LblProfitAmt.Caption = ""
    LBLSELPRICE.Caption = ""
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
    SERIAL_FLAG = False
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTBATCH As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            M_STOCK = 0
            'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If Trim(TXTITEMCODE.Text) = "" Then
                TXTPRODUCT.Enabled = True
                TXTITEMCODE.Enabled = False
                TXTPRODUCT.SetFocus
                Exit Sub
            End If
            'cmddelete.Enabled = False
            TXTQTY.Text = ""
            TXTEXPIRY.Text = "  /  "
            TXTAPPENDQTY.Text = ""
            TXTFREEAPPEND.Text = ""
            txtappendcomm.Text = ""
            TXTAPPENDTOTAL.Text = ""
            txtretail.Text = ""
            txtBatch.Text = ""
            TxtWarranty.Text = ""
            TxtWarranty_type.Text = ""
            TXTRETAILNOTAX.Text = ""
            TXTSALETYPE.Text = ""
            TXTFREE.Text = ""
            optnet.Value = True
            TxtMRP.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            TXTTAX.Text = ""
            TXTDISC.Text = ""
            txtcommi.Text = ""
            LBLSUBTOTAL.Caption = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
            
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, ITEM_COST, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP  From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.Text & "' ", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, ITEM_COST, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP  From ITEMMAST  WHERE ITEM_CODE = '" & Me.TXTITEMCODE.Text & "' ", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            Set grdtmp.DataSource = PHY
            
            If PHY.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "Sales"
                Exit Sub
            End If
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                TxtMRP.Text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
'                Set RSTBATCH = New ADODB.Recordset
'                RSTBATCH.Open "Select * From RTRXFILE WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockReadOnly
'                'If Not (RSTBATCH.EOF Or RSTBATCH.BOF) Then
'                    If RSTBATCH.RecordCount >= 1 Then
'                        Call FILL_BATCHGRID
'                        RSTBATCH.Close
'                        Set RSTBATCH = Nothing
'                        Exit Sub
'                    End If
'                'End If
'                RSTBATCH.Close
'                Set RSTBATCH = Nothing
                Call CONTINUE
            Else
                Call FILL_ITEMGRID
                Exit Sub
            End If
JUMPNONSTOCK:
            TXTPRODUCT.Enabled = False
            TXTITEMCODE.Enabled = False
            If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
                TXTTAX.Enabled = True
                TXTTAX.SetFocus
            Else
                If fRMEPRERATE.Visible = False Then Call FILL_PREVIIOUSRATE2
                TXTQTY.Enabled = True
                
                'TXTQTY.SetFocus
            End If
        Case vbKeyEscape
            If grdsales.rows > 1 Then
                TxtName1.Enabled = False
                
                CMDPRINT.Enabled = True
                CmdPrintA5.Enabled = True
                cmdRefresh.Enabled = True
                CmdPrintA5.SetFocus
            Else
                TxtName1.Enabled = False
                FRMEHEAD.Enabled = True
                TXTDEALER.Enabled = True
                TXTDEALER.SetFocus
            End If
            LBLDNORCN.Caption = ""
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TxtItemcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Function FILL_BATCHGRID()
    FRMEMAIN.Enabled = False
    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, ITEM_COST, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, EXP_DATE, PINV, VCH_DATE, TRX_YEAR From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, VCH_NO, LINE_NO, TRX_TYPE, VCH_DATE, ITEM_NAME, WARRANTY, WARRANTY_TYPE, ITEM_COST, P_WS, P_VAN, P_CRTN, CATEGORY, LOOSE_PACK, PACK_TYPE, COM_FLAG, COM_PER, COM_AMT, SALES_TAX, LINE_DISC, MRP, CRTN_PACK, EXP_DATE, PINV, VCH_DATE, TRX_YEAR From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "Serial No."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "VCH No"
    GRDPOPUP.Columns(3).Caption = "Line No"
    GRDPOPUP.Columns(4).Caption = "Trx Type"
    GRDPOPUP.Columns(7).Caption = "Warranty"
    GRDPOPUP.Columns(8).Caption = ""
    
    GRDPOPUP.Columns(0).Width = 4100
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 0
    GRDPOPUP.Columns(3).Width = 0
    GRDPOPUP.Columns(4).Width = 0
    GRDPOPUP.Columns(5).Width = 0
    GRDPOPUP.Columns(6).Width = 0
    GRDPOPUP.Columns(7).Width = 800
    'GRDPOPUP.Columns(7).Alignment
    GRDPOPUP.Columns(8).Width = 1200
    
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).Text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
    
    
End Function

Function FILL_PREVIIOUSRATE2()
    Set grdtmp.DataSource = Nothing
    grdtmp.Visible = False
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select ITEM_CODE, ITEM_NAME, VCH_DATE, PINV, QTY, BAL_QTY, REF_NO, EXP_DATE, MRP, ITEM_COST, DISC_FLAG, P_DISC, LOOSE_PACK, SALES_TAX, M_USER_ID, VCH_NO, LINE_NO, TRX_TYPE, CATEGORY, PACK_TYPE, LINE_DISC, TRX_YEAR From RTRXFILE  WHERE (TRX_TYPE ='PI' OR TRX_TYPE ='PW') AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select ITEM_CODE, ITEM_NAME, VCH_DATE, PINV, QTY, BAL_QTY, REF_NO, EXP_DATE, MRP, ITEM_COST, DISC_FLAG, P_DISC, LOOSE_PACK, SALES_TAX, M_USER_ID, VCH_NO, LINE_NO, TRX_TYPE, CATEGORY, PACK_TYPE, LINE_DISC, TRX_YEAR From RTRXFILE  WHERE (TRX_TYPE ='PI' OR TRX_TYPE ='PW') AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND M_USER_ID = '" & DataList2.BoundText & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        'PHY_PRERATE.Open "Select ITEM_CODE, VCH_DESC, VCH_DATE, QTY, P_RETAIL, M_USER_ID, VCH_NO, ITEM_NAME  From TRXFILE  WHERE TRX_TYPE ='PR' AND ITEM_CODE = '" & TXTITEMCODE.Text & "' AND VCH_NO <> " & Val(txtBillNo.Text) & " ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    End If
    
    If PHY_PRERATE.RecordCount > 0 Then
        'FRMEMAIN.Enabled = False
        fRMEPRERATE.Visible = True
        Set GRDPRERATE.DataSource = PHY_PRERATE
        GRDPRERATE.Columns(0).Caption = "ITEM CODE"
        GRDPRERATE.Columns(1).Caption = "ITEM NAME"
        GRDPRERATE.Columns(2).Caption = "INV DATE"
        GRDPRERATE.Columns(3).Caption = "INV NO."
        GRDPRERATE.Columns(4).Caption = "RCVD QTY"
        GRDPRERATE.Columns(5).Caption = "BAL QTY"
        GRDPRERATE.Columns(6).Caption = "SERIAL NO."
        GRDPRERATE.Columns(7).Caption = "EXP DATE"
        GRDPRERATE.Columns(8).Caption = "MRP"
        GRDPRERATE.Columns(9).Caption = "COST"
        GRDPRERATE.Columns(10).Caption = "DISC FLAG"
        GRDPRERATE.Columns(11).Caption = "DISC"
        GRDPRERATE.Columns(12).Caption = "PACK"
        GRDPRERATE.Columns(13).Caption = "TAX"
        GRDPRERATE.Columns(14).Caption = "SUPPLIER"
        GRDPRERATE.Columns(15).Caption = "COMP REF NO."
        
        GRDPRERATE.Columns(0).Visible = False
        'GRDPRERATE.Columns(1).Width = 2000
        GRDPRERATE.Columns(1).Visible = False
        GRDPRERATE.Columns(2).Width = 1200
        GRDPRERATE.Columns(3).Width = 1200
        GRDPRERATE.Columns(4).Width = 950
'        GRDPRERATE.Columns(5).Visible = False
'        GRDPRERATE.Columns(7).Visible = False
        GRDPRERATE.Columns(5).Width = 950
        GRDPRERATE.Columns(6).Width = 1100
        GRDPRERATE.Columns(7).Width = 1100
        GRDPRERATE.Columns(8).Width = 950
        GRDPRERATE.Columns(9).Width = 950
        GRDPRERATE.Columns(10).Visible = False
        GRDPRERATE.Columns(11).Width = 900
        GRDPRERATE.Columns(12).Width = 700
        GRDPRERATE.Columns(13).Width = 900
        'GRDPRERATE.Columns(11).Width = 1200
        GRDPRERATE.Columns(14).Visible = False
        GRDPRERATE.Columns(15).Width = 1200
        GRDPRERATE.Columns(16).Visible = False
        GRDPRERATE.Columns(17).Visible = False
        'GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(1).Text
        FRMEMAIN.Enabled = False
        GRDPRERATE.SetFocus
    Else
        FRMEMAIN.Enabled = True
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
    End If
End Function

Private Sub TxtVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub TxtWarranty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.Text) = 0 Then
                TxtWarranty.Enabled = False
                cmdadd.Enabled = True
                cmdadd.SetFocus
            Else
                TxtWarranty.Enabled = False
                TxtWarranty_type.Enabled = True
                TxtWarranty_type.SetFocus
            End If
        Case vbKeyEscape
            TXTDISC.Enabled = True
            TxtWarranty.Enabled = False
            TXTDISC.SetFocus
    End Select
End Sub

Private Sub TxtWarranty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtWarranty_type_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.Text) <> 0 And Trim(TxtWarranty_type.Text) = "" Then
                MsgBox "Please enter Period for Warranty", , "Sales"
                TxtWarranty_type.SetFocus
                Exit Sub
            End If
            If Val(TxtWarranty.Text) = 0 Then TxtWarranty_type.Text = ""
            TxtWarranty_type.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            TxtWarranty.Enabled = True
            TxtWarranty_type.Enabled = False
            TxtWarranty.SetFocus
    End Select
End Sub

Private Sub TxtWarranty_type_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
    rstBILL.Open "Select MAX(VCH_NO) FROM PURCAHSERETURN WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PR'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    LBLBILLNO.Caption = Val(txtBillNo.Text)
    
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
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
        Print #1, AlignLeft("ESTIMATE", 25)
        Print #1, RepeatString("-", 80)
        Print #1, AlignLeft("Sl", 2) & Space(1) & _
                AlignLeft("Comm Code", 14) & Space(1) & _
                AlignLeft("Description", 35) & _
                AlignLeft("Qty", 4) & Space(3) & _
                AlignLeft("Rate", 10) & Space(3) & _
                AlignLeft("Amount", 12) '& _
                Chr(27) & Chr(72)  '//Bold Ends
    
        Print #1, RepeatString("-", 80)
    
        For i = 1 To grdsales.rows - 1
            Print #1, AlignLeft(Val(i), 3) & _
                Space(15) & AlignLeft(grdsales.TextMatrix(i, 2), 34) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 4) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 9) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 13) '& _
                Chr(27) & Chr(72)  '//Bold Ends
        Next i
    
        Print #1, AlignRight("-------------", 80)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
            Print #1, AlignRight("DISC AMOUNT ", 65) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 65) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
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
        If Trim(TXTTIN.Text) <> "" Then
            Print #1, "Certified that all the particulars shown in the above Tax Invoice are true and correct"
            Print #1, "and that my/our Registration under KVAT ACT 2003 is valid as on the date of this bill"
            Print #1, RepeatString("-", 80)
        End If
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

Private Function ReportGeneratION_vpestimate()
    
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
        Print #1, AlignLeft("ESTIMATE", 25)
        Print #1, RepeatString("-", 55)
        Print #1, AlignLeft("Sl", 2) & Space(1) & _
                AlignLeft("Comm Code", 10) & Space(0) & _
                AlignLeft("Description", 22) & _
                AlignLeft("Qty", 4) & Space(3) & _
                AlignLeft("Rate", 7) & Space(0) & _
                AlignLeft("Amount", 9) '& _
                Chr(27) & Chr(72)  '//Bold Ends
    
        Print #1, RepeatString("-", 55)
    
        For i = 1 To grdsales.rows - 1
            'Print #1, AlignLeft(Val(i), 3) & _
                Space(10) & AlignLeft(Mid(grdsales.TextMatrix(i, 2), 1, 21), 21) & _
                AlignRight(Round(grdsales.TextMatrix(i, 3), 2), 4) & _
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 7)), 2), "0.00"), 8) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 12)), "0.00"), 9) '& _
                Chr(27) & Chr(72)  '//Bold Ends
        Next i
    
        Print #1, AlignRight("-------------", 55)
        If Val(LBLDISCAMT.Caption) <> 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 42) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
            Print #1, AlignRight("DISC AMOUNT ", 42) & AlignRight((Format(LBLDISCAMT.Caption, "####.00")), 12)
        ElseIf Val(LBLDISCAMT.Caption) = 0 Then
            Print #1, AlignRight("BILL AMOUNT ", 42) & AlignRight((Format(LBLTOTAL.Caption, "####.00")), 12)
        End If
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 55) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, AlignRight("Round off ", 42) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
        Print #1, Chr(13)
        Print #1, AlignRight("NET AMOUNT ", 42) & AlignRight((Format(Round(lblnetamount.Caption, 0), "####.00")), 12)
        'Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(lbltotalwodiscount.Caption) - Val(LBLRETAMT.Caption), "####.00")), 9)
        Num = CCur(Round(LBLTOTAL.Caption, 0))
        Print #1, AlignLeft("(Rupees " & Words_1_all(Num) & ")", 55)
        Print #1, RepeatString("-", 55)
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
    
    Close #1 '//Closing the file
    Exit Function

ERRHAND:
    MsgBox err.Description
End Function

Private Function CONTINUE_BATCH()

    M_STOCK = Val(GRDPOPUP.Columns(1))
    
    If M_STOCK <= 0 Then
        MsgBox "AVAILABLE STOCK IS  " & M_STOCK & " ", , "SALES"
        Exit Function
    End If
            
    Dim i As Double
    For i = 1 To grdsales.rows - 1
        If Trim(grdsales.TextMatrix(i, 13)) = Trim(TXTITEMCODE.Text) Then
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
    txtcategory.Text = IIf(IsNull(PHY!Category), "", PHY!Category)
    If UCase(txtcategory.Text) = "SERVICE CHARGE" Then
        TXTTAX.Enabled = True
        TXTTAX.SetFocus
        Exit Function
    End If
    'txtretail.Text = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
    lblwsale.Caption = IIf(IsNull(GRDPOPUP.Columns(10)), "", GRDPOPUP.Columns(10))
    lblvan.Caption = IIf(IsNull(GRDPOPUP.Columns(11)), "", GRDPOPUP.Columns(11))
    lblcase.Caption = IIf(IsNull(GRDPOPUP.Columns(12)), "", GRDPOPUP.Columns(12))
    lblcrtnpack.Caption = IIf(IsNull(GRDPOPUP.Columns(22)), "", GRDPOPUP.Columns(22))
    LblPack.Text = IIf(IsNull(GRDPOPUP.Columns(14)) Or GRDPOPUP.Columns(14) = "", "1", GRDPOPUP.Columns(14))
    lblOr_Pack.Caption = IIf(IsNull(GRDPOPUP.Columns(14)) Or GRDPOPUP.Columns(14) = "", "1", GRDPOPUP.Columns(14))
    If Val(LblPack.Text) = 0 Then LblPack.Text = 1
    TXTRETAILNOTAX.Text = IIf(IsNull(GRDPOPUP.Columns(9)), "", Val(LblPack.Text) * GRDPOPUP.Columns(9))
    lblretail.Caption = IIf(IsNull(GRDPOPUP.Columns(9)), "", Val(LblPack.Text) * GRDPOPUP.Columns(9))
    lblunit.Caption = IIf(IsNull(GRDPOPUP.Columns(15)), "Nos", GRDPOPUP.Columns(15))
    TxtWarranty.Text = IIf(IsNull(GRDPOPUP.Columns(7)), "", GRDPOPUP.Columns(7))
    TxtWarranty_type.Text = IIf(IsNull(GRDPOPUP.Columns(8)), "", GRDPOPUP.Columns(8))
    'txtinvnodate.Text = Trim(IIf(IsNull(GRDPOPUP.Columns(24)), "", GRDPOPUP.Columns(24))) '& " DTD " & IIf(IsDate(GRDPOPUP.Columns(25)), GRDPOPUP.Columns(25), ""))
    
    txtinvnodate.Text = Trim(IIf(IsNull(GRDPOPUP.Columns(24)), "", GRDPOPUP.Columns(24)))
    TxtInvoiceDate.Text = IIf(IsDate(GRDPOPUP.Columns(25)), Format(GRDPOPUP.Columns(25), "DD/MM/YYYY"), "  /  /    ")

    If GRDPOPUP.Columns(16) = "A" Then
        txtretaildummy.Text = IIf(IsNull(GRDPOPUP.Columns(18)), "P", GRDPOPUP.Columns(18))
        TxtRetailmode.Text = "A"
    Else
        txtretaildummy.Text = IIf(IsNull(GRDPOPUP.Columns(17)), "P", GRDPOPUP.Columns(17))
        TxtRetailmode.Text = "P"
    End If
    Select Case PHY!check_flag
        Case "M", "I"
            OPTTaxMRP.Value = True
            TXTTAX.Text = GRDPOPUP.Columns(19)
            TXTSALETYPE.Text = "2"
        Case "V"
            OPTVAT.Value = True
            TXTSALETYPE.Text = "2"
            TXTTAX.Text = GRDPOPUP.Columns(19)
        Case Else
            TXTSALETYPE.Text = "2"
            optnet.Value = True
            TXTTAX.Text = "0"
    End Select
    TXTUNIT.Text = GRDPOPUP.Columns(20)
                       
    'TXTPRODUCT.Enabled = False
    'TXTQTY.Enabled = True
    '
    'OptLoose.value = True
    'TXTQTY.SetFocus
    Exit Function
End Function

Private Function REMOVE_ITEM()
    Dim i As Long
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Function
      
    For i = Val(TXTSLNO.Text) To grdsales.rows - 2
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
    Next i
    grdsales.rows = grdsales.rows - 1
    
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLFOT.Caption = ""
    lblcomamt.Caption = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        Select Case grdsales.TextMatrix(i, 19)
            Case "CN"
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) - Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) - (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
            Case Else
                LBLTOTAL.Caption = Round(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 12)), 2)
                If Val(grdsales.TextMatrix(i, 20)) > 0 Then LBLFOT.Caption = Round(Val(LBLFOT.Caption) + (Val(grdsales.TextMatrix(i, 7)) - Val(grdsales.TextMatrix(i, 6))) * Val(grdsales.TextMatrix(i, 20)), 2)
                LBLFOT.Caption = ""
        End Select
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 24))
        Else
            lblcomamt.Caption = Val(lblcomamt.Caption) + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 24))
        End If
    Next i
    LBLTOTAL.Tag = Val(LBLTOTAL.Caption)
    TXTAMOUNT.Text = ""
    If OptDiscAmt.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(Val(TXTTOTALDISC.Text), 2)
    ElseIf OPTDISCPERCENT.Value = True And Val(TXTTOTALDISC.Text) > 0 Then
        TXTAMOUNT.Text = Round(((Val(LBLTOTAL.Caption) - Val(LBLFOT.Caption)) * Val(TXTTOTALDISC.Text) / 100), 2)
    End If
    LBLDISCAMT.Caption = Format(TXTAMOUNT.Text, "0.00")
    lblnetamount.Caption = Round(Val(LBLTOTAL.Caption) - (Val(TXTAMOUNT.Text) + Val(LBLRETAMT.Caption)), 2) + Val(LBLFOT.Caption) + Val(TxtFrieght.Text) + Val(Txthandle.Text)
    
    Call COSTCALCULATION
    Call Addcommission
    
    TXTSLNO.Text = Val(grdsales.rows)
    TXTPRODUCT.Text = ""
    txtcategory.Text = ""
    TxtName1.Text = ""
    TXTITEMCODE.Text = ""
    optnet.Value = True
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    TrxRYear.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTEXPIRY.Text = "  /  "
    TXTAPPENDQTY.Text = ""
    TXTFREEAPPEND.Text = ""
    txtappendcomm.Text = ""
    TXTAPPENDTOTAL.Text = ""
    txtretail.Text = ""
    txtBatch.Text = ""
    TxtWarranty.Text = ""
    TxtWarranty_type.Text = ""
    TXTTAX.Text = ""
    TXTRETAILNOTAX.Text = ""
    TXTSALETYPE.Text = ""
    TXTFREE.Text = ""
    TxtMRP.Text = ""
    txtinvnodate.Text = ""
    TxtInvoiceDate.Text = "  /  /    "
    txtmrpbt.Text = ""
    txtretaildummy.Text = ""
    txtcommi.Text = ""
    TxtRetailmode.Text = ""
    
    TXTDISC.Text = ""
    LBLSUBTOTAL.Caption = ""
    LBLDNORCN.Caption = ""
    cmdadd.Enabled = False
    'cmddelete.Enabled = False
    'CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    M_ADD = True
    TXTQTY.Enabled = False
    TXTITEMCODE.Enabled = True
    TXTITEMCODE.SetFocus
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
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 31  'ST_RATE
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(TXTsample.Text), "0.00")
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            grdsales.SetFocus
    End Select
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case grdsales.Col
        Case 31
             Select Case KeyAscii
                Case Asc("'")
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
    grdsales.SetFocus
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If grdsales.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113
            If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                Select Case grdsales.Col
                    Case 31
                        TXTsample.Visible = True
                        TXTsample.Top = grdsales.CellTop + 100
                        TXTsample.Left = grdsales.CellLeft + 0
                        TXTsample.Width = grdsales.CellWidth
                        TXTsample.Height = grdsales.CellHeight
                        TXTsample.Text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                        TXTsample.SetFocus
                End Select
            End If
    End Select
End Sub

Private Sub grdsales_Scroll()
    TXTsample.Visible = False
    grdsales.SetFocus
End Sub

Private Sub Txthandle_GotFocus()
    Txthandle.SelStart = 0
    Txthandle.SelLength = Len(Txthandle.Text)
End Sub

Private Sub Txthandle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            If TXTFREE.Enabled = True Then TXTFREE.SetFocus
            If TxtName1.Enabled = True Then TxtName1.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            'If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            'If TxtMRP.Enabled = True Then TxtMRP.SetFocus
            If TXTTAX.Enabled = True Then TXTTAX.SetFocus
            If TXTDISC.Enabled = True Then TXTDISC.SetFocus
            '''If txtcommi.Enabled = True Then txtcommi.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub Txthandle_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txthandle_LostFocus()
    Call TXTTOTALDISC_LostFocus
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
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
        Case vbKeyEscape
             If Len(Trim(TXTEXPIRY.Text)) = 1 Then GoTo Nextstep
            If Len(Trim(TXTEXPIRY.Text)) < 5 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
Nextstep:
            TXTEXPIRY.Enabled = False
            TxtMRP.Enabled = True
            TxtMRP.SetFocus
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    'TXTEXPIRY.Text = Format(TXTEXPIRY.Text, "MM/YY")
    'TXTEXPIRY.Visible = False
End Sub

Private Sub txtinvnodate_GotFocus()
    txtinvnodate.SelStart = 0
    txtinvnodate.SelLength = Len(txtinvnodate.Text)
End Sub

Private Sub txtinvnodate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtInvoiceDate.Enabled = True
            TxtInvoiceDate.SetFocus
        Case vbKeyEscape
            txtinvnodate.Enabled = False
            TxtInvoiceDate.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub txtinvnodate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub


