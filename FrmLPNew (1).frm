VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18585
   ControlBox      =   0   'False
   Icon            =   "FrmLPNew.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10995
   ScaleWidth      =   18585
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
      Left            =   17070
      TabIndex        =   174
      Top             =   6900
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
      Left            =   17085
      TabIndex        =   173
      Top             =   6630
      Width           =   1320
   End
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   2280
      TabIndex        =   117
      Top             =   3345
      Visible         =   0   'False
      Width           =   14820
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   2310
         Left            =   30
         TabIndex        =   118
         Top             =   270
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4075
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   13
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Index           =   2
         Left            =   3795
         TabIndex        =   120
         Top             =   15
         Width           =   11010
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   " PREVIOUS RATES FOR THE ITEM "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   1
         Left            =   30
         TabIndex        =   119
         Top             =   15
         Width           =   3780
      End
   End
   Begin VB.TextBox txtBillNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   76
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
      Height          =   435
      Left            =   4395
      TabIndex        =   48
      Top             =   7755
      Width           =   1155
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   2280
      TabIndex        =   63
      Top             =   2010
      Visible         =   0   'False
      Width           =   12420
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   3900
         Left            =   15
         TabIndex        =   64
         Top             =   15
         Width           =   12390
         _ExtentX        =   21855
         _ExtentY        =   6879
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
   End
   Begin VB.Frame Fram 
      BackColor       =   &H00D7F4F1&
      Caption         =   "Frame1"
      Height          =   11040
      Left            =   -135
      TabIndex        =   49
      Top             =   -90
      Width           =   18690
      Begin VB.ComboBox Cmbbarcode 
         Height          =   315
         ItemData        =   "FrmLPNew.frx":030A
         Left            =   14640
         List            =   "FrmLPNew.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   178
         Top             =   1155
         Width           =   2865
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<<&Previous"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12945
         TabIndex        =   154
         Top             =   135
         Width           =   1125
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Next>>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12945
         TabIndex        =   153
         Top             =   525
         Width           =   1125
      End
      Begin VB.CommandButton CmdTransfer 
         Caption         =   "Exp&ort Bill"
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
         Left            =   17490
         TabIndex        =   135
         Top             =   990
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00D7F4F1&
         Height          =   1575
         Left            =   150
         TabIndex        =   66
         Top             =   0
         Width           =   12795
         Begin VB.CheckBox Chktag 
            BackColor       =   &H00D7F4F1&
            Caption         =   "Print Tag"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   9720
            TabIndex        =   190
            Top             =   1260
            Width           =   1110
         End
         Begin VB.OptionButton OptDr 
            BackColor       =   &H00D7F4F1&
            Caption         =   "Debtors"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   75
            TabIndex        =   183
            Top             =   1230
            Width           =   1155
         End
         Begin VB.OptionButton OptCr 
            BackColor       =   &H00D7F4F1&
            Caption         =   "Creditors"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   60
            TabIndex        =   182
            Top             =   960
            Value           =   -1  'True
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
            TabIndex        =   102
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
            Left            =   12840
            TabIndex        =   74
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
            Height          =   315
            Left            =   6045
            MaxLength       =   150
            TabIndex        =   109
            Top             =   810
            Width           =   4815
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
            TabIndex        =   72
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
            Left            =   6450
            MaxLength       =   20
            TabIndex        =   105
            Top             =   135
            Width           =   2445
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   6450
            TabIndex        =   107
            Top             =   465
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
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
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1245
            TabIndex        =   103
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
         Begin MSDataListLib.DataCombo CMBPO 
            Height          =   1425
            Left            =   10875
            TabIndex        =   144
            Top             =   135
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   2514
            _Version        =   393216
            Appearance      =   0
            Style           =   1
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
         Begin MSMask.MaskEdBox TXTRCVDATE 
            Height          =   315
            Left            =   9075
            TabIndex        =   169
            Top             =   465
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
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
         Begin VB.Label lbllastdate 
            Height          =   150
            Left            =   9255
            TabIndex        =   194
            Top             =   210
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lbloldbills 
            Height          =   90
            Left            =   10620
            TabIndex        =   193
            Top             =   615
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "RCVD DATE"
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
            Left            =   7950
            TabIndex        =   170
            Top             =   495
            Width           =   1110
         End
         Begin MSForms.ComboBox CMBDISTRICT 
            Height          =   375
            Left            =   6045
            TabIndex        =   110
            Top             =   1140
            Width           =   3675
            VariousPropertyBits=   746604571
            ForeColor       =   255
            MaxLength       =   30
            DisplayStyle    =   3
            Size            =   "6482;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            DropButtonStyle =   0
            BorderColor     =   255
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "GODOWN"
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
            Index           =   2
            Left            =   5070
            TabIndex        =   161
            Top             =   1185
            Width           =   1290
         End
         Begin VB.Label lblcredit 
            Height          =   525
            Left            =   12960
            TabIndex        =   91
            Top             =   645
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "P.O No."
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
            Left            =   10155
            TabIndex        =   75
            Top             =   135
            Width           =   705
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
            Left            =   5070
            TabIndex        =   73
            Top             =   840
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
            TabIndex        =   71
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
            TabIndex        =   70
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
            Left            =   5070
            TabIndex        =   69
            Top             =   165
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
            Left            =   5070
            TabIndex        =   68
            Top             =   495
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
            TabIndex        =   67
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4440
         Left            =   150
         TabIndex        =   166
         Top             =   1485
         Width           =   18555
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
            Left            =   285
            TabIndex        =   167
            Top             =   1215
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4335
            Left            =   15
            TabIndex        =   168
            Top             =   90
            Width           =   18510
            _ExtentX        =   32650
            _ExtentY        =   7646
            _Version        =   393216
            Rows            =   1
            Cols            =   42
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            FocusRect       =   2
            HighLight       =   0
            AllowUserResizing=   3
            Appearance      =   0
            GridLineWidth   =   2
         End
      End
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00D7F4F1&
         Height          =   5160
         Left            =   150
         TabIndex        =   50
         Top             =   5835
         Width           =   18480
         Begin VB.TextBox TxtPoints 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   17490
            MaxLength       =   11
            TabIndex        =   191
            Top             =   465
            Width           =   885
         End
         Begin VB.CheckBox ChkFree 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "Free Warn"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   60
            TabIndex        =   189
            Top             =   1620
            Width           =   1065
         End
         Begin VB.Frame FRMEQTY 
            Caption         =   "Avaliable Qty"
            Height          =   750
            Left            =   14430
            TabIndex        =   184
            Top             =   2175
            Visible         =   0   'False
            Width           =   2595
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Same Barcode"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   63
               Left            =   1395
               TabIndex        =   188
               Top             =   165
               Width           =   1215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Same Item Code"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   62
               Left            =   75
               TabIndex        =   187
               Top             =   165
               Width           =   1215
            End
            Begin VB.Label lblbarqty 
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
               ForeColor       =   &H00004080&
               Height          =   360
               Left            =   1305
               TabIndex        =   186
               Top             =   345
               Width           =   1230
            End
            Begin VB.Label lblavlqty 
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
               ForeColor       =   &H00004080&
               Height          =   360
               Left            =   45
               TabIndex        =   185
               Top             =   345
               Width           =   1230
            End
         End
         Begin VB.TextBox TxtTotalexp 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   450
            Left            =   10500
            MaxLength       =   10
            TabIndex        =   162
            Top             =   2370
            Width           =   1155
         End
         Begin VB.TextBox TxtNetrate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   15195
            MaxLength       =   11
            TabIndex        =   17
            Top             =   465
            Width           =   960
         End
         Begin VB.TextBox TxTfree 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   11250
            MaxLength       =   8
            TabIndex        =   8
            Top             =   465
            Width           =   495
         End
         Begin VB.ComboBox cmbfull 
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
            ItemData        =   "FrmLPNew.frx":030E
            Left            =   9675
            List            =   "FrmLPNew.frx":0360
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   495
            Width           =   825
         End
         Begin VB.CommandButton CmdLabels 
            Caption         =   "Print &Labels"
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
            Left            =   17070
            TabIndex        =   155
            Top             =   1605
            Width           =   1335
         End
         Begin VB.TextBox TxtCustDisc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   9885
            MaxLength       =   7
            TabIndex        =   31
            Top             =   1140
            Width           =   1140
         End
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
            Height          =   420
            Left            =   17055
            TabIndex        =   150
            Top             =   2055
            Width           =   1335
         End
         Begin VB.TextBox TxtCessPer 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   12705
            MaxLength       =   7
            TabIndex        =   35
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   13365
            MaxLength       =   7
            TabIndex        =   36
            Top             =   1140
            Width           =   915
         End
         Begin VB.TextBox txtHSN 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   60
            MaxLength       =   15
            TabIndex        =   18
            Top             =   1125
            Width           =   960
         End
         Begin VB.TextBox TxtBarcode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   480
            MaxLength       =   20
            TabIndex        =   1
            Top             =   480
            Width           =   1785
         End
         Begin VB.TextBox TxtLWRate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   8865
            MaxLength       =   7
            TabIndex        =   30
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   34
            Top             =   1140
            Width           =   855
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
            Left            =   8070
            MaxLength       =   7
            TabIndex        =   42
            Top             =   4215
            Visible         =   0   'False
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
            Left            =   7110
            MaxLength       =   7
            TabIndex        =   41
            Top             =   4215
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox TxtExpense 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   3570
            MaxLength       =   7
            TabIndex        =   21
            Top             =   1140
            Width           =   840
         End
         Begin VB.TextBox txtcategory 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   2280
            TabIndex        =   2
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox TxtWarranty 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   14295
            MaxLength       =   4
            TabIndex        =   37
            Top             =   1140
            Width           =   315
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
            ItemData        =   "FrmLPNew.frx":03FC
            Left            =   14625
            List            =   "FrmLPNew.frx":0406
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1155
            Width           =   840
         End
         Begin VB.TextBox TxtRetailPercent 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Left            =   4425
            MaxLength       =   7
            TabIndex        =   23
            Top             =   1545
            Width           =   945
         End
         Begin VB.TextBox txtWsalePercent 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Left            =   5385
            MaxLength       =   7
            TabIndex        =   25
            Top             =   1545
            Width           =   870
         End
         Begin VB.TextBox txtSchPercent 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Left            =   6270
            MaxLength       =   7
            TabIndex        =   27
            Top             =   1545
            Width           =   900
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
            ItemData        =   "FrmLPNew.frx":0417
            Left            =   11760
            List            =   "FrmLPNew.frx":0469
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   495
            Width           =   780
         End
         Begin VB.TextBox Los_Pack 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   9225
            MaxLength       =   7
            TabIndex        =   5
            Top             =   480
            Width           =   435
         End
         Begin VB.TextBox Txtgrossamt 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   16170
            MaxLength       =   10
            TabIndex        =   16
            Top             =   465
            Width           =   1305
         End
         Begin VB.TextBox txtvanrate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   6270
            MaxLength       =   7
            TabIndex        =   26
            Top             =   1140
            Width           =   900
         End
         Begin VB.TextBox txtcrtnpack 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   7185
            MaxLength       =   7
            TabIndex        =   28
            Top             =   1140
            Width           =   705
         End
         Begin VB.TextBox TxtComper 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   11040
            MaxLength       =   7
            TabIndex        =   32
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   11775
            MaxLength       =   7
            TabIndex        =   33
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   7905
            MaxLength       =   7
            TabIndex        =   29
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   5385
            MaxLength       =   7
            TabIndex        =   24
            Top             =   1140
            Width           =   870
         End
         Begin VB.TextBox TXTRETAIL 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   4425
            MaxLength       =   7
            TabIndex        =   22
            Top             =   1140
            Width           =   945
         End
         Begin VB.TextBox txtPD 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   1035
            MaxLength       =   7
            TabIndex        =   19
            Top             =   1140
            Width           =   750
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
            TabIndex        =   98
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
            Left            =   15555
            MaxLength       =   7
            TabIndex        =   96
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
            Left            =   7815
            TabIndex        =   93
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
            Left            =   6735
            TabIndex        =   92
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
            TabIndex        =   85
            Top             =   2205
            Width           =   945
         End
         Begin VB.OptionButton OPTNET 
            BackColor       =   &H00D7F4F1&
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
            Left            =   15735
            TabIndex        =   83
            Top             =   3525
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00D7F4F1&
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
            Left            =   15720
            TabIndex        =   81
            Top             =   3045
            Width           =   1410
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00D7F4F1&
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
            Left            =   15735
            TabIndex        =   82
            Top             =   3285
            Width           =   1395
         End
         Begin VB.TextBox TxttaxMRP 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   14475
            MaxLength       =   7
            TabIndex        =   15
            Top             =   465
            Width           =   705
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
            Left            =   3375
            MaxLength       =   7
            TabIndex        =   77
            Top             =   3075
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   13470
            MaxLength       =   11
            TabIndex        =   11
            Top             =   465
            Width           =   990
         End
         Begin VB.TextBox TXTRATE 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   12540
            MaxLength       =   7
            TabIndex        =   10
            Top             =   465
            Width           =   915
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
            Height          =   435
            Left            =   60
            TabIndex        =   44
            Top             =   2010
            Width           =   1095
         End
         Begin VB.TextBox TXTSLNO 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   45
            TabIndex        =   0
            Top             =   480
            Width           =   420
         End
         Begin VB.TextBox TXTPRODUCT 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   3420
            TabIndex        =   3
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox TXTQTY 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   10500
            MaxLength       =   8
            TabIndex        =   7
            Top             =   465
            Width           =   735
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
            Height          =   435
            Left            =   2280
            TabIndex        =   46
            Top             =   2010
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
            Height          =   435
            Left            =   1185
            TabIndex        =   45
            Top             =   2010
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
            Left            =   60
            TabIndex        =   52
            Top             =   3075
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   12300
            MaxLength       =   35
            TabIndex        =   12
            Top             =   2685
            Visible         =   0   'False
            Width           =   1905
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
            TabIndex        =   51
            Top             =   3945
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdRefresh 
            BackColor       =   &H00000080&
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
            Height          =   435
            Left            =   3375
            TabIndex        =   47
            Top             =   2010
            Width           =   975
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   375
            Left            =   14220
            TabIndex        =   13
            Top             =   2535
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   661
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
            Height          =   375
            Left            =   14220
            TabIndex        =   14
            Top             =   2535
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   661
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
            BackColor       =   &H00D7F4F1&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   7185
            TabIndex        =   113
            Top             =   1455
            Width           =   2565
            Begin VB.OptionButton OptComAmt 
               BackColor       =   &H00D7F4F1&
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
               Left            =   1200
               TabIndex        =   40
               Top             =   195
               Width           =   1335
            End
            Begin VB.OptionButton OptComper 
               BackColor       =   &H00D7F4F1&
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
               Left            =   15
               TabIndex        =   39
               Top             =   180
               Value           =   -1  'True
               Width           =   1155
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00D7F4F1&
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   6345
            TabIndex        =   126
            Top             =   2475
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
               TabIndex        =   128
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
               TabIndex        =   127
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
               TabIndex        =   130
               Top             =   525
               Width           =   1470
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "TCS %"
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
               TabIndex        =   129
               Top             =   195
               Width           =   1050
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00D7F4F1&
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   1170
            TabIndex        =   121
            Top             =   1530
            Width           =   2070
            Begin VB.OptionButton optdiscper 
               BackColor       =   &H00D7F4F1&
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
               TabIndex        =   123
               Top             =   120
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton Optdiscamt 
               BackColor       =   &H00D7F4F1&
               Caption         =   "Disc Am&t"
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
               TabIndex        =   122
               Top             =   135
               Width           =   1125
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00D7F4F1&
            Height          =   2415
            Left            =   8970
            TabIndex        =   137
            Top             =   2880
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
         Begin MSDataListLib.DataCombo Cmbcategory 
            Height          =   360
            Left            =   7050
            TabIndex        =   4
            Top             =   495
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            ForeColor       =   255
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   270
            Index           =   64
            Left            =   16170
            TabIndex        =   192
            Top             =   195
            Width           =   1305
         End
         Begin VB.Label LBLNET 
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
            ForeColor       =   &H00004080&
            Height          =   450
            Left            =   12885
            TabIndex        =   177
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label LBLGROSSAMT 
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
            ForeColor       =   &H00008000&
            Height          =   450
            Left            =   13005
            TabIndex        =   176
            Top             =   1725
            Width           =   1470
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "GROSS AMT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   59
            Left            =   13005
            TabIndex        =   175
            Top             =   1500
            Width           =   1470
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL QTY"
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
            Height          =   225
            Index           =   58
            Left            =   11715
            TabIndex        =   172
            Top             =   2160
            Width           =   1080
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblqty 
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
            ForeColor       =   &H00004080&
            Height          =   450
            Left            =   11670
            TabIndex        =   171
            Top             =   2370
            Width           =   1200
         End
         Begin VB.Label LBLEXP 
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
            ForeColor       =   &H00800000&
            Height          =   450
            Left            =   15885
            TabIndex        =   165
            Top             =   1725
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL EXP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   57
            Left            =   15900
            TabIndex        =   164
            Top             =   1500
            Width           =   1140
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Enter total expenses here"
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
            Height          =   480
            Index           =   56
            Left            =   8955
            TabIndex        =   163
            Top             =   2355
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TAX AMOUNT"
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
            Height          =   210
            Index           =   55
            Left            =   14460
            TabIndex        =   160
            Top             =   1500
            Width           =   1470
            WordWrap        =   -1  'True
         End
         Begin VB.Label LBLTOTALTAX 
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
            ForeColor       =   &H00008000&
            Height          =   450
            Left            =   14490
            TabIndex        =   159
            Top             =   1725
            Width           =   1380
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
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   54
            Left            =   15195
            TabIndex        =   158
            Top             =   195
            Width           =   960
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
            Index           =   53
            Left            =   11760
            TabIndex        =   157
            Top             =   195
            Width           =   750
         End
         Begin VB.Label LBLPRE 
            Height          =   330
            Left            =   13275
            TabIndex        =   156
            Top             =   3780
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Cust Disc%"
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
            Index           =   52
            Left            =   9885
            TabIndex        =   152
            Top             =   885
            Width           =   1140
         End
         Begin VB.Label lblcategory 
            Height          =   345
            Left            =   15780
            TabIndex        =   151
            Top             =   2475
            Visible         =   0   'False
            Width           =   1935
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
            Left            =   12705
            TabIndex        =   149
            Top             =   885
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Adl. Cess"
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
            Height          =   255
            Index           =   49
            Left            =   13365
            TabIndex        =   148
            Top             =   885
            Width           =   915
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
            Left            =   2280
            TabIndex        =   147
            Top             =   195
            Width           =   1245
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
            Height          =   255
            Index           =   48
            Left            =   60
            TabIndex        =   146
            Top             =   885
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
            ForeColor       =   &H008080FF&
            Height          =   285
            Index           =   47
            Left            =   480
            TabIndex        =   145
            Top             =   195
            Width           =   1785
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
            Left            =   8865
            TabIndex        =   143
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
            Left            =   4170
            TabIndex        =   142
            Top             =   3045
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
            Left            =   4170
            TabIndex        =   141
            Top             =   2790
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Scheme"
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
            Left            =   1800
            TabIndex        =   140
            Top             =   885
            Width           =   855
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
            Left            =   8070
            TabIndex        =   139
            Top             =   3960
            Visible         =   0   'False
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
            Left            =   7110
            TabIndex        =   138
            Top             =   3960
            Visible         =   0   'False
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2670
            TabIndex        =   20
            Top             =   1155
            Width           =   870
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
            Left            =   3585
            TabIndex        =   136
            Top             =   885
            Width           =   825
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
            Left            =   14295
            TabIndex        =   132
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
            Left            =   3570
            TabIndex        =   131
            Top             =   1545
            Width           =   840
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
            Left            =   60
            TabIndex        =   125
            Top             =   2790
            Visible         =   0   'False
            Width           =   3300
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Loose Pack"
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
            Left            =   9225
            TabIndex        =   124
            Top             =   195
            Width           =   1260
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Points"
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
            Left            =   17490
            TabIndex        =   116
            Top             =   195
            Width           =   885
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
            Left            =   6270
            TabIndex        =   115
            Top             =   885
            Width           =   900
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
            Left            =   7185
            TabIndex        =   114
            Top             =   885
            Width           =   705
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
            Left            =   11040
            TabIndex        =   112
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
            Left            =   11775
            TabIndex        =   111
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
            Left            =   7905
            TabIndex        =   108
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
            Left            =   5385
            TabIndex        =   106
            Top             =   885
            Width           =   870
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
            Left            =   15555
            TabIndex        =   104
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
            Left            =   1035
            TabIndex        =   99
            Top             =   885
            Width           =   750
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
            Left            =   4425
            TabIndex        =   97
            Top             =   885
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TCS Amt"
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
            Left            =   7845
            TabIndex        =   95
            Top             =   1965
            Width           =   1020
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Round off"
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
            Left            =   6735
            TabIndex        =   94
            Top             =   1965
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
            Left            =   11595
            TabIndex        =   90
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
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   450
            Left            =   11385
            TabIndex        =   89
            Top             =   1725
            Width           =   1605
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
            TabIndex        =   88
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
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   450
            Left            =   9765
            TabIndex        =   87
            Top             =   1725
            Width           =   1605
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
            Left            =   9780
            TabIndex        =   86
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
            Left            =   11250
            TabIndex        =   84
            Top             =   195
            Width           =   495
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
            Height          =   270
            Index           =   13
            Left            =   2670
            TabIndex        =   80
            Top             =   885
            Width           =   885
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
            Left            =   14475
            TabIndex        =   79
            Top             =   195
            Width           =   705
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
            Left            =   3375
            TabIndex        =   78
            Top             =   2790
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
            Left            =   13470
            TabIndex        =   65
            Top             =   195
            Width           =   990
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
            TabIndex        =   62
            Top             =   195
            Width           =   420
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
            Left            =   3525
            TabIndex        =   61
            Top             =   195
            Width           =   3510
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
            Left            =   10500
            TabIndex        =   60
            Top             =   195
            Width           =   735
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
            Left            =   12540
            TabIndex        =   59
            Top             =   195
            Width           =   915
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
            Left            =   15465
            TabIndex        =   58
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
            TabIndex        =   57
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
            Left            =   14220
            TabIndex        =   56
            Top             =   2895
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Category"
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
            Index           =   7
            Left            =   7050
            TabIndex        =   55
            Top             =   195
            Width           =   2160
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
            Height          =   390
            Left            =   15465
            TabIndex        =   43
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
            TabIndex        =   54
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
            TabIndex        =   53
            Top             =   3615
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Add:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Index           =   61
         Left            =   14130
         TabIndex        =   181
         Top             =   120
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbladdress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   14595
         TabIndex        =   180
         Top             =   120
         Width           =   4035
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode Printer"
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
         Height          =   255
         Index           =   60
         Left            =   14640
         TabIndex        =   179
         Top             =   915
         Width           =   1620
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Index           =   50
         Left            =   12945
         TabIndex        =   134
         Top             =   900
         Width           =   1680
         WordWrap        =   -1  'True
      End
      Begin VB.Label LBLmonth 
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   12945
         TabIndex        =   133
         Top             =   1110
         Width           =   1680
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   135
         TabIndex        =   101
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   705
         TabIndex        =   100
         Top             =   45
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytData() As Byte
Dim ACT_PO As New ADODB.Recordset
Dim PO_FLAG As Boolean
Dim ACT_REC As New ADODB.Recordset
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim ACT_FLAG As Boolean
Dim PHY_CODE As New ADODB.Recordset
Dim PHYCODE_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT, M_ADD, OLD_BILL, NEW_BILL As Boolean
Dim PHY_PRERATE As New ADODB.Recordset
Dim PRERATE_FLAG As Boolean
Dim PONO As String
Dim CHANGE_FLAG As Boolean
Dim BARCODE_FLAG As Boolean
Dim ADDCLICK As Boolean
Dim BARPRINTER As String
Dim CAT_REC As New ADODB.Recordset

Private Sub Cmbbarcode_Click()
    BARPRINTER = Cmbbarcode.ListIndex
End Sub

Private Sub Cmbcategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            TXTQTY.SetFocus
         Case vbKeyEscape
             If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            Los_Pack.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub CMBDISTRICT_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            FRMECONTROLS.Enabled = True
            If CMBPO.VisibleCount = 0 Then
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            Else
                CMBPO.SetFocus
            End If
        Case vbKeyEscape
            TXTREMARKS.SetFocus
    End Select
End Sub

Private Sub cmbfull_GotFocus()
    cmbfull.BackColor = &H98F3C1
End Sub

Private Sub cmbfull_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If cmbfull.ListIndex = -1 Then cmbfull.ListIndex = 0
            TXTQTY.SetFocus
         Case vbKeyEscape
            'TXTUNIT.Text = ""
            Los_Pack.SetFocus
    End Select
End Sub

Private Sub cmbfull_LostFocus()
    cmbfull.BackColor = vbWhite
End Sub

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = "1"
            TXTRATE.SetFocus
         Case vbKeyEscape
            'TXTUNIT.Text = ""
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub CMBPO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMBPO.VisibleCount <> 0 And CMBPO.BoundText = "" Then
                If (MsgBox("Are you sure you want to continue without selecting the Purchase Order No.? !!!!", vbYesNo, "EzBiz") = vbNo) Then
                    CMBPO.SetFocus
                    Exit Sub
                End If
            End If
            If CMBPO.VisibleCount = 0 Then CMBPO.text = ""
            If CMBPO.text <> "" And CMBPO.MatchedWithList = False Then
                MsgBox "Please select a valid PO No. from the list", vbOKOnly, "EzBiz"
                On Error Resume Next
                CMBPO.SetFocus
                Exit Sub
            End If
            FRMECONTROLS.Enabled = True
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTREMARKS.SetFocus
    End Select
End Sub

Private Sub CMBPO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMDADD_Click()
    
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
    ADDCLICK = True
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
    If Val(Los_Pack.text) = 1 Then
         TxtLWRate.text = Val(txtWS.text)
         txtcrtn.text = Val(TXTRETAIL.text)
    End If
    
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
        
'    If MDIMAIN.StatusBar.Panels(14).text = "Y" Then
'        If optdiscper.Value = True Then
'            txtPD.Tag = Round((Val(TXTPTR.text) * Val(TXTQTY.text)) / (Val(TXTQTY.text) + Val(TXTFREE.text)), 3)
'            TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.text) / 100)) + ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.text) / 100)) * Val(TxttaxMRP.text) / 100)
'        Else
'            txtPD.Tag = Round((Val(TXTPTR.text) * Val(TXTQTY.text)) / (Val(TXTQTY.text) + Val(TXTFREE.text)), 3)
'            TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.text) / Val(TXTQTY.text))) + ((Val(txtPD.Tag) - (Val(txtPD.text) / Val(TXTQTY.text))) * Val(TxttaxMRP.text) / 100)
'        End If
'    Else
'        If optdiscper.Value = True Then
'            txtPD.Tag = Round((Val(TXTPTR.text) * Val(TXTQTY.text)) / (Val(TXTQTY.text) + Val(TXTFREE.text)), 3)
'            TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            txtPD.Tag = Round((Val(TXTPTR.text) * Val(TXTQTY.text)) / (Val(TXTQTY.text) + Val(TXTFREE.text)), 3)
'            TXTPTR.Tag = (Val(txtPD.Tag) - (Val(txtPD.text) / Val(TXTQTY.text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    End If
'
    TXTPTR.Tag = Val(TxtNetrate.text)
    If MDIMAIN.lblgst.Caption = "R" And MDIMAIN.StatusBar.Panels(14).text <> "Y" Then
        TXTPTR.Tag = Round(Val(TxtNetrate.text) * 100 / ((Val(TxttaxMRP.text)) + 100), 4)
    End If
    
    If Val(TXTRATE.text) <> 0 And Val(TXTRATE.text) < Val(TXTPTR.Tag) Then
        MsgBox "MRP less than cost", vbOKOnly, "Purchase....."
        TXTRATE.SetFocus
        Exit Sub
    End If
    
    If Val(TXTRATE.text) <> 0 And Val(TXTRETAIL.text) <> 0 And Val(TXTRETAIL.text) > Val(TXTRATE.text) Then
        MsgBox "Retail Price greater than MRP", vbOKOnly, "EzBiz"
        TXTRETAIL.SetFocus
        Exit Sub
    End If
    
    If Val(TXTRATE.text) <> 0 And Val(txtWS.text) <> 0 And Val(txtWS.text) > Val(TXTRATE.text) Then
        MsgBox "WS Price greater than MRP", vbOKOnly, "EzBiz"
        txtWS.SetFocus
        Exit Sub
    End If
    
    If Val(TXTRATE.text) <> 0 And Val(txtvanrate.text) <> 0 And Val(txtvanrate.text) > Val(TXTRATE.text) Then
        MsgBox "VAN Price greater than MRP", vbOKOnly, "EzBiz"
        txtvanrate.SetFocus
        Exit Sub
    End If
    
    If Val(TXTRETAIL.text) <> 0 And Val(TXTRETAIL.text) < Val(TXTPTR.Tag) Then
        If MsgBox("Retail Price less than cost. Are you sure?", vbYesNo + vbDefaultButton2, "Purchase.....") = vbNo Then
            TXTRETAIL.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtWS.text) <> 0 And Val(txtWS.text) < Val(TXTPTR.Tag) Then
        If MsgBox("WS Price less than cost. Are you sure?", vbYesNo + vbDefaultButton2, "Purchase.....") = vbNo Then
            txtWS.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtvanrate.text) <> 0 And Val(txtvanrate.text) < Val(TXTPTR.Tag) Then
        If MsgBox("Van Price less than cost. Are you sure?", vbYesNo + vbDefaultButton2, "Purchase.....") = vbNo Then
            txtvanrate.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(TXTRETAIL.text) <> 0 And Val(txtcrtn.text) <> 0 And Val(TXTRETAIL.text) < Val(txtcrtn.text) Then
        MsgBox "Retail Price less than Loose Price", vbOKOnly, "EzBiz"
        TXTRETAIL.SetFocus
        Exit Sub
    End If
    
    If Val(txtWS.text) <> 0 And Val(TxtLWRate.text) <> 0 And Val(txtWS.text) < Val(TxtLWRate.text) Then
        MsgBox "WS Price less than Loose Price", vbOKOnly, "EzBiz"
        txtWS.SetFocus
        Exit Sub
    End If
    
    'Call TXTPTR_LostFocus
    Call TXTQTY_LostFocus
    'Call Txtgrossamt_LostFocus
    Call txtPD_LostFocus
    Call txtcrtn_GotFocus
    Call TxtLWRate_GotFocus
    
    ADDCLICK = False
    txtcrtn.BackColor = vbWhite
    TxtLWRate.BackColor = vbWhite
    TXTRATE.BackColor = vbWhite
    
    Dim i As Single
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim M_DATA As Double
    
    M_DATA = 0
    Txtpack.text = 1
    If MDIMAIN.StatusBar.Panels(6).text = "Y" Then
        If Trim(TxtBarcode.text) = "" Or Trim(TXTITEMCODE.text) = Left(Trim(TxtBarcode.text), Len(Trim(TXTITEMCODE.text))) Then '(Trim(TxtBarcode.Text) = Trim(TXTITEMCODE.Text) & Val(LBLPRE.Caption)) Then
            TxtBarcode.text = Trim(TXTITEMCODE.text) & Val(TXTRETAIL.text)
'            If Trim(Txtsize.Text) = "" Then
'                TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text) & Trim(cmbcolor.Text)
'            Else
'
'            End If
'        Else
'            If Trim(Txtsize.Text) = "" Then
'                TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text) & Trim(cmbcolor.Text)
'            Else
'                TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text) & Left(Trim(Txtsize.Text), 2)
'            End If
        End If
    End If
    
    'If MDIMAIN.StatusBar.Panels(6).Text = "Y" Then If Trim(TxtBarcode.Text) = "" Then TxtBarcode.Text = Trim(TXTITEMCODE.Text) & Val(TXTRETAIL.Text)
    If grdsales.rows <= Val(TXTSLNO.text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.text), 0) = Val(TXTSLNO.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 1) = Trim(TXTITEMCODE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 2) = Trim(TXTPRODUCT.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 3) = Val(TXTQTY.text) + Val(TxTfree.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 4) = 1 'Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 5) = Val(Los_Pack.text) ' 1 'Val(TxtPack.Text)
    'grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Round(Val(TXTRATE.Text) / Val(Los_Pack.Text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 6) = Format(Val(TXTRATE.text), ".000")
    'grdsales.TextMatrix(Val(TXTSLNO.text), 8) = Format(Round(((Val(LblGross.Caption) / (Val(Los_Pack.text) * (Val(TXTQTY.text) + Val(TXTFREE.text)))) + ((Val(TxtExpense.text) / ((Val(TXTQTY.text) + Val(TXTFREE.text)) * Val(Los_Pack.text))))), 4), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 8) = Format(Round(Val(LblGross.Caption) / (Val(Los_Pack.text) * (Val(TXTQTY.text) + Val(TxTfree.text))), 4), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 9) = Format(Round(Val(TXTPTR.text) / Val(Los_Pack.text), 4), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 7) = Format((Val(txtprofit.text)), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 10) = IIf(Val(TxttaxMRP.text) = 0, "", Format(Val(TxttaxMRP.text), ".00")) 'TAX
    grdsales.TextMatrix(Val(TXTSLNO.text), 11) = Trim(txtBatch.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 12) = IIf(Trim(TXTEXPDATE.text) = "/  /", "", TXTEXPDATE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 13) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 14) = Val(TxTfree.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 17) = Val(txtPD.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 18) = Format(Val(TXTRETAIL.text), ".0000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 19) = Format(Val(txtWS.text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 25) = Format(Val(txtvanrate.text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 26) = Format(Val(Txtgrossamt.text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 20) = Format(Val(txtcrtn.text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.text), 37) = Format(Val(TxtLWRate.text), ".000")
    If OptComAmt.Value = True Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 21) = ""
        grdsales.TextMatrix(Val(TXTSLNO.text), 22) = Format(Val(TxtComAmt.text), ".00")
        grdsales.TextMatrix(Val(TXTSLNO.text), 23) = "A"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.text), 21) = Format(Val(TxtComper.text), ".00")
        grdsales.TextMatrix(Val(TXTSLNO.text), 22) = ""
        grdsales.TextMatrix(Val(TXTSLNO.text), 23) = "P"
    End If
    If optdiscper.Value = True Then
        grdsales.TextMatrix(Val(TXTSLNO.text), 27) = "P"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.text), 27) = "A"
    End If
    grdsales.TextMatrix(Val(TXTSLNO.text), 28) = Format(Val(Los_Pack.text), ".00")
    grdsales.TextMatrix(Val(TXTSLNO.text), 29) = Trim(CmbPack.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 30) = Val(TxtWarranty.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 31) = Trim(CmbWrnty.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 32) = Val(TxtExpense.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 33) = Val(TxtExDuty.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 34) = Val(TxtCSTper.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 35) = Val(TxtTrDisc.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 36) = Val(LblGross.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.text), 38) = Trim(TxtBarcode.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 39) = Val(txtCess.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 40) = Val(TxtCessPer.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 24) = Format(Val(txtcrtnpack.text), ".000")
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
    
    On Error GoTo ERRHAND
    'If OLD_BILL = False Then Call checklastbill
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    If OLD_BILL = False And Val(txtBillNo.text) <> 1 Then
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE VCH_NO= (SELECT MAX(VCH_NO) FROM TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PI')", db, adOpenStatic, adLockOptimistic, adCmdText
        txtBillNo.text = RSTTRXFILE!VCH_NO + 1
        db.BeginTrans
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "PI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = txtBillNo.text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE.Update
    Else
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "PI"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = txtBillNo.text
            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE.Update
        End If
    End If
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        RSTRTRXFILE.Properties("Update Criteria").Value = adCriteriaKey
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!TRX_TYPE = "PI"
        RSTRTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTRTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16))
        RSTRTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1))
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
        RSTRTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))

        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
'                If UCase(rststock!CATEGORY) = "CUTSHEET" Then
'                Else
                .Properties("Update Criteria").Value = adCriteriaKey
                !ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 8))
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                '!CLOSE_VAL = !CLOSE_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !CLOSE_VAL = Round(!ITEM_COST * !CLOSE_QTY, 3)
                !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                '!RCPT_VAL = !RCPT_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !RCPT_VAL = Round(!ITEM_COST * !RCPT_QTY, 3)
            
                !MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6))
                If Trim(txtHSN.text) <> "" Then !REMARKS = Trim(txtHSN.text)
                If cmbfull.ListIndex <> -1 Then !FULL_PACK = cmbfull.text
                !CUST_DISC = Val(TxtCustDisc.text)
                !SCH_POINTS = Val(TxtPoints.text)
                If !PRICE_CHANGE = "Y" And Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18)) <> 0 Then
                    db.Execute "Update RTRXFILE set P_RETAIL = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18)) & " WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 1) & "' AND BAL_QTY >0 "
                End If
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38)) <> "" Then !BARCODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18)) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 19)) <> 0 Then !P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 19))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20)) <> 0 Then !P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37)) <> 0 Then !P_LWS = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 25)) <> 0 Then !P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 25)) ' / Val(Los_Pack.Text), 3)
                'If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) <> 0 Then
                !cess_amt = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 39)) ' / Val(Los_Pack.Text), 3)
                'If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) <> 0 Then
                !CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 40)) ' / Val(Los_Pack.Text), 3)
                
                '!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24))) <> 0 Then !CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24))

                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 23)) = "A" Then
                    !COM_FLAG = "A"
                    !COM_PER = 0
                    !COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 22))
                Else
                    !COM_FLAG = "P"
                    !COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 21))
                    !COM_AMT = 0
                End If
                If ChkFree.Value = 0 Then
                    !FREE_WARN = "N"
                Else
                    !FREE_WARN = "Y"
                End If
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))) >= 5 Then !SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))
                '!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
                !check_flag = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15))
                !LOOSE_PACK = Val(Los_Pack.text)
                !PACK_TYPE = Trim(CmbPack.text)
                !WARRANTY = Val(TxtWarranty.text)
                !WARRANTY_TYPE = Trim(CmbWrnty.text)
                RSTRTRXFILE!MFGR = !MANUFACTURER
                RSTRTRXFILE!FOCUS_FLAG = !FOCUS_FLAG
                If Trim(Cmbcategory.text) <> "" Then !Category = Trim(Cmbcategory.text)
                    
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        
    Else
        RSTRTRXFILE.Properties("Update Criteria").Value = adCriteriaKey
        M_DATA = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
        M_DATA = M_DATA - (RSTRTRXFILE!QTY - RSTRTRXFILE!BAL_QTY)
        RSTRTRXFILE!BAL_QTY = M_DATA
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").Value = adCriteriaKey
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
                
                If Trim(txtHSN.text) <> "" Then !REMARKS = Trim(txtHSN.text)
                If cmbfull.ListIndex <> -1 Then !FULL_PACK = cmbfull.text
                !CUST_DISC = Val(TxtCustDisc.text)
                !SCH_POINTS = Val(TxtPoints.text)
                If !PRICE_CHANGE = "Y" And Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18)) <> 0 Then
                    db.Execute "Update RTRXFILE set P_RETAIL = " & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18)) & " WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 1) & "' AND BAL_QTY >0 "
                End If
                
                !MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38)) <> "" Then !BARCODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18)) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 19)) <> 0 Then !P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 19))
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20)) <> 0 Then !P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37)) <> 0 Then !P_LWS = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37)) ' / Val(Los_Pack.Text), 3)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.text), 25)) <> 0 Then !P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 25)) ' / Val(Los_Pack.Text), 3)
                'If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 39)) <> 0 Then
                !cess_amt = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 39)) ' / Val(Los_Pack.Text), 3)
                'If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 40)) <> 0 Then
                !CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 40)) ' / Val(Los_Pack.Text), 3)

                '!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24))) <> 0 Then !CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24))
                                    
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 23)) = "A" Then
                    !COM_FLAG = "A"
                    !COM_PER = 0
                    !COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 22))
                Else
                    !COM_FLAG = "P"
                    !COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 21))
                    !COM_AMT = 0
                End If
                If ChkFree.Value = 0 Then
                    !FREE_WARN = "N"
                Else
                    !FREE_WARN = "Y"
                End If
                If Val(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))) >= 5 Then !SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))
                '!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
                !check_flag = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15))
                !LOOSE_PACK = Val(Los_Pack.text)
                !PACK_TYPE = Trim(CmbPack.text)
                !WARRANTY = Val(TxtWarranty.text)
                !WARRANTY_TYPE = Trim(CmbWrnty.text)
                RSTRTRXFILE!MFGR = !MANUFACTURER
                RSTRTRXFILE!FOCUS_FLAG = !FOCUS_FLAG
                If Trim(Cmbcategory.text) <> "" Then !Category = Trim(Cmbcategory.text)
                
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
    End If
    
    If Trim(Cmbcategory.text) = "" Then
        RSTRTRXFILE!Category = "GENERAL"
    Else
        RSTRTRXFILE!Category = Trim(Cmbcategory.text)
    End If
    RSTRTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 13))
    RSTRTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "dd/mm/yyyy")
    If IsDate(TXTRCVDATE.text) Then
        RSTRTRXFILE!RCVD_DATE = Format(TXTRCVDATE.text, "DD/MM/YYYY")
    Else
        RSTRTRXFILE!RCVD_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    End If
    RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 2))
    RSTRTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 8))
    RSTRTRXFILE!ITEM_COST_PRICE = Round(Val(TXTPTR.text), 3)
    If (Val(TXTQTY.text) + Val(TxTfree.text)) = 0 Then
        RSTRTRXFILE!ITEM_NET_COST_PRICE = Round((Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.text)) + Val(TxtExpense.text), 3)
    Else
        RSTRTRXFILE!ITEM_NET_COST_PRICE = Round((Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.text)) + (Val(TxtExpense.text) / ((Val(TXTQTY.text) + Val(TxTfree.text)) * Val(Los_Pack.text))), 3)
    End If
    
    RSTRTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
    RSTRTRXFILE!P_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 17))
    RSTRTRXFILE!MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6))
    RSTRTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 9))
    RSTRTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 7))
    RSTRTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18))
    RSTRTRXFILE!P_WS = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 19))
    RSTRTRXFILE!P_CRTN = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20))
    RSTRTRXFILE!P_LWS = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37))
    RSTRTRXFILE!CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24))
    RSTRTRXFILE!P_VAN = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 25))
    RSTRTRXFILE!gross_amt = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 26))
    RSTRTRXFILE!BARCODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38))
    RSTRTRXFILE!cess_amt = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 39))
    RSTRTRXFILE!CESS_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 40))
    If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 23)) = "A" Then
        RSTRTRXFILE!COM_FLAG = "A"
        RSTRTRXFILE!COM_PER = 0
        RSTRTRXFILE!COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 22))
    Else
        RSTRTRXFILE!COM_FLAG = "P"
        RSTRTRXFILE!COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 21))
        RSTRTRXFILE!COM_AMT = 0
    End If
    RSTRTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 10))
    RSTRTRXFILE!LOOSE_PACK = Val(Los_Pack.text)
    RSTRTRXFILE!PACK_TYPE = Trim(CmbPack.text)
    RSTRTRXFILE!WARRANTY = Val(TxtWarranty.text)
    RSTRTRXFILE!WARRANTY_TYPE = Trim(CmbWrnty.text)
    RSTRTRXFILE!EXPENSE = Val(TxtExpense.text)
    RSTRTRXFILE!EXDUTY = Val(TxtExDuty.text)
    RSTRTRXFILE!CSTPER = Val(TxtCSTper.text)
    RSTRTRXFILE!TR_DISC = Val(TxtTrDisc.text)
    
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
    RSTRTRXFILE!SCHEME = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 14))
    'RSTRTRXFILE!EXP_DATE = IIf(grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = "", Null, Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "dd/mm/yyyy"))
    If IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 12)) Then
        RSTRTRXFILE!EXP_DATE = IIf(grdsales.TextMatrix(Val(TXTSLNO.text), 12) = "", Null, Format(grdsales.TextMatrix(Val(TXTSLNO.text), 12), "dd/mm/yyyy"))
    End If
    RSTRTRXFILE!FREE_QTY = 0
    RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!C_USER_ID = "SM"
    RSTRTRXFILE!check_flag = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15))
    
    'RSTRTRXFILE!M_USER_ID = DataList2.BoundText
    ''''RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))  'MODE OF TAX
    'RSTRTRXFILE!PINV = Trim(TXTINVOICE.Text)
    RSTRTRXFILE.Update
    db.CommitTrans
    RSTRTRXFILE.Close
    
    M_DATA = 0
    Set RSTRTRXFILE = Nothing
           
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    Dim GROSSVAL As Double
    For i = 1 To grdsales.rows - 1
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
        LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(i, 32)), ".00")
        lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(i, 3)), ".00")
        GROSSVAL = (Val(grdsales.TextMatrix(i, 9)) * IIf(Val(grdsales.TextMatrix(i, 5)) = 0, 1, Val(grdsales.TextMatrix(i, 5)))) * (Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 14)))
        If Trim(grdsales.TextMatrix(i, 27)) = "P" Then
            LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - (GROSSVAL * Val(grdsales.TextMatrix(i, 17)) / 100)) * Val(grdsales.TextMatrix(i, 10)) / 100, 2))
        ElseIf Trim(grdsales.TextMatrix(i, 27)) = "A" Then
            LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - Val(grdsales.TextMatrix(i, 17))) * Val(grdsales.TextMatrix(i, 10)) / 100, 2))
        End If
        LBLGROSSAMT.Caption = Val(LBLGROSSAMT.Caption) + Val(grdsales.TextMatrix(i, 8)) * Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
    Next i
    
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    If Roundflag = True Then
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    Else
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 2), "0.00")
    End If
    LBLNET.Caption = Val(LBLGROSSAMT.Caption) + Val(LBLTOTALTAX.Caption)
        
    Dim M, n, item_no As Integer
    Dim temp_file As String
    Dim ObjFile, objText, text
    If MDIMAIN.StatusBar.Panels(6).text = "Y" Then
        If MsgBox("Do you want to Print Barcode Labels now?", vbYesNo + vbDefaultButton2, "Purchase.....") = vbYes Then
            i = Val(InputBox("Enter number of lables to be print", "No. of labels..", Val(TXTQTY.text) + Val(TxTfree.text)))
            If i = 0 Then GoTo SKIP_BARCODE
            item_no = i
            If BARTEMPLATE = "Y" Then
                If Val(MDIMAIN.LBLLABELNOS.Caption) = 0 Then MDIMAIN.LBLLABELNOS.Caption = 1
                i = i / Val(MDIMAIN.LBLLABELNOS.Caption)
                If Math.Abs(i - Fix(i)) > 0 Then
                    i = Int(i) + 1
                End If
                If Chktag.Value = 0 Then
                    temp_file = "\template.txt"
                Else
                    temp_file = "\template1.txt"
                End If
                If FileExists(App.Path & temp_file) Then
                    Set ObjFile = CreateObject("Scripting.FileSystemObject")
                    Set objText = ObjFile.OpenTextFile(App.Path & temp_file)
                    text = objText.ReadAll
                    objText.Close
                
                    Set objText = Nothing
                    Set ObjFile = Nothing
                    text = Replace(text, "[AAAAAAAA]", "")   'REF (SPEC)
                    text = Replace(text, "[BBBBBBBB]", "") 'PACK
                    
                    If IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 12)) Then
                        If Val(Mid(grdsales.TextMatrix(Val(TXTSLNO.text), 12), 1, 2)) <> 0 And Val(Mid(grdsales.TextMatrix(Val(TXTSLNO.text), 12), 4, 5)) <= 12 And Val(Mid(grdsales.TextMatrix(Val(TXTSLNO.text), 12), 1, 2)) > 0 And Val(Mid(grdsales.TextMatrix(Val(TXTSLNO.text), 12), 4, 5)) > 0 Then
                            text = Replace(text, "[EEEEEEEE]", "" & Format(grdsales.TextMatrix(Val(TXTSLNO.text), 12), "dd/mm/yyyy") & "")  'EXP DATE
                        Else
                            text = Replace(text, "[EEEEEEEE]", "")   'EXP DATE
                        End If
                        text = Replace(text, "[CCCCCCCC]", "" & Format(Date, "dd/mm/yyyy") & "")  'PACK DATE
                    Else
                        text = Replace(text, "[EEEEEEEE]", "")   'EXP DATE
                        text = Replace(text, "[CCCCCCCC]", "")   'PACK DATE
                    End If
                    
                    text = Replace(text, "[DDDDDDDD]", "" & Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6)), "0.00") & "")  'MRP
                    text = Replace(text, "[FFFFFFFF]", "" & Left(Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 2)), 30) & "")  'ITEM NAME
                    text = Replace(text, "[KKKKKKKK]", "" & grdsales.TextMatrix(Val(TXTSLNO.text), 38) & "  /" & Val(TXTSLNO.text) & "-" & Val(TXTQTY.text) & "")    'BARCODE & QTY
                    text = Replace(text, "[GGGGGGGG]", "" & grdsales.TextMatrix(Val(TXTSLNO.text), 38) & "")  'BARCODE
                    'If BARFORMAT = "Y" Then
                        If Len(Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38))) Mod 2 = 0 Then
                            text = Replace(text, "[LLLLLLLL]", "" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38)) & "")  'BARCODE
                            text = Replace(text, "[MMMMMMMM]", "" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38)) & "")  'BARCODE
                        Else
                            text = Replace(text, "[LLLLLLLL]", "" & Mid(Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38)), 1, Len(Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38))) - 1) & "!100" & Right(Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38)), 1) & "") 'BARCODE
                            text = Replace(text, "[MMMMMMMM]", "" & Mid(Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38)), 1, Len(Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38))) - 1) & ">6" & Right(Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38)), 1) & "") 'BARCODE
                        End If
                    'End If
                    text = Replace(text, "[HHHHHHHH]", "" & Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18)), "0.00") & "")  'PRICE
                    text = Replace(text, "[IIIIIIII]", "" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 11)) & "")  'BATCH
                    text = Replace(text, "[JJJJJJJJ]", "" & Trim(MDIMAIN.StatusBar.Panels(5).text) & "")  'COMP NAME
                    item_no = item_no + 1
                    Dim intFile As Integer
                    Dim strFile As String
                    If FileExists(App.Path & "\BARCODE.PRN") Then
                        Kill (App.Path & "\BARCODE.PRN")
                    End If
                    strFile = App.Path & "\BARCODE.PRN" 'the file you want to save to
                    intFile = FreeFile
                    Open strFile For Output As #intFile
                        Print #intFile, text 'the data you want to save
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
    '            If MDIMAIN.barcode_profile.Caption = 0 Then
    '                If i > 0 Then Call print_3labels(i, Trim(TxtBarcode.Text), Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2)), Val(TXTRATE.Text), Val(txtretail.Text))
    '                '(i As Long, BAR_LABEL As String, itemname As String, itemmrp As Double, itemprice As Double)
    '            Else
    '                If i > 0 Then Call print_labels(i, Trim(TxtBarcode.Text), Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2)), Val(TXTRATE.Text), Val(txtretail.Text))
    '            End If
    '            Dim bar_category As String
                db.Execute "Delete from barprint"
                
    '            Set rststock = New ADODB.Recordset
    '            rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    '            With rststock
    '                If Not (.EOF And .BOF) Then
    '                    bar_category = IIf(IsNull(rststock!Category), "", rststock!Category)
    '                Else
    '                    bar_category = ""
    '                End If
    '            End With
    '            rststock.Close
    '            Set rststock = Nothing
                    
                
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "Select * From barprint", db, adOpenStatic, adLockOptimistic, adCmdText
                For M = 1 To i
                    RSTTRXFILE.AddNew
                    RSTTRXFILE!BARCODE = "*" & grdsales.TextMatrix(Val(TXTSLNO.text), 38) & "*"
                    RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 2))
                    RSTTRXFILE!item_Price = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18))
                    RSTTRXFILE!item_MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 6))
                    If IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 12)) Then
                        RSTTRXFILE!expdate = Format(grdsales.TextMatrix(Val(TXTSLNO.text), 12), "dd/mm/yyyy")
                        If IsDate(TXTINVDATE.text) Then
                            RSTTRXFILE!pckdate = Format(TXTINVDATE.text, "dd/mm/yyyy")
                        End If
                    End If
                    RSTTRXFILE!item_color = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 11))
                    RSTTRXFILE!REMARKS = ""
                    
                    
                    RSTTRXFILE!COMP_NAME = Trim(MDIMAIN.StatusBar.Panels(5).text)
                    RSTTRXFILE.Update
                Next M
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
                If BARPRINTER = barcodeprinter Then
                    ReportNameVar = Rptpath & "Rptbarprn"
                Else
                    ReportNameVar = Rptpath & "Rptbarprn1"
                End If
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
                
                Set Printer = Printers(BARPRINTER)
                Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
                Report.DiscardSavedData
                Report.VerifyOnEveryPrint = True
                Report.PrintOut (False)
                Set CRXFormulaFields = Nothing
                Set crxApplication = Nothing
                Set Report = Nothing
            End If
        Else
SKIP_BARCODE:
            If BARCODE_FLAG = False Then grdsales.TextMatrix(Val(TXTSLNO.text), 41) = Val(TXTQTY.text) + Val(TxTfree.text) 'Val(InputBox("Enter number of lables to be print", "No. of labels..", Val(TXTQTY.Text) + Val(TxtFree.Text)))
        End If
        '=======
    End If
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT DISTINCT CATEGORY FROM CATEGORY where CATEGORY = '" & Cmbcategory.text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (rststock.EOF And rststock.BOF) Then
        rststock.AddNew
        rststock!Category = Cmbcategory.text
        rststock.Update
    End If
    rststock.Close
    Set rststock = Nothing
                
    BARCODE_FLAG = False
    TXTSLNO.text = grdsales.rows
    TXTPRODUCT.text = ""
    
    Call fillcategory
    TXTITEMCODE.text = ""
    TxtBarcode.text = ""
    TXTPTR.text = ""
    TxtNetrate.text = ""
    Txtgrossamt.text = ""
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.text = ""
    'TxttaxMRP.text = ""
    TxtExDuty.text = ""
    TxtCSTper.text = ""
    TxtTrDisc.text = ""
    TxtCustDisc.text = ""
    TxtPoints.text = ""
    TxtCessPer.text = ""
    txtCess.text = ""
    txtPD.text = ""
    TxtExpense.text = ""
    txtprofit.text = ""
    TXTRETAIL.text = ""
    TxtRetailPercent.text = ""
    txtWsalePercent.text = ""
    txtSchPercent.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    Txtgrossamt.text = ""
    txtcrtn.text = ""
    TxtLWRate.text = ""
    txtcrtnpack.text = ""
    TXTRATE.text = ""
    TxtComAmt.text = ""
    TxtComper.text = ""
    txtmrpbt.text = ""
    txtBatch.text = ""
    'txtHSN.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
'    lblcategory.Caption = ""
'    Cmbcategory.text = ""
    LBLPRE.Caption = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CmdExit.Enabled = False
    'optnet.value = True
    'OptComper.value = True
    M_ADD = True
    Chkcancel.Value = 0
    OLD_BILL = True
    'txtcategory.Enabled = True
    txtBillNo.Enabled = False
    FRMEGRDTMP.Visible = False
    cmdRefresh.Enabled = True
    CMDTRANSFER.Enabled = True
    Los_Pack.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    Cmbcategory.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxtPoints.Enabled = False
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
    TxtCustDisc.Enabled = False
    TxtComper.Enabled = False
    TxtComAmt.Enabled = False
    OptComper.Enabled = False
    OptComAmt.Enabled = False
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
        If grdsales.rows >= 11 Then grdsales.TopRow = grdsales.rows - 1
        TxtBarcode.SetFocus
    End If
    M_EDIT = False
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = -2147206461 Then
        MsgBox err.Description
    ElseIf err.Number <> -2147168237 Then
        MsgBox err.Description
        On Error Resume Next
        db.RollbackTrans
    Else
        On Error Resume Next
        db.RollbackTrans
    End If
    
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
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    On Error GoTo ERRHAND
    db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.text), 16)) & ""
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    With rststock
        If Not (.EOF And .BOF) Then
            .Properties("Update Criteria").Value = adCriteriaKey
            !RCPT_QTY = !RCPT_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
            If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
            !RCPT_VAL = !RCPT_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 13))
            
            !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.text), 13))
            rststock.Update
        End If
    End With
    db.CommitTrans
    rststock.Close
    Set rststock = Nothing
    
    i = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(LINE_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTRTRXFILE.EOF
        RSTRTRXFILE!LINE_NO = i
        i = i + 1
        RSTRTRXFILE.Update
        RSTRTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    i = 1
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTRTRXFILE.EOF
        RSTRTRXFILE!LINE_NO = i
        i = i + 1
        RSTRTRXFILE.Update
        RSTRTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    grdsales.rows = 1
    i = 0
    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    LBLTOTALTAX.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    grdsales.rows = 1
    Dim GROSSVAL As Double
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
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
        grdsales.TextMatrix(i, 15) = IIf(IsNull(RSTRTRXFILE!check_flag), "N", RSTRTRXFILE!check_flag)
        grdsales.TextMatrix(i, 16) = RSTRTRXFILE!LINE_NO
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
        GROSSVAL = (Val(grdsales.TextMatrix(i, 9)) * IIf(Val(grdsales.TextMatrix(i, 5)) = 0, 1, Val(grdsales.TextMatrix(i, 5)))) * (Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 14)))
        If RSTRTRXFILE!DISC_FLAG = "P" Then
            grdsales.TextMatrix(i, 27) = "P"
            LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - (GROSSVAL * Val(grdsales.TextMatrix(i, 17)) / 100)) * Val(grdsales.TextMatrix(i, 10)) / 100, 2))
        Else
            grdsales.TextMatrix(i, 27) = "A"
            LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - Val(grdsales.TextMatrix(i, 17))) * Val(grdsales.TextMatrix(i, 10)) / 100, 2))
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
        grdsales.TextMatrix(i, 38) = IIf(IsNull(RSTRTRXFILE!BARCODE), "", RSTRTRXFILE!BARCODE)
        grdsales.TextMatrix(i, 39) = IIf(IsNull(RSTRTRXFILE!cess_amt), "", RSTRTRXFILE!cess_amt)
        grdsales.TextMatrix(i, 40) = IIf(IsNull(RSTRTRXFILE!CESS_PER), "", RSTRTRXFILE!CESS_PER)
        LBLGROSSAMT.Caption = Val(LBLGROSSAMT.Caption) + Val(grdsales.TextMatrix(i, 8)) * Val(RSTRTRXFILE!QTY)
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
        LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(i, 32)), ".00")
        lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(i, 3)), ".00")
        'TXTDEALER.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
        
        'TXTINVDATE.Text = Format(RSTRTRXFILE!VCH_DATE, "DD/MM/YYYY")
        'TXTREMARKS.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
        'TXTINVOICE.Text = IIf(IsNull(RSTRTRXFILE!PINV), "", RSTRTRXFILE!PINV)
        RSTRTRXFILE.MoveNext
    Loop
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    If Roundflag = True Then
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    Else
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 2), "0.00")
    End If
    LBLNET.Caption = Val(LBLGROSSAMT.Caption) + Val(LBLTOTALTAX.Caption)
    
    TXTSLNO.text = Val(grdsales.rows)
    TXTPRODUCT.text = ""
    TXTITEMCODE.text = ""
    TxtBarcode.text = ""
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.text = ""
    'TxttaxMRP.text = ""
    TxtExDuty.text = ""
    TxtCSTper.text = ""
    TxtTrDisc.text = ""
    TxtCustDisc.text = ""
    TxtPoints.text = ""
    TxtCessPer.text = ""
    txtCess.text = ""
    txtPD.text = ""
    TxtExpense.text = ""
    txtprofit.text = ""
    TXTRETAIL.text = ""
    TxtRetailPercent.text = ""
    
    txtWsalePercent.text = ""
    txtSchPercent.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    Txtgrossamt.text = ""
    txtcrtn.text = ""
    TxtLWRate.text = ""
    txtcrtnpack.text = ""
    TXTRATE.text = ""
    TxtComAmt.text = ""
    TxtComper.text = ""
    txtmrpbt.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    txtBatch.text = ""
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
'    lblcategory.Caption = ""
'    Cmbcategory.text = ""
    LBLPRE.Caption = ""
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CmdExit.Enabled = False
    M_ADD = True
    OLD_BILL = True
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub CmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
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
            TxTfree.text = ""
            'TxttaxMRP.text = ""
            TxtExDuty.text = ""
            TxtCSTper.text = ""
            TxtTrDisc.text = ""
            TxtCustDisc.text = ""
            TxtPoints.text = ""
            TxtCessPer.text = ""
            txtCess.text = ""
            txtPD.text = ""
            TxtExpense.text = ""
            txtprofit.text = ""
            TXTRETAIL.text = ""
            TxtRetailPercent.text = ""
            
            txtWsalePercent.text = ""
            txtSchPercent.text = ""
            txtWS.text = ""
            txtvanrate.text = ""
            Txtgrossamt.text = ""
            txtcrtn.text = ""
            TxtLWRate.text = ""
            txtcrtnpack.text = ""
            TXTRATE.text = ""
            TxtComAmt.text = ""
            TxtComper.text = ""
            txtmrpbt.text = ""
            TXTITEMCODE.text = ""
            TxtBarcode.text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            lbltaxamount.Caption = ""
'            lblcategory.Caption = ""
'            Cmbcategory.text = ""
            LBLPRE.Caption = ""
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

Private Sub CmdDeleteAll_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    On Error GoTo ERRHAND
    If Chkcancel.Value = 0 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE ALL", vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then Exit Sub
    
    TXTDEALER.text = ""
    DataList2.BoundText = ""
    TXTINVOICE.text = ""
    CMBDISTRICT.text = ""
    TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
    TXTRCVDATE.text = "  /  /    "
    TXTDATE.text = Format(Date, "DD/MM/YYYY")
    lbladdress.Caption = ""
    TXTREMARKS.text = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    LBLTOTAL.Caption = ""
    TXTDISCAMOUNT.text = ""
    TxtTotalexp.text = ""
    txtaddlamt.text = ""
        
    db.Execute "delete  From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " "
    db.Execute "delete FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'CR' AND INV_TRX_TYPE = 'PI'"
    db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'PY' AND INV_TRX_TYPE = 'PI' "
    'db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'PI'"
    For i = 1 To grdsales.rows - 1
        db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & ""
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        With rststock
            If Not (.EOF And .BOF) Then
                .Properties("Update Criteria").Value = adCriteriaKey
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
    grdsales.rows = 1
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    CMDTRANSFER.Enabled = False
    cmdRefresh.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.text = TXTLASTBILL.text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.text = "  /  /    "
    TXTRCVDATE.text = "  /  /    "
    TXTINVOICE.text = ""
    CMBDISTRICT.text = ""
    lbladdress.Caption = ""
    TXTREMARKS.text = ""
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TxtBarcode.text = ""
    TXTPRODUCT.text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.text = ""
    TxttaxMRP.text = ""
    TxtExDuty.text = ""
    TxtCSTper.text = ""
    TxtTrDisc.text = ""
    TxtCustDisc.text = ""
    TxtPoints.text = ""
    TxtCessPer.text = ""
    txtCess.text = ""
    txtPD.text = ""
    TxtExpense.text = ""
    txtprofit.text = ""
    TXTRETAIL.text = ""
    TxtRetailPercent.text = ""
    
    txtWsalePercent.text = ""
    txtSchPercent.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    Txtgrossamt.text = ""
    txtcrtn.text = ""
    TxtLWRate.text = ""
    txtcrtnpack.text = ""
    txtBatch.text = ""
    txtHSN.text = ""
    TXTRATE.text = ""
    txtmrpbt.text = ""
    TXTPTR.text = ""
    TxtNetrate.text = ""
    Txtgrossamt.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    lblcategory.Caption = ""
    Cmbcategory.text = ""
    LBLPRE.Caption = ""
    txtaddlamt.text = ""
    txtcramt.text = ""
    TxtInsurance.text = ""
    TxtCST.text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    TXTDISCAMOUNT.text = ""
    TxtTotalexp.text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.text = ""
    lbldealer.Caption = ""
    grdsales.rows = 1
    CmdExit.Enabled = True
    OptComper.Value = True
    txtBillNo.SetFocus
    M_ADD = False
    M_EDIT = False
    OLD_BILL = False
    NEW_BILL = True
    lbloldbills.Caption = "N"
    LBLmonth.Caption = "0.00"
    Chkcancel.Value = 0
    Call CLEAR_COMBO
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "EzBiz"
    Exit Sub

ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub CmdExit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdLabels_Click()
    
    If grdsales.rows <= 1 Then Exit Sub
    Dim i As Long
    Dim n, M As Long
    Dim sl As Single
    Dim sl_end As Single
    Dim temp_file As String
    Dim ObjFile, objText, text
    
    'If grdsales.Cols = 20 Then Exit Sub
    
    On Error GoTo ERRHAND
    
    If BARTEMPLATE = "Y" Then
        sl = Val(InputBox("Enter the Serial No. from which to be Print", "Label Printing", 1))
        sl_end = Val(InputBox("Enter the Serial No. upto which to be Print", "Label Printing", grdsales.rows - 1))
        If sl = 0 Then Exit Sub
        If sl_end = 0 Then sl_end = grdsales.rows - 1
        If sl_end > grdsales.rows - 1 Then Exit Sub
        If sl < sl_end Then Exit Sub
        
        If Val(MDIMAIN.LBLLABELNOS.Caption) = 0 Then MDIMAIN.LBLLABELNOS.Caption = 1
        sl = sl / Val(MDIMAIN.LBLLABELNOS.Caption)
        If sl / 10 <> 0 Then sl = Int(sl) + 1
        If Chktag.Value = 0 Then
            temp_file = "\template.txt"
        Else
            temp_file = "\template1.txt"
        End If
        If FileExists(App.Path & temp_file) Then
            For n = sl To sl_end
                Set ObjFile = CreateObject("Scripting.FileSystemObject")
                Set objText = ObjFile.OpenTextFile(App.Path & temp_file)
                text = objText.ReadAll
                objText.Close
            
                Set objText = Nothing
                Set ObjFile = Nothing
                text = Replace(text, "[AAAAAAAA]", "")   'REF (SPEC)
                text = Replace(text, "[BBBBBBBB]", "") 'PACK
                
                If IsDate(grdsales.TextMatrix(n, 12)) Then
                    If Val(Mid(grdsales.TextMatrix(n, 12), 1, 2)) <> 0 And Val(Mid(grdsales.TextMatrix(n, 12), 4, 5)) <= 12 And Val(Mid(grdsales.TextMatrix(n, 12), 1, 2)) > 0 And Val(Mid(grdsales.TextMatrix(n, 12), 4, 5)) > 0 Then
                        text = Replace(text, "[EEEEEEEE]", "" & Format(grdsales.TextMatrix(n, 12), "dd/mm/yyyy") & "")  'EXP DATE
                    Else
                        text = Replace(text, "[EEEEEEEE]", "")   'EXP DATE
                    End If
                    text = Replace(text, "[CCCCCCCC]", "" & Format(Date, "dd/mm/yyyy") & "")  'PACK DATE
                Else
                    text = Replace(text, "[EEEEEEEE]", "")   'EXP DATE
                    text = Replace(text, "[CCCCCCCC]", "")   'PACK DATE
                End If
                
                text = Replace(text, "[DDDDDDDD]", "" & Format(Val(grdsales.TextMatrix(n, 6)), "0.00") & "")  'MRP
                text = Replace(text, "[FFFFFFFF]", "" & Left(Trim(grdsales.TextMatrix(n, 2)), 30) & "") 'ITEM NAME
                text = Replace(text, "[KKKKKKKK]", "" & grdsales.TextMatrix(n, 38) & "  /" & n & "-" & Val(grdsales.TextMatrix(n, 3)) & "")    'BARCODE & QTY
                text = Replace(text, "[GGGGGGGG]", "" & grdsales.TextMatrix(n, 38) & "")  'BARCODE
                'If BARFORMAT = "Y" Then
                    If Len(Trim(grdsales.TextMatrix(n, 38))) Mod 2 = 0 Then
                        text = Replace(text, "[LLLLLLLL]", "" & Trim(grdsales.TextMatrix(n, 38)) & "")  'BARCODE
                        text = Replace(text, "[MMMMMMMM]", "" & Trim(grdsales.TextMatrix(n, 38)) & "")  'BARCODE
                    Else
                        text = Replace(text, "[LLLLLLLL]", "" & Mid(Trim(grdsales.TextMatrix(n, 38)), 1, Len(Trim(grdsales.TextMatrix(n, 38))) - 1) & "!100" & Right(Trim(grdsales.TextMatrix(n, 38)), 1) & "") 'BARCODE
                        text = Replace(text, "[MMMMMMMM]", "" & Mid(Trim(grdsales.TextMatrix(n, 38)), 1, Len(Trim(grdsales.TextMatrix(n, 38))) - 1) & ">6" & Right(Trim(grdsales.TextMatrix(n, 38)), 1) & "") 'BARCODE
                    End If
                'End If
                text = Replace(text, "[HHHHHHHH]", "" & Format(Val(grdsales.TextMatrix(n, 18)), "0.00") & "")  'PRICE
                text = Replace(text, "[IIIIIIII]", "" & Trim(grdsales.TextMatrix(n, 11)) & "")  'BATCH
                text = Replace(text, "[JJJJJJJJ]", "" & Trim(MDIMAIN.StatusBar.Panels(5).text) & "")  'COMP NAME
                
                Dim intFile As Integer
                Dim strFile As String
                If FileExists(App.Path & "\BARCODE.PRN") Then
                    Kill (App.Path & "\BARCODE.PRN")
                End If
                strFile = App.Path & "\BARCODE.PRN" 'the file you want to save to
                intFile = FreeFile
                Open strFile For Output As #intFile
                    Print #intFile, text 'the data you want to save
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
                For M = 1 To Val(grdsales.TextMatrix(n, 41))
                    Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & Rptpath & "REPO.BAT N", vbHide
                Next M
            Next n
        Else
            MsgBox "No template exists", , "EzBiz"
            Exit Sub
        End If
    Else
        db.Execute "Delete from barprint"
        Dim RSTTRXFILE As ADODB.Recordset
    '    Dim rststock As ADODB.Recordset
    '    Dim bar_category As String
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From barprint", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        
        sl = Val(InputBox("Enter the Serial No. from which to be Print", "Label Printing", 1))
        sl_end = Val(InputBox("Enter the Serial No. upto which to be Print", "Label Printing", grdsales.rows - 1))
        If sl = 0 Then Exit Sub
        If sl_end = 0 Then sl_end = grdsales.rows - 1
        If sl_end > grdsales.rows - 1 Then Exit Sub
        If sl < sl_end Then Exit Sub
        
        For n = sl To sl_end
    '        Set rststock = New ADODB.Recordset
    '        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(n, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    '        With rststock
    '            If Not (.EOF And .BOF) Then
    '                bar_category = IIf(IsNull(rststock!Category), "", rststock!Category)
    '            Else
    '                bar_category = ""
    '            End If
    '        End With
    '        rststock.Close
    '        Set rststock = Nothing
                
            For M = 1 To Val(grdsales.TextMatrix(n, 41))
                RSTTRXFILE.AddNew
                RSTTRXFILE!BARCODE = "*" & grdsales.TextMatrix(n, 38) & "*"
                RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(n, 2))
                RSTTRXFILE!item_Price = Val(grdsales.TextMatrix(n, 18))
                RSTTRXFILE!item_MRP = Val(grdsales.TextMatrix(n, 6))
                RSTTRXFILE!item_color = Trim(grdsales.TextMatrix(n, 11))
                If IsDate(grdsales.TextMatrix(n, 12)) Then
                    RSTTRXFILE!expdate = Format(grdsales.TextMatrix(n, 12), "dd/mm/yyyy")
                    If IsDate(TXTINVDATE.text) Then
                        RSTTRXFILE!pckdate = Format(TXTINVDATE.text, "dd/mm/yyyy")
                    End If
                End If
                RSTTRXFILE!REMARKS = ""
                RSTTRXFILE!COMP_NAME = Trim(MDIMAIN.StatusBar.Panels(5).text)
                RSTTRXFILE.Update
            Next M
        Next n
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        db.CommitTrans
        
        If BARPRINTER <> barcodeprinter Or Chktag.Value = 1 Then
            ReportNameVar = Rptpath & "Rptbarprn1"
        Else
            ReportNameVar = Rptpath & "Rptbarprn"
        End If
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
        Set Printer = Printers(BARPRINTER)
        Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
        Report.DiscardSavedData
        Report.VerifyOnEveryPrint = True
        Report.PrintOut (False)
        Set CRXFormulaFields = Nothing
        Set crxApplication = Nothing
        Set Report = Nothing
                
'        For i = 1 To Report.Database.Tables.COUNT
'            Report.Database.Tables.Item(i).SetLogOnInfo strConnection
'        Next i
'        Report.DiscardSavedData
'        frmreport.Caption = "BARCODE"
'        Call GENERATEREPORT
    End If


    Screen.MousePointer = vbNormal
        
Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = -2147206461 Then
        MsgBox err.Description
    ElseIf err.Number <> -2147168237 Then
        MsgBox err.Description
        On Error Resume Next
        db.RollbackTrans
    Else
        On Error Resume Next
        db.RollbackTrans
    End If
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.text) >= grdsales.rows Then Exit Sub
    
    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CmdExit.Enabled = False
'    Los_Pack.Enabled = True
'    CmbPack.Enabled = True
'    TXTQTY.Enabled = True
'    TxtFree.Enabled = True
'    TXTRATE.Enabled = True
'    TXTPTR.Enabled = True
'    TxttaxMRP.Enabled = True
'    TxtExDuty.Enabled = True
'    TxtTrDisc.Enabled = True
'    TxtCessPer.Enabled = True
'    txtCess.Enabled = True
'    TxtCSTper.Enabled = True
'    txtPD.Enabled = True
'    TxtExpense.Enabled = True
'    TXTRETAIL.Enabled = True
'    TxtRetailPercent.Enabled = True
'    txtWS.Enabled = True
'    txtWsalePercent.Enabled = True
'    txtvanrate.Enabled = True
'    txtSchPercent.Enabled = True
'    txtcrtnpack.Enabled = True
'    txtcrtn.Enabled = True
'    TxtLWRate.Enabled = True
'    TxtComper.Enabled = True
'    TxtComAmt.Enabled = True
'    cmdadd.Enabled = True
'    Txtgrossamt.Enabled = True
'    txtBatch.Enabled = True
'    txtHSN.Enabled = True
'    TxtWarranty.Enabled = True
'    CmbWrnty.Enabled = True
'    TXTEXPIRY.Visible = False
'    TXTEXPDATE.Enabled = True
'    TxtBarcode.Enabled = False
    If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
        Los_Pack.text = 1
        TXTQTY.text = 1
        TxTfree.text = ""
        TXTRATE.text = ""
        TXTPTR.Enabled = True
        TXTPTR.SetFocus
    Else
        Los_Pack.Enabled = True
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
    End If
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
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
            TxTfree.text = ""
            'TxttaxMRP.text = ""
            TxtExDuty.text = ""
            TxtCSTper.text = ""
            TxtTrDisc.text = ""
            TxtCustDisc.text = ""
            TxtPoints.text = ""
            TxtCessPer.text = ""
            txtCess.text = ""
            txtPD.text = ""
            TxtExpense.text = ""
            txtprofit.text = ""
            TXTRETAIL.text = ""
            TxtRetailPercent.text = ""
            
            txtWsalePercent.text = ""
            txtSchPercent.text = ""
            txtWS.text = ""
            txtvanrate.text = ""
            Txtgrossamt.text = ""
            txtcrtn.text = ""
            TxtLWRate.text = ""
            txtcrtnpack.text = ""
            TXTRATE.text = ""
            TxtComAmt.text = ""
            TxtComper.text = ""
            txtmrpbt.text = ""
            TXTITEMCODE.text = ""
            TxtBarcode.text = ""
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            lbltaxamount.Caption = ""
            lblcategory.Caption = ""
            Cmbcategory.text = ""
            LBLPRE.Caption = ""
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

Private Sub CmdPrint_Click()
    Dim i As Long
    
    On Error GoTo ERRHAND
     
    Screen.MousePointer = vbHourglass
    db.Execute "delete from TEMPTRXFILE"
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "LP"
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
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(i, 38))
        
        'RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!SCHEME = Val(grdsales.TextMatrix(i, 14))
        If IsDate(grdsales.TextMatrix(i, 12)) Then
            RSTTRXFILE!EXP_DATE = Format(grdsales.TextMatrix(i, 12), "dd/mm/yyyy")
        End If
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(i, 15))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        
        RSTTRXFILE.Update
    Next i
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
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
            
    ReportNameVar = Rptpath & "rptLP"
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
            If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.text = "'" & TXTINVOICE.text & "'"
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
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub cmdRefresh_Click()
    If ((frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4" And frmLogin.rs!Level <> "1") And NEW_BILL = False) Or (frmLogin.rs!Level <> "0" And lbloldbills.Caption = "Y") Then
        MsgBox "PERMISSION DENIED", vbOKOnly, "EzBiz"
        CmdExit.Enabled = True
        Exit Sub
    End If
    If CMBPO.VisibleCount = 0 Then CMBPO.text = ""
    If CMBPO.VisibleCount <> 0 And CMBPO.BoundText = "" Then
        If (MsgBox("Are you sure you want to save the Purchase Bill without selecting the Purchase Order No.? !!!!", vbYesNo, "EzBiz") = vbNo) Then Exit Sub
    End If
    If CMBPO.text <> "" And CMBPO.MatchedWithList = False Then
        MsgBox "Please select a valid PO No. from the list", vbOKOnly, "EzBiz"
        On Error Resume Next
        CMBPO.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(TXTRCVDATE.text) Then
        TXTRCVDATE.text = TXTINVDATE.text
    End If
    If (DateValue(TXTINVDATE.text) < DateValue(MDIMAIN.DTFROM.Value)) Or (DateValue(TXTINVDATE.text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.Value))) Then
        'db.Execute "delete from Users"
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
'    If (DateValue(TXTRCVDATE.Text) < DateValue(MDIMAIN.DTFROM.value)) Or (DateValue(TXTRCVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.value))) Then
'        'db.Execute "delete from Users"
'        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
'        TXTINVDATE.SetFocus
'        Exit Sub
'    End If
    
    If DateValue(TXTRCVDATE.text) < DateValue(TXTINVDATE.text) Then
        'db.Execute "delete from Users"
        MsgBox "Goods received date could not be less than Invice date", vbOKOnly, "EzBiz"
        TXTRCVDATE.SetFocus
        Exit Sub
    End If
    
    If DateValue(TXTRCVDATE.text) <> DateValue(TXTINVDATE.text) Then
        If (MsgBox("Invice date & Rcvd date are different. Are you sure?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo) Then
            TXTRCVDATE.SetFocus
            Exit Sub
        End If
    End If
    
    BARCODE_FLAG = False
    On Error GoTo ERRHAND
    If grdsales.rows <= 1 Then
        lblcredit.Caption = "0"
        Call appendpurchase
    Else
        If IsNull(DataList2.SelectedItem) Then
            MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
            DataList2.SetFocus
            Exit Sub
        End If
        If TXTINVOICE.text = "" Then
            MsgBox "Enter Supplier Invoice No.", vbOKOnly, "EzBiz"
            Exit Sub
        End If
        If Not IsDate(TXTINVDATE.text) Then
            MsgBox "Enter Supplier Invoice Date", vbOKOnly, "EzBiz"
            Exit Sub
        End If
        'Me.Enabled = False
        'MDIMAIN.cmdpurchase.Enabled = False
        'Set creditbill = Me
        'frmCREDIT.Show
        Call appendpurchase
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdRefresh_GotFocus()
    FRMEGRDTMP.Visible = False
End Sub

Private Sub CmdTransfer_Click()
    
    If grdsales.rows <= 1 Then Exit Sub
    Chkcancel.Value = 0
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Supplier From List", vbOKOnly, "Export Bill"
        FRMEMASTER.Enabled = True
        DataList2.SetFocus
        Exit Sub
    End If
    If TXTINVOICE.text = "" Then
        FRMEMASTER.Enabled = True
        MsgBox "Enter Supplier Invoice No.", vbOKOnly, "Export Bill"
        Exit Sub
    End If
    If Not IsDate(TXTINVDATE.text) Then
        FRMEMASTER.Enabled = True
        MsgBox "Enter Supplier Invoice Date", vbOKOnly, "Export Bill"
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to export the bill?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
    
    On Error GoTo ERRHAND
    Dim Strconnct As String
    
    Dim db2 As New ADODB.Connection
    Strconnct = "Driver={MySQL ODBC 5.1 Driver};Server=" & DBPath & ";Port=3306;Database=" & dbase2 & ";User=root; Password=###%%database%%###ret; Option=2;"
    db2.Open Strconnct
    db2.CursorLocation = adUseClient
    
    Dim RSTITEMMAST, rstTRXMAST As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    Dim rstBILL As ADODB.Recordset
    
    Set rstTRXMAST = New ADODB.Recordset
    rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND PINV = '" & Trim(TXTINVOICE.text) & "' AND ACT_NAME = '" & DataList2.text & "'", db2, adOpenStatic, adLockReadOnly
    If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
        MsgBox "You have already exported this Invoice of " & Trim(DataList2.text) & " System Ref: No. " & rstTRXMAST!VCH_NO, vbOKOnly, "Export Data"
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
        db2.Close
        Set db2 = Nothing
        Exit Sub
    End If
    rstTRXMAST.Close
    Set rstTRXMAST = Nothing
    
    Screen.MousePointer = vbHourglass
    
    Dim RSTRTRXFILE, rststock As ADODB.Recordset
    Dim M_DATA As Double
    Dim PR_CODE, PR_NAME As String
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PI'", db2, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Tag = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    For i = 1 To grdsales.rows - 1
        Set rstTRXMAST = New ADODB.Recordset
        rstTRXMAST.Open "Select * From ITEMMAST WHERE ITEM_NAME = '" & Trim(grdsales.TextMatrix(i, 2)) & "' ", db2, adOpenStatic, adLockReadOnly, adCmdText
        'rstTRXMAST.Open "Select * From ITEMMAST WHERE ITEM_CODE = '" & Trim(grdsales.TextMatrix(i, 1)) & "' ", db2, adOpenStatic, adLockReadOnly, adCmdText
        If (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db2, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                If IsNull(RSTITEMMAST.Fields(0)) Then
                    PR_CODE = 1
                Else
                    PR_CODE = Val(RSTITEMMAST.Fields(0)) + 1
                End If
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ITEMMAST", db2, adOpenStatic, adLockOptimistic, adCmdText
            RSTITEMMAST.AddNew
            RSTITEMMAST!ITEM_CODE = Val(PR_CODE)
            RSTITEMMAST!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
            PR_NAME = Trim(grdsales.TextMatrix(i, 2))
            If Trim(Cmbcategory.text) = "" Then
                RSTITEMMAST!Category = "GENERAL"
            Else
                RSTITEMMAST!Category = Trim(Cmbcategory.text)
            End If
            RSTITEMMAST!UNIT = 1
            RSTITEMMAST!MANUFACTURER = "GENERAL"
            RSTITEMMAST!DEAD_STOCK = "N"
            RSTITEMMAST!REMARKS = ""
            RSTITEMMAST!REORDER_QTY = 1
            RSTITEMMAST!PACK_TYPE = "Nos"
            RSTITEMMAST!FULL_PACK = "Nos"
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
            RSTITEMMAST!P_RETAIL = 0
            RSTITEMMAST!P_WS = 0
            RSTITEMMAST!CRTN_PACK = 1
            RSTITEMMAST!P_CRTN = 0
            RSTITEMMAST!LOOSE_PACK = 1
            RSTITEMMAST!UN_BILL = "N"
            If PC_FLAG = "Y" Then
                RSTITEMMAST!PRICE_CHANGE = "Y"
            Else
                RSTITEMMAST!PRICE_CHANGE = "N"
            End If
            RSTITEMMAST.Update
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        Else
            PR_CODE = rstTRXMAST!ITEM_CODE
            PR_NAME = rstTRXMAST!ITEM_NAME
        End If
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
    
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Tag) & " AND ITEM_CODE='" & PR_CODE & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "", db2, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
            RSTRTRXFILE.AddNew
            RSTRTRXFILE!TRX_TYPE = "PI"
            RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTRTRXFILE!VCH_NO = Val(txtBillNo.Tag)
            RSTRTRXFILE!LINE_NO = Val(grdsales.TextMatrix(i, 0))
            RSTRTRXFILE!ITEM_CODE = PR_CODE
            RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
            RSTRTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
    
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & PR_CODE & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
            With rststock
                If Not (.EOF And .BOF) Then
                    .Properties("Update Criteria").Value = adCriteriaKey
    '                If UCase(rststock!CATEGORY) = "CUTSHEET" Then
    '                Else
                    !ITEM_COST = Val(grdsales.TextMatrix(i, 8))
                    !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    '!CLOSE_VAL = !CLOSE_VAL + (Val(grdsales.TextMatrix(i, 13)) / Val(Los_Pack.Text))
                    !CLOSE_VAL = Round(!ITEM_COST * !CLOSE_QTY, 3)
                    !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                    If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                    '!RCPT_VAL = !RCPT_VAL + (Val(grdsales.TextMatrix(i, 13)) / Val(Los_Pack.Text))
                    !RCPT_VAL = Round(!ITEM_COST * !RCPT_QTY, 3)
                
                    !MRP = Val(grdsales.TextMatrix(i, 6))
                    'If Trim(txtHSN.Text) <> "" Then !Remarks = Trim(txtHSN.Text)
                    'If cmbfull.ListIndex <> -1 Then !FULL_PACK = cmbfull.Text
                    '!CUST_DISC = Val(TxtCustDisc.Text)
                    If !PRICE_CHANGE = "Y" And Val(grdsales.TextMatrix(i, 18)) <> 0 Then
                        db2.Execute "Update RTRXFILE set P_RETAIL = " & Val(grdsales.TextMatrix(i, 18)) & " WHERE ITEM_CODE = '" & PR_CODE & "' AND BAL_QTY >0 "
                    End If
                    If Trim(grdsales.TextMatrix(i, 38)) <> "" Then !BARCODE = Trim(grdsales.TextMatrix(i, 38))
                    If Val(grdsales.TextMatrix(i, 18)) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(i, 18))
                    If Val(grdsales.TextMatrix(i, 19)) <> 0 Then !P_WS = Val(grdsales.TextMatrix(i, 19))
                    If Val(grdsales.TextMatrix(i, 20)) <> 0 Then !P_CRTN = Val(grdsales.TextMatrix(i, 20)) ' / Val(Los_Pack.Text), 3)
                    If Val(grdsales.TextMatrix(i, 37)) <> 0 Then !P_LWS = Val(grdsales.TextMatrix(i, 37)) ' / Val(Los_Pack.Text), 3)
                    If Val(grdsales.TextMatrix(i, 25)) <> 0 Then !P_VAN = Val(grdsales.TextMatrix(i, 25)) ' / Val(Los_Pack.Text), 3)
                    If Val(grdsales.TextMatrix(i, 39)) <> 0 Then !cess_amt = Val(grdsales.TextMatrix(i, 39)) ' / Val(Los_Pack.Text), 3)
                    If Val(grdsales.TextMatrix(i, 40)) <> 0 Then !CESS_PER = Val(grdsales.TextMatrix(i, 40)) ' / Val(Los_Pack.Text), 3)
                    
                    '!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
                    If Val(Val(grdsales.TextMatrix(i, 24))) <> 0 Then !CRTN_PACK = Val(grdsales.TextMatrix(i, 24))
    
                    If Trim(grdsales.TextMatrix(i, 23)) = "A" Then
                        !COM_FLAG = "A"
                        !COM_PER = 0
                        !COM_AMT = Val(grdsales.TextMatrix(i, 22))
                    Else
                        !COM_FLAG = "P"
                        !COM_PER = Val(grdsales.TextMatrix(i, 21))
                        !COM_AMT = 0
                    End If
                    If Val(Val(grdsales.TextMatrix(i, 10))) >= 5 Then !SALES_TAX = Val(grdsales.TextMatrix(i, 10))
                    '!SALES_TAX = Val(grdsales.TextMatrix(i, 10))
                    !check_flag = Trim(grdsales.TextMatrix(i, 15))
                    !LOOSE_PACK = Val(grdsales.TextMatrix(i, 28))
                    !PACK_TYPE = Val(grdsales.TextMatrix(i, 29))
                    !WARRANTY = Val(grdsales.TextMatrix(i, 30))
                    !WARRANTY_TYPE = grdsales.TextMatrix(i, 31)
                    RSTRTRXFILE!MFGR = !MANUFACTURER
                    If Trim(Cmbcategory.text) = "" Then
                        RSTRTRXFILE!Category = "GENERAL"
                    Else
                        RSTRTRXFILE!Category = Trim(Cmbcategory.text)
                    End If
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
            rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & PR_CODE & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
            With rststock
                If Not (.EOF And .BOF) Then
                    .Properties("Update Criteria").Value = adCriteriaKey
                    '!ITEM_COST = Val(grdsales.TextMatrix(i, 8))
                    !ITEM_COST = Val(grdsales.TextMatrix(i, 8))
                    !CLOSE_QTY = !CLOSE_QTY - RSTRTRXFILE!QTY
                    !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    '!CLOSE_VAL = !CLOSE_VAL + (Val(grdsales.TextMatrix(i, 13)) / Val(Los_Pack.Text))
                    !CLOSE_VAL = Round(!ITEM_COST * !CLOSE_QTY, 3)
                    
                    !RCPT_QTY = !RCPT_QTY - RSTRTRXFILE!QTY
                    !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
                    If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                    '!RCPT_VAL =  !RCPT_VAL + (Val(grdsales.TextMatrix(i, 13)) / Val(Los_Pack.Text))
                    !RCPT_VAL = Round(!ITEM_COST * !RCPT_QTY, 3)
                    
                    'If Trim(txtHSN.Text) <> "" Then !Remarks = Trim(txtHSN.Text)
                    'If cmbfull.ListIndex <> -1 Then !FULL_PACK = cmbfull.Text
                    '!CUST_DISC = Val(TxtCustDisc.Text)
                
                    If !PRICE_CHANGE = "Y" And Val(grdsales.TextMatrix(i, 18)) <> 0 Then
                        db2.Execute "Update RTRXFILE set P_RETAIL = " & Val(grdsales.TextMatrix(i, 18)) & " WHERE ITEM_CODE = '" & PR_CODE & "' AND BAL_QTY >0 "
                    End If
                    
                    !MRP = Val(grdsales.TextMatrix(i, 6))
                    If Trim(grdsales.TextMatrix(i, 38)) <> "" Then !BARCODE = Trim(grdsales.TextMatrix(i, 38))
                    If Val(grdsales.TextMatrix(i, 18)) <> 0 Then !P_RETAIL = Val(grdsales.TextMatrix(i, 18))
                    If Val(grdsales.TextMatrix(i, 19)) <> 0 Then !P_WS = Val(grdsales.TextMatrix(i, 19))
                    If Val(grdsales.TextMatrix(i, 20)) <> 0 Then !P_CRTN = Val(grdsales.TextMatrix(i, 20)) ' / Val(Los_Pack.Text), 3)
                    If Val(grdsales.TextMatrix(i, 37)) <> 0 Then !P_LWS = Val(grdsales.TextMatrix(i, 37)) ' / Val(Los_Pack.Text), 3)
                    If Val(grdsales.TextMatrix(i, 25)) <> 0 Then !P_VAN = Val(grdsales.TextMatrix(i, 25)) ' / Val(Los_Pack.Text), 3)
                    If Val(grdsales.TextMatrix(i, 39)) <> 0 Then !cess_amt = Val(grdsales.TextMatrix(i, 39)) ' / Val(Los_Pack.Text), 3)
                    If Val(grdsales.TextMatrix(i, 40)) <> 0 Then !CESS_PER = Val(grdsales.TextMatrix(i, 40)) ' / Val(Los_Pack.Text), 3)
    
                    '!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
                    If Val(Val(grdsales.TextMatrix(i, 24))) <> 0 Then !CRTN_PACK = Val(grdsales.TextMatrix(i, 24))
                                        
                    If Trim(grdsales.TextMatrix(i, 23)) = "A" Then
                        !COM_FLAG = "A"
                        !COM_PER = 0
                        !COM_AMT = Val(grdsales.TextMatrix(i, 22))
                    Else
                        !COM_FLAG = "P"
                        !COM_PER = Val(grdsales.TextMatrix(i, 21))
                        !COM_AMT = 0
                    End If
                    If Val(Val(grdsales.TextMatrix(i, 10))) >= 5 Then !SALES_TAX = Val(grdsales.TextMatrix(i, 10))
                    '!SALES_TAX = Val(grdsales.TextMatrix(i, 10))
                    !check_flag = Trim(grdsales.TextMatrix(i, 15))
                    !LOOSE_PACK = Val(grdsales.TextMatrix(i, 28))
                    !PACK_TYPE = Val(grdsales.TextMatrix(i, 29))
                    !WARRANTY = Val(grdsales.TextMatrix(i, 30))
                    !WARRANTY_TYPE = grdsales.TextMatrix(i, 31)
                    RSTRTRXFILE!MFGR = !MANUFACTURER
                    If Trim(Cmbcategory.text) = "" Then
                        RSTRTRXFILE!Category = "GENERAL"
                    Else
                        RSTRTRXFILE!Category = Trim(Cmbcategory.text)
                    End If
                    rststock.Update
                End If
            End With
            rststock.Close
            Set rststock = Nothing
            RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 5))
        End If
        RSTRTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(i, 13))
        RSTRTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "dd/mm/yyyy")
        If IsDate(TXTRCVDATE.text) Then
            RSTRTRXFILE!RCVD_DATE = Format(TXTRCVDATE.text, "DD/MM/YYYY")
        Else
            RSTRTRXFILE!RCVD_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        End If
        RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
        RSTRTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(i, 8))
        RSTRTRXFILE!ITEM_COST_PRICE = Round(Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 5)), 4)
        'RSTRTRXFILE!ITEM_NET_COST_PRICE = Round((Val(grdsales.TextMatrix(i, 13)) / Val(grdsales.TextMatrix(i, 3))), 3)
        If (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))) = 0 Then
            RSTRTRXFILE!ITEM_NET_COST_PRICE = Round((Val(grdsales.TextMatrix(i, 13)) / Val(grdsales.TextMatrix(i, 3))) + Val(grdsales.TextMatrix(i, 32)), 3)
        Else
            RSTRTRXFILE!ITEM_NET_COST_PRICE = Round((Val(grdsales.TextMatrix(i, 13)) / Val(grdsales.TextMatrix(i, 3))) + (Val(grdsales.TextMatrix(i, 32)) / ((Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))) * Val(grdsales.TextMatrix(i, 5)))), 3)
        End If
    
        RSTRTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 5))
        RSTRTRXFILE!P_DISC = Val(grdsales.TextMatrix(i, 17))
        RSTRTRXFILE!MRP = Val(grdsales.TextMatrix(i, 6))
        RSTRTRXFILE!PTR = Val(grdsales.TextMatrix(i, 9))
        RSTRTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
        RSTRTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 18))
        RSTRTRXFILE!P_WS = Val(grdsales.TextMatrix(i, 19))
        RSTRTRXFILE!P_CRTN = Val(grdsales.TextMatrix(i, 20))
        RSTRTRXFILE!P_LWS = Val(grdsales.TextMatrix(i, 37))
        RSTRTRXFILE!CRTN_PACK = Val(grdsales.TextMatrix(i, 24))
        RSTRTRXFILE!P_VAN = Val(grdsales.TextMatrix(i, 25))
        RSTRTRXFILE!gross_amt = Val(grdsales.TextMatrix(i, 26))
        RSTRTRXFILE!BARCODE = Trim(grdsales.TextMatrix(i, 38))
        RSTRTRXFILE!cess_amt = Val(grdsales.TextMatrix(i, 39))
        RSTRTRXFILE!CESS_PER = Val(grdsales.TextMatrix(i, 40))
        If Trim(grdsales.TextMatrix(i, 23)) = "A" Then
            RSTRTRXFILE!COM_FLAG = "A"
            RSTRTRXFILE!COM_PER = 0
            RSTRTRXFILE!COM_AMT = Val(grdsales.TextMatrix(i, 22))
        Else
            RSTRTRXFILE!COM_FLAG = "P"
            RSTRTRXFILE!COM_PER = Val(grdsales.TextMatrix(i, 21))
            RSTRTRXFILE!COM_AMT = 0
        End If
        RSTRTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(i, 10))
        
        RSTRTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 28))
        RSTRTRXFILE!PACK_TYPE = Val(grdsales.TextMatrix(i, 29))
        RSTRTRXFILE!WARRANTY = Val(grdsales.TextMatrix(i, 30))
        RSTRTRXFILE!WARRANTY_TYPE = grdsales.TextMatrix(i, 31)
        RSTRTRXFILE!EXPENSE = 0
        RSTRTRXFILE!EXDUTY = 0 'Val(TxtExDuty.Text)
        RSTRTRXFILE!CSTPER = 0 'Val(TxtCSTper.Text)
        RSTRTRXFILE!TR_DISC = Val(grdsales.TextMatrix(i, 35))
        
        RSTRTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(i, 4))
        'RSTRTRXFILE!VCH_DESC = "Received From " & DataList2.Text
        RSTRTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 11))
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
        RSTRTRXFILE!check_flag = Trim(grdsales.TextMatrix(i, 15))
        RSTRTRXFILE.Update
        RSTRTRXFILE.Close
        
        M_DATA = 0
        Set RSTRTRXFILE = Nothing
    Next i

    db2.Execute "delete From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Tag) & ""
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Tag) & "", db2, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Tag)
        RSTTRXFILE!TRX_TYPE = "PI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        If IsDate(TXTRCVDATE.text) Then
            RSTTRXFILE!RCVD_DATE = Format(TXTRCVDATE.text, "DD/MM/YYYY")
        Else
            RSTTRXFILE!RCVD_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        End If
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = Trim(DataList2.text)
        RSTTRXFILE!VCH_AMOUNT = Val(lbltotalwodiscount.Caption)
        RSTTRXFILE!DISCOUNT = Val(TXTDISCAMOUNT.text)
        RSTTRXFILE!ADD_AMOUNT = Val(txtaddlamt.text)
        RSTTRXFILE!ROUNDED_OFF = 0
        RSTTRXFILE!OPEN_PAY = 0
        RSTTRXFILE!PAY_AMOUNT = 0
        RSTTRXFILE!REF_NO = ""
        If OptDr.Value = True Then
            RSTTRXFILE!SLSM_CODE = "DR"
        Else
            RSTTRXFILE!SLSM_CODE = "CR"
        End If
        RSTTRXFILE!check_flag = "N"
        'If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
        RSTTRXFILE!POST_FLAG = "N"
        RSTTRXFILE!CFORM_NO = ""
        RSTTRXFILE!CFORM_DATE = Date
        RSTTRXFILE!REMARKS = Trim(TXTREMARKS.text)
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
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    db2.Close
    Set db2 = Nothing

SKIP:
    Screen.MousePointer = vbNormal
    MsgBox "EXPORTED SUCCESSFULLY", vbOKOnly, "EzBiz"
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = 7 Then
        MsgBox "Select Supplier from the list", vbOKOnly, "EzBiz"
    Else
        MsgBox err.Description
    End If
End Sub


Private Sub Command4_Click()
    If CmdExit.Enabled = False Then Exit Sub
    If Val(txtBillNo.text) = 1 Then Exit Sub
    txtBillNo.text = Val(txtBillNo.text) - 1
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    CMDTRANSFER.Enabled = False
    cmdRefresh.Enabled = False
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.text = "  /  /    "
    TXTRCVDATE.text = "  /  /    "
    TXTINVOICE.text = ""
    CMBDISTRICT.text = ""
    lbladdress.Caption = ""
    TXTREMARKS.text = ""
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TxtBarcode.text = ""
    TXTPRODUCT.text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.text = ""
    'TxttaxMRP.text = ""
    TxtExDuty.text = ""
    TxtCSTper.text = ""
    TxtTrDisc.text = ""
    TxtCustDisc.text = ""
    TxtPoints.text = ""
    TxtCessPer.text = ""
    txtCess.text = ""
    txtPD.text = ""
    TxtExpense.text = ""
    txtprofit.text = ""
    TXTRETAIL.text = ""
    TxtRetailPercent.text = ""
    
    txtWsalePercent.text = ""
    txtSchPercent.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    Txtgrossamt.text = ""
    txtcrtn.text = ""
    TxtLWRate.text = ""
    txtcrtnpack.text = ""
    txtBatch.text = ""
    txtHSN.text = ""
    TXTRATE.text = ""
    txtmrpbt.text = ""
    TXTPTR.text = ""
    TxtNetrate.text = ""
    
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
'    lblcategory.Caption = ""
'    Cmbcategory.text = ""
    LBLPRE.Caption = ""
    txtaddlamt.text = ""
    txtcramt.text = ""
    TxtInsurance.text = ""
    TxtCST.text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    TXTDISCAMOUNT.text = ""
    TxtTotalexp.text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.text = ""
    lbldealer.Caption = ""
    grdsales.rows = 1
    CmdExit.Enabled = True
    OptComper.Value = True
    M_ADD = False
    OLD_BILL = False
    NEW_BILL = True
    lbloldbills.Caption = "N"
    LBLmonth.Caption = "0.00"
    
    Chkcancel.Value = 0
    Call txtBillNo_KeyDown(13, 0)
End Sub

Private Sub Command5_Click()
    If CmdExit.Enabled = False Then Exit Sub
    Dim rstBILL As ADODB.Recordset
    Dim lastbillno As Double
    On Error GoTo ERRHAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        lastbillno = IIf(IsNull(rstBILL.Fields(0)), 0, rstBILL.Fields(0))
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    If Val(txtBillNo.text) > lastbillno Then Exit Sub
    txtBillNo.text = Val(txtBillNo.text) + 1
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    CMDTRANSFER.Enabled = False
    cmdRefresh.Enabled = False
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.text = "  /  /    "
    TXTRCVDATE.text = "  /  /    "
    TXTINVOICE.text = ""
    CMBDISTRICT.text = ""
    lbladdress.Caption = ""
    TXTREMARKS.text = ""
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TxtBarcode.text = ""
    TXTPRODUCT.text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.text = ""
    'TxttaxMRP.text = ""
    TxtExDuty.text = ""
    TxtCSTper.text = ""
    TxtTrDisc.text = ""
    TxtCustDisc.text = ""
    TxtPoints.text = ""
    TxtCessPer.text = ""
    txtCess.text = ""
    txtPD.text = ""
    TxtExpense.text = ""
    txtprofit.text = ""
    TXTRETAIL.text = ""
    TxtRetailPercent.text = ""
    
    txtWsalePercent.text = ""
    txtSchPercent.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    Txtgrossamt.text = ""
    txtcrtn.text = ""
    TxtLWRate.text = ""
    txtcrtnpack.text = ""
    txtBatch.text = ""
    txtHSN.text = ""
    TXTRATE.text = ""
    txtmrpbt.text = ""
    TXTPTR.text = ""
    TxtNetrate.text = ""
    
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
'    lblcategory.Caption = ""
'    Cmbcategory.text = ""
    LBLPRE.Caption = ""
    txtaddlamt.text = ""
    txtcramt.text = ""
    TxtInsurance.text = ""
    TxtCST.text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    TXTDISCAMOUNT.text = ""
    TxtTotalexp.text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.text = ""
    lbldealer.Caption = ""
    grdsales.rows = 1
    CmdExit.Enabled = True
    OptComper.Value = True
    M_ADD = False
    OLD_BILL = False
    NEW_BILL = True
    lbloldbills.Caption = "N"
    LBLmonth.Caption = "0.00"
    
    Chkcancel.Value = 0
    Call txtBillNo_KeyDown(13, 0)
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub D_Click()
    
End Sub

Private Sub Form_Activate()
    'On Error GoTo ErrHand
    On Error Resume Next
    If txtBillNo.Enabled = True Then txtBillNo.SetFocus
    'if TXTDEALER.Enabled=True then TXTDEALER.SetFocus
'    Exit Sub
'ErrHand:
'    If Err.Number = 5 Then Exit Sub
'    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
        
    Dim p
    For Each p In Printers
        Cmbbarcode.AddItem (p.DeviceName)
    Next p
    
    Dim ObjFso
    Dim StrFileName
    Dim ObjFile
    If FileExists(App.Path & "\BillPrint") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\BillPrint")  'Reading from the file
        On Error Resume Next
        ObjFile.ReadLine
        ObjFile.ReadLine
        ObjFile.ReadLine
        Cmbbarcode.ListIndex = ObjFile.ReadLine
        err.Clear
        On Error GoTo ERRHAND
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    BARPRINTER = barcodeprinter
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    Call CLEAR_COMBO
    Call fillcategory
    
    M_EDIT = False
    NEW_BILL = True
    lbloldbills.Caption = "N"
    ACT_FLAG = True
    PO_FLAG = True
    PRERATE_FLAG = True
    OLD_BILL = False
    grdsales.ColWidth(0) = 500
    grdsales.ColWidth(1) = 1000
    grdsales.ColWidth(2) = 4000
    grdsales.ColWidth(3) = 1000
    grdsales.ColWidth(4) = 0 ' 800
    grdsales.ColWidth(5) = 0 '800
    grdsales.ColWidth(6) = 1000
    grdsales.ColWidth(7) = 0 '800
    grdsales.ColWidth(8) = 800
    grdsales.ColWidth(9) = 800
    grdsales.ColWidth(10) = 800
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 0 '1100
    grdsales.ColWidth(13) = 1700 '1100
    grdsales.ColWidth(16) = 0
    grdsales.ColWidth(14) = 800
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(17) = 800
    grdsales.ColWidth(18) = 800
    grdsales.ColWidth(19) = 800
    grdsales.ColWidth(20) = 800
    grdsales.ColWidth(37) = 800
    grdsales.ColWidth(21) = 0
    grdsales.ColWidth(22) = 0
    grdsales.ColWidth(23) = 0
    grdsales.ColWidth(24) = 700
    grdsales.ColWidth(25) = 0
    grdsales.ColWidth(26) = 1100
    grdsales.ColWidth(27) = 0
    grdsales.ColWidth(28) = 1100
    grdsales.ColWidth(29) = 0
    grdsales.ColWidth(30) = 0
    grdsales.ColWidth(31) = 0
    grdsales.ColWidth(33) = 0
    grdsales.ColWidth(34) = 0
    
    grdsales.ColAlignment(0) = 4
    grdsales.ColAlignment(1) = 1
    grdsales.ColAlignment(2) = 1
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
    grdsales.TextArray(24) = "L. Pck"
    grdsales.TextArray(25) = "Van Rate"
    grdsales.TextArray(26) = "GROSS AMT"
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
    grdsales.TextArray(41) = "Labels"
    
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
    TXTSLNO.text = 1
    TXTSLNO.Enabled = True
    FRMECONTROLS.Enabled = False
    FRMEMASTER.Enabled = False
    CLOSEALL = 1
    lblcredit.Caption = "1"
    TXTDEALER.text = ""
    M_ADD = False
    If MDIMAIN.lblExpEnable.Caption = "Y" Then CMDTRANSFER.Visible = True
    'Me.Width = 15135
    'Me.Height = 9660
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If PHYCODE_FLAG = False Then PHY_CODE.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If PO_FLAG = False Then ACT_PO.Close
        If PRERATE_FLAG = False Then PHY_PRERATE.Close
        If CAT_REC.State = 1 Then CAT_REC.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdsales_Click()
    TXTsample.Visible = False
End Sub

Private Sub grdsales_DblClick()
    If grdsales.rows <= 1 Then Exit Sub
    If M_EDIT = True Then Exit Sub
    If ((frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4" And frmLogin.rs!Level <> "1") And NEW_BILL = False) Or (frmLogin.rs!Level <> "0" And lbloldbills.Caption = "Y") Then Exit Sub
    TXTSLNO.text = grdsales.TextMatrix(grdsales.Row, 0)
    Call TXTSLNO_KeyDown(13, 0)
    CMDMODIFY_Click
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If ((frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4" And frmLogin.rs!Level <> "1") And NEW_BILL = False) Or (frmLogin.rs!Level <> "0" And lbloldbills.Caption = "Y") Then Exit Sub
            Select Case grdsales.Col
                Case 41
                        If grdsales.Cols = 20 Then Exit Sub
                        TXTsample.MaxLength = 3
                        TXTsample.Visible = True
                        TXTsample.Top = grdsales.CellTop + 100
                        TXTsample.Left = grdsales.CellLeft '+ 50
                        TXTsample.Width = grdsales.CellWidth
                        TXTsample.Height = grdsales.CellHeight
                        TXTsample.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                        TXTsample.SetFocus
            End Select
        Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            If txtBillNo.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then Exit Sub
            If Not IsDate(TXTINVDATE.text) Then Exit Sub
            If TXTQTY.Enabled = True Then Exit Sub
            If Los_Pack.Enabled = True Then Exit Sub
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyDelete
            If grdsales.rows <= 1 Then Exit Sub
            If M_EDIT = True Then Exit Sub
            If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(grdsales.Row, 2) & """", vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then
                grdsales.SetFocus
                Exit Sub
            End If
            
            Dim i As Long
            Dim rststock As ADODB.Recordset
            Dim RSTRTRXFILE As ADODB.Recordset
            Dim rstMaxNo As ADODB.Recordset
            
            Screen.MousePointer = vbHourglass
            On Error GoTo ERRHAND
            db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(grdsales.Row, 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(grdsales.Row, 16)) & ""
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            db.BeginTrans
            With rststock
                If Not (.EOF And .BOF) Then
                    .Properties("Update Criteria").Value = adCriteriaKey
                    !RCPT_QTY = !RCPT_QTY - Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 5))
                    If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                    !RCPT_VAL = !RCPT_VAL - Val(grdsales.TextMatrix(grdsales.Row, 13))
                    
                    !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(grdsales.Row, 3)) * Val(grdsales.TextMatrix(grdsales.Row, 5))
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(grdsales.Row, 13))
                    rststock.Update
                End If
            End With
            db.CommitTrans
            rststock.Close
            Set rststock = Nothing
            
            i = 0
            Set rstMaxNo = New ADODB.Recordset
            rstMaxNo.Open "Select MAX(LINE_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ", db, adOpenStatic, adLockReadOnly
            If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
                i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
            End If
            rstMaxNo.Close
            Set rstMaxNo = Nothing
            
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
            db.BeginTrans
            Do Until RSTRTRXFILE.EOF
                RSTRTRXFILE!LINE_NO = i
                i = i + 1
                RSTRTRXFILE.Update
                RSTRTRXFILE.MoveNext
            Loop
            db.CommitTrans
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            i = 1
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND  TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
            db.BeginTrans
            Do Until RSTRTRXFILE.EOF
                RSTRTRXFILE!LINE_NO = i
                i = i + 1
                RSTRTRXFILE.Update
                RSTRTRXFILE.MoveNext
            Loop
            db.CommitTrans
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            grdsales.rows = 1
            i = 0
            LBLTOTAL.Caption = ""
            lbltotalwodiscount = ""
            LBLTOTALTAX.Caption = ""
            LBLGROSSAMT.Caption = ""
            LBLEXP.Caption = ""
            lblqty.Caption = ""
            grdsales.rows = 1
            Dim GROSSVAL As Double
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
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
                grdsales.TextMatrix(i, 15) = IIf(IsNull(RSTRTRXFILE!check_flag), "N", RSTRTRXFILE!check_flag)
                grdsales.TextMatrix(i, 16) = RSTRTRXFILE!LINE_NO
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
                GROSSVAL = (Val(grdsales.TextMatrix(i, 9)) * IIf(Val(grdsales.TextMatrix(i, 5)) = 0, 1, Val(grdsales.TextMatrix(i, 5)))) * (Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 14)))
                If RSTRTRXFILE!DISC_FLAG = "P" Then
                    grdsales.TextMatrix(i, 27) = "P"
                    LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - (GROSSVAL * Val(grdsales.TextMatrix(i, 17)) / 100)) * Val(grdsales.TextMatrix(i, 10)) / 100, 2))
                Else
                    grdsales.TextMatrix(i, 27) = "A"
                    LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - Val(grdsales.TextMatrix(i, 17))) * Val(grdsales.TextMatrix(i, 10)) / 100, 2))
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
                grdsales.TextMatrix(i, 38) = IIf(IsNull(RSTRTRXFILE!BARCODE), "", RSTRTRXFILE!BARCODE)
                grdsales.TextMatrix(i, 39) = IIf(IsNull(RSTRTRXFILE!cess_amt), "", RSTRTRXFILE!cess_amt)
                grdsales.TextMatrix(i, 40) = IIf(IsNull(RSTRTRXFILE!CESS_PER), "", RSTRTRXFILE!CESS_PER)
                lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
                LBLGROSSAMT.Caption = Val(LBLGROSSAMT.Caption) + Val(grdsales.TextMatrix(i, 8)) * Val(RSTRTRXFILE!QTY)
                LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(i, 32)), ".00")
                lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(i, 3)), ".00")
                'TXTDEALER.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
                
                'TXTINVDATE.Text = Format(RSTRTRXFILE!VCH_DATE, "DD/MM/YYYY")
                'TXTREMARKS.Text = Mid(RSTRTRXFILE!VCH_DESC, 15)
                'TXTINVOICE.Text = IIf(IsNull(RSTRTRXFILE!PINV), "", RSTRTRXFILE!PINV)
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
            If Roundflag = True Then
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
            Else
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 2), "0.00")
            End If
            LBLNET.Caption = Val(LBLGROSSAMT.Caption) + Val(LBLTOTALTAX.Caption)
            TXTSLNO.text = Val(grdsales.rows)
            
            M_ADD = True
            OLD_BILL = True
            grdsales.SetFocus
            Screen.MousePointer = vbNormal
        End Select
        Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub grdsales_Scroll()
    TXTsample.Visible = False
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            On Error Resume Next
            TXTITEMCODE.text = grdtmp.Columns(0)
            TXTPRODUCT.text = grdtmp.Columns(1)
            'lblavlqty.Caption = IIf(IsNull(grdtmp.Columns(2)), "", grdtmp.Columns(2))
            On Error Resume Next
            Set Image1.DataSource = PHY
            If IsNull(PHY!PHOTO) Then
                Frame6.Visible = False
                Set Image1.DataSource = Nothing
                bytData = ""
            Else
                If err.Number = 545 Then
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
            On Error GoTo ERRHAND
            For i = 1 To grdsales.rows - 1
                If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                    If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                    Exit For
                End If
            Next i
            
            Set RSTRXFILE = New ADODB.Recordset
            'RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly, adCmdText
            RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND TRX_TYPE = 'PI' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                'RSTRXFILE.MoveLast
                TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                Los_Pack.text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
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
                TXTEXPIRY.text = IIf(IsDate(RSTRXFILE!EXP_DATE), Format(RSTRXFILE!EXP_DATE, "MM/YY"), "  /  ")
                If IsNull(RSTRXFILE!MRP) Then
                    TXTRATE.text = ""
                Else
                    TXTRATE.text = IIf(IsNull(RSTRXFILE!MRP), "", Format(Round(Val(RSTRXFILE!MRP), 2), ".000"))
                End If
                If IsNull(RSTRXFILE!MRP_BT) Then
                    txtmrpbt.text = 100 * Val(TXTRATE.text) / 105
                Else
                    txtmrpbt.text = Format(Val(RSTRXFILE!MRP_BT), ".000")
                End If
                If IsNull(RSTRXFILE!PTR) Then
                    TXTPTR.text = ""
                Else
                    TXTPTR.text = Format(Round(Val(RSTRXFILE!PTR) * Val(Los_Pack.text), 3), ".000")
                End If
                If IsNull(RSTRXFILE!P_RETAIL) Then
                    TXTRETAIL.text = ""
                Else
                    TXTRETAIL.text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                End If
                'TXTPTR.Text = IIf(IsNull(RSTRXFILE!PTR), "", Format(Round(Val(RSTRXFILE!PTR), 2), ".000"))
                'txtretail.Text = IIf(IsNull(RSTRXFILE!P_RETAIL), "", Format(Round(Val(RSTRXFILE!P_RETAIL) * Val(Los_Pack.Text), 2), ".000"))
                If IsNull(RSTRXFILE!P_WS) Then
                    txtWS.text = ""
                Else
                    txtWS.text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                End If
                If IsNull(RSTRXFILE!P_VAN) Then
                    txtvanrate.text = ""
                Else
                    txtvanrate.text = Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000")
                End If
                If IsNull(RSTRXFILE!P_CRTN) Then
                    txtcrtn.text = ""
                Else
                    txtcrtn.text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                End If
                If IsNull(RSTRXFILE!P_LWS) Then
                    TxtLWRate.text = ""
                Else
                    TxtLWRate.text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                End If
                If IsNull(RSTRXFILE!CRTN_PACK) Then
                    txtcrtnpack.text = ""
                Else
                    txtcrtnpack.text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                End If
                If IsNull(RSTRXFILE!SALES_PRICE) Then
                    txtprofit.text = ""
                Else
                    txtprofit.text = Format(Round(Val(RSTRXFILE!SALES_PRICE), 2), ".000")
                End If
                If IsNull(RSTRXFILE!SALES_TAX) Then
                    TxttaxMRP.text = ""
                Else
                    TxttaxMRP.text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                End If
                If IsNull(RSTRXFILE!EXDUTY) Then
                    TxtExDuty.text = ""
                Else
                    TxtExDuty.text = Format(Val(RSTRXFILE!EXDUTY), ".00")
                End If
                If IsNull(RSTRXFILE!CSTPER) Then
                    TxtCSTper.text = ""
                Else
                    TxtCSTper.text = Format(Val(RSTRXFILE!CSTPER), ".00")
                End If
                If IsNull(RSTRXFILE!TR_DISC) Then
                    TxtTrDisc.text = ""
                Else
                    TxtTrDisc.text = Format(Val(RSTRXFILE!TR_DISC), ".00")
                End If
                If IsNull(RSTRXFILE!cess_amt) Then
                    txtCess.text = ""
                Else
                    txtCess.text = Format(Val(RSTRXFILE!cess_amt), ".00")
                End If
                If IsNull(RSTRXFILE!CESS_PER) Then
                    TxtCessPer.text = ""
                Else
                    TxtCessPer.text = Format(Val(RSTRXFILE!CESS_PER), ".00")
                End If
                TxtWarranty.text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                If RSTRXFILE!COM_FLAG = "A" Then
                    TxtComAmt.text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                    OptComAmt.Value = True
                Else
                    TxtComper.text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                    OptComper.Value = True
                End If
                On Error Resume Next
                CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                CmbWrnty.text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                On Error GoTo ERRHAND
                
                ''TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                If RSTRXFILE!check_flag = "M" Then
                    OPTTaxMRP.Value = True
                ElseIf RSTRXFILE!check_flag = "V" Then
                    OPTVAT.Value = True
                Else
                    OPTNET.Value = True
                End If
                
                '==========
                If MDIMAIN.lblPerPurchase.Caption = "Y" Then
                    Txtgrossamt.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxtExDuty.text) / 100)
                    Txtgrossamt.Tag = Val(Txtgrossamt.Tag) + (Val(TXTPTR.text) * Val(TxtCSTper.text) / 100)
                    If OPTVAT.Value = True Then
                       If optdiscper.Value = True Then
                            LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                            LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                        Else
                            LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                            LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                        End If
                    Else
                        TxttaxMRP.text = 0
                        If optdiscper.Value = True Then
                            LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                            LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                        Else
                            LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                            LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                        End If
                    End If
                    Call TXTRETAIL_LostFocus
                    Call txtws_LostFocus
                    Call txtvanrate_LostFocus
                    TXTRETAIL.BackColor = vbWhite
                    txtWS.BackColor = vbWhite
                    txtvanrate.BackColor = vbWhite
                End If
                '==========
            Else
                TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                Txtpack.text = 1
                Los_Pack.text = 1
                TxtWarranty.text = ""
                On Error Resume Next
                CmbPack.text = "Nos"
                CmbWrnty.ListIndex = -1
                On Error GoTo ERRHAND
                
                TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                txtBatch.text = ""
                TXTEXPIRY.text = "  /  "
                TXTRATE.text = ""
                txtmrpbt.text = ""
                TXTPTR.text = ""
                TxtNetrate.text = ""
                TXTRETAIL.text = ""
                txtWS.text = ""
                txtvanrate.text = ""
                txtcrtn.text = ""
                TxtLWRate.text = ""
                txtcrtnpack.text = ""
                txtprofit.text = ""
                'TxttaxMRP.text = ""
                Los_Pack.text = "1"
                TxtWarranty.text = ""
                On Error Resume Next
                CmbPack.text = "Nos"
                CmbWrnty.ListIndex = -1
                On Error GoTo ERRHAND
                OPTVAT.Value = True
            End If
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing
            
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With RSTRXFILE
                If Not (.EOF And .BOF) Then
                    lblcategory.Caption = IIf(IsNull(RSTRXFILE!Category), "", RSTRXFILE!Category)
                    Cmbcategory.text = IIf(IsNull(RSTRXFILE!Category), "", RSTRXFILE!Category)
                    
                    If IsNull(RSTRXFILE!P_RETAIL) Then
                        TXTRETAIL.text = ""
                    Else
                        TXTRETAIL.text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.text = ""
                    Else
                        TxttaxMRP.text = Format(Round(Val(RSTRXFILE!SALES_TAX), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_WS) Then
                        txtWS.text = ""
                    Else
                        txtWS.text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_VAN) Then
                        txtvanrate.text = ""
                    Else
                        txtvanrate.text = Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000")
                    End If
                    If RSTRXFILE!COM_FLAG = "A" Then
                        TxtComAmt.text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                        OptComAmt.Value = True
                    Else
                        TxtComper.text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                        OptComper.Value = True
                    End If
                    If IsNull(RSTRXFILE!P_CRTN) Then
                        txtcrtn.text = ""
                    Else
                        txtcrtn.text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_LWS) Then
                        TxtLWRate.text = ""
                    Else
                        TxtLWRate.text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!CRTN_PACK) Then
                        txtcrtnpack.text = ""
                    Else
                        txtcrtnpack.text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.text = ""
                    Else
                        TxttaxMRP.text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                    End If
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    cmbfull.text = IIf(IsNull(RSTRXFILE!FULL_PACK), 0, RSTRXFILE!FULL_PACK)
                    On Error GoTo ERRHAND
                    '==========
                    If MDIMAIN.lblPerPurchase.Caption = "Y" Then
                        Txtgrossamt.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxtExDuty.text) / 100)
                        Txtgrossamt.Tag = Val(Txtgrossamt.Tag) + (Val(TXTPTR.text) * Val(TxtCSTper.text) / 100)
                        If OPTVAT.Value = True Then
                           If optdiscper.Value = True Then
                                LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                            Else
                                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                            End If
                        Else
                            TxttaxMRP.text = 0
                            If optdiscper.Value = True Then
                                LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                            Else
                                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                            End If
                        End If
                        Call TXTRETAIL_LostFocus
                        Call txtws_LostFocus
                        Call txtvanrate_LostFocus
                        TXTRETAIL.BackColor = vbWhite
                        txtWS.BackColor = vbWhite
                        txtvanrate.BackColor = vbWhite
                        
                    End If
                    '==========
                
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
            txtcategory.Enabled = False
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                Los_Pack.text = 1
                TXTQTY.text = 1
                TxTfree.text = ""
                TXTRATE.text = ""
                TXTPTR.Enabled = True
                TXTPTR.SetFocus
            Else
                Los_Pack.Enabled = True
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
            End If
            'TxtPack.Enabled = True
            'TxtPack.SetFocus
        Case vbKeyEscape
            TXTQTY.text = ""
            TxTfree.text = ""
            Fram.Enabled = True
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.Enabled = True
            TXTITEMCODE.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Los_Pack_LostFocus()
    Call CHANGEBOXCOLOR(Los_Pack, False)
End Sub

Private Sub OptComper_LostFocus()
    cmbfull.BackColor = vbWhite
End Sub

Private Sub OptCr_Click()
    Call TXTDEALER_Change
End Sub

Private Sub Optdiscamt_Click()
    Call TxttaxMRP_LostFocus
End Sub

Private Sub optdiscper_Click()
    Call TxttaxMRP_LostFocus
End Sub

Private Sub OptDr_Click()
    Call TXTDEALER_Change
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

Private Sub OPTNET_LostFocus()
    OPTNET.BackColor = vbWhite
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

Private Sub OPTTaxMRP_LostFocus()
    OPTTaxMRP.BackColor = vbWhite
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

Private Sub OPTVAT_LostFocus()
    OPTVAT.BackColor = vbWhite
End Sub

Private Sub txtbarcode_GotFocus()
    FRMEQTY.Visible = False
    Call CHANGEBOXCOLOR(TxtBarcode, True)
    Call CHANGEBOXCOLOR(txtcrtn, False)
    
    TxtBarcode.SelStart = 0
    TxtBarcode.SelLength = Len(TxtBarcode.text)
    FRMEGRDTMP.Visible = False
    TXTSLNO.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    Cmbcategory.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxtPoints.Enabled = False
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
    TxtCustDisc.Enabled = False
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

Private Sub TxtBarcode_LostFocus()
    Call CHANGEBOXCOLOR(TxtBarcode, False)
End Sub

Private Sub TXTBATCH_GotFocus()
    Call CHANGEBOXCOLOR(txtBatch, True)
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                TxttaxMRP.Enabled = True
                TxttaxMRP.SetFocus
            Else
                TXTEXPIRY.Visible = True
                TXTEXPIRY.SetFocus
            End If
        Case vbKeyEscape
            TXTPTR.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBatch_LostFocus()
    Call CHANGEBOXCOLOR(txtBatch, False)
End Sub

Private Sub txtBillNo_GotFocus()
    Call CHANGEBOXCOLOR(txtBillNo, True)
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.text)
    'txtBillNo.ForeColor = &HFFFF&
End Sub

Public Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim TRXFILE As ADODB.Recordset
    Dim i As Long

    On Error GoTo ERRHAND
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
            PONO = ""
            CMBPO.text = ""
            LBLTOTAL.Caption = ""
            lbltotalwodiscount = ""
            LBLTOTALTAX.Caption = ""
            LBLGROSSAMT.Caption = ""
            LBLEXP.Caption = ""
            lblqty.Caption = ""
            Dim GROSSVAL As Double
            grdsales.rows = 1
            OLD_BILL = False
            lbloldbills.Caption = "N"
            OptCr.Value = True
            
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                
                If rstTRXMAST!SLSM_CODE = "DR" Then
                    OptDr.Value = True
                Else
                    OptCr.Value = True
                End If
                TXTDISCAMOUNT.text = IIf(IsNull(rstTRXMAST!DISCOUNT), "", Format(rstTRXMAST!DISCOUNT, ".00"))
                txtaddlamt.text = IIf(IsNull(rstTRXMAST!ADD_AMOUNT), "", Format(rstTRXMAST!ADD_AMOUNT, ".00"))
                txtcramt.text = IIf(IsNull(rstTRXMAST!DISC_PERS), "", Format(rstTRXMAST!DISC_PERS, ".00"))
                TxtCST.text = IIf(IsNull(rstTRXMAST!CST_PER), "", Format(rstTRXMAST!CST_PER, ".00"))
                TxtInsurance.text = IIf(IsNull(rstTRXMAST!INS_PER), "", Format(rstTRXMAST!INS_PER, ".00"))
                'If rstTRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                lblcredit.Caption = "1"
                TXTREMARKS.text = IIf(IsNull(rstTRXMAST!REMARKS), "", rstTRXMAST!REMARKS)
                On Error Resume Next
'                If grdsales.Rows <= 1 Then
'                    TXTINVDATE.Text = "  /  /    "
'                    TXTRCVDATE.Text = "  /  /    "
'                    TXTDATE.Text = Format(Date, "DD/MM/YYYY")
'                Else
                    TXTINVDATE.text = IIf(IsDate(rstTRXMAST!VCH_DATE), Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY"), "  /  /    ")
                    TXTRCVDATE.text = IIf(IsDate(rstTRXMAST!RCVD_DATE), Format(rstTRXMAST!RCVD_DATE, "DD/MM/YYYY"), TXTINVDATE.text)
                    TXTDATE.text = IIf(IsDate(rstTRXMAST!VCH_DATE), Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY"), Format(rstTRXMAST!CREATE_DATE, "DD/MM/YYYY"))
'                End If
                On Error GoTo ERRHAND
                TXTINVOICE.text = IIf(IsNull(rstTRXMAST!PINV), "", rstTRXMAST!PINV)
                CMBDISTRICT.text = IIf(IsNull(rstTRXMAST!TRX_GODOWN), "", rstTRXMAST!TRX_GODOWN)
                TXTDEALER.text = IIf(IsNull(rstTRXMAST!ACT_NAME), "", rstTRXMAST!ACT_NAME)
                
                Set TRXFILE = New ADODB.Recordset
                TRXFILE.Open "SELECT max(VCH_DATE) FROM TRANSMAST ", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (TRXFILE.EOF And TRXFILE.BOF) Then
                    lbllastdate.Caption = IIf(IsNull(TRXFILE.Fields(0)), 1, TRXFILE.Fields(0))
                End If
                TRXFILE.Close
                Set TRXFILE = Nothing
                
                If IsDate(lbllastdate.Caption) And IsDate(TXTINVDATE.text) Then
                    If DateValue(lbllastdate.Caption) <> DateValue(Date) Then
                        lbloldbills.Caption = "Y"
                    End If
                    If DateValue(Date) <> DateValue(TXTINVDATE.text) Then
                        lbloldbills.Caption = "Y"
                    End If
                End If
                
                'OLD_BILL = True
            Else
                TXTDATE.text = Format(Date, "DD/MM/YYYY")
                TXTINVDATE.text = "  /  /    "
                TXTRCVDATE.text = "  /  /    "
                TXTREMARKS.text = ""
                TXTDEALER.text = ""
                TXTINVOICE.text = ""
                CMBDISTRICT.text = ""
                'OLD_BILL = False
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
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
                grdsales.TextMatrix(i, 6) = Format(rstTRXMAST!MRP, ".000")
                grdsales.TextMatrix(i, 7) = Format(rstTRXMAST!SALES_PRICE, ".00000")
                grdsales.TextMatrix(i, 8) = Format(rstTRXMAST!ITEM_COST, ".00000")
                grdsales.TextMatrix(i, 9) = Format(rstTRXMAST!PTR, ".0000")
                grdsales.TextMatrix(i, 10) = IIf(Val(rstTRXMAST!SALES_TAX) = 0, "", Format(rstTRXMAST!SALES_TAX, ".00"))
                grdsales.TextMatrix(i, 11) = IIf(IsNull(rstTRXMAST!REF_NO), "", rstTRXMAST!REF_NO)
                grdsales.TextMatrix(i, 12) = IIf(IsNull(rstTRXMAST!EXP_DATE), "", Format(rstTRXMAST!EXP_DATE, "DD/MM/YYYY"))
                grdsales.TextMatrix(i, 13) = Format(rstTRXMAST!TRX_TOTAL, ".000")
                grdsales.TextMatrix(i, 14) = IIf(IsNull(rstTRXMAST!SCHEME), "", rstTRXMAST!SCHEME)
                grdsales.TextMatrix(i, 15) = IIf(IsNull(rstTRXMAST!check_flag), "N", rstTRXMAST!check_flag)
                grdsales.TextMatrix(i, 16) = rstTRXMAST!LINE_NO
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
                grdsales.TextMatrix(i, 26) = IIf(IsNull(rstTRXMAST!gross_amt), 0, Format(rstTRXMAST!gross_amt, "0.00"))
                GROSSVAL = (Val(grdsales.TextMatrix(i, 9)) * IIf(Val(grdsales.TextMatrix(i, 5)) = 0, 1, Val(grdsales.TextMatrix(i, 5)))) * (Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 14)))
                If rstTRXMAST!DISC_FLAG = "P" Then
                    grdsales.TextMatrix(i, 27) = "P"
                    LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - (GROSSVAL * Val(grdsales.TextMatrix(i, 17)) / 100)) * Val(grdsales.TextMatrix(i, 10)) / 100, 2))
                Else
                    grdsales.TextMatrix(i, 27) = "A"
                    LBLTOTALTAX.Caption = Val(LBLTOTALTAX.Caption) + (Round((GROSSVAL - Val(grdsales.TextMatrix(i, 17))) * Val(grdsales.TextMatrix(i, 10)) / 100, 2))
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
                grdsales.TextMatrix(i, 39) = IIf(IsNull(rstTRXMAST!cess_amt), "", rstTRXMAST!cess_amt)
                grdsales.TextMatrix(i, 40) = IIf(IsNull(rstTRXMAST!CESS_PER), "", rstTRXMAST!CESS_PER)
                grdsales.TextMatrix(i, 41) = Val(grdsales.TextMatrix(i, 3))
                LBLGROSSAMT.Caption = Val(LBLGROSSAMT.Caption) + Val(grdsales.TextMatrix(i, 8)) * Val(rstTRXMAST!QTY)
                lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
                LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(i, 32)), ".00")
                lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(i, 3)), ".00")
                'TXTDEALER.Text = IIf(IsNull(rstTRXMAST!VCH_DESC), "", Mid(rstTRXMAST!VCH_DESC, 15))
                PONO = IIf(IsNull(rstTRXMAST!PO_NO), "", rstTRXMAST!PO_NO)
                On Error Resume Next
                TXTINVDATE.text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                TXTRCVDATE.text = Format(rstTRXMAST!RCVD_DATE, "DD/MM/YYYY")
                OLD_BILL = True
                NEW_BILL = False
                On Error GoTo ERRHAND
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            If OLD_BILL = False Then NEW_BILL = True
                
            LBLNET.Caption = Val(LBLGROSSAMT.Caption) + Val(LBLTOTALTAX.Caption)
            
            ''''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
            'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
            If Roundflag = True Then
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
            Else
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 2), "0.00")
            End If
            
            TXTSLNO.text = grdsales.rows
            TXTSLNO.Enabled = True
            txtBillNo.Enabled = False
            FRMEMASTER.Enabled = True
            If grdsales.rows > 1 Then
                FRMEMASTER.Enabled = True
                FRMECONTROLS.Enabled = True
                cmdRefresh.Enabled = True
                CMDTRANSFER.Enabled = True
                cmdRefresh.SetFocus
            Else
                TXTDEALER.SetFocus
            End If
            
'            Set RSTTRNSMAST = New ADODB.Recordset
'            RSTTRNSMAST.Open "Select CHECK_FLAG From TRANSMAST WHERE TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
'            If Not (RSTTRNSMAST.EOF Or RSTTRNSMAST.BOF) Then
'                If RSTTRNSMAST!CHECK_FLAG = "Y" Then FRMEMASTER.Enabled = False
'            End If
'            RSTTRNSMAST.Close
'            Set RSTTRNSMAST = Nothing
    
    End Select
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
    CMBPO.text = PONO
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
    Call CHANGEBOXCOLOR(txtBillNo, False)
    If Val(txtBillNo.text) = 0 Or Val(txtBillNo.text) > Val(TXTLASTBILL.text) Then txtBillNo.text = TXTLASTBILL.text
    'txtBillNo.BackColor = &HFFFFFF
    'txtBillNo.ForeColor = &H0&
End Sub

Private Sub txtcategory_Change()
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, MRP, ITEM_COST, SALES_TAX, ITEM_NET_COST, P_RETAIL, P_WS From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.text & "%' OR ITEM_NAME Like '" & Me.txtcategory.text & "%') AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHYFLAG = False
         Else
             PHY.Close
             'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, MRP, ITEM_COST, SALES_TAX, ITEM_NET_COST, P_RETAIL, P_WS From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.text & "%' OR ITEM_NAME Like '" & Me.txtcategory.text & "%') AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         End If
         
        Set grdtmp.DataSource = PHY
        
        If PHY.RecordCount > 0 Then
            FRMEGRDTMP.Visible = True
        Else
            FRMEGRDTMP.Visible = False
            Exit Sub
        End If
'        grdtmp.Columns(0).Visible = True
        grdtmp.Columns(0).Caption = "ITEM CODE"
        grdtmp.Columns(0).Width = 1300
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(1).Width = 5000
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(2).Width = 900
        grdtmp.Columns(2).Caption = "QTY"
        grdtmp.Columns(3).Width = 900
        grdtmp.Columns(3).Caption = "MRP"
        grdtmp.Columns(4).Width = 950
        grdtmp.Columns(4).Caption = "COST"
        grdtmp.Columns(5).Caption = "TAX%"
        grdtmp.Columns(5).Width = 800
        grdtmp.Columns(6).Caption = "NET COST"
        grdtmp.Columns(6).Width = 950
        grdtmp.Columns(7).Caption = "R. PRICE"
        grdtmp.Columns(7).Width = 950
        grdtmp.Columns(8).Caption = "W. PRICE"
        grdtmp.Columns(8).Width = 950

        Exit Sub
ERRHAND:
        MsgBox err.Description
End Sub

Private Sub txtcategory_GotFocus()
    FRMEQTY.Visible = False
    Call CHANGEBOXCOLOR(txtcategory, True)
    Call CHANGEBOXCOLOR(txtcrtn, False)
    Call CHANGEBOXCOLOR(TxtLWRate, False)
    
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.text)
    FRMEGRDTMP.Visible = False
    'TXTSLNO.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    Cmbcategory.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxtPoints.Enabled = False
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
    TxtCustDisc.Enabled = False
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

Private Sub txtcategory_LostFocus()
    Call CHANGEBOXCOLOR(txtcategory, False)
End Sub

Private Sub TxtCSTper_LostFocus()
    Call CHANGEBOXCOLOR(TxtCSTper, False)
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TXTDEALER_LostFocus()
    Call CHANGEBOXCOLOR(TXTDEALER, False)
End Sub

Private Sub TxtExDuty_LostFocus()
    Call CHANGEBOXCOLOR(TxtExDuty, False)
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TXTEXPDATE_GotFocus()
    TXTEXPDATE.BackColor = &H98F3C1
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
End Sub

Private Sub TXTEXPDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Len(Trim(TXTEXPDATE.text)) = 4 Then GoTo SKID
            If Not IsDate(TXTEXPDATE.text) Then Exit Sub
            If DateDiff("d", Date, TXTEXPDATE.text) < 0 Then
                MsgBox "Item Expired....", vbOKOnly, "EzBiz"
                TXTEXPDATE.SelStart = 0
                TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.text) < 60 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days", vbOKOnly, "EzBiz"
                TXTEXPDATE.SelStart = 0
                TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.text) < 180 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "EzBiz") = vbNo Then
                    TXTEXPDATE.SelStart = 0
                    TXTEXPDATE.SelLength = Len(TXTEXPDATE.text)
                    TXTEXPDATE.SetFocus
                    Exit Sub
                End If
            End If
SKID:
            TxttaxMRP.Enabled = True
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
            TxttaxMRP.SetFocus
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
    TXTEXPDATE.BackColor = vbWhite
    TXTEXPDATE.text = Format(TXTEXPDATE.text, "DD/MM/YYYY")
    If IsDate(TXTEXPDATE.text) Then TXTEXPIRY.text = Format(TXTEXPDATE.text, "MM/YY")
End Sub

Private Sub TxtExpense_LostFocus()
    Call CHANGEBOXCOLOR(TxtExpense, False)
End Sub

Private Sub TxtFree_GotFocus()
    Call CHANGEBOXCOLOR(TxTfree, True)
    TxTfree.SelStart = 0
    TxTfree.SelLength = Len(TxTfree.text)
End Sub

Private Sub TxtFree_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRATE.SetFocus
        Case vbKeyEscape
            TXTQTY.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TxTfree, False)
    If Val(TxTfree.text) = 0 Then TxTfree.text = 0
    TxTfree.text = Format(TxTfree.text, "0.00")
End Sub

Private Sub TxtHSN_LostFocus()
    Call CHANGEBOXCOLOR(txtHSN, False)
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.BackColor = &H98F3C1
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
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTINVDATE.SetFocus
                Exit Sub
            End If
            
            If (DateValue(TXTINVDATE.text) < DateValue(MDIMAIN.DTFROM.Value)) Or (DateValue(TXTINVDATE.text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.Value))) Then
                'db.Execute "delete from Users"
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTINVDATE.SetFocus
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

Private Sub TXTINVDATE_LostFocus()
    TXTINVDATE.BackColor = vbWhite
    If Not IsDate(TXTRCVDATE.text) And IsDate(TXTINVDATE.text) Then
       TXTRCVDATE.text = Format(TXTINVDATE, "DD/MM/YYYY")
    End If
End Sub

Private Sub TXTINVOICE_GotFocus()
    Call CHANGEBOXCOLOR(TXTINVOICE, True)
    TXTINVOICE.BackColor = &H98F3C1
    TXTINVOICE.SelStart = 0
    TXTINVOICE.SelLength = Len(TXTINVOICE.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTINVOICE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVOICE.text = "" Then
                MsgBox "Please enter the Invoice Number", vbOKOnly, "EzBiz"
                TXTINVOICE.SetFocus
                Exit Sub
            End If
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND PINV = '" & Trim(TXTINVOICE.text) & "' AND VCH_NO <> " & Val(txtBillNo.text) & " AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                MsgBox "You have already entered this Invoice number for " & Trim(DataList2.text) & " as Computer Bill No. " & rstTRXMAST!VCH_NO, vbOKOnly, "EzBiz"
                TXTINVOICE.SetFocus
            Else
                TXTINVDATE.SetFocus
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
        Case vbKeyEscape
            DataList2.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTINVOICE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTINVOICE_LostFocus()
    Call CHANGEBOXCOLOR(TXTINVOICE, False)
End Sub

Private Sub Txtpack_GotFocus()
    Call CHANGEBOXCOLOR(Txtpack, True)
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

Private Sub TxtPoints_GotFocus()
    Call CHANGEBOXCOLOR(TxtPoints, True)
    TxtPoints.SelStart = 0
    TxtPoints.SelLength = Len(TxtPoints.text)
End Sub

Private Sub TxtPoints_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(Txtgrossamt.Text) = 0 Then Exit Sub
            If Trim(txtHSN.text) = "" Then
                txtHSN.Enabled = True
                txtHSN.SetFocus
            Else
                txtPD.SetFocus
            End If
        Case vbKeyEscape
            TxttaxMRP.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtPoints_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtPoints_LostFocus()
    Call CHANGEBOXCOLOR(TxtPoints, False)
    TxtPoints.text = Format(TxtPoints.text, "0.00")
End Sub

Private Sub TXTPRODUCT_Change()
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, MRP, ITEM_COST, SALES_TAX, ITEM_NET_COST, P_RETAIL, P_WS From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHYFLAG = False
         Else
             PHY.Close
             PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, MRP, ITEM_COST, SALES_TAX, ITEM_NET_COST, P_RETAIL, P_WS From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         End If
         
        Set grdtmp.DataSource = PHY
        
        If PHY.RecordCount > 0 Then
            FRMEGRDTMP.Visible = True
        Else
            FRMEGRDTMP.Visible = False
            Exit Sub
        End If
        grdtmp.Columns(0).Caption = "ITEM CODE"
        grdtmp.Columns(0).Width = 1300
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(1).Width = 5000
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(2).Width = 900
        grdtmp.Columns(2).Caption = "QTY"
        grdtmp.Columns(3).Width = 900
        grdtmp.Columns(3).Caption = "MRP"
        grdtmp.Columns(4).Width = 950
        grdtmp.Columns(4).Caption = "COST"
        grdtmp.Columns(5).Caption = "TAX%"
        grdtmp.Columns(5).Width = 800
        grdtmp.Columns(6).Caption = "NET COST"
        grdtmp.Columns(6).Width = 950
        grdtmp.Columns(7).Caption = "R. PRICE"
        grdtmp.Columns(7).Width = 950
        grdtmp.Columns(8).Caption = "W. PRICE"
        grdtmp.Columns(8).Width = 950
        Exit Sub
ERRHAND:
        MsgBox err.Description
                
End Sub

Private Sub TXTPRODUCT_GotFocus()
    FRMEQTY.Visible = False
    Call CHANGEBOXCOLOR(TXTPRODUCT, True)
    Call CHANGEBOXCOLOR(txtcrtn, False)
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.text)
    If Trim(TXTPRODUCT.text) <> "" Or Trim(txtcategory.text) <> "" Then Call TXTPRODUCT_Change
    'TXTSLNO.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    Cmbcategory.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxtPoints.Enabled = False
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
    TxtCustDisc.Enabled = False
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
    Dim RSTRXFILE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
        Case vbKeyReturn
            On Error Resume Next
            TXTITEMCODE.text = ""
            TXTITEMCODE.text = grdtmp.Columns(0)
            If Trim(TXTPRODUCT.text) = "" Then Exit Sub
            If TXTITEMCODE.text <> "" Then
                Call TxtItemcode_KeyDown(13, 0)
                Exit Sub
            End If
            If Trim(txtcategory.text) = "" Then
                If MsgBox("Item not exists!!! Do You want to add this item?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
                TXTPRODUCT.Tag = ""
'                Set RSTITEMMAST = New ADODB.Recordset
'                RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
'                If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'                    If IsNull(RSTITEMMAST.Fields(0)) Then
'                        TXTPRODUCT.Tag = 1
'                    Else
'                        TXTPRODUCT.Tag = Val(RSTITEMMAST.Fields(0)) + 1
'                    End If
'                End If
'                RSTITEMMAST.Close
'                Set RSTITEMMAST = Nothing
                
                On Error GoTo ERRHAND
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "Select * From ITEMMAST WHERE ITEM_CODE= (SELECT MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) FROM ITEMMAST)", db, adOpenStatic, adLockOptimistic, adCmdText
                TXTPRODUCT.Tag = RSTITEMMAST!ITEM_CODE + 1
                If TXTPRODUCT.Tag = "" Then TXTPRODUCT.Tag = 1
                db.BeginTrans
                RSTITEMMAST.AddNew
                'RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
                RSTITEMMAST!ITEM_CODE = Val(TXTPRODUCT.Tag)
                RSTITEMMAST!ITEM_NAME = Trim(TXTPRODUCT.text)
                If Cmbcategory.text = "" Then
                    RSTITEMMAST!Category = "GENERAL"
                Else
                    RSTITEMMAST!Category = Cmbcategory.text
                End If
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
                RSTITEMMAST!SALES_TAX = Val(TxttaxMRP.text)
                RSTITEMMAST!ITEM_COST = 0
                RSTITEMMAST!P_RETAIL = 0
                RSTITEMMAST!P_WS = 0
                RSTITEMMAST!CRTN_PACK = 1
                RSTITEMMAST!P_CRTN = 0
                RSTITEMMAST!LOOSE_PACK = 1
                RSTITEMMAST!UN_BILL = "N"
                If PC_FLAG = "Y" Then
                    RSTITEMMAST!PRICE_CHANGE = "Y"
                Else
                    RSTITEMMAST!PRICE_CHANGE = "N"
                End If
                RSTITEMMAST.Update
                db.CommitTrans
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                TXTITEMCODE.text = TXTPRODUCT.Tag
                
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT CATEGORY FROM CATEGORY where CATEGORY = 'GENERAL'", db, adOpenStatic, adLockOptimistic, adCmdText
                If (rststock.EOF And rststock.BOF) Then
                    rststock.AddNew
                    rststock!Category = "GENERAL"
                    rststock.Update
                End If
                rststock.Close
                Set rststock = Nothing


                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT MANUFACTURER FROM MANUFACT where MANUFACTURER = 'GENERAL'", db, adOpenStatic, adLockOptimistic, adCmdText
                If (rststock.EOF And rststock.BOF) Then
                    rststock.AddNew
                    rststock!MANUFACTURER = "GENERAL"
                    rststock.Update
                End If
                rststock.Close
                Set rststock = Nothing
                    
                Call TxtItemcode_KeyDown(13, 0)
                Cmbcategory.Enabled = True
                Cmbcategory.SetFocus
                'frmitemmaster.Show
                'frmitemmaster.TXTITEM.Text = Trim(TXTPRODUCT.Text)
                'frmitemmaster.LBLLP.Caption = "P"
                'MsgBox "Item not found!!!!", , "EzBiz"
                Exit Sub
            Else
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME = '" & Trim(TXTPRODUCT.text) & "' ", db, adOpenForwardOnly
                If Not (RSTITEMMAST.EOF Or RSTITEMMAST.BOF) Then
                    MsgBox "Item Name already exists with Item Code " & RSTITEMMAST!ITEM_CODE, , "EzBiz"
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
                    Exit Sub
                End If
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                        
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(txtcategory.text) & "' ", db, adOpenForwardOnly
                If Not (RSTITEMMAST.EOF Or RSTITEMMAST.BOF) Then
                    If MsgBox("Item Code exists for " & RSTITEMMAST!ITEM_NAME & " Do You want to add this item with a system generated Item Code?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        Exit Sub
                    Else
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        TXTPRODUCT.Tag = ""
'                        Set RSTITEMMAST = New ADODB.Recordset
'                        RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
'                        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'                            If IsNull(RSTITEMMAST.Fields(0)) Then
'                                TXTPRODUCT.Tag = 1
'                            Else
'                                TXTPRODUCT.Tag = Val(RSTITEMMAST.Fields(0)) + 1
'                            End If
'                        End If
'                        RSTITEMMAST.Close
'                        Set RSTITEMMAST = Nothing
                        
                        Set RSTITEMMAST = New ADODB.Recordset
                        RSTITEMMAST.Open "Select * From ITEMMAST WHERE ITEM_CODE= (SELECT MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) FROM ITEMMAST)", db, adOpenStatic, adLockOptimistic, adCmdText
                        TXTPRODUCT.Tag = RSTITEMMAST!ITEM_CODE + 1
'                        Set RSTITEMMAST = New ADODB.Recordset
'                        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & TXTPRODUCT.Tag & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        db.BeginTrans
                        RSTITEMMAST.AddNew
                        'RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
                        RSTITEMMAST!ITEM_CODE = Val(TXTPRODUCT.Tag)
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
                        RSTITEMMAST!UN_BILL = "N"
                        If PC_FLAG = "Y" Then
                            RSTITEMMAST!PRICE_CHANGE = "Y"
                        Else
                            RSTITEMMAST!PRICE_CHANGE = "N"
                        End If
                        RSTITEMMAST.Update
                        db.CommitTrans
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        TXTITEMCODE.text = TXTPRODUCT.Tag
                        txtcategory.text = TXTPRODUCT.Tag
                        Call TxtItemcode_KeyDown(13, 0)
                        Exit Sub
                    End If
                Else
                    If MsgBox("Are you sure you want to add this item with this Item Code?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        Exit Sub
                    Else
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        
                        db.BeginTrans
                        Set RSTITEMMAST = New ADODB.Recordset
                        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(txtcategory.text) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
                        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                            RSTITEMMAST.AddNew
                            RSTITEMMAST!ITEM_CODE = Trim(txtcategory.text)
                        End If
                        'RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
                        
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
                        RSTITEMMAST!UN_BILL = "N"
                        If PC_FLAG = "Y" Then
                            RSTITEMMAST!PRICE_CHANGE = "Y"
                        Else
                            RSTITEMMAST!PRICE_CHANGE = "N"
                        End If
                        RSTITEMMAST.Update
                        db.CommitTrans
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                        TXTITEMCODE.text = Trim(txtcategory.text)
                        Call TxtItemcode_KeyDown(13, 0)
                        Exit Sub
                    End If
                End If
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                'Call TxtItemcode_KeyDown(13, 0)
            End If
            Exit Sub
            If Trim(TXTPRODUCT.text) = "" Then Exit Sub
'            If Trim(TXTPRODUCT.Text) = "" Then
'                txtcategory.Enabled = True
'                txtcategory.SetFocus
'                Exit Sub
'            End If
            CmdDelete.Enabled = False
                
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
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
                'RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND TRX_TYPE = 'PI' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    'RSTRXFILE.MoveLast
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.text = 1
                    Los_Pack.text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
                    TxtWarranty.text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    
                    CmbWrnty.text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo ERRHAND
                    
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.text = IIf(IsNull(RSTRXFILE!REF_NO), "", RSTRXFILE!REF_NO)
                    TXTEXPIRY.text = IIf(IsDate(RSTRXFILE!EXP_DATE), Format(RSTRXFILE!EXP_DATE, "MM/YY"), "  /  ")
                    If (IsNull(RSTRXFILE!MRP)) Then
                        TXTRATE.text = ""
                    Else
                        TXTRATE.text = Format(Round(Val(RSTRXFILE!MRP), 2), ".000")
                    End If
                    If (IsNull(RSTRXFILE!MRP_BT)) Then
                        txtmrpbt.text = 100 * Val(TXTRATE.text) / 105
                    Else
                        txtmrpbt.text = Val(TXTRATE.text)
                    End If
                    If IsNull(RSTRXFILE!PTR) Then
                        TXTPTR.text = ""
                    Else
                        TXTPTR.text = Format(Round(Val(RSTRXFILE!PTR) * Val(Los_Pack.text), 3), ".000")
                        'TXTPTR.Text = Format(Round(Val(RSTRXFILE!PTR), 3), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_RETAIL) Then
                        TXTRETAIL.text = ""
                    Else
                        TXTRETAIL.text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_WS) Then
                        txtWS.text = ""
                    Else
                        txtWS.text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_VAN) Then
                        txtvanrate.text = ""
                    Else
                        txtvanrate.text = Format(Round(Val(RSTRXFILE!P_VAN) * Val(Los_Pack.text), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_CRTN) Then
                        txtcrtn.text = ""
                    Else
                        txtcrtn.text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_LWS) Then
                        TxtLWRate.text = ""
                    Else
                        TxtLWRate.text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!CRTN_PACK) Then
                        txtcrtnpack.text = ""
                    Else
                        txtcrtnpack.text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_PRICE) Then
                        txtprofit.text = ""
                    Else
                        txtprofit.text = Format(Round(Val(RSTRXFILE!SALES_PRICE), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.text = ""
                    Else
                        TxttaxMRP.text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                    End If
                    If IsNull(RSTRXFILE!EXDUTY) Then
                        TxtExDuty.text = ""
                    Else
                        TxtExDuty.text = Format(Val(RSTRXFILE!EXDUTY), ".00")
                    End If
                    If IsNull(RSTRXFILE!CSTPER) Then
                        TxtCSTper.text = ""
                    Else
                        TxtCSTper.text = Format(Val(RSTRXFILE!CSTPER), ".00")
                    End If
                    If IsNull(RSTRXFILE!TR_DISC) Then
                        TxtTrDisc.text = ""
                    Else
                        TxtTrDisc.text = Format(Val(RSTRXFILE!TR_DISC), ".00")
                    End If
                    If IsNull(RSTRXFILE!cess_amt) Then
                        txtCess.text = ""
                    Else
                        txtCess.text = Format(Val(RSTRXFILE!cess_amt), ".00")
                    End If
                    If IsNull(RSTRXFILE!CESS_PER) Then
                        TxtCessPer.text = ""
                    Else
                        TxtCessPer.text = Format(Val(RSTRXFILE!CESS_PER), ".00")
                    End If
                    Los_Pack.text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
                    TxtWarranty.text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                    If RSTRXFILE!COM_FLAG = "A" Then
                        TxtComAmt.text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                        OptComAmt.Value = True
                    Else
                        TxtComper.text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                        OptComper.Value = True
                    End If
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo ERRHAND
                
                    'TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                    If RSTRXFILE!check_flag = "M" Then
                        OPTTaxMRP.Value = True
                    ElseIf RSTRXFILE!check_flag = "V" Then
                        OPTVAT.Value = True
                    Else
                        OPTNET.Value = True
                    End If
                    
                    '==========
                    If MDIMAIN.lblPerPurchase.Caption = "Y" Then
                        Txtgrossamt.Tag = Val(TXTPTR.text) + (Val(TXTPTR.text) * Val(TxtExDuty.text) / 100)
                        Txtgrossamt.Tag = Val(Txtgrossamt.Tag) + (Val(TXTPTR.text) * Val(TxtCSTper.text) / 100)
                        If OPTVAT.Value = True Then
                           If optdiscper.Value = True Then
                                LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                            Else
                                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                            End If
                        Else
                            TxttaxMRP.text = 0
                            If optdiscper.Value = True Then
                                LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                            Else
                                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
                            End If
                        End If
                        Call TXTRETAIL_LostFocus
                        Call txtws_LostFocus
                        Call txtvanrate_LostFocus
                        TXTRETAIL.BackColor = vbWhite
                        txtWS.BackColor = vbWhite
                        txtvanrate.BackColor = vbWhite
                    End If
                    '==========
                Else
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = 1
                    Los_Pack.text = 1
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND
                    
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.text = ""
                    txtHSN.text = ""
                    TXTEXPIRY.text = "  /  "
                    TXTRATE.text = ""
                    txtmrpbt.text = ""
                    TXTPTR.text = ""
                    TxtNetrate.text = ""
                    TXTRETAIL.text = ""
                    txtWS.text = ""
                    txtvanrate.text = ""
                    txtcrtn.text = ""
                    TxtLWRate.text = ""
                    txtcrtnpack.text = ""
                    txtprofit.text = ""
                    'TxttaxMRP.text = ""
                    TxtExDuty.text = ""
                    TxtCSTper.text = ""
                    TxtTrDisc.text = ""
                    TxtCustDisc.text = ""
                    TxtPoints.text = ""
                    TxtCessPer.text = ""
                    txtCess.text = ""
                    Los_Pack.text = "1"
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND
                    OPTVAT.Value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                
                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    TXTITEMCODE.Enabled = False
                    Los_Pack.Enabled = True
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
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPRODUCT_LostFocus()
    Call CHANGEBOXCOLOR(TXTPRODUCT, False)
End Sub

Private Sub TXTPTR_GotFocus()
    Call CHANGEBOXCOLOR(TXTPTR, True)
    TXTPTR.SelStart = 0
    TXTPTR.SelLength = Len(TXTPTR.text)
    Call FILL_PREVIIOUSRATE
    
    TXTRATE.Enabled = True
    TXTPTR.Enabled = True
    TxtNetrate.Enabled = True
    TxtPoints.Enabled = True
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
    TxtCustDisc.Enabled = True
    TxtComper.Enabled = True
    TxtComAmt.Enabled = True
    cmdadd.Enabled = True
    Txtgrossamt.Enabled = True
    txtBatch.Enabled = True
    txtHSN.Enabled = True
    TxtWarranty.Enabled = True
    CmbWrnty.Enabled = True
    TXTEXPDATE.Enabled = True
    End Sub

Private Sub TXTPTR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTPTR.text) = 0 Then Exit Sub
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
        Case vbKeyEscape
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" And M_EDIT = True Then Exit Sub
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                Set GRDPRERATE.DataSource = Nothing
                fRMEPRERATE.Visible = False
                TXTPRODUCT.Enabled = True
                TXTPRODUCT.SetFocus
            Else
                TXTRATE.SetFocus
            End If
        Case 116
            Call FILL_PREVIIOUSRATE
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TXTPTR, False)
    'tXTptrdummy.Text = Format(Val(TXTPTR.Text) / Val(TXTUNIT.Text), ".000")
    Txtgrossamt.text = Val(TXTPTR.text) * Val(TXTQTY.text)
    TXTPTR.text = Format(TXTPTR.text, ".0000")
    'TxtNetrate.Text = Round(Val(TXTPTR.Text) + Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100, 4)
    Call TxttaxMRP_LostFocus
    Call TXTRETAIL_LostFocus
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            TxtRetailPercent.text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtWsalePercent.text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtSchPercent.text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtSchPercent_LostFocus
    End If
    
    'TXTRETAIL.Text = Round(Val(txtmrpbt.Text) * 0.8, 2)
'    txtretail.Text = Format(Round(Val(TXTRATE.Text) - (Val(txtmrpbt.Text) * 20 / 100), 3), ".000")
'    txtprofit.Text = Format(Round(Val(txtretail.Text) - Val(txtretail.Text) * 10 / 100, 3), ".000")
End Sub

Private Sub TXTQTY_GotFocus()
    Call CHANGEBOXCOLOR(TXTQTY, True)
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.text)
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
    If cmbfull.ListIndex = -1 Then cmbfull.text = "Nos"
    If Val(Los_Pack.text) = 1 Then CmbPack.text = cmbfull.text
    FRMEGRDTMP.Visible = False
    CmbPack.Enabled = True
    cmbfull.Enabled = True
    Cmbcategory.Enabled = True
    Los_Pack.Enabled = True
    TXTQTY.Enabled = True
    TxTfree.Enabled = True
    TXTRATE.Enabled = True
    TXTPTR.Enabled = True
    TxtNetrate.Enabled = True
    TxtPoints.Enabled = True
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
    TxtCustDisc.Enabled = True
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
    rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    With rststock
        If Not (.EOF And .BOF) Then
            txtHSN.text = IIf(IsNull(rststock!REMARKS), "", rststock!REMARKS)
            TxtCustDisc.text = IIf(IsNull(rststock!CUST_DISC), "", rststock!CUST_DISC)
            TxtPoints.text = IIf(IsNull(rststock!SCH_POINTS), "", rststock!SCH_POINTS)
            On Error Resume Next
            If cmbfull.ListIndex = -1 Then cmbfull.text = IIf(IsNull(rststock!FULL_PACK), 0, rststock!FULL_PACK)
            On Error GoTo ERRHAND
            If IsNull(rststock!FREE_WARN) Or rststock!FREE_WARN = "N" Or rststock!FREE_WARN = "" Then
                ChkFree.Value = 0
            Else
                ChkFree.Value = 1
            End If
        Else
            ChkFree.Value = 0
            txtHSN.text = ""
            TxtCustDisc.text = ""
            TxtPoints.text = ""
            On Error Resume Next
            If cmbfull.ListIndex = -1 Then cmbfull.text = CmbPack.text
            ChkFree.Value = 0
            On Error GoTo ERRHAND
        End If
    End With
    rststock.Close
    Set rststock = Nothing
    
    If Trim(TxtBarcode.text) = "" Then
        Set rststock = New ADODB.Recordset
        rststock.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
        If Not (rststock.EOF Or rststock.BOF) Then
            TxtBarcode.text = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
        End If
        rststock.Close
        Set rststock = Nothing
    End If
    
    FRMEQTY.Visible = False
    lblbarqty.Caption = ""
    lblavlqty.Caption = ""
    If M_EDIT = False Then
        Dim rstTRXMAST As ADODB.Recordset
        Set rstTRXMAST = New ADODB.Recordset
        rstTRXMAST.Open "SELECT CLOSE_QTY FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' ", db, adOpenStatic, adLockReadOnly
        If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
            lblavlqty.Caption = IIf(IsNull(rstTRXMAST!CLOSE_QTY), 0, rstTRXMAST!CLOSE_QTY)
        End If
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
        If Trim(TxtBarcode.text) <> "" Then
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "SELECT SUM(BAL_QTY) FROM RTRXFILE WHERE BARCODE= '" & Trim(TxtBarcode.text) & "'AND  BAL_QTY >0", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                lblbarqty.Caption = IIf(IsNull(rstTRXMAST.Fields(0)), 0, rstTRXMAST.Fields(0))
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
        End If
        FRMEQTY.Visible = True
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.text) = 0 Then Exit Sub
            TxTfree.Enabled = True
            TxTfree.SetFocus
        Case vbKeyEscape
            cmbfull.Enabled = True
            cmbfull.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TXTQTY, False)
    TXTQTY.text = Format(TXTQTY.text, ".00")
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.text) * Round(Val(TXTPTR.text), 2)), ".000")
    LblGross.Caption = Format((Val(TXTQTY.text) * Round(Val(TXTPTR.text), 2)), ".000")
    Call TXTPTR_LostFocus
End Sub

Private Sub TXTRATE_Change()
'    If Val(TXTRATE.Text) > 0 Then TXTRETAIL.Text = Val(TXTRATE.Text)

'    If val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTRATE.Text) > 0 Then
'        TXTRETAIL.Text = Val(TXTRATE.Text)
'    End If
'    If val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTRATE.Text) > 0 Then
'        txtWS.Text = Val(TXTRATE.Text)
'    End If
'    If val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTRATE.Text) > 0 Then
'        txtvanrate.Text = Val(TXTRATE.Text)
'    End If
End Sub

Private Sub TXTRATE_GotFocus()
    TXTRATE.Tag = Val(TXTRATE.text)
    Call CHANGEBOXCOLOR(TXTRATE, True)
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
            TxTfree.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    If Val(TXTRATE.text) <> 0 And (Val(TXTRETAIL.text) = 0 Or Val(TXTRATE.text) <> Val(TXTRATE.Tag)) Then
        TXTRETAIL.text = Val(TXTRATE.text)
        Call TxttaxMRP_LostFocus
        Call TXTRETAIL_LostFocus
    End If
    Call CHANGEBOXCOLOR(TXTRATE, False)
    TXTRATE.text = Format(TXTRATE.text, ".000")
    txtmrpbt.text = 100 * Val(TXTRATE.text) / 105 '(100 + Val(TxttaxMRP.Text))
End Sub

Private Sub txtremarks_GotFocus()
    Call CHANGEBOXCOLOR(TXTREMARKS, True)
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If txtBillNo.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Please select the supplier from the list", vbOKOnly, "EzBiz"
                TXTDEALER.SetFocus
                Exit Sub
            End If
            'If TXTINVOICE.Text = "" Then Exit Sub
            If Not IsDate(TXTINVDATE.text) Then Exit Sub
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND PINV = '" & Trim(TXTINVOICE.text) & "' AND VCH_NO <> " & Val(txtBillNo.text) & " AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                MsgBox "You have already entered this Invoice number for " & Trim(DataList2.text) & " as Computer Bill No. " & rstTRXMAST!VCH_NO, vbOKOnly, "EzBiz"
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                TXTINVOICE.SetFocus
                Exit Sub
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            CMBDISTRICT.SetFocus
        Case vbKeyEscape
            TXTINVDATE.SetFocus
    End Select
    Exit Sub
ERRHAND:
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

Private Sub TXTREMARKS_LostFocus()
    Call CHANGEBOXCOLOR(TXTREMARKS, False)
End Sub

Private Sub TxtRetailPercent_GotFocus()
    Call CHANGEBOXCOLOR(TxtRetailPercent, True)
    TxtRetailPercent.SelStart = 0
    TxtRetailPercent.SelLength = Len(TxtRetailPercent.text)
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.text) <> 0 Then Label1(38).Caption = "Disc% on MRP"
End Sub

Private Sub TxtRetailPercent_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            If MDIMAIN.lblPerPurchase.Caption = "Y" Then
                txtWsalePercent.SetFocus
            Else
                txtWS.SetFocus
            End If
         Case vbKeyEscape
            TXTRETAIL.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtRetailPercent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtRetailPercent_LostFocus()
    Call CHANGEBOXCOLOR(TxtRetailPercent, False)
    On Error Resume Next
'    If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) + ((Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) + ((Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    Else
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    End If
''    If Val(TXTRATE.Text) = 0 Then
''        txtretail.Text = Round((Val(TXTPTR.Tag) * Val(TxtRetailPercent.Text) / 100) + Val(TXTPTR.Tag), 0)
''    Else
''        'txtretail.Text = Round(Val(TXTRATE.Text) / 1.12, 2) - (Round(Val(TXTRATE.Text) / 1.12, 2) * Val(TxtRetailPercent.Text) / 100)
''        txtretail.Text = Round(Val(TXTRATE.Text) * 100 / (Val(TxtRetailPercent.Text) + 100), 0)
''    End If
    
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.text) <> 0 Then
        If Val(TxtRetailPercent.text) > 0 Then
            TXTRETAIL.text = Round(Val(TXTRATE.text) - (Val(TXTRATE.text) * Val(TxtRetailPercent.text) / 100), 2)
        Else
            TXTRETAIL.text = Val(TXTRATE.text)
        End If
        Call TXTRETAIL_LostFocus
    Else
        If MDIMAIN.lblgst.Caption <> "R" Then
            If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
                TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption)) + ((Val(TxtExpense.text)))), 4)
            Else
                TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
            End If
        Else
            If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
                TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.text)))), 4)
            Else
                TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
            End If
        End If
        If MDIMAIN.lblgst.Caption <> "R" Then
            TXTRETAIL.text = Round((Val(TXTPTR.Tag) * Val(TxtRetailPercent.text) / 100) + Val(TXTPTR.Tag), 2)
        Else
            If MDIMAIN.StatusBar.Panels(14).text = "Y" Then
                TXTRETAIL.text = (Val(TXTPTR.Tag) * Val(TxtRetailPercent.text) / 100) + Val(TXTPTR.Tag)
                TXTRETAIL.text = Round(Val(TXTRETAIL.text) + (Val(TXTRETAIL.text) * (Val(TxttaxMRP.text) + Val(TxtCessPer.text)) / 100), 2)
                'TXTRETAIL.Tag = Round(Val(TXTRETAIL.Text) * 100 / ((Val(TxttaxMRP.Text)) + 100), 2)
            Else
                TXTRETAIL.text = Round((Val(TXTPTR.Tag) * Val(TxtRetailPercent.text) / 100) + Val(TXTPTR.Tag), 2)
            End If
        End If
    End If
    
    TXTRETAIL.text = Format(Val(TXTRETAIL.text), "0.00")
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub
'
Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                  Case 41  'BARCODE COUNT
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Trim(TXTsample.text)
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
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
        Case 3, 5, 6, 7, 8, 9, 10
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 41
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub txtSchPercent_GotFocus()
    Call CHANGEBOXCOLOR(txtSchPercent, True)
    txtSchPercent.SelStart = 0
    txtSchPercent.SelLength = Len(txtSchPercent.text)
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.text) <> 0 Then Label1(38).Caption = "Disc% on MRP"
End Sub

Private Sub txtSchPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcrtnpack.SetFocus
         Case vbKeyEscape
            txtvanrate.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtSchPercent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSchPercent_LostFocus()
    Call CHANGEBOXCOLOR(txtSchPercent, False)
    On Error Resume Next
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.text) <> 0 Then
        If Val(txtSchPercent.text) > 0 Then
            txtvanrate.text = Round(Val(TXTRATE.text) - (Val(TXTRATE.text) * Val(txtSchPercent.text) / 100), 2)
        Else
            txtvanrate.text = Val(TXTRATE.text)
        End If
        Call txtvanrate_LostFocus
    Else
        If MDIMAIN.lblgst.Caption <> "R" Then
            If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
                TXTPTR.Tag = Val(LBLSUBTOTAL.Caption)
            Else
                TXTPTR.Tag = Round(Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.text) + Val(TxTfree.text), 4)
            End If
        Else
            If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
                TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.text)))), 4)
            Else
                TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
            End If
        End If
        If MDIMAIN.lblgst.Caption <> "R" Then
            txtvanrate.text = Round((Val(TXTPTR.Tag) * Val(txtSchPercent.text) / 100) + Val(TXTPTR.Tag), 2)
        Else
            If MDIMAIN.StatusBar.Panels(14).text = "Y" Then
                txtvanrate.text = (Val(TXTPTR.Tag) * Val(txtSchPercent.text) / 100) + Val(TXTPTR.Tag)
                txtvanrate.text = Round(Val(txtvanrate.text) + (Val(txtvanrate.text) * (Val(TxttaxMRP.text) + Val(TxtCessPer.text)) / 100), 2)
            Else
                txtvanrate.text = Round((Val(TXTPTR.Tag) * Val(txtSchPercent.text) / 100) + Val(TXTPTR.Tag), 2)
            End If
        End If
    End If
    
    txtvanrate.text = Format(Val(txtvanrate.text), "0.00")
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub TXTSLNO_GotFocus()
    FRMEQTY.Visible = False
    Call CHANGEBOXCOLOR(TXTSLNO, True)
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    TxtBarcode.Enabled = False
    txtcategory.Enabled = False
    TXTPRODUCT.Enabled = False
    
    Los_Pack.Enabled = False
    CmbPack.Enabled = False
    cmbfull.Enabled = False
    Cmbcategory.Enabled = False
    TXTQTY.Enabled = False
    TxTfree.Enabled = False
    TXTRATE.Enabled = False
    TXTPTR.Enabled = False
    TxtNetrate.Enabled = False
    TxtPoints.Enabled = False
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
    TxtCustDisc.Enabled = False
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
    
    
    BARCODE_FLAG = False
    Set grdtmp.DataSource = Nothing
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If ((frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4" And frmLogin.rs!Level <> "1") And NEW_BILL = False) Or (frmLogin.rs!Level <> "0" And lbloldbills.Caption = "Y") Then Exit Sub
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
                TXTPTR.text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 9)) * Val(grdsales.TextMatrix(Val(TXTSLNO.text), 5)), 4), "0.0000")
                txtprofit.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 7)), "0.00")
                TXTRETAIL.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 18)), "0.00")
                LBLPRE.Caption = Val(TXTRETAIL.text)
                txtWS.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 19)), "0.00")
                txtvanrate.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 25)), "0.00")
                Txtgrossamt.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 26)), "0.00")
                txtcrtn.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 20)), "0.00")
                TxtLWRate.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 37)), "0.00")
                txtcrtnpack.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 24)), "0.00")
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 23)) = "A" Then
                    OptComAmt.Value = True
                    TxtComper.text = ""
                    TxtComAmt.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 22)), "0.00")
                Else
                    OptComper.Value = True
                    TxtComper.text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.text), 21)), "0.00")
                    TxtComAmt.text = ""
                End If
                
                'TXTPTR.Text = Format((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))) * Val(Los_Pack.Text), "0.000")

                txtBatch.text = grdsales.TextMatrix(Val(TXTSLNO.text), 11)
                TXTEXPDATE.text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 12)), grdsales.TextMatrix(Val(TXTSLNO.text), 12), "  /  /    ")
                TXTEXPIRY.text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.text), 12)), Format(grdsales.TextMatrix(Val(TXTSLNO.text), 12), "mm/yy"), "  /  ")
                'LBLSUBTOTAL.Caption = Format(Val(TXTQTY.Text) * (Val(TXTPTR.Text) + Val(lbltaxamount.Caption)), ".000")
                If Optdiscamt.Value = True Then
                    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text) + Val(lbltaxamount.Caption) - Val(txtPD.text), ".000")
                    LblGross.Caption = Format(Val(Txtgrossamt.text) - Val(txtPD.text), ".000")
                Else
                    LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100), ".000")
                    LblGross.Caption = Format(Val(Txtgrossamt.text) - (Val(Val(Txtgrossamt.text) * Val(txtPD.text) / 100)), ".000")
                End If
                TxTfree.text = grdsales.TextMatrix(Val(TXTSLNO.text), 14)
                TxttaxMRP.text = grdsales.TextMatrix(Val(TXTSLNO.text), 10)
                txtmrpbt.text = 100 * Val(TXTRATE.text) / 105 '(100 + Val(TxttaxMRP.Text))
                txtPD.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 17))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) = "V" Then
                    OPTVAT.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 15)) = "M" Then
                    OPTTaxMRP.Value = True
                Else
                    OPTNET.Value = True
                End If
                
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 27)) = "P" Then
                    optdiscper.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 27)) = "A" Then
                    Optdiscamt.Value = True
                End If
                On Error Resume Next
                Los_Pack.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 28))
                CmbPack.text = grdsales.TextMatrix(Val(TXTSLNO.text), 29)
                TxtWarranty.text = grdsales.TextMatrix(Val(TXTSLNO.text), 30)
                CmbWrnty.text = grdsales.TextMatrix(Val(TXTSLNO.text), 31)
                TxtExpense.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 32))
                TxtExDuty.text = "" 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 33))
                TxtCSTper.text = "" 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 34))
                TxtTrDisc.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 35))
                LblGross.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 36))
                TxtBarcode.text = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 38))
                txtCess.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 39))
                TxtCessPer.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 40))
                TxtNetrate.text = Round(Val(TXTPTR.text) + Val(TXTPTR.text) * Val(TxttaxMRP.text) / 100, 4)
                FRMEGRDTMP.Visible = False
                err.Clear
                
                On Error GoTo ERRHAND
                Dim rststock As ADODB.Recordset
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                With rststock
                    If Not (.EOF And .BOF) Then
                        lblcategory.Caption = IIf(IsNull(rststock!Category), "", rststock!Category)
                        Cmbcategory.text = IIf(IsNull(rststock!Category), "", rststock!Category)
                        On Error Resume Next
                        cmbfull.text = IIf(IsNull(rststock!FULL_PACK), 0, rststock!FULL_PACK)
                        err.Clear
                        On Error GoTo ERRHAND
                    Else
'                        lblcategory.Caption = ""
'                        Cmbcategory.text = ""
                    End If
                End With
                rststock.Close
                Set rststock = Nothing
                
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
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
            'TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.text = Val(grdsales.rows)
                TXTPRODUCT.text = ""
                TXTITEMCODE.text = ""
                TxtBarcode.text = ""
                TXTQTY.text = ""
                Txtpack.text = 1 '""
                Los_Pack.text = ""
                CmbPack.ListIndex = -1
                TxtWarranty.text = ""
                CmbWrnty.ListIndex = -1
                TxTfree.text = ""
                'TxttaxMRP.text = ""
                TxtExDuty.text = ""
                TxtCSTper.text = ""
                TxtTrDisc.text = ""
                TxtCustDisc.text = ""
                TxtPoints.text = ""
                TxtCessPer.text = ""
                txtCess.text = ""
                txtPD.text = ""
                TxtExpense.text = ""
                txtprofit.text = ""
                TXTRETAIL.text = ""
                TxtRetailPercent.text = ""
                
                txtWsalePercent.text = ""
                txtSchPercent.text = ""
                txtWS.text = ""
                txtvanrate.text = ""
                Txtgrossamt.text = ""
                txtcrtn.text = ""
                TxtLWRate.text = ""
                txtcrtnpack.text = ""
                OptComper.Value = True
                TXTRATE.text = ""
                TxtComAmt.text = ""
                TxtComper.text = ""
                txtmrpbt.text = ""
                LBLSUBTOTAL.Caption = ""
                LblGross.Caption = ""
                lbltaxamount.Caption = ""
'                lblcategory.Caption = ""
'                Cmbcategory.text = ""
                TXTEXPDATE.text = "  /  /    "
                TXTEXPIRY.text = "  /  "
                txtBatch.text = ""
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            Else
                CMDTRANSFER.Enabled = True
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
    Exit Sub
ERRHAND:
    MsgBox err.Description
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
    'Call CHANGEBOXCOLOR(TXTEXPIRY, True)
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
            If Len(Trim(TXTEXPIRY.text)) = 1 Then GoTo SKIP
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
                MsgBox "Item Expired....", vbOKOnly, "EzBiz"
                TXTEXPDATE.text = "  /  /    "
                TXTEXPIRY.SelStart = 0
                TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
                TXTEXPIRY.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.text) < 60 Then
                MsgBox "Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days", vbOKOnly, "EzBiz"
                TXTEXPDATE.text = "  /  /    "
                TXTEXPIRY.SelStart = 0
                TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
                TXTEXPIRY.SetFocus
                Exit Sub
            End If
            
            If DateDiff("d", Date, TXTEXPDATE.text) < 180 Then
                If MsgBox("Expiry < " & Val(DateDiff("d", Date, TXTEXPDATE.text)) & " Days.. DO YOU WANT TO CONTINUE...", vbYesNo, "EzBiz") = vbNo Then
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
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            txtBatch.Enabled = True
            TXTEXPDATE.Enabled = False
            txtBatch.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    'Call CHANGEBOXCOLOR(TXTEXPIRY, False)
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

Private Sub TXTSLNO_LostFocus()
    Call CHANGEBOXCOLOR(TXTSLNO, False)
End Sub

Private Sub TxttaxMRP_GotFocus()
    Call CHANGEBOXCOLOR(TxttaxMRP, True)
    TxttaxMRP.SelStart = 0
    TxttaxMRP.SelLength = Len(TxttaxMRP.text)
End Sub

Private Sub TxttaxMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxttaxMRP.text) <> 0 And OPTNET.Value = True Then
                OPTVAT.Value = True
                OPTVAT.SetFocus
                Exit Sub
            End If
            TxtPoints.SetFocus
         Case vbKeyEscape
            TXTPTR.Enabled = True
            TXTPTR.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TxttaxMRP, False)
    txtmrpbt.text = 100 * Val(TXTRATE.text) / (100 + Val(TxttaxMRP.text))
    Txtgrossamt.text = Val(TXTPTR.text) * Val(TXTQTY.text)
    Txtgrossamt.Tag = Val(Txtgrossamt.text) + (Val(Txtgrossamt.text) * Val(TxtExDuty.text) / 100)
    Txtgrossamt.Tag = Val(Txtgrossamt.Tag) + (Val(Txtgrossamt.text) * Val(TxtCSTper.text) / 100)
    'Txtgrossamt.Tag = Val(Txtgrossamt.Tag) + Val(txtCess.Text)
'    If Val(TxttaxMRP.Text) = 0 Then
'
'        TxttaxMRP.Text = 0
'        lbltaxamount.Caption = 0
'        lbltaxamount.Caption = ""
'        If optdiscper.value = True Then
'            LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag)) - Val(Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)
'            LblGross.Caption = (Val(Txtgrossamt.Tag)) - Val(Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)
'        Else
'            LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) - Val(TxtPD.Text))
'            LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(TxtPD.Text))
'        End If
'    Else
'        If OPTTaxMRP.value = True Then
'            lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100
'            If optdiscper.value = True Then
'                LBLSUBTOTAL.Caption = (Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption)
'                LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - (Val(LBLSUBTOTAL.Caption) * Val(TxtPD.Text) / 100)
'            Else
'                LBLSUBTOTAL.Caption = (Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption)
'                LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - Val(TxtPD.Text)
'            End If
'            LblGross.Caption = LBLSUBTOTAL.Caption
        If OPTVAT.Value = True Then
           If optdiscper.Value = True Then
                lbltaxamount.Tag = (Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                lbltaxamount.Caption = Round((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.text) / 100)) * Val(TxttaxMRP.text) / 100, 2)
                'lbltaxamount.Caption = Round((Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                'LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)
                'LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - (Val(LBLSUBTOTAL.Caption) - Val(lbltaxamount.Caption)) * Val(TxtTrDisc.Text) / 100
                LBLSUBTOTAL.Caption = Round(((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.text) / 100))) + Val(lbltaxamount.Caption), 2)
                LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
            Else
                lbltaxamount.Tag = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                lbltaxamount.Caption = Round((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.text) / 100)) * Val(TxttaxMRP.text) / 100, 2)
                'LBLSUBTOTAL.Caption = Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption) - Val(TxtPD.Text)
                LBLSUBTOTAL.Caption = Round(((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.text) / 100))) + Val(lbltaxamount.Caption), 2)
                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
            End If
            LBLSUBTOTAL.Caption = LBLSUBTOTAL.Caption + (Val(LblGross.Caption) * Val(TxtCessPer.text) / 100)
            LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) + (Val(txtCess.text) * Val(TXTQTY.text))
        Else
            TxttaxMRP.text = 0
            If optdiscper.Value = True Then
                lbltaxamount.Tag = (Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                lbltaxamount.Caption = Round((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.text) / 100)) * Val(TxttaxMRP.text) / 100, 2)
                'lbltaxamount.Caption = Round((Val(Txtgrossamt.Tag) - (Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                'LBLSUBTOTAL.Caption = (Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Tag) * Val(TxtPD.Text) / 100)
                'LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) - (Val(LBLSUBTOTAL.Caption) - Val(lbltaxamount.Caption)) * Val(TxtTrDisc.Text) / 100
                LBLSUBTOTAL.Caption = Round(((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.text) / 100))) + Val(lbltaxamount.Caption), 2)
                LblGross.Caption = (Val(Txtgrossamt.Tag) - Val(Val(Txtgrossamt.Tag) * Val(txtPD.text) / 100))
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
            Else
                lbltaxamount.Tag = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                lbltaxamount.Caption = Round((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.text) / 100)) * Val(TxttaxMRP.text) / 100, 2)
                'LBLSUBTOTAL.Caption = Val(Txtgrossamt.Tag) + Val(lbltaxamount.Caption) - Val(TxtPD.Text)
                LBLSUBTOTAL.Caption = Round(((Val(lbltaxamount.Tag) - (Val(lbltaxamount.Tag) * Val(TxtTrDisc.text) / 100))) + Val(lbltaxamount.Caption), 2)
                LblGross.Caption = Val(Txtgrossamt.Tag) - Val(txtPD.text)
                LblGross.Caption = Val(LblGross.Caption) - (Val(LblGross.Caption) * Val(TxtTrDisc.text) / 100)
            End If
            LBLSUBTOTAL.Caption = LBLSUBTOTAL.Caption + (Val(LblGross.Caption) * Val(TxtCessPer.text) / 100)
            LBLSUBTOTAL.Caption = Val(LBLSUBTOTAL.Caption) + (Val(txtCess.text) * Val(TXTQTY.text))
        End If
'    End If
    'LBLSUBTOTAL.Caption = Round(Val(LBLSUBTOTAL.Caption) + Val(txtCess.Text), 2)
    'TxtNetrate.Text = Round(Val(TXTPTR.Text) + Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100, 4)
    LBLSUBTOTAL.Caption = Format(Round(LBLSUBTOTAL.Caption, 3), "0.00")
    LblGross.Caption = Format(LblGross.Caption, "0.00")
    TxttaxMRP.text = Format(TxttaxMRP.text, "0.00")
    lbltaxamount.Caption = Format(lbltaxamount.Caption, "0.00")
    If Val(TXTQTY.text) > 0 Then TxtNetrate.text = Val(LBLSUBTOTAL) / Val(TXTQTY.text)
    
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            TxtRetailPercent.text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtWsalePercent.text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtSchPercent.text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtSchPercent_LostFocus
    End If
    
End Sub

Private Sub TxtTotalexp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape
            Call find_small_number
    End Select
End Sub

Private Sub TxtTotalexp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtTrDisc_LostFocus()
    Call CHANGEBOXCOLOR(TxtTrDisc, False)
    Call TxttaxMRP_LostFocus
    'If Val(TXTQTY.Text) <> 0 Then TxtNetrate.Text = Val(LBLSUBTOTAL) / Val(TXTQTY.Text)
End Sub

Private Sub TXTUNIT_GotFocus()
    Call CHANGEBOXCOLOR(TXTUNIT, True)
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
            TxTfree.text = ""
            'TxttaxMRP.text = ""
            TxtExDuty.text = ""
            TxtCSTper.text = ""
            TxtTrDisc.text = ""
            TxtCustDisc.text = ""
            TxtPoints.text = ""
            TxtCessPer.text = ""
            txtCess.text = ""
            txtprofit.text = ""
            TXTRETAIL.text = ""
            TxtRetailPercent.text = ""
            
            txtWsalePercent.text = ""
            txtSchPercent.text = ""
            txtWS.text = ""
            txtvanrate.text = ""
            Txtgrossamt.text = ""
            txtcrtn.text = ""
            TxtLWRate.text = ""
            txtcrtnpack.text = ""
            txtPD.text = ""
            TxtExpense.text = ""
            txtBatch.text = ""
            TXTRATE.text = ""
            txtmrpbt.text = ""
            TXTPTR.text = ""
            TxtNetrate.text = ""
            Txtgrossamt.text = ""
            TXTEXPDATE.text = "  /  /    "
            TXTEXPIRY.text = "  /  "
            LBLSUBTOTAL.Caption = ""
            LblGross.Caption = ""
            lbltaxamount.Caption = ""
'            lblcategory.Caption = ""
'            Cmbcategory.text = ""
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
    Call CHANGEBOXCOLOR(TXTDISCAMOUNT, False)
    Dim DISC As Currency
    
    On Error GoTo ERRHAND
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
    If Roundflag = True Then
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
            Else
                LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 2), "0.00")
            End If
    ''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
ERRHAND:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    TXTDISCAMOUNT.SetFocus
End Sub

Private Sub TXTDISCAMOUNT_GotFocus()
    Call CHANGEBOXCOLOR(TXTDISCAMOUNT, True)
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
            'If txtcategory.Enabled = True Then txtcategory.SetFocus
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
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    
    'If OLD_BILL = False Then Call checklastbill
    Set RSTTRXFILE = New ADODB.Recordset
    If OLD_BILL = False And Val(txtBillNo.text) <> 1 Then
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE VCH_NO= (SELECT MAX(VCH_NO) FROM TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PI')", db, adOpenStatic, adLockOptimistic, adCmdText
        txtBillNo.text = RSTTRXFILE!VCH_NO + 1
        db.BeginTrans
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "PI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = txtBillNo.text
        RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    Else
        RSTTRXFILE.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "PI"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = txtBillNo.text
            RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
            RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        End If
    End If
    If Not IsDate(TXTDATE.text) Then TXTDATE.text = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!CREATE_DATE = Format(TXTDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    If IsDate(TXTRCVDATE.text) Then
        RSTTRXFILE!RCVD_DATE = Format(TXTRCVDATE.text, "DD/MM/YYYY")
    End If
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
    If OptDr.Value = True Then
        RSTTRXFILE!SLSM_CODE = "DR"
    Else
        RSTTRXFILE!SLSM_CODE = "CR"
    End If
    RSTTRXFILE!check_flag = "N"
    'If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
    RSTTRXFILE!POST_FLAG = "N"
    RSTTRXFILE!CFORM_NO = ""
    RSTTRXFILE!CFORM_DATE = Date
    RSTTRXFILE!REMARKS = Trim(TXTREMARKS.text)
    RSTTRXFILE!DISC_PERS = Val(txtcramt.text)
    RSTTRXFILE!CST_PER = Val(TxtCST.text)
    RSTTRXFILE!INS_PER = Val(TxtInsurance.text)
    RSTTRXFILE!LETTER_NO = 0
    RSTTRXFILE!LETTER_DATE = Date
    RSTTRXFILE!INV_MSGS = ""
    RSTTRXFILE!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!PINV = Trim(TXTINVOICE.text)
    RSTTRXFILE!TRX_GODOWN = Trim(CMBDISTRICT.text)
    RSTTRXFILE.Update
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    'db.Execute "delete From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete FROM CRDTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'CR' AND INV_TRX_TYPE = 'PI'"
    db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'PY' AND INV_TRX_TYPE = 'PI' "
    'db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'PI'"
    If grdsales.rows = 1 Then GoTo SKIP
                
    Dim BillNO As Long
    If OptCr.Value = True Then
        i = 0
        Set rstMaxNo = New ADODB.Recordset
        rstMaxNo.Open "Select MAX(CR_NO) From CRDTPYMT", db, adOpenStatic, adLockReadOnly
        If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
            i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        End If
        rstMaxNo.Close
        Set rstMaxNo = Nothing
        
        'If lblcredit.Caption = "1" Then
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
    '        If lblcredit.Caption = "0" Then
    '            RSTITEMMAST!CHECK_FLAG = "Y"
    '            RSTITEMMAST!BAL_AMT = 0
    '        Else
    '            RSTITEMMAST!CHECK_FLAG = "N"
    '            RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption) - RSTITEMMAST!RCPT_AMOUNT
    '        End If
            RSTITEMMAST!RCPT_AMOUNT = 0
            RSTITEMMAST!check_flag = "N"
            RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption)
            RSTITEMMAST!PINV = Trim(TXTINVOICE.text)
            RSTITEMMAST!ACT_CODE = DataList2.BoundText
            RSTITEMMAST!ACT_NAME = DataList2.text
            RSTITEMMAST!REMARKS = Left(Trim(TXTREMARKS.text), 50)
            RSTITEMMAST.Update
            db.CommitTrans
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        'End If
    Else
        i = 0
        Set rstMaxNo = New ADODB.Recordset
        rstMaxNo.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'PY' AND '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenForwardOnly
        If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
            i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        End If
        rstMaxNo.Close
        Set rstMaxNo = Nothing
        
        BillNO = 1
        Set rstMaxNo = New ADODB.Recordset
        rstMaxNo.Open "Select MAX(REC_NO) From DBTPYMT WHERE TRX_TYPE = 'PY' AND '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenForwardOnly
        If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
            BillNO = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        End If
        rstMaxNo.Close
        Set rstMaxNo = Nothing
    
    'If lblcredit.Caption = "1" Then
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.text) & " AND TRX_TYPE = 'PY' AND INV_TRX_TYPE = 'PI'", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!TRX_TYPE = "PY"
            RSTITEMMAST!INV_TRX_TYPE = "PI"
            RSTITEMMAST!INV_NO = Val(txtBillNo.text)
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        End If
        RSTITEMMAST!CR_NO = i
        RSTITEMMAST!REC_NO = BillNO
        RSTITEMMAST!RCPT_AMT = Val(LBLTOTAL.Caption)
        RSTITEMMAST!RCPT_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = DataList2.text
        RSTITEMMAST!INV_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTITEMMAST!INV_AMT = 0
        'RSTITEMMAST!INV_NO = 0
        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
        RSTITEMMAST!BANK_FLAG = "N"
        RSTITEMMAST!B_TRX_TYPE = Null
        'RSTITEMMAST!B_TRX_NO = Null
        RSTITEMMAST!B_BILL_TRX_TYPE = Null
        RSTITEMMAST!B_TRX_YEAR = Null
        RSTITEMMAST!BANK_CODE = Null
        RSTITEMMAST!BANK_NAME = ""
        RSTITEMMAST!C_TRX_TYPE = Null
        'RSTITEMMAST!C_REC_NO = Null
        RSTITEMMAST!C_INV_TRX_TYPE = Null
        RSTITEMMAST!C_INV_TYPE = Null
        'RSTITEMMAST!RCPT_AMOUNT = 0
        'RSTITEMMAST!CHECK_FLAG = "N"
        'RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption)
        'RSTITEMMAST!PINV = Trim(TXTINVOICE.Text)
        
        RSTITEMMAST.Update
        db.CommitTrans
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
    
    Call find_small_number
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(Trim(TXTINVDATE.text), "dd/mm/yyyy")
        If IsDate(TXTRCVDATE.text) Then
            RSTTRXFILE!RCVD_DATE = Format(TXTRCVDATE.text, "DD/MM/YYYY")
        Else
            RSTTRXFILE!RCVD_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        End If
        RSTTRXFILE!VCH_DESC = "Received From " & Left(DataList2.text, 85)
        RSTTRXFILE!PINV = Trim(TXTINVOICE.text)
        RSTTRXFILE!TRX_GODOWN = Trim(CMBDISTRICT.text)
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        If CMBPO.text <> "" Then
            RSTTRXFILE!PO_NO = IIf(CMBPO.text = "", Null, CMBPO.text)
        End If
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
'    If lblcredit.Caption = "0" Then
'        i = 0
'        Set rstMaxRec = New ADODB.Recordset
'        rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
'        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
'            i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
'        End If
'        rstMaxRec.Close
'        Set rstMaxRec = Nothing
'
'        Set RSTITEMMAST = New ADODB.Recordset
'        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'PI'", db, adOpenStatic, adLockOptimistic, adCmdText
'        db.BeginTrans
'        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'            RSTITEMMAST.AddNew
'            RSTITEMMAST!rec_no = i + 1
'            RSTITEMMAST!INV_TYPE = "PY"
'            RSTITEMMAST!INV_TRX_TYPE = "PI"
'            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'            RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
'        End If
'        If lblcredit.Caption = "0" Then
'            RSTITEMMAST!TRX_TYPE = "DR"
'        Else
'            RSTITEMMAST!TRX_TYPE = "CR"
'        End If
'        RSTITEMMAST!TRX_TYPE = "CR"
'        RSTITEMMAST!act_code = DataList2.BoundText
'        RSTITEMMAST!act_name = Trim(DataList2.Text)
'        RSTITEMMAST!AMOUNT = Val(LBLTOTAL.Caption)
'        RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
'        RSTITEMMAST!CHECK_FLAG = "P"
'        RSTITEMMAST.Update
'        db.CommitTrans
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
'    End If

SKIP:
    
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PI'", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    CMDTRANSFER.Enabled = False
    cmdRefresh.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.text = TXTLASTBILL.text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.text = "  /  /    "
    TXTRCVDATE.text = "  /  /    "
    TXTINVOICE.text = ""
    CMBDISTRICT.text = ""
    lbladdress.Caption = ""
    TXTREMARKS.text = ""
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TxtBarcode.text = ""
    TXTPRODUCT.text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.text = ""
    Txtpack.text = 1 '""
    Los_Pack.text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.text = ""
    CmbWrnty.ListIndex = -1
    TxTfree.text = ""
    TxttaxMRP.text = ""
    TxtExDuty.text = ""
    TxtCSTper.text = ""
    TxtTrDisc.text = ""
    TxtCustDisc.text = ""
    TxtPoints.text = ""
    TxtCessPer.text = ""
    txtCess.text = ""
    txtPD.text = ""
    TxtExpense.text = ""
    txtprofit.text = ""
    TXTRETAIL.text = ""
    TxtRetailPercent.text = ""
    
    txtWsalePercent.text = ""
    txtSchPercent.text = ""
    txtWS.text = ""
    txtvanrate.text = ""
    Txtgrossamt.text = ""
    txtcrtn.text = ""
    TxtLWRate.text = ""
    txtcrtnpack.text = ""
    txtBatch.text = ""
    txtHSN.text = ""
    TXTRATE.text = ""
    txtmrpbt.text = ""
    TXTPTR.text = ""
    TxtNetrate.text = ""
    Txtgrossamt.text = ""
    TXTEXPDATE.text = "  /  /    "
    TXTEXPIRY.text = "  /  "
    LBLSUBTOTAL.Caption = ""
    LblGross.Caption = ""
    lbltaxamount.Caption = ""
    lblcategory.Caption = ""
    Cmbcategory.text = ""
    LBLPRE.Caption = ""
    txtaddlamt.text = ""
    txtcramt.text = ""
    TxtInsurance.text = ""
    TxtCST.text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    LBLTOTALTAX.Caption = ""
    LBLGROSSAMT.Caption = ""
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    TXTDISCAMOUNT.text = ""
    TxtTotalexp.text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.text = ""
    lbldealer.Caption = ""
    grdsales.rows = 1
    CmdExit.Enabled = True
    OptComper.Value = True
    txtBillNo.SetFocus
    M_ADD = False
    M_EDIT = False
    OLD_BILL = False
    NEW_BILL = True
    lbloldbills.Caption = "N"
    LBLmonth.Caption = "0.00"
    Chkcancel.Value = 0
    Call CLEAR_COMBO
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "EzBiz"
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = 7 Then
        MsgBox "Select Supplier from the list", vbOKOnly, "EzBiz"
    Else
        Screen.MousePointer = vbNormal
        If err.Number <> -2147168237 Then
            MsgBox err.Description
        End If
        On Error Resume Next
        db.RollbackTrans
    End If
End Sub


Private Sub txtaddlamt_GotFocus()
    Call CHANGEBOXCOLOR(txtaddlamt, True)
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
            'If txtcategory.Enabled = True Then txtcategory.SetFocus
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
    Call CHANGEBOXCOLOR(txtaddlamt, False)
    Dim DISC As Currency
    
    On Error GoTo ERRHAND
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
    If Roundflag = True Then
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 0), "0.00")
    Else
        LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 2), "0.00")
    End If
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
ERRHAND:
    MsgBox "Please enter a Numeric Value for Discount", , "DISCOUNT.."
    txtaddlamt.SetFocus
End Sub

Private Sub txtcramt_GotFocus()
    Call CHANGEBOXCOLOR(txtcramt, True)
    txtcramt.SelStart = 0
    txtcramt.SelLength = Len(txtcramt.text)
End Sub

Private Sub txtcramt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("."), Asc("-")
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
            'If txtcategory.Enabled = True Then txtcategory.SetFocus
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
    Call CHANGEBOXCOLOR(txtcramt, False)
    Dim DISC As Currency
    
    On Error GoTo ERRHAND
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 2), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    Exit Sub
ERRHAND:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    txtcramt.SetFocus
End Sub

Private Sub OPTTaxMRP_GotFocus()
    OPTTaxMRP.BackColor = &H98F3C1
    'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text) + Val(TxtFree.Text)) * Val(TxttaxMRP.Text) / 100
    'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
    lbltaxamount.Caption = ((Val(TXTRATE.text) * (Val(TXTQTY.text) + Val(TxTfree.text)) * 55 / 100)) * Val(TxttaxMRP.text) / 100
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
    OPTVAT.BackColor = &H98F3C1
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
    OPTNET.BackColor = &H98F3C1
    lbltaxamount.Caption = ""
    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.text), ".000")
    LblGross.Caption = Format(Val(Txtgrossamt.text), ".000")
End Sub

Private Sub txtprofit_GotFocus()
    Call CHANGEBOXCOLOR(txtprofit, True)
    txtprofit.SelStart = 0
    txtprofit.SelLength = Len(txtprofit.text)
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
    Call CHANGEBOXCOLOR(txtprofit, False)
    txtprofit.text = Format(txtprofit.text, "0.00")
End Sub

Private Sub txtPD_GotFocus()
    Call CHANGEBOXCOLOR(txtPD, True)
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
            If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                Call CMDADD_Click
            Else
                TxtTrDisc.SetFocus
            End If
         Case vbKeyEscape
            TxtPoints.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(txtPD, False)
    Call TxttaxMRP_LostFocus
    
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            TxtRetailPercent.text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtWsalePercent.text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtSchPercent.text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtSchPercent_LostFocus
    End If
    'If Val(TXTQTY.Text) <> 0 Then TxtNetrate.Text = Val(LBLSUBTOTAL) / Val(TXTQTY.Text)
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
    On Error GoTo ERRHAND
    If OptCr.Value = True Then
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
    Else
        If flagchange.Caption <> "1" Then
            If ACT_FLAG = True Then
                ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                ACT_FLAG = False
            Else
                ACT_REC.Close
                ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTDEALER_GotFocus()
    Call CHANGEBOXCOLOR(TXTDEALER, True)
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
    If DataList2.BoundText = "" Then Call TXTDEALER_Change
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Trim(TXTDEALER.text) = "" Then Exit Sub
            If DataList2.VisibleCount = 0 Then
                If MsgBox("No such supplier exists. Do you want to create a new supplier", vbYesNo, "EzBiz") = vbNo Then
                    TXTDEALER.SetFocus
                    Exit Sub
                Else
                    frmsuppliermast.Show
                    frmsuppliermast.SetFocus
                    frmsuppliermast.txtsupplier.text = Trim(TXTDEALER.text)
                    Exit Sub
                End If
            End If
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
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
    Call fillcombo
    Call Monthly_purchase
    On Error GoTo ERRHAND
    Dim rstCustomer As ADODB.Recordset
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from ACTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address)) & Chr(13) & "GSTIN: " & IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
    Else
        lbladdress.Caption = ""
    End If
        
    'TXTDEALER.Text = lbldealer.Caption
    'LBL.Caption = ""
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "Purchase Bill..."
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
    'Call CHANGEBOXCOLOR(DataList2, True)
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub DataList2_LostFocus()
    'Call CHANGEBOXCOLOR(txtBillNo, False)
    DataList2.BackColor = vbWhite
    flagchange.Caption = ""
End Sub

Private Sub TXTRETAIL_GotFocus()
    If Val(TXTRETAIL.text) = 0 And Val(TXTRATE.text) <> 0 Then TXTRETAIL.text = Val(TXTRATE.text)
    If Val(TXTRETAIL.text) = 0 Then TXTRETAIL.text = ""
    Call CHANGEBOXCOLOR(TXTRETAIL, True)
    Call FILL_PREVIIOUSRATE
    TXTRETAIL.SelStart = 0
    TXTRETAIL.SelLength = Len(TXTRETAIL.text)
End Sub

Private Sub TXTRETAIL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTRETAIL.text) = 0 Then
                TxtRetailPercent.SetFocus
            Else
                txtWS.SetFocus
            End If
         Case vbKeyEscape
            txtPD.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TXTRETAIL, False)
    On Error Resume Next
    TXTRETAIL.text = Format(TXTRETAIL.text, "0.00")
    If Val(TXTPTR.text) = 0 Then
        TxtRetailPercent.text = ""
        Exit Sub
    End If
'    If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) + ((Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) + ((Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    Else
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    End If
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MDIMAIN.lblgst.Caption <> "R" Then
        If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption)) + ((Val(TxtExpense.text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
        End If
    Else
        If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
            TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
        End If
    End If
    If MDIMAIN.lblgst.Caption <> "R" Then
        TXTRETAIL.Tag = TXTRETAIL.text
    Else
        'TXTRETAIL.Tag = (Val(TXTRETAIL.Text) - Val(txtCess.Text)) / (1 + ((Val(TxttaxMRP.Text)) / 100) + (Val(TxtCessPer.Text) / 100))
        'TXTRETAIL.Tag = Round(Val(TXTRETAIL.Tag) + (Val(TXTRETAIL.Tag) * Val(TxttaxMRP.Text) / 100), 4)
        If MDIMAIN.StatusBar.Panels(14).text = "Y" Then
            'txtretail.Tag = Val(txtretail.Text) * 100 / ((Val(TxttaxMRP.Text) + Val(TxtCessPer.Text)) + 100)
            'txtretail.Tag = Val(txtretail.Tag) * 100 / ((Val(TxtCessPer.Text)) + 100)
            'txtretail.Tag = Round(Val(txtretail.Tag) - Val(txtCess.Text), 2)
            
            TXTRETAIL.Tag = (Val(TXTRETAIL.text) - Val(txtCess.text)) / (1 + ((Val(TxttaxMRP.text)) / 100) + (Val(TxtCessPer.text) / 100))
            txtCess.Tag = Round(Val(TXTRETAIL.Tag) + (Val(TXTRETAIL.Tag) * Val(TxttaxMRP.text) / 100), 4)
            TXTRETAIL.Tag = Round(Val(txtCess.Tag) * 100 / ((Val(TxttaxMRP.text)) + 100), 4)
            
'            TXTRETAILNOTAX.Text = (Val(txtNetrate.Text) - Val(TxtCessAmt.Text)) / (1 + ((Val(TXTTAX.Text) + Val(TxtKFC.Caption)) / 100) + (Val(TxtCessPer.Text) / 100))
'            txtretail.Text = Round(Val(TXTRETAILNOTAX.Text) + (Val(TXTRETAILNOTAX.Text) * Val(TXTTAX.Text) / 100), 4)
'            TXTRETAILNOTAX.Text = Round(Val(txtretail.Text) * 100 / ((Val(TXTTAX.Text)) + 100), 4)
        
        Else
            TXTRETAIL.Tag = TXTRETAIL.text
        End If
    End If
    If Val(Val(TXTPTR.Tag)) <> 0 Then
        TxtRetailPercent.text = Round(((Val(TXTRETAIL.Tag) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
        TxtRetailPercent.text = Format(Val(TxtRetailPercent.text), "0.00")
    Else
         TxtRetailPercent.text = Round(((Val(TXTRETAIL.Tag) - Val(TXTPTR.Tag)) * 100), 2)
        TxtRetailPercent.text = Format(Val(TxtRetailPercent.text), "0.00")
    End If
    
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub TxtWarranty_LostFocus()
    Call CHANGEBOXCOLOR(TxtWarranty, False)
End Sub

Private Sub txtws_GotFocus()
    Call CHANGEBOXCOLOR(txtWS, True)
    txtWS.SelStart = 0
    txtWS.SelLength = Len(txtWS.text)
End Sub

Private Sub txtws_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtWS.text) = 0 Then
                txtWsalePercent.SetFocus
            Else
                txtvanrate.SetFocus
            End If
         Case vbKeyEscape
            TXTRETAIL.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(txtWS, False)
    On Error Resume Next
    txtWS.text = Format(txtWS.text, "0.00")
'    If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) + ((Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) + ((Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    Else
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    End If
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MDIMAIN.lblgst.Caption <> "R" Then
        If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption)) + ((Val(TxtExpense.text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
        End If
    Else
        If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
            TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
        End If
    End If
    If MDIMAIN.lblgst.Caption <> "R" Then
        txtWS.Tag = txtWS.text
    Else
        'txtws.Tag = (Val(txtws.Text) - Val(txtCess.Text)) / (1 + ((Val(TxttaxMRP.Text)) / 100) + (Val(TxtCessPer.Text) / 100))
        'txtws.Tag = Round(Val(txtws.Tag) + (Val(txtws.Tag) * Val(TxttaxMRP.Text) / 100), 4)
        If MDIMAIN.StatusBar.Panels(14).text = "Y" Then
            'TxtWS.Tag = Round(Val(TxtWS.Text) * 100 / ((Val(TxttaxMRP.Text)) + 100), 2)
            txtWS.Tag = (Val(txtWS.text) - Val(txtCess.text)) / (1 + ((Val(TxttaxMRP.text)) / 100) + (Val(TxtCessPer.text) / 100))
            txtCess.Tag = Round(Val(txtWS.Tag) + (Val(txtWS.Tag) * Val(TxttaxMRP.text) / 100), 4)
            txtWS.Tag = Round(Val(txtCess.Tag) * 100 / ((Val(TxttaxMRP.text)) + 100), 4)
        Else
            txtWS.Tag = txtWS.text
        End If
    End If
    If Val(Val(TXTPTR.Tag)) <> 0 Then
        txtWsalePercent.text = Round(((Val(txtWS.Tag) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
        txtWsalePercent.text = Format(Val(txtWsalePercent.text), "0.00")
    Else
        txtWsalePercent.text = Round(((Val(txtWS.Tag) - Val(TXTPTR.Tag)) * 100), 2)
        txtWsalePercent.text = Format(Val(txtWsalePercent.text), "0.00")
    End If
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub txtcrtn_GotFocus()
    Call CHANGEBOXCOLOR(txtcrtn, True)
    If Val(txtcrtnpack.text) = 0 Then txtcrtnpack.text = "1"
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = "1"
    If Val(Los_Pack.text) = 1 Then
        txtcrtn.text = Format(Val(TXTRETAIL.text), "0.00")
        txtcrtnpack.text = "1"
    Else
        If Val(txtcrtn.text) = 0 Then
            If Val(txtcrtnpack.text) = 1 Then
                txtcrtn.text = Format(Round(Val(TXTRETAIL.text) / Val(Los_Pack.text), 2), "0.00")
            Else
                txtcrtn.text = Format(Round((Val(TXTRETAIL.text) / Val(Los_Pack.text)) * Val(txtcrtnpack.text), 2), "0.00")
            End If
        End If
    End If
    
    txtcrtn.SelStart = 0
    txtcrtn.SelLength = Len(txtcrtn.text)
End Sub

Private Sub txtcrtn_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtcrtn.text) <> 0 And Val(txtcrtnpack.text) = 0 Then
                MsgBox "Please enter the Pack Qty for Loose Qty", vbOKOnly, "EzBiz"
                txtcrtnpack.SetFocus
                Exit Sub
            End If
            If Val(Los_Pack.text) = 1 Then
                txtcrtn.text = Format(Val(TXTRETAIL.text), "0.00")
                txtcrtnpack.text = "1"
            End If
           TxtLWRate.SetFocus
         Case vbKeyEscape
            txtcrtnpack.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(txtcrtn, False)
    txtcrtn.text = Format(txtcrtn.text, "0.00")
End Sub

Private Sub TxtComper_GotFocus()
    Call CHANGEBOXCOLOR(TxtComper, True)
    TxtComper.SelStart = 0
    TxtComper.SelLength = Len(TxtComper.text)
    OptComper.Value = True
End Sub

Private Sub TxtComper_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn
            TxtCessPer.SetFocus
        Case vbKeyEscape
            TxtCustDisc.SetFocus
        Case vbKeyDown
'            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
'            If Val(TXTQTY.Text) = 0 Then Exit Sub
'            If Val(TXTPTR.Text) = 0 Then Exit Sub
'            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TxtComper, False)
    TxtComper.text = Format(TxtComper.text, "0.00")
End Sub

Private Sub TxtComAmt_GotFocus()
    Call CHANGEBOXCOLOR(TxtComAmt, True)
    TxtComAmt.SelStart = 0
    TxtComAmt.SelLength = Len(TxtComAmt.text)
    OptComAmt.Value = True
End Sub

Private Sub TxtComAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtCessPer.SetFocus
        Case vbKeyEscape
            TxtCustDisc.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TxtComAmt, False)
    TxtComAmt.text = Format(TxtComAmt.text, "0.00")
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
    Call CHANGEBOXCOLOR(TxtComper, True)
    TxtComper.text = ""
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
    cmbfull.BackColor = &H98F3C1
    TxtComAmt.text = ""
    TxtComAmt.Enabled = False
    TxtComper.Enabled = True
    TxtComper.SetFocus
End Sub

Private Sub txtcrtnpack_GotFocus()
    Call CHANGEBOXCOLOR(txtcrtnpack, True)
    If Val(Los_Pack.text) = 1 Then
        txtcrtn.text = Format(Val(TXTRETAIL.text), "0.00")
        txtcrtnpack.text = "1"
    End If
    txtcrtnpack.SelStart = 0
    txtcrtnpack.SelLength = Len(txtcrtnpack.text)
End Sub

Private Sub txtcrtnpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtcrtnpack.text) = 0 Then txtcrtnpack.text = "1"
            txtcrtn.SetFocus
         Case vbKeyEscape
            txtvanrate.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(txtcrtnpack, False)
    txtcrtnpack.text = Format(txtcrtnpack.text, "0.00")
End Sub

Private Sub txtvanrate_GotFocus()
    Call CHANGEBOXCOLOR(txtvanrate, True)
    txtvanrate.SelStart = 0
    txtvanrate.SelLength = Len(txtvanrate.text)
End Sub

Private Sub txtvanrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtvanrate.text) = 0 Then
                txtSchPercent.SetFocus
            Else
                txtcrtnpack.SetFocus
            End If
         Case vbKeyEscape
            txtWS.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(txtvanrate, False)
    On Error Resume Next
    txtvanrate.text = Format(txtvanrate.text, "0.00")
'    If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) + ((Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) + ((Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    Else
'        If optdiscper.value = True Then
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            'TXTPTR.Tag = Val(TXTPTR.Text) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Tag) * Val(TxtPD.Text) / 100)) '+ ((Val(txtPD.Tag) - (Val(txtPD.Tag) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100)
'        Else
'            'TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            'TXTPTR.Tag = Val(TXTPTR.Text) / Val(Los_Pack.Text) + (Val(TXTPTR.Text) / Val(Los_Pack.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
'            TxtPD.Tag = Round((Val(TXTPTR.Text) * Val(TXTQTY.Text)) / (Val(TXTQTY.Text) + Val(TxtFree.Text)), 3)
'            TXTPTR.Tag = (Val(TxtPD.Tag) - (Val(TxtPD.Text) / Val(TXTQTY.Text))) '+ ((Val(txtPD.Tag) - (Val(txtPD.Text) / Val(TXTQTY.Text))) * Val(TxttaxMRP.Text) / 100)
'        End If
'    End If
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MDIMAIN.lblgst.Caption <> "R" Then
        If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption)) + ((Val(TxtExpense.text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LBLSUBTOTAL.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
        End If
    Else
        If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
            TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.text)))), 4)
        Else
            TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
        End If
    End If
    If MDIMAIN.lblgst.Caption <> "R" Then
        txtvanrate.Tag = txtvanrate.text
    Else
        'txtvanrate.Tag = (Val(txtvanrate.Text) - Val(txtCess.Text)) / (1 + ((Val(TxttaxMRP.Text)) / 100) + (Val(TxtCessPer.Text) / 100))
        'txtvanrate.Tag = Round(Val(txtvanrate.Tag) + (Val(txtvanrate.Tag) * Val(TxttaxMRP.Text) / 100), 4)
        If MDIMAIN.StatusBar.Panels(14).text = "Y" Then
            'txtvanrate.Tag = Round(Val(txtvanrate.Text) * 100 / ((Val(TxttaxMRP.Text)) + 100), 2)
            txtvanrate.Tag = (Val(txtvanrate.text) - Val(txtCess.text)) / (1 + ((Val(TxttaxMRP.text)) / 100) + (Val(TxtCessPer.text) / 100))
            txtCess.Tag = Round(Val(txtvanrate.Tag) + (Val(txtvanrate.Tag) * Val(TxttaxMRP.text) / 100), 4)
            txtvanrate.Tag = Round(Val(txtCess.Tag) * 100 / ((Val(TxttaxMRP.text)) + 100), 4)
        Else
            txtvanrate.Tag = txtvanrate.text
        End If
    End If
    If Val(Val(TXTPTR.Tag)) <> 0 Then
        txtSchPercent.text = Round(((Val(txtvanrate.Tag) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
        txtSchPercent.text = Format(Val(txtSchPercent.text), "0.00")
    Else
        txtSchPercent.text = Round(((Val(txtvanrate.Tag) - Val(TXTPTR.Tag)) * 100), 2)
        txtSchPercent.text = Format(Val(txtSchPercent.text), "0.00")
    End If
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub Txtgrossamt_GotFocus()
    Call CHANGEBOXCOLOR(Txtgrossamt, True)
    Txtgrossamt.SelStart = 0
    Txtgrossamt.SelLength = Len(Txtgrossamt.text)
End Sub

Private Sub Txtgrossamt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(Txtgrossamt.Text) = 0 Then Exit Sub
            TxtPoints.SetFocus
        Case vbKeyEscape
            TxttaxMRP.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(Txtgrossamt, False)
    If Val(Txtgrossamt.text) <> 0 Then
        Txtgrossamt.text = Format(Txtgrossamt.text, ".000")
        If Val(TXTQTY.text) <> 0 Then
            TXTPTR.text = Format(Round(Val(Txtgrossamt.text) / Val(TXTQTY.text), 4), "0.0000")
        ElseIf Val(TXTPTR.text) <> 0 Then
            TXTQTY.text = Format(Round(Val(Txtgrossamt.text) / Val(TXTPTR.text), 4), "0.0000")
        End If
    End If
    Call TxttaxMRP_LostFocus
    
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            TxtRetailPercent.text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtWsalePercent.text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtSchPercent.text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtSchPercent_LostFocus
    End If
    
End Sub

Function FILL_PREVIIOUSRATE()
    Set GRDPRERATE.DataSource = Nothing
    
    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, ITEM_COST_PRICE, ITEM_NET_COST_PRICE, P_RETAIL, P_WS, VCH_NO, VCH_DATE, VCH_DESC  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND (TRX_TYPE = 'PI' OR TRX_TYPE = 'PW' OR TRX_TYPE = 'LP') ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, LOOSE_PACK, PACK_TYPE, ITEM_COST_PRICE, ITEM_NET_COST_PRICE, P_RETAIL, P_WS, VCH_NO, VCH_DATE, VCH_DESC  From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND (TRX_TYPE = 'PI' OR TRX_TYPE = 'PW' OR TRX_TYPE = 'LP') ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
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
    Call CHANGEBOXCOLOR(Los_Pack, True)
    Los_Pack.SelStart = 0
    Los_Pack.SelLength = Len(Los_Pack.text)
    FRMEGRDTMP.Visible = False
    CmbPack.Enabled = True
    cmbfull.Enabled = True
    Cmbcategory.Enabled = True
    TXTQTY.Enabled = True
    TxTfree.Enabled = True
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
    TxtCustDisc.Enabled = True
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
    rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    With rststock
        If Not (.EOF And .BOF) Then
            On Error Resume Next
            cmbfull.text = IIf(IsNull(rststock!FULL_PACK), 0, rststock!FULL_PACK)
            On Error GoTo ERRHAND
        Else
            On Error Resume Next
            cmbfull.text = CmbPack.text
            On Error GoTo ERRHAND
        End If
    End With
    rststock.Close
    Set rststock = Nothing
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Los_Pack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            cmbfull.SetFocus
         Case vbKeyEscape
'             If M_EDIT = True Then Exit Sub
'            'TXTUNIT.Text = ""
'            Los_Pack.Enabled = False
            Cmbcategory.Enabled = True
            Cmbcategory.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub Los_Pack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtItemcode_GotFocus()
    Call CHANGEBOXCOLOR(TXTITEMCODE, True)
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
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
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            Else
                PHY_CODE.Close
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.text & "' AND ucase(CATEGORY) <> 'SERVICES' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'OWN' ", db, adOpenStatic, adLockReadOnly
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
                'lblavlqty.Caption = IIf(IsNull(PHY_CODE!CLOSE_QTY), "", PHY_CODE!CLOSE_QTY)
                lblcategory.Caption = IIf(IsNull(PHY_CODE!Category), "", PHY_CODE!Category)
                Cmbcategory.text = IIf(IsNull(PHY_CODE!Category), "", PHY_CODE!Category)
                On Error Resume Next
                Set Image1.DataSource = PHY
                If IsNull(PHY!PHOTO) Then
                    Frame6.Visible = False
                    Set Image1.DataSource = Nothing
                    bytData = ""
                Else
                    If err.Number = 545 Then
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
                On Error GoTo ERRHAND
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                        Exit For
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                'RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                'RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' AND TRX_TYPE <> 'ST' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "' AND TRX_TYPE = 'PI' ORDER BY VCH_DATE DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    'RSTRXFILE.MoveLast
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.text = 1
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.text = IIf(IsNull(RSTRXFILE!REF_NO), "", RSTRXFILE!REF_NO)
                    TXTEXPIRY.text = IIf(IsDate(RSTRXFILE!EXP_DATE), Format(RSTRXFILE!EXP_DATE, "MM/YY"), "  /  ")
                    Los_Pack.text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
                    If (IsNull(RSTRXFILE!MRP)) Then
                        TXTRATE.text = ""
                    Else
                        TXTRATE.text = Format(Round(Val(RSTRXFILE!MRP), 2), ".000")
                    End If
                    If (IsNull(RSTRXFILE!MRP_BT)) Then
                        txtmrpbt.text = 100 * Val(TXTRATE.text) / 105
                    Else
                        txtmrpbt.text = Val(TXTRATE.text)
                    End If
                    If IsNull(RSTRXFILE!PTR) Then
                        TXTPTR.text = ""
                    Else
                        TXTPTR.text = Format(Round(Val(RSTRXFILE!PTR) * Val(Los_Pack.text), 3), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_RETAIL) Then
                        TXTRETAIL.text = ""
                    Else
                        TXTRETAIL.text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_WS) Then
                        txtWS.text = ""
                    Else
                        txtWS.text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_VAN) Then
                        txtvanrate.text = ""
                    Else
                        txtvanrate.text = Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_CRTN) Then
                        txtcrtn.text = ""
                    Else
                        txtcrtn.text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!P_LWS) Then
                        TxtLWRate.text = ""
                    Else
                        TxtLWRate.text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!CRTN_PACK) Then
                        txtcrtnpack.text = ""
                    Else
                        txtcrtnpack.text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_PRICE) Then
                        txtprofit.text = ""
                    Else
                        txtprofit.text = Format(Round(Val(RSTRXFILE!SALES_PRICE), 2), ".000")
                    End If
                    If IsNull(RSTRXFILE!SALES_TAX) Then
                        TxttaxMRP.text = ""
                    Else
                        TxttaxMRP.text = Format(Val(RSTRXFILE!SALES_TAX), ".00")
                    End If
                    If IsNull(RSTRXFILE!EXDUTY) Then
                        TxtExDuty.text = ""
                    Else
                        TxtExDuty.text = Format(Val(RSTRXFILE!EXDUTY), ".00")
                    End If
                    If IsNull(RSTRXFILE!CSTPER) Then
                        TxtCSTper.text = ""
                    Else
                        TxtCSTper.text = Format(Val(RSTRXFILE!CSTPER), ".00")
                    End If
                    If IsNull(RSTRXFILE!TR_DISC) Then
                        TxtTrDisc.text = ""
                    Else
                        TxtTrDisc.text = Format(Val(RSTRXFILE!TR_DISC), ".00")
                    End If
                    If IsNull(RSTRXFILE!cess_amt) Then
                        txtCess.text = ""
                    Else
                        txtCess.text = Format(Val(RSTRXFILE!cess_amt), ".00")
                    End If
                    If IsNull(RSTRXFILE!CESS_PER) Then
                        txtCess.text = ""
                    Else
                        txtCess.text = Format(Val(RSTRXFILE!CESS_PER), ".00")
                    End If
                    TxtWarranty.text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
                    If RSTRXFILE!COM_FLAG = "A" Then
                        TxtComAmt.text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                        OptComAmt.Value = True
                    Else
                        TxtComper.text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                        OptComper.Value = True
                    End If
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo ERRHAND
                    
                    'TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", Format(Val(RSTRXFILE!SALES_TAX), ".00"))
                    If RSTRXFILE!check_flag = "M" Then
                        OPTTaxMRP.Value = True
                    ElseIf RSTRXFILE!check_flag = "V" Then
                        OPTVAT.Value = True
                    Else
                        OPTNET.Value = True
                    End If
                Else
                    TXTUNIT.text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.text = 1
                    Los_Pack.text = 1
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND
                    
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.text = ""
                    txtHSN.text = ""
                    TXTEXPIRY.text = "  /  "
                    TXTRATE.text = ""
                    txtmrpbt.text = ""
                    TXTPTR.text = ""
                    TxtNetrate.text = ""
                    TXTRETAIL.text = ""
                    txtWS.text = ""
                    txtvanrate.text = ""
                    txtcrtn.text = ""
                    TxtLWRate.text = ""
                    txtcrtnpack.text = ""
                    txtprofit.text = ""
                    'TxttaxMRP.text = ""
                    TxtExDuty.text = ""
                    TxtCSTper.text = ""
                    TxtTrDisc.text = ""
                    TxtCustDisc.text = ""
                    TxtPoints.text = ""
                    TxtCessPer.text = ""
                    txtCess.text = ""
                    Los_Pack.text = "1"
                    TxtWarranty.text = ""
                    On Error Resume Next
                    CmbPack.text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND
                    OPTVAT.Value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                
                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                With RSTRXFILE
                    If Not (.EOF And .BOF) Then
                        If IsNull(RSTRXFILE!P_RETAIL) Then
                            TXTRETAIL.text = ""
                        Else
                            TXTRETAIL.text = Format(Round(Val(RSTRXFILE!P_RETAIL), 2), ".000")
                        End If
                        If IsNull(RSTRXFILE!P_WS) Then
                            txtWS.text = ""
                        Else
                            txtWS.text = Format(Round(Val(RSTRXFILE!P_WS), 2), ".000")
                        End If
                        If IsNull(RSTRXFILE!P_VAN) Then
                            txtvanrate.text = ""
                        Else
                            txtvanrate.text = Format(Round(Val(RSTRXFILE!P_VAN), 2), ".000")
                        End If
                        If RSTRXFILE!COM_FLAG = "A" Then
                            TxtComAmt.text = IIf(IsNull(RSTRXFILE!COM_AMT), 0, RSTRXFILE!COM_AMT)
                            OptComAmt.Value = True
                        Else
                            TxtComper.text = IIf(IsNull(RSTRXFILE!COM_PER), 0, RSTRXFILE!COM_PER)
                            OptComper.Value = True
                        End If
                        If IsNull(RSTRXFILE!P_CRTN) Then
                            txtcrtn.text = ""
                        Else
                            txtcrtn.text = Format(Round(Val(RSTRXFILE!P_CRTN), 2), ".000")
                        End If
                        If IsNull(RSTRXFILE!P_LWS) Then
                            TxtLWRate.text = ""
                        Else
                            TxtLWRate.text = Format(Round(Val(RSTRXFILE!P_LWS), 2), ".000")
                        End If
                        If IsNull(RSTRXFILE!CRTN_PACK) Then
                            txtcrtnpack.text = ""
                        Else
                            txtcrtnpack.text = Format(Round(Val(RSTRXFILE!CRTN_PACK), 2), ".000")
                        End If
                        On Error Resume Next
                        cmbfull.text = IIf(IsNull(RSTRXFILE!FULL_PACK), 0, RSTRXFILE!FULL_PACK)
                        CmbPack.text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                        On Error GoTo ERRHAND
                    End If
                End With
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                
                If PHY_CODE.RecordCount = 1 Then
                    If Trim(UCase(lblcategory.Caption)) = "SERVICE CHARGE" Then
                        TXTITEMCODE.Enabled = False
                        TXTPRODUCT.Enabled = False
                        txtcategory.Enabled = False
                        Los_Pack.text = 1
                        TXTQTY.text = 1
                        TxTfree.text = ""
                        TXTRATE.text = ""
                        TXTPTR.Enabled = True
                        TXTPTR.SetFocus
                        'TxtPack.Enabled = True
                        'TxtPack.SetFocus
                    Else
                        TXTITEMCODE.Enabled = False
                        TXTPRODUCT.Enabled = False
                        Los_Pack.Enabled = True
                        TXTQTY.Enabled = True
                        TXTQTY.SetFocus
                        'TxtPack.Enabled = True
                        'TxtPack.SetFocus
                    End If
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
ERRHAND:
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
    Call CHANGEBOXCOLOR(TxtCST, True)
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
            'If txtcategory.Enabled = True Then txtcategory.SetFocus
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
    Call CHANGEBOXCOLOR(TxtCST, False)
    Dim DISC As Currency
    
    On Error GoTo ERRHAND
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 2), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtCST.Text)), 0), ".00")
    Exit Sub
ERRHAND:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtCST.SetFocus
End Sub

Private Sub TxtInsurance_GotFocus()
    Call CHANGEBOXCOLOR(TxtInsurance, True)
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
            'If txtcategory.Enabled = True Then txtcategory.SetFocus
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
    Call CHANGEBOXCOLOR(TxtInsurance, False)
    Dim DISC As Currency
    
    On Error GoTo ERRHAND
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
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.text)) * Val(TxtCST.text) / 100) + Val(TxtInsurance.text) + Val(txtaddlamt.text) - (Val(TXTDISCAMOUNT.text) + Val(txtcramt.text)), 2), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtInsurance.Text)), 0), ".00")
    Exit Sub
ERRHAND:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    TxtInsurance.SetFocus
End Sub

Private Sub txtWsalePercent_GotFocus()
    Call CHANGEBOXCOLOR(txtWsalePercent, True)
    txtWsalePercent.SelStart = 0
    txtWsalePercent.SelLength = Len(txtWsalePercent.text)
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.text) <> 0 Then Label1(38).Caption = "Disc% on MRP"
End Sub

Private Sub txtWsalePercent_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            If MDIMAIN.lblPerPurchase.Caption = "Y" Then
                txtSchPercent.SetFocus
            Else
                txtvanrate.SetFocus
            End If
         Case vbKeyEscape
            txtWS.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub txtWsalePercent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtWsalePercent_LostFocus()
    Call CHANGEBOXCOLOR(txtWsalePercent, False)
    On Error Resume Next
    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
    'If MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
    If MRPDISC_FLAG = "Y" And Val(TXTRATE.text) <> 0 Then
        If Val(txtWsalePercent.text) > 0 Then
            txtWS.text = Round(Val(TXTRATE.text) - (Val(TXTRATE.text) * Val(txtWsalePercent.text) / 100), 2)
        Else
            txtWS.text = Val(TXTRATE.text)
        End If
        Call txtws_LostFocus
    Else
        If MDIMAIN.lblgst.Caption <> "R" Then
            If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
                TXTPTR.Tag = Val(LBLSUBTOTAL.Caption)
            Else
                TXTPTR.Tag = Round(Val(LBLSUBTOTAL.Caption) / Val(TXTQTY.text) + Val(TxTfree.text), 4)
            End If
        Else
            If Val(TXTQTY.text) + Val(TxTfree.text) = 0 Then
                TXTPTR.Tag = Round(((Val(LblGross.Caption)) + ((Val(TxtExpense.text)))), 4)
            Else
                TXTPTR.Tag = Round(((Val(LblGross.Caption) / ((Val(TXTQTY.text) + Val(TxTfree.text)))) + ((Val(TxtExpense.text) / (Val(TXTQTY.text) + Val(TxTfree.text))))), 4)
            End If
        End If
        If MDIMAIN.lblgst.Caption <> "R" Then
            txtWS.text = Round((Val(TXTPTR.Tag) * Val(txtWsalePercent.text) / 100) + Val(TXTPTR.Tag), 2)
        Else
            If MDIMAIN.StatusBar.Panels(14).text = "Y" Then
                txtWS.text = (Val(TXTPTR.Tag) * Val(txtWsalePercent.text) / 100) + Val(TXTPTR.Tag)
                txtWS.text = Round(Val(txtWS.text) + (Val(txtWS.text) * (Val(TxttaxMRP.text) + Val(TxtCessPer.text)) / 100), 2)
            Else
                txtWS.text = Round((Val(TXTPTR.Tag) * Val(txtWsalePercent.text) / 100) + Val(TXTPTR.Tag), 2)
            End If
        End If
    End If
    
    txtWS.text = Format(Val(txtWS.text), "0.00")
    Label1(38).Caption = "% of   Profit"
End Sub

Private Sub TxtWarranty_GotFocus()
    Call CHANGEBOXCOLOR(TxtWarranty, True)
    TxtWarranty.SelStart = 0
    TxtWarranty.SelLength = Len(TxtWarranty.text)
End Sub

Private Sub TxtWarranty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtWarranty.text) = 0 Then
                cmdadd.SetFocus
            Else
                CmbWrnty.SetFocus
            End If
         Case vbKeyEscape
            txtCess.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
            If Val(TxtWarranty.text) <> 0 And CmbWrnty.ListIndex = -1 Then
                MsgBox "Please select the Warranty Period", , "EzBiz"
                CmbWrnty.SetFocus
                Exit Sub
            End If
            If Val(TxtWarranty.text) = 0 Then CmbWrnty.ListIndex = -1
            cmdadd.SetFocus
         Case vbKeyEscape
            TxtWarranty.SetFocus
    End Select
End Sub

Private Function checklastbill()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ERRHAND
    
    Dim BillNO As Double
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'PI'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        BillNO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    If Val(txtBillNo.text) >= BillNO Then
        txtBillNo.text = BillNO
    End If
Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Function Monthly_purchase()
    Dim rstTRANX As ADODB.Recordset
    Dim TOT_SALE As Long
    Dim FROM_DATE As Date
    
    FROM_DATE = "01/" & Month(Date) & "/" & Year(Date)
    On Error GoTo ERRHAND
    TOT_SALE = 0
    LBLmonth.Caption = "0.00"
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(NET_AMOUNT) FROM TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(Date, "yyyy/mm/dd") & "' AND ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        TOT_SALE = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    LBLmonth.Caption = Format(TOT_SALE, "0.00")
    
    
    'LBLRETURNED.Caption = Format(TOT_RET, "0.00")
    
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub TxtExpense_GotFocus()
    Call CHANGEBOXCOLOR(TxtExpense, True)
    TxtExpense.SelStart = 0
    TxtExpense.SelLength = Len(TxtExpense.text)
End Sub

Private Sub TxtExpense_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If MDIMAIN.lblPerPurchase.Caption = "Y" Then
                TxtRetailPercent.SetFocus
            Else
                TXTRETAIL.SetFocus
            End If
         Case vbKeyEscape
            TxtTrDisc.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TxtExDuty, True)
    TxtExDuty.SelStart = 0
    TxtExDuty.SelLength = Len(TxtExDuty.text)
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
    Call CHANGEBOXCOLOR(TxtCSTper, True)
    TxtCSTper.SelStart = 0
    TxtCSTper.SelLength = Len(TxtCSTper.text)
End Sub

Private Sub TxtCSTper_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTRETAIL.SetFocus
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
    Call CHANGEBOXCOLOR(TxtTrDisc, True)
    TxtTrDisc.SelStart = 0
    TxtTrDisc.SelLength = Len(TxtTrDisc.text)
End Sub

Private Sub TxtTrDisc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtExpense.SetFocus
        Case vbKeyEscape
            txtPD.SetFocus
'            Frame1.Enabled = True
'            If OptComper.value = True Then
'                TxtComper.SetFocus
'            Else
'                TxtComAmt.SetFocus
'            End If
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TxtLWRate, True)
    On Error Resume Next
    If Val(txtcrtnpack.text) = 0 Then txtcrtnpack.text = "1"
    If Val(Los_Pack.text) = 1 Then
        TxtLWRate.text = Format(Val(txtWS.text), "0.00")
        txtcrtnpack.text = "1"
    Else
        If Val(TxtLWRate.text) = 0 Then
            If Val(txtcrtnpack.text) = 1 Then
                TxtLWRate.text = Format(Round(Val(txtWS.text) / Val(Los_Pack.text), 2), "0.00")
            Else
                TxtLWRate.text = Format(Round((Val(txtWS.text) / Val(Los_Pack.text)) * Val(txtcrtnpack.text), 2), "0.00")
            End If
        End If
    End If
    
    TxtLWRate.SelStart = 0
    TxtLWRate.SelLength = Len(TxtLWRate.text)
End Sub

Private Sub TxtLWRate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtLWRate.text) <> 0 And Val(txtcrtnpack.text) = 0 Then
                MsgBox "Please enter the Pack Qty for Loose Qty", vbOKOnly, "EzBiz"
                txtcrtnpack.SetFocus
                Exit Sub
            End If
            If Val(Los_Pack.text) = 1 Then
                TxtLWRate.text = Format(Val(txtWS.text), "0.00")
                txtcrtnpack.text = "1"
            End If
            TxtCustDisc.SetFocus
         Case vbKeyEscape
            txtcrtn.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TxtLWRate, False)
    TxtLWRate.text = Format(TxtLWRate.text, "0.00")
End Sub

Private Function fillcombo()
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    Set CMBPO.DataSource = Nothing
    If PO_FLAG = True Then
        ACT_PO.Open "Select VCH_NO, ACT_CODE from POMAST  WHERE (VCH_NO = " & Val(PONO) & " OR (ISNULL(STATUS) OR STATUS = 'N')) AND ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_DATE ASC", db, adOpenStatic, adLockReadOnly, adCmdText
        PO_FLAG = False
    Else
        ACT_PO.Close
        ACT_PO.Open "Select VCH_NO, ACT_CODE from POMAST  WHERE (VCH_NO = " & Val(PONO) & " OR (ISNULL(STATUS) OR STATUS = 'N')) AND ACT_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_DATE ASC", db, adOpenStatic, adLockReadOnly, adCmdText
        PO_FLAG = False
    End If
    
    Set Me.CMBPO.RowSource = ACT_PO
    CMBPO.ListField = "VCH_NO"
    CMBPO.BoundColumn = "ACT_CODE"
    Screen.MousePointer = vbNormal
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TxtBarcode.text) = "" Then
                txtcategory.Enabled = True
                TXTPRODUCT.Enabled = True
                TXTPRODUCT.SetFocus
                Exit Sub
            End If
            
            Set rstTRXMAST = New ADODB.Recordset
            'MFG_REC.Open "SELECT DISTINCT CATEGORY FROM ITEMMAST RIGHT JOIN RTRXFILE ON ITEMMAST.ITEM_CODE = RTRXFILE.ITEM_CODE WHERE RTRXFILE.BAL_QTY > 0 ORDER BY ITEMMAST.MANUFACTURER", db, adOpenForwardOnly ' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y')
            'rstTRXMAST.Open "Select * From RTRXFILE LEFT JOIN ON ITEMMAST.ITEM_CODE = RTRXFILE.ITEM_CODE WHERE RTRXFILE.BARCODE= '" & Trim(txtBarcode.Text) & "' AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') ORDER BY VCH_NO ", db, adOpenStatic, adLockReadOnly
            'WHERE RTRXFILE.BARCODE= '" & Trim(txtBarcode.Text) & "' AND ITEMMAST.UN_BILL <> 'Y' ORDER BY VCH_NO
            rstTRXMAST.Open "Select * From RTRXFILE LEFT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE RTRXFILE.BARCODE= '" & Trim(TxtBarcode.text) & "' AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                CHANGE_FLAG = True
                TXTITEMCODE.text = IIf(IsNull(rstTRXMAST!ITEM_CODE), "", rstTRXMAST!ITEM_CODE)
                TXTPRODUCT.text = IIf(IsNull(rstTRXMAST!ITEM_NAME), "", rstTRXMAST!ITEM_NAME)
                CHANGE_FLAG = False
                TXTUNIT.text = 1 'IIf(IsNull(rstTRXMAST!UNIT), "", rstTRXMAST!UNIT)
                
                Set RSTITEM = New ADODB.Recordset
                RSTITEM.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & rstTRXMAST!TRX_YEAR & "' AND TRX_TYPE= '" & rstTRXMAST!TRX_TYPE & "' AND VCH_NO = " & rstTRXMAST!VCH_NO & " AND LINE_NO = " & rstTRXMAST!LINE_NO & "", db, adOpenStatic, adLockReadOnly
                If Not (RSTITEM.EOF Or RSTITEM.BOF) Then
                    Txtpack.text = IIf(IsNull(RSTITEM!LINE_DISC), "", RSTITEM!LINE_DISC)
                    Txtpack.text = 1
                    Los_Pack.text = IIf(IsNull(RSTITEM!LOOSE_PACK), "1", RSTITEM!LOOSE_PACK)
                    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
                    TxtWarranty.text = IIf(IsNull(RSTITEM!WARRANTY), "", RSTITEM!WARRANTY)
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTITEM!PACK_TYPE), "Nos", RSTITEM!PACK_TYPE)
                    CmbWrnty.text = IIf(IsNull(RSTITEM!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTITEM!WARRANTY_TYPE)
                    'cmbcolor.Text = IIf(IsNull(rstITEM!ITEM_COLOR), CmbWrnty.ListIndex = -1, rstITEM!ITEM_COLOR)
                    On Error GoTo ERRHAND
                    'Txtsize.Text = IIf(IsNull(rstITEM!ITEM_SIZE), "", rstITEM!ITEM_SIZE)
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(rstITEM!EXP_DATE), "  /  /    ", Format(rstITEM!EXP_DATE, "DD/MM/YYYY"))
                    txtBatch.text = IIf(IsNull(RSTITEM!REF_NO), "", RSTITEM!REF_NO)
                    TXTEXPIRY.text = IIf(IsDate(RSTITEM!EXP_DATE), Format(RSTITEM!EXP_DATE, "MM/YY"), "  /  ")
                    TXTRATE.text = IIf(IsNull(RSTITEM!MRP), "", Format(Round(Val(RSTITEM!MRP), 2), ".000"))
                    If (IsNull(RSTITEM!MRP_BT)) Then
                        txtmrpbt.text = 100 * Val(TXTRATE.text) / 105
                    Else
                        txtmrpbt.text = Val(TXTRATE.text)
                    End If
                    If IsNull(RSTITEM!PTR) Then
                        TXTPTR.text = ""
                    Else
                        TXTPTR.text = Format(Round(Val(RSTITEM!PTR) * Val(Los_Pack.text), 2), ".000")
                    End If
                    If IsNull(RSTITEM!P_RETAIL) Then
                        TXTRETAIL.text = ""
                    Else
                        TXTRETAIL.text = Format(Round(Val(RSTITEM!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(RSTITEM!P_WS) Then
                        txtWS.text = ""
                    Else
                        txtWS.text = Format(Round(Val(RSTITEM!P_WS), 2), ".000")
                    End If
                    If IsNull(RSTITEM!P_VAN) Then
                        txtvanrate.text = ""
                    Else
                        txtvanrate.text = Format(Round(Val(RSTITEM!P_VAN), 2), ".000")
                    End If
                    If IsNull(RSTITEM!P_CRTN) Then
                        txtcrtn.text = ""
                    Else
                        txtcrtn.text = Format(Round(Val(RSTITEM!P_CRTN), 2), ".000")
                    End If
                    If IsNull(RSTITEM!CRTN_PACK) Then
                        txtcrtnpack.text = ""
                    Else
                        txtcrtnpack.text = Format(Round(Val(RSTITEM!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(RSTITEM!SALES_PRICE) Then
                        txtprofit.text = ""
                    Else
                        txtprofit.text = Format(Round(Val(RSTITEM!SALES_PRICE), 2), ".000")
                    End If
                    If IsNull(RSTITEM!SALES_TAX) Then
                        TxttaxMRP.text = ""
                    Else
                        TxttaxMRP.text = Format(Val(RSTITEM!SALES_TAX), ".00")
                    End If
                    Los_Pack.text = IIf(IsNull(RSTITEM!LOOSE_PACK), "1", RSTITEM!LOOSE_PACK)
                    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
                    TxtWarranty.text = IIf(IsNull(RSTITEM!WARRANTY), "", RSTITEM!WARRANTY)
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(RSTITEM!PACK_TYPE), "Nos", RSTITEM!PACK_TYPE)
                    CmbWrnty.text = IIf(IsNull(RSTITEM!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTITEM!WARRANTY_TYPE)
                    On Error GoTo ERRHAND
                    txtPD.text = IIf(IsNull(RSTITEM!P_DISC), "", RSTITEM!P_DISC)
                    Select Case RSTITEM!DISC_FLAG
                        Case "P"
                            optdiscper.Value = True
                        Case "A"
                            Optdiscamt.Value = True
                    End Select
                    'TxttaxMRP.Text = IIf(IsNull(rstITEM!SALES_TAX), "", Format(Val(rstITEM!SALES_TAX), ".00"))
                    If RSTITEM!check_flag = "M" Then
                        OPTTaxMRP.Value = True
                    ElseIf RSTITEM!check_flag = "V" Then
                        OPTVAT.Value = True
                    Else
                        OPTNET.Value = True
                    End If
                End If
                RSTITEM.Close
                Set RSTITEM = Nothing
                
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                'txtbarcode.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
            Else
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                
                Set rstTRXMAST = New ADODB.Recordset
                rstTRXMAST.Open "Select * From ITEMMAST WHERE BARCODE= '" & Trim(TxtBarcode.text) & "' AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
                If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                    CHANGE_FLAG = True
                    TXTITEMCODE.text = IIf(IsNull(rstTRXMAST!ITEM_CODE), "", rstTRXMAST!ITEM_CODE)
                    TXTPRODUCT.text = IIf(IsNull(rstTRXMAST!ITEM_NAME), "", rstTRXMAST!ITEM_NAME)
                    CHANGE_FLAG = False
                    TXTUNIT.text = 1 'IIf(IsNull(rstTRXMAST!UNIT), "", rstTRXMAST!UNIT)
                    Txtpack.text = 1
                    Los_Pack.text = IIf(IsNull(rstTRXMAST!LOOSE_PACK), "1", rstTRXMAST!LOOSE_PACK)
                    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
                    TxtWarranty.text = IIf(IsNull(rstTRXMAST!WARRANTY), "", rstTRXMAST!WARRANTY)
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)
                    CmbWrnty.text = IIf(IsNull(rstTRXMAST!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, rstTRXMAST!WARRANTY_TYPE)
                    'cmbcolor.Text = IIf(IsNull(rstTRXMAST!ITEM_COLOR), CmbWrnty.ListIndex = -1, rstTRXMAST!ITEM_COLOR)
                    On Error GoTo ERRHAND
                    'Txtsize.Text = IIf(IsNull(rstTRXMAST!ITEM_SIZE), "", rstTRXMAST!ITEM_SIZE)
                    TXTEXPDATE.text = "  /  /    " 'IIf(IsNull(rstTRXMAST!EXP_DATE), "  /  /    ", Format(rstTRXMAST!EXP_DATE, "DD/MM/YYYY"))
                    'txtBatch.Text = IIf(IsNull(rstTRXMAST!REF_NO), "", rstTRXMAST!REF_NO)
                    'TXTEXPIRY.Text = IIf(IsDate(rstTRXMAST!EXP_DATE), Format(rstTRXMAST!EXP_DATE, "MM/YY"), "  /  ")
                    TXTRATE.text = IIf(IsNull(rstTRXMAST!MRP), "", Format(Round(Val(rstTRXMAST!MRP), 2), ".000"))
                    If IsNull(rstTRXMAST!ITEM_COST) Then
                        TXTPTR.text = ""
                    Else
                        TXTPTR.text = Format(Round(Val(rstTRXMAST!ITEM_COST) * Val(Los_Pack.text), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!P_RETAIL) Then
                        TXTRETAIL.text = ""
                    Else
                        TXTRETAIL.text = Format(Round(Val(rstTRXMAST!P_RETAIL), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!P_WS) Then
                        txtWS.text = ""
                    Else
                        txtWS.text = Format(Round(Val(rstTRXMAST!P_WS), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!P_VAN) Then
                        txtvanrate.text = ""
                    Else
                        txtvanrate.text = Format(Round(Val(rstTRXMAST!P_VAN), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!P_CRTN) Then
                        txtcrtn.text = ""
                    Else
                        txtcrtn.text = Format(Round(Val(rstTRXMAST!P_CRTN), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!CRTN_PACK) Then
                        txtcrtnpack.text = ""
                    Else
                        txtcrtnpack.text = Format(Round(Val(rstTRXMAST!CRTN_PACK), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!SALES_PRICE) Then
                        txtprofit.text = ""
                    Else
                        txtprofit.text = Format(Round(Val(rstTRXMAST!SALES_PRICE), 2), ".000")
                    End If
                    If IsNull(rstTRXMAST!SALES_TAX) Then
                        TxttaxMRP.text = ""
                    Else
                        TxttaxMRP.text = Format(Val(rstTRXMAST!SALES_TAX), ".00")
                    End If
                    Los_Pack.text = IIf(IsNull(rstTRXMAST!LOOSE_PACK), "1", rstTRXMAST!LOOSE_PACK)
                    If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
                    TxtWarranty.text = IIf(IsNull(rstTRXMAST!WARRANTY), "", rstTRXMAST!WARRANTY)
                    On Error Resume Next
                    CmbPack.text = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)
                    CmbWrnty.text = IIf(IsNull(rstTRXMAST!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, rstTRXMAST!WARRANTY_TYPE)
                    On Error GoTo ERRHAND
                    txtPD.text = IIf(IsNull(rstTRXMAST!CUST_DISC), "", rstTRXMAST!CUST_DISC)
                    optdiscper.Value = True
                    OPTVAT.Value = True
                    'TxttaxMRP.Text = IIf(IsNull(rstTRXMAST!SALES_TAX), "", Format(Val(rstTRXMAST!SALES_TAX), ".00"))
                    
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
                    TXTPRODUCT.Enabled = True
                    TXTPRODUCT.SetFocus
                End If
            End If
            If Trim(TxtBarcode.text) = "" Then
                BARCODE_FLAG = False
            Else
                BARCODE_FLAG = True
            End If
            
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
        Case vbKeyMultiply, vbKeyInsert
            If grdsales.rows <= 1 Then Exit Sub
            If Trim(grdsales.TextMatrix(grdsales.rows - 1, 38)) = "" Then
                TXTITEMCODE.text = grdsales.TextMatrix(grdsales.rows - 1, 1)
                TXTPRODUCT.text = grdsales.TextMatrix(grdsales.rows - 1, 2)
                Call TXTPRODUCT_KeyDown(13, 0)
            Else
                TxtBarcode.text = Trim(grdsales.TextMatrix(grdsales.rows - 1, 38))
            End If
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCess_GotFocus()
    Call CHANGEBOXCOLOR(txtCess, True)
    txtCess.SelStart = 0
    txtCess.SelLength = Len(txtCess.text)
End Sub

Private Sub txtCess_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.SetFocus
        Case vbKeyEscape
            TxtCessPer.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(txtCess, False)
    Call TxttaxMRP_LostFocus
End Sub

Private Function print_labels(i As Long, BAR_LABEL As String, itemname As String, itemmrp As Double, itemprice As Double)
    Dim wid As Single
    Dim hgt As Single
    
    On Error GoTo ERRHAND
    
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
        
    
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub TxtCessPer_GotFocus()
    Call CHANGEBOXCOLOR(TxtCessPer, True)
    TxtCessPer.SelStart = 0
    TxtCessPer.SelLength = Len(TxtCessPer.text)
End Sub

Private Sub TxtCessPer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtCess.SetFocus
        Case vbKeyEscape
            Frame1.Enabled = True
            If OptComper.Value = True Then
                TxtComper.SetFocus
            Else
                TxtComAmt.SetFocus
            End If

        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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
    Call CHANGEBOXCOLOR(TxtCessPer, False)
    Call TxttaxMRP_LostFocus
End Sub

Private Sub TxtHSN_GotFocus()
    Call CHANGEBOXCOLOR(txtHSN, True)
    txtHSN.SelStart = 0
    txtHSN.SelLength = Len(txtHSN.text)
End Sub

Private Sub TxtHSN_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtHSN.text) = "" And MDIMAIN.lblgst.Caption = "R" And Trim(UCase(lblcategory.Caption)) <> "SERVICE CHARGE" Then
                If MsgBox("HSN Code not entered. Are you sure?", vbYesNo + vbDefaultButton2, "PURCHASE ENTRY") = vbNo Then Exit Sub
            End If
            txtPD.Enabled = True
            txtPD.SetFocus
         Case vbKeyEscape
            TxtPoints.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
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

Private Sub TxtCustDisc_GotFocus()
    Call CHANGEBOXCOLOR(TxtCustDisc, True)
    TxtCustDisc.SelStart = 0
    TxtCustDisc.SelLength = Len(TxtCustDisc.text)
End Sub

Private Sub TxtCustDisc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Frame1.Enabled = True
            If OptComper.Value = True Then
                TxtComper.SetFocus
            Else
                TxtComAmt.SetFocus
            End If
        Case vbKeyEscape
            TxtLWRate.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtCustDisc_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCustDisc_LostFocus()
    Call CHANGEBOXCOLOR(TxtCustDisc, False)
    TxtCustDisc.text = Format(TxtCustDisc.text, "0.00")
End Sub


Private Sub txtNetrate_GotFocus()
    Call CHANGEBOXCOLOR(TxtNetrate, True)
    TxtNetrate.SelStart = 0
    TxtNetrate.SelLength = Len(TxtNetrate.text)
End Sub

Private Sub txtNetrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TxtNetrate.Text) = 0 Then Exit Sub
            If Trim(txtHSN.text) = "" Then
                txtHSN.Enabled = True
                txtHSN.SetFocus
            Else
                txtPD.SetFocus
            End If
        Case vbKeyEscape
            TxttaxMRP.SetFocus
        Case vbKeyDown
            If Val(Los_Pack.text) = 0 Then Los_Pack.text = 1
            If Val(TXTQTY.text) = 0 Then Exit Sub
            If Val(TXTPTR.text) = 0 Then Exit Sub
            Call CMDADD_Click
    End Select
End Sub

Private Sub TxtNetrate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNetrate_LostFocus()
    Call CHANGEBOXCOLOR(TxtNetrate, False)
    If Val(TxtNetrate.text) <> 0 Then
        TxtNetrate.text = Format(TxtNetrate.text, ".00")
        TXTPTR.Tag = Format(Round(Val(TxtNetrate.text) * 100 / (100 - Val(txtPD.text)), 4), "0.0000")
        TXTPTR.text = Format(Round(Val(TXTPTR.Tag) * 100 / (Val(TxttaxMRP.text) + 100), 4), "0.0000")
    End If
    Call TxttaxMRP_LostFocus
    
    If ADDCLICK = False Then
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            TxtRetailPercent.text = Val(MDIMAIN.LBLRT.Caption)
        End If
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtWsalePercent.text = Val(MDIMAIN.LBLWS.Caption)
        End If
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then
            txtSchPercent.text = Val(MDIMAIN.lblvp.Caption)
        End If
        If Val(MDIMAIN.LBLRT.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call TxtRetailPercent_LostFocus
        If Val(MDIMAIN.LBLWS.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtWsalePercent_LostFocus
        If Val(MDIMAIN.lblvp.Caption) > 0 And Val(TXTPTR.text) > 0 Then Call txtSchPercent_LostFocus
    End If
    
End Sub

Private Sub CHANGEBOXCOLOR(BOX As TextBox, texton As Boolean)
    If texton Then
        BOX.BackColor = &H98F3C1
    Else
        BOX.BackColor = vbWhite
    End If
End Sub

Private Function CLEAR_COMBO()
    CMBDISTRICT.Clear
    Dim rstfillcombo As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Set rstfillcombo = New ADODB.Recordset
    rstfillcombo.Open "Select DISTINCT TRX_GODOWN From TRANSMAST ORDER BY TRX_GODOWN", db, adOpenStatic, adLockReadOnly
    Do Until rstfillcombo.EOF
        If Not IsNull(rstfillcombo!TRX_GODOWN) Then CMBDISTRICT.AddItem (rstfillcombo!TRX_GODOWN)
        rstfillcombo.MoveNext
    Loop
    rstfillcombo.Close
    Set rstfillcombo = Nothing
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Function find_small_number()
    Dim i As Integer
    Dim sum_ary As Double
    Dim GROSSAMT As Double
    Dim totexpn As Double
    Dim NETCOST As Double
    On Error GoTo ERRHAND
    sum_ary = 0
    GROSSAMT = 0
    For i = 1 To grdsales.rows - 1
        'If Aray(i) < sn Then sn = Aray(i)
        sum_ary = sum_ary + Val(grdsales.TextMatrix(i, 3))
    Next i
    
    LBLEXP.Caption = ""
    lblqty.Caption = ""
    totexpn = Val(TxtTotalexp.text) + Val(txtaddlamt.text) + Val(TxtInsurance.text) + (Val(lbltotalwodiscount.Caption) * Val(TxtCST.text) / 100)
    For i = 1 To grdsales.rows - 1
        'grdsales.TextMatrix(i, 8) = Format(Round(((grossamt / (Val(grdsales.TextMatrix(i, 5)) * (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))))) + ((Val(grdsales.TextMatrix(i, 32)) / Val(grdsales.TextMatrix(i, 5))))), 4), ".0000")
        'grdsales.TextMatrix(i, 8) = Format(Round(((GROSSAMT / (Val(grdsales.TextMatrix(i, 5)) * (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))))) + ((Val(grdsales.TextMatrix(i, 32)) / ((Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))) * Val(grdsales.TextMatrix(i, 5)))))), 4), ".0000")
        
        If totexpn <> 0 Then
            grdsales.TextMatrix(i, 32) = Round((totexpn / sum_ary) * Val(grdsales.TextMatrix(i, 3)), 3)
        End If
        LBLEXP.Caption = Format(Val(LBLEXP.Caption) + Val(grdsales.TextMatrix(i, 32)), ".00")
        
        'GROSSAMT = Round((Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 14))) * (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 5))), 3)
        If (Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))) = 0 Then
            NETCOST = Round((Val(grdsales.TextMatrix(i, 13)) / Val(grdsales.TextMatrix(i, 3))) + Val(grdsales.TextMatrix(i, 32)), 3)
        Else
            NETCOST = Round((Val(grdsales.TextMatrix(i, 13)) / Val(grdsales.TextMatrix(i, 3))) + (Val(grdsales.TextMatrix(i, 32)) / ((Val(grdsales.TextMatrix(i, 3)) + Val(grdsales.TextMatrix(i, 14))) * Val(grdsales.TextMatrix(i, 5)))), 3)
        End If
        
        lblqty.Caption = Format(Val(lblqty.Caption) + Val(grdsales.TextMatrix(i, 3)), ".00")
        db.Execute "Update RTRXFILE set ITEM_NET_COST_PRICE = " & NETCOST & ", EXPENSE = " & Val(grdsales.TextMatrix(i, 32)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
        'db.Execute "Update RTRXFILE set EXPENSE = " & Val(grdsales.TextMatrix(i, 32)) & " WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='PI' AND VCH_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & "  "
       
    Next i
    Exit Function
ERRHAND:
    'MsgBox "Smallest Number is: " & sn
End Function

Private Sub TXTRCVDATE_GotFocus()
    If Not IsDate(TXTRCVDATE.text) And IsDate(TXTINVDATE.text) Then
       TXTRCVDATE.text = Format(TXTINVDATE, "DD/MM/YYYY")
    End If
    TXTRCVDATE.BackColor = &H98F3C1
    TXTRCVDATE.SelStart = 0
    TXTRCVDATE.SelLength = Len(TXTRCVDATE.text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTRCVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTRCVDATE.text = "  /  /    " Then
                TXTRCVDATE.text = Format(Date, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTRCVDATE.text) Then
                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
                TXTRCVDATE.SetFocus
                Exit Sub
            End If
            
'            If (DateValue(TXTRCVDATE.Text) < DateValue(MDIMAIN.DTFROM.value)) Or (DateValue(TXTRCVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.value))) Then
'                'db.Execute "delete from Users"
'                MsgBox "Please check the Date", vbOKOnly, "EzBiz"
'                TXTRCVDATE.SetFocus
'                Exit Sub
'            End If
            If Not IsDate(TXTRCVDATE.text) Then
                TXTRCVDATE.SetFocus
            Else
                TXTRCVDATE.text = Format(TXTRCVDATE.text, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
            End If
        Case vbKeyEscape
            TXTINVOICE.SetFocus
    End Select
End Sub

Private Sub TXTRCVDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRCVDATE_LostFocus()
    TXTRCVDATE.BackColor = vbWhite
End Sub

Private Function fillcategory()
        
    On Error GoTo ERRHAND
    Set Cmbcategory.DataSource = Nothing
    If CAT_REC.State = 1 Then
        CAT_REC.Close
        CAT_REC.Open "SELECT DISTINCT CATEGORY FROM CATEGORY ORDER BY CATEGORY", db, adOpenForwardOnly
    Else
        CAT_REC.Open "SELECT DISTINCT CATEGORY FROM CATEGORY ORDER BY CATEGORY", db, adOpenForwardOnly
    End If
    Set Cmbcategory.RowSource = CAT_REC
    Cmbcategory.ListField = "CATEGORY"
    
    Exit Function
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Function

