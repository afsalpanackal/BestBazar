VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSalesReturn 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Return"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19590
   Icon            =   "FrmSalesReturn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   19590
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   600
      TabIndex        =   39
      Top             =   2340
      Visible         =   0   'False
      Width           =   12255
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   3570
         Left            =   15
         TabIndex        =   40
         Top             =   15
         Width           =   12210
         _ExtentX        =   21537
         _ExtentY        =   6297
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
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   1815
      TabIndex        =   99
      Top             =   1650
      Visible         =   0   'False
      Width           =   13935
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   3840
         Left            =   30
         TabIndex        =   100
         Top             =   390
         Width           =   13890
         _ExtentX        =   24500
         _ExtentY        =   6773
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
         TabIndex        =   102
         Top             =   15
         Width           =   10110
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   "SOLD RATES FOR THE ITEM "
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
         TabIndex        =   101
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
      Left            =   1275
      TabIndex        =   53
      Top             =   165
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
      Left            =   6525
      TabIndex        =   24
      Top             =   7725
      Width           =   960
   End
   Begin VB.Frame Fram 
      BackColor       =   &H00F5EDDA&
      Height          =   9750
      Left            =   -120
      TabIndex        =   26
      Top             =   -45
      Width           =   19665
      Begin VB.TextBox TxtKMS 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   18825
         MaxLength       =   4
         TabIndex        =   163
         Top             =   1260
         Width           =   450
      End
      Begin VB.TextBox txtPin 
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
         Left            =   17040
         MaxLength       =   35
         TabIndex        =   161
         Top             =   675
         Width           =   2220
      End
      Begin VB.CommandButton CmDJSON 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "Print E-Invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15135
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   6870
         Width           =   1335
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
         Height          =   450
         Left            =   11535
         TabIndex        =   148
         Top             =   615
         Width           =   1155
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
         Height          =   450
         Left            =   11535
         TabIndex        =   147
         Top             =   120
         Width           =   1155
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
         Height          =   465
         Left            =   15150
         TabIndex        =   128
         Top             =   6375
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
         Height          =   390
         Left            =   15150
         TabIndex        =   127
         Top             =   5970
         Width           =   1320
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00F5EDDA&
         Caption         =   "Actual Address"
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
         Height          =   1605
         Left            =   12690
         TabIndex        =   121
         Top             =   -15
         Width           =   4305
         Begin VB.TextBox TxtArea 
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
            Left            =   615
            MaxLength       =   35
            TabIndex        =   162
            Top             =   570
            Width           =   3630
         End
         Begin VB.TextBox TxtPhone 
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
            Left            =   1245
            MaxLength       =   35
            TabIndex        =   123
            Top             =   1260
            Width           =   3000
         End
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
            Height          =   315
            Left            =   1245
            MaxLength       =   35
            TabIndex        =   122
            Top             =   915
            Width           =   3000
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Area"
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
            Index           =   51
            Left            =   45
            TabIndex        =   168
            Top             =   630
            Width           =   840
         End
         Begin MSForms.TextBox TxtBillAddress 
            Height          =   345
            Left            =   45
            TabIndex        =   126
            Top             =   210
            Width           =   4200
            VariousPropertyBits=   -1400879077
            MaxLength       =   150
            BorderStyle     =   1
            Size            =   "7408;609"
            SpecialEffect   =   0
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            Left            =   90
            TabIndex        =   125
            Top             =   1290
            Width           =   660
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "GST No"
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
            Left            =   75
            TabIndex        =   124
            Top             =   960
            Width           =   660
         End
      End
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00F5EDDA&
         Height          =   1575
         Left            =   150
         TabIndex        =   42
         Top             =   15
         Width           =   11370
         Begin VB.ComboBox cmbtype 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            ItemData        =   "FrmSalesReturn.frx":030A
            Left            =   6630
            List            =   "FrmSalesReturn.frx":031A
            Style           =   2  'Dropdown List
            TabIndex        =   145
            Top             =   1170
            Width           =   1860
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
            TabIndex        =   81
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
            TabIndex        =   51
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
            Left            =   6630
            MaxLength       =   100
            TabIndex        =   48
            Top             =   495
            Width           =   4680
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
            Left            =   3720
            MaxLength       =   10
            TabIndex        =   47
            Top             =   210
            Width           =   1260
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   6630
            TabIndex        =   50
            Top             =   150
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
            TabIndex        =   82
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
         Begin VB.Label lblIGST 
            BackColor       =   &H00F5EDDA&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   75
            TabIndex        =   167
            Top             =   1215
            Width           =   255
         End
         Begin VB.Label lblinvdetails 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6630
            TabIndex        =   156
            Top             =   810
            Width           =   4680
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Details"
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
            Left            =   5070
            TabIndex        =   155
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Cust Type"
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
            Left            =   5085
            TabIndex        =   146
            Top             =   1215
            Width           =   1530
         End
         Begin VB.Label lbltype 
            Height          =   375
            Left            =   5535
            TabIndex        =   120
            Top             =   1710
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblcredit 
            Height          =   525
            Left            =   9480
            TabIndex        =   70
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
            TabIndex        =   52
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
            Left            =   5085
            TabIndex        =   49
            Top             =   510
            Width           =   1530
         End
         Begin VB.Label INVDATE 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   0
            Left            =   2595
            TabIndex        =   46
            Top             =   210
            Width           =   1110
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
            Left            =   165
            TabIndex        =   45
            Top             =   210
            Width           =   870
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Return Date"
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
            Left            =   5085
            TabIndex        =   44
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            TabIndex        =   43
            Top             =   600
            Width           =   1005
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdsales 
         Height          =   4275
         Left            =   135
         TabIndex        =   93
         Top             =   1590
         Width           =   19395
         _ExtentX        =   34211
         _ExtentY        =   7541
         _Version        =   393216
         Rows            =   1
         Cols            =   40
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
         BackColor       =   &H00F5EDDA&
         Height          =   3645
         Left            =   135
         TabIndex        =   27
         Top             =   5790
         Width           =   15000
         Begin VB.TextBox Txtbarcode 
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
            TabIndex        =   158
            Top             =   480
            Width           =   2520
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
            Height          =   360
            Left            =   13305
            MaxLength       =   60
            TabIndex        =   7
            Top             =   480
            Width           =   1650
         End
         Begin VB.CommandButton CMDPRINTA5 
            Caption         =   "Prin&t Small"
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
            Left            =   4335
            TabIndex        =   22
            Top             =   1980
            Width           =   1140
         End
         Begin VB.TextBox txtNetrate 
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
            Left            =   7995
            MaxLength       =   7
            TabIndex        =   15
            Top             =   1155
            Width           =   1185
         End
         Begin VB.TextBox txtcategory 
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
            Height          =   360
            Left            =   4620
            TabIndex        =   2
            Top             =   480
            Width           =   1260
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Print"
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
            Left            =   3345
            TabIndex        =   21
            Top             =   1980
            Width           =   960
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
            Left            =   13485
            MaxLength       =   4
            TabIndex        =   118
            Top             =   3345
            Visible         =   0   'False
            Width           =   795
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
            Left            =   14280
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   3345
            Visible         =   0   'False
            Width           =   1350
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
            Left            =   6735
            MaxLength       =   7
            TabIndex        =   115
            Top             =   3600
            Visible         =   0   'False
            Width           =   1050
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
            Height          =   405
            Left            =   7815
            MaxLength       =   7
            TabIndex        =   114
            Top             =   3600
            Visible         =   0   'False
            Width           =   1065
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
            Height          =   405
            Left            =   12645
            MaxLength       =   7
            TabIndex        =   113
            Top             =   3675
            Visible         =   0   'False
            Width           =   1275
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
            ItemData        =   "FrmSalesReturn.frx":032F
            Left            =   11340
            List            =   "FrmSalesReturn.frx":0384
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   480
            Width           =   1095
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
            Left            =   10665
            MaxLength       =   7
            TabIndex        =   4
            Top             =   480
            Width           =   660
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
            Left            =   13005
            MaxLength       =   7
            TabIndex        =   17
            Top             =   1140
            Width           =   1560
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
            Left            =   12645
            MaxLength       =   7
            TabIndex        =   96
            Top             =   3345
            Visible         =   0   'False
            Width           =   1275
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
            Left            =   12975
            MaxLength       =   7
            TabIndex        =   94
            Top             =   3345
            Visible         =   0   'False
            Width           =   1020
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
            Left            =   8910
            MaxLength       =   7
            TabIndex        =   90
            Top             =   3225
            Visible         =   0   'False
            Width           =   795
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
            Left            =   9720
            MaxLength       =   7
            TabIndex        =   89
            Top             =   3960
            Visible         =   0   'False
            Width           =   1020
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
            Left            =   11910
            MaxLength       =   7
            TabIndex        =   87
            Top             =   3345
            Visible         =   0   'False
            Width           =   1035
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
            Left            =   7815
            MaxLength       =   7
            TabIndex        =   85
            Top             =   3210
            Visible         =   0   'False
            Width           =   1065
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
            Height          =   360
            Left            =   6735
            MaxLength       =   7
            TabIndex        =   83
            Top             =   3225
            Visible         =   0   'False
            Width           =   1050
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
            Left            =   9195
            MaxLength       =   7
            TabIndex        =   16
            Top             =   1140
            Width           =   960
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
            TabIndex        =   77
            Top             =   3240
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
            TabIndex        =   75
            Top             =   3990
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
            Height          =   405
            Left            =   11025
            TabIndex        =   72
            Top             =   4050
            Visible         =   0   'False
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
            Height          =   420
            Left            =   12165
            TabIndex        =   71
            Top             =   4035
            Visible         =   0   'False
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
            Height          =   405
            Left            =   9855
            TabIndex        =   64
            Top             =   4050
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton OPTNET 
            BackColor       =   &H00F5EDDA&
            Caption         =   "NET"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   8085
            TabIndex        =   61
            Top             =   1605
            Value           =   -1  'True
            Width           =   870
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
            Left            =   12630
            MaxLength       =   7
            TabIndex        =   62
            Top             =   3825
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.OptionButton OPTTaxMRP 
            BackColor       =   &H00F5EDDA&
            Caption         =   "Tax on MRP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   10095
            TabIndex        =   59
            Top             =   2820
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.OptionButton OPTVAT 
            BackColor       =   &H00F5EDDA&
            Caption         =   "TAX %"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6825
            TabIndex        =   60
            Top             =   1605
            Width           =   1140
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
            Left            =   6885
            MaxLength       =   7
            TabIndex        =   14
            Top             =   1155
            Width           =   1095
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
            Left            =   5925
            MaxLength       =   7
            TabIndex        =   54
            Top             =   3255
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
            Height          =   360
            Left            =   5700
            MaxLength       =   7
            TabIndex        =   13
            Top             =   1155
            Width           =   1170
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
            Height          =   360
            Left            =   4590
            MaxLength       =   7
            TabIndex        =   12
            Top             =   1155
            Width           =   1095
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
            Left            =   75
            TabIndex        =   18
            Top             =   1980
            Width           =   1125
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
            Left            =   5895
            TabIndex        =   3
            Top             =   480
            Width           =   4755
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
            Left            =   12435
            MaxLength       =   9
            TabIndex        =   6
            Top             =   480
            Width           =   855
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
            Left            =   2355
            TabIndex        =   20
            Top             =   1980
            Width           =   960
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
            Left            =   1245
            TabIndex        =   19
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
            Left            =   3135
            TabIndex        =   1
            Top             =   480
            Width           =   1470
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
            Left            =   1275
            MaxLength       =   15
            TabIndex        =   9
            Top             =   1170
            Width           =   2055
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
            TabIndex        =   28
            Top             =   3930
            Visible         =   0   'False
            Width           =   765
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
            Left            =   5520
            TabIndex        =   23
            Top             =   1980
            Width           =   960
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   345
            Left            =   3360
            TabIndex        =   10
            Top             =   1170
            Width           =   1215
            _ExtentX        =   2143
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
            Left            =   3360
            TabIndex        =   11
            Top             =   1170
            Width           =   1215
            _ExtentX        =   2143
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
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   9315
            TabIndex        =   108
            Top             =   3645
            Visible         =   0   'False
            Width           =   2595
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
               TabIndex        =   110
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
               TabIndex        =   109
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
               TabIndex        =   112
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
               TabIndex        =   111
               Top             =   195
               Width           =   1050
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5EDDA&
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   9195
            TabIndex        =   103
            Top             =   1455
            Width           =   2235
            Begin VB.OptionButton optdiscper 
               BackColor       =   &H00F5EDDA&
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
               Left            =   60
               TabIndex        =   105
               Top             =   150
               Value           =   -1  'True
               Width           =   960
            End
            Begin VB.OptionButton Optdiscamt 
               BackColor       =   &H00F5EDDA&
               Caption         =   "Dis&c Amt"
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
               Left            =   1080
               TabIndex        =   104
               Top             =   150
               Width           =   1125
            End
         End
         Begin MSMask.MaskEdBox TxtInvoiceDate 
            Height          =   375
            Left            =   45
            TabIndex        =   8
            Top             =   1155
            Width           =   1215
            _ExtentX        =   2143
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
            Index           =   49
            Left            =   600
            TabIndex        =   159
            Top             =   195
            Width           =   2520
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
            Height          =   285
            Index           =   48
            Left            =   45
            TabIndex        =   157
            Top             =   885
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Index           =   47
            Left            =   13305
            TabIndex        =   154
            Top             =   195
            Width           =   1650
         End
         Begin VB.Label lbllineno 
            Height          =   240
            Left            =   3075
            TabIndex        =   153
            Top             =   3180
            Width           =   750
         End
         Begin VB.Label lbltrxyear 
            Height          =   285
            Left            =   3075
            TabIndex        =   152
            Top             =   2700
            Width           =   735
         End
         Begin VB.Label lbltrxtype 
            Height          =   315
            Left            =   1140
            TabIndex        =   151
            Top             =   2730
            Width           =   675
         End
         Begin VB.Label lblvchno 
            Height          =   255
            Left            =   345
            TabIndex        =   150
            Top             =   2790
            Width           =   615
         End
         Begin VB.Label lblcost 
            Height          =   375
            Left            =   13500
            TabIndex        =   149
            Top             =   2760
            Visible         =   0   'False
            Width           =   1035
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
            Index           =   46
            Left            =   7995
            TabIndex        =   144
            Top             =   885
            Width           =   1185
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
            Index           =   45
            Left            =   7485
            TabIndex        =   143
            Top             =   2010
            Width           =   480
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
            Index           =   44
            Left            =   8835
            TabIndex        =   142
            Top             =   2010
            Width           =   510
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
            Index           =   43
            Left            =   10215
            TabIndex        =   141
            Top             =   2010
            Width           =   555
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
            Left            =   7995
            TabIndex        =   140
            Top             =   2010
            Width           =   825
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
            Left            =   9360
            TabIndex        =   139
            Top             =   2010
            Width           =   825
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
            Left            =   10785
            TabIndex        =   138
            Top             =   2010
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L.R.Price"
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
            Left            =   8940
            TabIndex        =   137
            Top             =   2340
            Width           =   885
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
            Left            =   17310
            TabIndex        =   136
            Top             =   2070
            Width           =   645
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
            Left            =   16005
            TabIndex        =   135
            Top             =   2070
            Width           =   780
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
            Left            =   16890
            TabIndex        =   134
            Top             =   2070
            Width           =   405
         End
         Begin VB.Label lblLWPrice 
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
            Left            =   9870
            TabIndex        =   133
            Top             =   2340
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "L.W.Price"
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
            Index           =   63
            Left            =   15000
            TabIndex        =   132
            Top             =   2070
            Width           =   900
         End
         Begin VB.Label LBLMRP 
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
            Left            =   8130
            TabIndex        =   131
            Top             =   2340
            Width           =   810
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
            Height          =   300
            Index           =   67
            Left            =   7485
            TabIndex        =   130
            Top             =   2340
            Width           =   630
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
            Left            =   4620
            TabIndex        =   129
            Top             =   195
            Width           =   1290
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
            Left            =   13485
            TabIndex        =   119
            Top             =   3090
            Visible         =   0   'False
            Width           =   2130
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
            Height          =   405
            Index           =   38
            Left            =   5745
            TabIndex        =   116
            Top             =   3600
            Visible         =   0   'False
            Width           =   960
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
            Left            =   3135
            TabIndex        =   107
            Top             =   195
            Width           =   1470
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
            Left            =   10170
            TabIndex        =   58
            Top             =   1140
            Width           =   1305
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
            Left            =   10665
            TabIndex        =   106
            Top             =   195
            Width           =   1740
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
            Height          =   255
            Index           =   33
            Left            =   13005
            TabIndex        =   98
            Top             =   885
            Width           =   1560
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Scheme Rate"
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
            Left            =   12645
            TabIndex        =   97
            Top             =   3090
            Visible         =   0   'False
            Width           =   1275
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
            Left            =   12975
            TabIndex        =   95
            Top             =   3090
            Visible         =   0   'False
            Width           =   1020
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
            Left            =   8910
            TabIndex        =   92
            Top             =   2970
            Visible         =   0   'False
            Width           =   795
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
            Left            =   9720
            TabIndex        =   91
            Top             =   3705
            Visible         =   0   'False
            Width           =   1020
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
            Left            =   11910
            TabIndex        =   88
            Top             =   3510
            Visible         =   0   'False
            Width           =   1035
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
            Index           =   27
            Left            =   7815
            TabIndex        =   86
            Top             =   2970
            Visible         =   0   'False
            Width           =   1065
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
            TabIndex        =   84
            Top             =   3750
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   25
            Left            =   9195
            TabIndex        =   78
            Top             =   885
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "PTR"
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
            Left            =   6735
            TabIndex        =   76
            Top             =   2970
            Visible         =   0   'False
            Width           =   1050
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
            Left            =   11055
            TabIndex        =   74
            Top             =   3780
            Visible         =   0   'False
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
            Left            =   12180
            TabIndex        =   73
            Top             =   3780
            Visible         =   0   'False
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
            Height          =   225
            Index           =   21
            Left            =   13440
            TabIndex        =   69
            Top             =   1485
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
            Height          =   495
            Left            =   13215
            TabIndex        =   68
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
            Left            =   9930
            TabIndex        =   67
            Top             =   3780
            Visible         =   0   'False
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
            Height          =   495
            Left            =   11490
            TabIndex        =   66
            Top             =   1725
            Width           =   1695
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
            ForeColor       =   &H00800080&
            Height          =   240
            Index           =   6
            Left            =   11520
            TabIndex        =   65
            Top             =   1485
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
            Left            =   12630
            TabIndex        =   63
            Top             =   3540
            Visible         =   0   'False
            Width           =   810
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
            Left            =   10170
            TabIndex        =   57
            Top             =   885
            Width           =   1305
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
            Left            =   6885
            TabIndex        =   56
            Top             =   885
            Width           =   1095
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
            Left            =   5925
            TabIndex        =   55
            Top             =   2970
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
            Left            =   5700
            TabIndex        =   41
            Top             =   885
            Width           =   1170
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
            TabIndex        =   38
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
            Left            =   5895
            TabIndex        =   37
            Top             =   195
            Width           =   4755
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
            Left            =   12435
            TabIndex        =   36
            Top             =   195
            Width           =   855
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
            Left            =   4590
            TabIndex        =   35
            Top             =   885
            Width           =   1095
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
            Left            =   11490
            TabIndex        =   34
            Top             =   885
            Width           =   1500
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
            TabIndex        =   33
            Top             =   3870
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
            Left            =   3360
            TabIndex        =   32
            Top             =   885
            Width           =   1215
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
            ForeColor       =   &H008080FF&
            Height          =   285
            Index           =   7
            Left            =   1275
            TabIndex        =   31
            Top             =   885
            Width           =   2055
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
            Left            =   11490
            TabIndex        =   25
            Top             =   1140
            Width           =   1500
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   3600
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
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
         Index           =   53
         Left            =   17040
         TabIndex        =   169
         Top             =   420
         Width           =   915
      End
      Begin MSForms.ComboBox TxtVehicle 
         Height          =   285
         Left            =   17025
         TabIndex        =   166
         Top             =   1260
         Width           =   1725
         VariousPropertyBits=   746604571
         ForeColor       =   255
         MaxLength       =   30
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3043;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         DropButtonStyle =   0
         BorderColor     =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle No."
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
         Index           =   50
         Left            =   17025
         TabIndex        =   165
         Top             =   1035
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kms"
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
         Index           =   79
         Left            =   18840
         TabIndex        =   164
         Top             =   1050
         Width           =   390
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   135
         TabIndex        =   80
         Top             =   1575
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   705
         TabIndex        =   79
         Top             =   1575
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmSalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim ACT_FLAG As Boolean
Dim PHY_CODE As New ADODB.Recordset
Dim PHYCODE_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT, M_ADD As Boolean
Dim PHY_PRERATE As New ADODB.Recordset
Dim PRERATE_FLAG As Boolean
Dim CHANGE_FLAG As Boolean
Dim Small_Print As Boolean

Private Sub CmbPack_GotFocus()
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = "1"
            'CmbPack.Enabled = False
            TXTQTY.Enabled = True
            Call FILL_PREVIIOUSRATE
            'TXTQTY.SetFocus
         Case vbKeyEscape
            'TXTUNIT.Text = ""
            'CmbPack.Enabled = False
            Los_Pack.Enabled = True
            Los_Pack.SetFocus
    End Select
End Sub

Private Sub cmbtype_GotFocus()
    FRMEGRDTMP.Visible = False
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
    
    
    'Call TXTPTR_LostFocus
    Call TXTQTY_LostFocus
    'Call Txtgrossamt_LostFocus
    Call txtPD_LostFocus
    
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim M_DATA As Double

    M_DATA = 0
    
    Txtpack.Text = 1
    If grdsales.rows <= Val(TXTSLNO.Text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text) + Val(TXTFREE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = 1 'Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Val(Los_Pack.Text) ' 1 'Val(TxtPack.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Round(Val(TXTRATE.Text) / Val(Los_Pack.Text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Format(Round(Val(TXTPTR.Text) / Val(Los_Pack.Text), 3), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Format((Val(txtprofit.Text)), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = IIf(Val(TxttaxMRP.Text) = 0, "", Format(Val(TxttaxMRP.Text), ".00")) 'TAX
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = IIf(IsDate(TXTEXPDATE.Text), TXTEXPDATE.Text, "") 'IIf(Trim(TXTEXPDATE.Text) = "/  /", "", TXTEXPDATE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = Val(TXTFREE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = Val(txtPD.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 26) = Format(Val(Txtgrossamt.Text), ".00")
    If optdiscper.Value = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 27) = "P"
    Else
        grdsales.TextMatrix(Val(TXTSLNO.Text), 27) = "A"
    End If
    grdsales.TextMatrix(Val(TXTSLNO.Text), 28) = Format(Val(Los_Pack.Text), ".00")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 29) = Trim(CmbPack.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 30) = Val(TxtWarranty.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 31) = Trim(CmbWrnty.Text)
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 33) = Trim(lbltrxtype.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 34) = Trim(lbltrxyear.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 35) = Val(lblvchno.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 36) = Val(lbllineno.Caption)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 37) = Trim(txtinvnodate.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 38) = IIf(IsDate(TxtInvoiceDate.Text), TxtInvoiceDate.Text, "")
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
    RSTRTRXFILE.Open "SELECT * From TRXSUB WHERE TRX_YEAR= '" & lbltrxyear.Caption & "' AND TRX_TYPE= '" & lbltrxtype.Caption & "' AND VCH_NO = " & Val(lblvchno.Caption) & " AND LINE_NO= " & Val(lbllineno.Caption) & "", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        Dim RSTRTRXFILE2 As ADODB.Recordset
        Set RSTRTRXFILE2 = New ADODB.Recordset
        RSTRTRXFILE2.Open "SELECT * From RTRXFILE WHERE ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "' AND TRX_YEAR= '" & RSTRTRXFILE!R_TRX_YEAR & "' AND TRX_TYPE= '" & RSTRTRXFILE!R_TRX_TYPE & "' AND VCH_NO = " & RSTRTRXFILE!R_VCH_NO & " AND LINE_NO= " & RSTRTRXFILE!R_LINE_NO & "", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTRTRXFILE2.EOF And RSTRTRXFILE2.BOF) Then
            grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = IIf(IsNull(RSTRTRXFILE2!item_COST), 0, RSTRTRXFILE2!item_COST)
            grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = IIf(IsNull(RSTRTRXFILE2!P_RETAIL), 0, RSTRTRXFILE2!P_RETAIL)
            grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = IIf(IsNull(RSTRTRXFILE2!P_WS), 0, RSTRTRXFILE2!P_WS)
            grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = IIf(IsNull(RSTRTRXFILE2!P_VAN), 0, RSTRTRXFILE2!P_VAN)
            grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = IIf(IsNull(RSTRTRXFILE2!P_CRTN), 0, RSTRTRXFILE2!P_CRTN)
            grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = IIf(IsNull(RSTRTRXFILE2!CRTN_PACK), 0, RSTRTRXFILE2!CRTN_PACK)
        End If
        RSTRTRXFILE2.Close
        Set RSTRTRXFILE2 = Nothing
    End If
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) = 0 Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(LBLCOST.Caption)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 32) = Val(LBLCOST.Caption)
    End If
    
'    If OLD_BILL = False Then Call checklastbill
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!TRX_TYPE = "SR"
        RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTRTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTRTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16))
        RSTRTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1))
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
        RSTRTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))

        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With rststock
            If Not (.EOF And .BOF) Then
                'rststock!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
                '!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text)
                
'                If (Val(TXTQTY.Text) + Val(TXTFREE.Text)) = 0 Then
'                    !ITEM_NET_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
'                Else
'                    !ITEM_NET_COST = Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text)) + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10)) / 100), 3)
'                End If
                
                RSTRTRXFILE!item_COST = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text), 3)
                RSTRTRXFILE!ITEM_COST_PRICE = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text), 3)
                If Val(Los_Pack.Text) = 0 Then
                    RSTRTRXFILE!ITEM_NET_COST_PRICE = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10)) / 100), 3)
                Else
                    RSTRTRXFILE!ITEM_NET_COST_PRICE = Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text)) + ((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10)) / 100), 3)
                End If
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                '!CLOSE_VAL = !CLOSE_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !CLOSE_VAL = Round(!item_COST * !CLOSE_QTY, 3)
                !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                '!RCPT_VAL = !RCPT_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !RCPT_VAL = Round(!item_COST * !RCPT_QTY, 3)

                '!P_RETAIL = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) / Val(Los_Pack.Text), 3)
                '!P_WS = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)) / Val(Los_Pack.Text), 3)
                '!P_CRTN = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) / Val(Los_Pack.Text), 3)
                '!P_VAN = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) / Val(Los_Pack.Text), 3)
                '!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
                '!CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23)) = "A" Then
                    !COM_FLAG = "A"
                    !COM_PER = 0
                    !COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22))
                Else
                    !COM_FLAG = "P"
                    !COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21))
                    !COM_AMT = 0
                End If
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = IIf(IsNull(!P_RETAIL), 0, !P_RETAIL)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = IIf(IsNull(!P_WS), 0, !P_WS)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = IIf(IsNull(!P_VAN), 0, !P_VAN)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = IIf(IsNull(!P_CRTN), 0, !P_CRTN)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = IIf(IsNull(!CRTN_PACK), 0, !CRTN_PACK)
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
                'rststock!Category = IIf(IsNull(rststock!Category), "OTHERS", rststock!Category)
                '!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
'                If (Val(TXTQTY.Text) + Val(TXTFREE.Text)) = 0 Then
'                    !ITEM_NET_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
'                Else
'                    !ITEM_NET_COST = Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text)) + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10)) / 100), 3)
'                End If
                
                RSTRTRXFILE!item_COST = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text), 3)
                RSTRTRXFILE!ITEM_COST_PRICE = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text), 3)
                If Val(Los_Pack.Text) = 0 Then
                    RSTRTRXFILE!ITEM_NET_COST_PRICE = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10)) / 100), 3)
                Else
                    RSTRTRXFILE!ITEM_NET_COST_PRICE = Round((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text)) + ((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) / Val(Los_Pack.Text)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10)) / 100), 3)
                End If
                
                !CLOSE_QTY = !CLOSE_QTY - RSTRTRXFILE!QTY
                !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                '!CLOSE_VAL = !CLOSE_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !CLOSE_VAL = Round(!item_COST * !CLOSE_QTY, 3)

                !RCPT_QTY = !RCPT_QTY - RSTRTRXFILE!QTY
                !RCPT_QTY = !RCPT_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
                If (IsNull(!RCPT_VAL)) Then !RCPT_VAL = 0
                '!RCPT_VAL =  !RCPT_VAL + (Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13)) / Val(Los_Pack.Text))
                !RCPT_VAL = Round(!item_COST * !RCPT_QTY, 3)
                '!P_RETAIL = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) / Val(Los_Pack.Text), 3)
                '!P_WS = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)) / Val(Los_Pack.Text), 3)
                '!P_CRTN = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) / Val(Los_Pack.Text), 3)
                '!P_VAN = Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) / Val(Los_Pack.Text), 3)
                '!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
                '!CRTN_PACK = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 23)) = "A" Then
                    !COM_FLAG = "A"
                    !COM_PER = 0
                    !COM_AMT = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 22))
                Else
                    !COM_FLAG = "P"
                    !COM_PER = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 21))
                    !COM_AMT = 0
                End If
                '!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 10))
                '!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))
                '!LOOSE_PACK = Val(Los_Pack.Text)
                '!PACK_TYPE = Trim(CmbPack.Text)
                '!WARRANTY = Val(TxtWarranty.Text)
                '!WARRANTY_TYPE = Trim(CmbWrnty.Text)
                'RSTRTRXFILE!MFGR = !MANUFACTURER
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = IIf(IsNull(!P_RETAIL), 0, !P_RETAIL)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 19) = IIf(IsNull(!P_WS), 0, !P_WS)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 25) = IIf(IsNull(!P_VAN), 0, !P_VAN)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 20) = IIf(IsNull(!P_CRTN), 0, !P_CRTN)
                If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 24) = IIf(IsNull(!CRTN_PACK), 0, !CRTN_PACK)
                rststock.Update
            End If
        End With
        rststock.Close
        Set rststock = Nothing
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
    End If
    RSTRTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    RSTRTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
    RSTRTRXFILE!VCH_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2))
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
    RSTRTRXFILE!gross_amt = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 26))
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
    RSTRTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
    RSTRTRXFILE!VCH_DESC = "Received From " & DataList2.Text
    RSTRTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
    'RSTRTRXFILE!ISSUE_QTY = 0
    RSTRTRXFILE!CST = 0
    If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 27)) = "P" Then
        RSTRTRXFILE!DISC_FLAG = "P"
    Else
        RSTRTRXFILE!DISC_FLAG = "A"
    End If
    RSTRTRXFILE!SCHEME = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))
    RSTRTRXFILE!EXP_DATE = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 12)), Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "dd/mm/yyyy"), Null)
    RSTRTRXFILE!FREE_QTY = 0
    RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTRTRXFILE!C_USER_ID = "SM"
    RSTRTRXFILE!check_flag = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))
    RSTRTRXFILE!S_TRX_TYPE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 33))
    RSTRTRXFILE!S_TRX_YEAR = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 34))
    RSTRTRXFILE!S_VCH_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 35))
    RSTRTRXFILE!S_LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 36))
    RSTRTRXFILE!INV_DETAILS = grdsales.TextMatrix(Val(TXTSLNO.Text), 37)
    RSTRTRXFILE!INV_DATE = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 38)), Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 38), "dd/mm/yyyy"), Null)
    'RSTRTRXFILE!M_USER_ID = DataList2.BoundText
    ''''RSTRTRXFILE!CHECK_FLAG = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15))  'MODE OF TAX
    'RSTRTRXFILE!PINV = Trim(TXTINVOICE.Text)
    RSTRTRXFILE.Update
    RSTRTRXFILE.Close

    M_DATA = 0
    Set RSTRTRXFILE = Nothing


    LBLTOTAL.Caption = ""
    lbltotalwodiscount = ""
    For i = 1 To grdsales.rows - 1
        lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
    Next i
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    TXTSLNO.Text = grdsales.rows
    TXTPRODUCT.Text = ""

    TXTITEMCODE.Text = ""
    TXTPTR.Text = ""
    txtNetrate.Text = ""
    Txtgrossamt.Text = ""
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.Text = ""
    TxttaxMRP.Text = ""
    txtPD.Text = ""
    txtprofit.Text = ""
    txtretail.Text = ""
    TxtRetailPercent.Text = ""
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    txtcrtnpack.Text = ""
    TXTRATE.Text = ""
    TxtComAmt.Text = ""
    TxtComper.Text = ""
    txtmrpbt.Text = ""
    txtBatch.Text = ""
    txtinvnodate.Text = ""
    TxtInvoiceDate.Text = "  /  /    "
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    OPTVAT.Value = True
    M_EDIT = False
    M_ADD = True
    cmdRefresh.Enabled = True
    CMDPRINT.Enabled = True
    CmdPrintA5.Enabled = True
    If TXTITEMCODE.Visible = True Then
        TXTITEMCODE.Enabled = True
        TXTITEMCODE.SetFocus
    Else
        txtcategory.Enabled = True
        txtcategory.SetFocus
    End If
    'TXTITEMCODE.Enabled = True
    'TXTITEMCODE.SetFocus
    CmbPack.Enabled = False
    TXTRATE.Enabled = False
    txtNetrate.Enabled = False
    TXTPTR.Enabled = False
    Los_Pack.Enabled = False
    TXTQTY.Enabled = False
    txtinvnodate.Enabled = False
    TxtInvoiceDate.Enabled = False
    txtBatch.Enabled = False
    TXTEXPIRY.Visible = False
    TXTEXPDATE.Enabled = False
    TxttaxMRP.Enabled = False
    txtPD.Enabled = False
    Txtgrossamt.Enabled = False
    txtBillNo.Enabled = False
    

    If grdsales.rows >= 18 Then grdsales.TopRow = grdsales.rows - 1

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

Private Sub CmdDelete_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset

    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub

    On Error GoTo ERRHAND
    db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 16)) & ""
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
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
    rstMaxNo.Open "Select MAX(LINE_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & " ", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing

    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
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
    RSTRTRXFILE.Open "SELECT * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
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
    RSTRTRXFILE.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
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
        grdsales.TextMatrix(i, 8) = Format(RSTRTRXFILE!item_COST, ".000")
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

    TXTSLNO.Text = Val(grdsales.rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.Text = ""
    TxttaxMRP.Text = ""
    txtPD.Text = ""
    txtprofit.Text = ""
    txtretail.Text = ""
    TxtRetailPercent.Text = ""
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    txtcrtnpack.Text = ""
    TXTRATE.Text = ""
    TxtComAmt.Text = ""
    TxtComper.Text = ""
    txtmrpbt.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    txtBatch.Text = ""
    txtinvnodate.Text = ""
    TxtInvoiceDate.Text = "  /  /    "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_ADD = True
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
     Dim rststock As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            Txtpack.Text = 1 '""
            Los_Pack.Text = ""
            CmbPack.ListIndex = -1
            TxtWarranty.Text = ""
            CmbWrnty.ListIndex = -1
            TXTFREE.Text = ""
            TxttaxMRP.Text = ""
            txtPD.Text = ""
            txtprofit.Text = ""
            txtretail.Text = ""
            TxtRetailPercent.Text = ""
            txtWsalePercent.Text = ""
            txtSchPercent.Text = ""
            txtWS.Text = ""
            txtvanrate.Text = ""
            Txtgrossamt.Text = ""
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
            txtBatch.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            txtcategory.Enabled = False
            TXTITEMCODE.Enabled = False
            TxtBarcode.Enabled = False
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

Private Sub CmdDeleteAll_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    On Error GoTo ERRHAND
    If Chkcancel.Value = 0 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE ALL", vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then Exit Sub
   
    For i = 1 To grdsales.rows - 1
        db.Execute "delete  From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='SR' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(i, 1)) & "' AND LINE_NO=" & Val(grdsales.TextMatrix(i, 16)) & ""
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
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
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdExit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdJSON_Click()
    If grdsales.rows <= 1 Then Exit Sub
    If MsgBox("Are you sure you want to generate E-Invoice?", vbYesNo, "EzBiz") = vbNo Then Exit Sub
    
    Call JSON_REPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"

End Sub

Private Sub CMDMODIFY_Click()

    If Val(TXTSLNO.Text) >= grdsales.rows Then Exit Sub

    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    CmbPack.Enabled = True
    TXTRATE.Enabled = True
    txtNetrate.Enabled = True
    TXTPTR.Enabled = True
    Los_Pack.Enabled = True
    txtinvnodate.Enabled = True
    TxtInvoiceDate.Enabled = True
    txtBatch.Enabled = True
    TXTEXPIRY.Visible = True
    TXTEXPDATE.Enabled = True
    TxttaxMRP.Enabled = True
    txtPD.Enabled = True
    Txtgrossamt.Enabled = True
    TXTQTY.Enabled = True
    TXTQTY.SetFocus

End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            Txtpack.Text = 1 '""
            Los_Pack.Text = ""
            CmbPack.ListIndex = -1
            TxtWarranty.Text = ""
            CmbWrnty.ListIndex = -1
            TXTFREE.Text = ""
            TxttaxMRP.Text = ""
            txtPD.Text = ""
            txtprofit.Text = ""
            txtretail.Text = ""
            TxtRetailPercent.Text = ""
            txtWsalePercent.Text = ""
            txtSchPercent.Text = ""
            txtWS.Text = ""
            txtvanrate.Text = ""
            Txtgrossamt.Text = ""
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
            txtBatch.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            txtcategory.Enabled = False
            TXTITEMCODE.Enabled = False
            TxtBarcode.Enabled = False
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

Private Sub CmdPrint_Click()
    
   If grdsales.rows = 1 Then Exit Sub
'    If Month(MDIMAIN.DTFROM.value) >= 4 And Year(MDIMAIN.DTFROM.value) >= 2021 Then Exit Sub
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        DataList2.SetFocus
        Exit Sub
    End If
    
    If MDIMAIN.StatusBar.Panels(8).Text <> "Y" Then
        Small_Print = False
        Call Generateprint
    Else
        Call ReportGeneratION_estimate
        On Error GoTo CLOSEFILE
        Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
    End If
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
    Print #1, "EXIT"
    Close #1
    
    '//HERE write the proper path where your command.com file exist
    Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & Rptpath & "REPO.BAT N", vbHide
    cmdRefresh.SetFocus
    'Call cmdRefresh_Click
    Screen.MousePointer = vbNormal
    End If
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Public Function Generateprint()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Long
    Dim b As Integer
    Dim Num As Currency

    On Error GoTo ERRHAND
    b = 0
        
    Screen.MousePointer = vbHourglass
    db.Execute "delete From TEMPTRXFILE "
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        RSTTRXFILE.AddNew
        
        RSTTRXFILE!TRX_TYPE = "SR"
        'RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = grdsales.TextMatrix(i, 25)
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        
        
        RSTTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(i, 13))
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
        RSTTRXFILE!item_COST = Val(grdsales.TextMatrix(i, 8))
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(i, 17))
        'RSTTRXFILE!P_DISC = Val(grdsales.TextMatrix(i, 17))
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(i, 6))
        
        If MDIMAIN.lblgst.Caption = "R" Then
            RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 9)) + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)
            RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 7))
            RSTTRXFILE!P_RETAIL = (Val(grdsales.TextMatrix(i, 9)) + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)) * Val(grdsales.TextMatrix(i, 28))
            RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 28))   ''+ (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)
            RSTTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(i, 10))
        Else
            RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 9)) + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)
            RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 9)) + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)
            RSTTRXFILE!P_RETAIL = (Val(grdsales.TextMatrix(i, 9)) + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)) * Val(grdsales.TextMatrix(i, 28))
            RSTTRXFILE!P_RETAILWOTAX = (Val(grdsales.TextMatrix(i, 9)) + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 10)) / 100)) * Val(grdsales.TextMatrix(i, 28))
            RSTTRXFILE!SALES_TAX = 0
        End If
        'RSTTRXFILE!P_WS = Val(grdsales.TextMatrix(i, 19))
        'RSTTRXFILE!P_CRTN = Val(grdsales.TextMatrix(i, 20))
        'RSTTRXFILE!CRTN_PACK = Val(grdsales.TextMatrix(i, 24))
        'RSTTRXFILE!P_VAN = Val(grdsales.TextMatrix(i, 25))
        'RSTTRXFILE!GROSS_AMT = Val(grdsales.TextMatrix(i, 26))
        
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 28))
        RSTTRXFILE!PACK_TYPE = "Nos" 'Trim(CmbPack.Text)
        RSTTRXFILE!WARRANTY = Val(TxtWarranty.Text)
        RSTTRXFILE!WARRANTY_TYPE = Trim(CmbWrnty.Text)
        RSTTRXFILE!UNIT = 1 'Val(grdsales.TextMatrix(I, 4))
        RSTTRXFILE!VCH_DESC = grdsales.TextMatrix(i, 37) & IIf(IsDate(grdsales.TextMatrix(i, 38)), " DTD " & Format(grdsales.TextMatrix(i, 38), "dd/mm/yyyy"), "")
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 11))
        'RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!SCHEME = Val(grdsales.TextMatrix(i, 14))
        If IsDate(grdsales.TextMatrix(i, 12)) Then
            RSTTRXFILE!EXP_DATE = Format(grdsales.TextMatrix(i, 12), "dd/mm/yyyy")
        End If
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        'RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!check_flag = Trim(grdsales.TextMatrix(i, 15))
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        'RSTTRXFILE!C_USER_ID = "SM"
        'RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE!CESS_PER = 0
        RSTTRXFILE!cess_amt = 0
        RSTTRXFILE!kfc_tax = 0
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTTRXFILE!C_USER_ID = IIf(IsNull(RSTITEMMAST!REMARKS), "", Left(RSTITEMMAST!REMARKS, 8))
            RSTTRXFILE!MFGR = IIf(IsNull(RSTITEMMAST!ITEM_MAL), "", RSTITEMMAST!ITEM_MAL)
            RSTTRXFILE!M_USER_ID = IIf(IsNull(RSTITEMMAST!FULL_PACK), "", RSTITEMMAST!FULL_PACK)
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE.Update
    Next i
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    'lblnetamount.Tag = Round(Val(Round(Val(LBLTOTAL.Caption), 0)) - Val(Round(Val(LBLTOTAL.Caption), 2)), 2)
    'Num = CCur(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0))
    
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompTin, CompCST As String
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        CompName = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
        CompAddress1 = IIf(IsNull(RSTCOMPANY!Address), "", RSTCOMPANY!Address)
        CompAddress2 = IIf(IsNull(RSTCOMPANY!HO_NAME), "", RSTCOMPANY!HO_NAME)
        If Trim(CompAddress2) = "" Then
            CompAddress2 = "Ph: " & IIf(IsNull(RSTCOMPANY!TEL_NO), "", RSTCOMPANY!TEL_NO) & IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", ", " & RSTCOMPANY!FAX_NO) & _
                        IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!FAX_NO)
        Else
            CompAddress3 = "Ph: " & IIf(IsNull(RSTCOMPANY!TEL_NO), "", RSTCOMPANY!TEL_NO) & IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", ", " & RSTCOMPANY!FAX_NO) & _
                        IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!FAX_NO)
        End If
        CompTin = IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", "", "GSTIN No. " & RSTCOMPANY!CST)
        CompCST = IIf(IsNull(RSTCOMPANY!DL_NO) Or RSTCOMPANY!DL_NO = "", "", "CST No. " & RSTCOMPANY!DL_NO)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Sleep (300)
    If MDIMAIN.lblgst.Caption = "R" Then
        If Small_Print = True Then
            ReportNameVar = Rptpath & "rptSRA5"
        Else
            ReportNameVar = Rptpath & "rptSR"
        End If
    Else
        If Small_Print = True Then
            ReportNameVar = Rptpath & "rptSRCA5"
        Else
            ReportNameVar = Rptpath & "rptSRC"
        End If
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.RecordSelectionFormula = "( {TRXFILE.TRX_TYPE}='SR' AND {TRXFILE.VCH_NO}= " & Val(txtBillNo.Text) & " )"
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
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.Text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
        If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.Text = "'" & CompName & "'"
        If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.Text = "'" & CompAddress1 & "'"
        If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.Text = "'" & CompAddress2 & "'"
        If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.Text = "'" & CompAddress3 & "'"
        If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.Text = "'" & CompTin & "'"
        If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.Text = "'" & CompCST & "'"
        If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.Text = "'" & TXTDEALER.Text & "'"
        If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'" & Trim(TxtBillAddress.Text) & "'"
        If CRXFormulaField.Name = "{@HSNSUM_FLAG}" Then
            If Val(LBLTOTAL.Caption) >= Val(MDIMAIN.LBLHSNSUM.Caption) Then
                CRXFormulaField.Text = "'Y'"
            Else
                CRXFormulaField.Text = "'N'"
            End If
        End If
        If Trim(TXTTIN.Text) <> "" Then If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'GSTIN: ' & '" & Trim(TXTTIN.Text) & "'"
        'If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
        If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & TxtPhone.Text & "'"
'        If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.Text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
'        If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
'        If CRXFormulaField.Name = "{@TIN}" Then CRXFormulaField.Text = "'" & TXTTIN.Text & "'"
'        If CRXFormulaField.Name = "{@Phone}" Then CRXFormulaField.Text = "'" & TxtPhone.Text & "'"
        If CRXFormulaField.Name = "{@VCH_NO}" Then
            Me.Tag = "SR-" & Format(Trim(txtBillNo.Text), bill_for)
            CRXFormulaField.Text = "'" & Me.Tag & "' "
        End If
'        If CRXFormulaField.Name = "{@Vehicle}" Then CRXFormulaField.Text = "'" & Trim(TxtVehicle.Text) & "'"
'        If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = "'" & Format(Val(LBLDISCAMT.Caption), "0.00") & "'"
'        If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
'        If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.Text = "'" & Trim(TxtCode.Text) & "'"
'        If CRXFormulaField.Name = "{@P_Bal}" Then CRXFormulaField.Text = "'" & Format(Val(txtOutstanding.Text), "0.00") & "'"
'
'        'If CRXFormulaField.Name = "{@unit}" Then CRXFormulaField.Text = "'" & Trim(lblunit.Caption) & "'"
'        If Trim(TXTTIN.Text) = "" Then
'            If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'TAX INVOICE FORM 8B'"
'        Else
'            If CRXFormulaField.Name = "{@ZFORM}" Then CRXFormulaField.Text = "'TAX INVOICE FORM 8'"
'        End If
'        If lblcredit.Caption = "0" Then
'            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Cash'"
'        Else
'            If Val(txtcrdays.Text) > 0 Then
'                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'" & txtcrdays.Text & "'" & "' Days Credit'"
'            Else
'                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.Text = "'Credit'"
'            End If
'        End If
    Next
    
    'Bill
    If Small_Print = True Then
        Set Printer = Printers(thermalprinter)
        Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
        If MDIMAIN.LBLTHPREVIEW.Caption = "Y" Then
            'Preview
            frmreport.Caption = "CREDIT NOTE"
            Call GENERATEREPORT
            Screen.MousePointer = vbNormal
        Else
            '    '''No Preview
            Report.PrintOut (False)
            Set CRXFormulaFields = Nothing
            Set CRXFormulaField = Nothing
            Set crxApplication = Nothing
            Set Report = Nothing
            Call cmdRefresh_Click
            Screen.MousePointer = vbNormal
            Exit Function
        End If
    Else
        Set Printer = Printers(billprinter)
        Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
        If MDIMAIN.StatusBar.Panels(13).Text = "Y" Then
            'Preview
            frmreport.Caption = "CREDIT NOTE"
            Call GENERATEREPORT
            Screen.MousePointer = vbNormal
        Else
            '    '''No Preview
            Report.PrintOut (False)
            Set CRXFormulaFields = Nothing
            Set CRXFormulaField = Nothing
            Set crxApplication = Nothing
            Set Report = Nothing
            Call cmdRefresh_Click
            Screen.MousePointer = vbNormal
            Exit Function
        End If
    End If
    
        
'    frmreport.Caption = "SALES RETURN"
'    Call GENERATEREPORT

    CMDEXIT.Enabled = False
    Screen.MousePointer = vbNormal
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Function


Private Sub CmdPrintA5_Click()
    If grdsales.rows = 1 Then Exit Sub
'    If Month(MDIMAIN.DTFROM.value) >= 4 And Year(MDIMAIN.DTFROM.value) >= 2021 Then Exit Sub
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Customer From List", vbOKOnly, "Sale Bill..."
        DataList2.SetFocus
        Exit Sub
    End If
    
    If MDIMAIN.StatusBar.Panels(8).Text <> "Y" Then
        Small_Print = True
        Call Generateprint
    Else
        Call ReportGeneratION_estimate
        On Error GoTo CLOSEFILE
    Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
    End If
    On Error GoTo ERRHAND
    
    Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
    Print #1, "EXIT"
    Close #1
    
    '//HERE write the proper path where your command.com file exist
    Shell "C:\WINDOWS\SYSTEM32\CMD.EXE /C " & Rptpath & "REPO.BAT N", vbHide
    cmdRefresh.SetFocus
    'Call cmdRefresh_Click
    Screen.MousePointer = vbNormal
    End If
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
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
            
    If grdsales.rows <= 1 Or DataList2.BoundText = "130000" Then
        lblcredit.Caption = "0"
        Call appendpurchase
    Else
        If IsNull(DataList2.SelectedItem) Then
            MsgBox "Select Customer From List", vbOKOnly, "Sales Return..."
            DataList2.SetFocus
            Exit Sub
        End If
        If Not IsDate(TXTINVDATE.Text) Then
            MsgBox "Enter Returned Date", vbOKOnly, "Sales Return"
            Exit Sub
        End If
        Me.Enabled = False
        MDIMAIN.cmdpurchase.Enabled = False
        Set creditbill = Me
        frmCREDIT.Show
    End If
    Set grdtmp.DataSource = Nothing
    FRMEGRDTMP.Visible = False
End Sub

Private Sub Command4_Click()
    If CMDEXIT.Enabled = False Then Exit Sub
    If Val(txtBillNo.Text) = 1 Then Exit Sub
    txtBillNo.Text = Val(txtBillNo.Text) - 1
    
    grdsales.rows = 1
    TXTSLNO.Text = 1
    cmdRefresh.Enabled = False
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTREMARKS.Text = ""
    lblinvdetails.Caption = ""
    TXTSLNO.Text = ""
    TXTITEMCODE.Text = ""
    TXTPRODUCT.Text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.Text = ""
    TxttaxMRP.Text = ""
    txtPD.Text = ""
    txtprofit.Text = ""
    txtretail.Text = ""
    TxtRetailPercent.Text = ""
    
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    txtcrtnpack.Text = ""
    txtBatch.Text = ""
    txtinvnodate.Text = ""
    TxtInvoiceDate.Text = "  /  /    "
    TXTRATE.Text = ""
    txtmrpbt.Text = ""
    TXTPTR.Text = ""
    txtNetrate.Text = ""
    Txtgrossamt.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    txtcategory.Text = ""
    txtaddlamt.Text = ""
    txtcramt.Text = ""
    TxtInsurance.Text = ""
    TxtCST.Text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    TXTDISCAMOUNT.Text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    M_ADD = False
    
    Chkcancel.Value = 0
    Call txtBillNo_KeyDown(13, 0)
End Sub

Private Sub Command5_Click()
    If CMDEXIT.Enabled = False Then Exit Sub
    Dim rstBILL As ADODB.Recordset
    Dim lastbillno As Double
    On Error GoTo ERRHAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'SR'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        lastbillno = IIf(IsNull(rstBILL.Fields(0)), 0, rstBILL.Fields(0))
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    If Val(txtBillNo.Text) > lastbillno Then Exit Sub
    txtBillNo.Text = Val(txtBillNo.Text) + 1
    
    grdsales.rows = 1
    TXTSLNO.Text = 1
    cmdRefresh.Enabled = False
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTREMARKS.Text = ""
    lblinvdetails.Caption = ""
    TXTSLNO.Text = ""
    TXTITEMCODE.Text = ""
    TXTPRODUCT.Text = ""
    FRMEGRDTMP.Visible = False
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.Text = ""
    TxttaxMRP.Text = ""
    txtPD.Text = ""
    txtprofit.Text = ""
    txtretail.Text = ""
    TxtRetailPercent.Text = ""
    
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    txtcrtnpack.Text = ""
    txtBatch.Text = ""
    txtinvnodate.Text = ""
    TxtInvoiceDate.Text = "  /  /    "
    TXTRATE.Text = ""
    txtmrpbt.Text = ""
    TXTPTR.Text = ""
    txtNetrate.Text = ""
    Txtgrossamt.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    txtcategory.Text = ""
    txtaddlamt.Text = ""
    txtcramt.Text = ""
    TxtInsurance.Text = ""
    TxtCST.Text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    TXTDISCAMOUNT.Text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    M_ADD = False
    
    Chkcancel.Value = 0
    Call txtBillNo_KeyDown(13, 0)
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Form_Activate()
    On Error GoTo ERRHAND
    txtBillNo.SetFocus
    Exit Sub
ERRHAND:
    If err.Number = 5 Then Exit Sub
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    'txtBillNo.Visible = False
    'cmditemcreate.Visible = False
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    
    Set TRXMAST = New ADODB.Recordset
    
    TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' ", db, adOpenStatic, adLockReadOnly
    'TRXMAST.Open "Select MAX(VCH_NO) From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'SR'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    ACT_FLAG = True
    PRERATE_FLAG = True
    grdsales.ColWidth(0) = 500
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 3700
    grdsales.ColWidth(3) = 1000
    grdsales.ColWidth(4) = 0 ' 800
    grdsales.ColWidth(5) = 0 '800
    grdsales.ColWidth(6) = 1200
    grdsales.ColWidth(7) = 0 '800
    grdsales.ColWidth(9) = 1000
    grdsales.ColWidth(10) = 1100
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 0 '1100
    grdsales.ColWidth(13) = 1700 '1100
    grdsales.ColWidth(16) = 0
    grdsales.ColWidth(15) = 0
    grdsales.ColWidth(17) = 1000
    grdsales.ColWidth(18) = 1000
    grdsales.ColWidth(19) = 1000
    grdsales.ColWidth(20) = 0
    grdsales.ColWidth(21) = 0
    grdsales.ColWidth(22) = 0
    grdsales.ColWidth(23) = 0
    grdsales.ColWidth(24) = 0
    grdsales.ColWidth(25) = 0
    grdsales.ColWidth(26) = 1700
    grdsales.ColWidth(27) = 0
    grdsales.ColWidth(28) = 0
    grdsales.ColWidth(29) = 0
    grdsales.ColWidth(30) = 0
    grdsales.ColWidth(31) = 0
    grdsales.ColWidth(32) = 0
    grdsales.ColWidth(33) = 0
    grdsales.ColWidth(34) = 0
    grdsales.ColWidth(35) = 0
    grdsales.ColWidth(36) = 0
    grdsales.ColWidth(37) = 0
    grdsales.ColWidth(38) = 0
    grdsales.ColWidth(39) = 0
    
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
    grdsales.TextArray(7) = "" '"PTS"
    If frmLogin.rs!Level = "2" Or frmLogin.rs!Level = "5" Then
        grdsales.TextArray(8) = ""
        grdsales.ColWidth(8) = 0
    Else
        grdsales.TextArray(8) = "COST"
        grdsales.ColWidth(8) = 800
    End If
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
    grdsales.TextArray(20) = "" '"Cartn Price"
    grdsales.TextArray(21) = "" '"Comm %"
    grdsales.TextArray(22) = "" '"Comm Amt"
    grdsales.TextArray(23) = "" '"Comm Flag"
    grdsales.TextArray(24) = "" '"Cnt Pck"
    grdsales.TextArray(25) = "" '"Van Rate"
    grdsales.TextArray(26) = "GROSS AMOUNT"
    grdsales.TextArray(27) = "DISC_FLAG"
    
'    If MDIMAIN.StatusBar.Panels(15).Text = "Y" Then
'        Label1(35).Visible = False
'        TXTITEMCODE.Visible = False
'        Label1(40).Left = 600
'        Label1(40).Width = Val(Label1(40).Width) + 1450
'        txtcategory.Left = 600
'        txtcategory.Width = Val(Label1(40).Width)
'    End If
        
    PHYFLAG = True
    PHYCODE_FLAG = True
    TXTPRODUCT.Enabled = False
    txtcategory.Enabled = False
    TXTITEMCODE.Enabled = False
    TxtBarcode.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TXTDATE.Text = Date
    TXTEXPDATE.Enabled = False
    txtBatch.Enabled = False
    txtinvnodate.Enabled = False
    TxtInvoiceDate.Enabled = False
    CmdDelete.Enabled = False
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
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If grdsales.rows <= 1 Then db.Execute "delete From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & ""
        If PHYFLAG = False Then PHY.Close
        If PHYCODE_FLAG = False Then PHY_CODE.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If PRERATE_FLAG = False Then PHY_PRERATE.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdsales_DblClick()
    If grdsales.rows <= 1 Then Exit Sub
    If M_EDIT = True Then Exit Sub
    TXTSLNO.Text = grdsales.TextMatrix(grdsales.Row, 0)
    Call TXTSLNO_KeyDown(13, 0)
    CMDMODIFY_Click
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long

    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn

    
            TXTITEMCODE.Text = grdtmp.Columns(0)
            TXTPRODUCT.Text = grdtmp.Columns(1)
            For i = 1 To grdsales.rows - 1
                If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                End If
            Next i
            
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                On Error Resume Next
                CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                CmbWrnty.Text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                On Error GoTo ERRHAND
                txtBatch.Text = ""
                TXTEXPIRY.Text = "  /  "
                TXTRATE.Text = IIf(IsNull(RSTRXFILE!MRP), "", RSTRXFILE!MRP)
                txtmrpbt.Text = ""
                Select Case cmbtype.ListIndex
                    Case 1
                        TXTPTR.Text = IIf(IsNull(RSTRXFILE!P_WS), "", RSTRXFILE!P_WS)
                        txtNetrate.Text = IIf(IsNull(RSTRXFILE!P_WS), "", RSTRXFILE!P_WS)
                    Case 2
                        TXTPTR.Text = IIf(IsNull(RSTRXFILE!P_VAN), "", RSTRXFILE!P_VAN)
                        txtNetrate.Text = IIf(IsNull(RSTRXFILE!P_VAN), "", RSTRXFILE!P_VAN)
                    Case 3
                        TXTPTR.Text = IIf(IsNull(RSTRXFILE!MRP), "", RSTRXFILE!MRP)
                        txtNetrate.Text = IIf(IsNull(RSTRXFILE!MRP), "", RSTRXFILE!MRP)
                    Case Else
                        TXTPTR.Text = IIf(IsNull(RSTRXFILE!P_RETAIL), "", RSTRXFILE!P_RETAIL)
                        txtNetrate.Text = IIf(IsNull(RSTRXFILE!P_RETAIL), "", RSTRXFILE!P_RETAIL)
                End Select
                TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", RSTRXFILE!SALES_TAX)
                OPTVAT.Value = True
            End If
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing
            
'            Set RSTRXFILE = New ADODB.Recordset
'            RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ORDER BY CREATE_DATE", db, adOpenStatic, adLockReadOnly, adCmdText
'            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
'                RSTRXFILE.MoveLast
'                TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
'                If IsNull(RSTRXFILE!LINE_DISC) Then
'                    Txtpack.Text = ""
'                Else
'                    Txtpack.Text = RSTRXFILE!LINE_DISC
'                End If
'                Txtpack.Text = 1
'                Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
'                On Error Resume Next
'                CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
'                CmbWrnty.Text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
'                On Error GoTo eRRhAND
'            Else
'                TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
'                Txtpack.Text = 1
'                Los_Pack.Text = 1
'                TxtWarranty.Text = ""
'                On Error Resume Next
'                CmbPack.Text = "Nos"
'                CmbWrnty.ListIndex = -1
'                TXTEXPDATE.Text = IIf(IsNull(RSTRXFILE!EXP_DATE), "  /  /    ", Format(RSTRXFILE!EXP_DATE, "DD/MM/YYYY"))
'                On Error GoTo eRRhAND
'                txtBatch.Text = ""
'                TXTEXPIRY.Text = "  /  "
'                TXTRATE.Text = ""
'                txtmrpbt.Text = ""
'                TXTPTR.Text = ""
'                txtNetrate.Text = ""
'                txtretail.Text = ""
'                txtWS.Text = ""
'                txtvanrate.Text = ""
'                txtcrtn.Text = ""
'                txtcrtnpack.Text = ""
'                txtprofit.Text = ""
'                TxttaxMRP.Text = ""
'                Los_Pack.Text = "1"
'                TxtWarranty.Text = ""
'                On Error Resume Next
'                CmbPack.Text = "Nos"
'                CmbWrnty.ListIndex = -1
'                On Error GoTo eRRhAND
'                OPTVAT.value = True
'            End If
'            RSTRXFILE.Close
'            Set RSTRXFILE = Nothing

            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Fram.Enabled = True
            TXTPRODUCT.Enabled = False
            txtcategory.Enabled = False
            TXTITEMCODE.Enabled = False
            TxtBarcode.Enabled = False
            
            CmbPack.Enabled = True
            TXTRATE.Enabled = True
            txtNetrate.Enabled = True
            TXTPTR.Enabled = True
            Los_Pack.Enabled = True
            txtinvnodate.Enabled = True
            TxtInvoiceDate.Enabled = True
            txtBatch.Enabled = True
            TXTEXPIRY.Visible = True
            TXTEXPDATE.Enabled = True
            TxttaxMRP.Enabled = True
            txtPD.Enabled = True
            Txtgrossamt.Enabled = True
            TXTQTY.Enabled = True
            
            'TXTQTY.SetFocus
            Call FILL_PREVIIOUSRATE
            'TxtPack.Enabled = True
            'TxtPack.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTFREE.Text = ""
            Fram.Enabled = True
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
ERRHAND:
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
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "Sales Return"
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
'                    MsgBox "SELECT MODE OF TAX ....", vbOKOnly, "Sales Return"
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
    txtBatch.SelLength = Len(txtBatch.Text)
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(txtBatch.Text) = "" Then Exit Sub
            'txtBatch.Enabled = False
            TXTEXPIRY.Visible = True
            TXTEXPIRY.SetFocus
        Case vbKeyEscape
            Call FILL_PREVIIOUSRATE
            If Not IsDate(TxtInvoiceDate.Text) Then
                TxtInvoiceDate.Enabled = True
                TxtInvoiceDate.SetFocus
            Else
                TXTQTY.Enabled = True
                Call FILL_PREVIIOUSRATE
            End If
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

Private Sub txtBillNo_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTDIST As ADODB.Recordset
    Dim RSTTRNSMAST As ADODB.Recordset
    Dim i As Long

    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            grdsales.rows = 1
            i = 0
            LBLTOTAL.Caption = ""
            lbltotalwodiscount = ""
            grdsales.rows = 1
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
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
                grdsales.TextMatrix(i, 7) = Format(rstTRXMAST!SALES_PRICE, ".000")
                grdsales.TextMatrix(i, 8) = Format(rstTRXMAST!item_COST, ".000")
                grdsales.TextMatrix(i, 9) = Format(rstTRXMAST!PTR, ".000")
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
                If rstTRXMAST!DISC_FLAG = "P" Then
                    grdsales.TextMatrix(i, 27) = "P"
                Else
                    grdsales.TextMatrix(i, 27) = "A"
                End If
                grdsales.TextMatrix(i, 28) = IIf(IsNull(rstTRXMAST!LOOSE_PACK), 1, rstTRXMAST!LOOSE_PACK)
                grdsales.TextMatrix(i, 29) = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)

                grdsales.TextMatrix(i, 30) = IIf(IsNull(rstTRXMAST!WARRANTY), "", rstTRXMAST!WARRANTY)
                grdsales.TextMatrix(i, 31) = IIf(IsNull(rstTRXMAST!WARRANTY_TYPE), "", rstTRXMAST!WARRANTY_TYPE)
                
                grdsales.TextMatrix(i, 32) = IIf(IsNull(rstTRXMAST!item_COST), "", rstTRXMAST!item_COST)
                grdsales.TextMatrix(i, 33) = IIf(IsNull(rstTRXMAST!S_TRX_TYPE), "", rstTRXMAST!S_TRX_TYPE)
                grdsales.TextMatrix(i, 34) = IIf(IsNull(rstTRXMAST!S_TRX_YEAR), "", rstTRXMAST!S_TRX_YEAR)
                grdsales.TextMatrix(i, 35) = IIf(IsNull(rstTRXMAST!S_VCH_NO), "", rstTRXMAST!S_VCH_NO)
                grdsales.TextMatrix(i, 36) = IIf(IsNull(rstTRXMAST!S_LINE_NO), "", rstTRXMAST!S_LINE_NO)
                grdsales.TextMatrix(i, 37) = IIf(IsNull(rstTRXMAST!INV_DETAILS), "", rstTRXMAST!INV_DETAILS)
                grdsales.TextMatrix(i, 38) = IIf(IsDate(rstTRXMAST!INV_DATE), Format(rstTRXMAST!INV_DATE, "DD/MM/YYYY"), "")
                lbltotalwodiscount.Caption = Format(Val(lbltotalwodiscount.Caption) + Val(grdsales.TextMatrix(i, 13)), ".00")
                'TXTDEALER.Text = IIf(IsNull(rstTRXMAST!VCH_DESC), "", Mid(rstTRXMAST!VCH_DESC, 15))
                TXTINVDATE.Text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                'TXTREMARKS.Text = IIf(IsNull(rstTRXMAST!VCH_DESC), "", Mid(rstTRXMAST!VCH_DESC, 15))
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            lblinvdetails.Caption = ""
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
            If (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                rstTRXMAST.AddNew
                rstTRXMAST!VCH_NO = Val(txtBillNo.Text)
                rstTRXMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                rstTRXMAST!TRX_TYPE = "SR"
                rstTRXMAST.Update
            Else
                If rstTRXMAST!POST_FLAG = "Y" Then lblcredit.Caption = "0" Else lblcredit.Caption = "1"
                TXTREMARKS.Text = IIf(IsNull(rstTRXMAST!REMARKS), "", rstTRXMAST!REMARKS)
                lblinvdetails.Caption = IIf(IsNull(rstTRXMAST!INV_DETAILS), "", rstTRXMAST!INV_DETAILS)
                TXTDEALER.Text = IIf(IsNull(rstTRXMAST!ACT_NAME), "", rstTRXMAST!ACT_NAME)
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing

            ''''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
            'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
            LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")

            TXTSLNO.Text = grdsales.rows
            TXTSLNO.Enabled = True
            txtBillNo.Enabled = False
            FRMEMASTER.Enabled = True
            If i > 0 Or (Val(txtBillNo.Text) < Val(TXTLASTBILL.Text)) Then
                FRMEMASTER.Enabled = True
                FRMECONTROLS.Enabled = True
                cmdRefresh.Enabled = True
                CMDPRINT.Enabled = True
                CmdPrintA5.Enabled = True
                CMDPRINT.SetFocus
            Else
                TXTDEALER.SetFocus
            End If

            Set RSTTRNSMAST = New ADODB.Recordset
            RSTTRNSMAST.Open "Select CHECK_FLAG From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRNSMAST.EOF Or RSTTRNSMAST.BOF) Then
                If RSTTRNSMAST!check_flag = "Y" Then FRMEMASTER.Enabled = False
            End If
            RSTTRNSMAST.Close
            Set RSTTRNSMAST = Nothing
        
    End Select
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
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
    If Val(txtBillNo.Text) = 0 Or Val(txtBillNo.Text) > Val(TXTLASTBILL.Text) Then txtBillNo.Text = TXTLASTBILL.Text
End Sub

Private Sub txtcategory_Change()
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            'PHY.Open "Select * From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '%" & Me.txtcategory.Text & "%') AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_TAX, P_RETAIL, P_WS, P_VAN From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '" & Me.txtcategory.Text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_CODE, ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         Else
             PHY.Close
             'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_TAX, P_RETAIL, P_WS, P_VAN From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '" & Me.txtcategory.Text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_CODE, ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         End If
         
        Set grdtmp.DataSource = PHY
        
        If PHY.RecordCount > 0 Then
            FRMEGRDTMP.Visible = True
        Else
            FRMEGRDTMP.Visible = False
            Exit Sub
        End If
        grdtmp.Columns(0).Caption = "CODE"
        grdtmp.Columns(0).Width = 1500
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(1).Width = 4500
        grdtmp.Columns(2).Caption = "QTY"
        grdtmp.Columns(2).Width = 1100
        grdtmp.Columns(3).Caption = "TAX%"
        grdtmp.Columns(3).Width = 1100
        grdtmp.Columns(4).Caption = "R. RATE"
        grdtmp.Columns(4).Width = 1100
        grdtmp.Columns(5).Caption = "W. RATE"
        grdtmp.Columns(5).Width = 1100
        grdtmp.Columns(6).Caption = "V. RATE"
        grdtmp.Columns(6).Width = 1100
        Exit Sub
ERRHAND:
        MsgBox err.Description
End Sub

Private Sub txtcategory_GotFocus()
    lblretail.Caption = 0
    lblwsale.Caption = 0
    lblvan.Caption = 0
    LBLMRP.Caption = 0
    lblcase.Caption = 0
    lblLWPrice.Caption = 0
    lblcrtnpack.Caption = 1
    
    LBLCOST.Caption = 0
    lblvchno.Caption = ""
    lbltrxtype.Caption = ""
    lbltrxyear.Caption = ""
    lbllineno.Caption = ""
    
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.Text)
    'Call txtcategory_Change
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
            If TXTITEMCODE.Visible = True Then
                TXTITEMCODE.Enabled = True
                TXTITEMCODE.SetFocus
            Else
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            End If
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
            'TXTEXPDATE.Enabled = False
            TXTRATE.SetFocus
        Case vbKeyEscape
            If TXTEXPDATE.Text = "  /  /    " Then GoTo SKIP
            If Not IsDate(TXTEXPDATE.Text) Then Exit Sub
SKIP:
            txtBatch.Enabled = True
            'TXTEXPDATE.Enabled = False
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
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                TXTREMARKS.SetFocus
            End If
        Case vbKeyEscape
            TXTREMARKS.SetFocus
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

Private Sub txtinvnodate_GotFocus()
    txtinvnodate.SelStart = 0
    txtinvnodate.SelLength = Len(txtinvnodate.Text)
End Sub

Private Sub txtinvnodate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'txtinvnodate.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
            'txtinvnodate.Enabled = False
            TXTQTY.Enabled = True
            Call FILL_PREVIIOUSRATE
            'TXTQTY.SetFocus
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
            txtBatch.SetFocus
        Case vbKeyEscape
            If Trim(txtinvnodate.Text) = "" Then
                txtinvnodate.Enabled = True
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

Private Sub txtNetrate_GotFocus()
    'TxtNetrate.Text = Format(Round(Val(TXTPTR.Text) + Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100, 2), "0.00")
    txtNetrate.SelStart = 0
    txtNetrate.SelLength = Len(txtNetrate.Text)
End Sub

Private Sub txtNetrate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TxtNetrate.Text) = 0 Then Exit Sub
            'TXTPTR.Enabled = False
            'TxtNetrate.Enabled = False
            txtPD.Enabled = True
            txtPD.SetFocus
        Case vbKeyEscape
            TxttaxMRP.Enabled = True
            TxttaxMRP.SetFocus
'        Case vbKeyDown
'            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
'            If Val(TXTQTY.Text) = 0 Then Exit Sub
'            If Val(TXTPTR.Text) = 0 Then Exit Sub
'            Call CMDADD_Click
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
    If Val(txtNetrate.Text) <> 0 Then
        txtNetrate.Text = Format(txtNetrate.Text, ".00")
        TXTPTR.Text = Format(Round(Val(txtNetrate.Text) * 100 / (Val(TxttaxMRP.Text) + 100), 3), "0.000")
    End If
    Call TxttaxMRP_LostFocus
    
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
            txtcategory.Enabled = True
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
    On Error GoTo ERRHAND
        If CHANGE_FLAG = True Then Exit Sub
         'If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
         Set grdtmp.DataSource = Nothing
         If PHYFLAG = True Then
            'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_TAX, P_RETAIL, P_WS, P_VAN From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '" & Me.txtcategory.Text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            PHYFLAG = False
         Else
             PHY.Close
             'PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ITEM_CODE Like '" & Me.txtcategory.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_TAX, P_RETAIL, P_WS, P_VAN From ITEMMAST  WHERE (ITEM_CODE Like '" & Me.txtcategory.Text & "%' OR ITEM_NAME Like '" & Me.txtcategory.Text & "%') AND ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
             PHYFLAG = False
         End If
         
        Set grdtmp.DataSource = PHY
        
        If PHY.RecordCount > 0 Then
            FRMEGRDTMP.Visible = True
        Else
            FRMEGRDTMP.Visible = False
            Exit Sub
        End If
        grdtmp.Columns(0).Caption = "CODE"
        grdtmp.Columns(0).Width = 1500
        grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
        grdtmp.Columns(1).Width = 4500
        grdtmp.Columns(2).Caption = "QTY"
        grdtmp.Columns(2).Width = 1100
        grdtmp.Columns(3).Caption = "TAX%"
        grdtmp.Columns(3).Width = 1100
        grdtmp.Columns(4).Caption = "R. RATE"
        grdtmp.Columns(4).Width = 1100
        grdtmp.Columns(5).Caption = "W. RATE"
        grdtmp.Columns(5).Width = 1100
        grdtmp.Columns(6).Caption = "V. RATE"
        grdtmp.Columns(6).Width = 1100
        Exit Sub
ERRHAND:
        MsgBox err.Description
                
End Sub

Private Sub TXTPRODUCT_GotFocus()
    lblretail.Caption = 0
    lblwsale.Caption = 0
    lblvan.Caption = 0
    LBLMRP.Caption = 0
    LBLCOST.Caption = 0
    lblcase.Caption = 0
    lblLWPrice.Caption = 0
    lblcrtnpack.Caption = 1
    
    lblvchno.Caption = ""
    lbltrxtype.Caption = ""
    lbltrxyear.Caption = ""
    lbllineno.Caption = ""
    
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            On Error Resume Next
            grdtmp.SetFocus
        Case vbKeyReturn

            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
'            If Trim(TXTPRODUCT.Text) = "" Then
'                TXTITEMCODE.Enabled = True
'                TXTITEMCODE.SetFocus
'                Exit Sub
'            End If
            CmdDelete.Enabled = False

            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If

            Set grdtmp.DataSource = PHY

            If PHY.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "Sales Return"
                Exit Sub
            End If

            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i
                
                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.Text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
                    On Error GoTo ERRHAND
                    txtBatch.Text = ""
                    TXTEXPIRY.Text = "  /  "
                    TXTRATE.Text = IIf(IsNull(RSTRXFILE!MRP), "", RSTRXFILE!MRP)
                    txtmrpbt.Text = ""
                    Select Case cmbtype.ListIndex
                        Case 1
                            TXTPTR.Text = IIf(IsNull(RSTRXFILE!P_WS), "", RSTRXFILE!P_WS)
                            txtNetrate.Text = IIf(IsNull(RSTRXFILE!P_WS), "", RSTRXFILE!P_WS)
                        Case 2
                            TXTPTR.Text = IIf(IsNull(RSTRXFILE!P_VAN), "", RSTRXFILE!P_VAN)
                            txtNetrate.Text = IIf(IsNull(RSTRXFILE!P_VAN), "", RSTRXFILE!P_VAN)
                        Case 3
                            TXTPTR.Text = IIf(IsNull(RSTRXFILE!MRP), "", RSTRXFILE!MRP)
                            txtNetrate.Text = IIf(IsNull(RSTRXFILE!MRP), "", RSTRXFILE!MRP)
                        Case Else
                            TXTPTR.Text = IIf(IsNull(RSTRXFILE!P_RETAIL), "", RSTRXFILE!P_RETAIL)
                            txtNetrate.Text = IIf(IsNull(RSTRXFILE!P_RETAIL), "", RSTRXFILE!P_RETAIL)
                    End Select
                    TxttaxMRP.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", RSTRXFILE!SALES_TAX)
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                Set grdtmp.DataSource = Nothing
                FRMEGRDTMP.Visible = False
                
'                Set RSTRXFILE = New ADODB.Recordset
'                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ORDER BY CREATE_DATE", db, adOpenStatic, adLockReadOnly
'                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
'                    RSTRXFILE.MoveLast
'                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
'                    Txtpack.Text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
'                    Txtpack.Text = 1
'                    Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "1", RSTRXFILE!LOOSE_PACK)
'                    TxtWarranty.Text = IIf(IsNull(RSTRXFILE!WARRANTY), "", RSTRXFILE!WARRANTY)
'                    On Error Resume Next
'                    CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
'                    CmbWrnty.Text = IIf(IsNull(RSTRXFILE!WARRANTY_TYPE), CmbWrnty.ListIndex = -1, RSTRXFILE!WARRANTY_TYPE)
'                    On Error GoTo eRRhAND
'                Else
'                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
'                    Txtpack.Text = 1
'                    Los_Pack.Text = 1
'                    TxtWarranty.Text = ""
'                    On Error Resume Next
'                    CmbPack.Text = "Nos"
'                    CmbWrnty.ListIndex = -1
'                    On Error GoTo eRRhAND
'
'                    TXTEXPDATE.Text = "  /  /    "
'                    txtBatch.Text = ""
'                    TXTEXPIRY.Text = "  /  "
'                    TXTRATE.Text = ""
'                    txtmrpbt.Text = ""
'                    TXTPTR.Text = ""
'                    txtNetrate.Text = ""
'                    TXTRETAIL.Text = ""
'                    txtWS.Text = ""
'                    txtvanrate.Text = ""
'                    txtcrtn.Text = ""
'                    txtcrtnpack.Text = ""
'                    txtprofit.Text = ""
'                    TxttaxMRP.Text = ""
'                    Los_Pack.Text = "1"
'                    TxtWarranty.Text = ""
'                    On Error Resume Next
'                    CmbPack.Text = "Nos"
'                    CmbWrnty.ListIndex = -1
'                    On Error GoTo eRRhAND
'                    OPTVAT.value = True
'                End If
'                RSTRXFILE.Close
'                Set RSTRXFILE = Nothing

                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    txtcategory.Enabled = False
                    TXTITEMCODE.Enabled = False
                    TxtBarcode.Enabled = False
                    
                    CmbPack.Enabled = True
                    TXTRATE.Enabled = True
                    txtNetrate.Enabled = True
                    TXTPTR.Enabled = True
                    Los_Pack.Enabled = True
                    txtinvnodate.Enabled = True
                    TxtInvoiceDate.Enabled = True
                    txtBatch.Enabled = True
                    TXTEXPIRY.Visible = True
                    TXTEXPDATE.Enabled = True
                    TxttaxMRP.Enabled = True
                    txtPD.Enabled = True
                    Txtgrossamt.Enabled = True
                    TXTQTY.Enabled = True
                    
                    Call FILL_PREVIIOUSRATE
                    'TXTQTY.SetFocus
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
            TXTITEMCODE.Enabled = False
            'TXTPRODUCT.Enabled = False
            txtcategory.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            txtcategory.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
ERRHAND:
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
    TXTPTR.SelLength = Len(TXTPTR.Text)
End Sub

Private Sub TXTPTR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTPTR.Text) = 0 Then Exit Sub
            TxttaxMRP.Enabled = True
            'TxtNetrate.Enabled = False
            'TXTPTR.Enabled = False
            TxttaxMRP.SetFocus
        Case vbKeyEscape
            'TxtNetrate.Enabled = False
            'TXTPTR.Enabled = False
            TXTQTY.Enabled = True
            Call FILL_PREVIIOUSRATE
            'TXTQTY.SetFocus
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
    txtNetrate.Text = Format(Round(Val(TXTPTR.Text) + Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100, 2), "0.00")
    'TXTRETAIL.Text = Round(Val(txtmrpbt.Text) * 0.8, 2)
'    txtretail.Text = Format(Round(Val(TXTRATE.Text) - (Val(txtmrpbt.Text) * 20 / 100), 3), ".000")
'    txtprofit.Text = Format(Round(Val(txtretail.Text) - Val(txtretail.Text) * 10 / 100, 3), ".000")
End Sub

Private Sub TXTQTY_GotFocus()
    FRMEGRDTMP.Visible = False
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
    'Call FILL_PREVIIOUSRATE
    
    On Error GoTo ERRHAND
    lblretail.Caption = ""
    lblwsale.Caption = ""
    lblvan.Caption = ""
    LBLMRP.Caption = ""
    lblcase.Caption = ""
    lblLWPrice.Caption = ""
    lblcrtnpack.Caption = 1
        
    Dim RSTITEMMAST As ADODB.Recordset
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        If Val(lblretail.Caption) = 0 Then lblretail.Caption = IIf(IsNull(RSTITEMMAST!P_RETAIL) Or RSTITEMMAST!P_RETAIL = "", "", RSTITEMMAST!P_RETAIL)
        If Val(lblwsale.Caption) = 0 Then lblwsale.Caption = IIf(IsNull(RSTITEMMAST!P_WS) Or RSTITEMMAST!P_WS = 0, "", RSTITEMMAST!P_WS)
        If Val(lblvan.Caption) = 0 Then lblvan.Caption = IIf(IsNull(RSTITEMMAST!P_VAN) Or RSTITEMMAST!P_VAN = 0, "", RSTITEMMAST!P_VAN)
        If Val(LBLMRP.Caption) = 0 Then LBLMRP.Caption = IIf(IsNull(RSTITEMMAST!MRP) Or RSTITEMMAST!MRP = 0, "", RSTITEMMAST!MRP)
        If Val(LBLCOST.Caption) = 0 Then LBLCOST.Caption = IIf(IsNull(RSTITEMMAST!item_COST) Or RSTITEMMAST!item_COST = 0, "", RSTITEMMAST!item_COST)
        lblcase.Caption = IIf(IsNull(RSTITEMMAST!P_CRTN) Or RSTITEMMAST!P_CRTN = 0, "", RSTITEMMAST!P_CRTN)
        lblLWPrice.Caption = IIf(IsNull(RSTITEMMAST!P_LWS) Or RSTITEMMAST!P_LWS = 0, "", RSTITEMMAST!P_LWS)
        lblcrtnpack.Caption = IIf(IsNull(RSTITEMMAST!CRTN_PACK) Or RSTITEMMAST!CRTN_PACK = 0, 1, RSTITEMMAST!CRTN_PACK)
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            'TXTQTY.Enabled = False
            txtinvnodate.Enabled = True
            'txtBatch.Enabled = True
            If Trim(txtinvnodate.Text) = "" Then
                txtinvnodate.SetFocus
            Else
                txtBatch.Enabled = True
                txtBatch.SetFocus
            End If
        Case vbKeyEscape
'            TXTQTY.Text = ""
'            TXTFREE.Text = ""
'            TxttaxMRP.Text = ""
'            txtprofit.Text = ""
'            txtretail.Text = ""
'            txtWS.Text = ""
'            txtvanrate.Text = ""
'            Txtgrossamt.Text = ""
'            txtcrtn.Text = ""
'            txtcrtnpack.Text = ""
'            txtPD.Text = ""
'            txtBatch.Text = ""
'            TXTRATE.Text = ""
'            txtmrpbt.Text = ""
'            TXTPTR.Text = ""
'            Txtgrossamt.Text = ""
'            LBLSUBTOTAL.Caption = ""
'            lbltaxamount.Caption = ""
            'TXTQTY.Enabled = False
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
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Or MDIMAIN.StatusBar.Panels(14).Text = "Y" Then
                txtNetrate.Enabled = True
                txtNetrate.SetFocus
            Else
                If Val(TxttaxMRP.Text) = 0 Then
                    TxttaxMRP.Enabled = True
                    TxttaxMRP.SetFocus
                Else
                    If MDIMAIN.StatusBar.Panels(14).Text <> "Y" Then
                        TXTPTR.Enabled = True
                        TXTPTR.SetFocus
                    Else
                        If Val(TXTRATE.Text) <> 0 And Val(TXTRATE.Text) = Val(TXTPTR.Text) And MDIMAIN.LBLMRPPLUS.Caption = "Y" Then
                            TXTPTR.Enabled = True
                            TXTPTR.SetFocus
                        Else
                            txtNetrate.Enabled = True
                            txtNetrate.SetFocus
                        End If
                    End If
                End If
            End If
                        
            'TXTRATE.Enabled = False
            'txtNetrate.Enabled = True
            'TXTPTR.Enabled = True
            'TXTPTR.SetFocus
         Case vbKeyEscape
            'TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = True
            TXTEXPDATE.SetFocus
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
    FRMEGRDTMP.Visible = False
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If txtBillNo.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then Exit Sub
            'If TXTINVOICE.Text = "" Then Exit Sub
            FRMECONTROLS.Enabled = True
            
            CmbPack.Enabled = False
            TXTRATE.Enabled = False
            txtNetrate.Enabled = False
            TXTPTR.Enabled = False
            Los_Pack.Enabled = False
            TXTQTY.Enabled = False
            txtinvnodate.Enabled = False
            TxtInvoiceDate.Enabled = False
            txtBatch.Enabled = False
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
            TxttaxMRP.Enabled = False
            txtPD.Enabled = False
            Txtgrossamt.Enabled = False
            
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
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

Private Sub TxtRetailPercent_GotFocus()
    TxtRetailPercent.SelStart = 0
    TxtRetailPercent.SelLength = Len(TxtRetailPercent.Text)
End Sub

Private Sub TxtRetailPercent_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            txtretail.Enabled = False
            TxtRetailPercent.Enabled = False
            txtWS.Enabled = True
            txtWsalePercent.Enabled = True
            txtWS.SetFocus
            'TXTRETAIL.SetFocus
         Case vbKeyEscape
            txtretail.SetFocus
    End Select
End Sub

Private Sub TxtRetailPercent_LostFocus()
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
    End If
    txtretail.Text = Round((Val(TXTPTR.Tag) * Val(TxtRetailPercent.Text) / 100) + Val(TXTPTR.Tag), 2)
    'txtretail.Text = Round(((Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) + Val(TXTPTR.Text)) * Val(TxtRetailPercent.Text) / 100 + ((Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) + Val(TXTPTR.Text)), 2)
    txtretail.Text = Format(Val(txtretail.Text), "0.0000")
End Sub

Private Sub txtSchPercent_GotFocus()
    txtSchPercent.SelStart = 0
    txtSchPercent.SelLength = Len(txtSchPercent.Text)
End Sub

Private Sub txtSchPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtvanrate.Enabled = False
            txtSchPercent.Enabled = False
            'Frame1.Enabled = True
         Case vbKeyEscape
            txtvanrate.SetFocus
    End Select
End Sub

Private Sub txtSchPercent_LostFocus()
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
    End If
    txtvanrate.Text = Round((Val(TXTPTR.Tag) * Val(txtSchPercent.Text) / 100) + Val(TXTPTR.Tag), 2)
    txtvanrate.Text = Format(Val(txtvanrate.Text), "0.000")
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.Text) = 0 Then
                TXTSLNO.Text = grdsales.rows
                CmdDelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.Text) >= grdsales.rows Then
                TXTSLNO.Text = grdsales.rows
                CmdDelete.Enabled = False
                CMDMODIFY.Enabled = False
            End If

            If Val(TXTSLNO.Text) < grdsales.rows Then
                TXTSLNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 0)
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 1)
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))
                TXTUNIT.Text = 1 'grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                Txtpack.Text = 1 'grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                LBLCOST.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
                TXTRATE.Text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)), 2), "0.000")
                TXTPTR.Text = Format(Round(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 9)) * Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5)), 2), "0.000")
                txtprofit.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 7)), "0.00")
                txtretail.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 18)), "0.00")
                txtWS.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 19)), "0.00")
                txtvanrate.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 25)), "0.00")
                Txtgrossamt.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 26)), "0.00")
                txtcrtn.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 20)), "0.00")
                txtcrtnpack.Text = Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 24)), "0.00")
                'TXTPTR.Text = Format((Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))) * Val(Los_Pack.Text), "0.000")

                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 11)
                TXTEXPDATE.Text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 12)), grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "  /  /    ")
                TXTEXPIRY.Text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 12)), Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 12), "mm/yy"), "  /  ")
                'LBLSUBTOTAL.Caption = Format(Val(TXTQTY.Text) * (Val(TXTPTR.Text) + Val(lbltaxamount.Caption)), ".000")
                If OptDiscAmt.Value = True Then
                    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption) - Val(txtPD.Text), ".000")
                Else
                    LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100), ".000")
                End If
                TXTFREE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 14)
                TxttaxMRP.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105 '(100 + Val(TxttaxMRP.Text))
                txtPD.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 17))
                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) = "V" Then
                    OPTVAT.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 15)) = "M" Then
                    OPTTaxMRP.Value = True
                Else
                    OPTVAT.Value = True
                End If

                If Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 27)) = "P" Then
                    optdiscper.Value = True
                ElseIf Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 27)) = "A" Then
                    OptDiscAmt.Value = True
                End If
                On Error Resume Next
                Los_Pack.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 28)
                CmbPack.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 29)
                TxtWarranty.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 30)
                CmbWrnty.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 31)
                
                LBLCOST.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
                lbltrxtype.Caption = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 33))
                lbltrxyear.Caption = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 34))
                lblvchno.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 35))
                lbllineno.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 36))
                txtinvnodate.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 37))
                TxtInvoiceDate.Text = IIf(IsDate(grdsales.TextMatrix(Val(TXTSLNO.Text), 38)), grdsales.TextMatrix(Val(TXTSLNO.Text), 38), "  /  /    ")
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                txtcategory.Enabled = False
                TXTITEMCODE.Enabled = False
                TxtBarcode.Enabled = False
                TXTQTY.Enabled = False
                TXTRATE.Enabled = False
                TXTEXPDATE.Enabled = False
                txtBatch.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Set grdtmp.DataSource = Nothing
                FRMEGRDTMP.Visible = False
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            TXTPRODUCT.Enabled = True
            txtcategory.Enabled = True
            txtcategory.SetFocus
            'TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TXTQTY.Text = ""
                Txtpack.Text = 1 '""
                Los_Pack.Text = ""
                CmbPack.ListIndex = -1
                TxtWarranty.Text = ""
                CmbWrnty.ListIndex = -1
                TXTFREE.Text = ""
                TxttaxMRP.Text = ""
                txtPD.Text = ""
                txtprofit.Text = ""
                txtretail.Text = ""
                TxtRetailPercent.Text = ""
                txtWsalePercent.Text = ""
                txtSchPercent.Text = ""
                txtWS.Text = ""
                txtvanrate.Text = ""
                Txtgrossamt.Text = ""
                txtcrtn.Text = ""
                txtcrtnpack.Text = ""
                TXTRATE.Text = ""
                TxtComAmt.Text = ""
                TxtComper.Text = ""
                txtmrpbt.Text = ""
                LBLSUBTOTAL.Caption = ""
                lbltaxamount.Caption = ""
                TXTEXPDATE.Text = "  /  /    "
                TXTEXPIRY.Text = "  /  "
                txtBatch.Text = ""
                txtinvnodate.Text = ""
                TxtInvoiceDate.Text = "  /  /    "
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            Else
                cmdRefresh.Enabled = True
                CMDPRINT.Enabled = True
                CmdPrintA5.Enabled = True
                CMDPRINT.SetFocus
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
            'TXTEXPDATE.Enabled = False
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            txtBatch.Enabled = True
            'TXTEXPDATE.Enabled = False
            txtBatch.SetFocus

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
            If Val(TxttaxMRP.Text) <> 0 And optnet.Value = True Then
                OPTVAT.Value = True
                OPTVAT.SetFocus
                Exit Sub
            End If
            'TxttaxMRP.Enabled = False
            txtPD.Enabled = True
            txtPD.SetFocus
         Case vbKeyEscape
            'TxttaxMRP.Enabled = False
            txtNetrate.Enabled = True
            TXTPTR.Enabled = True
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
    Txtgrossamt.Text = Val(TXTPTR.Text) * Val(TXTQTY.Text)
    txtmrpbt.Text = 100 * Val(TXTRATE.Text) / (100 + Val(TxttaxMRP.Text))
    If Val(TxttaxMRP.Text) = 0 Then
        TxttaxMRP.Text = 0
        lbltaxamount.Caption = 0
        lbltaxamount.Caption = ""
        If optdiscper.Value = True Then
            LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.Text)) - Val(Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100), ".000")
        Else
            LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text) - Val(txtPD.Text), ".000")
        End If

    Else
        If OPTTaxMRP.Value = True Then
            If optdiscper.Value = True Then
                lbltaxamount.Caption = (Val(txtmrpbt.Text) - (Val(TXTRATE.Text) * Val(txtPD.Text) / 100)) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
                LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
            Else
                lbltaxamount.Caption = (Val(txtmrpbt.Text) - Val(txtPD.Text)) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
                LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
            End If
        ElseIf OPTVAT.Value = True Then
            'lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text) + Val(TxtFree.Text))
            'lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text))

            If optdiscper.Value = True Then
                lbltaxamount.Caption = Round((Val(Txtgrossamt.Text) - (Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
                LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100), ".000")
            Else
                lbltaxamount.Caption = Round((Val(Txtgrossamt.Text) - Val(txtPD.Text)) * Val(TxttaxMRP.Text) / 100, 2)
                LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption) - Val(txtPD.Text), ".000")
            End If
            'LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.Text)) + Val(lbltaxamount.Caption), ".000")
        Else
            lbltaxamount.Caption = ""
            LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text), ".000")
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
            Txtpack.Enabled = True
            Txtpack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTQTY.Text = ""
            TXTFREE.Text = ""
            TxttaxMRP.Text = ""
            txtprofit.Text = ""
            txtretail.Text = ""
            TxtRetailPercent.Text = ""
            txtWsalePercent.Text = ""
            txtSchPercent.Text = ""
            txtWS.Text = ""
            txtvanrate.Text = ""
            Txtgrossamt.Text = ""
            txtcrtn.Text = ""
            txtcrtnpack.Text = ""
            txtPD.Text = ""
            txtBatch.Text = ""
            txtinvnodate.Text = ""
            TxtInvoiceDate.Text = "  /  /    "
            TXTRATE.Text = ""
            txtmrpbt.Text = ""
            TXTPTR.Text = ""
            txtNetrate.Text = ""
            Txtgrossamt.Text = ""
            TXTEXPDATE.Text = "  /  /    "
            TXTEXPIRY.Text = "  /  "
            LBLSUBTOTAL.Caption = ""
            lbltaxamount.Caption = ""
            TXTPRODUCT.Enabled = True
            txtcategory.Enabled = True
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

    On Error GoTo ERRHAND
    If (TXTDISCAMOUNT.Text = "") Then
        DISC = 0
    Else
        DISC = TXTDISCAMOUNT.Text
    End If
    If grdsales.rows = 1 Then
        TXTDISCAMOUNT.Text = "0"
    ElseIf Val(TXTDISCAMOUNT.Text) > Val(lbltotalwodiscount.Caption) Then
        MsgBox "Discount Amount More than Bill Amount", , "SALES..."
        TXTDISCAMOUNT.SelStart = 0
        TXTDISCAMOUNT.SelLength = Len(TXTDISCAMOUNT.Text)
        TXTDISCAMOUNT.SetFocus
        Exit Sub
    End If
    TXTDISCAMOUNT.Text = Format(TXTDISCAMOUNT.Text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text)) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    ''LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) - Val(TXTDISCAMOUNT.Text), 0), ".00")
    Exit Sub
ERRHAND:
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
            'If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Public Sub appendpurchase()

    Dim rstMaxRec As ADODB.Recordset
    Dim RSTLINK As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    Dim TRXTYPE, INVTRXTYPE, INVTYPE As String
    Dim RECNO, INVNO As Long
    
    Dim M_DATA As Double
    Dim i As Long

    'On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass

    db.Execute "delete From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'SR' AND INV_TRX_TYPE = 'SR'"
    db.Execute "delete FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE= 'DB' AND INV_TRX_TYPE = 'DN'"
    
    lblinvdetails.Caption = ""
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT DISTINCT INV_DETAILS FROM RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTITEMMAST.EOF
        lblinvdetails.Caption = lblinvdetails.Caption & IIf(IsNull(RSTITEMMAST!INV_DETAILS) Or RSTITEMMAST!INV_DETAILS = "", "", RSTITEMMAST!INV_DETAILS) & ", "
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    lblinvdetails.Caption = Trim(lblinvdetails.Caption)
    If Len(lblinvdetails.Caption) > 2 And Right(lblinvdetails.Caption, 1) = "," Then
        lblinvdetails.Caption = Mid(lblinvdetails.Caption, 1, Len(lblinvdetails.Caption) - 1)
        lblinvdetails.Caption = Trim(lblinvdetails.Caption)
    End If
    lblinvdetails.Caption = Left(Trim(lblinvdetails.Caption), 200)
    
    If grdsales.rows = 1 Then
        db.Execute "delete FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE= 'SR' AND INV_TRX_TYPE = 'SR'"
        GoTo SKIP
    End If
    
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
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'SR' AND INV_TRX_TYPE = 'SR'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!REC_NO = i + 1
            RSTITEMMAST!INV_TYPE = "SR"
            RSTITEMMAST!INV_TRX_TYPE = "SR"
            RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        End If
        RSTITEMMAST!TRX_TYPE = "DR"
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = Trim(DataList2.Text)
        RSTITEMMAST!AMOUNT = Val(LBLTOTAL.Caption)
        RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
        RSTITEMMAST!check_flag = "P"
        RSTITEMMAST!REMARKS = "Cash Returned"
        
        RECNO = RSTITEMMAST!REC_NO
        INVNO = RSTITEMMAST!INV_NO
        TRXTYPE = RSTITEMMAST!TRX_TYPE
        INVTRXTYPE = RSTITEMMAST!INV_TRX_TYPE
        INVTYPE = RSTITEMMAST!INV_TYPE
        
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "SR"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = Trim(DataList2.Text)
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!DISCOUNT = Val(TXTDISCAMOUNT.Text)
        RSTTRXFILE!ADD_AMOUNT = Val(txtaddlamt.Text)
        RSTTRXFILE!ROUNDED_OFF = 0
        RSTTRXFILE!OPEN_PAY = 0
        RSTTRXFILE!PAY_AMOUNT = 0
        RSTTRXFILE!REF_NO = ""
        RSTTRXFILE!SLSM_CODE = "CS"
        RSTTRXFILE!check_flag = "N"
        If lblcredit.Caption = "0" Then RSTTRXFILE!POST_FLAG = "Y" Else RSTTRXFILE!POST_FLAG = "N"
        RSTTRXFILE!CFORM_NO = ""
        RSTTRXFILE!TIN = Trim(TXTTIN.Text)
        RSTTRXFILE!CFORM_DATE = Date
        RSTTRXFILE!REMARKS = Trim(TXTREMARKS.Text)
        RSTTRXFILE!INV_DETAILS = Trim(lblinvdetails.Caption)
        RSTTRXFILE!DISC_PERS = Val(txtcramt.Text)
        RSTTRXFILE!CST_PER = Val(TxtCST.Text)
        RSTTRXFILE!INS_PER = Val(TxtInsurance.Text)
        RSTTRXFILE!LETTER_NO = 0
        RSTTRXFILE!LETTER_DATE = Date
        RSTTRXFILE!INV_MSGS = ""
        RSTTRXFILE!CREATE_DATE = Format(TXTDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!MODIFY_DATE = Format(TXTDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    i = 0
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(CR_NO) From DBTPYMT", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing

    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE = 'SR' AND INV_TRX_TYPE = 'SR'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!TRX_TYPE = "SR"
        RSTITEMMAST!INV_TRX_TYPE = "SR"
        RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTITEMMAST!CR_NO = i
        RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
        RSTITEMMAST!INV_AMT = 0
    End If
    RSTITEMMAST!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTITEMMAST!RCPT_AMT = Val(LBLTOTAL.Caption)
    If lblcredit.Caption = "0" Then
        RSTITEMMAST!check_flag = "Y"
        RSTITEMMAST!BAL_AMT = 0
    Else
        RSTITEMMAST!check_flag = "N"
        RSTITEMMAST!BAL_AMT = Val(LBLTOTAL.Caption) - RSTITEMMAST!RCPT_AMT
    End If
    RSTITEMMAST!REF_NO = Trim(TXTREMARKS.Text)
    RSTITEMMAST!ACT_CODE = DataList2.BoundText
    RSTITEMMAST!ACT_NAME = DataList2.Text
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    If lblcredit.Caption = "0" Then
        Set rstMaxNo = New ADODB.Recordset
        rstMaxNo.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'DB'", db, adOpenForwardOnly
        If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
            i = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        End If
        rstMaxNo.Close
        Set rstMaxNo = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "DB"
        RSTTRXFILE!INV_TRX_TYPE = "DN"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!CR_NO = i
        RSTTRXFILE!INV_NO = Val(txtBillNo.Text)
        RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!RCPT_AMT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.Text
        RSTTRXFILE!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!REF_NO = "Cash Returned"
        RSTTRXFILE!INV_AMT = Null
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!BANK_FLAG = "N"
        RSTTRXFILE!B_TRX_TYPE = Null
        'RSTTRXFILE!B_TRX_NO = Null
        RSTTRXFILE!B_BILL_TRX_TYPE = Null
        RSTTRXFILE!B_TRX_YEAR = Null
        RSTTRXFILE!BANK_CODE = Null
        RSTTRXFILE!C_TRX_TYPE = TRXTYPE
        RSTTRXFILE!C_REC_NO = RECNO
        RSTTRXFILE!C_INV_TRX_TYPE = INVTRXTYPE
        RSTTRXFILE!C_INV_TYPE = INVTYPE
        RSTTRXFILE!C_INV_NO = INVNO
        
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!VCH_DATE = Format(Trim(TXTINVDATE.Text), "dd/mm/yyyy")
        RSTTRXFILE!VCH_DESC = "Received From " & Mid(DataList2.Text, 1, 80)
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

SKIP:
    Set rstMaxNo = New ADODB.Recordset
    'rstMaxNo.Open "Select MAX(VCH_NO) From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= 'SR'", db, adOpenStatic, adLockReadOnly
    rstMaxNo.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' ", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    FRMEGRDTMP.Visible = False
    grdsales.rows = 1
    TXTSLNO.Text = 1
    cmdRefresh.Enabled = False
    CMDPRINT.Enabled = False
    CmdPrintA5.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.Text = TXTLASTBILL.Text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTREMARKS.Text = ""
    lblinvdetails.Caption = ""
    TXTSLNO.Text = ""
    TXTITEMCODE.Text = ""
    TXTPRODUCT.Text = ""
    TXTQTY.Text = ""
    Txtpack.Text = 1 '""
    Los_Pack.Text = ""
    CmbPack.ListIndex = -1
    TxtWarranty.Text = ""
    CmbWrnty.ListIndex = -1
    TXTFREE.Text = ""
    TxttaxMRP.Text = ""
    txtPD.Text = ""
    txtprofit.Text = ""
    txtretail.Text = ""
    TxtRetailPercent.Text = ""
    txtWsalePercent.Text = ""
    txtSchPercent.Text = ""
    txtWS.Text = ""
    txtvanrate.Text = ""
    Txtgrossamt.Text = ""
    txtcrtn.Text = ""
    txtcrtnpack.Text = ""
    txtBatch.Text = ""
    txtinvnodate.Text = ""
    TxtInvoiceDate.Text = "  /  /    "
    TXTRATE.Text = ""
    txtmrpbt.Text = ""
    TXTPTR.Text = ""
    txtNetrate.Text = ""
    Txtgrossamt.Text = ""
    TXTEXPDATE.Text = "  /  /    "
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    lbltaxamount.Caption = ""
    txtaddlamt.Text = ""
    txtcramt.Text = ""
    TxtInsurance.Text = ""
    txtcategory.Text = ""
    TxtCST.Text = ""
    LBLTOTAL.Caption = ""
    lbltotalwodiscount.Caption = ""
    TXTDISCAMOUNT.Text = ""
    lblcredit.Caption = "1"
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
    txtBillNo.Visible = True
    Chkcancel.Value = 0
    Set grdtmp.DataSource = Nothing
    FRMEGRDTMP.Visible = False
    Screen.MousePointer = vbNormal
    '''MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "Sales Return ENTRY"
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = 7 Then
        MsgBox "Select Customer from the list", vbOKOnly, "Sales Return"
    Else
        MsgBox err.Description
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
            'If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
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

    On Error GoTo ERRHAND
    If (txtaddlamt.Text = "") Then
        DISC = 0
    Else
        DISC = txtaddlamt.Text
    End If
    If grdsales.rows = 1 Then
        txtaddlamt.Text = "0"
    ElseIf Val(txtaddlamt.Text) > Val(lbltotalwodiscount.Caption) Then
        MsgBox "Discount Amount More than Bill Amount", , "Sales Return..."
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
ERRHAND:
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
            'If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
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

    On Error GoTo ERRHAND
    If (txtcramt.Text = "") Then
        DISC = 0
    Else
        DISC = txtcramt.Text
    End If
    If grdsales.rows = 1 Then
        txtcramt.Text = "0"
    ElseIf Val(txtcramt.Text) > Val(lbltotalwodiscount.Caption) Then
        MsgBox "Credit Note Amount More than Bill Amount", , "Sales Return..."
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
ERRHAND:
    MsgBox "Please enter a Numeric Value", , "Cr. Note.."
    txtcramt.SetFocus
End Sub

Private Sub OPTTaxMRP_GotFocus()
    If optdiscper.Value = True Then
        lbltaxamount.Caption = (Val(txtmrpbt.Text) - (Val(TXTRATE.Text) * Val(txtPD.Text) / 100)) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")

        'lbltaxamount.Caption = Val(txtmrpbt.Text) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
        'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
    Else
        lbltaxamount.Caption = (Val(txtmrpbt.Text) - Val(txtPD.Text)) * (Val(TXTQTY.Text)) * Val(TxttaxMRP.Text) / 100
        LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Val(TXTPTR.Text)) + Val(lbltaxamount.Caption), ".000")
    End If
End Sub

Private Sub OPTVAT_GotFocus()
    'lbltaxamount.Caption = (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) * (Val(TXTQTY.Text) + Val(TxtFree.Text))
    If optdiscper.Value = True Then
        lbltaxamount.Caption = Round((Val(Txtgrossamt.Text) - (Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100)) * Val(TxttaxMRP.Text) / 100, 2)
        LBLSUBTOTAL.Caption = Format((Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption)) - Val(Val(Txtgrossamt.Text) * Val(txtPD.Text) / 100), ".000")
    Else
        lbltaxamount.Caption = Round((Val(Txtgrossamt.Text) - Val(txtPD.Text)) * Val(TxttaxMRP.Text) / 100, 2)
        LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text) + Val(lbltaxamount.Caption) - Val(txtPD.Text), ".000")
    End If
End Sub

Private Sub OPTNET_GotFocus()
    lbltaxamount.Caption = ""
    LBLSUBTOTAL.Caption = Format(Val(Txtgrossamt.Text), ".000")
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
            'txtPD.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            'txtPD.Enabled = False
            TxttaxMRP.Enabled = True
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

    If Val(txtretail.Text) = 0 Then txtretail.Text = Format(Round(Val(txtmrpbt.Text) - (Val(txtmrpbt.Text) * 20 / 100), 3), ".000")
    If Val(txtWS.Text) = 0 Then txtWS.Text = Format(Round(Val(txtretail.Text) - (Val(txtretail.Text) * 10 / 100), 3), ".000")

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
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_NAME Like '" & Me.TXTDEALER.Text & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_NAME Like '" & Me.TXTDEALER.Text & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
    FRMEGRDTMP.Visible = False
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
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
    On Error GoTo ERRHAND
    Dim rstCustomer As ADODB.Recordset
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        'If TxtBillName.Text = "" Then TxtBillName.Text = DataList2.Text
        'TxtBillName.Text = DataList2.Text
        TxtBillAddress.Text = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
        'TxtBillAddress.Text = IIf(IsNull(rstCustomer!ADDRESS), "", Trim(rstCustomer!ADDRESS))
        TXTTIN.Text = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
        TxtPhone.Text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO))
        TXTAREA.Text = IIf(IsNull(rstCustomer!Area), "", Trim(rstCustomer!Area))
        txtPin.Text = IIf(IsNull(rstCustomer!PINCODE), "", Trim(rstCustomer!PINCODE))
        lblIGST.Caption = IIf(IsNull(rstCustomer!CUST_IGST), "N", rstCustomer!CUST_IGST)
        'TXTAREA.Text = IIf(IsNull(rstCustomer!Area), "", Trim(rstCustomer!Area))
        'TxtDL1.Text = IIf(IsNull(rstCustomer!DL_NO), "", Trim(rstCustomer!DL_NO))
        'TxtDL2.Text = IIf(IsNull(rstCustomer!REMARKS), "", Trim(rstCustomer!REMARKS))
        'TxtCST.Text = IIf(IsNull(rstCustomer!CST), "", Trim(rstCustomer!CST))
        Select Case rstCustomer!Type
            Case "W"
                cmbtype.ListIndex = 1
                'TXTTYPE.Text = 2
            Case "V"
                cmbtype.ListIndex = 2
                'TXTTYPE.Text = 3
            Case "M"
                cmbtype.ListIndex = 3
                'TXTTYPE.Text = 4
            Case Else
                cmbtype.ListIndex = 0
                'TXTTYPE.Text = 1
        End Select
    Else
        lblIGST.Caption = ""
        TxtBillAddress = ""
        TXTTIN.Text = ""
        TxtPhone.Text = ""
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Sales Return..."
                DataList2.SetFocus
                Exit Sub
            End If
            TXTINVDATE.SetFocus
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
    FRMEGRDTMP.Visible = False
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
            If Val(txtretail.Text) = 0 Then
                TxtRetailPercent.SetFocus
            Else
                txtretail.Enabled = False
                TxtRetailPercent.Enabled = False
                txtWS.Enabled = True
                txtWsalePercent.Enabled = True
                txtWS.SetFocus
            End If
            Exit Sub
            If Val(txtretail.Text) = 0 Then
                TxtRetailPercent.SetFocus
                Exit Sub
            End If
            txtretail.Enabled = False
            TxtRetailPercent.Enabled = False
            'cmdadd.Enabled = True
            'cmdadd.SetFocus
            txtWS.Enabled = True
            txtWsalePercent.Enabled = True
            txtWS.SetFocus
         Case vbKeyEscape
            txtretail.Enabled = False
            TxtRetailPercent.Enabled = False
            txtPD.Enabled = True
            txtPD.SetFocus
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
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
    End If
    TxtRetailPercent.Text = Round(((Val(txtretail.Text) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
    TxtRetailPercent.Text = Format(Val(TxtRetailPercent.Text), "0.00")
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
                txtWS.Enabled = False
                txtWsalePercent.Enabled = False
                cmdadd.Enabled = True
                cmdadd.SetFocus
            End If
            Exit Sub
            If Val(txtWS.Text) = 0 Then
                txtWsalePercent.SetFocus
                Exit Sub
            End If
            txtWS.Enabled = False
            txtWsalePercent.Enabled = False
         Case vbKeyEscape
            txtWS.Enabled = False
            txtWsalePercent.Enabled = False
            txtretail.Enabled = True
            TxtRetailPercent.Enabled = True
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
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
    End If
    txtWsalePercent.Text = Round(((Val(txtWS.Text) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
    txtWsalePercent.Text = Format(Val(txtWsalePercent.Text), "0.00")
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
            txtSchPercent.Enabled = True
            txtWS.SetFocus
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

Private Sub txtcrtnpack_GotFocus()
    If Val(txtcrtn.Text) = 0 Then txtcrtnpack.Text = 0
    txtcrtnpack.SelStart = 0
    txtcrtnpack.SelLength = Len(txtcrtnpack.Text)
End Sub

Private Sub txtcrtnpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtcrtn.Text) <> 0 And Val(txtcrtnpack.Text) = 0 Then
                MsgBox "Please enter the Pack Qty for Carton", vbOKOnly, "Sales Return"
                txtcrtnpack.SetFocus
                Exit Sub
            End If
            If Val(txtcrtn.Text) = 0 And Val(txtcrtnpack.Text) <> 0 Then
                MsgBox "Please enter the Rate for Carton", vbOKOnly, "Sales Return"
                txtcrtnpack.Enabled = False
                txtcrtn.Enabled = True
                txtcrtn.SetFocus
                Exit Sub
            End If
            txtcrtnpack.Enabled = False
            
         Case vbKeyEscape
            txtcrtnpack.Enabled = False
            txtcrtn.Enabled = True
            txtcrtn.SetFocus
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
                txtvanrate.Enabled = False
                txtSchPercent.Enabled = False
                TxtWarranty.Enabled = True
                CmbWrnty.Enabled = True
                TxtWarranty.SetFocus
            End If
            Exit Sub
            If Val(txtvanrate.Text) = 0 Then
                txtSchPercent.SetFocus
                Exit Sub
            End If
            txtvanrate.Enabled = False
            txtSchPercent.Enabled = False
            TxtWarranty.Enabled = True
            CmbWrnty.Enabled = True
            TxtWarranty.SetFocus
            'txtcrtn.Enabled = True
            'txtcrtn.SetFocus
         Case vbKeyEscape
            txtvanrate.Enabled = False
            txtSchPercent.Enabled = False
            txtWS.Enabled = True
            txtWsalePercent.Enabled = True
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
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
    End If
    txtSchPercent.Text = Round(((Val(txtvanrate.Text) - Val(TXTPTR.Tag)) * 100) / Val(TXTPTR.Tag), 2)
    txtSchPercent.Text = Format(Val(txtSchPercent.Text), "0.00")
End Sub

Private Sub Txtgrossamt_GotFocus()
    Txtgrossamt.SelStart = 0
    Txtgrossamt.SelLength = Len(Txtgrossamt.Text)
End Sub

Private Sub Txtgrossamt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(Txtgrossamt.Text) = 0 Then Exit Sub
            TXTRATE.Enabled = True
            Txtgrossamt.Enabled = False
            TXTRATE.SetFocus
        Case vbKeyEscape
            Txtgrossamt.Enabled = False
            TXTQTY.Enabled = True
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
        TXTPTR.Text = Format(Round(Val(Txtgrossamt.Text) / Val(TXTQTY.Text), 1), "0.00")
    End If
    Call TxttaxMRP_LostFocus
End Sub

Function FILL_PREVIIOUSRATE()
    
    Set GRDPRERATE.DataSource = Nothing
    If Trim(TxtBarcode.Text) <> "" Then
        If PRERATE_FLAG = True Then
            PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, P_RETAIL, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, TRX_YEAR, VCH_NO, LINE_NO, LINE_DISC, BARCODE, LOOSE_PACK From TRXFILE  WHERE BARCODE = '" & TxtBarcode.Text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SV' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
            PRERATE_FLAG = False
        Else
            PHY_PRERATE.Close
            PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, P_RETAIL, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, TRX_YEAR, VCH_NO, LINE_NO, LINE_DISC, BARCODE, LOOSE_PACK From TRXFILE  WHERE BARCODE = '" & TxtBarcode.Text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SV' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
            PRERATE_FLAG = False
        End If
        If PHY_PRERATE.RecordCount = 0 Then
            If PRERATE_FLAG = True Then
                PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, P_RETAIL, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, TRX_YEAR, VCH_NO, LINE_NO, LINE_DISC, BARCODE, LOOSE_PACK From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SV' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
                PRERATE_FLAG = False
            Else
                PHY_PRERATE.Close
                PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, P_RETAIL, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, TRX_YEAR, VCH_NO, LINE_NO, LINE_DISC, BARCODE, LOOSE_PACK From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SV' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
                PRERATE_FLAG = False
            End If
        End If
    Else
        If PRERATE_FLAG = True Then
            PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, P_RETAIL, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, TRX_YEAR, VCH_NO, LINE_NO, LINE_DISC, BARCODE, LOOSE_PACK From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SV' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
            PRERATE_FLAG = False
        Else
            PHY_PRERATE.Close
            PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, P_RETAIL, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, TRX_YEAR, VCH_NO, LINE_NO, LINE_DISC, BARCODE, LOOSE_PACK From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SV' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
            PRERATE_FLAG = False
        End If
    End If
    
    If PHY_PRERATE.RecordCount > 0 Then
        'Fram.Enabled = False
        fRMEPRERATE.Visible = True
        Set GRDPRERATE.DataSource = PHY_PRERATE
        
        GRDPRERATE.Columns(0).Caption = "TYPE"
        GRDPRERATE.Columns(1).Caption = "ITEM CODE"
        GRDPRERATE.Columns(2).Caption = "ITEM NAME"
        GRDPRERATE.Columns(3).Caption = "BILL DATE"
        GRDPRERATE.Columns(4).Caption = "SOLD QTY"
        GRDPRERATE.Columns(5).Caption = "MRP"
        GRDPRERATE.Columns(6).Caption = "RATE"
        GRDPRERATE.Columns(7).Caption = "NET RATE"
        GRDPRERATE.Columns(8).Caption = "TAX"
        GRDPRERATE.Columns(9).Caption = "TAX MODE"
        GRDPRERATE.Columns(10).Caption = "COST"
        GRDPRERATE.Columns(11).Caption = "BATCH"
        GRDPRERATE.Columns(12).Caption = "EXPIRY"
        GRDPRERATE.Columns(13).Caption = "YEAR"
        GRDPRERATE.Columns(14).Caption = "BILL NO"
        GRDPRERATE.Columns(15).Caption = "LINE NO"
        GRDPRERATE.Columns(16).Caption = "Disc%"

        GRDPRERATE.Columns(0).Visible = False
        GRDPRERATE.Columns(1).Visible = False
        GRDPRERATE.Columns(2).Width = 3500
        GRDPRERATE.Columns(3).Width = 1200
        GRDPRERATE.Columns(4).Width = 1000
        GRDPRERATE.Columns(5).Width = 1000
        GRDPRERATE.Columns(6).Width = 1000
        GRDPRERATE.Columns(7).Width = 1000
        GRDPRERATE.Columns(8).Width = 800
        GRDPRERATE.Columns(9).Width = 0
        GRDPRERATE.Columns(10).Width = 1000
        GRDPRERATE.Columns(11).Width = 1100
        GRDPRERATE.Columns(12).Width = 1100
        GRDPRERATE.Columns(13).Width = 0
        GRDPRERATE.Columns(14).Width = 1000
        GRDPRERATE.Columns(15).Width = 0
        GRDPRERATE.Columns(16).Width = 1000



        'GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(2).Text
        GRDPRERATE.SetFocus
    Else
'        Set GRDPRERATE.DataSource = Nothing
'
'        If PRERATE_FLAG = True Then
'            PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, P_RETAIL, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, TRX_YEAR, VCH_NO, LINE_NO  From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SV' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
'            PRERATE_FLAG = False
'        Else
'            PHY_PRERATE.Close
'            PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, MRP, P_RETAILWOTAX, P_RETAIL, SALES_TAX, CHECK_FLAG, ITEM_COST, REF_NO, EXP_DATE, TRX_YEAR, VCH_NO, LINE_NO  From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI' OR TRX_TYPE = 'SV' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') ORDER BY TRX_YEAR DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
'            PRERATE_FLAG = False
'        End If
'
'        If PHY_PRERATE.RecordCount > 0 Then
'            'Fram.Enabled = False
'            fRMEPRERATE.Visible = True
'            Set GRDPRERATE.DataSource = PHY_PRERATE
'
'            GRDPRERATE.Columns(0).Caption = "TYPE"
'            GRDPRERATE.Columns(1).Caption = "ITEM CODE"
'            GRDPRERATE.Columns(2).Caption = "ITEM NAME"
'            GRDPRERATE.Columns(3).Caption = "BILL DATE"
'            GRDPRERATE.Columns(4).Caption = "SOLD QTY"
'            GRDPRERATE.Columns(5).Caption = "MRP"
'            GRDPRERATE.Columns(6).Caption = "RATE"
'            GRDPRERATE.Columns(7).Caption = "NET RATE"
'            GRDPRERATE.Columns(8).Caption = "TAX"
'            GRDPRERATE.Columns(9).Caption = "TAX MODE"
'            GRDPRERATE.Columns(10).Caption = "COST"
'            GRDPRERATE.Columns(11).Caption = "BATCH"
'            GRDPRERATE.Columns(12).Caption = "EXPIRY"
'            GRDPRERATE.Columns(13).Caption = "YEAR"
'            GRDPRERATE.Columns(14).Caption = "VCH NO"
'            GRDPRERATE.Columns(15).Caption = "LINE NO"
'
'            GRDPRERATE.Columns(0).Visible = False
'            GRDPRERATE.Columns(1).Visible = False
'            GRDPRERATE.Columns(2).Width = 3500
'            GRDPRERATE.Columns(3).Width = 1400
'            GRDPRERATE.Columns(4).Width = 1200
'            GRDPRERATE.Columns(5).Width = 1200
'            GRDPRERATE.Columns(6).Width = 1200
'            GRDPRERATE.Columns(7).Width = 1200
'            GRDPRERATE.Columns(8).Width = 1200
'            GRDPRERATE.Columns(9).Width = 1300
'            GRDPRERATE.Columns(10).Width = 1300
'            GRDPRERATE.Columns(11).Width = 1300
'            GRDPRERATE.Columns(12).Width = 1300
'            GRDPRERATE.Columns(13).Width = 500
'            GRDPRERATE.Columns(14).Width = 500
'            GRDPRERATE.Columns(15).Width = 500
'
'
'            'GRDPRERATE.SetFocus
'            LBLHEAD(2).Caption = GRDPRERATE.Columns(2).Text
'            GRDPRERATE.SetFocus
'        Else
'            'If MsgBox("This Item has not been sold to " & DataList2.Text & " Yet!! Do You Want to Continue...?", vbYesNo, "SALES RETURN..") = vbYes Then
'                TXTQTY.Enabled = True
'                TXTQTY.SetFocus
'            'Else
'            '    TXTQTY.Enabled = False
'            '    CmbPack.Enabled = True
'            '    CmbPack.SetFocus
'            'End If
'        End If
        TXTQTY.Enabled = True
        TXTQTY.SetFocus
    End If

End Function

Private Sub GRDPRERATE_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            txtBatch.Text = IIf(IsNull(GRDPRERATE.Columns(11)), "", GRDPRERATE.Columns(11))
            LBLCOST.Caption = IIf(IsNull(GRDPRERATE.Columns(10)), "", GRDPRERATE.Columns(10))
            
            lblvchno.Caption = IIf(IsNull(GRDPRERATE.Columns(14)), "", GRDPRERATE.Columns(14))
            lbltrxtype.Caption = IIf(IsNull(GRDPRERATE.Columns(0)), "", GRDPRERATE.Columns(0))
            lbltrxyear.Caption = IIf(IsNull(GRDPRERATE.Columns(13)), "", GRDPRERATE.Columns(13))
            lbllineno.Caption = IIf(IsNull(GRDPRERATE.Columns(15)), "", GRDPRERATE.Columns(15))
            Los_Pack.Text = IIf(IsNull(GRDPRERATE.Columns(18)), 1, GRDPRERATE.Columns(18))
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = "1"
            txtPD.Text = IIf(IsNull(GRDPRERATE.Columns(16)), "", GRDPRERATE.Columns(16))
            TXTRATE.Text = IIf(IsNull(GRDPRERATE.Columns(5)), "", Format(Round(Val(GRDPRERATE.Columns(5)), 2), ".000"))
            txtmrpbt.Text = 100 * Val(TXTRATE.Text) / 105
            TXTPTR.Text = IIf(IsNull(GRDPRERATE.Columns(6)), "", Format(Round(Val(GRDPRERATE.Columns(6)), 3), ".000"))
            txtNetrate.Text = IIf(IsNull(GRDPRERATE.Columns(7)), "", Format(Round(Val(GRDPRERATE.Columns(7)), 3), ".000"))
            TxttaxMRP.Text = IIf(IsNull(GRDPRERATE.Columns(8)), "", Format(Val(GRDPRERATE.Columns(8)), ".00"))
            If GRDPRERATE.Columns(9) = "M" Then
                OPTTaxMRP.Value = True
            ElseIf GRDPRERATE.Columns(9) = "V" Then
                OPTVAT.Value = True
            Else
                OPTVAT.Value = True
            End If
            txtinvnodate.Text = Trim(IIf(IsNull(GRDPRERATE.Columns(14)), "", GRDPRERATE.Columns(14)))
            TxtInvoiceDate.Text = IIf(IsDate(GRDPRERATE.Columns(3)), Format(GRDPRERATE.Columns(3), "DD/MM/YYYY"), "  /  /    ")
            TXTEXPDATE.Text = IIf(IsDate(GRDPRERATE.Columns(11)), Format(GRDPRERATE.Columns(11), "DD/MM/YYYY"), "  /  /    ")
            TXTEXPIRY.Text = IIf(IsDate(GRDPRERATE.Columns(11)), Format(GRDPRERATE.Columns(11), "MM/YY"), "  /  ")
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            
            Dim BIL_PRE, BILL_SUF As String
            Dim RSTCOMPANY As ADODB.Recordset
            Set RSTCOMPANY = New ADODB.Recordset
            RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & lbltrxyear.Caption & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
                If lbltrxtype.Caption = "HI" Then
                    BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
                    BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
                ElseIf lbltrxtype.Caption = "GI" Then
                    BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
                    BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
                ElseIf lbltrxtype.Caption = "SV" Then
                    BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
                    BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
                End If
            End If
            RSTCOMPANY.Close
            Set RSTCOMPANY = Nothing
            txtinvnodate.Text = BIL_PRE & Format(Trim(txtinvnodate.Text), "0000") & BILL_SUF
                            
            Fram.Enabled = True
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Fram.Enabled = True
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
    End Select
    Exit Sub
ERRHAND:
End Sub

Private Sub Los_Pack_GotFocus()
    FRMEGRDTMP.Visible = False
    Los_Pack.SelStart = 0
    Los_Pack.SelLength = Len(Los_Pack.Text)
End Sub

Private Sub Los_Pack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Los_Pack.Text) = 0 Then Los_Pack.Text = 1
            'Los_Pack.Enabled = False
            CmbPack.Enabled = True
            TXTQTY.Enabled = True
            Call FILL_PREVIIOUSRATE
            'TXTQTY.SetFocus
         Case vbKeyEscape
             If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            CmbPack.Enabled = False
            TXTRATE.Enabled = False
            txtNetrate.Enabled = False
            TXTPTR.Enabled = False
            Los_Pack.Enabled = False
            TXTQTY.Enabled = False
            txtinvnodate.Enabled = False
            TxtInvoiceDate.Enabled = False
            txtBatch.Enabled = False
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
            TxttaxMRP.Enabled = False
            txtPD.Enabled = False
            Txtgrossamt.Enabled = False
    
            
            TXTPRODUCT.Enabled = True
            txtcategory.Enabled = True
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
    lblretail.Caption = 0
    lblwsale.Caption = 0
    lblvan.Caption = 0
    LBLMRP.Caption = 0
    LBLCOST.Caption = 0
    lblcase.Caption = 0
    lblLWPrice.Caption = 0
    lblcrtnpack.Caption = 1
    
    lblvchno.Caption = ""
    lbltrxtype.Caption = ""
    lbltrxyear.Caption = ""
    lbllineno.Caption = ""
            
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn

            If Trim(TXTITEMCODE.Text) = "" Then
                txtcategory.Enabled = True
                TXTPRODUCT.Enabled = True
                txtcategory.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False

            Set grdtmp.DataSource = Nothing
            If PHYCODE_FLAG = True Then
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            Else
                PHY_CODE.Close
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            End If

            Set grdtmp.DataSource = PHY_CODE

            If PHY_CODE.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "Sales Return"
                Exit Sub
            End If

            If PHY_CODE.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ORDER BY CREATE_DATE", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    RSTRXFILE.MoveLast
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.Text = 1
                Else
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = 1
                    Los_Pack.Text = 1
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND

                    TXTEXPDATE.Text = "  /  /    "
                    txtBatch.Text = ""
                    txtinvnodate.Text = ""
                    TxtInvoiceDate.Text = "  /  /    "
                    TXTEXPIRY.Text = "  /  "
                    TXTRATE.Text = ""
                    txtmrpbt.Text = ""
                    TXTPTR.Text = ""
                    txtNetrate.Text = ""
                    txtretail.Text = ""
                    txtWS.Text = ""
                    txtvanrate.Text = ""
                    txtcrtn.Text = ""
                    txtcrtnpack.Text = ""
                    txtprofit.Text = ""
                    TxttaxMRP.Text = ""
                    Los_Pack.Text = "1"
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND
                    OPTVAT.Value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing

                If PHY_CODE.RecordCount = 1 Then
                    TXTITEMCODE.Enabled = False
                    TxtBarcode.Enabled = False
                    TXTPRODUCT.Enabled = False
                    txtcategory.Enabled = False
                    
                    CmbPack.Enabled = True
                    TXTRATE.Enabled = True
                    txtNetrate.Enabled = True
                    TXTPTR.Enabled = True
                    Los_Pack.Enabled = True
                    txtinvnodate.Enabled = True
                    TxtInvoiceDate.Enabled = True
                    txtBatch.Enabled = True
                    TXTEXPIRY.Visible = True
                    TXTEXPDATE.Enabled = True
                    TxttaxMRP.Enabled = True
                    txtPD.Enabled = True
                    Txtgrossamt.Enabled = True
                    TXTQTY.Enabled = True
                    
                    Call FILL_PREVIIOUSRATE
                    'TXTQTY.SetFocus
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
            'TXTSLNO.Enabled = True
            TxtBarcode.Enabled = True
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            TxtBarcode.SetFocus
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
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            'If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub TxtCST_LostFocus()
    Dim DISC As Currency

    On Error GoTo ERRHAND
    If (TxtCST.Text = "") Then
        DISC = 0
    Else
        DISC = TxtCST.Text
    End If
    If grdsales.rows = 1 Then
        TxtCST.Text = "0"
        Exit Sub
    End If
    TxtCST.Text = Format(TxtCST.Text, ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtCST.Text)), 0), ".00")
    Exit Sub
ERRHAND:
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
            If txtcategory.Enabled = True Then txtcategory.SetFocus
            If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
            'If TXTITEMCODE.Enabled = True Then TXTITEMCODE.SetFocus
            If TXTQTY.Enabled = True Then TXTQTY.SetFocus
            If TXTRATE.Enabled = True Then TXTRATE.SetFocus
            If TXTEXPIRY.Visible = True Then TXTEXPIRY.SetFocus
            If TXTEXPDATE.Enabled = True Then TXTEXPDATE.SetFocus
            If txtBatch.Enabled = True Then txtBatch.SetFocus
            If cmdadd.Enabled = True Then cmdadd.SetFocus
    End Select
End Sub

Private Sub TxtInsurance_LostFocus()
    Dim DISC As Currency

    On Error GoTo ERRHAND
    If (TxtInsurance.Text = "") Then
        DISC = 0
    Else
        DISC = TxtInsurance.Text
    End If
    If grdsales.rows = 1 Then
        TxtInsurance.Text = "0"
        Exit Sub
    End If
    TxtInsurance.Text = Format(TxtInsurance.Text, ".00")
    'LBLTOTAL.Caption = Format(Round((Val(lbltotalwodiscount.Caption) + (Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(txtcst.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), ".00")
    LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + ((Val(lbltotalwodiscount.Caption) + Val(TxtInsurance.Text)) * Val(TxtCST.Text) / 100) + Val(TxtInsurance.Text) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(txtcramt.Text)), 0), "0.00")
    'LBLTOTAL.Caption = Format(Round(Val(lbltotalwodiscount.Caption) + Val(txtaddlamt.Text) - (Val(TXTDISCAMOUNT.Text) + Val(TxtInsurance.Text)), 0), ".00")
    Exit Sub
ERRHAND:
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
            txtWS.Enabled = False
            txtWsalePercent.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            txtWS.SetFocus
    End Select
End Sub

Private Sub txtWsalePercent_LostFocus()
    If optdiscper.Value = True Then
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - (Val(TXTPTR.Text) * Val(txtPD.Text) / 100)
    Else
        TXTPTR.Tag = Val(TXTPTR.Text) + (Val(TXTPTR.Text) * Val(TxttaxMRP.Text) / 100) - Val(txtPD.Text)
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
                TxtWarranty.Enabled = False
                CmbWrnty.Enabled = False
                cmdadd.Enabled = True
                cmdadd.SetFocus
            Else
                CmbWrnty.Enabled = True
                CmbWrnty.SetFocus
            End If
         Case vbKeyEscape
            TxtWarranty.Enabled = False
            CmbWrnty.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
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
                MsgBox "Please select the Warranty Period", , "Sales Return"
                CmbWrnty.SetFocus
                Exit Sub
            End If
            If Val(TxtWarranty.Text) = 0 Then CmbWrnty.ListIndex = -1
            TxtWarranty.Enabled = False
            CmbWrnty.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            TxtWarranty.SetFocus
    End Select
End Sub

Private Function checklastbill()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ERRHAND

    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'SR' ", db, adOpenStatic, adLockReadOnly
    'rstBILL.Open "Select MAX(VCH_NO) From RETURNMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= 'SR'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing

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
        
    Screen.MousePointer = vbHourglass
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
        Print #1, AlignLeft("RETURN", 25)
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
                AlignRight(Format(Round(Val(grdsales.TextMatrix(i, 18)), 2), "0.00"), 9) & _
                AlignRight(Format(Val(grdsales.TextMatrix(i, 13)), "0.00"), 13) '& _
                Chr(27) & Chr(72)  '//Bold Ends
        Next i
    
        Print #1, AlignRight("-------------", 80)
        'Print #1, Chr(27) & Chr(71) & Space(10) & AlignRight("Amount ", 57) & AlignRight(Format(LBLTOTAL.Caption, "####.00"), 10)
        Print #1, AlignRight("Round off ", 65) & AlignRight(Format(Round(LBLTOTAL.Caption, 0) - Val(LBLTOTAL.Caption), "0.00"), 12)
        Print #1, Chr(13)
        Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(18) & AlignRight("NET AMOUNT: ", 11) & AlignRight((Format(Val(LBLTOTAL.Caption), "####.00")), 9)
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
    Screen.MousePointer = vbNormal
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Function

Private Sub txtbarcode_GotFocus()
    lblretail.Caption = 0
    lblwsale.Caption = 0
    lblvan.Caption = 0
    LBLMRP.Caption = 0
    LBLCOST.Caption = 0
    lblcase.Caption = 0
    lblLWPrice.Caption = 0
    lblcrtnpack.Caption = 1
    
    lblvchno.Caption = ""
    lbltrxtype.Caption = ""
    lbltrxyear.Caption = ""
    lbllineno.Caption = ""
            
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
    FRMEGRDTMP.Visible = False
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn


            CmdDelete.Enabled = False
            
            If Trim(TxtBarcode.Text) <> "" Then
                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From TRXFILE  WHERE BARCODE = '" & Trim(TxtBarcode.Text) & "' ORDER BY TRX_YEAR DESC, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    TXTITEMCODE.Text = RSTRXFILE!ITEM_CODE
                    TXTPRODUCT.Text = IIf(IsNull(RSTRXFILE!ITEM_NAME), "", RSTRXFILE!ITEM_NAME)
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = 1 'IIf(IsNull(RSTRXFILE!LOOSE_PACK), "", RSTRXFILE!LOOSE_PACK) 'RSTRXFILE!LOOSE_PACK
                    Los_Pack.Text = IIf(IsNull(RSTRXFILE!LOOSE_PACK), "", RSTRXFILE!LOOSE_PACK) 'RSTRXFILE!LOOSE_PACK
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), "Nos", RSTRXFILE!PACK_TYPE)
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND
    
                    TXTEXPDATE.Text = "  /  /    "
                    txtBatch.Text = ""
                    txtinvnodate.Text = ""
                    TxtInvoiceDate.Text = "  /  /    "
                    TXTEXPIRY.Text = "  /  "
                    TXTRATE.Text = ""
                    txtmrpbt.Text = ""
                    TXTPTR.Text = ""
                    txtNetrate.Text = ""
                    txtretail.Text = ""
                    txtWS.Text = ""
                    txtvanrate.Text = ""
                    txtcrtn.Text = ""
                    txtcrtnpack.Text = ""
                    txtprofit.Text = ""
                    TxttaxMRP.Text = ""
                    Los_Pack.Text = "1"
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND
                    OPTVAT.Value = True
                    RSTRXFILE.Close
                    Set RSTRXFILE = Nothing
                    
                    TxtBarcode.Enabled = False
                    TXTITEMCODE.Enabled = False
                    TXTPRODUCT.Enabled = False
                    txtcategory.Enabled = False
                    
                    CmbPack.Enabled = True
                    TXTRATE.Enabled = True
                    txtNetrate.Enabled = True
                    TXTPTR.Enabled = True
                    Los_Pack.Enabled = True
                    txtinvnodate.Enabled = True
                    TxtInvoiceDate.Enabled = True
                    txtBatch.Enabled = True
                    TXTEXPIRY.Visible = True
                    TXTEXPDATE.Enabled = True
                    TxttaxMRP.Enabled = True
                    txtPD.Enabled = True
                    Txtgrossamt.Enabled = True
                    TXTQTY.Enabled = True
                    FRMEGRDTMP.Visible = False
                    Call FILL_PREVIIOUSRATE
                    'TXTQTY.SetFocus
                    
                    Exit Sub
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
            End If
            
            If Trim(TXTITEMCODE.Text) = "" Then
                txtcategory.Enabled = True
                TXTPRODUCT.Enabled = True
                txtcategory.SetFocus
                Exit Sub
            End If
            Set grdtmp.DataSource = Nothing
            If PHYCODE_FLAG = True Then
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            Else
                PHY_CODE.Close
                PHY_CODE.Open "Select * From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockReadOnly
                PHYCODE_FLAG = False
            End If

            Set grdtmp.DataSource = PHY_CODE

            If PHY_CODE.RecordCount = 0 Then
                MsgBox "Item not found!!!!", , "Sales Return"
                Exit Sub
            End If

            If PHY_CODE.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "' ORDER BY CREATE_DATE", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    RSTRXFILE.MoveLast
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = IIf(IsNull(RSTRXFILE!LINE_DISC), "", RSTRXFILE!LINE_DISC)
                    Txtpack.Text = 1
                Else
                    TXTUNIT.Text = 1 'IIf(IsNull(RSTRXFILE!UNIT), "", RSTRXFILE!UNIT)
                    Txtpack.Text = 1
                    Los_Pack.Text = 1
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND

                    TXTEXPDATE.Text = "  /  /    "
                    txtBatch.Text = ""
                    txtinvnodate.Text = ""
                    TxtInvoiceDate.Text = "  /  /    "
                    TXTEXPIRY.Text = "  /  "
                    TXTRATE.Text = ""
                    txtmrpbt.Text = ""
                    TXTPTR.Text = ""
                    txtNetrate.Text = ""
                    txtretail.Text = ""
                    txtWS.Text = ""
                    txtvanrate.Text = ""
                    txtcrtn.Text = ""
                    txtcrtnpack.Text = ""
                    txtprofit.Text = ""
                    TxttaxMRP.Text = ""
                    Los_Pack.Text = "1"
                    TxtWarranty.Text = ""
                    On Error Resume Next
                    CmbPack.Text = "Nos"
                    CmbWrnty.ListIndex = -1
                    On Error GoTo ERRHAND
                    OPTVAT.Value = True
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing

                If PHY_CODE.RecordCount = 1 Then
                    TxtBarcode.Enabled = False
                    TXTITEMCODE.Enabled = False
                    TXTPRODUCT.Enabled = False
                    txtcategory.Enabled = False
                    
                    CmbPack.Enabled = True
                    TXTRATE.Enabled = True
                    txtNetrate.Enabled = True
                    TXTPTR.Enabled = True
                    Los_Pack.Enabled = True
                    txtinvnodate.Enabled = True
                    TxtInvoiceDate.Enabled = True
                    txtBatch.Enabled = True
                    TXTEXPIRY.Visible = True
                    TXTEXPDATE.Enabled = True
                    TxttaxMRP.Enabled = True
                    txtPD.Enabled = True
                    Txtgrossamt.Enabled = True
                    TXTQTY.Enabled = True
                    
                    Call FILL_PREVIIOUSRATE
                    'TXTQTY.SetFocus
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
            TXTITEMCODE.Enabled = False
            TxtBarcode.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Function JSON_REPORT()

    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open Rptpath & "JSON\e_invoice.json" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "JSON\e_invoice.json" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
    
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompTin, CompPin, Auth_ky As String
    Dim EINV_DATE As String
    Dim Eway_Falg As Boolean
    Dim gross_amt As Double
    Dim DISC_AMT As Double
    Dim TAX_AMT As Double
    Dim tot_val As Double
    Dim cess_amt As Double
    Dim ONLINEBILL As Boolean
    
    EINV_DATE = ""
    Screen.MousePointer = vbHourglass
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        CompName = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", Left(RSTCOMPANY!COMP_NAME, 100))
        CompAddress1 = IIf(IsNull(RSTCOMPANY!Address), "", Left(RSTCOMPANY!Address, 100))
        CompAddress2 = IIf(IsNull(RSTCOMPANY!HO_NAME), "", Left(RSTCOMPANY!HO_NAME, 100))
        CompAddress3 = IIf(IsNull(RSTCOMPANY!TEL_NO) Or RSTCOMPANY!TEL_NO = "", "", Left(RSTCOMPANY!TEL_NO, 10))
        CompTin = IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", "", RSTCOMPANY!CST)
        CompPin = IIf(IsNull(RSTCOMPANY!PINCODE) Or RSTCOMPANY!CST = "", "", RSTCOMPANY!PINCODE)
        Auth_ky = IIf(IsNull(RSTCOMPANY!auth_key), "", RSTCOMPANY!auth_key)
        If Auth_ky <> "" And IsDate(EncryptString(RSTCOMPANY!auth_date, "_einv*")) Then
            EINV_DATE = EncryptString(RSTCOMPANY!auth_date, "_einv*")
        Else
            EINV_DATE = ""
        End If
        If Not IsNull(RSTCOMPANY!HSN_SUM) And Val(LBLTOTAL.Caption) > RSTCOMPANY!HSN_SUM Then Eway_Falg = True
        If IsNull(RSTCOMPANY!ONLINE_BILL) Or RSTCOMPANY!ONLINE_BILL = "N" Then
            ONLINEBILL = False
        Else
            ONLINEBILL = True
        End If
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    If EINV_DATE = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Activation expired. Please renew the key", , "EzBiz"
        Exit Function
    End If
    If DateDiff("d", TXTINVDATE.Text, EINV_DATE) <= 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "Activation expired. Please renew the key", , "EzBiz"
        Exit Function
    End If
    
    If DateDiff("d", TXTINVDATE.Text, EINV_DATE) <= 30 Then
        Screen.MousePointer = vbNormal
        MsgBox DateDiff("d", TXTINVDATE.Text, EINV_DATE) & " days remaining. Please renew the key before it expires", , "EzBiz"
    End If
    
    If Eway_Falg = True Then
        If Val(TxtKMS.Text) = 0 Then
            Screen.MousePointer = vbNormal
            MsgBox "Please enter the approx. distance in KMs", , "EzBiz"
            Exit Function
        End If
        
        If Trim(TxtVehicle.Text) = "" Then
            Screen.MousePointer = vbNormal
            MsgBox "Please enter Vehicle details", , "EzBiz"
            Exit Function
        End If
    End If
        
    If Trim(CompTin) = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Seller GST No. not available", , "EzBiz"
        Exit Function
    End If
    If Len(Trim(CompPin)) <> 6 Then
        Screen.MousePointer = vbNormal
        MsgBox "Seller Pincode not entered", , "EzBiz"
        Exit Function
    End If
    If Len(Trim(CompTin)) <> 15 Then
        Screen.MousePointer = vbNormal
        MsgBox "Not a valid GST No.", , "EzBiz"
        Exit Function
    End If
    If Len(Trim(txtPin.Text)) <> 6 Then
        Screen.MousePointer = vbNormal
        MsgBox "Buyer Pincode not entered", , "EzBiz"
        Exit Function
    End If
    If Len(Trim(TXTTIN.Text)) <> 15 Then
        Screen.MousePointer = vbNormal
        MsgBox "Please check the buyer GST No.", , "EzBiz"
        Exit Function
    End If
    If Trim(TXTAREA.Text) = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "Please enter the area", , "EzBiz"
        Exit Function
    End If
    
    'Print #1, "["
    Print #1, "  {"
    Print #1, "    " & Chr(34) & "Version" & Chr(34) & ": " & Chr(34) & "1.1" & Chr(34) & ","
    Print #1, "    " & Chr(34) & "TranDtls" & Chr(34) & ": " & "{"
    Print #1, "      " & Chr(34) & "TaxSch" & Chr(34) & ": " & Chr(34) & "GST" & Chr(34) & ","
    Print #1, "      " & Chr(34) & "SupTyp" & Chr(34) & ": " & Chr(34) & "B2B" & Chr(34) & ","
'    If lblIGST.Caption = "Y" Then
'        Print #1, "      " & Chr(34) & "IgstOnIntra" & Chr(34) & ": " & Chr(34) & "Y" & Chr(34) & ","
'        Print #1, "      " & Chr(34) & "RegRev" & Chr(34) & ": " & Chr(34) & "Y" & Chr(34) & ","
'    Else
'        Print #1, "      " & Chr(34) & "IgstOnIntra" & Chr(34) & ": " & Chr(34) & "N" & Chr(34) & ","
'        Print #1, "      " & Chr(34) & "RegRev" & Chr(34) & ": " & Chr(34) & "N" & Chr(34) & ","
'    End If
    Print #1, "      " & Chr(34) & "IgstOnIntra" & Chr(34) & ": " & Chr(34) & "N" & Chr(34) & ","
    Print #1, "      " & Chr(34) & "RegRev" & Chr(34) & ": " & Chr(34) & "N" & Chr(34) & ","
    Print #1, "      " & Chr(34) & "EcmGstin" & Chr(34) & ": " & "null"
    Print #1, "    },"
    
    Print #1, "    " & Chr(34) & "DocDtls" & Chr(34) & ": " & "{"
    Print #1, "      " & Chr(34) & "Typ" & Chr(34) & ": " & Chr(34) & "CRN" & Chr(34) & ","
    Print #1, "      " & Chr(34) & "No" & Chr(34) & ": " & Chr(34) & "SR-" & txtBillNo.Text & Chr(34) & ","
    Print #1, "      " & Chr(34) & "Dt" & Chr(34) & ": " & Chr(34) & Format(TXTINVDATE.Text, "DD/MM/YYYY") & Chr(34)
    Print #1, "    },"
    
    'seller
    Print #1, "    " & Chr(34) & "SellerDtls" & Chr(34) & ": " & "{"
    If Trim(CompTin) = "" Then
        Print #1, "      " & Chr(34) & "Gstin" & Chr(34) & ": " & "null" & ","
    Else
        Print #1, "      " & Chr(34) & "Gstin" & Chr(34) & ": " & Chr(34) & Trim(CompTin) & Chr(34) & ","
    End If
    Print #1, "      " & Chr(34) & "LglNm" & Chr(34) & ": " & Chr(34) & Trim(CompName) & Chr(34) & ","
    If Trim(CompAddress1) = "" Then
        Print #1, "      " & Chr(34) & "Addr1" & Chr(34) & ": " & "null" & ","
    Else
        Print #1, "      " & Chr(34) & "Addr1" & Chr(34) & ": " & Chr(34) & Trim(CompAddress1) & Chr(34) & ","
    End If
    If Trim(CompAddress2) = "" Then
        Print #1, "      " & Chr(34) & "Addr2" & Chr(34) & ": " & "null" & ","
    Else
        Print #1, "      " & Chr(34) & "Addr2" & Chr(34) & ": " & Chr(34) & Trim(CompAddress2) & Chr(34) & ","
    End If
    If Trim(CompAddress2) = "" Then
        Print #1, "      " & Chr(34) & "Loc" & Chr(34) & ": " & "null" & ","
    Else
        Print #1, "      " & Chr(34) & "Loc" & Chr(34) & ": " & Chr(34) & Left(Trim(CompAddress2), 50) & Chr(34) & ","
    End If
    Print #1, "      " & Chr(34) & "Pin" & Chr(34) & ": " & Trim(CompPin) & ","
    If Trim(MDIMAIN.LBLSTATE.Caption) = "" Then
        Print #1, "      " & Chr(34) & "Stcd" & Chr(34) & ": " & Chr(34) & "32" & Chr(34) & ","
    Else
        Print #1, "      " & Chr(34) & "Stcd" & Chr(34) & ": " & Chr(34) & Trim(MDIMAIN.LBLSTATE.Caption) & Chr(34) & ","
    End If
    If Trim(CompAddress3) = "" Then
        Print #1, "      " & Chr(34) & "Ph" & Chr(34) & ": " & "null" & ","
    Else
        Print #1, "      " & Chr(34) & "Ph" & Chr(34) & ": " & Chr(34) & Trim(CompAddress3) & Chr(34) & ","
    End If
    Print #1, "      " & Chr(34) & "Em" & Chr(34) & ": " & "null"
    Print #1, "    },"
    
    'Buyer
    Print #1, "    " & Chr(34) & "BuyerDtls" & Chr(34) & ": " & "{"
    If Trim(TXTTIN.Text) = "" Then
        Print #1, "      " & Chr(34) & "Gstin" & Chr(34) & ": " & Chr(34) & "URP" & Chr(34) & ","
    Else
        Print #1, "      " & Chr(34) & "Gstin" & Chr(34) & ": " & Chr(34) & Trim(TXTTIN.Text) & Chr(34) & ","
    End If
    Print #1, "      " & Chr(34) & "LglNm" & Chr(34) & ": " & Chr(34) & Left(Trim(TXTDEALER.Text), 100) & Chr(34) & ","
    If Trim(TxtBillAddress.Text) = "" Then
        Print #1, "      " & Chr(34) & "Addr1" & Chr(34) & ": " & "null" & ","
    Else
        Print #1, "      " & Chr(34) & "Addr1" & Chr(34) & ": " & Chr(34) & Left(Trim(TxtBillAddress.Text), 100) & Chr(34) & ","
    End If
    Print #1, "      " & Chr(34) & "Addr2" & Chr(34) & ": " & "null" & ","
    If Trim(TXTAREA.Text) = "" Then
        Print #1, "      " & Chr(34) & "Loc" & Chr(34) & ": " & "null" & ","
    Else
        Print #1, "      " & Chr(34) & "Loc" & Chr(34) & ": " & Chr(34) & Left(Trim(TXTAREA.Text), 50) & Chr(34) & ","
    End If
    Print #1, "      " & Chr(34) & "Pin" & Chr(34) & ": " & Trim(txtPin.Text) & ","
    If Left(Trim(TXTTIN.Text), 2) = "" Then
        Print #1, "      " & Chr(34) & "Pos" & Chr(34) & ": " & Chr(34) & Trim(MDIMAIN.LBLSTATE.Caption) & Chr(34) & ","
        Print #1, "      " & Chr(34) & "Stcd" & Chr(34) & ": " & Chr(34) & Trim(MDIMAIN.LBLSTATE.Caption) & Chr(34) & ","
    Else
        Print #1, "      " & Chr(34) & "Pos" & Chr(34) & ": " & Chr(34) & Left(Trim(TXTTIN.Text), 2) & Chr(34) & ","
        Print #1, "      " & Chr(34) & "Stcd" & Chr(34) & ": " & Chr(34) & Left(Trim(TXTTIN.Text), 2) & Chr(34) & ","
    End If
    If Trim(TxtPhone.Text) = "" Then
        Print #1, "      " & Chr(34) & "Ph" & Chr(34) & ": " & "null" & ","
    Else
        Print #1, "      " & Chr(34) & "Ph" & Chr(34) & ": " & Chr(34) & Left(Trim(TxtPhone.Text), 10) & Chr(34) & ","
    End If
    Print #1, "      " & Chr(34) & "Em" & Chr(34) & ": " & "null"
    Print #1, "    },"
    
    
    gross_amt = 0
    tot_val = 0
    DISC_AMT = 0
    cess_amt = 0
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 27) <> "A" Then
            gross_amt = Round(gross_amt + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 3))) - (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 17)) / 100, 2)
        Else
            gross_amt = Round(gross_amt + (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 3))) - Val(grdsales.TextMatrix(i, 17)), 2)
        End If
        tot_val = Round(tot_val + Val(grdsales.TextMatrix(i, 13)), 2)
        'cess_amt = cess_amt + Val(grdsales.TextMatrix(i, 48))
    Next i
    
    Print #1, "    " & Chr(34) & "ValDtls" & Chr(34) & ": " & "{"
    Print #1, "    " & Chr(34) & "AssVal" & Chr(34) & ": " & gross_amt & ","
    If lblIGST.Caption = "Y" Then
        Print #1, "    " & Chr(34) & "IgstVal" & Chr(34) & ": " & Round((tot_val - gross_amt) - cess_amt, 2) & ","
        Print #1, "    " & Chr(34) & "CgstVal" & Chr(34) & ": " & 0 & ","
        Print #1, "    " & Chr(34) & "SgstVal" & Chr(34) & ": " & 0 & ","
    Else
        Print #1, "    " & Chr(34) & "IgstVal" & Chr(34) & ": " & 0 & ","
        Print #1, "    " & Chr(34) & "CgstVal" & Chr(34) & ": " & Round(((tot_val - gross_amt) - cess_amt) / 2, 2) & ","
        Print #1, "    " & Chr(34) & "SgstVal" & Chr(34) & ": " & Round(((tot_val - gross_amt) - cess_amt) / 2, 2) & ","
    End If
    Print #1, "    " & Chr(34) & "CesVal" & Chr(34) & ": " & Round(cess_amt, 2) & ","
    Print #1, "    " & Chr(34) & "StCesVal" & Chr(34) & ": " & 0 & ","
    Print #1, "    " & Chr(34) & "Discount" & Chr(34) & ": " & DISC_AMT & ","
    Print #1, "    " & Chr(34) & "OthChrg" & Chr(34) & ": " & 0 & ","
    Print #1, "    " & Chr(34) & "RndOffAmt" & Chr(34) & ": " & Round(Round(tot_val, 0) - Round(tot_val, 2), 2) & ","
    Print #1, "    " & Chr(34) & "TotInvVal" & Chr(34) & ": " & Round(tot_val, 0)
    Print #1, "    },"
    
    If Eway_Falg = False Then
        Print #1, "    " & Chr(34) & "EwbDtls" & Chr(34) & ": " & "null" & ","
    Else
        Print #1, "    " & Chr(34) & "EwbDtls" & Chr(34) & ": " & "{"
        Print #1, "    " & Chr(34) & "TransId" & Chr(34) & ": " & "null" & ","
        Print #1, "    " & Chr(34) & "TransName" & Chr(34) & ": " & "null" & ","
        Print #1, "    " & Chr(34) & "TransMode" & Chr(34) & ": " & Chr(34) & "1" & Chr(34) & ","
        Print #1, "    " & Chr(34) & "Distance" & Chr(34) & ": " & Val(TxtKMS.Text) & ","
        Print #1, "    " & Chr(34) & "TransDocNo" & Chr(34) & ": " & "null" & ","
        Print #1, "    " & Chr(34) & "TransDocDt" & Chr(34) & ": " & Chr(34) & Format(TXTINVDATE.Text, "DD/MM/YYYY") & Chr(34) & ","
        Print #1, "    " & Chr(34) & "VehNo" & Chr(34) & ": " & Chr(34) & Trim(TxtVehicle.Text) & Chr(34) & ","
        Print #1, "    " & Chr(34) & "VehType" & Chr(34) & ": " & Chr(34) & "R" & Chr(34)
        Print #1, "    },"
    End If
    
    Print #1, "    " & Chr(34) & "RefDtls" & Chr(34) & ": " & "{"
    Print #1, "    " & Chr(34) & "InvRm" & Chr(34) & ": " & Chr(34) & "NICGEPP2.0" & Chr(34)
    Print #1, "    },"
    
    Dim RSTITEMMAST As ADODB.Recordset
    Dim HSNCODE As String
    Dim uqccode As String
    Dim prod_name As String
    Print #1, "    " & Chr(34) & "ItemList" & Chr(34) & ": " & "["
    For i = 1 To grdsales.rows - 1
        gross_amt = 0
        tot_val = 0
        DISC_AMT = 0
        TAX_AMT = 0
        HSNCODE = ""
        uqccode = ""
        prod_name = Left(grdsales.TextMatrix(i, 2), 300)
        prod_name = Replace(prod_name, "'", "")
        prod_name = Replace(prod_name, Chr(34), "")
        prod_name = Replace(prod_name, "{", "")
        prod_name = Replace(prod_name, "}", "")
        prod_name = Replace(prod_name, "[", "")
        prod_name = Replace(prod_name, "]", "")
        prod_name = Replace(prod_name, "\", "")
        
        Print #1, "      {"
        Print #1, "        " & Chr(34) & "SlNo" & Chr(34) & ": " & Chr(34) & i & Chr(34) & ","
        Print #1, "        " & Chr(34) & "PrdDesc" & Chr(34) & ": " & Chr(34) & prod_name & Chr(34) & ","
        Print #1, "        " & Chr(34) & "IsServc" & Chr(34) & ": " & Chr(34) & "N" & Chr(34) & ","
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            HSNCODE = IIf(IsNull(RSTITEMMAST!REMARKS), "", RSTITEMMAST!REMARKS)
            uqccode = IIf(IsNull(RSTITEMMAST!UQC) Or RSTITEMMAST!UQC = "", "OTH", RSTITEMMAST!UQC)
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        If HSNCODE = "" Then
            Screen.MousePointer = vbNormal
            MsgBox "HSN Code not found in item", , "EzBiz"
            Exit Function
        End If
        If uqccode = "" Then uqccode = "OTH"
        
        Print #1, "        " & Chr(34) & "HsnCd" & Chr(34) & ": " & Chr(34) & HSNCODE & Chr(34) & ","
        Print #1, "        " & Chr(34) & "Qty" & Chr(34) & ": " & grdsales.TextMatrix(i, 3) & ","
        Print #1, "        " & Chr(34) & "FreeQty" & Chr(34) & ": " & grdsales.TextMatrix(i, 14) & ","
        Print #1, "        " & Chr(34) & "Unit" & Chr(34) & ": " & Chr(34) & uqccode & Chr(34) & ","
        Print #1, "        " & Chr(34) & "UnitPrice" & Chr(34) & ": " & Round(Val(grdsales.TextMatrix(i, 9)), 3) & ","
        Print #1, "        " & Chr(34) & "TotAmt" & Chr(34) & ": " & Round(Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 3)), 2) & ","
        If grdsales.TextMatrix(i, 27) <> "A" Then
            Print #1, "        " & Chr(34) & "Discount" & Chr(34) & ": " & Round((Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 17)) / 100, 2) & ","
            DISC_AMT = (Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 17)) / 100
        Else
            Print #1, "        " & Chr(34) & "Discount" & Chr(34) & ": " & Round(Val(grdsales.TextMatrix(i, 17)), 2) & ","
            DISC_AMT = Val(grdsales.TextMatrix(i, 17))
        End If
        gross_amt = Round((Val(grdsales.TextMatrix(i, 9)) * Val(grdsales.TextMatrix(i, 3))) - DISC_AMT, 2)
        DISC_AMT = Round(DISC_AMT, 2)
        TAX_AMT = Round(gross_amt * Val(grdsales.TextMatrix(i, 10)) / 100, 3)
        tot_val = Round(Val(grdsales.TextMatrix(i, 13)), 2)
        
        Print #1, "        " & Chr(34) & "PreTaxVal" & Chr(34) & ": " & 0 & ","
        Print #1, "        " & Chr(34) & "AssAmt" & Chr(34) & ": " & gross_amt & ","
        Print #1, "        " & Chr(34) & "GstRt" & Chr(34) & ": " & Val(grdsales.TextMatrix(i, 10)) & ","
        If lblIGST.Caption = "Y" Then
            Print #1, "        " & Chr(34) & "IgstAmt" & Chr(34) & ": " & TAX_AMT & ","
            Print #1, "        " & Chr(34) & "CgstAmt" & Chr(34) & ": " & 0 & ","
            Print #1, "        " & Chr(34) & "SgstAmt" & Chr(34) & ": " & 0 & ","
        Else
            Print #1, "        " & Chr(34) & "IgstAmt" & Chr(34) & ": " & 0 & ","
            Print #1, "        " & Chr(34) & "CgstAmt" & Chr(34) & ": " & Round(TAX_AMT / 2, 2) & ","
            Print #1, "        " & Chr(34) & "SgstAmt" & Chr(34) & ": " & Round(TAX_AMT / 2, 2) & ","
        End If
        Print #1, "        " & Chr(34) & "CesRt" & Chr(34) & ": " & 0 & ","
        Print #1, "        " & Chr(34) & "CesAmt" & Chr(34) & ": " & 0 & ","
        Print #1, "        " & Chr(34) & "CesNonAdvlAmt" & Chr(34) & ": " & 0 & ","
        Print #1, "        " & Chr(34) & "StateCesRt" & Chr(34) & ": " & 0 & ","
        Print #1, "        " & Chr(34) & "StateCesAmt" & Chr(34) & ": " & 0 & ","
        Print #1, "        " & Chr(34) & "StateCesNonAdvlAmt" & Chr(34) & ": " & 0 & ","
        Print #1, "        " & Chr(34) & "OthChrg" & Chr(34) & ": " & 0 & ","
        Print #1, "        " & Chr(34) & "TotItemVal" & Chr(34) & ": " & tot_val
        If i = grdsales.rows - 1 Then
            Print #1, "        }"
        Else
            Print #1, "        },"
        End If
    Next i
    Print #1, "      ]"
    Print #1, "    }"
    'Print #1, "  ]"
    Close #1 '//Closing the file
    
    If Auth_ky = "" Then
        Screen.MousePointer = vbNormal
        MsgBox "JSON Generated and saved successfuly." & Chr(13) & "Cannot upload automatically since not registered.", , "EzBiz"
        Exit Function
    End If

    '
    
    ' Replace this with your E-Invoice JSON content
  ' For samples, refer to https://my.gstzen.in/p/e-invoice-samples-list/
  

    Dim fso As New FileSystemObject
    Dim requestJSON As TextStream
    Dim ReadFileIntoString As String
    Set requestJSON = fso.OpenTextFile(Rptpath & "JSON\e_invoice.json")
    ReadFileIntoString = requestJSON.ReadAll
    
'    Screen.MousePointer = vbNormal
'    MsgBox ReadFileIntoString, , "EzBiz"
'    Screen.MousePointer = vbHourglass
'  Dim requestJSON As String
'  requestJSON = "{ \"SellerDtls\": {\"Gstin\": \"27AADCG4992P1ZT\"} }"

  ' The Authentication Token that you have received from GSTZen
  ' Replace with your token
'  Dim token As String
'  token = "de3a3a01-273a-4a81-8b75-13fe37f14dc6"
    'token="5d34161d-f610-4a73-a414-0b65f53b8396"
  'Debug.Print ReadFileIntoString

    Dim responseJSON As String
  'responseJSON = GetDataFromURL("https://my.gstzen.in/~gstzen/a/post-einvoice-data/einvoice-json/", "POST", ReadFileIntoString, Auth_ky)
  responseJSON = GetDataFromURL("http://your.gstzen.in/~gstzen/a/post-einvoice-data/einvoice-json/", "POST", ReadFileIntoString, Auth_ky)
'  Screen.MousePointer = vbNormal
'  MsgBox responseJSON, , "EzBiz"
'  Screen.MousePointer = vbHourglass
  
  'https://my.gstzen.in/~gstzen/a/post-einvoice-data/einvoice-json/genewb/
  
  'Debug.Print responseJSON
'https://my.gstzen.in/~gstzen/a/post-einvoice-data/einvoice-json/a/invoices/b5e6615a-4348-4860-acd1-edf967b529ce/einvoice/.pdf2/
'https://my.gstzen.in/~gstzen/a/post-einvoice-data/einvoice-json/

    If Len(responseJSON) < 25 Then
        MsgBox responseJSON, , "EzBiz"
        Exit Function
    End If
    
    Dim p
    Set p = JSON.parse(responseJSON)
    Screen.MousePointer = vbNormal
    MsgBox p.Item("message"), , "EzBiz"
    Screen.MousePointer = vbHourglass
    
    Dim r As Long
    If p.Item("status") = 1 Then
        r = ShellExecute(0, "open", "https://my.gstzen.in/~gstzen" & Mid(p.Item("InvoicePdfUrl"), 13), 0, 0, 1)
    Else
        Exit Function
    End If
    Screen.MousePointer = vbNormal
    '    Dim qrcode As String
    '    qrcode = "https://my.gstzen.in/~gstzen" & Mid(p.Item("SignedQrCodeImgUrl"), 13)
       
        'If MsgBox("Do you want to print QR Code generated Bill", vbYesNo, "E-Invoice") = vbNo Then Exit Function
    
    
    
'    If Eway_Falg = True Then
'        responseJSON = GetDataFromURL("https://my.gstzen.in/~gstzen/a/post-einvoice-data/einvoice-json/genewb/", "POST", ReadFileIntoString, Auth_ky)
'    End If
    Exit Function
''    'Dim I As Long
''
'''    ITM = Split(URL, ":")
'''    If ITM(1) = 1 Then
'''        MsgBox "Message: " & ITM(3)
''
''        If UCase(Trim(ITM(34))) = Chr(34) & "INVOICEPDFURL" & Chr(34) Then
''            r = ShellExecute(0, "open", "https://my.gstzen.in/~gstzen" & Mid(ITM(41), 15, Len(ITM(41)) - 16), 0, 0, 1)
''        End If
''        If UCase(Trim(ITM(40))) = Chr(34) & "INVOICEPDFURL" & Chr(34) Then
''            r = ShellExecute(0, "open", "https://my.gstzen.in/~gstzen" & Mid(ITM(41), 15, Len(ITM(41)) - 16), 0, 0, 1)
''        Else
''            r = ShellExecute(0, "open", "https://my.gstzen.in/~gstzen" & Mid(ITM(35), 15, Len(ITM(35)) - 16), 0, 0, 1)
''        End If
'''    Else
'''        MsgBox "Error" & ITM(3)
'''
'''    End If
    

ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
    
End Function

Function GetDataFromURL(strURL, strMethod, strPostData, strToken)
  Dim lngTimeout
  Dim strUserAgentString
  Dim intSslErrorIgnoreFlags
  Dim blnEnableRedirects
  Dim blnEnableHttpsToHttpRedirects
  Dim strHostOverride
  Dim strLogin
  Dim strPassword
  Dim strResponseText
  Dim objWinHttp
  
  'On Error GoTo ErrHand
  lngTimeout = 59000
  strUserAgentString = "zen_request/0.1"
  intSslErrorIgnoreFlags = 13056 ' 13056: ignore all err, 0: accept no err
  blnEnableRedirects = True
  blnEnableHttpsToHttpRedirects = True
  strHostOverride = ""
  strLogin = ""
  strPassword = ""
  Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
  objWinHttp.setTimeouts lngTimeout, lngTimeout, lngTimeout, lngTimeout
  objWinHttp.Open strMethod, strURL
  If strMethod = "POST" Then
    objWinHttp.setRequestHeader "Content-type", _
      "application/json"

    objWinHttp.setRequestHeader "Token", _
      strToken

  End If
  If strHostOverride <> "" Then
    objWinHttp.setRequestHeader "Host", strHostOverride
  End If
  objWinHttp.Option(0) = strUserAgentString
  objWinHttp.Option(4) = intSslErrorIgnoreFlags
  objWinHttp.Option(6) = blnEnableRedirects
  objWinHttp.Option(12) = blnEnableHttpsToHttpRedirects
  If (strLogin <> "") And (strPassword <> "") Then
    objWinHttp.SetCredentials strLogin, strPassword, 0
  End If
  On Error Resume Next
  objWinHttp.send (strPostData)
  If err.Number = 0 Then
    If objWinHttp.Status = "200" Then
      GetDataFromURL = objWinHttp.responseText
    Else
      GetDataFromURL = "HTTP " & objWinHttp.Status & " " & _
        objWinHttp.statusText
    End If
  Else
    GetDataFromURL = "Error " & err.Number & " " & err.source & " " & _
      err.Description
  End If
  On Error GoTo 0
  Set objWinHttp = Nothing
  Exit Function
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Function


