VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmitemmaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Creation"
   ClientHeight    =   9165
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   14535
   Icon            =   "frmitemmastermini.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   14535
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   9285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtProduct 
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
      Left            =   1320
      TabIndex        =   0
      Top             =   285
      Visible         =   0   'False
      Width           =   4965
   End
   Begin VB.TextBox TxtItemcode 
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
      Left            =   1320
      MaxLength       =   21
      TabIndex        =   34
      Top             =   285
      Width           =   4950
   End
   Begin VB.Frame FRAME 
      Height          =   8490
      Left            =   15
      TabIndex        =   1
      Top             =   645
      Width           =   8025
      Begin VB.TextBox txtExpense 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   4935
         MaxLength       =   5
         TabIndex        =   90
         Top             =   4755
         Width           =   990
      End
      Begin VB.ComboBox CmbQAC 
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
         ItemData        =   "frmitemmastermini.frx":0442
         Left            =   6570
         List            =   "frmitemmastermini.frx":04CA
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   4785
         Width           =   1365
      End
      Begin VB.CheckBox ChkFreeWarn 
         Appearance      =   0  'Flat
         Caption         =   "Free Qty Warning"
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
         Height          =   195
         Left            =   75
         TabIndex        =   86
         Top             =   7530
         Width           =   1995
      End
      Begin VB.CheckBox chkfocusQty 
         Appearance      =   0  'Flat
         Caption         =   "Focus on Qty"
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
         Height          =   195
         Left            =   75
         TabIndex        =   78
         Top             =   7305
         Width           =   1575
      End
      Begin VB.TextBox txtpackdet 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   7410
         MaxLength       =   5
         TabIndex        =   15
         Top             =   3165
         Width           =   525
      End
      Begin VB.TextBox Txtpackdes 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   6225
         MaxLength       =   5
         TabIndex        =   14
         Top             =   3165
         Width           =   570
      End
      Begin VB.TextBox Txtbarcode 
         Appearance      =   0  'Flat
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
         Height          =   435
         Left            =   4935
         MaxLength       =   30
         TabIndex        =   25
         Top             =   5205
         Width           =   2985
      End
      Begin VB.TextBox TxtSpec 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2040
         TabIndex        =   63
         Top             =   660
         Width           =   5910
      End
      Begin VB.TextBox TxtMRP 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   3525
         TabIndex        =   18
         Top             =   3630
         Width           =   900
      End
      Begin VB.TextBox TxtLocation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Kerala"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   6705
         MaxLength       =   10
         TabIndex        =   24
         Top             =   4095
         Width           =   1215
      End
      Begin VB.TextBox TxtHSN 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   3630
         MaxLength       =   10
         TabIndex        =   23
         Top             =   4095
         Width           =   1620
      End
      Begin VB.CheckBox cHKcHANGE 
         Appearance      =   0  'Flat
         Caption         =   "Price Changing Items"
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
         Height          =   195
         Left            =   75
         TabIndex        =   59
         Top             =   6870
         Width           =   2160
      End
      Begin VB.CheckBox chkunbill 
         Appearance      =   0  'Flat
         Caption         =   "Un Bill Items"
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
         Height          =   195
         Left            =   75
         TabIndex        =   58
         Top             =   6645
         Width           =   1575
      End
      Begin VB.CommandButton cmddelphoto 
         Caption         =   "Remove Photo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3555
         TabIndex        =   55
         Top             =   6630
         Width           =   1350
      End
      Begin VB.CommandButton CMDBROWSE 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3555
         TabIndex        =   54
         Top             =   6240
         Width           =   1350
      End
      Begin VB.CheckBox ChkDead 
         Appearance      =   0  'Flat
         Caption         =   "Dead Stock"
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
         Height          =   195
         Left            =   75
         TabIndex        =   53
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox txtPack 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   2010
         TabIndex        =   11
         Top             =   3165
         Width           =   480
      End
      Begin VB.TextBox txtLPrice 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   2280
         TabIndex        =   22
         Top             =   4095
         Width           =   900
      End
      Begin VB.TextBox TxtCost 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   2010
         TabIndex        =   17
         Top             =   3630
         Width           =   1035
      End
      Begin VB.TextBox TxtLPack 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   975
         TabIndex        =   21
         Top             =   4095
         Width           =   555
      End
      Begin VB.TextBox TxtWS 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   6960
         TabIndex        =   20
         Top             =   3630
         Width           =   975
      End
      Begin VB.TextBox txtRT 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   5190
         TabIndex        =   19
         Top             =   3630
         Width           =   930
      End
      Begin VB.TextBox TxtTax 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   975
         TabIndex        =   16
         Top             =   3630
         Width           =   555
      End
      Begin VB.ComboBox cmbfullpack 
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
         ItemData        =   "frmitemmastermini.frx":05AA
         Left            =   4170
         List            =   "frmitemmastermini.frx":05F3
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3210
         Width           =   1020
      End
      Begin VB.ComboBox CmbPack 
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
         ItemData        =   "frmitemmastermini.frx":0676
         Left            =   2505
         List            =   "frmitemmastermini.frx":06BF
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3195
         Width           =   1005
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00400000&
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
         Left            =   6630
         MaskColor       =   &H80000007&
         TabIndex        =   28
         Top             =   6105
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00400000&
         Caption         =   "&CANCEL"
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
         Left            =   5310
         MaskColor       =   &H80000007&
         TabIndex        =   27
         Top             =   6600
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00400000&
         Caption         =   "&SAVE"
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
         Left            =   5310
         MaskColor       =   &H80000007&
         TabIndex        =   26
         Top             =   6105
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.TextBox TxtMinQty 
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   975
         TabIndex        =   10
         Top             =   3165
         Width           =   555
      End
      Begin VB.TextBox txtcategory 
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
         ForeColor       =   &H00004080&
         Height          =   390
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1050
         Width           =   2895
      End
      Begin VB.CheckBox chknewcategory 
         Caption         =   "N&ew Category"
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
         Height          =   210
         Left            =   1050
         TabIndex        =   8
         Top             =   2805
         Width           =   1725
      End
      Begin VB.CommandButton CMDDELETE 
         BackColor       =   &H00400000&
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
         Left            =   6615
         MaskColor       =   &H80000007&
         TabIndex        =   30
         Top             =   6600
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.TextBox TXTITEM 
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
         Left            =   2040
         TabIndex        =   2
         Top             =   165
         Width           =   5910
      End
      Begin MSDataListLib.DataList Datacategory 
         Height          =   1320
         Left            =   1050
         TabIndex        =   29
         Top             =   1455
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2328
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16512
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
      Begin VB.Frame FrmeCompany 
         BorderStyle     =   0  'None
         Height          =   2070
         Left            =   4020
         TabIndex        =   3
         Top             =   945
         Width           =   3945
         Begin VB.CheckBox chknewcomp 
            Caption         =   "&New Company"
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
            Height          =   210
            Left            =   1035
            TabIndex        =   5
            Top             =   1860
            Width           =   1695
         End
         Begin VB.TextBox txtcompany 
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
            ForeColor       =   &H00004080&
            Height          =   405
            Left            =   1035
            MaxLength       =   25
            TabIndex        =   4
            Top             =   105
            Width           =   2895
         End
         Begin MSDataListLib.DataList Datacompany 
            Height          =   1320
            Left            =   1035
            TabIndex        =   6
            Top             =   525
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   2328
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16512
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
         Begin VB.Label LBLLP 
            Height          =   375
            Left            =   2760
            TabIndex        =   43
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
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
            Index           =   4
            Left            =   0
            TabIndex        =   7
            Top             =   210
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Offer Details"
         ForeColor       =   &H000000C0&
         Height          =   1440
         Left            =   2340
         TabIndex        =   70
         Top             =   6990
         Width           =   5640
         Begin VB.TextBox txtPrice4 
            Appearance      =   0  'Flat
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
            Height          =   420
            Left            =   4440
            MaxLength       =   8
            TabIndex        =   82
            Top             =   705
            Width           =   1110
         End
         Begin VB.TextBox txtPrice3 
            Appearance      =   0  'Flat
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
            Height          =   420
            Left            =   3045
            MaxLength       =   8
            TabIndex        =   81
            Top             =   705
            Width           =   1095
         End
         Begin VB.TextBox txtPrice2 
            Appearance      =   0  'Flat
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
            Height          =   420
            Left            =   1740
            MaxLength       =   8
            TabIndex        =   80
            Top             =   705
            Width           =   1095
         End
         Begin VB.OptionButton OptQtyOffer 
            Caption         =   "Qty"
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
            Left            =   30
            TabIndex        =   79
            Top             =   795
            Width           =   810
         End
         Begin VB.TextBox TxtOffer2 
            Appearance      =   0  'Flat
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
            Height          =   420
            Left            =   3045
            MaxLength       =   8
            TabIndex        =   76
            Top             =   225
            Width           =   1095
         End
         Begin VB.TextBox TxtOfferPrice 
            Appearance      =   0  'Flat
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
            Height          =   420
            Left            =   4440
            MaxLength       =   8
            TabIndex        =   74
            Top             =   225
            Width           =   1110
         End
         Begin VB.TextBox TxtOffer 
            Appearance      =   0  'Flat
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
            Height          =   420
            Left            =   1740
            MaxLength       =   8
            TabIndex        =   73
            Top             =   225
            Width           =   1095
         End
         Begin VB.OptionButton OptQty 
            Caption         =   "On Qty >="
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
            Left            =   30
            TabIndex        =   72
            Top             =   525
            Width           =   1710
         End
         Begin VB.OptionButton Optvalue 
            Caption         =   "Total Bill Value >"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   15
            TabIndex        =   71
            Top             =   270
            Value           =   -1  'True
            Width           =   1830
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "for 4 && above"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Index           =   23
            Left            =   4470
            TabIndex        =   85
            Top             =   1140
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Price for 3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Index           =   22
            Left            =   3105
            TabIndex        =   84
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Price for 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Index           =   21
            Left            =   1800
            TabIndex        =   83
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2895
            TabIndex        =   77
            Top             =   330
            Width           =   180
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Rs."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   20
            Left            =   4170
            TabIndex        =   75
            Top             =   315
            Width           =   300
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1800
         Left            =   30
         TabIndex        =   57
         Top             =   4815
         Width           =   2835
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1680
            Left            =   0
            Top             =   90
            Width           =   2805
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   25
         Left            =   3885
         TabIndex        =   91
         Top             =   4845
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UQC"
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
         Left            =   6105
         TabIndex        =   89
         Top             =   4845
         Width           =   420
      End
      Begin MSForms.ComboBox Txtmalay 
         Height          =   390
         Left            =   4935
         TabIndex        =   87
         Top             =   5670
         Width           =   2985
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5265;688"
         MatchEntry      =   1
         SpecialEffect   =   0
         FontName        =   "Kartika"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LBLPACKDESC 
         BackStyle       =   0  'Transparent
         Caption         =   "Pack Details"
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
         Height          =   435
         Index           =   0
         Left            =   6810
         TabIndex        =   69
         Top             =   3165
         Width           =   1020
      End
      Begin VB.Label LBLPACKDESC 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Items in Box"
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
         Height          =   435
         Index           =   20
         Left            =   5190
         TabIndex        =   68
         Top             =   3165
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   19
         Left            =   3885
         TabIndex        =   67
         Top             =   5280
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Malayalam"
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
         Height          =   240
         Index           =   18
         Left            =   3885
         TabIndex        =   65
         Top             =   5715
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Specification"
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
         Index           =   17
         Left            =   75
         TabIndex        =   64
         Top             =   690
         Width           =   1995
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   16
         Left            =   3075
         TabIndex        =   62
         Top             =   3690
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loc /Remarks"
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
         Index           =   15
         Left            =   5325
         TabIndex        =   61
         Top             =   4185
         Width           =   1350
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "HSN"
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
         Index           =   14
         Left            =   3210
         TabIndex        =   60
         Top             =   4185
         Width           =   405
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(Size 150 x 250 Pix)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   34
         Left            =   1980
         TabIndex        =   56
         Top             =   6645
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "L. Price"
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
         Index           =   13
         Left            =   1560
         TabIndex        =   52
         Top             =   4185
         Width           =   780
      End
      Begin VB.Label lblPack 
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
         Height          =   315
         Left            =   2130
         TabIndex        =   51
         Top             =   5055
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Tax"
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
         Index           =   12
         Left            =   45
         TabIndex        =   50
         Top             =   3690
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   11
         Left            =   60
         TabIndex        =   49
         Top             =   4185
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "W. Price"
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
         Index           =   10
         Left            =   6135
         TabIndex        =   48
         Top             =   3690
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "R. Price"
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
         Index           =   9
         Left            =   4440
         TabIndex        =   47
         Top             =   3690
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   7
         Left            =   1560
         TabIndex        =   46
         Top             =   3690
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   3495
         TabIndex        =   44
         Top             =   3240
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   1530
         TabIndex        =   42
         Top             =   3225
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-order"
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
         Index           =   6
         Left            =   45
         TabIndex        =   33
         Top             =   3225
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   45
         TabIndex        =   32
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
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
         Left            =   75
         TabIndex        =   31
         Top             =   255
         Width           =   1995
      End
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   1320
      TabIndex        =   35
      Top             =   810
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2858
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
   Begin MSDataGridLib.DataGrid grdtmp 
      Height          =   9075
      Left            =   8055
      TabIndex        =   66
      Top             =   30
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   16007
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
   Begin VB.Label LBLITEMNAME 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   480
      Left            =   8910
      TabIndex        =   45
      Top             =   120
      Width           =   5205
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   6390
      TabIndex        =   41
      Top             =   630
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   7500
      TabIndex        =   40
      Top             =   465
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   7290
      TabIndex        =   39
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   6420
      TabIndex        =   38
      Top             =   60
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM CODE"
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
      Left            =   195
      TabIndex        =   37
      Top             =   435
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F3 to Search...... Press Esc to Exit.."
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
      Index           =   8
      Left            =   1320
      TabIndex        =   36
      Top             =   0
      Width           =   3960
   End
End
Attribute VB_Name = "frmitemmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytData() As Byte
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim REPFLAG As Boolean
Dim COMPANYFLAG As Boolean
Dim CATEGORYFLAG As Boolean
Dim RSTREP As New ADODB.Recordset
Dim RSTCOMPANY As New ADODB.Recordset
Dim RSTCATEGORY As New ADODB.Recordset

Private Sub chknewcategory_Click()
    On Error Resume Next
    txtcategory.SetFocus
End Sub

Private Sub chknewcomp_Click()
    On Error Resume Next
    txtcompany.SetFocus
End Sub

Private Sub cmbfullpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            On Error Resume Next
            If cmbfullpack.ListIndex = -1 Then cmbfullpack.Text = CmbPack.Text
            Txtpackdes.SetFocus
        Case vbKeyEscape
            CmbPack.SetFocus
    End Select
End Sub

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            cmbfullpack.SetFocus
        Case vbKeyEscape
            TxtMinQty.SetFocus
    End Select
End Sub

Private Sub CmbPack_LostFocus()
    LblPack.Caption = CmbPack.Text
End Sub

Private Sub CmbQAC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtBarcode.SetFocus
        Case vbKeyEscape
            TxtExpense.SetFocus
    End Select
End Sub

Private Sub cmdcancel_Click()
    
    Set Image1.DataSource = Nothing
    Image1.Picture = LoadPicture("")
    TXTPRODUCT.Text = ""
    TXTITEM.Text = ""
    TxtSpec.Text = ""
    lblitemname.Caption = ""
    txtcategory.Text = ""
    txtcompany.Text = ""
    TxtMinQty.Text = ""
    TXTTAX.Text = ""
    txtHSN.Text = ""
    TxtOffer.Text = ""
    TxtOffer2.Text = ""
    TxtOfferPrice.Text = ""
    txtPrice2.Text = ""
    txtPrice3.Text = ""
    txtPrice4.Text = ""
    TxtBarcode.Text = ""
    TxtLocation.Text = ""
    TxtMalay.Text = ""
    TxtCost.Text = ""
    txtRT.Text = ""
    txtWS.Text = ""
    TxtMRP.Text = ""
    TxtExpense.Text = ""
    TxtLPack.Text = ""
    txtLPrice.Text = ""
    Txtpackdes.Text = ""
    txtpackdet.Text = ""
    CmbPack.ListIndex = -1
    cmbfullpack.ListIndex = -1
    CmbQAC.ListIndex = -1
    Set DataList2.RowSource = Nothing
    TXTITEMCODE.Enabled = True
    DataList2.Enabled = True
    FRAME.Visible = False
    TXTPRODUCT.Visible = False
    DataList2.Visible = False
    TXTITEMCODE.SetFocus
    chknewcategory.Value = 0
    chknewcomp.Value = 0
    ChkDead.Value = 0
    chkfocusQty.Value = 0
    ChkFreeWarn.Value = 0
    chkunbill.Value = 0
    Optvalue.Value = True
    If PC_FLAG = "Y" Then
        cHKcHANGE.Value = 1
    Else
        cHKcHANGE.Value = 0
    End If
    
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        If IsNull(TRXMAST.Fields(0)) Then
            TXTITEMCODE.Text = 1
        Else
            TXTITEMCODE.Text = Val(TRXMAST.Fields(0)) + 1
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If MDIMAIN.StatusBar.Panels(9).Text = "Y" Then Exit Sub
    Dim RSTITEMMAST As ADODB.Recordset
    On Error Resume Next
    If Val(Txtpackdes.Text) = 0 Then Txtpackdes.Text = 1
    If Val(txtpackdet.Text) = 0 Then txtpackdet.Text = 1
    If TXTITEM.Text = "" Then
        MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
        TXTITEM.SetFocus
        Exit Sub
    End If
    If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then CmbQAC.Text = "OTH"
    
    If MDIMAIN.lblgst.Caption = "R" And CmbQAC.ListIndex = -1 Then
        MsgBox "Please select Unique Quantity Code. It is mandatory for GST", vbOKOnly, "PRODUCT MASTER"
        CmbQAC.SetFocus
        Exit Sub
    End If
'    If txtcompany.Visible = True Then
'        If UCase(Datacategory.BoundText) <> "SERVICE CHARGE" And Trim(txtcompany.text) = "" Then
'            MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
'            txtcompany.SetFocus
'            Exit Sub
'        End If
'        If UCase(Datacategory.BoundText) <> "SERVICE CHARGE" And chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
'            MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
'            txtcompany.SetFocus
'            Exit Sub
'        End If
'    End If
'    If Trim(txtcategory.text) = "" Then
'        MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
'        txtcategory.SetFocus
'        Exit Sub
'    End If
'
'    If chknewcategory.Value = 0 And Datacategory.BoundText = "" Then
'        MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
'        txtcategory.SetFocus
'        Exit Sub
'    End If
    
    If Val(TxtOffer.Text) = 0 And Val(TxtOfferPrice.Text) <> 0 Then
        MsgBox "Please enter offer details correctly", vbOKOnly, "PRODUCT MASTER"
        TxtOffer.SetFocus
        Exit Sub
    End If
    
    If Val(TxtOfferPrice.Text) = 0 And Val(TxtOffer.Text) <> 0 Then
        MsgBox "Please enter offer details correctly", vbOKOnly, "PRODUCT MASTER"
        TxtOfferPrice.SetFocus
        Exit Sub
    End If
    
    If OptQty.Value = True And Val(TxtOffer.Text) <> 0 And Val(TxtOffer.Text) <= 1 Then
        MsgBox "Please enter offer details correctly", vbOKOnly, "PRODUCT MASTER"
        TxtOfferPrice.SetFocus
        Exit Sub
    End If
    
    If Optvalue.Value = True And Val(TxtOffer.Text) <> 0 And Val(TxtOffer.Text) <= 100 Then
        MsgBox "Please enter offer details correctly", vbOKOnly, "PRODUCT MASTER"
        TxtOfferPrice.SetFocus
        Exit Sub
    End If
    
    If Val(TxtOffer.Text) <> 0 And Val(TxtOfferPrice.Text) >= Val(txtRT.Text) Then
        MsgBox "Please enter offer details correctly", vbOKOnly, "PRODUCT MASTER"
        TxtOfferPrice.SetFocus
        Exit Sub
    End If
    
    If OptQtyOffer.Value = True And Val(txtPrice2.Text) = 0 And Val(txtPrice3.Text) = 0 And Val(txtPrice4.Text) = 0 Then
    ElseIf OptQtyOffer.Value = True And (Val(txtPrice2.Text) = 0 Or Val(txtPrice3.Text) = 0 Or Val(txtPrice4.Text) = 0) Then
        MsgBox "Please enter offer details correctly", vbOKOnly, "PRODUCT MASTER"
        txtPrice2.SetFocus
        Exit Sub
    End If
    On Error GoTo ERRHAND
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME = '" & Trim(TXTITEM.Text) & "' AND ITEM_CODE <> '" & Trim(TxtItemcode.Text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        MsgBox "The Item name already exists...", vbOKOnly, "Item Master"
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
'        Exit Sub
'    End If
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.Properties("Update Criteria").Value = adCriteriaKey
        RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        If chknewcategory.Value = 1 Then RSTITEMMAST!Category = txtcategory.Text Else RSTITEMMAST!Category = Datacategory.BoundText
        RSTITEMMAST!UNIT = 1
        If chknewcomp.Value = 1 Then RSTITEMMAST!MANUFACTURER = txtcompany.Text Else RSTITEMMAST!MANUFACTURER = Datacompany.BoundText
        RSTITEMMAST!REORDER_QTY = Val(TxtMinQty.Text)
        RSTITEMMAST!PACK_TYPE = CmbPack.Text
        RSTITEMMAST!FULL_PACK = cmbfullpack.Text
        RSTITEMMAST!UQC = CmbQAC.Text
        RSTITEMMAST!BIN_LOCATION = Trim(TxtLocation.Text)
        RSTITEMMAST!ITEM_MAL = Trim(TxtMalay.Text)
        RSTITEMMAST!SALES_TAX = Val(TXTTAX.Text)
        RSTITEMMAST!REMARKS = Trim(txtHSN.Text)
        RSTITEMMAST!ITEM_SPEC = Trim(TxtSpec.Text)
        If Val(TXTTAX.Text) > 0 Then RSTITEMMAST!check_flag = "V"
        RSTITEMMAST!item_COST = Val(TxtCost.Text)
        RSTITEMMAST!P_RETAIL = Val(txtRT.Text)
        RSTITEMMAST!MRP = Val(TxtMRP.Text)
        RSTITEMMAST!ITEM_EXPENSE = Val(TxtExpense.Text)
        RSTITEMMAST!P_WS = Val(txtWS.Text)
        RSTITEMMAST!CRTN_PACK = IIf(Val(TxtLPack.Text) = 0, 1, Val(TxtLPack.Text))
        RSTITEMMAST!P_CRTN = Val(txtLPrice.Text)
        RSTITEMMAST!LOOSE_PACK = IIf(Val(Txtpack.Text) = 0, 1, Val(Txtpack.Text))
        RSTITEMMAST!PACK_DESC = Val(Txtpackdes.Text)
        RSTITEMMAST!PACK_DET = Val(txtpackdet.Text)
        RSTITEMMAST!BARCODE = Trim(TxtBarcode.Text)
        If ChkDead.Value = 0 Then
            RSTITEMMAST!DEAD_STOCK = "N"
        Else
            RSTITEMMAST!DEAD_STOCK = "Y"
        End If
        If chkfocusQty.Value = 0 Then
            RSTITEMMAST!FOCUS_FLAG = "N"
        Else
            RSTITEMMAST!FOCUS_FLAG = "Y"
        End If
        If ChkFreeWarn.Value = 0 Then
            RSTITEMMAST!FREE_WARN = "N"
        Else
            RSTITEMMAST!FREE_WARN = "Y"
        End If
        If chkunbill.Value = 0 Then
            RSTITEMMAST!UN_BILL = "N"
        Else
            RSTITEMMAST!UN_BILL = "Y"
        End If
        If cHKcHANGE.Value = 0 Then
            RSTITEMMAST!PRICE_CHANGE = "N"
        Else
            RSTITEMMAST!PRICE_CHANGE = "Y"
        End If
        If Optvalue.Value = True And Val(TxtOffer.Text) <> 0 Then
            RSTITEMMAST!OFFER_FLAG = "1"
            RSTITEMMAST!OFFER_VALUE = Val(TxtOffer.Text)
            RSTITEMMAST!OFFER_VALUE2 = Val(TxtOffer2.Text)
            RSTITEMMAST!OFFER_PRICE = Val(TxtOfferPrice.Text)
        ElseIf OptQty.Value = True And Val(TxtOffer.Text) <> 0 Then
            RSTITEMMAST!OFFER_FLAG = "2"
            RSTITEMMAST!OFFER_VALUE = Val(TxtOffer.Text)
            RSTITEMMAST!OFFER_VALUE2 = Val(TxtOffer2.Text)
            RSTITEMMAST!OFFER_PRICE = Val(TxtOfferPrice.Text)
        ElseIf OptQtyOffer.Value = True Then
            RSTITEMMAST!OFFER_FLAG = "3"
            RSTITEMMAST!OFFER_VALUE = Val(txtPrice2.Text)
            RSTITEMMAST!OFFER_VALUE2 = Val(txtPrice3.Text)
            RSTITEMMAST!OFFER_PRICE = Val(txtPrice4.Text)
        Else
            RSTITEMMAST!OFFER_FLAG = "0"
            RSTITEMMAST!OFFER_VALUE = 0
            RSTITEMMAST!OFFER_VALUE2 = 0
            RSTITEMMAST!OFFER_PRICE = 0
        End If
        
        RSTITEMMAST.Update
    Else
        RSTITEMMAST.AddNew
        RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
        RSTITEMMAST!ITEM_CODE = Trim(TXTITEMCODE.Text)
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        If chknewcategory.Value = 1 Then RSTITEMMAST!Category = txtcategory.Text Else RSTITEMMAST!Category = Datacategory.BoundText
        RSTITEMMAST!UNIT = 1
        If chknewcomp.Value = 1 Then RSTITEMMAST!MANUFACTURER = txtcompany.Text Else RSTITEMMAST!MANUFACTURER = Trim(Datacompany.BoundText)
        If ChkDead.Value = 0 Then
            RSTITEMMAST!DEAD_STOCK = "N"
        Else
            RSTITEMMAST!DEAD_STOCK = "Y"
        End If
        If chkfocusQty.Value = 0 Then
            RSTITEMMAST!FOCUS_FLAG = "N"
        Else
            RSTITEMMAST!FOCUS_FLAG = "Y"
        End If
        If ChkFreeWarn.Value = 0 Then
            RSTITEMMAST!FREE_WARN = "N"
        Else
            RSTITEMMAST!FREE_WARN = "Y"
        End If
        If chkunbill.Value = 0 Then
            RSTITEMMAST!UN_BILL = "N"
        Else
            RSTITEMMAST!UN_BILL = "Y"
        End If
        If cHKcHANGE.Value = 0 Then
            RSTITEMMAST!PRICE_CHANGE = "N"
        Else
            RSTITEMMAST!PRICE_CHANGE = "Y"
        End If
        
        RSTITEMMAST!REMARKS = Trim(txtHSN.Text)
        RSTITEMMAST!ITEM_SPEC = Trim(TxtSpec.Text)
        RSTITEMMAST!REORDER_QTY = Val(TxtMinQty.Text)
        If CmbPack.ListIndex = -1 Then
            RSTITEMMAST!PACK_TYPE = "Nos"
        Else
            RSTITEMMAST!PACK_TYPE = CmbPack.Text
        End If
        If cmbfullpack.ListIndex = -1 Then
            RSTITEMMAST!FULL_PACK = "Nos"
        Else
            RSTITEMMAST!FULL_PACK = CmbPack.Text
        End If
        RSTITEMMAST!UQC = CmbQAC.Text
        RSTITEMMAST!BIN_LOCATION = Trim(TxtLocation.Text)
        RSTITEMMAST!ITEM_MAL = Trim(TxtMalay.Text)
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
        RSTITEMMAST!SALES_TAX = Val(TXTTAX.Text)
        If Val(TXTTAX.Text) > 0 Then RSTITEMMAST!check_flag = "V"
        RSTITEMMAST!item_COST = Val(TxtCost.Text)
        RSTITEMMAST!P_RETAIL = Val(txtRT.Text)
        RSTITEMMAST!MRP = Val(TxtMRP.Text)
        RSTITEMMAST!ITEM_EXPENSE = Val(TxtExpense.Text)
        RSTITEMMAST!P_WS = Val(txtWS.Text)
        RSTITEMMAST!CRTN_PACK = IIf(Val(TxtLPack.Text) = 0, 1, Val(TxtLPack.Text))
        RSTITEMMAST!P_CRTN = Val(txtLPrice.Text)
        RSTITEMMAST!LOOSE_PACK = IIf(Val(Txtpack.Text) = 0, 1, Val(Txtpack.Text))
        RSTITEMMAST!PACK_DESC = Val(Txtpackdes.Text)
        RSTITEMMAST!PACK_DET = Val(txtpackdet.Text)
        RSTITEMMAST!BARCODE = Trim(TxtBarcode.Text)
        If Optvalue.Value = True And Val(TxtOffer.Text) <> 0 Then
            RSTITEMMAST!OFFER_FLAG = "1"
            RSTITEMMAST!OFFER_VALUE = Val(TxtOffer.Text)
            RSTITEMMAST!OFFER_VALUE2 = Val(TxtOffer2.Text)
            RSTITEMMAST!OFFER_PRICE = Val(TxtOfferPrice.Text)
        ElseIf OptQty.Value = True And Val(TxtOffer.Text) <> 0 Then
            RSTITEMMAST!OFFER_FLAG = "2"
            RSTITEMMAST!OFFER_VALUE = Val(TxtOffer.Text)
            RSTITEMMAST!OFFER_VALUE2 = Val(TxtOffer2.Text)
            RSTITEMMAST!OFFER_PRICE = Val(TxtOfferPrice.Text)
        ElseIf OptQtyOffer.Value = True Then
            RSTITEMMAST!OFFER_FLAG = "3"
            RSTITEMMAST!OFFER_VALUE = Val(txtPrice2.Text)
            RSTITEMMAST!OFFER_VALUE2 = Val(txtPrice3.Text)
            RSTITEMMAST!OFFER_PRICE = Val(txtPrice4.Text)
        Else
            RSTITEMMAST!OFFER_FLAG = "0"
            RSTITEMMAST!OFFER_VALUE = 0
            RSTITEMMAST!OFFER_VALUE2 = 0
            RSTITEMMAST!OFFER_PRICE = 0
        End If
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT MANUFACTURER FROM MANUFACT WHERE MANUFACTURER = '" & Trim(txtcompany.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!MANUFACTURER = Trim(txtcompany.Text)
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT CATEGORY FROM CATEGORY WHERE CATEGORY = '" & Trim(txtcategory.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!Category = Trim(txtcategory.Text)
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTITEMMAST.EOF
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        RSTITEMMAST!MFGR = Trim(txtcompany.Text)
        RSTITEMMAST!Category = Trim(txtcategory.Text)
        If chkfocusQty.Value = 0 Then
            RSTITEMMAST!FOCUS_FLAG = "N"
        Else
            RSTITEMMAST!FOCUS_FLAG = "Y"
        End If
        RSTITEMMAST.Update
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
'    db.Execute "Update RTRXFILE set ITEM_NAME = '" & Trim(TXTITEM.Text) & "', MFGR = '" & Trim(txtcompany.Text) & "', Category = '" & Trim(txtcategory.Text) & "' where ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'"
    db.Execute "Update trxfile set Category = '" & Trim(txtcategory.Text) & "' where ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'"

                    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    cmdcancel_Click
Exit Sub
ERRHAND:
    MsgBox (err.Description)
        
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtcategory.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ERRHAND
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT ITEM_CODE FROM ITEMMAST WHERE ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                TXTITEMCODE.Text = RSTITEMMAST!ITEM_CODE
            End If
            TXTPRODUCT.Visible = False
            DataList2.Visible = False
            TXTITEMCODE.SetFocus
        Case vbKeyEscape
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox (err.Description)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call txtcompany_Change
    TXTITEMCODE.SetFocus
End Sub

Private Sub Form_Load()
    
    PHYFLAG = True
    REPFLAG = True
    COMPANYFLAG = True
    CATEGORYFLAG = True
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        If IsNull(TRXMAST.Fields(0)) Then
            TXTITEMCODE.Text = 1
        Else
            TXTITEMCODE.Text = Val(TRXMAST.Fields(0)) + 1
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    If PC_FLAG = "Y" Then
        cHKcHANGE.Value = 1
    Else
        cHKcHANGE.Value = 0
    End If
    
    Call txtcategory_Change
    'TMPFLAG = True
    'Width = 8385
    'Height = 4575
    Left = 3500
    Top = 0
    FRAME.Visible = False
    'txtunit.Visible = False
    
    

    Exit Sub
ERRHAND:
    MsgBox (err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    If COMPANYFLAG = False Then RSTCOMPANY.Close
    If CATEGORYFLAG = False Then RSTCATEGORY.Close
    If PHYFLAG = False Then PHY.Close
    'If TMPFLAG = False Then rstTMP.Close
    'MDIMAIN.Enabled = True
    'FrmCrimedata.Enabled = True
End Sub

Private Sub grdtmp_DblClick()
    On Error Resume Next
    If FRAME.Visible = True Then
        TXTITEM.Text = Trim(grdtmp.Columns(1))
        TXTITEM.SetFocus
    End If
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            On Error Resume Next
            TXTITEM.Text = Trim(grdtmp.Columns(1))
            TXTITEM.SetFocus
    End Select
End Sub

Private Sub TxtCost_GotFocus()
    TxtCost.SelStart = 0
    TxtCost.SelLength = Len(TxtCost.Text)
End Sub

Private Sub TxtCost_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtMRP.SetFocus
        Case vbKeyEscape
            TXTTAX.SetFocus
    End Select
End Sub

Private Sub TxtCost_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCost_LostFocus()
    TxtCost.Text = Format(Val(TxtCost.Text), "0.00")
End Sub

Private Sub TXTITEM_Change()
    On Error GoTo ERRHAND
    Set grdtmp.DataSource = Nothing
    If PHYFLAG = True Then
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    Else
        PHY.Close
        'PHY.Open "Select ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' AND ITEM_CODE <> '" & Trim(TXTITEMCODE.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    End If
    Set grdtmp.DataSource = PHY
    grdtmp.Columns(0).Caption = "Code"
    'grdtmp.Columns(8).Caption = ""
    
    grdtmp.Columns(0).Width = 1000
    grdtmp.Columns(1).Width = 3800
    grdtmp.Columns(2).Width = 1200
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTITEM_GotFocus()
    TXTITEM.SelStart = 0
    TXTITEM.SelLength = Len(TXTITEM.Text)
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTITEM.Text = "" Then
                MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
                TXTITEM.SetFocus
                Exit Sub
            End If
            TxtSpec.SetFocus
    End Select
End Sub

Private Sub TXTITEM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTITEMCODE_Change()
    On Error GoTo ERRHAND
    Set grdtmp.DataSource = Nothing
    If PHYFLAG = True Then
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_CODE Like '" & Trim(Me.TXTITEMCODE.Text) & "%' ORDER BY ITEM_CODE ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    Else
        PHY.Close
        'PHY.Open "Select ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' AND ITEM_CODE <> '" & Trim(TXTITEMCODE.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_CODE Like '" & Trim(Me.TXTITEMCODE.Text) & "%' ORDER BY ITEM_CODE ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    End If
    Set grdtmp.DataSource = PHY
    grdtmp.Columns(0).Caption = "Code"
    'grdtmp.Columns(8).Caption = ""
    
    grdtmp.Columns(0).Width = 1000
    grdtmp.Columns(1).Width = 3800
    grdtmp.Columns(2).Width = 1200
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TxtItemcode_GotFocus()
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
End Sub

Public Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ERRHAND
            If Trim(TXTITEMCODE.Text) = "" Then Exit Sub
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                On Error Resume Next
                Set Image1.DataSource = RSTITEMMAST
                If err.Number = 545 Then
                    Set Image1.DataSource = Nothing
                    bytData = ""
                Else
                    Set Image1.DataSource = RSTITEMMAST 'setting image1’s datasource
                    Image1.DataField = "PHOTO"
                    bytData = RSTITEMMAST!PHOTO
                End If
                On Error GoTo ERRHAND
            
                TXTITEM.Text = IIf(IsNull(RSTITEMMAST!ITEM_NAME), "", RSTITEMMAST!ITEM_NAME)
                txtcategory.Text = IIf(IsNull(RSTITEMMAST!Category), "", RSTITEMMAST!Category)
                txtcompany.Text = IIf(IsNull(RSTITEMMAST!MANUFACTURER), "", RSTITEMMAST!MANUFACTURER)
                TxtMinQty.Text = IIf(IsNull(RSTITEMMAST!REORDER_QTY), 0, RSTITEMMAST!REORDER_QTY)
                TXTTAX.Text = IIf(IsNull(RSTITEMMAST!SALES_TAX), 0, RSTITEMMAST!SALES_TAX)
                txtHSN.Text = IIf(IsNull(RSTITEMMAST!REMARKS), "", RSTITEMMAST!REMARKS)
                TxtBarcode.Text = IIf(IsNull(RSTITEMMAST!BARCODE), "", RSTITEMMAST!BARCODE)
                TxtSpec.Text = IIf(IsNull(RSTITEMMAST!ITEM_SPEC), "", RSTITEMMAST!ITEM_SPEC)
                TxtLocation.Text = IIf(IsNull(RSTITEMMAST!BIN_LOCATION), "", RSTITEMMAST!BIN_LOCATION)
                TxtMalay.Text = IIf(IsNull(RSTITEMMAST!ITEM_MAL), "", RSTITEMMAST!ITEM_MAL)
                TxtCost.Text = IIf(IsNull(RSTITEMMAST!item_COST), 0, RSTITEMMAST!item_COST)
                txtRT.Text = IIf(IsNull(RSTITEMMAST!P_RETAIL), 0, RSTITEMMAST!P_RETAIL)
                txtWS.Text = IIf(IsNull(RSTITEMMAST!P_WS), 0, RSTITEMMAST!P_WS)
                TxtMRP.Text = IIf(IsNull(RSTITEMMAST!MRP), 0, RSTITEMMAST!MRP)
                TxtExpense.Text = IIf(IsNull(RSTITEMMAST!ITEM_EXPENSE), 0, RSTITEMMAST!ITEM_EXPENSE)
                TxtLPack.Text = IIf(IsNull(RSTITEMMAST!CRTN_PACK), 1, RSTITEMMAST!CRTN_PACK)
                txtLPrice.Text = IIf(IsNull(RSTITEMMAST!P_CRTN), 0, RSTITEMMAST!P_CRTN)
                Txtpackdes.Text = IIf(IsNull(RSTITEMMAST!PACK_DESC), 1, RSTITEMMAST!PACK_DESC)
                txtpackdet.Text = IIf(IsNull(RSTITEMMAST!PACK_DET), 1, RSTITEMMAST!PACK_DET)
                Txtpack.Text = IIf(IsNull(RSTITEMMAST!LOOSE_PACK), 1, RSTITEMMAST!LOOSE_PACK)
                If IsNull(RSTITEMMAST!DEAD_STOCK) Or RSTITEMMAST!DEAD_STOCK = "N" Or RSTITEMMAST!DEAD_STOCK = "" Then
                    ChkDead.Value = 0
                Else
                    ChkDead.Value = 1
                End If
                If IsNull(RSTITEMMAST!FOCUS_FLAG) Or RSTITEMMAST!FOCUS_FLAG = "N" Or RSTITEMMAST!FOCUS_FLAG = "" Then
                    chkfocusQty.Value = 0
                Else
                    chkfocusQty.Value = 1
                End If
                If IsNull(RSTITEMMAST!FREE_WARN) Or RSTITEMMAST!FREE_WARN = "N" Or RSTITEMMAST!FREE_WARN = "" Then
                    ChkFreeWarn.Value = 0
                Else
                    ChkFreeWarn.Value = 1
                End If
                If IsNull(RSTITEMMAST!UN_BILL) Or RSTITEMMAST!UN_BILL = "N" Or RSTITEMMAST!UN_BILL = "" Then
                    chkunbill.Value = 0
                Else
                    chkunbill.Value = 1
                End If
                If IsNull(RSTITEMMAST!PRICE_CHANGE) Or RSTITEMMAST!PRICE_CHANGE = "N" Or RSTITEMMAST!PRICE_CHANGE = "" Then
                    cHKcHANGE.Value = 0
                Else
                    cHKcHANGE.Value = 1
                End If
                On Error Resume Next
                CmbPack.Text = IIf(IsNull(RSTITEMMAST!PACK_TYPE), 0, RSTITEMMAST!PACK_TYPE)
                cmbfullpack.Text = IIf(IsNull(RSTITEMMAST!FULL_PACK), 0, RSTITEMMAST!FULL_PACK)
                CmbQAC.Text = IIf(IsNull(RSTITEMMAST!UQC), 0, RSTITEMMAST!UQC)
                If RSTITEMMAST!OFFER_FLAG = "1" Then
                    Optvalue.Value = True
                    TxtOffer.Text = IIf(IsNull(RSTITEMMAST!OFFER_VALUE), "", RSTITEMMAST!OFFER_VALUE)
                    TxtOffer2.Text = IIf(IsNull(RSTITEMMAST!OFFER_VALUE2), "", RSTITEMMAST!OFFER_VALUE2)
                    TxtOfferPrice.Text = IIf(IsNull(RSTITEMMAST!OFFER_PRICE), "", RSTITEMMAST!OFFER_PRICE)
                ElseIf RSTITEMMAST!OFFER_FLAG = "2" Then
                    OptQty.Value = True
                    TxtOffer.Text = IIf(IsNull(RSTITEMMAST!OFFER_VALUE), "", RSTITEMMAST!OFFER_VALUE)
                    TxtOffer2.Text = IIf(IsNull(RSTITEMMAST!OFFER_VALUE2), "", RSTITEMMAST!OFFER_VALUE2)
                    TxtOfferPrice.Text = IIf(IsNull(RSTITEMMAST!OFFER_PRICE), "", RSTITEMMAST!OFFER_PRICE)
                ElseIf RSTITEMMAST!OFFER_FLAG = "3" Then
                    OptQtyOffer.Value = True
                    txtPrice2.Text = IIf(IsNull(RSTITEMMAST!OFFER_VALUE), "", RSTITEMMAST!OFFER_VALUE)
                    txtPrice3.Text = IIf(IsNull(RSTITEMMAST!OFFER_VALUE2), "", RSTITEMMAST!OFFER_VALUE2)
                    txtPrice4.Text = IIf(IsNull(RSTITEMMAST!OFFER_PRICE), "", RSTITEMMAST!OFFER_PRICE)
                Else
                    Optvalue.Value = True
                    TxtOffer.Text = ""
                    TxtOffer2.Text = ""
                    TxtOfferPrice.Text = ""
                    txtPrice2.Text = ""
                    txtPrice3.Text = ""
                    txtPrice4.Text = ""
                End If
                
                On Error GoTo ERRHAND
                Datacategory.Text = txtcategory.Text
                Call Datacategory_Click
                Datacompany.Text = txtcompany.Text
                Call Datacompany_Click
                If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
                If Val(TxtLPack.Text) = 0 Then TxtLPack.Text = 1
                If Val(txtLPrice.Text) = 0 Then txtLPrice.Text = Val(txtRT.Text) / Val(TxtLPack.Text)
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        
            TXTITEMCODE.Enabled = False
            FRAME.Visible = True
            TXTITEM.SetFocus
        Case 114
            TXTPRODUCT.Visible = True
            DataList2.Visible = True
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            Call CmdExit_Click
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

Private Sub TxtLPack_Change()
    txtLPrice.Text = ""
End Sub

Private Sub TxtLPack_GotFocus()
    If Val(TxtLPack.Text) = 0 Then
        TxtLPack.Text = 1
    End If
    TxtLPack.SelStart = 0
    TxtLPack.SelLength = Len(TxtLPack.Text)
End Sub

Private Sub TxtLPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtLPrice.SetFocus
        Case vbKeyEscape
            txtWS.SetFocus
    End Select
End Sub

Private Sub TxtLPack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtLPrice_GotFocus()
    If Val(TxtLPack.Text) = 0 Then TxtLPack.Text = "1"
    If Val(txtLPrice.Text) = 0 Then
        txtLPrice.Text = Format(Round(Val(txtRT.Text) / Val(TxtLPack.Text), 2), "0.00")
    End If
    txtLPrice.SelStart = 0
    txtLPrice.SelLength = Len(txtLPrice.Text)
End Sub

Private Sub txtLPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtHSN.SetFocus
        Case vbKeyEscape
            TxtLPack.SetFocus
    End Select
End Sub

Private Sub txtLPrice_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtLPrice_LostFocus()
    txtLPrice.Text = Format(Val(txtLPrice.Text), "0.00")
End Sub

Private Sub TxtMalay_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmdSave_Click
        Case vbKeyEscape
            TxtLocation.SetFocus
    End Select
End Sub

Private Sub TxtMalay_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtminqty_GotFocus()
    If TxtMinQty.Text = "" Then
        TxtMinQty.Text = 1
    End If
    TxtMinQty.SelStart = 0
    TxtMinQty.SelLength = Len(TxtMinQty.Text)
End Sub

Private Sub TxtMinQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txtpack.SetFocus
        Case vbKeyEscape
            If FrmeCompany.Visible = True Then
                txtcompany.SetFocus
            Else
                txtcategory.SetFocus
            End If
    End Select
End Sub

Private Sub TxtMinQty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtpack_GotFocus()
    If Val(Txtpack.Text) = 0 Then
        Txtpack.Text = 1
    End If
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.Text)
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmbPack.SetFocus
        Case vbKeyEscape
            TxtMinQty.SetFocus
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

Private Sub txtPrice2_GotFocus()
    OptQtyOffer.Value = True
End Sub

Private Sub txtPrice3_GotFocus()
    OptQtyOffer.Value = True
End Sub

Private Sub txtPrice4_GotFocus()
    OptQtyOffer.Value = True
End Sub

Private Sub TXTPRODUCT_Change()
    On Error GoTo ERRHAND
    Set grdtmp.DataSource = Nothing
    If REPFLAG = True Then
        RSTREP.Open "Select DISTINCT ITEM_CODE,ITEM_NAME From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select DISTINCT ITEM_CODE,ITEM_NAME From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    End If
    Set DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Set grdtmp.DataSource = RSTREP
    grdtmp.Columns(0).Caption = "Code"
    
    grdtmp.Columns(0).Width = 1000
    grdtmp.Columns(1).Width = 3800
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTPRODUCT.Text = "" Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            TXTPRODUCT.Visible = False
            DataList2.Visible = False
            TXTITEMCODE.SetFocus
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

Private Sub CmdDelete_Click()
    If MDIMAIN.StatusBar.Panels(9).Text = "Y" Then Exit Sub
    Dim rststock As ADODB.Recordset
    
    If TXTITEMCODE.Text = "" Then Exit Sub
    On Error GoTo ERRHAND
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from RTRXFILE where RTRXFILE.ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFILE where ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULASUB where FOR_NAME = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Transactions is Available in Formula", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULAMAST where ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Transactions is Available in Formula", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & TXTITEM.Text & " ****", vbYesNo, "DELETING ITEM....") = vbNo Then Exit Sub
    'db.Execute ("DELETE from RTRXFILE where RTRXFILE.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
    db.Execute ("DELETE from PRODLINK where PRODLINK.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
    db.Execute ("DELETE from ITEMMAST where ITEMMAST.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
    
    'tXTMEDICINE.Tag = tXTMEDICINE.Text
    'tXTMEDICINE.Text = ""
    'tXTMEDICINE.Text = tXTMEDICINE.Tag
    'TXTQTY.Text = ""
    MsgBox "ITEM " & TXTITEM.Text & "DELETED SUCCESSFULLY", vbInformation, "DELETING ITEM...."
    Call cmdcancel_Click
    Exit Sub
   
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub txtcategory_Change()

    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If CATEGORYFLAG = True Then
            RSTCATEGORY.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & txtcategory.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
            CATEGORYFLAG = False
        Else
            RSTCATEGORY.Close
            RSTCATEGORY.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & txtcategory.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
            CATEGORYFLAG = False
        End If
        If (RSTCATEGORY.EOF And RSTCATEGORY.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = RSTCATEGORY!Category
        End If
        Set Me.Datacategory.RowSource = RSTCATEGORY
        Datacategory.ListField = "CATEGORY"
        Datacategory.BoundColumn = "CATEGORY"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description

End Sub

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.Text)
End Sub

Private Sub txtcategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If chknewcategory.Value = 0 And Datacategory.BoundText = "" Then
                If Datacategory.VisibleCount = 0 Then Exit Sub
'                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
'                Txtcompany.SetFocus
                Datacategory.SetFocus
                Exit Sub
            Else
                If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
                    FrmeCompany.Visible = False
                    TxtMinQty.Visible = False
                    Txtpackdes.Visible = False
                    txtpackdet.Visible = False
                    CmbPack.Visible = False
                    cmbfullpack.Visible = False
                    Label1(6).Visible = False
                    CmdSave.SetFocus
                Else
                    FrmeCompany.Visible = True
                    TxtMinQty.Visible = True
                    CmbPack.Visible = True
                    cmbfullpack.Visible = True
                    Txtpackdes.Visible = True
                    txtpackdet.Visible = True
                    Label1(6).Visible = True
                    txtcompany.SetFocus
                End If
            End If
        Case vbKeyEscape
            TXTITEM.SetFocus
    End Select
End Sub

Private Sub txtcategory_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacategory_Click()

    txtcategory.Text = Datacategory.Text
    lbldealer.Caption = txtcategory.Text

    If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
        FrmeCompany.Visible = False
        TxtMinQty.Visible = False
        CmbPack.Visible = False
        cmbfullpack.Visible = False
        Label1(6).Visible = False
        Txtpackdes.Visible = False
        txtpackdet.Visible = False
    Else
        FrmeCompany.Visible = True
        TxtMinQty.Visible = True
        CmbPack.Visible = True
        cmbfullpack.Visible = True
        Label1(6).Visible = True
        Txtpackdes.Visible = True
        txtpackdet.Visible = True
    End If

End Sub

Private Sub Datacategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtcategory.Text = "" Then
                MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
                txtcategory.SetFocus
                Exit Sub
            End If

            If chknewcategory.Value = 0 And Datacategory.BoundText = "" Then
                MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
                txtcategory.SetFocus
                Exit Sub
            End If
            If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
                FrmeCompany.Visible = False
                TxtMinQty.Visible = False
                CmbPack.Visible = False
                cmbfullpack.Visible = False
                Label1(6).Visible = False
                Txtpackdes.Visible = False
                txtpackdet.Visible = False
                CmdSave.SetFocus
            Else
                FrmeCompany.Visible = True
                TxtMinQty.Visible = True
                CmbPack.Visible = True
                cmbfullpack.Visible = True
                Label1(6).Visible = True
                Txtpackdes.Visible = True
                txtpackdet.Visible = True
                txtcompany.SetFocus
            End If
        Case vbKeyEscape
            txtcategory.SetFocus
    End Select
End Sub

Private Sub Datacategory_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacategory_GotFocus()
    flagchange.Caption = 1
    txtcategory.Text = lbldealer.Caption
    Datacategory.Text = txtcategory.Text
    Call Datacategory_Click
End Sub

Private Sub Datacategory_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub txtcompany_Change()

    On Error GoTo ERRHAND
    If flagchange2.Caption <> "1" Then
        If COMPANYFLAG = True Then
            RSTCOMPANY.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & txtcompany.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            COMPANYFLAG = False
        Else
            RSTCOMPANY.Close
            RSTCOMPANY.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & txtcompany.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            COMPANYFLAG = False
        End If
        If (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
            LBLDEALER2.Caption = ""
        Else
            LBLDEALER2.Caption = RSTCOMPANY!MANUFACTURER
        End If
        Set Me.Datacompany.RowSource = RSTCOMPANY
        Datacompany.ListField = "MANUFACTURER"
        Datacompany.BoundColumn = "MANUFACTURER"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description

End Sub

Private Sub txtcompany_GotFocus()
    txtcompany.SelStart = 0
    txtcompany.SelLength = Len(txtcompany.Text)
End Sub

Private Sub txtcompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
                If Datacompany.VisibleCount = 0 Then Exit Sub
'                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
'                Txtcompany.SetFocus
                Datacompany.SetFocus
                Exit Sub
            Else
                TxtMinQty.SetFocus
            End If
        Case vbKeyEscape
            txtcategory.SetFocus
    End Select

End Sub

Private Sub txtcompany_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacompany_Click()

    txtcompany.Text = Datacompany.Text
    LBLDEALER2.Caption = txtcompany.Text

End Sub

Private Sub Datacompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtcompany.Text = "" Then
                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
                txtcompany.SetFocus
                Exit Sub
            End If
            If chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
                txtcompany.SetFocus
                Exit Sub
            End If

            TxtMinQty.SetFocus
        Case vbKeyEscape
            txtcompany.SetFocus
    End Select
End Sub

Private Sub Datacompany_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacompany_GotFocus()
    flagchange2.Caption = 1
    txtcompany.Text = LBLDEALER2.Caption
    Datacompany.Text = txtcompany.Text
    Call Datacompany_Click
End Sub

Private Sub Datacompany_LostFocus()
     flagchange2.Caption = ""
End Sub

Private Sub txtcategory_LostFocus()
    If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
        FrmeCompany.Visible = False
        TxtMinQty.Visible = False
        CmbPack.Visible = False
        cmbfullpack.Visible = False
        Label1(6).Visible = False
        Txtpackdes.Visible = False
        txtpackdet.Visible = False
    Else
        FrmeCompany.Visible = True
        CmbPack.Visible = True
        cmbfullpack.Visible = True
        TxtMinQty.Visible = True
        Label1(6).Visible = True
        Txtpackdes.Visible = True
        txtpackdet.Visible = True
    End If
End Sub

Private Sub txtRT_GotFocus()
    txtRT.SelStart = 0
    txtRT.SelLength = Len(txtRT.Text)
End Sub

Private Sub txtRT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtWS.SetFocus
        Case vbKeyEscape
            TxtMRP.SetFocus
    End Select
End Sub

Private Sub txtRT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtRT_LostFocus()
    txtRT.Text = Format(Val(txtRT.Text), "0.00")
End Sub

Private Sub TxtTax_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
End Sub

Private Sub TxtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtCost.SetFocus
        Case vbKeyEscape
            Txtpackdes.SetFocus
    End Select
End Sub

Private Sub TxtTax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTTAX_LostFocus()
    TXTTAX.Text = Format(Val(TXTTAX.Text), "0.00")
End Sub

Private Sub txtws_GotFocus()
    txtWS.SelStart = 0
    txtWS.SelLength = Len(txtWS.Text)
End Sub

Private Sub txtws_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtLPack.SetFocus
        Case vbKeyEscape
            txtRT.SetFocus
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
    txtWS.Text = Format(Val(txtWS.Text), "0.00")
End Sub

Private Sub cmddelphoto_Click()
        
    On Error GoTo errHandler
    CommonDialog1.FileName = ""
    Set Image1.DataSource = Nothing
    Image1.Picture = LoadPicture("")
    
    bytData = ""
    Exit Sub
errHandler:
    MsgBox "Unexpected error. Err " & err & " : " & Error
End Sub

Private Sub CMDBROWSE_Click()
    On Error GoTo errHandler
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "Picture Files (*.jpg)|*.jpg"
    CommonDialog1.ShowOpen
    Image1.Picture = LoadPicture(CommonDialog1.FileName)
    
    Open CommonDialog1.FileName For Binary As #1
    ReDim bytData(FileLen(CommonDialog1.FileName))
    
    Get #1, , bytData
    Close #1
    Exit Sub
errHandler:
    Select Case err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        MsgBox "Unexpected error. Err " & err & " : " & Error
    End Select
End Sub

Private Sub TxtHSN_GotFocus()
    txtHSN.SelStart = 0
    txtHSN.SelLength = Len(txtHSN.Text)
End Sub

Private Sub TxtHSN_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtLocation.SetFocus
        Case vbKeyEscape
            txtLPrice.SetFocus
    End Select
End Sub

Private Sub txtHSN_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLocation_GotFocus()
    TxtLocation.SelStart = 0
    TxtLocation.SelLength = Len(TxtLocation.Text)
End Sub

Private Sub TxtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtExpense.SetFocus
        Case vbKeyEscape
            txtLPrice.SetFocus
    End Select
End Sub

Private Sub TxtLocation_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
            txtRT.SetFocus
        Case vbKeyEscape
            TxtCost.SetFocus
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
    TxtMRP.Text = Format(Val(TxtMRP.Text), "0.00")
End Sub

Private Sub TxtExpense_GotFocus()
    TxtExpense.SelStart = 0
    TxtExpense.SelLength = Len(TxtExpense.Text)
End Sub

Private Sub TxtExpense_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmbQAC.SetFocus
        Case vbKeyEscape
            TxtLocation.SetFocus
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

Private Sub TxtExpense_LostFocus()
    TxtExpense.Text = Format(Val(TxtExpense.Text), "0.00")
End Sub

Private Sub TxtSpec_GotFocus()
    TxtSpec.SelStart = 0
    TxtSpec.SelLength = Len(TxtSpec.Text)
End Sub

Private Sub TxtSpec_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcategory.SetFocus
    End Select
End Sub

Private Sub TxtSpec_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtMalay_GotFocus()
    TxtMalay.SelStart = 0
    TxtMalay.SelLength = Len(TxtMalay.Text)
End Sub

Private Sub txtbarcode_GotFocus()
    TxtBarcode.SelStart = 0
    TxtBarcode.SelLength = Len(TxtBarcode.Text)
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmdSave_Click
        Case vbKeyEscape
            TxtLocation.SetFocus
    End Select
End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtpackdes_GotFocus()
    If Val(Txtpackdes.Text) = 0 Then
        Txtpackdes.Text = 1
    End If
    Txtpackdes.SelStart = 0
    Txtpackdes.SelLength = Len(Txtpackdes.Text)
End Sub

Private Sub Txtpackdes_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtpackdet.SetFocus
        Case vbKeyEscape
            cmbfullpack.SetFocus
    End Select
End Sub

Private Sub Txtpackdes_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtpackdet_GotFocus()
    If Val(txtpackdet.Text) = 0 Then
        txtpackdet.Text = 1
    End If
    txtpackdet.SelStart = 0
    txtpackdet.SelLength = Len(txtpackdet.Text)
End Sub

Private Sub txtpackdet_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTTAX.SetFocus
        Case vbKeyEscape
            Txtpackdes.SetFocus
    End Select
End Sub

Private Sub txtpackdet_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtOffer_GotFocus()
    TxtOffer.SelStart = 0
    TxtOffer.SelLength = Len(TxtOffer.Text)
End Sub

Private Sub TxtOffer_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtOffer_LostFocus()
    TxtOffer.Text = Format(Val(TxtOffer.Text), "0.00")
End Sub

Private Sub TxtOffer2_GotFocus()
    TxtOffer2.SelStart = 0
    TxtOffer2.SelLength = Len(TxtOffer2.Text)
End Sub

Private Sub TxtOffer2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtOffer2_LostFocus()
    TxtOffer2.Text = Format(Val(TxtOffer2.Text), "0.00")
End Sub

