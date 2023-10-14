VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmplumaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Creation with PLU Code"
   ClientHeight    =   5280
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   14535
   Icon            =   "frmplumaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
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
      MaxLength       =   5
      TabIndex        =   0
      Top             =   360
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
      TabIndex        =   17
      Top             =   360
      Width           =   4950
   End
   Begin VB.Frame FRAME 
      Height          =   4455
      Left            =   15
      TabIndex        =   1
      Top             =   750
      Width           =   8025
      Begin VB.TextBox TXTPLU 
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
         MaxLength       =   3
         TabIndex        =   51
         Top             =   870
         Width           =   555
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
         Left            =   2820
         MaxLength       =   5
         TabIndex        =   49
         Top             =   2835
         Visible         =   0   'False
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
         Left            =   1635
         MaxLength       =   5
         TabIndex        =   47
         Top             =   2835
         Visible         =   0   'False
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
         TabIndex        =   45
         Top             =   2250
         Width           =   2985
      End
      Begin VB.TextBox TxtMalay 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Kerala"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   4935
         MaxLength       =   30
         TabIndex        =   42
         Top             =   2700
         Width           =   2985
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
         Left            =   3630
         TabIndex        =   40
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TxtLocation 
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
         Left            =   6465
         MaxLength       =   10
         TabIndex        =   38
         Top             =   1800
         Width           =   1455
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
         Left            =   4575
         MaxLength       =   10
         TabIndex        =   36
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox chkunbill 
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
         ForeColor       =   &H00000040&
         Height          =   270
         Left            =   75
         TabIndex        =   35
         Top             =   3915
         Width           =   1575
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
         TabIndex        =   3
         Top             =   870
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
         TabIndex        =   11
         Top             =   1800
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1155
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
         TabIndex        =   10
         Top             =   1800
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
         Left            =   6990
         TabIndex        =   9
         Top             =   1320
         Width           =   930
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
         Left            =   5250
         TabIndex        =   8
         Top             =   1320
         Width           =   915
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
         TabIndex        =   6
         Top             =   1320
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
         ItemData        =   "frmplumaster.frx":0442
         Left            =   4170
         List            =   "frmplumaster.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   900
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
         ItemData        =   "frmplumaster.frx":0459
         Left            =   2505
         List            =   "frmplumaster.frx":0463
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
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
         Height          =   480
         Left            =   6645
         MaskColor       =   &H80000007&
         TabIndex        =   14
         Top             =   3240
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
         Height          =   480
         Left            =   5310
         MaskColor       =   &H80000007&
         TabIndex        =   13
         Top             =   3240
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
         Height          =   480
         Left            =   3975
         MaskColor       =   &H80000007&
         TabIndex        =   12
         Top             =   3240
         UseMaskColor    =   -1  'True
         Width           =   1275
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
         Height          =   495
         Left            =   6630
         MaskColor       =   &H80000007&
         TabIndex        =   15
         Top             =   3750
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
         Top             =   345
         Width           =   5910
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PLU CODE"
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
         TabIndex        =   52
         Top             =   945
         Width           =   915
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
         Left            =   2220
         TabIndex        =   50
         Top             =   2835
         Visible         =   0   'False
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
         Left            =   600
         TabIndex        =   48
         Top             =   2835
         Visible         =   0   'False
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
         TabIndex        =   46
         Top             =   2325
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
         Left            =   3870
         TabIndex        =   43
         Top             =   2775
         Width           =   1035
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
         Left            =   3180
         TabIndex        =   41
         Top             =   1395
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loc"
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
         Left            =   6075
         TabIndex        =   39
         Top             =   1845
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   14
         Left            =   3630
         TabIndex        =   37
         Top             =   1845
         Width           =   1035
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
         TabIndex        =   34
         Top             =   1845
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   1395
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
         TabIndex        =   31
         Top             =   1845
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
         Left            =   6180
         TabIndex        =   30
         Top             =   1380
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
         Left            =   4485
         TabIndex        =   29
         Top             =   1380
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
         TabIndex        =   28
         Top             =   1380
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
         TabIndex        =   26
         Top             =   930
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
         TabIndex        =   25
         Top             =   930
         Width           =   555
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
         TabIndex        =   16
         Top             =   420
         Width           =   1995
      End
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   1320
      TabIndex        =   18
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
      Height          =   5160
      Left            =   8055
      TabIndex        =   44
      Top             =   30
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   9102
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
      TabIndex        =   27
      Top             =   120
      Width           =   5205
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   6600
      TabIndex        =   24
      Top             =   375
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   7680
      TabIndex        =   23
      Top             =   210
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   7290
      TabIndex        =   22
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   6420
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   0
      Width           =   3960
   End
End
Attribute VB_Name = "frmplumaster"
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

Private Sub cmbfullpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            On Error Resume Next
            If cmbfullpack.ListIndex = -1 Then cmbfullpack.text = CmbPack.text
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
            TXTITEM.SetFocus
    End Select
End Sub

Private Sub CmbPack_LostFocus()
    lblPack.Caption = CmbPack.text
End Sub

Private Sub cmdcancel_Click()
    
    TXTPRODUCT.text = ""
    TXTITEM.text = ""
    TXTPLU.text = ""
    LBLITEMNAME.Caption = ""
    TxtTax.text = ""
    txtHSN.text = ""
    TxtBarcode.text = ""
    TxtLocation.text = ""
    TxtMalay.text = ""
    TxtCost.text = ""
    txtRT.text = ""
    txtWS.text = ""
    TxtMRP.text = ""
    TxtLPack.text = ""
    txtLPrice.text = ""
    Txtpackdes.text = ""
    txtpackdet.text = ""
    CmbPack.ListIndex = -1
    cmbfullpack.ListIndex = -1
    Set DataList2.RowSource = Nothing
    TXTITEMCODE.Enabled = True
    DataList2.Enabled = True
    FRAME.Visible = False
    TXTPRODUCT.Visible = False
    DataList2.Visible = False
    TXTITEMCODE.SetFocus
    chkunbill.Value = 0
    
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST where LENGTH(PLU_CODE)>0", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        If IsNull(TRXMAST.Fields(0)) Then
            TXTITEMCODE.text = 1
        Else
            TXTITEMCODE.text = Val(TRXMAST.Fields(0)) + 1
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    TXTITEMCODE.text = Format(TXTITEMCODE.text, "00000")
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If MDIMAIN.StatusBar.Panels(9).text = "Y" Then Exit Sub
    Dim RSTITEMMAST As ADODB.Recordset
    On Error Resume Next
    If Val(Txtpackdes.text) = 0 Then Txtpackdes.text = 1
    If Val(txtpackdet.text) = 0 Then txtpackdet.text = 1
    If TXTITEM.text = "" Then
        MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
        TXTITEM.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(TXTPLU.text)) = 0 Then
        If MsgBox("PLU Code not entered. Are you sure?", vbYesNo + vbDefaultButton2, "PRODUCT MASTER") = vbNo Then
            TXTITEM.SetFocus
            Exit Sub
        End If
    End If
    
    On Error GoTo ERRHAND
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE PLU_CODE = '" & Trim(TXTPLU.text) & "' AND ITEM_CODE <> '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        MsgBox "PLU CODE ALREADY ASSIGNED TO " & IIf(IsNull(RSTITEMMAST!ITEM_NAME), "", RSTITEMMAST!ITEM_NAME) & IIf(IsNull(RSTITEMMAST!ITEM_CODE), "", " (" & RSTITEMMAST!ITEM_CODE & ")"), vbOKOnly, "Item Master"
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        Exit Sub
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME = '" & Trim(TXTITEM.text) & "' AND ITEM_CODE <> '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        MsgBox "The Item name already exists...", vbOKOnly, "Item Master"
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        Exit Sub
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.Properties("Update Criteria").Value = adCriteriaKey
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.text)
        RSTITEMMAST!UNIT = 1
        RSTITEMMAST!REORDER_QTY = 0
        RSTITEMMAST!PACK_TYPE = CmbPack.text
        RSTITEMMAST!FULL_PACK = cmbfullpack.text
        RSTITEMMAST!BIN_LOCATION = Trim(TxtLocation.text)
        RSTITEMMAST!ITEM_MAL = Trim(TxtMalay.text)
        RSTITEMMAST!SALES_TAX = Val(TxtTax.text)
        RSTITEMMAST!REMARKS = Trim(txtHSN.text)
        If Val(TxtTax.text) > 0 Then RSTITEMMAST!check_flag = "V"
        RSTITEMMAST!ITEM_COST = Val(TxtCost.text)
        RSTITEMMAST!P_RETAIL = Val(txtRT.text)
        RSTITEMMAST!MRP = Val(TxtMRP.text)
        RSTITEMMAST!P_WS = Val(txtWS.text)
        RSTITEMMAST!CRTN_PACK = IIf(Val(TxtLPack.text) = 0, 1, Val(TxtLPack.text))
        RSTITEMMAST!P_CRTN = Val(txtLPrice.text)
        RSTITEMMAST!LOOSE_PACK = IIf(Val(Txtpack.text) = 0, 1, Val(Txtpack.text))
        RSTITEMMAST!PACK_DESC = Val(Txtpackdes.text)
        RSTITEMMAST!PACK_DET = Val(txtpackdet.text)
        RSTITEMMAST!BARCODE = Trim(TxtBarcode.text)
        'RSTITEMMAST!DEAD_STOCK = "N"
        If chkunbill.Value = 0 Then
            RSTITEMMAST!UN_BILL = "N"
        Else
            RSTITEMMAST!UN_BILL = "Y"
        End If
        RSTITEMMAST!PRICE_CHANGE = "Y"
        RSTITEMMAST!PLU_CODE = Trim(TXTPLU.text)
        RSTITEMMAST.Update
    Else
        RSTITEMMAST.AddNew
        RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
        RSTITEMMAST!ITEM_CODE = Trim(TXTITEMCODE.text)
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.text)
        RSTITEMMAST!Category = "GENERAL"
        RSTITEMMAST!UNIT = 1
        RSTITEMMAST!MANUFACTURER = "GENERAL"
        RSTITEMMAST!DEAD_STOCK = "N"
        If chkunbill.Value = 0 Then
            RSTITEMMAST!UN_BILL = "N"
        Else
            RSTITEMMAST!UN_BILL = "Y"
        End If
        RSTITEMMAST!PRICE_CHANGE = "Y"
        
        RSTITEMMAST!REMARKS = Trim(txtHSN.text)
        RSTITEMMAST!ITEM_SPEC = ""
        RSTITEMMAST!REORDER_QTY = 0
        If CmbPack.ListIndex = -1 Then
            RSTITEMMAST!PACK_TYPE = "Kg"
        Else
            RSTITEMMAST!PACK_TYPE = CmbPack.text
        End If
        If cmbfullpack.ListIndex = -1 Then
            RSTITEMMAST!FULL_PACK = "Kg"
        Else
            RSTITEMMAST!FULL_PACK = CmbPack.text
        End If
        RSTITEMMAST!BIN_LOCATION = Trim(TxtLocation.text)
        RSTITEMMAST!ITEM_MAL = Trim(TxtMalay.text)
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
        RSTITEMMAST!SALES_TAX = Val(TxtTax.text)
        If Val(TxtTax.text) > 0 Then RSTITEMMAST!check_flag = "V"
        RSTITEMMAST!ITEM_COST = Val(TxtCost.text)
        RSTITEMMAST!P_RETAIL = Val(txtRT.text)
        RSTITEMMAST!MRP = Val(TxtMRP.text)
        RSTITEMMAST!P_WS = Val(txtWS.text)
        RSTITEMMAST!CRTN_PACK = IIf(Val(TxtLPack.text) = 0, 1, Val(TxtLPack.text))
        RSTITEMMAST!P_CRTN = Val(txtLPrice.text)
        RSTITEMMAST!LOOSE_PACK = IIf(Val(Txtpack.text) = 0, 1, Val(Txtpack.text))
        RSTITEMMAST!PACK_DESC = Val(Txtpackdes.text)
        RSTITEMMAST!PACK_DET = Val(txtpackdet.text)
        RSTITEMMAST!BARCODE = Trim(TxtBarcode.text)
        RSTITEMMAST!PLU_CODE = Trim(TXTPLU.text)
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & Trim(TxtItemcode.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until RSTITEMMAST.EOF
'        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
'        RSTITEMMAST!MFGR = Trim(txtcompany.Text)
'        RSTITEMMAST!Category = Trim(txtcategory.Text)
'        RSTITEMMAST.Update
'        RSTITEMMAST.MoveNext
'    Loop
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    
                    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    cmdcancel_Click
Exit Sub
ERRHAND:
    MsgBox (err.Description)
        
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TxtMalay.SetFocus
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
                TXTITEMCODE.text = RSTITEMMAST!ITEM_CODE
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
    TRXMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST where LENGTH(PLU_CODE)>0", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        If IsNull(TRXMAST.Fields(0)) Then
            TXTITEMCODE.text = 1
        Else
            TXTITEMCODE.text = Val(TRXMAST.Fields(0)) + 1
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    TXTITEMCODE.text = Format(TXTITEMCODE.text, "00000")
    
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
        TXTITEM.text = Trim(grdtmp.Columns(1))
        TXTITEM.SetFocus
    End If
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            On Error Resume Next
            TXTITEM.text = Trim(grdtmp.Columns(1))
            TXTITEM.SetFocus
    End Select
End Sub

Private Sub TxtCost_GotFocus()
    TxtCost.SelStart = 0
    TxtCost.SelLength = Len(TxtCost.text)
End Sub

Private Sub TxtCost_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtMRP.SetFocus
        Case vbKeyEscape
            TxtTax.SetFocus
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
    TxtCost.text = Format(Val(TxtCost.text), "0.00")
End Sub

Private Sub TXTITEM_Change()
    On Error GoTo ERRHAND
    Set grdtmp.DataSource = Nothing
    If PHYFLAG = True Then
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.text) & "%' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    Else
        PHY.Close
        'PHY.Open "Select ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' AND ITEM_CODE <> '" & Trim(TXTITEMCODE.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.text) & "%' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
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
    TXTITEM.SelLength = Len(TXTITEM.text)
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTITEM.text = "" Then
                MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
                TXTITEM.SetFocus
                Exit Sub
            End If
            If Trim(TXTPLU.text) = "" Then
                TXTPLU.SetFocus
            Else
                Txtpack.SetFocus
            End If
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
        PHY.Open "Select ITEM_CODE,ITEM_NAME, PLU_CODE From ITEMMAST  WHERE ITEM_CODE Like '" & Trim(Me.TXTITEMCODE.text) & "%' AND LENGTH(PLU_CODE)>0 ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    Else
        PHY.Close
        PHY.Open "Select ITEM_CODE,ITEM_NAME, PLU_CODE From ITEMMAST  WHERE ITEM_CODE Like '" & Trim(Me.TXTITEMCODE.text) & "%' AND LENGTH(PLU_CODE)>0 ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    End If
    Set grdtmp.DataSource = PHY
    grdtmp.Columns(0).Caption = "Item Code"
    grdtmp.Columns(0).Caption = "Item Description"
    grdtmp.Columns(0).Caption = "PLU Code"
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
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.text)
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ERRHAND
            If Trim(TXTITEMCODE.text) = "" Then Exit Sub
            If Len(Trim(TXTITEMCODE.text)) <> 5 Then
                MsgBox "Item Code should be 5 characters", , "PLU Master"
                TXTITEMCODE.SetFocus
                Exit Sub
            End If
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                TXTPLU.text = IIf(IsNull(RSTITEMMAST!PLU_CODE), "", RSTITEMMAST!PLU_CODE)
                TXTITEM.text = IIf(IsNull(RSTITEMMAST!ITEM_NAME), "", RSTITEMMAST!ITEM_NAME)
                TxtTax.text = IIf(IsNull(RSTITEMMAST!SALES_TAX), 0, RSTITEMMAST!SALES_TAX)
                txtHSN.text = IIf(IsNull(RSTITEMMAST!REMARKS), "", RSTITEMMAST!REMARKS)
                TxtBarcode.text = IIf(IsNull(RSTITEMMAST!BARCODE), "", RSTITEMMAST!BARCODE)
                TxtLocation.text = IIf(IsNull(RSTITEMMAST!BIN_LOCATION), "", RSTITEMMAST!BIN_LOCATION)
                TxtMalay.text = IIf(IsNull(RSTITEMMAST!ITEM_MAL), "", RSTITEMMAST!ITEM_MAL)
                TxtCost.text = IIf(IsNull(RSTITEMMAST!ITEM_COST), 0, RSTITEMMAST!ITEM_COST)
                txtRT.text = IIf(IsNull(RSTITEMMAST!P_RETAIL), 0, RSTITEMMAST!P_RETAIL)
                txtWS.text = IIf(IsNull(RSTITEMMAST!P_WS), 0, RSTITEMMAST!P_WS)
                TxtMRP.text = IIf(IsNull(RSTITEMMAST!MRP), 0, RSTITEMMAST!MRP)
                TxtLPack.text = IIf(IsNull(RSTITEMMAST!CRTN_PACK), 1, RSTITEMMAST!CRTN_PACK)
                txtLPrice.text = IIf(IsNull(RSTITEMMAST!P_CRTN), 0, RSTITEMMAST!P_CRTN)
                Txtpackdes.text = IIf(IsNull(RSTITEMMAST!PACK_DESC), 1, RSTITEMMAST!PACK_DESC)
                txtpackdet.text = IIf(IsNull(RSTITEMMAST!PACK_DET), 1, RSTITEMMAST!PACK_DET)
                Txtpack.text = IIf(IsNull(RSTITEMMAST!LOOSE_PACK), 1, RSTITEMMAST!LOOSE_PACK)
                If IsNull(RSTITEMMAST!UN_BILL) Or RSTITEMMAST!UN_BILL = "N" Or RSTITEMMAST!UN_BILL = "" Then
                    chkunbill.Value = 0
                Else
                    chkunbill.Value = 1
                End If
                On Error Resume Next
                CmbPack.text = IIf(IsNull(RSTITEMMAST!PACK_TYPE), 0, RSTITEMMAST!PACK_TYPE)
                cmbfullpack.text = IIf(IsNull(RSTITEMMAST!FULL_PACK), 0, RSTITEMMAST!FULL_PACK)
                On Error GoTo ERRHAND
                If Val(Txtpack.text) = 0 Then Txtpack.text = 1
                If Val(TxtLPack.text) = 0 Then TxtLPack.text = 1
                If Val(txtLPrice.text) = 0 Then txtLPrice.text = Val(txtRT.text) / Val(TxtLPack.text)
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
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLPack_Change()
    txtLPrice.text = ""
End Sub

Private Sub TxtLPack_GotFocus()
    If Val(TxtLPack.text) = 0 Then
        TxtLPack.text = 1
    End If
    TxtLPack.SelStart = 0
    TxtLPack.SelLength = Len(TxtLPack.text)
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
    If Val(TxtLPack.text) = 0 Then TxtLPack.text = "1"
    If Val(txtLPrice.text) = 0 Then
        txtLPrice.text = Format(Round(Val(txtRT.text) / Val(TxtLPack.text), 2), "0.00")
    End If
    txtLPrice.SelStart = 0
    txtLPrice.SelLength = Len(txtLPrice.text)
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
    txtLPrice.text = Format(Val(txtLPrice.text), "0.00")
End Sub

Private Sub Txtpack_GotFocus()
    If Val(Txtpack.text) = 0 Then
        Txtpack.text = 1
    End If
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.text)
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmbPack.SetFocus
        Case vbKeyEscape
            TXTPLU.SetFocus
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

Private Sub TXTPLU_GotFocus()
    TXTPLU.SelStart = 0
    TXTPLU.SelLength = Len(TXTPLU.text)
End Sub

Private Sub TXTPLU_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmbPack.SetFocus
        Case vbKeyEscape
            TXTITEM.SetFocus
    End Select
End Sub

Private Sub TXTPLU_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPRODUCT_Change()
    On Error GoTo ERRHAND
    Set grdtmp.DataSource = Nothing
    If REPFLAG = True Then
        RSTREP.Open "Select ITEM_CODE,ITEM_NAME, PLU_CODE From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.text & "%' AND LENGTH(PLU_CODE)>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select ITEM_CODE,ITEM_NAME, PLU_CODE From ITEMMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.text & "%' AND LENGTH(PLU_CODE)>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    End If
    Set DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Set grdtmp.DataSource = RSTREP
    grdtmp.Columns(0).Caption = "Code"
    grdtmp.Columns(1).Caption = "Item Description"
    grdtmp.Columns(2).Caption = "PLU"
    
    grdtmp.Columns(0).Width = 1000
    grdtmp.Columns(1).Width = 3800
    grdtmp.Columns(2).Width = 1000
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTPRODUCT.text = "" Then Exit Sub
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
    If MDIMAIN.StatusBar.Panels(9).text = "Y" Then Exit Sub
    Dim rststock As ADODB.Recordset
    
    If TXTITEMCODE.text = "" Then Exit Sub
    On Error GoTo ERRHAND
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from RTRXFILE where RTRXFILE.ITEM_CODE = '" & TXTITEMCODE.text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.text & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFILE where ITEM_CODE = '" & TXTITEMCODE.text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.text & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULASUB where FOR_NAME = '" & TXTITEMCODE.text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.text & " Since Transactions is Available in Formula", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULAMAST where ITEM_CODE = '" & TXTITEMCODE.text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.text & " Since Transactions is Available in Formula", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & TXTITEM.text & " ****", vbYesNo, "DELETING ITEM....") = vbNo Then Exit Sub
    'db.Execute ("DELETE from RTRXFILE where RTRXFILE.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
    db.Execute ("DELETE from PRODLINK where PRODLINK.ITEM_CODE = '" & TXTITEMCODE.text & "'")
    db.Execute ("DELETE from ITEMMAST where ITEMMAST.ITEM_CODE = '" & TXTITEMCODE.text & "'")
    
    'tXTMEDICINE.Tag = tXTMEDICINE.Text
    'tXTMEDICINE.Text = ""
    'tXTMEDICINE.Text = tXTMEDICINE.Tag
    'TXTQTY.Text = ""
    MsgBox "ITEM " & TXTITEM.text & "DELETED SUCCESSFULLY", vbInformation, "DELETING ITEM...."
    Call cmdcancel_Click
    Exit Sub
   
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub txtRT_GotFocus()
    txtRT.SelStart = 0
    txtRT.SelLength = Len(txtRT.text)
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
    txtRT.text = Format(Val(txtRT.text), "0.00")
End Sub

Private Sub TxtTax_GotFocus()
    TxtTax.SelStart = 0
    TxtTax.SelLength = Len(TxtTax.text)
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
    TxtTax.text = Format(Val(TxtTax.text), "0.00")
End Sub

Private Sub txtws_GotFocus()
    txtWS.SelStart = 0
    txtWS.SelLength = Len(txtWS.text)
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
    txtWS.text = Format(Val(txtWS.text), "0.00")
End Sub

Private Sub TxtHSN_GotFocus()
    txtHSN.SelStart = 0
    txtHSN.SelLength = Len(txtHSN.text)
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
    TxtLocation.SelLength = Len(TxtLocation.text)
End Sub

Private Sub TxtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmdSave_Click
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
    TxtMRP.SelLength = Len(TxtMRP.text)
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
    TxtMRP.text = Format(Val(TxtMRP.text), "0.00")
End Sub

Private Sub TxtMalay_GotFocus()
    TxtMalay.SelStart = 0
    TxtMalay.SelLength = Len(TxtMalay.text)
End Sub

Private Sub TxtMalay_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmdSave_Click
        Case vbKeyEscape
            TxtLocation.SetFocus
    End Select
End Sub

Private Sub TxtMalay_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtbarcode_GotFocus()
    TxtBarcode.SelStart = 0
    TxtBarcode.SelLength = Len(TxtBarcode.text)
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmdSave_Click
        Case vbKeyEscape
            txtLPrice.SetFocus
    End Select
End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtpackdes_GotFocus()
    If Val(Txtpackdes.text) = 0 Then
        Txtpackdes.text = 1
    End If
    Txtpackdes.SelStart = 0
    Txtpackdes.SelLength = Len(Txtpackdes.text)
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
    If Val(txtpackdet.text) = 0 Then
        txtpackdet.text = 1
    End If
    txtpackdet.SelStart = 0
    txtpackdet.SelLength = Len(txtpackdet.text)
End Sub

Private Sub txtpackdet_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtTax.SetFocus
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

