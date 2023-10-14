VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmStkmovmnt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK MOVEMENT"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19755
   ClipControls    =   0   'False
   Icon            =   "frmstkmvmt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   19755
   Begin VB.CommandButton Command3 
      Caption         =   "Display Summary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   17520
      TabIndex        =   49
      Top             =   7605
      Width           =   1380
   End
   Begin VB.TextBox TxtBarcode 
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
      Left            =   6660
      TabIndex        =   47
      Top             =   210
      Width           =   2445
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Report for All Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16080
      TabIndex        =   46
      Top             =   7620
      Width           =   1380
   End
   Begin VB.CommandButton cmdstkcrct 
      Caption         =   "Stock Crrection"
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
      Left            =   9750
      TabIndex        =   45
      Top             =   270
      Width           =   1380
   End
   Begin VB.TextBox LBLITEMCODE 
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
      Left            =   5055
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   585
      Width           =   4050
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
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   210
      Width           =   1605
   End
   Begin VB.TextBox tXTCODE 
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
      Left            =   5055
      TabIndex        =   2
      Top             =   210
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
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
      Left            =   14805
      TabIndex        =   33
      Top             =   7620
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SORT ORDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   11940
      TabIndex        =   25
      Top             =   15
      Width           =   7245
      Begin VB.OptionButton OPTOUTDATE 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   105
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.OptionButton OPTOUTCUST 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1770
         TabIndex        =   10
         Top             =   240
         Width           =   3285
      End
   End
   Begin VB.TextBox tXTMEDICINE 
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
      Left            =   1620
      TabIndex        =   1
      Top             =   210
      Width           =   3420
   End
   Begin VB.CommandButton CmdExit 
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
      Height          =   495
      Left            =   13470
      TabIndex        =   13
      Top             =   7620
      Width           =   1260
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   5685
      Left            =   0
      TabIndex        =   11
      Top             =   1860
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   10028
      _Version        =   393216
      Rows            =   1
      Cols            =   21
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   8438015
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   3
      Appearance      =   0
      GridLineWidth   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1035
      Left            =   0
      TabIndex        =   3
      Top             =   570
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   1826
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SORT ORDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   870
      Left            =   9120
      TabIndex        =   16
      Top             =   720
      Width           =   2010
      Begin VB.OptionButton OPTBALQTY 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Available Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   30
         TabIndex        =   8
         Top             =   615
         Width           =   1950
      End
      Begin VB.OptionButton optsupplier 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Supplier"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   30
         TabIndex        =   7
         Top             =   390
         Width           =   1920
      End
      Begin VB.OptionButton optdate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   30
         TabIndex        =   6
         Top             =   165
         Value           =   -1  'True
         Width           =   1620
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRDOUTWARD 
      Height          =   6675
      Left            =   11940
      TabIndex        =   12
      Top             =   870
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   11774
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   8438015
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   3
      Appearance      =   0
      GridLineWidth   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTFROM 
      Height          =   315
      Left            =   5670
      TabIndex        =   4
      Top             =   945
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   393216
      CalendarForeColor=   0
      CalendarTitleForeColor=   16576
      CalendarTrailingForeColor=   255
      Format          =   112852993
      CurrentDate     =   41640
      MinDate         =   40179
   End
   Begin MSComCtl2.DTPicker DTTO 
      Height          =   315
      Left            =   7275
      TabIndex        =   5
      Top             =   930
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   112852993
      CurrentDate     =   41640
      MinDate         =   40179
   End
   Begin VB.Frame FrmeAll 
      Height          =   435
      Left            =   5055
      TabIndex        =   50
      Top             =   1185
      Visible         =   0   'False
      Width           =   1920
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   30
         TabIndex        =   51
         Top             =   150
         Value           =   -1  'True
         Width           =   1845
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode"
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
      Height          =   210
      Index           =   11
      Left            =   6660
      TabIndex        =   48
      Top             =   15
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Height          =   210
      Index           =   13
      Left            =   5070
      TabIndex        =   43
      Top             =   15
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
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
      Height          =   225
      Index           =   12
      Left            =   15
      TabIndex        =   42
      Top             =   15
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AVAIL. VALUE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   10
      Left            =   6975
      TabIndex        =   41
      Top             =   7935
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OUT. VALUE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Index           =   9
      Left            =   3420
      TabIndex        =   40
      Top             =   7935
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IN. VALUE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Index           =   8
      Left            =   45
      TabIndex        =   39
      Top             =   7950
      Width           =   1545
   End
   Begin VB.Label lblavailval 
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
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   8805
      TabIndex        =   38
      Top             =   7950
      Width           =   2235
   End
   Begin VB.Label lbloutval 
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
      Height          =   315
      Left            =   5100
      TabIndex        =   37
      Top             =   7950
      Width           =   1740
   End
   Begin VB.Label lblinval 
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
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   1500
      TabIndex        =   36
      Top             =   7935
      Width           =   1860
   End
   Begin VB.Label LBLWASTE 
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
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   5100
      TabIndex        =   35
      Top             =   8280
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WASTAGE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Index           =   5
      Left            =   3420
      TabIndex        =   34
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OP. QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   4
      Left            =   60
      TabIndex        =   32
      Top             =   8280
      Width           =   1365
   End
   Begin VB.Label LBLOPQTY 
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
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   1500
      TabIndex        =   31
      Top             =   8280
      Width           =   1860
   End
   Begin VB.Label LBLTOTAL 
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
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
      Height          =   195
      Index           =   5
      Left            =   7005
      TabIndex        =   30
      Top             =   945
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period"
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
      Height          =   270
      Index           =   7
      Left            =   5055
      TabIndex        =   29
      Top             =   975
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   3
      Left            =   6480
      TabIndex        =   28
      Top             =   8325
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOOSE QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   2
      Left            =   3360
      TabIndex        =   27
      Top             =   8970
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label LblLoose 
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
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   5235
      TabIndex        =   26
      Top             =   8910
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label lblmanual 
      Caption         =   "*Adjusted Manually"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11220
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Index           =   1
      Left            =   11925
      TabIndex        =   23
      Top             =   600
      Width           =   7230
   End
   Begin VB.Label LBLOUTWARD 
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
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   5100
      TabIndex        =   22
      Top             =   7575
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AVAILABLE QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Index           =   1
      Left            =   6960
      TabIndex        =   21
      Top             =   7620
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OUTWARD QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   0
      Left            =   3405
      TabIndex        =   20
      Top             =   7620
      Width           =   1785
   End
   Begin VB.Label LBLBALANCE 
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
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   8805
      TabIndex        =   19
      Top             =   7575
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INWARD QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Index           =   6
      Left            =   60
      TabIndex        =   18
      Top             =   7620
      Width           =   1545
   End
   Begin VB.Label LBLINWARD 
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
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   1500
      TabIndex        =   17
      Top             =   7575
      Width           =   1860
   End
   Begin VB.Label LBLHEAD 
      BackColor       =   &H00000000&
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
      Left            =   3090
      TabIndex        =   15
      Top             =   1620
      Width           =   8835
   End
   Begin VB.Label LBLHEAD 
      BackColor       =   &H00000000&
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
      Left            =   0
      TabIndex        =   14
      Top             =   1620
      Width           =   3750
   End
End
Attribute VB_Name = "FrmStkmovmnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPFLAG As Boolean 'REP
Dim RSTREP As New ADODB.Recordset

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdStkCrct_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    Dim RSTBALQTY As ADODB.Recordset
    Dim INWARD As Double
    Dim OUTWARD As Double
    Dim BALQTY As Double
    Dim DIFFQTY As Double
    Dim i As Long
''''    db.Execute "delete from cashatrxfile"
''''    db.Execute "delete from dbtpymt"
''''    db.Execute "delete from BANK_TRX"
''''    db.Execute "delete from CATEGORY"
''''    Exit Sub
    
    If DataList2.BoundText = "" Then Exit Sub
    If MsgBox("DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    On Error GoTo ERRHAND
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTITEMMAST.Properties("Update Criteria").Value = adCriteriaKey
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        
        BALQTY = 0
        Set RSTBALQTY = New ADODB.Recordset
        RSTBALQTY.Open "Select SUM(BAL_QTY) FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "'", db, adOpenForwardOnly
        If Not (RSTBALQTY.EOF And RSTBALQTY.BOF) Then
            BALQTY = IIf(IsNull(RSTBALQTY.Fields(0)), 0, RSTBALQTY.Fields(0))
        End If
        RSTBALQTY.Close
        Set RSTBALQTY = Nothing
        
        INWARD = 0
        OUTWARD = 0
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            INWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
            
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(FREE_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            INWARD = INWARD + IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
        
                
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockOptimistic, adCmdText
'        Do Until rststock.EOF
'            OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
'            OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
'            rststock.MoveNext
'        Loop
'        rststock.Close
'        Set rststock = Nothing
        
        Set rststock = New ADODB.Recordset
        rststock.Open "Select SUM((QTY + FREE_QTY) * LOOSE_PACK) FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rststock.EOF And rststock.BOF) Then
            OUTWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
        
'        Set rststock = New ADODB.Recordset
'        rststock.Open "Select SUM(FREE_QTY * LOOSE_PACK) FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI' OR TRX_TYPE='GI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (rststock.EOF And rststock.BOF) Then
'            OUTWARD = OUTWARD + IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
'        End If
'        rststock.Close
'        Set rststock = Nothing
        
        If Round(INWARD - OUTWARD, 2) = Round(BALQTY, 2) Then GoTo SKIP_BALCHECK
        
        
        db.Execute "Update RTRXFILE set BAL_QTY = QTY where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' "
        BALQTY = 0
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until rststock.EOF
            BALQTY = 0
            Set RSTBALQTY = New ADODB.Recordset
            RSTBALQTY.Open "Select SUM(QTY) FROM TRXSUB WHERE R_TRX_YEAR ='" & rststock!TRX_YEAR & "' AND R_TRX_TYPE='" & rststock!TRX_TYPE & "' AND R_VCH_NO = " & rststock!VCH_NO & " AND R_LINE_NO = " & rststock!LINE_NO & "", db, adOpenForwardOnly
            If Not (RSTBALQTY.EOF And RSTBALQTY.BOF) Then
                BALQTY = IIf(IsNull(RSTBALQTY.Fields(0)), 0, RSTBALQTY.Fields(0))
            End If
            RSTBALQTY.Close
            Set RSTBALQTY = Nothing
            
            rststock!BAL_QTY = rststock!BAL_QTY - BALQTY
            rststock.Update
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        
        
        
        db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY <0"
        BALQTY = 0
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(BAL_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            BALQTY = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
        
        If Round(INWARD - OUTWARD, 2) < BALQTY Then
            DIFFQTY = BALQTY - (Round(INWARD - OUTWARD, 2))
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until rststock.EOF
                If DIFFQTY - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) >= 0 Then
                    DIFFQTY = DIFFQTY - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY)
                    rststock!BAL_QTY = 0
                    rststock.Update
                Else
                    rststock!BAL_QTY = Round(IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) - DIFFQTY, 2)
                    DIFFQTY = 0
                    rststock.Update
                End If
                If DIFFQTY <= 0 Then Exit Do
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
        ElseIf Round(INWARD - OUTWARD, 2) > BALQTY Then
            DIFFQTY = Round((INWARD - OUTWARD), 2) - BALQTY
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until rststock.EOF
                If DIFFQTY <= IIf(IsNull(rststock!QTY), 0, rststock!QTY) - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) Then
                    rststock!BAL_QTY = Round(IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) + DIFFQTY, 2)
                    DIFFQTY = 0
                Else
                    If Not rststock!BAL_QTY = IIf(IsNull(rststock!QTY), 0, rststock!QTY) Then
                        rststock!BAL_QTY = Round(IIf(IsNull(rststock!QTY), 0, rststock!QTY), 2)
                        DIFFQTY = DIFFQTY - IIf(IsNull(rststock!QTY), 0, rststock!QTY)
                    End If
                End If
                rststock.Update
                If DIFFQTY <= 0 Then Exit Do
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            'MsgBox ""
        End If
        
SKIP_BALCHECK:
        RSTITEMMAST!CLOSE_QTY = Round(INWARD - OUTWARD, 2)
        RSTITEMMAST!RCPT_QTY = INWARD
        RSTITEMMAST!ISSUE_QTY = OUTWARD
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Screen.MousePointer = vbNormal

    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTREPORT"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.OpenSubreport("Rptinward").RecordSelectionFormula = "( ({TRXFILE.TRX_TYPE} = 'OG' OR {TRXFILE.TRX_TYPE} = 'PI' OR {TRXFILE.TRX_TYPE} = 'OP') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    Report.OpenSubreport("RPTINWRD.rpt").RecordSelectionFormula = "({RTRXFILE.ITEM_CODE} = '" & DataList2.BoundText & "' AND {RTRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    'Report.OpenSubreport("RPTINWRD.rpt").RecordSelectionFormula = "({TRANSMAST.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRANSMAST.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTINWRD.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTINWRD.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTINWRD.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RPTINWRD.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    
    
    Report.OpenSubreport("RPTOUTWARD.rpt").RecordSelectionFormula = "({TRXFILE.ITEM_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTOUTWARD.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTOUTWARD.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTOUTWARD.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RPTOUTWARD.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    
    Report.OpenSubreport("RPTINWRD.rpt").DiscardSavedData
    Report.OpenSubreport("RPTINWRD.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTINWRD.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    
    Report.OpenSubreport("RPTOUTWARD.rpt").DiscardSavedData
    Report.OpenSubreport("RPTOUTWARD.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTOUTWARD.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "ITEM WISE INWARD OUTWARD MOVEMENT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTREPORT2"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.OpenSubreport("Rptinward").RecordSelectionFormula = "( ({TRXFILE.TRX_TYPE} = 'OG' OR {TRXFILE.TRX_TYPE} = 'PI' OR {TRXFILE.TRX_TYPE} = 'OP') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    Report.OpenSubreport("RPTINWRD.rpt").RecordSelectionFormula = "({RTRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    'Report.OpenSubreport("RPTINWRD.rpt").RecordSelectionFormula = "({TRANSMAST.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRANSMAST.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTINWRD.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTINWRD.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTINWRD.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RPTINWRD.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    
    
    Report.OpenSubreport("RPTOUTWARD.rpt").RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTOUTWARD.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTOUTWARD.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTOUTWARD.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RPTOUTWARD.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    
    Report.OpenSubreport("RPTINWRD.rpt").DiscardSavedData
    Report.OpenSubreport("RPTINWRD.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTINWRD.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    
    Report.OpenSubreport("RPTOUTWARD.rpt").DiscardSavedData
    Report.OpenSubreport("RPTOUTWARD.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTOUTWARD.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "ITEM WISE INWARD OUTWARD MOVEMENT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Command3_Click()
    frmstkmvmreport.Show
    frmstkmvmreport.SetFocus
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            GRDSTOCK.SetFocus
            'DataList2.SetFocus
        Case vbKeyEscape
            TxtItemName.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    REPFLAG = True
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "TYPE"
    GRDSTOCK.TextMatrix(0, 4) = "SUPPLIER"
    GRDSTOCK.TextMatrix(0, 5) = "QTY"
    GRDSTOCK.TextMatrix(0, 6) = "INV DATE"
    GRDSTOCK.TextMatrix(0, 7) = "INV NO"
    GRDSTOCK.TextMatrix(0, 8) = "COMP REF"
    GRDSTOCK.TextMatrix(0, 9) = "" '"PACK"
    GRDSTOCK.TextMatrix(0, 10) = "Serial No"
    GRDSTOCK.TextMatrix(0, 11) = "EXPIRY"
    GRDSTOCK.TextMatrix(0, 12) = "MRP"
    GRDSTOCK.TextMatrix(0, 13) = "Cost"
    GRDSTOCK.TextMatrix(0, 14) = "Net Cost"
    GRDSTOCK.TextMatrix(0, 15) = "BAL QTY"
    GRDSTOCK.TextMatrix(0, 16) = "PACK"
    GRDSTOCK.TextMatrix(0, 17) = "LINE"
    GRDSTOCK.TextMatrix(0, 20) = "MRP"
    
    GRDSTOCK.ColWidth(0) = 400
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 0
    GRDSTOCK.ColWidth(3) = 900
    GRDSTOCK.ColWidth(4) = 2000
    GRDSTOCK.ColWidth(5) = 1100
    GRDSTOCK.ColWidth(6) = 1100
    GRDSTOCK.ColWidth(7) = 900
    GRDSTOCK.ColWidth(8) = 900
    GRDSTOCK.ColWidth(9) = 0 '700
    GRDSTOCK.ColWidth(10) = 800
    GRDSTOCK.ColWidth(11) = 0
    GRDSTOCK.ColWidth(12) = 0
    If frmLogin.rs!Level = "0" Then
        GRDSTOCK.ColWidth(13) = 900
        GRDSTOCK.ColWidth(14) = 900
    Else
        GRDSTOCK.ColWidth(13) = 0
        GRDSTOCK.ColWidth(14) = 0
    End If
    GRDSTOCK.ColWidth(15) = 900
    GRDSTOCK.ColWidth(16) = 500
    GRDSTOCK.ColWidth(17) = 400
    GRDSTOCK.ColWidth(18) = 0
    GRDSTOCK.ColWidth(19) = 0
    GRDSTOCK.ColWidth(20) = 1100
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 1
    GRDSTOCK.ColAlignment(4) = 1
    GRDSTOCK.ColAlignment(5) = 4
    GRDSTOCK.ColAlignment(6) = 4
    GRDSTOCK.ColAlignment(7) = 1
     GRDSTOCK.ColAlignment(8) = 1
    GRDSTOCK.ColAlignment(9) = 4
    GRDSTOCK.ColAlignment(10) = 1
    GRDSTOCK.ColAlignment(11) = 4
    GRDSTOCK.ColAlignment(12) = 1
    GRDSTOCK.ColAlignment(13) = 4
    GRDSTOCK.ColAlignment(14) = 4
    GRDSTOCK.ColAlignment(15) = 4
    GRDSTOCK.ColAlignment(16) = 4
    
    GRDOUTWARD.TextMatrix(0, 0) = "SL"
    GRDOUTWARD.TextMatrix(0, 1) = "TYPE"
    GRDOUTWARD.TextMatrix(0, 2) = "CUSTOMER"
    GRDOUTWARD.TextMatrix(0, 3) = "QTY"
    GRDOUTWARD.TextMatrix(0, 4) = "FREE"
    GRDOUTWARD.TextMatrix(0, 5) = "Pack"
    GRDOUTWARD.TextMatrix(0, 6) = "RATE"
    GRDOUTWARD.TextMatrix(0, 7) = "INV #"
    GRDOUTWARD.TextMatrix(0, 8) = "INV DATE"
    GRDOUTWARD.TextMatrix(0, 9) = "Serial No"
    

    GRDOUTWARD.ColWidth(0) = 400
    GRDOUTWARD.ColWidth(1) = 0
    GRDOUTWARD.ColWidth(2) = 2100
    GRDOUTWARD.ColWidth(3) = 1000
    GRDOUTWARD.ColWidth(4) = 700
    GRDOUTWARD.ColWidth(5) = 800
    GRDOUTWARD.ColWidth(6) = 900
    GRDOUTWARD.ColWidth(7) = 1000
    GRDOUTWARD.ColWidth(8) = 1200
    GRDOUTWARD.ColWidth(9) = 900
    
    GRDOUTWARD.ColWidth(11) = 0
    GRDOUTWARD.ColWidth(12) = 0
    GRDOUTWARD.ColWidth(13) = 0
    
    GRDOUTWARD.ColAlignment(0) = 4
    GRDOUTWARD.ColAlignment(1) = 1
    GRDOUTWARD.ColAlignment(2) = 1
    GRDOUTWARD.ColAlignment(3) = 1
    GRDOUTWARD.ColAlignment(4) = 1
    GRDOUTWARD.ColAlignment(5) = 4
    GRDOUTWARD.ColAlignment(6) = 1
    GRDOUTWARD.ColAlignment(7) = 4
    GRDOUTWARD.ColAlignment(8) = 4
    GRDOUTWARD.ColAlignment(9) = 4
    
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    DTFROM.Value = "01/04/" & Year(MDIMAIN.DTFROM.Value)
    'Me.Height = 9990
    'Me.Width = 18555
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If REPFLAG = False Then RSTREP.Close
    If RSTREP.State = 1 Then RSTREP.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDOUTWARD_DblClick()
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    If GRDOUTWARD.rows <= 1 Then Exit Sub
    
    If frmLogin.rs!Level = "2" Or frmLogin.rs!Level = "5" Then Exit Sub
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    Select Case Trim(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 11))
        Case "HI"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 13)) Then Exit Sub
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                If IsFormLoaded(frmsales) <> True Then
                    frmsales.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                    frmsales.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                    frmsales.Show
                    frmsales.SetFocus
                    Call frmsales.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMSALES1) <> True Then
                    FRMSALES1.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                    FRMSALES1.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                    FRMSALES1.Show
                    FRMSALES1.SetFocus
                    Call FRMSALES1.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMSALES2) <> True Then
                    FRMSALES2.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                    FRMSALES2.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                    FRMSALES2.Show
                    FRMSALES2.SetFocus
                    Call FRMSALES2.txtBillNo_KeyDown(13, 0)
                End If
            Else
                If SALESLT_FLAG = "Y" Then
                    If IsFormLoaded(FRMGSTRSM1) <> True Then
                        FRMGSTRSM1.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTRSM1.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTRSM1.Show
                        FRMGSTRSM1.SetFocus
                        Call FRMGSTRSM1.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTRSM2) <> True Then
                        FRMGSTRSM2.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTRSM2.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTRSM2.Show
                        FRMGSTRSM2.SetFocus
                        Call FRMGSTRSM2.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTRSM3) <> True Then
                        FRMGSTRSM3.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTRSM3.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTRSM3.Show
                        FRMGSTRSM3.SetFocus
                        Call FRMGSTRSM3.txtBillNo_KeyDown(13, 0)
                    End If
                Else
                    If IsFormLoaded(FRMGSTR) <> True Then
                        FRMGSTR.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTR.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTR.Show
                        FRMGSTR.SetFocus
                        Call FRMGSTR.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTR1) <> True Then
                        FRMGSTR1.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTR1.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTR1.Show
                        FRMGSTR1.SetFocus
                        Call FRMGSTR1.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTR2) <> True Then
                        FRMGSTR2.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTR2.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                        FRMGSTR2.Show
                        FRMGSTR2.SetFocus
                        Call FRMGSTR2.txtBillNo_KeyDown(13, 0)
                    End If
                End If
            End If
        Case "GI"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 13)) Then Exit Sub
            If IsFormLoaded(FRMGST) <> True Then
                FRMGST.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMGST.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMGST.Show
                FRMGST.SetFocus
                Call FRMGST.txtBillNo_KeyDown(13, 0)
            ElseIf IsFormLoaded(FRMGST1) <> True Then
                FRMGST1.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMGST1.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMGST1.Show
                FRMGST1.SetFocus
                Call FRMGST1.txtBillNo_KeyDown(13, 0)
            End If
        Case "SV"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 13)) Then Exit Sub
            If IsFormLoaded(FRMService) <> True Then
                FRMService.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMService.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMService.Show
                FRMService.SetFocus
                Call FRMService.txtBillNo_KeyDown(13, 0)
            ElseIf IsFormLoaded(FRMGST1) <> True Then
                FRMGST1.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMGST1.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMGST1.Show
                FRMGST1.SetFocus
                Call FRMGST1.txtBillNo_KeyDown(13, 0)
            End If
        Case "DN"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 13)) Then Exit Sub
            If IsFormLoaded(FRMDELIVERY) <> True Then
                FRMDELIVERY.txtBillNo.text = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMDELIVERY.LBLBILLNO.Caption = Val(GRDOUTWARD.TextMatrix(GRDOUTWARD.Row, 12))
                FRMDELIVERY.Show
                FRMDELIVERY.SetFocus
                Call FRMDELIVERY.txtBillNo_KeyDown(13, 0)
            End If
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub GRDSTOCK_DblClick()
    If GRDSTOCK.rows <= 1 Then Exit Sub
    If frmLogin.rs!Level = "2" Or frmLogin.rs!Level = "5" Then Exit Sub
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    
    Select Case Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 18))
        Case "PI"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 19)) Then Exit Sub
            If MDIMAIN.LBLSHOPRT.Caption = "Y" Then
                If IsFormLoaded(frmLPS) <> True Then
                    frmLPS.txtBillNo.text = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
                    frmLPS.Show
                    frmLPS.SetFocus
                    Call frmLPS.txtBillNo_KeyDown(13, 0)
                End If
            Else
                If MDIMAIN.lblcategory.Caption = "Y" Then
                    If IsFormLoaded(frmLP) <> True Then
                        frmLP.txtBillNo.text = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
                        frmLP.Show
                        frmLP.SetFocus
                        Call frmLP.txtBillNo_KeyDown(13, 0)
                    End If
                Else
                    If IsFormLoaded(frmLP1) <> True Then
                        frmLP1.txtBillNo.text = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
                        frmLP1.Show
                        frmLP1.SetFocus
                        Call frmLP1.txtBillNo_KeyDown(13, 0)
                    End If
                End If
            End If
        Case "OP"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 19)) Then Exit Sub
            If IsFormLoaded(frmOP) <> True Then
                    frmOP.txtBillNo.text = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
                    frmOP.Show
                    frmOP.SetFocus
                    Call frmOP.txtBillNo_KeyDown(13, 0)
                End If
        Case "WO"
    End Select
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            
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

Private Sub Label1_DblClick(index As Integer)
    If index <> 7 Then Exit Sub
    If Frmeall.Visible = True Then
        Frmeall.Visible = False
    Else
        Frmeall.Visible = True
    End If
End Sub

Private Sub OPTBALQTY_Click()
    Call Fillgrid
    Call Fillgrid2
End Sub

Private Sub optdate_Click()
    Screen.MousePointer = vbHourglass
    Call Fillgrid
    Call Fillgrid2
    Screen.MousePointer = vbNormal
End Sub

Private Sub OPTOUTCUST_Click()
    Screen.MousePointer = vbHourglass
    Call Fillgrid
    Call Fillgrid2
    Screen.MousePointer = vbNormal
End Sub

Private Sub OPTOUTDATE_Click()
    Screen.MousePointer = vbHourglass
    Call Fillgrid
    Call Fillgrid2
    Screen.MousePointer = vbNormal
End Sub

Private Sub OptSupplier_Click()
    Screen.MousePointer = vbHourglass
    Call Fillgrid
    Call Fillgrid2
    Screen.MousePointer = vbNormal
End Sub

Private Sub TXTCODE_Change()
    On Error GoTo ERRHAND
    If REPFLAG = True Then
        If Frmeall.Visible = False Then
            RSTREP.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_CODE Like '" & TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        Else
            RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_CODE Like '" & TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If Frmeall.Visible = False Then
            RSTREP.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_CODE Like '" & TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        Else
            RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_CODE Like '" & TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub tXTMEDICINE_Change()
    
    On Error GoTo ERRHAND
    If REPFLAG = True Then
        If Frmeall.Visible = False Then
            RSTREP.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        Else
            RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If Frmeall.Visible = False Then
            RSTREP.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        Else
            RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.text)
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
            If Trim(tXTMEDICINE.text) = "" Then
                TxtCode.SetFocus
            Else
                DataList2.SetFocus
            End If
        Case vbKeyEscape
            TxtItemName.SetFocus
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Public Sub DataList2_Click()
    LBLITEMCODE.text = DataList2.BoundText
    Call Fillgrid
    Call Fillgrid2
    lblbalance.Caption = Format(Round((Val(LBLOPQTY.Caption) + Val(LBLINWARD.Caption)) - ((Val(LBLOUTWARD.Caption) + Val(LblLoose.Caption) + Val(LBLWASTE.Caption))), 2), "0.00")
    lblbalance.Visible = True
    Screen.MousePointer = vbNormal
    ''''''''LBLBALANCE.Caption = Val(LBLINWARD.Caption) - Val(LBLOUTWARD.Caption)
End Sub

Private Function Fillgrid()
    Dim OPQTY, OPVAL, RCVD_OP As Double
    Dim rststock, RSTITEM As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim FROMDATE As Date
    Dim TODATE As Date
    
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTo.Value, "MM,DD,YYYY")
        
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set RSTITEM = New ADODB.Recordset
    RSTITEM.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEM.EOF And RSTITEM.BOF) Then
        OPQTY = 0
        OPVAL = 0
        
        OPQTY = IIf(IsNull(RSTITEM!OPEN_QTY), 0, RSTITEM!OPEN_QTY)
        OPVAL = IIf(IsNull(RSTITEM!OPEN_VAL), 0, RSTITEM!OPEN_VAL)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select SUM(QTY) FROM RTRXFILE WHERE ITEM_CODE = '" & RSTITEM!ITEM_CODE & "' AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OPQTY = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & RSTITEM!ITEM_CODE & "' AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        Do Until RSTTRXFILE.EOF
'            OPQTY = OPQTY + IIf(IsNull(RSTTRXFILE!QTY), 0, RSTTRXFILE!QTY)
'            OPVAL = OPVAL + IIf(IsNull(RSTTRXFILE!TRX_TOTAL), 0, RSTTRXFILE!TRX_TOTAL)
'            RSTTRXFILE.MoveNext
'        Loop
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        'RSTTRXFILE.Open "Select SUM(QTY * LOOSE_PACK) FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEM!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI' OR TRX_TYPE='GI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        RSTTRXFILE.Open "Select SUM((QTY + FREE_QTY) * LOOSE_PACK) FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEM!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RCVD_OP = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "Select SUM(FREE_QTY * LOOSE_PACK) FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEM!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI' OR TRX_TYPE='GI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'            RCVD_OP = RCVD_OP + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
'        End If
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
        
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEM!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI' OR TRX_TYPE='GI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        'rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI' OR TRX_TYPE='GI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DESC ASC, VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
'        Do Until RSTTRXFILE.EOF
'            RCVD_OP = RCVD_OP + ((RSTTRXFILE!QTY + IIf(IsNull(RSTTRXFILE!FREE_QTY), 0, RSTTRXFILE!FREE_QTY)) * IIf(IsNull(RSTTRXFILE!LOOSE_PACK) Or RSTTRXFILE!LOOSE_PACK = 0, 1, RSTTRXFILE!LOOSE_PACK))
'            'ISSVAL = ISSVAL + IIf(IsNull(RSTTRXFILE!TRX_TOTAL), 0, RSTTRXFILE!TRX_TOTAL)
'            'RCVD_OP = RCVD_OP + IIf(IsNull(RSTTRXFILE!FREE_QTY), 0, RSTTRXFILE!FREE_QTY)
'            'FREEVAL = FREEVAL + RSTTRXFILE!SALES_PRICE * FREEQTY
'            RSTTRXFILE.MoveNext
'        Loop
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
        
        OPQTY = OPQTY - RCVD_OP
    End If
    RSTITEM.Close
    Set RSTITEM = Nothing
    
    LBLOPQTY.Caption = OPQTY
    Dim rststock_bal As ADODB.Recordset
    Dim i As Long
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo ERRHAND
    
    i = 0
    Screen.MousePointer = vbHourglass
    LBLINWARD.Caption = ""
    lblinval.Caption = ""
    lblbalance.Caption = ""
    
    GRDSTOCK.rows = 1
    Set rststock = New ADODB.Recordset
    If optdate.Value = True Then rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    If optsupplier.Value = True Then rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY TRX_TYPE,VCH_DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    If OPTBALQTY.Value = True Then rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        Select Case rststock!TRX_TYPE
            Case "CN", "SR", "HI", "GI", "WO", "SV", "RW"
                GRDSTOCK.TextMatrix(i, 3) = "Sales Return"
                GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!VCH_DESC), "", Mid(rststock!VCH_DESC, 15))
            Case "XX", "OP", "ST"
                GRDSTOCK.TextMatrix(i, 3) = "OP. Stock"
            Case "WR"
                GRDSTOCK.TextMatrix(i, 3) = "Warranty Replacement"
                GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!VCH_DESC), "", Mid(rststock!VCH_DESC, 15))
            Case "MI"
                GRDSTOCK.TextMatrix(i, 3) = "Mixture"
                GRDSTOCK.TextMatrix(i, 4) = "Received from Production"
            Case "RM"
                GRDSTOCK.TextMatrix(i, 3) = "Mixture2"
                GRDSTOCK.TextMatrix(i, 4) = "Received from Production2"
            Case "PC"
                GRDSTOCK.TextMatrix(i, 3) = "Process"
                GRDSTOCK.TextMatrix(i, 4) = "Received from Process"
            Case "TF"
                GRDSTOCK.TextMatrix(i, 3) = "Stock Transfer"
                GRDSTOCK.TextMatrix(i, 4) = "Received Via Transfer"
            Case Else
                GRDSTOCK.TextMatrix(i, 3) = "Purchase"
                GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!VCH_DESC), "", Mid(rststock!VCH_DESC, 15))
        End Select
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!QTY), "", rststock!QTY) ''(rststock!QTY / rststock!UNIT) * rststock!LINE_DISC
        'GRDSTOCK.TextMatrix(i, 3) = rststock!BAL_QTY
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!VCH_DATE), "", rststock!VCH_DATE)
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!PINV), "", rststock!PINV)
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!VCH_NO), "", rststock!VCH_NO)
        If IsNull(rststock!UNIT) Then
            GRDSTOCK.TextMatrix(i, 9) = 1
        Else
            GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!UNIT), "1", rststock!UNIT)
        End If
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
        GRDSTOCK.TextMatrix(i, 11) = ""
        GRDSTOCK.TextMatrix(i, 12) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.000"))
        GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST, "0.000"))
        GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!ITEM_NET_COST_PRICE) Or rststock!ITEM_NET_COST_PRICE < Val(GRDSTOCK.TextMatrix(i, 13)), Val(GRDSTOCK.TextMatrix(i, 13)), Format(rststock!ITEM_NET_COST_PRICE, "0.000"))   'IIf(IsNull(rststock!ITEM_COST), 0, rststock!ITEM_COST + (rststock!ITEM_COST * rststock!SALES_TAX / 100)) + IIf(IsNull(rststock!ITEM_COST), 0, (rststock!ITEM_COST * IIf(IsNull(rststock!CESS_PER), 0, rststock!CESS_PER) / 100)) + IIf(IsNull(rststock!cess_amt), 0, rststock!cess_amt) 'IIf(IsNull(rststock!TRX_TOTAL), "", rststock!TRX_TOTAL / IIf(IsNull(rststock!QTY), "1", rststock!QTY))
'        If (IIf(IsNull(rststock!QTY), 0, rststock!QTY) + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY)) = 0 Then
'            GRDSTOCK.TextMatrix(i, 14) = Format(Round(Val(GRDSTOCK.TextMatrix(i, 14)), 3), "0.000")
'        Else
'            GRDSTOCK.TextMatrix(i, 14) = Format(Round(Val(GRDSTOCK.TextMatrix(i, 14)) + IIf(IsNull(rststock!EXPENSE), 0, rststock!EXPENSE) / ((IIf(IsNull(rststock!QTY), 0, rststock!QTY) + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY)) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)), 3), "0.000")
'        End If
        'grdsales.TextMatrix(Val(TXTSLNO.text), 8) = Format(Round(((Val(LblGross.Caption) / (Val(Los_Pack.text) * (Val(TXTQTY.text) + Val(TXTFREE.text)))) + ((Val(TxtExpense.text) / ((Val(TXTQTY.text) + Val(TXTFREE.text)) * Val(Los_Pack.text))))), 4), ".0000")
        'Format(Round(IIf(IsNull(rststock!ITEM_COST), 0, rststock!ITEM_COST + (rststock!ITEM_COST * rststock!SALES_TAX / 100)) + IIf(IsNull(rststock!ITEM_COST), 0, (rststock!ITEM_COST * rststock!CESS_PER / 100)) + IIf(IsNull(rststock!CESS_AMT), 0, rststock!CESS_AMT), 3), "0.000") 'IIf(IsNull(rststock!TRX_TOTAL), "", rststock!TRX_TOTAL / IIf(IsNull(rststock!QTY), "1", rststock!QTY))
        GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!BAL_QTY), 0, Format(rststock!BAL_QTY, "0.000"))
        GRDSTOCK.TextMatrix(i, 16) = IIf(IsNull(rststock!LOOSE_PACK), "", rststock!LOOSE_PACK)
        GRDSTOCK.TextMatrix(i, 17) = rststock!LINE_NO
        GRDSTOCK.TextMatrix(i, 18) = rststock!TRX_TYPE
        GRDSTOCK.TextMatrix(i, 19) = rststock!TRX_YEAR
        GRDSTOCK.TextMatrix(i, 20) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.000"))
        'LBLINWARD.Caption = ""
        If Not (IsNull(rststock!QTY)) Then
            LBLINWARD.Caption = Val(LBLINWARD.Caption) + rststock!QTY  '''(rststock!QTY / rststock!UNIT) * rststock!LINE_DISC
            lblinval.Caption = Val(lblinval.Caption) + (rststock!QTY * rststock!ITEM_COST)
        End If
    
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
'    LBLBALANCE.Caption = ""
'    Set rststock_bal = New ADODB.Recordset
'    rststock_bal.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
'    If Not (rststock_bal.EOF And rststock_bal.BOF) Then
'        Select Case rststock_bal!LOOSE_PACK
'            Case Is > 1
'                LBLBALANCE.Caption = Int(rststock_bal!CLOSE_QTY) & " + " & IIf(Val(Round((rststock_bal!CLOSE_QTY - Int(rststock_bal!CLOSE_QTY)) * rststock_bal!LOOSE_PACK, 0)) = 0, "", Round((rststock_bal!CLOSE_QTY - Int(rststock_bal!CLOSE_QTY)) * rststock_bal!LOOSE_PACK, 0) & " " & rststock_bal!PACK_TYPE)
'            Case Else
'                LBLBALANCE.Caption = Int(rststock_bal!CLOSE_QTY)
'        End Select
'
'        'LBLBALANCE.Caption = IIf(IsNull(rststock_bal!CLOSE_QTY), "", rststock_bal!CLOSE_QTY)
'    End If
'    rststock_bal.Close
'    Set rststock_bal = Nothing
    
    LBLINWARD.Caption = Format(Round(Val(LBLINWARD.Caption), 2), "0.00")
    LBLHEAD(9).Caption = "INWARD DETAILS FOR THE ITEM "
    LBLHEAD(0).Caption = DataList2.text
    Screen.MousePointer = vbNormal
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Function Fillgrid2()
    Dim rststock As ADODB.Recordset
    Dim RSTTEMP As ADODB.Recordset
    Dim M As Integer
    Dim E_DATE As Date
    Dim i, Full_Qty, Loose_qty As Double
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    LBLOUTWARD.Caption = ""
    lbloutval.Caption = ""
    LBLWASTE.Caption = ""
    LblLoose.Caption = ""
    Label1(3).Caption = ""
    
    'db.Execute "delete From TEMPTRX"
    'Set RSTTEMP = New ADODB.Recordset
    'RSTTEMP.Open "SELECT *  FROM TEMPTRX", db, adOpenStatic, adLockOptimistic, adCmdText
    i = 0
    Full_Qty = 0
    Loose_qty = 0
    GRDOUTWARD.rows = 1
    Set rststock = New ADODB.Recordset
    If OPTOUTDATE.Value = True Then rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    If OPTOUTCUST.Value = True Then rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DESC ASC, VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    'rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND CST <>2", db, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        i = i + 1
        GRDOUTWARD.rows = GRDOUTWARD.rows + 1
        GRDOUTWARD.FixedRows = 1
        GRDOUTWARD.TextMatrix(i, 0) = i
'        Select Case rststock!CST
'            Case 0
'               GRDOUTWARD.TextMatrix(i, 1) = "SALES"
'            Case 1
'               GRDOUTWARD.TextMatrix(i, 1) = "DELIVEREY"
'            Case 2
'               GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN"
'        End Select
        GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
        Select Case rststock!TRX_TYPE
            Case "SI"
                GRDOUTWARD.TextMatrix(i, 1) = "WHOLESALE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "RI"
                GRDOUTWARD.TextMatrix(i, 1) = "RETAIL"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "VI"
                GRDOUTWARD.TextMatrix(i, 1) = "VAN SALE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "GI"
                GRDOUTWARD.TextMatrix(i, 1) = "B2B"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "SV"
                GRDOUTWARD.TextMatrix(i, 1) = "Service Bills"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "HI"
                GRDOUTWARD.TextMatrix(i, 1) = "GST-RETAIL"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "TF"
                GRDOUTWARD.TextMatrix(i, 1) = "TRANSFER"
                GRDOUTWARD.TextMatrix(i, 2) = "Stock Transfer"
            Case "WO"
                GRDOUTWARD.TextMatrix(i, 1) = "WO"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "DN"
                GRDOUTWARD.TextMatrix(i, 1) = "DELIVERY"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "PR"
                GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "WP"
                GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "DG", "DM"
                GRDOUTWARD.TextMatrix(i, 1) = "DAMAGED GOODS"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 1, 30)
            Case "SR", "RW"
                GRDOUTWARD.TextMatrix(i, 1) = "To Service"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "GF"
                GRDOUTWARD.TextMatrix(i, 1) = "SAMPLE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "MI"
                GRDOUTWARD.TextMatrix(i, 1) = "FACTORY"
                GRDOUTWARD.TextMatrix(i, 2) = "FACTORY"
            Case "RM"
                GRDOUTWARD.TextMatrix(i, 1) = "FACTORY2"
                GRDOUTWARD.TextMatrix(i, 2) = "FACTORY2"
        End Select
        
        GRDOUTWARD.Tag = ""
        GRDSTOCK.Tag = ""
        Set RSTTEMP = New ADODB.Recordset
        RSTTEMP.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly
        If Not (RSTTEMP.EOF And RSTTEMP.BOF) Then
            GRDOUTWARD.Tag = IIf(IsNull(RSTTEMP!PACK_TYPE), "", RSTTEMP!PACK_TYPE)
            GRDSTOCK.Tag = IIf(IsNull(RSTTEMP!LOOSE_PACK), "", RSTTEMP!LOOSE_PACK)
        End If
        RSTTEMP.Close
        Set RSTTEMP = Nothing
        
        
        GRDOUTWARD.TextMatrix(i, 4) = IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY)
        GRDOUTWARD.TextMatrix(i, 5) = IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
        If rststock!TRX_TYPE = "GI" And rststock!CST = 1 Then
            GRDOUTWARD.TextMatrix(i, 3) = IIf(IsNull(rststock!QTY), 0, rststock!QTY) & "*"
            'Full_Qty = Full_Qty + (Val(GRDOUTWARD.TextMatrix(i, 3)) + Val(GRDOUTWARD.TextMatrix(i, 4))) * GRDOUTWARD.TextMatrix(i, 5)
        Else
            GRDOUTWARD.TextMatrix(i, 3) = IIf(IsNull(rststock!QTY), 0, rststock!QTY)
            Full_Qty = Full_Qty + (Val(GRDOUTWARD.TextMatrix(i, 3)) + Val(GRDOUTWARD.TextMatrix(i, 4))) * GRDOUTWARD.TextMatrix(i, 5)
        End If
        
        'Full_Qty = Full_Qty + (Val(GRDOUTWARD.TextMatrix(i, 3)) + Val(GRDOUTWARD.TextMatrix(i, 4))) * GRDOUTWARD.TextMatrix(i, 5)
        
        'GRDOUTWARD.TextMatrix(i, 3) = rststock!QTY
        'GRDOUTWARD.TextMatrix(i, 4) = IIf(IsNull(rststock!FREE_QTY), "", rststock!FREE_QTY)
        GRDOUTWARD.TextMatrix(i, 6) = Format(rststock!P_RETAIL, "0.00")
        GRDOUTWARD.TextMatrix(i, 7) = rststock!TRX_TYPE & "-" & rststock!VCH_NO
        GRDOUTWARD.TextMatrix(i, 8) = Format(rststock!VCH_DATE, "dd/mm/yyyy")
        GRDOUTWARD.TextMatrix(i, 9) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
        GRDOUTWARD.TextMatrix(i, 10) = IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
        GRDOUTWARD.TextMatrix(i, 11) = rststock!TRX_TYPE
        GRDOUTWARD.TextMatrix(i, 12) = rststock!VCH_NO
        GRDOUTWARD.TextMatrix(i, 13) = rststock!TRX_YEAR
        GRDOUTWARD.TextMatrix(i, 14) = IIf(IsNull(rststock!INV_DETAILS), "", rststock!INV_DETAILS)
        LBLWASTE.Caption = Val(LBLWASTE.Caption) + Val(GRDOUTWARD.TextMatrix(i, 10))
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    LBLOUTWARD.Caption = Full_Qty
    If Loose_qty > 0 Then
        LblLoose.Visible = True
        Label1(2).Visible = True
        Label1(3).Visible = True
        LblLoose.Caption = Loose_qty
        Label1(3).Caption = GRDOUTWARD.Tag
    Else
        LblLoose.Visible = False
        Label1(2).Visible = False
        Label1(3).Visible = False
        Label1(3).Caption = ""
        LblLoose.Caption = ""
    End If
    
    LBLOUTWARD.Caption = Format(Round(Val(LBLOUTWARD.Caption), 2), "0.00")
    LBLHEAD(1).Caption = "OUTWARD DETAILS"
    Screen.MousePointer = vbNormal
    'If Val(LBLINWARD.Caption) - Val(LBLOUTWARD.Caption) <> Val(LBLBALANCE.Caption) Then lblmanual.Visible = True Else lblmanual.Visible = False
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
            If Trim(TxtCode.text) = "" Then
                TxtItemName.SetFocus
            Else
                DataList2.SetFocus
            End If
        Case vbKeyEscape
            tXTMEDICINE.SetFocus
    End Select

End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TxtItemName_Change()
    
    On Error GoTo ERRHAND
    If REPFLAG = True Then
        If Frmeall.Visible = False Then
            RSTREP.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        Else
            RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If Frmeall.Visible = False Then
            RSTREP.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        Else
            RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TxtItemName_GotFocus()
    TxtItemName.SelStart = 0
    TxtItemName.SelLength = Len(TxtItemName.text)
End Sub

Private Sub TxtItemName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
            If Trim(TxtItemName.text) = "" Then
                tXTMEDICINE.SetFocus
            Else
                DataList2.SetFocus
            End If
        Case vbKeyEscape
            Call cmdexit_Click
    End Select

End Sub

Private Sub TxtItemName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TxtBarcode_Change()
'    On Error GoTo Errhand
'    If REPFLAG = True Then
'        RSTREP.Open "Select DISTINCT BARCODE From RTRXFILE  WHERE BARCODE Like '" & TxtBarcode.Text & "%' ", db, adOpenStatic, adLockReadOnly
'        REPFLAG = False
'    Else
'        RSTREP.Close
'        RSTREP.Open "Select DISTINCT BARCODE From RTRXFILE  WHERE BARCODE Like '" & TxtBarcode.Text & "%' ", db, adOpenStatic, adLockReadOnly
'        REPFLAG = False
'    End If
'    Set Me.DataList2.RowSource = RSTREP
'    DataList2.ListField = "BARCODE"
'    DataList2.BoundColumn = "BARCODE"
'
'    Exit Sub
''RSTREP.Close
''TMPFLAG = False
'Errhand:
'    MsgBox Err.Description
End Sub

Private Sub txtbarcode_GotFocus()
    TxtBarcode.SelStart = 0
    TxtBarcode.SelLength = Len(TxtBarcode.text)
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TxtBarcode.text) = "" Then Exit Sub
            Call Fillgrid3
            Call FILLGRID4
            'LBLBALANCE.Caption = Format(Round((Val(LBLOPQTY.Caption) + Val(LBLINWARD.Caption)) - ((Val(LBLOUTWARD.Caption) + Val(LblLoose.Caption) + Val(LBLWASTE.Caption))), 2), "0.00")
            lblbalance.Visible = False
            Screen.MousePointer = vbNormal
        Case vbKeyEscape
            tXTMEDICINE.SetFocus
    End Select

End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Function Fillgrid3()
    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And Frmeall.Visible = False Then Exit Function
    Dim OPQTY, OPVAL, RCVD_OP As Double
    Dim rststock, RSTITEM As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim FROMDATE As Date
    Dim TODATE As Date
    
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTo.Value, "MM,DD,YYYY")
        
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
'    Set RSTITEM = New ADODB.Recordset
'    RSTITEM.Open "SELECT *  FROM rtrxfile WHERE BARCODE = '" & Trim(TxtBarcode.Text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'    If Not (RSTITEM.EOF And RSTITEM.BOF) Then
'        OPQTY = 0
'        OPVAL = 0
'
'        'OPQTY = IIf(IsNull(RSTITEM!OPEN_QTY), 0, RSTITEM!OPEN_QTY)
'        'OPVAL = IIf(IsNull(RSTITEM!OPEN_VAL), 0, RSTITEM!OPEN_VAL)
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE BARCODE = '" & RSTITEM!BARCODE & "' AND VCH_DATE <'" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        Do Until RSTTRXFILE.EOF
'            OPQTY = OPQTY + IIf(IsNull(RSTTRXFILE!QTY), 0, RSTTRXFILE!QTY)
'            OPVAL = OPVAL + IIf(IsNull(RSTTRXFILE!TRX_TOTAL), 0, RSTTRXFILE!TRX_TOTAL)
'            RSTTRXFILE.MoveNext
'        Loop
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
'
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "SELECT * FROM TRXFILE WHERE  BARCODE = '" & RSTITEM!BARCODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='PR' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI') AND VCH_DATE <'" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        Do Until RSTTRXFILE.EOF
'            RCVD_OP = RCVD_OP + RSTTRXFILE!QTY
'            'ISSVAL = ISSVAL + IIf(IsNull(RSTTRXFILE!TRX_TOTAL), 0, RSTTRXFILE!TRX_TOTAL)
'            RCVD_OP = RCVD_OP + IIf(IsNull(RSTTRXFILE!FREE_QTY), 0, RSTTRXFILE!FREE_QTY)
'            'FREEVAL = FREEVAL + RSTTRXFILE!SALES_PRICE * FREEQTY
'            RSTTRXFILE.MoveNext
'        Loop
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
'
'        OPQTY = OPQTY - RCVD_OP
'    End If
'    RSTITEM.Close
'    Set RSTITEM = Nothing
    
    LBLOPQTY.Caption = OPQTY
    Dim rststock_bal As ADODB.Recordset
    Dim i As Long
    
    If Trim(TxtBarcode.text) = "" Then Exit Function
    On Error GoTo ERRHAND
    
    i = 0
    Screen.MousePointer = vbHourglass
    LBLINWARD.Caption = ""
    lblinval.Caption = ""
    lblbalance.Caption = ""
    
    GRDSTOCK.rows = 1
    Set rststock = New ADODB.Recordset
    If optdate.Value = True Then rststock.Open "SELECT * FROM RTRXFILE WHERE  BARCODE = '" & Trim(TxtBarcode.text) & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    If optsupplier.Value = True Then rststock.Open "SELECT * FROM RTRXFILE WHERE  BARCODE = '" & Trim(TxtBarcode.text) & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY TRX_TYPE,VCH_DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    If OPTBALQTY.Value = True Then rststock.Open "SELECT * FROM RTRXFILE WHERE  BARCODE = '" & Trim(TxtBarcode.text) & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        Select Case rststock!TRX_TYPE
            Case "CN", "SR", "HI", "GI", "WO", "SV", "RW"
                GRDSTOCK.TextMatrix(i, 3) = "Sales Return"
                GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!VCH_DESC), "", Mid(rststock!VCH_DESC, 15))
            Case "XX", "OP", "ST"
                GRDSTOCK.TextMatrix(i, 3) = "OP. Stock"
            Case "WR"
                GRDSTOCK.TextMatrix(i, 3) = "Warranty Replacement"
                GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!VCH_DESC), "", Mid(rststock!VCH_DESC, 15))
            Case "MI"
                GRDSTOCK.TextMatrix(i, 3) = "Mixture"
                GRDSTOCK.TextMatrix(i, 4) = "Received from Production"
            Case "RM"
                GRDSTOCK.TextMatrix(i, 3) = "Mixture2"
                GRDSTOCK.TextMatrix(i, 4) = "Received from Production2"
            Case "PC"
                GRDSTOCK.TextMatrix(i, 3) = "Process"
                GRDSTOCK.TextMatrix(i, 4) = "Received from Process"
            Case "TF"
                GRDSTOCK.TextMatrix(i, 3) = "Stock Transfer"
                GRDSTOCK.TextMatrix(i, 4) = "Received Via Transfer"
            Case Else
                GRDSTOCK.TextMatrix(i, 3) = "Purchase"
                GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!VCH_DESC), "", Mid(rststock!VCH_DESC, 15))
        End Select
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!QTY), "", rststock!QTY) ''(rststock!QTY / rststock!UNIT) * rststock!LINE_DISC
        'GRDSTOCK.TextMatrix(i, 3) = rststock!BAL_QTY
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!VCH_DATE), "", rststock!VCH_DATE)
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!PINV), "", rststock!PINV)
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!VCH_NO), "", rststock!VCH_NO)
        If IsNull(rststock!UNIT) Then
            GRDSTOCK.TextMatrix(i, 9) = 1
        Else
            GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!UNIT), "1", rststock!UNIT)
        End If
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
        GRDSTOCK.TextMatrix(i, 11) = ""
        GRDSTOCK.TextMatrix(i, 12) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.000"))
        GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST, "0.000"))
        GRDSTOCK.TextMatrix(i, 14) = Format(Round(IIf(IsNull(rststock!ITEM_COST), 0, rststock!ITEM_COST + (rststock!ITEM_COST * rststock!SALES_TAX / 100)) + IIf(IsNull(rststock!ITEM_COST), 0, (rststock!ITEM_COST * IIf(IsNull(rststock!CESS_PER), 0, rststock!CESS_PER) / 100)) + IIf(IsNull(rststock!cess_amt), 0, rststock!cess_amt), 3), "0.000")  'IIf(IsNull(rststock!TRX_TOTAL), "", rststock!TRX_TOTAL / IIf(IsNull(rststock!QTY), "1", rststock!QTY))
        'Format(Round(IIf(IsNull(rststock!ITEM_COST), 0, rststock!ITEM_COST + (rststock!ITEM_COST * rststock!SALES_TAX / 100)) + IIf(IsNull(rststock!ITEM_COST), 0, (rststock!ITEM_COST * rststock!CESS_PER / 100)) + IIf(IsNull(rststock!CESS_AMT), 0, rststock!CESS_AMT), 3), "0.000") 'IIf(IsNull(rststock!TRX_TOTAL), "", rststock!TRX_TOTAL / IIf(IsNull(rststock!QTY), "1", rststock!QTY))
        GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!BAL_QTY), 0, Format(rststock!BAL_QTY, "0.000"))
        GRDSTOCK.TextMatrix(i, 16) = IIf(IsNull(rststock!LOOSE_PACK), "", rststock!LOOSE_PACK)
        GRDSTOCK.TextMatrix(i, 17) = rststock!LINE_NO
        GRDSTOCK.TextMatrix(i, 18) = rststock!TRX_TYPE
        GRDSTOCK.TextMatrix(i, 19) = rststock!TRX_YEAR
        'LBLINWARD.Caption = ""
        If Not (IsNull(rststock!QTY)) Then
            LBLINWARD.Caption = Val(LBLINWARD.Caption) + rststock!QTY  '''(rststock!QTY / rststock!UNIT) * rststock!LINE_DISC
            lblinval.Caption = Val(lblinval.Caption) + (rststock!QTY * rststock!ITEM_COST)
        End If
    
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
'    LBLBALANCE.Caption = ""
'    Set rststock_bal = New ADODB.Recordset
'    rststock_bal.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & Trim(TxtBarcode.Text) & "'", db, adOpenStatic, adLockReadOnly
'    If Not (rststock_bal.EOF And rststock_bal.BOF) Then
'        Select Case rststock_bal!LOOSE_PACK
'            Case Is > 1
'                LBLBALANCE.Caption = Int(rststock_bal!CLOSE_QTY) & " + " & IIf(Val(Round((rststock_bal!CLOSE_QTY - Int(rststock_bal!CLOSE_QTY)) * rststock_bal!LOOSE_PACK, 0)) = 0, "", Round((rststock_bal!CLOSE_QTY - Int(rststock_bal!CLOSE_QTY)) * rststock_bal!LOOSE_PACK, 0) & " " & rststock_bal!PACK_TYPE)
'            Case Else
'                LBLBALANCE.Caption = Int(rststock_bal!CLOSE_QTY)
'        End Select
'
'        'LBLBALANCE.Caption = IIf(IsNull(rststock_bal!CLOSE_QTY), "", rststock_bal!CLOSE_QTY)
'    End If
'    rststock_bal.Close
'    Set rststock_bal = Nothing
    
    LBLINWARD.Caption = Format(Round(Val(LBLINWARD.Caption), 2), "0.00")
    LBLHEAD(9).Caption = "INWARD DETAILS FOR THE ITEM "
    LBLHEAD(0).Caption = DataList2.text
    Screen.MousePointer = vbNormal
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Function FILLGRID4()
    
    If Not (UCase(DUPCODE) = "DUP" Or DUPCODE = "") And Frmeall.Visible = False Then Exit Function
    Dim rststock As ADODB.Recordset
    Dim RSTTEMP As ADODB.Recordset
    Dim M As Integer
    Dim E_DATE As Date
    Dim i, Full_Qty, Loose_qty As Double
    
    If Trim(TxtBarcode.text) = "" Then Exit Function
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    LBLOUTWARD.Caption = ""
    lbloutval.Caption = ""
    LBLWASTE.Caption = ""
    LblLoose.Caption = ""
    Label1(3).Caption = ""
    
    'db.Execute "delete From TEMPTRX"
    'Set RSTTEMP = New ADODB.Recordset
    'RSTTEMP.Open "SELECT *  FROM TEMPTRX", db, adOpenStatic, adLockOptimistic, adCmdText
    i = 0
    Full_Qty = 0
    Loose_qty = 0
    GRDOUTWARD.rows = 1
    Set rststock = New ADODB.Recordset
    If OPTOUTDATE.Value = True Then rststock.Open "SELECT * FROM TRXFILE WHERE  BARCODE = '" & Trim(TxtBarcode.text) & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    If OPTOUTCUST.Value = True Then rststock.Open "SELECT * FROM TRXFILE WHERE  BARCODE = '" & Trim(TxtBarcode.text) & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DESC ASC, VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    'rststock.Open "SELECT * FROM TRXFILE WHERE  BARCODE = '" & Trim(TxtBarcode.Text) & "' AND CST <>2", db, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        i = i + 1
        GRDOUTWARD.rows = GRDOUTWARD.rows + 1
        GRDOUTWARD.FixedRows = 1
        GRDOUTWARD.TextMatrix(i, 0) = i
'        Select Case rststock!CST
'            Case 0
'               GRDOUTWARD.TextMatrix(i, 1) = "SALES"
'            Case 1
'               GRDOUTWARD.TextMatrix(i, 1) = "DELIVEREY"
'            Case 2
'               GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN"
'        End Select
        GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
        Select Case rststock!TRX_TYPE
            Case "SI"
                GRDOUTWARD.TextMatrix(i, 1) = "WHOLESALE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "RI"
                GRDOUTWARD.TextMatrix(i, 1) = "RETAIL"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "VI"
                GRDOUTWARD.TextMatrix(i, 1) = "VAN SALE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "GI"
                GRDOUTWARD.TextMatrix(i, 1) = "B2B"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "SV"
                GRDOUTWARD.TextMatrix(i, 1) = "Service Bills"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "HI"
                GRDOUTWARD.TextMatrix(i, 1) = "GST-RETAIL"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "TF"
                GRDOUTWARD.TextMatrix(i, 1) = "TRANSFER"
                GRDOUTWARD.TextMatrix(i, 2) = "Stock Transfer"
            Case "WO"
                GRDOUTWARD.TextMatrix(i, 1) = "WO"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "DN"
                GRDOUTWARD.TextMatrix(i, 1) = "DELIVERY"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "PR"
                GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "WP"
                GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN(W)"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "DG", "DM"
                GRDOUTWARD.TextMatrix(i, 1) = "DAMAGED GOODS"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "SR", "RW"
                GRDOUTWARD.TextMatrix(i, 1) = "To Service"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "GF"
                GRDOUTWARD.TextMatrix(i, 1) = "SAMPLE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "MI"
                GRDOUTWARD.TextMatrix(i, 1) = "FACTORY"
                GRDOUTWARD.TextMatrix(i, 2) = "FACTORY"
            Case "RM"
                GRDOUTWARD.TextMatrix(i, 1) = "FACTORY2"
                GRDOUTWARD.TextMatrix(i, 2) = "FACTORY2"
        End Select
        
        GRDOUTWARD.Tag = ""
        GRDSTOCK.Tag = ""
        Set RSTTEMP = New ADODB.Recordset
        RSTTEMP.Open "SELECT * FROM ITEMMAST WHERE  BARCODE = '" & Trim(TxtBarcode.text) & "' ", db, adOpenStatic, adLockReadOnly
        If Not (RSTTEMP.EOF And RSTTEMP.BOF) Then
            GRDOUTWARD.Tag = IIf(IsNull(RSTTEMP!PACK_TYPE), "", RSTTEMP!PACK_TYPE)
            GRDSTOCK.Tag = IIf(IsNull(RSTTEMP!LOOSE_PACK), "", RSTTEMP!LOOSE_PACK)
        End If
        RSTTEMP.Close
        Set RSTTEMP = Nothing
        
        GRDOUTWARD.TextMatrix(i, 3) = IIf(IsNull(rststock!QTY), 0, rststock!QTY)
        GRDOUTWARD.TextMatrix(i, 4) = IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY)
        GRDOUTWARD.TextMatrix(i, 5) = IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
        Full_Qty = Full_Qty + (Val(GRDOUTWARD.TextMatrix(i, 3)) + Val(GRDOUTWARD.TextMatrix(i, 4))) * GRDOUTWARD.TextMatrix(i, 5)
        
        'GRDOUTWARD.TextMatrix(i, 3) = rststock!QTY
        'GRDOUTWARD.TextMatrix(i, 4) = IIf(IsNull(rststock!FREE_QTY), "", rststock!FREE_QTY)
        GRDOUTWARD.TextMatrix(i, 6) = Format(rststock!P_RETAIL, "0.00")
        GRDOUTWARD.TextMatrix(i, 7) = rststock!TRX_TYPE & "-" & rststock!VCH_NO
        GRDOUTWARD.TextMatrix(i, 8) = Format(rststock!VCH_DATE, "dd/mm/yyyy")
        GRDOUTWARD.TextMatrix(i, 9) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
        GRDOUTWARD.TextMatrix(i, 10) = IIf(IsNull(rststock!WASTE_QTY), 0, rststock!WASTE_QTY)
        GRDOUTWARD.TextMatrix(i, 11) = rststock!TRX_TYPE
        GRDOUTWARD.TextMatrix(i, 12) = rststock!VCH_NO
        GRDOUTWARD.TextMatrix(i, 13) = rststock!TRX_YEAR
        GRDOUTWARD.TextMatrix(i, 14) = IIf(IsNull(rststock!INV_DETAILS), "", rststock!INV_DETAILS)
        LBLWASTE.Caption = Val(LBLWASTE.Caption) + Val(GRDOUTWARD.TextMatrix(i, 10))
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    LBLOUTWARD.Caption = Full_Qty
    If Loose_qty > 0 Then
        LblLoose.Visible = True
        Label1(2).Visible = True
        Label1(3).Visible = True
        LblLoose.Caption = Loose_qty
        Label1(3).Caption = GRDOUTWARD.Tag
    Else
        LblLoose.Visible = False
        Label1(2).Visible = False
        Label1(3).Visible = False
        Label1(3).Caption = ""
        LblLoose.Caption = ""
    End If
    
    LBLOUTWARD.Caption = Format(Round(Val(LBLOUTWARD.Caption), 2), "0.00")
    LBLHEAD(1).Caption = "OUTWARD DETAILS"
    Screen.MousePointer = vbNormal
    'If Val(LBLINWARD.Caption) - Val(LBLOUTWARD.Caption) <> Val(LBLBALANCE.Caption) Then lblmanual.Visible = True Else lblmanual.Visible = False
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function
