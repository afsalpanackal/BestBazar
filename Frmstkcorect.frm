VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmStockCorrect 
   BackColor       =   &H00E8DFEC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price Analysis"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18705
   ClipControls    =   0   'False
   Icon            =   "Frmstkcorect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   18705
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export Items"
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
      Left            =   8595
      TabIndex        =   46
      Top             =   1515
      Width           =   1155
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Assign Company"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   15375
      TabIndex        =   45
      Top             =   915
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Assign Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14220
      TabIndex        =   44
      Top             =   930
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Import Items"
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
      Left            =   9765
      TabIndex        =   43
      Top             =   1530
      Width           =   1155
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   9195
      TabIndex        =   37
      Top             =   -60
      Width           =   1755
      Begin VB.CheckBox chkhidecomp 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   40
         TabIndex        =   42
         Top             =   1125
         Width           =   1635
      End
      Begin VB.CheckBox chkhidecat 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   40
         TabIndex        =   41
         Top             =   885
         Width           =   1635
      End
      Begin VB.CheckBox chklwp 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide LW Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   40
         TabIndex        =   40
         Top             =   645
         Width           =   1635
      End
      Begin VB.CheckBox chkvp 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide VP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   40
         TabIndex        =   39
         Top             =   405
         Width           =   1635
      End
      Begin VB.CheckBox chkws 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide WS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   40
         TabIndex        =   38
         Top             =   165
         Width           =   1635
      End
   End
   Begin VB.TextBox TxtHSNCODE 
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
      Left            =   5775
      TabIndex        =   4
      Top             =   270
      Width           =   990
   End
   Begin VB.TextBox TxtTax 
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
      Left            =   4905
      TabIndex        =   3
      Top             =   270
      Width           =   855
   End
   Begin VB.TextBox TxtItemcode 
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
      Left            =   3750
      TabIndex        =   2
      Top             =   270
      Width           =   1140
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete Item (Ctrl +D)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   15375
      TabIndex        =   30
      Top             =   1365
      Width           =   1140
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New Item (Ctrl +I)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   14220
      TabIndex        =   29
      Top             =   1365
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Assign HSN to all"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14220
      TabIndex        =   28
      Top             =   495
      Width           =   1095
   End
   Begin VB.TextBox TxtHSN 
      Alignment       =   2  'Center
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
      Height          =   405
      Left            =   15330
      TabIndex        =   27
      Top             =   480
      Width           =   1155
   End
   Begin VB.CheckBox chkcategory 
      BackColor       =   &H00E8DFEC&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   12735
      TabIndex        =   24
      Top             =   45
      Width           =   1410
   End
   Begin VB.TextBox TXTDEALER2 
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
      Height          =   330
      Left            =   10950
      TabIndex        =   20
      Top             =   300
      Width           =   3225
   End
   Begin VB.CheckBox CHKCATEGORY2 
      BackColor       =   &H00E8DFEC&
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   10980
      TabIndex        =   19
      Top             =   45
      Width           =   1590
   End
   Begin VB.TextBox TxtDisc 
      Alignment       =   2  'Center
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
      Height          =   405
      Left            =   15330
      TabIndex        =   18
      Top             =   45
      Width           =   1155
   End
   Begin VB.CommandButton CmdDisc 
      Caption         =   "Assign &Tax to all"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14220
      TabIndex        =   17
      Top             =   45
      Width           =   1095
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Re- Load"
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
      Left            =   6810
      TabIndex        =   12
      Top             =   975
      Width           =   1215
   End
   Begin VB.TextBox TxtCode 
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
      Left            =   2685
      TabIndex        =   1
      Top             =   270
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1035
      Left            =   6795
      TabIndex        =   7
      Top             =   -75
      Width           =   2385
      Begin VB.OptionButton OptPC 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Price Changing Items"
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
         Height          =   315
         Left            =   30
         TabIndex        =   26
         Top             =   660
         Width           =   2340
      End
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Display All"
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
         Height          =   315
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptStock 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Stock Items Only"
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
         Height          =   315
         Left            =   30
         TabIndex        =   8
         Top             =   390
         Width           =   1935
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
      Left            =   45
      TabIndex        =   0
      Top             =   270
      Width           =   2625
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
      Height          =   405
      Left            =   8055
      TabIndex        =   6
      Top             =   975
      Width           =   1125
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   45
      TabIndex        =   5
      Top             =   630
      Width           =   6720
      _ExtentX        =   11853
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
   Begin VB.Frame Frame1 
      Height          =   5985
      Left            =   45
      TabIndex        =   10
      Top             =   2295
      Width           =   18645
      Begin MSDataGridLib.DataGrid grdmsc 
         Height          =   5820
         Left            =   15
         TabIndex        =   47
         Top             =   120
         Width           =   18585
         _ExtentX        =   32782
         _ExtentY        =   10266
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   19
         RowDividerStyle =   4
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
   Begin MSComCtl2.DTPicker DTFROM 
      Height          =   375
      Left            =   6795
      TabIndex        =   15
      Top             =   1575
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   0
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   255
      CheckBox        =   -1  'True
      Format          =   117440513
      CurrentDate     =   40498
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   780
      Left            =   10950
      TabIndex        =   21
      Top             =   645
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   1376
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
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
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   5
      Left            =   5775
      TabIndex        =   36
      Top             =   15
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Tax"
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
      Index           =   1
      Left            =   4905
      TabIndex        =   35
      Top             =   15
      Width           =   855
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
      Index           =   0
      Left            =   3750
      TabIndex        =   34
      Top             =   15
      Width           =   1140
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
      Left            =   45
      TabIndex        =   33
      Top             =   15
      Width           =   3690
   End
   Begin VB.Label lblnetvalue 
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
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   10965
      TabIndex        =   32
      Top             =   1905
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   4
      Left            =   10980
      TabIndex        =   31
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label lblitemname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   6795
      TabIndex        =   25
      Top             =   1950
      Width           =   4140
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   0
      TabIndex        =   23
      Top             =   810
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   0
      TabIndex        =   22
      Top             =   450
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OP. Stock Entry Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Index           =   3
      Left            =   6735
      TabIndex        =   16
      Top             =   1380
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   6
      Left            =   12885
      TabIndex        =   14
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label lblpvalue 
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
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   12915
      TabIndex        =   13
      Top             =   1905
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   300
      Index           =   2
      Left            =   60
      TabIndex        =   11
      Top             =   660
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "FrmStockCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPFLAG As Boolean 'REP
Dim MFG_REC As New ADODB.Recordset
Dim CAT_REC As New ADODB.Recordset
Dim RSTREP As New ADODB.Recordset
Dim PHY_FLAG As Boolean 'REP
Dim PHY_REC As New ADODB.Recordset
Dim frmloadflag As Boolean
Private adoGrid As ADODB.Recordset
Dim SEARCHFLAG As Integer

Private Sub CHKCATEGORY_Click()
    CHKCATEGORY2.Value = 0
End Sub

Private Sub CHKCATEGORY2_Click()
    chkcategory.Value = 0
End Sub

Private Sub chkhidecat_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo ERRHAND
    If chkhidecat.Value = 1 Then
        db.Execute "Update COMPINFO set hide_category = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 18) = ""
'        GRDSTOCK.ColWidth(18) = 0
    Else
        db.Execute "Update COMPINFO set hide_category = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 18) = "Category"
'        GRDSTOCK.ColWidth(18) = 1000
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub chkhidecomp_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo ERRHAND
    If chkhidecomp.Value = 1 Then
        db.Execute "Update COMPINFO set hide_company = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 19) = ""
'        GRDSTOCK.ColWidth(19) = 0
    Else
        db.Execute "Update COMPINFO set hide_company = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 19) = "Company"
'        GRDSTOCK.ColWidth(19) = 1000
    End If
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub chklwp_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo ERRHAND
    If chklwp.Value = 1 Then
        db.Execute "Update COMPINFO set hide_lwp = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 17) = ""
'        GRDSTOCK.ColWidth(17) = 0
    Else
        db.Execute "Update COMPINFO set hide_lwp = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 17) = "L.W.Price"
'        GRDSTOCK.ColWidth(17) = 1000
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub chkvp_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo ERRHAND
    If chkvp.Value = 1 Then
        db.Execute "Update COMPINFO set hide_van = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 9) = ""
'        GRDSTOCK.ColWidth(9) = 0
    Else
        db.Execute "Update COMPINFO set hide_van = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 9) = "VP"
'        GRDSTOCK.ColWidth(9) = 1000
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub chkws_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo ERRHAND
    If chkws.Value = 1 Then
        db.Execute "Update COMPINFO set hide_ws = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 8) = ""
'        GRDSTOCK.ColWidth(8) = 0
    Else
        db.Execute "Update COMPINFO set hide_ws = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 8) = "WS"
'        GRDSTOCK.ColWidth(8) = 1000
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub


Private Sub CmdDisc_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If Trim(TxtDisc.text) = "" Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THESE TAX", vbYesNo + vbDefaultButton2, "Assign TAX....") = vbNo Then Exit Sub
'    For i = 1 To GRDSTOCK.Rows - 1
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (rststock.EOF And rststock.BOF) Then
'            rststock!SALES_TAX = Val(TxtDisc.Text)
'            rststock!CHECK_FLAG = "V"
'            'rststock!P_RETAIL = rststock!MRP
'            GRDSTOCK.TextMatrix(i, 10) = Val(TxtDisc.Text)
'            rststock.Update
'        End If
'        rststock.Close
'        Set rststock = Nothing
'
''        Set rststock = New ADODB.Recordset
''        rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' WHERE BAL_QTY >0", db, adOpenStatic, adLockOptimistic, adCmdText
''        Do Until rststock.EOF
''            rststock!CUST_DISC = Val(TxtDisc.Text)
''            'rststock!P_RETAIL = rststock!MRP
''            GRDSTOCK.TextMatrix(i, 17) = Val(TxtDisc.Text)
''            rststock.Update
''            rststock.MoveNext
''        Loop
''        rststock.Close
''        Set rststock = Nothing
'
'    Next i
    TxtDisc.text = ""
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

'Private Sub CMDEXPORT_Click()
'    If frmLogin.rs!Level <> "0" Then MsgBox "Permission Denied", vbOKOnly, "Price Analysis"
'    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
'    Dim oApp As Excel.Application
'    Dim oWB As Excel.Workbook
'    Dim oWS As Excel.Worksheet
'    Dim xlRange As Excel.Range
'    Dim i, N As Long
'
'    On Error GoTo eRRHAND
'    Screen.MousePointer = vbHourglass
'    'Create an Excel instalce.
'    Set oApp = CreateObject("Excel.Application")
'    Set oWB = oApp.Workbooks.Add
'    Set oWS = oWB.Worksheets(1)
'
'
'
'
''    xlRange = oWS.Range("A1", "C1")
''    xlRange.Font.Bold = True
''    xlRange.ColumnWidth = 15
''    'xlRange.Value = {"First Name", "Last Name", "Last Service"}
''    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
''    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
''
''    xlRange = oWS.Range("C1", "C999")
''    xlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
''    xlRange.ColumnWidth = 12
'
'    'If Sum_flag = False Then
'        oWS.Range("A1", "J1").Merge
'        oWS.Range("A1", "J1").HorizontalAlignment = xlCenter
'        oWS.Range("A2", "J2").Merge
'        oWS.Range("A2", "J2").HorizontalAlignment = xlCenter
'    'End If
'    oWS.Range("A:A").ColumnWidth = 6
'    oWS.Range("B:B").ColumnWidth = 10
'    oWS.Range("C:C").ColumnWidth = 12
'    oWS.Range("D:D").ColumnWidth = 12
'    oWS.Range("E:E").ColumnWidth = 12
'    oWS.Range("F:F").ColumnWidth = 12
'    oWS.Range("G:G").ColumnWidth = 12
'    oWS.Range("H:H").ColumnWidth = 12
'    oWS.Range("I:I").ColumnWidth = 12
'    oWS.Range("J:J").ColumnWidth = 12
'    oWS.Range("K:K").ColumnWidth = 12
'    oWS.Range("L:L").ColumnWidth = 12
'    oWS.Range("M:M").ColumnWidth = 12
'    oWS.Range("N:N").ColumnWidth = 12
'    oWS.Range("O:O").ColumnWidth = 12
'    oWS.Range("P:P").ColumnWidth = 12
'    oWS.Range("Q:Q").ColumnWidth = 12
'    oWS.Range("R:R").ColumnWidth = 12
'    oWS.Range("S:S").ColumnWidth = 12
'    oWS.Range("T:T").ColumnWidth = 12
'    oWS.Range("U:U").ColumnWidth = 12
'    oWS.Range("V:V").ColumnWidth = 12
'
'    oWS.Range("A1").Select                      '-- particular cell selection
'    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
'    oApp.Selection.Font.Size = 14            '-- enabled bold cell style
'    oApp.Selection.Font.Bold = True
'    'oApp.Columns("A:A").EntireColumn.AutoFit     '-- autofitted column
'
'    oWS.Range("A2").Select                      '-- particular cell selection
'    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
'    oApp.Selection.Font.Size = 11            '-- enabled bold cell style
'    oApp.Selection.Font.Bold = True
'
''    Range("C2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("C:C").EntireColumn.AutoFit     '-- autofitted column
''
''
''    Range("D2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("D:D").EntireColumn.AutoFit     '-- autofitted column
''
''    Range("E2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("E:E").EntireColumn.AutoFit     '-- autofitted column
''
''    Range("F2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("F:F").EntireColumn.AutoFit     '-- autofitted column
''
''    Range("G2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("G:G").EntireColumn.AutoFit     '-- autofitted column
''
''    Range("H2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("H:H").EntireColumn.AutoFit     '-- autofitted column
''
''    Range("I2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("I:I").EntireColumn.AutoFit     '-- autofitted column
''
''    Range("J2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("J:J").EntireColumn.AutoFit     '-- autofitted column
''
''    Range("K2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("K:K").EntireColumn.AutoFit     '-- autofitted column
''
''    Range("L2").Select                      '-- particular cell selection
''    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
''    Selection.Font.Bold = True              '-- enabled bold cell style
''    Columns("L:L").EntireColumn.AutoFit     '-- autofitted column
'
''    oWB.ActiveSheet.Font.Name = "Arial"
''    oApp.ActiveSheet.Font.Name = "Arial"
''    oWB.Font.Size = "11"
''    oWB.Font.Bold = True
'    oWS.Range("A" & 1).value = MDIMAIN.StatusBar.Panels(5).Text
'    oWS.Range("A" & 2).value = "STOCK REPORT"
'
'    'oApp.Selection.Font.Bold = False
'    oWS.Range("A" & 3).value = GRDSTOCK.TextMatrix(0, 0)
'    oWS.Range("B" & 3).value = GRDSTOCK.TextMatrix(0, 1)
'    oWS.Range("C" & 3).value = GRDSTOCK.TextMatrix(0, 2)
'    oWS.Range("D" & 3).value = GRDSTOCK.TextMatrix(0, 3)
'    On Error Resume Next
'    oWS.Range("E" & 3).value = GRDSTOCK.TextMatrix(0, 4)
'    oWS.Range("F" & 3).value = GRDSTOCK.TextMatrix(0, 5)
'    oWS.Range("G" & 3).value = GRDSTOCK.TextMatrix(0, 6)
'    oWS.Range("H" & 3).value = GRDSTOCK.TextMatrix(0, 7)
'    oWS.Range("I" & 3).value = GRDSTOCK.TextMatrix(0, 8)
'    oWS.Range("J" & 3).value = GRDSTOCK.TextMatrix(0, 9)
'    oWS.Range("K" & 3).value = GRDSTOCK.TextMatrix(0, 10)
'    oWS.Range("L" & 3).value = GRDSTOCK.TextMatrix(0, 11)
'    oWS.Range("M" & 3).value = GRDSTOCK.TextMatrix(0, 12)
'    oWS.Range("0" & 3).value = GRDSTOCK.TextMatrix(0, 13)
'    oWS.Range("O" & 3).value = GRDSTOCK.TextMatrix(0, 14)
'    oWS.Range("P" & 3).value = GRDSTOCK.TextMatrix(0, 15)
'    oWS.Range("Q" & 3).value = GRDSTOCK.TextMatrix(0, 16)
'    oWS.Range("R" & 3).value = GRDSTOCK.TextMatrix(0, 17)
'    oWS.Range("S" & 3).value = GRDSTOCK.TextMatrix(0, 18)
'    oWS.Range("T" & 3).value = GRDSTOCK.TextMatrix(0, 19)
'    oWS.Range("U" & 3).value = GRDSTOCK.TextMatrix(0, 20)
'    oWS.Range("V" & 3).value = GRDSTOCK.TextMatrix(0, 21)
'    oWS.Range("W" & 3).value = GRDSTOCK.TextMatrix(0, 22)
'    oWS.Range("X" & 3).value = GRDSTOCK.TextMatrix(0, 23)
'    oWS.Range("Y" & 3).value = GRDSTOCK.TextMatrix(0, 24)
'    oWS.Range("Z" & 3).value = GRDSTOCK.TextMatrix(0, 25)
'    On Error GoTo eRRHAND
'
'    i = 4
'    For N = 1 To GRDSTOCK.Rows - 1
'        oWS.Range("A" & i).value = GRDSTOCK.TextMatrix(N, 0)
'        oWS.Range("B" & i).value = GRDSTOCK.TextMatrix(N, 1)
'        oWS.Range("C" & i).value = GRDSTOCK.TextMatrix(N, 2)
'        oWS.Range("D" & i).value = GRDSTOCK.TextMatrix(N, 3)
'        oWS.Range("E" & i).value = GRDSTOCK.TextMatrix(N, 4)
'        oWS.Range("F" & i).value = GRDSTOCK.TextMatrix(N, 5)
'        oWS.Range("G" & i).value = GRDSTOCK.TextMatrix(N, 6)
'        oWS.Range("H" & i).value = GRDSTOCK.TextMatrix(N, 7)
'        oWS.Range("I" & i).value = GRDSTOCK.TextMatrix(N, 8)
'        oWS.Range("J" & i).value = GRDSTOCK.TextMatrix(N, 9)
'        oWS.Range("K" & i).value = GRDSTOCK.TextMatrix(N, 10)
'        oWS.Range("L" & i).value = GRDSTOCK.TextMatrix(N, 11)
'        oWS.Range("M" & i).value = GRDSTOCK.TextMatrix(N, 12)
'        oWS.Range("N" & i).value = GRDSTOCK.TextMatrix(N, 13)
'        oWS.Range("O" & i).value = GRDSTOCK.TextMatrix(N, 14)
'        oWS.Range("P" & i).value = GRDSTOCK.TextMatrix(N, 15)
'        oWS.Range("Q" & i).value = GRDSTOCK.TextMatrix(N, 16)
'        oWS.Range("R" & i).value = GRDSTOCK.TextMatrix(N, 17)
'        oWS.Range("S" & i).value = GRDSTOCK.TextMatrix(N, 18)
'        oWS.Range("T" & i).value = GRDSTOCK.TextMatrix(N, 19)
'        oWS.Range("U" & i).value = GRDSTOCK.TextMatrix(N, 20)
'        oWS.Range("V" & i).value = GRDSTOCK.TextMatrix(N, 21)
'        oWS.Range("W" & i).value = GRDSTOCK.TextMatrix(N, 22)
'        oWS.Range("X" & i).value = GRDSTOCK.TextMatrix(N, 23)
'        oWS.Range("Y" & i).value = GRDSTOCK.TextMatrix(N, 24)
'        oWS.Range("Z" & i).value = GRDSTOCK.TextMatrix(N, 25)
'        On Error GoTo eRRHAND
'        i = i + 1
'    Next N
'    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
'    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    oApp.Selection.HorizontalAlignment = xlRight
'    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
'    oApp.Selection.Font.Size = 13            '-- enabled bold cell style
'    oApp.Selection.Font.Bold = True
'
'
'SKIP:
'    oApp.Visible = True
'
'    If Sum_flag = True Then
'        'oWS.Columns("C:C").Select
'        oWS.Columns("C:C").NumberFormat = "0"
'        oWS.Columns("A:Z").EntireColumn.AutoFit
'    End If
'
''    Set oWB = Nothing
''    oApp.Quit
''    Set oApp = Nothing
''
'
'    Screen.MousePointer = vbNormal
'    Exit Sub
'eRRHAND:
'    'On Error Resume Next
'    Screen.MousePointer = vbNormal
'    Set oWB = Nothing
'    'oApp.Quit
'    'Set oApp = Nothing
'    MsgBox Err.Description
'End Sub

Private Sub CmdLoad_Click()
    Call Fillgrid
End Sub


Private Sub Command1_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If Trim(TxtHSN.text) = "" Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THESE HSN CODES", vbYesNo + vbDefaultButton2, "Assign HSN CODES....") = vbNo Then Exit Sub
'    For i = 1 To GRDSTOCK.Rows - 1
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (rststock.EOF And rststock.BOF) Then
'            rststock!REMARKS = Trim(TxtHSN.Text)
'            GRDSTOCK.TextMatrix(i, 14) = Trim(TxtHSN.Text)
'            rststock.Update
'        End If
'        rststock.Close
'        Set rststock = Nothing
'    Next i
    TxtHSN.text = ""
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Command2_Click()
'    If frmLogin.rs!Level <> "0" Then MsgBox "Permission Denied", vbOKOnly, "Price Analysis"
'    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Import Stock Items") = vbNo Then Exit Sub
'    If MsgBox("Sheet Name should be 'ITEMS' and First coloumn should be Item Code and Second coloumn should be Item name", vbYesNo, "Import Items") = vbNo Then Exit Sub
'    On Error GoTo eRRHAND
'    CommonDialog1.CancelError = True
'    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
'    CommonDialog1.Filter = "Excel Files (*.xls*)|*.xls*"
'    CommonDialog1.ShowOpen
'
'    Screen.MousePointer = vbHourglass
'    Set xlApp = New Excel.Application
'
'    'Set wb = xlApp.Workbooks.Open("PATH TO YOUR EXCEL FILE")
'    Set wb = xlApp.Workbooks.Open(CommonDialog1.FileName)
'
'    Set ws = wb.Worksheets("ITEMS") 'Specify your worksheet name
'    var = ws.Range("A1").value
'
''    db.Execute "dELETE FROM ITEMMAST"
''    db.Execute "dELETE FROM RTRXFILE"
'
'    Dim RSTITEMMAST As ADODB.Recordset
'    Dim RSTITEMTRX As ADODB.Recordset
'    Dim itemcode As String
'    Dim sl As Integer
'    Dim lastno As Integer
'    sl = 1
'    lastno = 1
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'ST'", db, adOpenStatic, adLockReadOnly
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        lastno = IIf(IsNull(RSTITEMMAST.Fields(0)), 1, RSTITEMMAST.Fields(0) + 1)
'    End If
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
'
'    For i = 2 To 7000
'        If Trim(ws.Range("A" & i).value) = "" Then Exit For
'
'        'If MsgBox("Item not exists!!! Do You want to add this item?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
'        Set RSTITEMTRX = New ADODB.Recordset
'        RSTITEMTRX.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME = '" & Trim(ws.Range("B" & i).value) & "' AND ITEM_CODE = '" & Trim(ws.Range("A" & i).value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (RSTITEMTRX.EOF And RSTITEMTRX.BOF) Then
'            MsgBox "Duplicate Name. Item " & Trim(ws.Range("B" & i).value) & " Skipped", vbOKOnly, "IMPORT ITEMS"
'            RSTITEMTRX.Close
'            Set RSTITEMTRX = Nothing
'            GoTo SKIP
'        End If
'        RSTITEMTRX.Close
'        Set RSTITEMTRX = Nothing
'
'        Set RSTITEMTRX = New ADODB.Recordset
'        RSTITEMTRX.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(ws.Range("A" & i).value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (RSTITEMTRX.EOF And RSTITEMTRX.BOF) Then
'            itemcode = ""
'            Set RSTITEMMAST = New ADODB.Recordset
'            RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
'            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'                If IsNull(RSTITEMMAST.Fields(0)) Then
'                    itemcode = 1
'                Else
'                    itemcode = Val(RSTITEMMAST.Fields(0)) + 1
'                End If
'            End If
'            RSTITEMMAST.Close
'            Set RSTITEMMAST = Nothing
'        Else
'            itemcode = Trim(ws.Range("A" & i).value)
'        End If
'
'        Set RSTITEMMAST = New ADODB.Recordset
'        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & itemcode & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        db.BeginTrans
'
'        RSTITEMMAST.AddNew
'        'RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
'        RSTITEMMAST!ITEM_CODE = itemcode
'        RSTITEMMAST!ITEM_NAME = Trim(ws.Range("B" & i).value)
'        RSTITEMMAST!Category = "GENERAL"
'        RSTITEMMAST!UNIT = 1
'        RSTITEMMAST!MANUFACTURER = "GENERAL"
'        RSTITEMMAST!DEAD_STOCK = "N"
'        RSTITEMMAST!REMARKS = Trim(ws.Range("I" & i).value)
'        RSTITEMMAST!REORDER_QTY = 1
'        RSTITEMMAST!PACK_TYPE = Trim(ws.Range("D" & i).value)
'        RSTITEMMAST!FULL_PACK = Trim(ws.Range("E" & i).value)
'        RSTITEMMAST!BIN_LOCATION = ""
'        RSTITEMMAST!MRP = Val(ws.Range("F" & i).value)
'        RSTITEMMAST!PTR = Val(ws.Range("G" & i).value)
'        RSTITEMMAST!CST = 0
'        RSTITEMMAST!OPEN_QTY = 0
'        RSTITEMMAST!OPEN_VAL = 0
'        RSTITEMMAST!RCPT_QTY = Val(ws.Range("C" & i).value)
'        RSTITEMMAST!RCPT_VAL = Val(ws.Range("C" & i).value) * Val(ws.Range("G" & i).value)
'        RSTITEMMAST!ISSUE_QTY = 0
'        RSTITEMMAST!ISSUE_VAL = 0
'        RSTITEMMAST!CLOSE_QTY = Val(ws.Range("C" & i).value)
'        RSTITEMMAST!CLOSE_VAL = Val(ws.Range("C" & i).value) * Val(ws.Range("G" & i).value)
'        RSTITEMMAST!DAM_QTY = 0
'        RSTITEMMAST!DAM_VAL = 0
'        RSTITEMMAST!DISC = 0
'        RSTITEMMAST!SALES_TAX = Val(ws.Range("H" & i).value)
'        RSTITEMMAST!ITEM_COST = Val(ws.Range("G" & i).value)
'        RSTITEMMAST!P_RETAIL = Val(ws.Range("J" & i).value)
'        RSTITEMMAST!P_WS = Val(ws.Range("K" & i).value)
'        RSTITEMMAST!P_VAN = Val(ws.Range("L" & i).value)
'        RSTITEMMAST!CRTN_PACK = 1
'        RSTITEMMAST!P_CRTN = Val(ws.Range("J" & i).value)
'        RSTITEMMAST!LOOSE_PACK = 1
'        RSTITEMMAST!CHECK_FLAG = "V"
'        RSTITEMMAST!UN_BILL = "N"
'        RSTITEMMAST.Update
'        db.CommitTrans
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
'
'        If Val(ws.Range("C" & i).value) > 0 Then
'            Set rststock = New ADODB.Recordset
'            rststock.Open "SELECT * FROM RTRXFILE ", db, adOpenStatic, adLockOptimistic, adCmdText
'            rststock.AddNew
'            rststock!TRX_TYPE = "ST"
'            rststock!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'            rststock!VCH_NO = lastno
'            rststock!line_no = sl
'            rststock!ITEM_CODE = itemcode
'            rststock!BAL_QTY = Val(ws.Range("C" & i).value)
'            rststock!QTY = Val(ws.Range("C" & i).value)
'            rststock!TRX_TOTAL = Val(ws.Range("C" & i).value) * Val(ws.Range("G" & i).value)
'            rststock!VCH_DATE = Format(Date, "dd/mm/yyyy")
'            rststock!ITEM_NAME = Trim(ws.Range("B" & i).value)
'            rststock!ITEM_COST = Val(ws.Range("G" & i).value)
'            rststock!LINE_DISC = 1
'            rststock!P_DISC = 0
'            rststock!MRP = Val(ws.Range("F" & i).value)
'            rststock!PTR = Val(ws.Range("G" & i).value)
'            rststock!SALES_PRICE = Val(ws.Range("J" & i).value)
'            rststock!P_RETAIL = Val(ws.Range("J" & i).value)
'            rststock!P_WS = Val(ws.Range("K" & i).value)
'            rststock!P_VAN = Val(ws.Range("L" & i).value)
'            rststock!P_CRTN = Val(ws.Range("J" & i).value)
'            rststock!P_LWS = Val(ws.Range("K" & i).value)
'            rststock!CRTN_PACK = 1
'            rststock!Category = "GENERAL"
'            rststock!GROSS_AMT = 0
'            rststock!COM_FLAG = "P"
'            rststock!COM_PER = 0
'            rststock!COM_AMT = 0
'            rststock!SALES_TAX = Val(ws.Range("H" & i).value)
'            rststock!LOOSE_PACK = 1
'            rststock!PACK_TYPE = Trim(ws.Range("D" & i).value)
'            rststock!WARRANTY = Null
'            rststock!WARRANTY_TYPE = Null
'            rststock!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
'            'rststock!VCH_DESC = "Received From " & DataList2.Text
'            rststock!REF_NO = ""
'            'rststock!ISSUE_QTY = 0
'            rststock!CST = 0
'            rststock!DISC_FLAG = "P"
'            rststock!SCHEME = 0
'            rststock!EXP_DATE = Null
'            rststock!FREE_QTY = 0
'            rststock!CREATE_DATE = Format(Date, "dd/mm/yyyy")
'            rststock!C_USER_ID = "SM"
'            rststock!CHECK_FLAG = "V"
'
'            'rststock!M_USER_ID = DataList2.BoundText
'            'rststock!PINV = Trim(TXTINVOICE.Text)
'            rststock.Update
'            rststock.Close
'            Set rststock = Nothing
'            sl = sl + 1
'        End If
'
'SKIP:
'    Next i
'    wb.Close
'
'    xlApp.Quit
'
'    Set ws = Nothing
'    Set wb = Nothing
'    Set xlApp = Nothing
'    Screen.MousePointer = vbNormal
'
'    Call CmdLoad_Click
'    MsgBox "Success", vbOKOnly
'    Exit Sub
'eRRHAND:
'    Screen.MousePointer = vbNormal
'    If Err.Number = 9 Then
'        MsgBox "NO SUCH FILE PRESENT!!", vbOKOnly, "IMPORT ITEMS"
'        wb.Close
'        xlApp.Quit
'        Set ws = Nothing
'        Set wb = Nothing
'        Set xlApp = Nothing
'    ElseIf Err.Number = 32755 Then
'
'    Else
'        MsgBox Err.Description
'    End If
End Sub

Private Sub Command3_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If Trim(TXTDEALER2.text) = "" Then
        MsgBox "Please enter a Category Name in the Text Box", vbOKOnly, "Price Analysis"
        TXTDEALER2.SetFocus
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THE CATEGORY TO ALL LISTED ITEMS", vbYesNo + vbDefaultButton2, "Assign CATEGORY....") = vbNo Then Exit Sub
'    For i = 1 To GRDSTOCK.Rows - 1
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (rststock.EOF And rststock.BOF) Then
'            rststock!Category = Trim(TXTDEALER2.Text)
'            GRDSTOCK.TextMatrix(i, 18) = Trim(TXTDEALER2.Text)
'            rststock.Update
'        End If
'        rststock.Close
'        Set rststock = Nothing
'    Next i
    TXTDEALER2.text = ""
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Command4_Click()
    Dim i As Long
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If Trim(TXTDEALER2.text) = "" Then
        MsgBox "Please enter a Company Name in the Text Box", vbOKOnly, "Price Analysis"
        TXTDEALER2.SetFocus
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THE COMPANY TO ALL LISTED ITEMS", vbYesNo + vbDefaultButton2, "Assign COMPANY....") = vbNo Then Exit Sub
'    For i = 1 To GRDSTOCK.Rows - 1
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (rststock.EOF And rststock.BOF) Then
'            rststock!MANUFACTURER = Trim(TXTDEALER2.Text)
'            GRDSTOCK.TextMatrix(i, 19) = Trim(TXTDEALER2.Text)
'            rststock.Update
'        End If
'        rststock.Close
'        Set rststock = Nothing
'    Next i
    TXTDEALER2.text = ""
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            tXTMEDICINE.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo ERRHAND
    
    frmloadflag = True
    db.Execute "Update itemmast set ITEM_NET_COST = ITEM_COST + ITEM_COST * SALES_TAX /100 "
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        If RSTCOMPANY!hide_ws = "Y" Then
            chkws.Value = 1
        Else
            chkws.Value = 0
        End If
        If RSTCOMPANY!hide_van = "Y" Then
            chkvp.Value = 1
        Else
            chkvp.Value = 0
        End If
        If RSTCOMPANY!hide_lwp = "Y" Then
            chklwp.Value = 1
        Else
            chklwp.Value = 0
        End If
        If RSTCOMPANY!hide_category = "Y" Then
            chkhidecat.Value = 1
        Else
            chkhidecat.Value = 0
        End If
        If RSTCOMPANY!hide_company = "Y" Then
            chkhidecomp.Value = 1
        Else
            chkhidecomp.Value = 0
        End If
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
           
    frmloadflag = False
    REPFLAG = True
    PHY_FLAG = True
    
'    If chkws.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 8) = ""
'        GRDSTOCK.ColWidth(8) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 8) = "WS"
'        GRDSTOCK.ColWidth(8) = 1000
'    End If
'    If chkvp.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 9) = ""
'        GRDSTOCK.ColWidth(9) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 9) = "VP"
'        GRDSTOCK.ColWidth(9) = 1000
'    End If
'    If chklwp.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 17) = ""
'        GRDSTOCK.ColWidth(17) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 17) = "L.W.Price"
'        GRDSTOCK.ColWidth(17) = 1000
'    End If
'    If chkhidecat.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 18) = ""
'        GRDSTOCK.ColWidth(18) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 18) = "Category"
'        GRDSTOCK.ColWidth(18) = 1000
'    End If
'    If chkhidecomp.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 19) = ""
'        GRDSTOCK.ColWidth(19) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 19) = "Company"
'        GRDSTOCK.ColWidth(19) = 1000
'    End If
'    If (frmLogin.rs!Level = "0" or frmLogin.rs!Level = "4") Then
'        GRDSTOCK.TextMatrix(0, 11) = "Per Rate"
'        GRDSTOCK.TextMatrix(0, 12) = "Net Cost"
'        GRDSTOCK.TextMatrix(0, 20) = "Profit%"
'        GRDSTOCK.ColWidth(11) = 1000
'        GRDSTOCK.ColWidth(12) = 900
'        GRDSTOCK.ColWidth(20) = 1000
'    Else
'        GRDSTOCK.TextMatrix(0, 11) = ""
'        GRDSTOCK.TextMatrix(0, 12) = ""
'        GRDSTOCK.TextMatrix(0, 20) = ""
'        GRDSTOCK.ColWidth(11) = 0
'        GRDSTOCK.ColWidth(12) = 0
'        GRDSTOCK.ColWidth(20) = 0
'    End If
    'Call Fillgrid
    DTFROM.Value = Format(Date, "DD/MM/YYYY")
    DTFROM.Value = Null
    Call Fillgrid
    'Me.Height = 8415
    'Me.Width = 6465
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If RSTREP.State = 1 Then RSTREP.Close
    If PHY_REC.State = 1 Then PHY_REC.Close
    If adoGrid.State = 1 Then
        adoGrid.Close
        Set adoGrid = Nothing
    End If
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub grdmsc_AfterColEdit(ByVal ColIndex As Integer)
'    Dim rststock, RSTITEMMAST As ADODB.Recordset
'
'    On Error GoTo eRRHAND
'
'        Select Case grdmsc.Col

'

'            Case 8  'WS
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT P_WS from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_WS) AND P_WS <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_WS = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 15)) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 15) = 1
'                        rststock!CRTN_PACK = 1
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 13)) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 13) = 1
'                        rststock!LOOSE_PACK = 1
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!P_WS = Val(TXTsample.Text)
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 15)) = 0 Then
'                        rststock!CRTN_PACK = 1
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 13)) = 0 Then
'                        rststock!LOOSE_PACK = 1
'                    End If
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 15  'CRTN_PACK
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!CRTN_PACK = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Val(TXTsample.Text)
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!CRTN_PACK = Val(TXTsample.Text)
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 16  'L. R. PRICE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT P_CRTN from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_CRTN) AND P_CRTN <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_CRTN = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                    If grdmsc.TextMatrix(grdmsc.Row, 15) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 15) = 1
'                        rststock!CRTN_PACK = 1
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!P_CRTN = Val(TXTsample.Text)
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 17  'L. W. PRICE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT P_LWS from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_LWS) AND P_LWS <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_LWS = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                    If grdmsc.TextMatrix(grdmsc.Row, 15) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 15) = 1
'                        rststock!CRTN_PACK = 1
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!P_LWS = Val(TXTsample.Text)
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 9  'VAN
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT P_VAN from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_VAN) AND P_VAN <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_VAN = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 15)) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 15) = 1
'                        rststock!CRTN_PACK = 1
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 13)) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 13) = 1
'                        rststock!LOOSE_PACK = 1
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!P_WS = Val(TXTsample.Text)
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 15)) = 0 Then
'                        rststock!CRTN_PACK = 1
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 13)) = 0 Then
'                        rststock!LOOSE_PACK = 1
'                    End If
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 18  'CATEGORY
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!Category = Trim(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Trim(TXTsample.Text)
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
''                Case 11  'LOC
''                    Set rststock = New ADODB.Recordset
''                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
''                    If Not (rststock.EOF And rststock.BOF) Then
''                        rststock!BIN_LOCATION = Trim(TXTsample.Text)
''                        grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Trim(TXTsample.Text)
''                        rststock.Update
''                    End If
''                    rststock.Close
''                    Set rststock = Nothing
''                    grdmsc.Enabled = True
''                    TXTsample.Visible = False
''                    grdmsc.SetFocus
'
'            Case 11  'COST
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!ITEM_COST = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    grdmsc.TextMatrix(grdmsc.Row, 12) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) + (Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 10)) / 100)
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 12)) <> 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 20) = Format(Round(((Val(grdmsc.TextMatrix(grdmsc.Row, 7)) - Val(grdmsc.TextMatrix(grdmsc.Row, 12))) * 100) / Val(grdmsc.TextMatrix(grdmsc.Row, 12)), 2), "0.00")
'                    Else
'                        grdmsc.TextMatrix(grdmsc.Row, 20) = 0
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.TextMatrix(grdmsc.Row, 25) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 3))
'                Call Toatal_value
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 25  'VALUE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 3)) <> 0 Then
'                        rststock!ITEM_COST = Round(Val(TXTsample.Text) / Val(grdmsc.TextMatrix(grdmsc.Row, 3)), 3)
'                        grdmsc.TextMatrix(grdmsc.Row, 11) = Format(Round(Val(TXTsample.Text) / Val(grdmsc.TextMatrix(grdmsc.Row, 3)), 3), "0.000")
'                        grdmsc.TextMatrix(grdmsc.Row, 12) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) + (Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 10)) / 100)
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 11)) <> 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 20) = Format(Round(((Val(grdmsc.TextMatrix(grdmsc.Row, 7)) - Val(grdmsc.TextMatrix(grdmsc.Row, 11))) * 100) / Val(grdmsc.TextMatrix(grdmsc.Row, 11)), 2), "0.00")
'                    Else
'                        grdmsc.TextMatrix(grdmsc.Row, 20) = 0
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.TextMatrix(grdmsc.Row, 25) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 3))
'                Call Toatal_value
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 6  'MRP
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!MRP = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 10  'TAX
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!SALES_TAX = Val(TXTsample.Text)
'                    rststock!CHECK_FLAG = "V"
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Val(TXTsample.Text)
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.TextMatrix(grdmsc.Row, 12) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) + (Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 10)) / 100)
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'                Call Toatal_value
'            Case 20  'Profit %
'                grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                grdmsc.TextMatrix(grdmsc.Row, 7) = Format(Round(((Val(grdmsc.TextMatrix(grdmsc.Row, 12)) * grdmsc.TextMatrix(grdmsc.Row, 13)) * Val(TXTsample.Text) / 100) + (Val(grdmsc.TextMatrix(grdmsc.Row, 12)) * grdmsc.TextMatrix(grdmsc.Row, 13)), 2), "0.000")
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_RETAIL = Val(grdmsc.TextMatrix(grdmsc.Row, 7))
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 21  'Cust Disc
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!CUST_DISC = Val(TXTsample.Text)
'                    rststock!DISC_AMT = 0
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    grdmsc.TextMatrix(grdmsc.Row, 22) = ""
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 22  'Cust Disc Amt
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!DISC_AMT = Val(TXTsample.Text)
'                    rststock!CUST_DISC = 0
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    grdmsc.TextMatrix(grdmsc.Row, 21) = ""
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 14  'HSN CODE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!REMARKS = Trim(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Trim(TXTsample.Text)
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 26  'CESS%
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!CESS_PER = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 27  'CESS RATE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!CESS_AMT = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 13  'UNIT
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!LOOSE_PACK = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
                
'        End Select
'        Exit Sub
'eRRHAND:
'    MsgBox Err.Description
End Sub

Private Sub grdmsc_AfterColUpdate(ByVal ColIndex As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    
    Select Case ColIndex
        Case 0  ' Item Code
            db.Execute "Update RTRXFILE set ITEM_CODE = '" & grdmsc.Columns(0) & "' where ITEM_CODE = '" & grdmsc.Tag & "' "
            db.Execute "Update TRXFILE set ITEM_CODE = '" & grdmsc.Columns(0) & "' where ITEM_CODE = '" & grdmsc.Tag & "' "
        Case 1  ' Item Name
            db.Execute "Update RTRXFILE set ITEM_NAME = '" & grdmsc.Columns(1) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' "
            db.Execute "Update TRXFILE set ITEM_NAME = '" & grdmsc.Columns(1) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' "
        Case 2
                Dim INWARD, OUTWARD, BAL_QTY As Double
                Dim TRXMAST As ADODB.Recordset
                Dim RSTITEMMAST As ADODB.Recordset

                Screen.MousePointer = vbHourglass
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & grdmsc.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                    INWARD = 0
                    OUTWARD = 0
                    BAL_QTY = 0

'                    Set TRXMAST = New ADODB.Recordset
'                    TRXMAST.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE = '" & Format(Date, "yyyy/mm/dd") & "' AND TRX_TYPE = 'ST' ", db, adOpenStatic, adLockOptimistic, adCmdText
'                    Set rststock = New ADODB.Recordset
'                    If TRXMAST.RecordCount > 0 Then
'                        rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' and TRX_TYPE <> 'ST'", db, adOpenStatic, adLockReadOnly
'                    Else
'                        rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
'                    End If
'                    TRXMAST.Close
'                    Set TRXMAST = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
                    Do Until rststock.EOF
                        INWARD = INWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                        INWARD = INWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                        BAL_QTY = BAL_QTY + IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing

                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='MI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockReadOnly
                    Do Until rststock.EOF
                        OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                        OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing

                    Dim BILL_NO, M_DATA As Double
                    Set TRXMAST = New ADODB.Recordset
                    TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'ST'", db, adOpenStatic, adLockReadOnly
                    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
                        BILL_NO = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
                    End If
                    TRXMAST.Close
                    Set TRXMAST = Nothing

                    If Not (grdmsc.Columns(2) - (Val(INWARD - OUTWARD)) = 0) Then
                        Set rststock = New ADODB.Recordset
                        'rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE = '" & Format(Date, "yyyy/mm/dd") & "' AND TRX_TYPE = 'ST' ", db, adOpenStatic, adLockOptimistic, adCmdText
                        rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                        'If (rststock.EOF And rststock.BOF) Then
                            rststock.AddNew
                            rststock!TRX_TYPE = "ST"
                            rststock!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                            rststock!VCH_NO = BILL_NO
                            rststock!LINE_NO = 1
                            rststock!ITEM_CODE = RSTITEMMAST!ITEM_CODE
                        'End If
                        rststock!BAL_QTY = grdmsc.Columns(2) - (Val(BAL_QTY))
                        rststock!QTY = grdmsc.Columns(2) - (Val(INWARD - OUTWARD))
                        rststock!TRX_TOTAL = 0
                        rststock!VCH_DATE = Format(DTFROM.Value, "dd/mm/yyyy")
                        rststock!ITEM_NAME = grdmsc.Columns(1)
                        rststock!ITEM_COST = grdmsc.Columns(9)
                        rststock!LINE_DISC = 1
                        rststock!P_DISC = 0
                        rststock!MRP = grdmsc.Columns(4)
                        rststock!PTR = grdmsc.Columns(9)
                        rststock!SALES_PRICE = grdmsc.Columns(5)
                        rststock!P_RETAIL = Val(grdmsc.Columns(5))
                        rststock!P_WS = Val(grdmsc.Columns(6))
                        rststock!P_VAN = Val(grdmsc.Columns(7))
                        rststock!P_CRTN = Val(grdmsc.Columns(10))
                        rststock!P_LWS = Val(grdmsc.Columns(11))
                        rststock!CRTN_PACK = IIf(IsNull(RSTITEMMAST!CRTN_PACK) Or RSTITEMMAST!CRTN_PACK = 0, 1, RSTITEMMAST!CRTN_PACK)
                        rststock!Category = grdmsc.Columns(12)
                        rststock!gross_amt = 0
                        rststock!COM_FLAG = "P"
                        rststock!COM_PER = 0
                        rststock!COM_AMT = 0
                        rststock!SALES_TAX = Val(grdmsc.Columns(8))
                        rststock!LOOSE_PACK = RSTITEMMAST!LOOSE_PACK
                        rststock!PACK_TYPE = grdmsc.Columns(3)
                        rststock!WARRANTY = Null
                        rststock!WARRANTY_TYPE = Null
                        rststock!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
                        'rststock!VCH_DESC = "Received From " & DataList2.Text
                        rststock!REF_NO = ""
                        'rststock!ISSUE_QTY = 0
                        rststock!CST = 0
                        rststock!DISC_FLAG = "P"
                        rststock!SCHEME = 0
                        'rststock!EXP_DATE = Null
                        rststock!FREE_QTY = 0
                        rststock!CREATE_DATE = Format(Date, "dd/mm/yyyy")
                        rststock!C_USER_ID = "SM"
                        rststock!check_flag = "V"

                        'rststock!M_USER_ID = DataList2.BoundText
                        'rststock!PINV = Trim(TXTINVOICE.Text)
                        rststock.Update
                        rststock.Close
                        Set rststock = Nothing

                        'RSTITEMMAST!CLOSE_QTY = grdmsc.Columns(2)
                        'RSTITEMMAST!RCPT_QTY = INWARD + grdmsc.Columns(2)
                        'RSTITEMMAST!ISSUE_QTY = OUTWARD
                        'RSTITEMMAST.Update
                    End If
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
                End If
                Screen.MousePointer = vbNormal
            Case 5
                db.Execute "Update RTRXFILE set P_RETAIL = '" & grdmsc.Columns(5) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 "
            Case 6
                db.Execute "Update RTRXFILE set P_WS = '" & grdmsc.Columns(6) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 "
            Case 7
                db.Execute "Update RTRXFILE set P_VAN = '" & grdmsc.Columns(7) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 "
               
                

               

            
    End Select
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub grdmsc_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    Select Case ColIndex
        Case 0
            grdmsc.Tag = OldValue
        Case 2
            If IsNull(DTFROM.Value) Then
                MsgBox "Select the Date for Opening Qty", vbOKOnly, "Price Analysis"
                Cancel = 1
                Exit Sub
            End If
            If (DTFROM.Value) > Date Then
                MsgBox "The date could not be greater than Today", vbOKOnly, "Price Analysis"
                Cancel = 1
                Exit Sub
            End If
            If Val(grdmsc.Columns(9)) = 0 Then
                MsgBox "Please enter the cost", vbOKOnly, "Price Analysis"
                Cancel = 1
                Exit Sub
            End If
        
        Case 5
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT DISTINCT P_RETAIL from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(P_RETAIL) AND P_RETAIL <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
            If rststock.RecordCount > 1 Then
                If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                    rststock.Close
                    Set rststock = Nothing
                    Cancel = 1
                    Exit Sub
                End If
            End If
            rststock.Close
            Set rststock = Nothing

            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
            If rststock.RecordCount > 1 Then
                If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                    rststock.Close
                    Set rststock = Nothing
                    Cancel = 1
                    Exit Sub
                End If
            End If
            rststock.Close
            Set rststock = Nothing

            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If rststock.RecordCount > 1 Then
                If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                    rststock.Close
                    Set rststock = Nothing
                    Cancel = 1
                    Exit Sub
                End If
            End If
            rststock.Close
            Set rststock = Nothing
    
        Case 6
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT P_WS from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(P_WS) AND P_WS <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
        
        Case 7
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT P_VAN from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(P_VAN) AND P_VAN <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    End Select
    Exit Sub
ERRHAND:
    Cancel = 1
    MsgBox err.Description
    
End Sub

Private Sub OptAll_Click()
    tXTMEDICINE.SetFocus
End Sub

Private Sub TxtHSNCODE_Change()
    SEARCHFLAG = 1
    Call fillitemlist
End Sub

Private Sub TXTITEMCODE_Change()
    SEARCHFLAG = 2
    Call fillitemlist
End Sub

Private Sub tXTMEDICINE_Change()
    SEARCHFLAG = 1
    Call fillitemlist
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.text)
    'Call Fillgrid
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
            TxtCode.SetFocus
        Case vbKeyEscape
            Call CmdExit_Click
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


Private Sub TXTCODE_Change()
    SEARCHFLAG = 1
    Call fillitemlist
    Exit Sub
    On Error GoTo ERRHAND
    Call Fillgrid
    If REPFLAG = True Then
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
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

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
    'Call Fillgrid
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtItemcode.SetFocus
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


Private Function Toatal_value()
    Dim Stk_Val As Double
    Dim i As Long
    lblpvalue.Caption = ""
    lblnetvalue.Caption = ""
'    For i = 1 To GRDSTOCK.Rows - 1
'        lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 25)), "0.00")
''        If (Val(GRDSTOCK.TextMatrix(i, 12)) * Val(GRDSTOCK.TextMatrix(i, 3))) <> Val(GRDSTOCK.TextMatrix(i, 25)) Then
''            MsgBox ""
''        End If
'        lblnetvalue.Caption = Format(Val(lblnetvalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 12)) * Val(GRDSTOCK.TextMatrix(i, 3))), "0.00")
'    Next i
End Function


Private Sub TXTDEALER2_Change()
    
    On Error GoTo ERRHAND
    If FLAGCHANGE2.Caption <> "1" Then
        If chkcategory.Value = 1 Then
            If PHY_FLAG = True Then
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            End If
            If (PHY_REC.EOF And PHY_REC.BOF) Then
                LBLDEALER2.Caption = ""
            Else
                LBLDEALER2.Caption = PHY_REC!Category
            End If
            Set Me.DataList1.RowSource = PHY_REC
            DataList1.ListField = "CATEGORY"
            DataList1.BoundColumn = "CATEGORY"
        Else
            If PHY_FLAG = True Then
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            End If
            If (PHY_REC.EOF And PHY_REC.BOF) Then
                LBLDEALER2.Caption = ""
            Else
                LBLDEALER2.Caption = PHY_REC!MANUFACTURER
            End If
            Set Me.DataList1.RowSource = PHY_REC
            DataList1.ListField = "MANUFACTURER"
            DataList1.BoundColumn = "MANUFACTURER"

        End If
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub


Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.text)
    'CHKCATEGORY2.value = 1
End Sub

Private Sub TXTDEALER2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList1.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList1.SetFocus
    End Select

End Sub

Private Sub TXTDEALER2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_Click()
        
    TXTDEALER2.text = DataList1.text
    LBLDEALER2.Caption = TXTDEALER2.text
    Call Fillgrid
    tXTMEDICINE.SetFocus
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER2.text) = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Category From List", vbOKOnly, "Category List..."
                DataList1.SetFocus
                Exit Sub
            End If
        Case vbKeyEscape
            TXTDEALER2.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_GotFocus()
    FLAGCHANGE2.Caption = 1
    TXTDEALER2.text = LBLDEALER2.Caption
    DataList1.text = TXTDEALER2.text
    Call DataList1_Click
    'CHKCATEGORY2.value = 1
End Sub

Private Sub DataList1_LostFocus()
     FLAGCHANGE2.Caption = ""
End Sub

Private Sub TxtTax_Change()
    On Error GoTo ERRHAND
    SEARCHFLAG = 1
    Call Fillgrid
    If REPFLAG = True Then
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND SALES_TAX = " & Val(Me.TxtTax.text) & " AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND SALES_TAX = " & Val(Me.TxtTax.text) & " ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND SALES_TAX = " & Val(Me.TxtTax.text) & " AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND SALES_TAX = " & Val(Me.TxtTax.text) & " ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND SALES_TAX = " & Val(Me.TxtTax.text) & " AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND SALES_TAX = " & Val(Me.TxtTax.text) & " ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND SALES_TAX = " & Val(Me.TxtTax.text) & " AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND SALES_TAX = " & Val(Me.TxtTax.text) & " ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
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

Private Sub TxtItemcode_GotFocus()
    TxtItemcode.SelStart = 0
    TxtItemcode.SelLength = Len(TxtItemcode.text)
    'Call Fillgrid
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtTax.SetFocus
        Case vbKeyEscape
            TxtCode.SetFocus
    End Select

End Sub

Private Sub TxtItemcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TxtTax_GotFocus()
    TxtTax.SelStart = 0
    TxtTax.SelLength = Len(TxtTax.text)
    'Call Fillgrid
End Sub

Private Sub TxtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtHSNCODE.SetFocus
        Case vbKeyEscape
            TxtItemcode.SetFocus
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

Private Sub TxtHSNCODE_GotFocus()
    TxtHSNCODE.SelStart = 0
    TxtHSNCODE.SelLength = Len(TxtHSNCODE.text)
    'Call Fillgrid
End Sub

Private Sub TxtHSNCODE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'Call CmdLoad_Click
        Case vbKeyEscape
            TxtTax.SetFocus
    End Select

End Sub

Private Sub TxtHSNCODE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Function Fillgrid()
    On Error GoTo ERRHAND
    
    db.Execute "Update ITEMMAST Set CLOSE_QTY =0 where isnull(CLOSE_QTY)"
    db.Execute "Update ITEMMAST Set ITEM_COST =0 where isnull(ITEM_COST)"
    db.Execute "Update ITEMMAST Set REMARKS ='' where isnull(REMARKS)"
    
    Dim RSTTRXFILE As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Set grdmsc.DataSource = Nothing
    Set adoGrid = New ADODB.Recordset
    With adoGrid
        .CursorLocation = adUseClient
        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
            If OptStock.Value = True Then
                If SEARCHFLAG = 2 Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ", db, adOpenForwardOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                Else
                    .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ", db, adOpenForwardOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
            ElseIf OptPC.Value = True Then
                If SEARCHFLAG = 2 Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ", db, adOpenForwardOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                Else
                    .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ", db, adOpenForwardOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
            Else
                If SEARCHFLAG = 2 Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ", db, adOpenForwardOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                Else
                    .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ", db, adOpenForwardOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                End If
            End If
        Else
            If CHKCATEGORY2.Value = 1 Then
                If OptStock.Value = True Then
                    If SEARCHFLAG = 2 Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    End If
                ElseIf OptPC.Value = True Then
                    If SEARCHFLAG = 2 Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    End If
                Else
                    If SEARCHFLAG = 2 Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    End If
                End If
            Else
                If OptStock.Value = True Then
                    If SEARCHFLAG = 2 Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    End If
                ElseIf OptPC.Value = True Then
                    If SEARCHFLAG = 2 Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    End If
                Else
                    If SEARCHFLAG = 2 Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_COST, ITEM_NET_COST, CUST_DISC, P_CRTN, P_LWS, CATEGORY, MANUFACTURER, CESS_PER, CESS_AMT FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT SUM(ITEM_COST * CLOSE_QTY) FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ", db, adOpenForwardOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            lblpvalue.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                    End If
                End If
    
            End If
        End If
    End With
    Set grdmsc.DataSource = adoGrid
    
    grdmsc.Columns(0).Caption = "ITEM CODE"
    grdmsc.Columns(1).Caption = "ITEM NAME"
    grdmsc.Columns(2).Caption = "QTY"
    grdmsc.Columns(3).Caption = "UOM"
    grdmsc.Columns(4).Caption = "MRP"
    grdmsc.Columns(5).Caption = "R. PRICE"
    grdmsc.Columns(6).Caption = "W. PRICE"
    grdmsc.Columns(7).Caption = "V. PRICE"
    grdmsc.Columns(8).Caption = "GST%"
    grdmsc.Columns(9).Caption = "COST"
    grdmsc.Columns(10).Caption = "NET COST"
    grdmsc.Columns(11).Caption = "CUST DISC"
    grdmsc.Columns(12).Caption = "LR. PRICE"
    grdmsc.Columns(13).Caption = "LW. PRICE"
    grdmsc.Columns(14).Caption = "CATEGORY"
    grdmsc.Columns(15).Caption = "COMPANY"
    grdmsc.Columns(16).Caption = "CESS%"
    grdmsc.Columns(17).Caption = "ADDL CESS"
        
    grdmsc.Columns(0).Width = 1200
    grdmsc.Columns(1).Width = 6000
    grdmsc.Columns(2).Width = 1200
    grdmsc.Columns(3).Width = 900
    grdmsc.Columns(4).Width = 1000
    grdmsc.Columns(5).Width = 1000
    grdmsc.Columns(6).Width = 1000
    grdmsc.Columns(7).Width = 1000
    grdmsc.Columns(8).Width = 900
    grdmsc.Columns(9).Width = 1000
    grdmsc.Columns(10).Width = 1000
    grdmsc.Columns(11).Width = 1000
    grdmsc.Columns(12).Width = 1100
    grdmsc.Columns(13).Width = 1100
    grdmsc.Columns(14).Width = 1500
    grdmsc.Columns(15).Width = 1500
    grdmsc.Columns(16).Width = 1000
    grdmsc.Columns(17).Width = 1000
    lblpvalue.Caption = Format(Round(Val(lblpvalue.Caption), 2), "0.00")
    Screen.MousePointer = vbNormal
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Function

Private Function fillitemlist()
    On Error GoTo ERRHAND
    Call Fillgrid
    If REPFLAG = True Then
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.text & "%' AND ITEM_CODE Like '" & Me.TxtItemcode.text & "' AND REMARKS Like '%" & Me.TxtHSNCODE.text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    Exit Function
'RSTREP.Close
'TMPFLAG = False
ERRHAND:
    MsgBox err.Description
End Function
