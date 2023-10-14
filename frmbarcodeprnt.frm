VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmBarcodePrint 
   BackColor       =   &H00E8DFEC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price Analysis"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   20400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmbarcodeprnt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   20400
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
      Left            =   18165
      TabIndex        =   54
      Top             =   2055
      Width           =   1110
   End
   Begin VB.TextBox txtbarcode 
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
      Left            =   6555
      TabIndex        =   3
      Top             =   555
      Width           =   3885
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
      Left            =   5355
      TabIndex        =   2
      Top             =   270
      Width           =   1170
   End
   Begin VB.Frame frmunbill 
      BackColor       =   &H00FFC0C0&
      Height          =   630
      Left            =   16500
      TabIndex        =   49
      Top             =   -15
      Visible         =   0   'False
      Width           =   2565
      Begin VB.CheckBox Chkunbill 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Show Un Bill items"
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
         Left            =   45
         TabIndex        =   51
         Top             =   120
         Width           =   1875
      End
      Begin VB.CheckBox chkonlyunbill 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Show Only Un Bill Items"
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
         Left            =   45
         TabIndex        =   50
         Top             =   375
         Visible         =   0   'False
         Width           =   2385
      End
   End
   Begin VB.ComboBox Cmbbarcode 
      Height          =   315
      ItemData        =   "frmbarcodeprnt.frx":000C
      Left            =   14250
      List            =   "frmbarcodeprnt.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   1980
      Width           =   3870
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
      Height          =   510
      Left            =   15360
      TabIndex        =   43
      Top             =   1230
      Width           =   1155
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
      Height          =   510
      Left            =   14235
      TabIndex        =   42
      Top             =   1230
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Assign HSN to all"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   14220
      TabIndex        =   40
      Top             =   600
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
      TabIndex        =   39
      Top             =   630
      Width           =   1155
   End
   Begin VB.Frame Frame 
      Height          =   2190
      Left            =   1920
      TabIndex        =   30
      Top             =   3900
      Visible         =   0   'False
      Width           =   3945
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   2640
         TabIndex        =   37
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   1335
         TabIndex        =   36
         Top             =   1665
         Width           =   1200
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Commission Type"
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
         Height          =   1470
         Left            =   75
         TabIndex        =   31
         Top             =   150
         Width           =   3780
         Begin VB.OptionButton OptAmt 
            BackColor       =   &H00FFC0C0&
            Caption         =   "&Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1890
            TabIndex        =   34
            Top             =   285
            Width           =   1680
         End
         Begin VB.OptionButton OptPercent 
            BackColor       =   &H00FFC0C0&
            Caption         =   "&Percentage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   165
            TabIndex        =   33
            Top             =   285
            Width           =   1680
         End
         Begin VB.TextBox TxtComper 
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
            Left            =   1470
            TabIndex        =   32
            Top             =   765
            Width           =   1650
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            Caption         =   "Commission"
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
            Index           =   24
            Left            =   195
            TabIndex        =   35
            Top             =   765
            Width           =   1260
         End
      End
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
      Height          =   315
      Left            =   12480
      TabIndex        =   28
      Top             =   0
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
      Left            =   10545
      TabIndex        =   24
      Top             =   330
      Width           =   3360
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
      Height          =   315
      Left            =   10740
      TabIndex        =   23
      Top             =   -15
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
      TabIndex        =   22
      Top             =   75
      Width           =   1155
   End
   Begin VB.CommandButton CmdDisc 
      Caption         =   "&Assign Tax to all"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   14220
      TabIndex        =   21
      Top             =   30
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   19320
      Top             =   1935
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   450
      Left            =   9060
      TabIndex        =   15
      Top             =   945
      Width           =   1380
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
      Left            =   3765
      TabIndex        =   1
      Top             =   270
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   990
      Left            =   6555
      TabIndex        =   6
      Top             =   915
      Width           =   2460
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
         Height          =   240
         Left            =   60
         TabIndex        =   38
         Top             =   690
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
         Height          =   240
         Left            =   75
         TabIndex        =   8
         Top             =   135
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
         Height          =   240
         Left            =   60
         TabIndex        =   7
         Top             =   405
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
      Width           =   3705
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
      Height          =   435
      Left            =   9060
      TabIndex        =   5
      Top             =   1440
      Width           =   1395
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   45
      TabIndex        =   4
      Top             =   630
      Width           =   6480
      _ExtentX        =   11430
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
      Height          =   6210
      Left            =   45
      TabIndex        =   9
      Top             =   2295
      Width           =   20370
      Begin MSMask.MaskEdBox TXTEXPIRY 
         Height          =   345
         Left            =   10035
         TabIndex        =   46
         Top             =   3405
         Visible         =   0   'False
         Width           =   1590
         _ExtentX        =   2805
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
      Begin MSDataListLib.DataCombo Cmbcategory 
         Height          =   360
         Left            =   9645
         TabIndex        =   41
         Top             =   2040
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   635
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo CMBMFGR 
         Height          =   360
         Left            =   6120
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   635
         _Version        =   393216
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
         Left            =   210
         TabIndex        =   12
         Top             =   870
         Visible         =   0   'False
         Width           =   1350
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
         ItemData        =   "frmbarcodeprnt.frx":0010
         Left            =   2385
         List            =   "frmbarcodeprnt.frx":0062
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   795
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
         Height          =   6075
         Left            =   30
         TabIndex        =   10
         Top             =   45
         Width           =   20310
         _ExtentX        =   35825
         _ExtentY        =   10716
         _Version        =   393216
         Rows            =   1
         Cols            =   29
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   400
         BackColor       =   15985374
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DTFROM 
      Height          =   390
      Left            =   8595
      TabIndex        =   19
      Top             =   15
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   0
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   255
      CheckBox        =   -1  'True
      Format          =   113836033
      CurrentDate     =   40498
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   780
      Left            =   10545
      TabIndex        =   25
      Top             =   675
      Width           =   3360
      _ExtentX        =   5927
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
      BackStyle       =   0  'Transparent
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
      Height          =   300
      Index           =   1
      Left            =   5355
      TabIndex        =   53
      Top             =   30
      Width           =   1050
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
      Height          =   300
      Index           =   5
      Left            =   6555
      TabIndex        =   52
      Top             =   315
      Width           =   1485
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
      Left            =   14250
      TabIndex        =   48
      Top             =   1740
      Width           =   1620
   End
   Begin VB.Label lblnetvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   12090
      TabIndex        =   45
      Top             =   1905
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Net Value"
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
      Left            =   10545
      TabIndex        =   44
      Top             =   1920
      Width           =   1545
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
      Left            =   6555
      TabIndex        =   29
      Top             =   1905
      Width           =   3900
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   0
      TabIndex        =   27
      Top             =   810
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   0
      TabIndex        =   26
      Top             =   450
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Op Stock Entry Date"
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
      Height          =   270
      Index           =   3
      Left            =   6600
      TabIndex        =   20
      Top             =   60
      Width           =   1980
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
      Left            =   10545
      TabIndex        =   18
      Top             =   1515
      Width           =   1500
   End
   Begin VB.Label lblpvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   12075
      TabIndex        =   17
      Top             =   1470
      Width           =   2100
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
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
      Left            =   75
      TabIndex        =   14
      Top             =   30
      Width           =   4380
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
      TabIndex        =   13
      Top             =   660
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "FrmBarcodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim REPFLAG As Boolean 'REP
Dim MFG_REC As New ADODB.Recordset
Dim CAT_REC As New ADODB.Recordset
Dim RSTREP As New ADODB.Recordset
Dim PHY_FLAG As Boolean 'REP
Dim PHY_REC As New ADODB.Recordset
Dim BARPRINTER As String

Private Sub Cmbbarcode_Click()
    BARPRINTER = Cmbbarcode.ListIndex
End Sub

Private Sub CHKCATEGORY_Click()
    CHKCATEGORY2.Value = 0
End Sub

Private Sub CHKCATEGORY2_Click()
    chkcategory.Value = 0
End Sub

Private Sub Cmbcategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            Select Case GRDSTOCK.Col
                Case 18  'CATEGORY
                    If Cmbcategory.Text = "" Then
                        MsgBox "Please select Category from the List", vbOKOnly, "EzBiz"
                        Exit Sub
                    End If
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!Category = Cmbcategory.Text
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Cmbcategory.Text
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT CATEGORY FROM CATEGORY where CATEGORY = '" & Cmbcategory.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If (rststock.EOF And rststock.BOF) Then
                        rststock.AddNew
                        rststock!Category = Cmbcategory.Text
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set Cmbcategory.DataSource = Nothing
                    If CAT_REC.State = 1 Then
                        CAT_REC.Close
                        CAT_REC.Open "SELECT DISTINCT CATEGORY FROM CATEGORY ORDER BY CATEGORY", db, adOpenForwardOnly
                    Else
                        CAT_REC.Open "SELECT DISTINCT CATEGORY FROM CATEGORY ORDER BY CATEGORY", db, adOpenForwardOnly
                    End If
                    Set Cmbcategory.RowSource = CAT_REC
                    Cmbcategory.ListField = "CATEGORY"
    
                    GRDSTOCK.Enabled = True
                    Cmbcategory.Visible = False
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            Cmbcategory.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub CMBMFGR_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            Select Case GRDSTOCK.Col
                Case 19  'MFGR
                    If CMBMFGR.Text = "" Then
                        MsgBox "Please select Company from the List", vbOKOnly, "EzBiz"
                        Exit Sub
                    End If
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MANUFACTURER = CMBMFGR.Text
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = CMBMFGR.Text
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT MANUFACTURER FROM MANUFACT where MANUFACTURER = '" & CMBMFGR.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If (rststock.EOF And rststock.BOF) Then
                        rststock.AddNew
                        rststock!MANUFACTURER = CMBMFGR.Text
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set CMBMFGR.DataSource = Nothing
                    If MFG_REC.State = 1 Then
                        MFG_REC.Close
                        MFG_REC.Open "SELECT DISTINCT MANUFACTURER FROM ITEMMAST ORDER BY MANUFACTURER", db, adOpenForwardOnly
                    Else
                        MFG_REC.Open "SELECT DISTINCT MANUFACTURER FROM ITEMMAST ORDER BY MANUFACTURER", db, adOpenForwardOnly
                    End If
                    Set CMBMFGR.RowSource = MFG_REC
                    CMBMFGR.ListField = "MANUFACTURER"
    
                    GRDSTOCK.Enabled = True
                    CMBMFGR.Visible = False
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            CMBMFGR.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            Select Case GRDSTOCK.Col
                Case 5  'pack
                    If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
                    db.Execute "Update ITEMMAST set PACK_TYPE = '" & CmbPack.Text & "',  UQC = '" & CmbPack.Text & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' "
                    db.Execute "Update RTRXFILE set PACK_TYPE = '" & CmbPack.Text & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0"
                    GRDSTOCK.Enabled = True
                    CmbPack.Visible = False
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = CmbPack.Text
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            CmbPack.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub CmdDelete_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If MDIMAIN.StatusBar.Panels(9).Text = "Y" Then Exit Sub
    Dim rststock As ADODB.Recordset
    
    If GRDSTOCK.rows <= 1 Then Exit Sub
    On Error GoTo eRRhAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from RTRXFILE where RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULASUB where FOR_NAME = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " Since Transactions is Available in Formula", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULAMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " Since Transactions is Available in Formula", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT BAL_QTY from RTRXFILE where RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
'    Do Until rststock.EOF
'        i = i + rststock!BAL_QTY
'        rststock.MoveNext
'    Loop
'    rststock.Close
'    Set rststock = Nothing
    
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " ****", vbYesNo + vbDefaultButton2, "DELETING ITEM....") = vbNo Then Exit Sub
    'db.Execute ("DELETE from RTRXFILE where RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'")
    db.Execute ("DELETE from PRODLINK where PRODLINK.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'")
    db.Execute ("DELETE from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'")
    
    'tXTMEDICINE.Tag = tXTMEDICINE.Text
    'tXTMEDICINE.Text = ""
    'tXTMEDICINE.Text = tXTMEDICINE.Tag
    'TXTQTY.Text = ""
'    Dim n As Long
'    n = GRDSTOCK.Row
'    MsgBox "ITEM " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & "DELETED SUCCESSFULLY", vbInformation, "DELETING ITEM...."
'    For i = GRDSTOCK.Row - 1 To GRDSTOCK.Rows - 2
'        GRDSTOCK.TextMatrix(n, 0) = i
'        GRDSTOCK.TextMatrix(n, 1) = GRDSTOCK.TextMatrix(i + 1, 1)
'        GRDSTOCK.TextMatrix(n, 2) = GRDSTOCK.TextMatrix(i + 1, 2)
'        GRDSTOCK.TextMatrix(n, 3) = GRDSTOCK.TextMatrix(i + 1, 3)
'        GRDSTOCK.TextMatrix(n, 4) = GRDSTOCK.TextMatrix(i + 1, 4)
'        GRDSTOCK.TextMatrix(n, 6) = GRDSTOCK.TextMatrix(i + 1, 6)
'        GRDSTOCK.TextMatrix(n, 5) = GRDSTOCK.TextMatrix(i + 1, 5)
'        GRDSTOCK.TextMatrix(n, 7) = GRDSTOCK.TextMatrix(i + 1, 7)
'        GRDSTOCK.TextMatrix(n, 8) = GRDSTOCK.TextMatrix(i + 1, 8)
'        GRDSTOCK.TextMatrix(n, 9) = GRDSTOCK.TextMatrix(i + 1, 9)
'        GRDSTOCK.TextMatrix(n, 10) = GRDSTOCK.TextMatrix(i + 1, 10)
'        GRDSTOCK.TextMatrix(n, 11) = GRDSTOCK.TextMatrix(i + 1, 11)
'        GRDSTOCK.TextMatrix(n, 12) = GRDSTOCK.TextMatrix(i + 1, 12)
'        GRDSTOCK.TextMatrix(n, 13) = GRDSTOCK.TextMatrix(i + 1, 13)
'        GRDSTOCK.TextMatrix(n, 14) = GRDSTOCK.TextMatrix(i + 1, 14)
'        GRDSTOCK.TextMatrix(n, 15) = GRDSTOCK.TextMatrix(i + 1, 15)
'        GRDSTOCK.TextMatrix(n, 16) = GRDSTOCK.TextMatrix(i + 1, 16)
'        GRDSTOCK.TextMatrix(n, 17) = GRDSTOCK.TextMatrix(i + 1, 17)
'        GRDSTOCK.TextMatrix(n, 18) = GRDSTOCK.TextMatrix(i + 1, 18)
'        GRDSTOCK.TextMatrix(n, 19) = GRDSTOCK.TextMatrix(i + 1, 19)
'        GRDSTOCK.TextMatrix(n, 20) = GRDSTOCK.TextMatrix(i + 1, 20)
'        GRDSTOCK.TextMatrix(n, 21) = GRDSTOCK.TextMatrix(i + 1, 21)
'        GRDSTOCK.TextMatrix(n, 22) = GRDSTOCK.TextMatrix(i + 1, 22)
'        GRDSTOCK.TextMatrix(n, 23) = GRDSTOCK.TextMatrix(i + 1, 23)
'        GRDSTOCK.TextMatrix(n, 24) = GRDSTOCK.TextMatrix(i + 1, 24)
'        GRDSTOCK.TextMatrix(n, 25) = GRDSTOCK.TextMatrix(i + 1, 25)
'        GRDSTOCK.TextMatrix(n, 26) = GRDSTOCK.TextMatrix(i + 1, 26)
'        n = n - 1
'    Next i
'    GRDSTOCK.Rows = GRDSTOCK.Rows - 1
    Exit Sub
   
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub CmdDisc_Click()
    Dim i As Integer
    Dim rststock As ADODB.Recordset
    On Error GoTo eRRhAND
    If Trim(TXTDISC.Text) = "" Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THESE TAX", vbYesNo + vbDefaultButton2, "Assign TAX....") = vbNo Then Exit Sub
    For i = 1 To GRDSTOCK.rows - 1
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (rststock.EOF And rststock.BOF) Then
            rststock!SALES_TAX = Val(TXTDISC.Text)
            rststock!check_flag = "V"
            'rststock!P_RETAIL = rststock!MRP
            GRDSTOCK.TextMatrix(i, 10) = Val(TXTDISC.Text)
            rststock.Update
        End If
        rststock.Close
        Set rststock = Nothing
        
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' WHERE BAL_QTY >0", db, adOpenStatic, adLockOptimistic, adCmdText
'        Do Until rststock.EOF
'            rststock!CUST_DISC = Val(TxtDisc.Text)
'            'rststock!P_RETAIL = rststock!MRP
'            GRDSTOCK.TextMatrix(i, 17) = Val(TxtDisc.Text)
'            rststock.Update
'            rststock.MoveNext
'        Loop
'        rststock.Close
'        Set rststock = Nothing
        
    Next i
    TXTDISC.Text = ""
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdLoad_Click()
    Call Fillgrid
End Sub

Private Sub cmdnew_Click()
    If GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 2) <> "" Then
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 0) = GRDSTOCK.rows - 1
        
        Dim TRXMAST As ADODB.Recordset
        On Error GoTo eRRhAND
        
        Set TRXMAST = New ADODB.Recordset
        TRXMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
        If Not (TRXMAST.EOF And TRXMAST.BOF) Then
            If IsNull(TRXMAST.Fields(0)) Then
                GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 1) = 1
            Else
                GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 1) = Val(TRXMAST.Fields(0)) + 1
            End If
        End If
        TRXMAST.Close
        Set TRXMAST = Nothing
    End If
    TXTsample.Visible = False
    TXTEXPIRY.Visible = False
    GRDSTOCK.TopRow = GRDSTOCK.rows - 1
    GRDSTOCK.Row = GRDSTOCK.rows - 1
    GRDSTOCK.Col = 2
    GRDSTOCK.SetFocus
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

'Private Sub Command1_Click()
'    Dim i As Integer
'    Dim rststock As ADODB.Recordset
'    On Error GoTo Errhand
'    If Trim(TxtHSN.Text) = "" Then Exit Sub
'    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THESE HSN CODES", vbYesNo + vbDefaultButton2, "Assign HSN CODES....") = vbNo Then Exit Sub
'    For i = 1 To GRDSTOCK.Rows - 1
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (rststock.EOF And rststock.BOF) Then
'            rststock!Remarks = Trim(TxtHSN.Text)
'            GRDSTOCK.TextMatrix(i, 14) = Trim(TxtHSN.Text)
'            rststock.Update
'        End If
'        rststock.Close
'        Set rststock = Nothing
'    Next i
'    TxtHSN.Text = ""
'    Exit Sub
'Errhand:
'    MsgBox Err.Description
'End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            tXTMEDICINE.SetFocus
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case 73
                Call cmdnew_Click
        End Select
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo eRRhAND
        
    If frmLogin.rs!Level <> "0" Then
        lblpvalue.Visible = False
        lblnetvalue.Visible = False
        Label1(6).Visible = False
    End If
    
    db.Execute "Update itemmast set category = '' where isnull(category) "
    db.Execute "Update itemmast set REMARKS = '' where isnull(REMARKS) "
    
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
        On Error GoTo eRRhAND
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    BARPRINTER = barcodeprinter
    
    Set CMBMFGR.DataSource = Nothing
    MFG_REC.Open "SELECT DISTINCT MANUFACTURER FROM ITEMMAST ORDER BY MANUFACTURER", db, adOpenForwardOnly
    Set CMBMFGR.RowSource = MFG_REC
    CMBMFGR.ListField = "MANUFACTURER"
    
    Set Cmbcategory.DataSource = Nothing
    CAT_REC.Open "SELECT DISTINCT CATEGORY FROM CATEGORY ORDER BY CATEGORY", db, adOpenForwardOnly
    Set Cmbcategory.RowSource = CAT_REC
    Cmbcategory.ListField = "CATEGORY"
    
    REPFLAG = True
    PHY_FLAG = True
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "QTY"
    GRDSTOCK.TextMatrix(0, 4) = "Box Qty"
    GRDSTOCK.TextMatrix(0, 5) = "UOM"
    GRDSTOCK.TextMatrix(0, 6) = "MRP"
    GRDSTOCK.TextMatrix(0, 7) = "RT"
    GRDSTOCK.TextMatrix(0, 8) = "Expiry" '"WS"
    GRDSTOCK.TextMatrix(0, 9) = "Batch" '"VP"
    GRDSTOCK.TextMatrix(0, 10) = "Tax"
    GRDSTOCK.TextMatrix(0, 11) = "Per Rate"
    GRDSTOCK.TextMatrix(0, 12) = "Net Cost"
    GRDSTOCK.TextMatrix(0, 13) = "Pack"
    GRDSTOCK.TextMatrix(0, 14) = "Barcode"
    GRDSTOCK.TextMatrix(0, 15) = "" '"L.Pack"
    GRDSTOCK.TextMatrix(0, 16) = "L.Price"
    GRDSTOCK.TextMatrix(0, 17) = "" '"L.W.Price"
    GRDSTOCK.TextMatrix(0, 18) = "Category"
    GRDSTOCK.TextMatrix(0, 19) = "Company"
    GRDSTOCK.TextMatrix(0, 20) = "Profit%"
    GRDSTOCK.TextMatrix(0, 21) = "Cust Disc"
    GRDSTOCK.TextMatrix(0, 22) = "" '"Commi"
    GRDSTOCK.TextMatrix(0, 23) = "" '"Type"
    GRDSTOCK.TextMatrix(0, 24) = "Value"
    GRDSTOCK.TextMatrix(0, 25) = "" '"Cess%"
    GRDSTOCK.TextMatrix(0, 26) = "" '"Cess Rate"
    GRDSTOCK.TextMatrix(0, 27) = "Spec" '"Cess Rate"
    GRDSTOCK.TextMatrix(0, 28) = "Remarks" '"Cess Rate"
    
    GRDSTOCK.ColWidth(0) = 500
    GRDSTOCK.ColWidth(1) = 800
    GRDSTOCK.ColWidth(2) = 3700
    GRDSTOCK.ColWidth(3) = 950
    GRDSTOCK.ColWidth(4) = 950
    GRDSTOCK.ColWidth(5) = 600
    GRDSTOCK.ColWidth(6) = 950
    GRDSTOCK.ColWidth(7) = 1000
    GRDSTOCK.ColWidth(8) = 1200
    GRDSTOCK.ColWidth(9) = 1000
    GRDSTOCK.ColWidth(10) = 700
    GRDSTOCK.ColWidth(13) = 700
    GRDSTOCK.ColWidth(14) = 2500
    GRDSTOCK.ColWidth(15) = 0
    GRDSTOCK.ColWidth(16) = 950
    GRDSTOCK.ColWidth(17) = 0
    GRDSTOCK.ColWidth(18) = 1250
    GRDSTOCK.ColWidth(19) = 1250
    GRDSTOCK.ColWidth(21) = 900
    GRDSTOCK.ColWidth(22) = 0
    GRDSTOCK.ColWidth(23) = 0
    GRDSTOCK.ColWidth(24) = 1500
    GRDSTOCK.ColWidth(25) = 0 '900
    GRDSTOCK.ColWidth(26) = 0 '1200
    GRDSTOCK.ColWidth(27) = 1000 '1200
    GRDSTOCK.ColWidth(28) = 1000 '1200
    
    If frmLogin.rs!Level = "0" Then
        GRDSTOCK.ColWidth(11) = 900
        GRDSTOCK.ColWidth(12) = 900
        GRDSTOCK.ColWidth(20) = 900
    Else
        GRDSTOCK.ColWidth(11) = 0
        GRDSTOCK.ColWidth(12) = 0
        GRDSTOCK.ColWidth(20) = 0
    End If
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 4
    GRDSTOCK.ColAlignment(5) = 4
    GRDSTOCK.ColAlignment(6) = 4
    GRDSTOCK.ColAlignment(7) = 4
    GRDSTOCK.ColAlignment(8) = 4
    GRDSTOCK.ColAlignment(9) = 4
    GRDSTOCK.ColAlignment(10) = 4
    GRDSTOCK.ColAlignment(11) = 4
    GRDSTOCK.ColAlignment(12) = 4
    GRDSTOCK.ColAlignment(13) = 4
    GRDSTOCK.ColAlignment(14) = 4
    GRDSTOCK.ColAlignment(15) = 4
    GRDSTOCK.ColAlignment(16) = 4
    GRDSTOCK.ColAlignment(17) = 4
    GRDSTOCK.ColAlignment(18) = 1
    GRDSTOCK.ColAlignment(19) = 1
    GRDSTOCK.ColAlignment(20) = 4
    GRDSTOCK.ColAlignment(21) = 4
    GRDSTOCK.ColAlignment(22) = 4
    GRDSTOCK.ColAlignment(23) = 4
    GRDSTOCK.ColAlignment(24) = 1
    GRDSTOCK.ColAlignment(25) = 4
    GRDSTOCK.ColAlignment(26) = 4
    GRDSTOCK.ColAlignment(27) = 4
    GRDSTOCK.ColAlignment(28) = 4
    
    DTFROM.Value = Format(Date, "DD/MM/YYYY")
    DTFROM.Value = Null
    'Call Fillgrid
    'Me.Height = 8415
    'Me.Width = 6465
    Me.Left = 0
    Me.Top = 0
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    If PHY_FLAG = False Then PHY_REC.Close
    MFG_REC.Close
    CAT_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDSTOCK_Click()
    TXTsample.Visible = False
    TXTEXPIRY.Visible = False
    CmbPack.Visible = False
    CMBMFGR.Visible = False
    FRAME.Visible = False
    GRDSTOCK.SetFocus
End Sub

Private Sub GRDSTOCK_DblClick()
    Dim i As Single
    Dim M As Long
    Dim temp_file As String
    
    If GRDSTOCK.rows <= 1 Then Exit Sub
    If GRDSTOCK.Cols = 20 Then Exit Sub
    
'    If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 22)) = "" Then
'
'    End If
    
    i = Val(InputBox("Enter number of lables to be print", "No. of labels..", GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)))
    If i <= 0 Then Exit Sub
    
    'If Not IsDate(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) And Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 28)) <> 0 And Len(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 28))) > 0 And Len(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 28))) < 3 Then
    If IsNumeric(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 28)) And Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 28)) > 0 And Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 28)) < 366 Then
        If IsNull(DTFROM.Value) Then
            MsgBox "Please select the date of manufacturing", , "Barcode printing"
            Exit Sub
        End If
        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8) = Format(DateAdd("d", Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 28)), DTFROM.Value), "dd/mm/yyyy")
    End If
    
    On Error GoTo eRRhAND
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
        'If FileExists(App.Path & "\template.txt") Then
            Set ObjFile = CreateObject("Scripting.FileSystemObject")
            Set objText = ObjFile.OpenTextFile(App.Path & temp_file)
            Text = objText.ReadAll
            objText.Close
        
            Set objText = Nothing
            Set ObjFile = Nothing
            
            Dim rstformula As ADODB.Recordset
            Dim pergr As Integer
            pergr = 0
            Set rstformula = New ADODB.Recordset
            rstformula.Open "select * from ITEMMAST where ITEM_CODE = '" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstformula.EOF Or rstformula.BOF) Then
                If IsNull(rstformula!ITEM_SPEC) Then
                    pergr = 0
                Else
                    pergr = IIf(IsNull(rstformula!ITEM_SPEC), 0, Val(rstformula!ITEM_SPEC))
                End If
            End If
            rstformula.Close
            Set rstformula = Nothing
            If pergr > 1 And Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) <> 0 Then
                Text = Replace(Text, "[PPPPPPPP]", "" & Round(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) / pergr, 3) & "") 'pergram
            Else
                Text = Replace(Text, "[PPPPPPPP]", "")   'REF (SPEC)
            End If
                    
            Text = Replace(Text, "[AAAAAAAA]", "" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 27)) & "")  'REF (SPEC)
            Text = Replace(Text, "[BBBBBBBB]", "") 'PACK
            'if Format(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), "dd/mm/yyyy")
            If IsDate(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) Then
                If Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 1, 2)) <> 0 And Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 4, 5)) <= 12 And Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 1, 2)) > 0 And Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 4, 5)) > 0 Then
                    Text = Replace(Text, "[EEEEEEEE]", "" & Format(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), "dd/mm/yyyy") & "")  'EXP DATE
                Else
                    Text = Replace(Text, "[EEEEEEEE]", "")   'EXP DATE
                End If
                If IsNull(DTFROM.Value) Then
                    Text = Replace(Text, "[CCCCCCCC]", "" & Format(Date, "dd/mm/yyyy") & "")  'PACK DATE
                Else
                    Text = Replace(Text, "[CCCCCCCC]", "" & Format(DTFROM.Value, "dd/mm/yyyy") & "")  'PACK DATE
                End If
            Else
                Text = Replace(Text, "[EEEEEEEE]", "")   'EXP DATE
                Text = Replace(Text, "[CCCCCCCC]", "")  'PACK DATE
            End If
            
            Text = Replace(Text, "[DDDDDDDD]", "" & Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6)), "0.00") & "")  'MRP
            Text = Replace(Text, "[FFFFFFFF]", "" & Left(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)), 30) & "")  'ITEM NAME
            Text = Replace(Text, "[NNNNNNNN]", "" & Left(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)), 30) & "")  'ITEM CODE
            If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) = "" Then
                Text = Replace(Text, "[GGGGGGGG]", "" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "")  'BARCODE
            Else
                Text = Replace(Text, "[GGGGGGGG]", "" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14) & "")  'BARCODE
            End If
            'If BARFORMAT = "Y" Then
            If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) = "" Then
                If Len(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1))) Mod 2 = 0 Then
                    Text = Replace(Text, "[LLLLLLLL]", "" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)) & "")  'BARCODE
                    Text = Replace(Text, "[MMMMMMMM]", "" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)) & "")  'BARCODE
                Else
                    Text = Replace(Text, "[LLLLLLLL]", "" & Mid(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)), 1, Len(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1))) - 1) & "!100" & Right(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)), 1) & "") 'BARCODE
                    Text = Replace(Text, "[MMMMMMMM]", "" & Mid(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)), 1, Len(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1))) - 1) & ">6" & Right(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)), 1) & "") 'BARCODE
                End If
            Else
                If Len(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14))) Mod 2 = 0 Then
                    Text = Replace(Text, "[LLLLLLLL]", "" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & "")  'BARCODE
                    Text = Replace(Text, "[MMMMMMMM]", "" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & "")  'BARCODE
                Else
                    Text = Replace(Text, "[LLLLLLLL]", "" & Mid(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)), 1, Len(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14))) - 1) & "!100" & Right(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)), 1) & "") 'BARCODE
                    Text = Replace(Text, "[MMMMMMMM]", "" & Mid(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)), 1, Len(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14))) - 1) & ">6" & Right(Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)), 1) & "") 'BARCODE
                End If
            End If
            'End If
            Text = Replace(Text, "[QQQQQQQQ]", "" & Decode_Cost(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12))))  'COST
            Text = Replace(Text, "[HHHHHHHH]", "" & Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)), "0.00") & "")  'PRICE
            Text = Replace(Text, "[IIIIIIII]", "" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9)) & "")  'BATCH
            Text = Replace(Text, "[JJJJJJJJ]", "" & Trim(MDIMAIN.StatusBar.Panels(5).Text) & "")  'COMP NAME
            
            Dim intFile As Integer
            Dim strFile As String
            If FileExists(App.Path & "\BARCODE.PRN") Then
                Kill (App.Path & "\BARCODE.PRN")
            End If
            strFile = App.Path & "\BARCODE.PRN" 'the file you want to save to
            intFile = FreeFile
            Open strFile For Output As #intFile
                Print #intFile, Text 'the data you want to save
            Close #intFile
            
            On Error GoTo CLOSEFILE
            Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
            If err.Number = 55 Then
                Close #1
                Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
            End If
            On Error GoTo eRRhAND
            
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
        db.Execute "Delete from barprint"
        Dim RSTTRXFILE As ADODB.Recordset
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From barprint", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        For M = 1 To i
            RSTTRXFILE.AddNew
            If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) = "" Then
                RSTTRXFILE!BARCODE = "*" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "*"
            Else
                RSTTRXFILE!BARCODE = "*" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14) & "*"
            End If
            RSTTRXFILE!ITEM_NAME = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2))
            RSTTRXFILE!item_Price = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
            RSTTRXFILE!item_COST = Decode_Cost(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12)))
            RSTTRXFILE!item_MRP = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6))
            
            If IsDate(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) Then
                If Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 1, 2)) <> 0 And Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 4, 5)) <= 12 And Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 1, 2)) > 0 And Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 4, 5)) > 0 Then
                    RSTTRXFILE!expdate = Format(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), "dd/mm/yyyy")
                End If
            End If
            If IsNull(DTFROM.Value) Then
                RSTTRXFILE!pckdate = Format(Date, "dd/mm/yyyy")
            Else
                RSTTRXFILE!pckdate = Format(DTFROM.Value, "dd/mm/yyyy")
            End If
            RSTTRXFILE!item_color = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9))
            RSTTRXFILE!REMARKS = ""
                    
            RSTTRXFILE!COMP_NAME = Trim(MDIMAIN.StatusBar.Panels(5).Text)
            RSTTRXFILE.Update
        Next M
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        db.CommitTrans
        
        If BARPRINTER = barcodeprinter Then
            ReportNameVar = Rptpath & "Rptbarprn"
        Else
            ReportNameVar = Rptpath & "Rptbarprn1"
        End If
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Set CRXFormulaFields = Report.FormulaFields
    
        For M = 1 To Report.Database.Tables.COUNT
            Report.Database.Tables.Item(M).SetLogOnInfo strConnection
            If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                Set oRs = New ADODB.Recordset
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(M).Name & " ")
                Report.Database.SetDataSource oRs, 3, M
                Set oRs = Nothing
            End If
        Next M
        Set Printer = Printers(BARPRINTER)
        Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
        Report.DiscardSavedData
        Report.VerifyOnEveryPrint = True
        Report.PrintOut (False)
        Set CRXFormulaFields = Nothing
        Set crxApplication = Nothing
        Set Report = Nothing
    End If
    Screen.MousePointer = vbNormal
    Exit Sub

eRRhAND:
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

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Integer
    If GRDSTOCK.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            If Not (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then Exit Sub
            Select Case GRDSTOCK.Col
                Case 3
                    If UCase(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 18)) = "SELF" Then
                        
                        Exit Sub
                    End If
                        
                    If IsNull(DTFROM.Value) Then
                        MsgBox "Select the Date for Opening Qty", vbOKOnly, "Price Analysis"
                        GRDSTOCK.SetFocus
                        Exit Sub
                    End If
                    If (DTFROM.Value) > Date Then
                        MsgBox "The date could not be greater than Today", vbOKOnly, "Price Analysis"
                        GRDSTOCK.SetFocus
                        Exit Sub
                    End If
'                    If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) = 0 Then
'                        MsgBox "Please enter the cost", vbOKOnly, "Price Analysis"
'                        GRDSTOCK.SetFocus
'                        Exit Sub
'                    End If
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 100
                    TXTsample.Left = GRDSTOCK.CellLeft '+ 60
                    TXTsample.Width = GRDSTOCK.CellWidth
                    TXTsample.Height = GRDSTOCK.CellHeight
                    TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
                'Case 1, 2, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17, 20, 21, 25, 26
                Case 1, 2, 21, 14, 7, 9, 10, 11, 6, 13, 16, 27, 28 ', 7, 8, 9, 10, 11, 13, 14, 15,  17, 20, 21, 25, 26
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 100
                    TXTsample.Left = GRDSTOCK.CellLeft '+ 60
                    TXTsample.Width = GRDSTOCK.CellWidth
                    TXTsample.Height = GRDSTOCK.CellHeight
                    TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
                Case 5
                    CmbPack.Visible = True
                    CmbPack.Top = GRDSTOCK.CellTop + 100
                    CmbPack.Left = GRDSTOCK.CellLeft '+ 60
                    CmbPack.Width = GRDSTOCK.CellWidth
                    'CmbPack.Height = GRDSTOCK.CellHeight
                    On Error Resume Next
                    CmbPack.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    CmbPack.SetFocus
                Case 8
                    TXTEXPIRY.Visible = True
                    TXTEXPIRY.Top = GRDSTOCK.CellTop + 100
                    TXTEXPIRY.Left = GRDSTOCK.CellLeft '+ 50
                    TXTEXPIRY.Width = GRDSTOCK.CellWidth '- 25
                    If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)) <> 0 Then
                        TXTEXPIRY.Text = IIf(IsDate(GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)), Format(GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col), "DD/MM/YYYY"), "  /  /    ")
                    Else
                        TXTEXPIRY.Text = "  /  /    "
                    End If
                    TXTEXPIRY.SetFocus
                Case 18
                    Cmbcategory.Visible = True
                    Cmbcategory.Top = GRDSTOCK.CellTop + 100
                    Cmbcategory.Left = GRDSTOCK.CellLeft '+ 60
                    Cmbcategory.Width = GRDSTOCK.CellWidth
                    'CmbPack.Height = GRDSTOCK.CellHeight
                    On Error Resume Next
                    Cmbcategory.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    Cmbcategory.SetFocus
                Case 19
                    CMBMFGR.Visible = True
                    CMBMFGR.Top = GRDSTOCK.CellTop + 100
                    CMBMFGR.Left = GRDSTOCK.CellLeft '+ 60
                    CMBMFGR.Width = GRDSTOCK.CellWidth
                    'CmbPack.Height = GRDSTOCK.CellHeight
                    On Error Resume Next
                    CMBMFGR.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    CMBMFGR.SetFocus
                Case 22
                    FRAME.Visible = True
                    FRAME.Top = GRDSTOCK.CellTop - 300
                    FRAME.Left = GRDSTOCK.CellLeft - 1500
                    'Frame.Width = GRDSTOCK.CellWidth - 25
                    TxtComper.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    If GRDSTOCK.TextMatrix(GRDSTOCK.Row, 23) = "Rs" Then
                        OptAmt.Value = True
                    Else
                        OptPercent.Value = True
                    End If
                    TxtComper.SetFocus
            End Select
        Case 114
            sitem = UCase(InputBox("Item Name...?", "STOCK"))
            For i = 1 To GRDSTOCK.rows - 1
                If UCase(Mid(GRDSTOCK.TextMatrix(i, 2), 1, Len(sitem))) = sitem Then
                    GRDSTOCK.Row = i
                    GRDSTOCK.TopRow = i
                    Exit For
                End If
            Next i
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub GRDSTOCK_RowColChange()
    lblitemname.Caption = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)
End Sub

Private Sub GRDSTOCK_Scroll()
    TXTsample.Visible = False
    TXTEXPIRY.Visible = False
    CmbPack.Visible = False
    CMBMFGR.Visible = False
    FRAME.Visible = False
End Sub

Private Sub Label1_DblClick(index As Integer)
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If frmunbill.Visible = False Then
        frmunbill.Visible = True
        chkunbill.Value = 0
        chkonlyunbill.Value = 0
    Else
        frmunbill.Visible = False
        chkunbill.Value = 0
        chkonlyunbill.Value = 0
    End If
End Sub

Private Sub Chkunbill_Click()
    If chkunbill.Value = 1 Then
        chkonlyunbill.Visible = True
        chkonlyunbill.Value = 0
    Else
        chkonlyunbill.Visible = False
        chkonlyunbill.Value = 0
    End If
End Sub

Private Sub OptAll_Click()
    tXTMEDICINE.SetFocus
End Sub

Private Sub OptAmt_Click()
    TxtComper.SetFocus
End Sub

Private Sub OptPercent_Click()
    TxtComper.SetFocus
End Sub

Private Sub OptStock_Click()
    tXTMEDICINE.SetFocus
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Dim rststock As ADODB.Recordset
            Dim i As Long
            db.Execute "Update Itemmast Set BARCODE ='' where isnull(BARCODE)"
            On Error GoTo eRRhAND
            
            i = 0
            Screen.MousePointer = vbHourglass
            
            GRDSTOCK.rows = 1
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE = '" & TxtBarcode.Text & "' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY BARCODE", db, adOpenForwardOnly
            Do Until rststock.EOF
                i = i + 1
                GRDSTOCK.rows = GRDSTOCK.rows + 1
                GRDSTOCK.FixedRows = 1
                GRDSTOCK.TextMatrix(i, 0) = i
                GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
                GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
                GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!CLOSE_QTY), 0, rststock!CLOSE_QTY)
                GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                If Val(GRDSTOCK.TextMatrix(i, 13)) = 0 Then GRDSTOCK.TextMatrix(i, 13) = 1
                GRDSTOCK.TextMatrix(i, 4) = Round(Val(GRDSTOCK.TextMatrix(i, 3)) / Val(GRDSTOCK.TextMatrix(i, 13)), 0)
                GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!PACK_TYPE), "", rststock!PACK_TYPE)
                GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.00"))
                GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 8) = "" 'IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 9) = "" 'IIf(IsNull(rststock!P_VAN), "", Format(rststock!P_VAN, "0.00"))
                GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!SALES_TAX), "", Format(rststock!SALES_TAX, "0.00"))
                GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!item_COST), "", Format(rststock!item_COST, "0.00"))
                GRDSTOCK.TextMatrix(i, 12) = (Val(GRDSTOCK.TextMatrix(i, 11)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 10)) / 100)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 25)) / 100) + Val(GRDSTOCK.TextMatrix(i, 26))
                GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
                GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!CRTN_PACK), "", rststock!CRTN_PACK)
                GRDSTOCK.TextMatrix(i, 16) = IIf(IsNull(rststock!P_CRTN), "", Format(Round(rststock!P_CRTN, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 17) = IIf(IsNull(rststock!P_LWS), "", Format(Round(rststock!P_LWS, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 18) = IIf(IsNull(rststock!Category), "", rststock!Category)
                GRDSTOCK.TextMatrix(i, 19) = IIf(IsNull(rststock!MANUFACTURER), "", rststock!MANUFACTURER)
                If Val(GRDSTOCK.TextMatrix(i, 11)) <> 0 Then
                    GRDSTOCK.TextMatrix(i, 20) = Round((((Val(GRDSTOCK.TextMatrix(i, 7)) / Val(GRDSTOCK.TextMatrix(i, 13))) - Val(GRDSTOCK.TextMatrix(i, 12))) * 100) / Val(GRDSTOCK.TextMatrix(i, 12)), 2)
                Else
                    GRDSTOCK.TextMatrix(i, 20) = 0
                End If
                GRDSTOCK.TextMatrix(i, 21) = IIf(IsNull(rststock!CUST_DISC), "", Format(rststock!CUST_DISC, "0.00"))
                Select Case rststock!COM_FLAG
                    Case "P"
                        GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!COM_PER), "", Format(rststock!COM_PER, "0.00"))
                        GRDSTOCK.TextMatrix(i, 23) = "%"
                    Case "A"
                        GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!COM_AMT), "", Format(rststock!COM_AMT, "0.00"))
                        GRDSTOCK.TextMatrix(i, 23) = "Rs"
                End Select
                GRDSTOCK.TextMatrix(i, 24) = IIf(IsNull(rststock!item_COST), "", Format(rststock!item_COST * rststock!CLOSE_QTY, "0.00"))
                GRDSTOCK.TextMatrix(i, 25) = IIf(IsNull(rststock!CESS_PER), "", Format(rststock!CESS_PER, "0.00"))
                GRDSTOCK.TextMatrix(i, 26) = IIf(IsNull(rststock!cess_amt), "", Format(rststock!cess_amt, "0.00"))
                GRDSTOCK.TextMatrix(i, 27) = IIf(IsNull(rststock!ITEM_SPEC), "", rststock!ITEM_SPEC)
                GRDSTOCK.TextMatrix(i, 28) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
                
                If rststock!UN_BILL = "Y" Then GRDSTOCK.TextMatrix(i, 0) = GRDSTOCK.TextMatrix(i, 0) & "*"
                    
        '        Set rstopstock = New ADODB.Recordset
        '        rstopstock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & rststock!ITEM_CODE & "' AND TRX_TYPE ='ST'", db, adOpenStatic, adLockReadOnly
        '        If Not (rstopstock.EOF And rstopstock.BOF) Then
        '            GRDSTOCK.TextMatrix(i, 22) = "*"
        '        Else
        '            GRDSTOCK.TextMatrix(i, 22) = ""
        '        End If
        '        rstopstock.Close
        '        Set rstopstock = Nothing
                
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            Call Toatal_value
            Screen.MousePointer = vbNormal
    End Select
    Screen.MousePointer = vbNormal
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim rststock As ADODB.Recordset
'    Dim M_DATE As Date
'    Dim D As Integer
'    Dim M As Integer
'    Dim Y As Integer
'
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            'IF NOT ISDATE
'            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
'            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
'            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
'
'            M = Val(Mid(TXTEXPIRY.Text, 1, 2))
'            Y = Val(Right(TXTEXPIRY.Text, 2))
'            Y = 2000 + Y
'            M_DATE = "01" & "/" & M & "/" & Y
'            D = LastDayOfMonth(M_DATE)
            If IsDate(TXTEXPIRY.Text) Then
                'M_DATE = Format(TXTEXPIRY.Text, "DD/MM/YYYY")
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic
                If Not (rststock.EOF And rststock.BOF) Then
                    rststock!EXP_DATE = Format(TXTEXPIRY.Text, "DD/MM/YYYY")
                    'rststock!VCH_DATE = Format(M_DATE, "dd/mm/yyyy")
                    rststock.Update
                End If
                rststock.Close
                Set rststock = Nothing
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(TXTEXPIRY.Text, "DD/MM/YYYY")
            Else
                'M_DATE = Null
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic
                If Not (rststock.EOF And rststock.BOF) Then
                    rststock!EXP_DATE = Null
                    'rststock!VCH_DATE = Format(M_DATE, "dd/mm/yyyy")
                    rststock.Update
                End If
                rststock.Close
                Set rststock = Nothing
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = ""
            End If
            'M_DATE = D & "/" & M & "/" & Y
            
            
            
            TXTEXPIRY.Visible = False
            
            GRDSTOCK.Enabled = True
            GRDSTOCK.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            GRDSTOCK.SetFocus
    End Select
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub TXTITEMCODE_Change()
    On Error GoTo eRRhAND
    'Call Fillgrid
    If REPFLAG = True Then
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_CODE Like '" & Trim(TXTITEMCODE.Text) & "%'  AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Trim(TXTITEMCODE.Text) & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Trim(TXTITEMCODE.Text) & "%'  AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Trim(TXTITEMCODE.Text) & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Trim(TXTITEMCODE.Text) & "%'  AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Trim(TXTITEMCODE.Text) & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Trim(TXTITEMCODE.Text) & "%'  AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Trim(TXTITEMCODE.Text) & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
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
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Dim rststock As ADODB.Recordset
            Dim rstopstock As ADODB.Recordset
            Dim i As Long
        
            On Error GoTo eRRhAND
            
            i = 0
            Screen.MousePointer = vbHourglass
            
            GRDSTOCK.rows = 1
            Set rststock = New ADODB.Recordset
            If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y'  AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            '==========================================================================================
            ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y'  AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
            '=======================================================================================
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y'  AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
        
            Do Until rststock.EOF
                i = i + 1
                GRDSTOCK.rows = GRDSTOCK.rows + 1
                GRDSTOCK.FixedRows = 1
                GRDSTOCK.TextMatrix(i, 0) = i
                GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
                GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
                GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!CLOSE_QTY), 0, rststock!CLOSE_QTY)
                GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                If Val(GRDSTOCK.TextMatrix(i, 13)) = 0 Then GRDSTOCK.TextMatrix(i, 13) = 1
                GRDSTOCK.TextMatrix(i, 4) = Round(Val(GRDSTOCK.TextMatrix(i, 3)) / Val(GRDSTOCK.TextMatrix(i, 13)), 0)
                GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!PACK_TYPE), "", rststock!PACK_TYPE)
                GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.00"))
                GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 8) = "" 'IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 9) = "" 'IIf(IsNull(rststock!P_VAN), "", Format(rststock!P_VAN, "0.00"))
                GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!SALES_TAX), "", Format(rststock!SALES_TAX, "0.00"))
                GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!item_COST), "", Format(rststock!item_COST, "0.00"))
                GRDSTOCK.TextMatrix(i, 12) = (Val(GRDSTOCK.TextMatrix(i, 11)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 10)) / 100)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 25)) / 100) + Val(GRDSTOCK.TextMatrix(i, 26))
                GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
                GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!CRTN_PACK), "", rststock!CRTN_PACK)
                GRDSTOCK.TextMatrix(i, 16) = IIf(IsNull(rststock!P_CRTN), "", Format(Round(rststock!P_CRTN, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 17) = IIf(IsNull(rststock!P_LWS), "", Format(Round(rststock!P_LWS, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 18) = IIf(IsNull(rststock!Category), "", rststock!Category)
                'GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
                GRDSTOCK.TextMatrix(i, 19) = IIf(IsNull(rststock!MANUFACTURER), "", rststock!MANUFACTURER)
                If Val(GRDSTOCK.TextMatrix(i, 11)) <> 0 Then
                    GRDSTOCK.TextMatrix(i, 20) = Round((((Val(GRDSTOCK.TextMatrix(i, 7)) / Val(GRDSTOCK.TextMatrix(i, 13))) - Val(GRDSTOCK.TextMatrix(i, 12))) * 100) / Val(GRDSTOCK.TextMatrix(i, 12)), 2)
                Else
                    GRDSTOCK.TextMatrix(i, 20) = 0
                End If
                GRDSTOCK.TextMatrix(i, 21) = IIf(IsNull(rststock!CUST_DISC), "", Format(rststock!CUST_DISC, "0.00"))
                Select Case rststock!COM_FLAG
                    Case "P"
                        GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!COM_PER), "", Format(rststock!COM_PER, "0.00"))
                        GRDSTOCK.TextMatrix(i, 23) = "%"
                    Case "A"
                        GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!COM_AMT), "", Format(rststock!COM_AMT, "0.00"))
                        GRDSTOCK.TextMatrix(i, 23) = "Rs"
                End Select
                GRDSTOCK.TextMatrix(i, 24) = IIf(IsNull(rststock!item_COST), "", Format(rststock!item_COST * rststock!CLOSE_QTY, "0.00"))
                GRDSTOCK.TextMatrix(i, 25) = IIf(IsNull(rststock!CESS_PER), "", Format(rststock!CESS_PER, "0.00"))
                GRDSTOCK.TextMatrix(i, 26) = IIf(IsNull(rststock!cess_amt), "", Format(rststock!cess_amt, "0.00"))
                GRDSTOCK.TextMatrix(i, 27) = IIf(IsNull(rststock!ITEM_SPEC), "", rststock!ITEM_SPEC)
                GRDSTOCK.TextMatrix(i, 28) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
                If rststock!UN_BILL = "Y" Then GRDSTOCK.TextMatrix(i, 0) = GRDSTOCK.TextMatrix(i, 0) & "*"
                    
        '        Set rstopstock = New ADODB.Recordset
        '        rstopstock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & rststock!ITEM_CODE & "' AND TRX_TYPE ='ST'", db, adOpenStatic, adLockReadOnly
        '        If Not (rstopstock.EOF And rstopstock.BOF) Then
        '            GRDSTOCK.TextMatrix(i, 22) = "*"
        '        Else
        '            GRDSTOCK.TextMatrix(i, 22) = ""
        '        End If
        '        rstopstock.Close
        '        Set rstopstock = Nothing
                
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            Call Toatal_value
            
            DTFROM.Value = Null
            Screen.MousePointer = vbNormal
    
    End Select
Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub tXTMEDICINE_Change()
    On Error GoTo eRRhAND
    If Trim(TxtCode.Text) <> "" Then Call Fillgrid
    If REPFLAG = True Then
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
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
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.Text)
    'Call Fillgrid
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            'TxtCode.SetFocus
            Call CmdLoad_Click
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

Private Function Fillgrid()
    Dim rststock As ADODB.Recordset
    Dim rstopstock As ADODB.Recordset
    Dim i As Long

    On Error GoTo eRRhAND
    
    i = 0
    Screen.MousePointer = vbHourglass
    
    GRDSTOCK.rows = 1
    Set rststock = New ADODB.Recordset
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
            If OptStock.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            ElseIf OptPC.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If CHKCATEGORY2.Value = 1 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
    
            End If
        End If
    '==========================================================================================
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
            If OptStock.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            ElseIf OptPC.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If CHKCATEGORY2.Value = 1 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
    
            End If
        End If
    Else
    '=======================================================================================
        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
            If OptStock.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            ElseIf OptPC.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If CHKCATEGORY2.Value = 1 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%'  AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
    
            End If
        End If
    End If

    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!CLOSE_QTY), 0, rststock!CLOSE_QTY)
        GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
        If Val(GRDSTOCK.TextMatrix(i, 13)) = 0 Then GRDSTOCK.TextMatrix(i, 13) = 1
        GRDSTOCK.TextMatrix(i, 4) = Round(Val(GRDSTOCK.TextMatrix(i, 3)) / Val(GRDSTOCK.TextMatrix(i, 13)), 0)
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!PACK_TYPE), "", rststock!PACK_TYPE)
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.00"))
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 1), "0.000"))
        GRDSTOCK.TextMatrix(i, 8) = "" 'IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 1), "0.000"))
        GRDSTOCK.TextMatrix(i, 9) = "" 'IIf(IsNull(rststock!P_VAN), "", Format(rststock!P_VAN, "0.00"))
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!SALES_TAX), "", Format(rststock!SALES_TAX, "0.00"))
        
        If frmLogin.rs!Level = "0" Then
            GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!item_COST), "", Format(rststock!item_COST, "0.00"))
            GRDSTOCK.TextMatrix(i, 12) = (Val(GRDSTOCK.TextMatrix(i, 11)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 10)) / 100)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 25)) / 100) + Val(GRDSTOCK.TextMatrix(i, 26))
            If Val(GRDSTOCK.TextMatrix(i, 11)) <> 0 Then
                GRDSTOCK.TextMatrix(i, 20) = Round((((Val(GRDSTOCK.TextMatrix(i, 7)) / Val(GRDSTOCK.TextMatrix(i, 13))) - Val(GRDSTOCK.TextMatrix(i, 12))) * 100) / Val(GRDSTOCK.TextMatrix(i, 12)), 2)
            Else
                GRDSTOCK.TextMatrix(i, 20) = 0
            End If
        End If
        GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
        GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!CRTN_PACK), "", rststock!CRTN_PACK)
        GRDSTOCK.TextMatrix(i, 16) = IIf(IsNull(rststock!P_CRTN), "", Format(Round(rststock!P_CRTN, 1), "0.000"))
        GRDSTOCK.TextMatrix(i, 17) = IIf(IsNull(rststock!P_LWS), "", Format(Round(rststock!P_LWS, 1), "0.000"))
        GRDSTOCK.TextMatrix(i, 18) = IIf(IsNull(rststock!Category), "", rststock!Category)
        
        GRDSTOCK.TextMatrix(i, 19) = IIf(IsNull(rststock!MANUFACTURER), "", rststock!MANUFACTURER)
        
        GRDSTOCK.TextMatrix(i, 21) = IIf(IsNull(rststock!CUST_DISC), "", Format(rststock!CUST_DISC, "0.00"))
        Select Case rststock!COM_FLAG
            Case "P"
                GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!COM_PER), "", Format(rststock!COM_PER, "0.00"))
                GRDSTOCK.TextMatrix(i, 23) = "%"
            Case "A"
                GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!COM_AMT), "", Format(rststock!COM_AMT, "0.00"))
                GRDSTOCK.TextMatrix(i, 23) = "Rs"
        End Select
        GRDSTOCK.TextMatrix(i, 24) = IIf(IsNull(rststock!item_COST), "", Format(rststock!item_COST * rststock!CLOSE_QTY, "0.00"))
        GRDSTOCK.TextMatrix(i, 25) = IIf(IsNull(rststock!CESS_PER), "", Format(rststock!CESS_PER, "0.00"))
        GRDSTOCK.TextMatrix(i, 26) = IIf(IsNull(rststock!cess_amt), "", Format(rststock!cess_amt, "0.00"))
        GRDSTOCK.TextMatrix(i, 27) = IIf(IsNull(rststock!ITEM_SPEC), "", rststock!ITEM_SPEC)
        GRDSTOCK.TextMatrix(i, 28) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
        'ITEM_SPEC
        If rststock!UN_BILL = "Y" Then GRDSTOCK.TextMatrix(i, 0) = GRDSTOCK.TextMatrix(i, 0) & "*"
            
'        Set rstopstock = New ADODB.Recordset
'        rstopstock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & rststock!ITEM_CODE & "' AND TRX_TYPE ='ST'", db, adOpenStatic, adLockReadOnly
'        If Not (rstopstock.EOF And rstopstock.BOF) Then
'            GRDSTOCK.TextMatrix(i, 22) = "*"
'        Else
'            GRDSTOCK.TextMatrix(i, 22) = ""
'        End If
'        rstopstock.Close
'        Set rstopstock = Nothing
        
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    Call Toatal_value
    
    DTFROM.Value = Null
    Screen.MousePointer = vbNormal
    Exit Function

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock, RSTITEMMAST As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            Select Case GRDSTOCK.Col
                Case 1  ' Item Code
                    If Trim(TXTsample.Text) = "" Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_CODE = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!ITEM_CODE = Trim(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from TRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!ITEM_CODE = Trim(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 2  ' Item Name
                    If Trim(TXTsample.Text) = "" Then Exit Sub
                    Dim item_unbill As Boolean
                    item_unbill = False
                    If frmunbill.Visible = True And chkunbill.Value = 1 Then
                        If MsgBox("Click No to create unbill Item or else click Yes", vbYesNo, "Item Creation") = vbNo Then
                            item_unbill = True
                        Else
                            item_unbill = False
                        End If
                    End If
                
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If (rststock.EOF And rststock.BOF) Then
                        rststock.AddNew
                        rststock!ITEM_CODE = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)
                        rststock!ITEM_NAME = Trim(TXTsample.Text)
                        rststock!Category = "GENERAL"
                        rststock!UNIT = 1
                        rststock!MANUFACTURER = "GENERAL"
                        rststock!DEAD_STOCK = "N"
                        rststock!REMARKS = ""
                        rststock!REORDER_QTY = 1
                        rststock!PACK_TYPE = "Nos"
                        rststock!FULL_PACK = "Nos"
                        rststock!BIN_LOCATION = ""
                        rststock!MRP = 0
                        rststock!PTR = 0
                        rststock!CST = 0
                        rststock!OPEN_QTY = 0
                        rststock!OPEN_VAL = 0
                        rststock!RCPT_QTY = 0
                        rststock!RCPT_VAL = 0
                        rststock!ISSUE_QTY = 0
                        rststock!ISSUE_VAL = 0
                        rststock!CLOSE_QTY = 0
                        rststock!CLOSE_VAL = 0
                        rststock!DAM_QTY = 0
                        rststock!DAM_VAL = 0
                        rststock!DISC = 0
                        rststock!SALES_TAX = 0
                        rststock!item_COST = 0
                        rststock!P_RETAIL = 0
                        rststock!P_WS = 0
                        rststock!CRTN_PACK = 1
                        rststock!P_CRTN = 0
                        rststock!LOOSE_PACK = 1
                        If PC_FLAG = "Y" Then
                            rststock!PRICE_CHANGE = "Y"
                        Else
                            rststock!PRICE_CHANGE = "N"
                        End If
                        If item_unbill = True Then
                            rststock!UN_BILL = "Y"
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 0) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 0) & "*"
                        Else
                            rststock!UN_BILL = "N"
                        End If
                        
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5) = "Nos"
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = 1
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 18) = "GENERAL"
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 19) = "GENERAL"
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3) = 0
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 4) = 0
                    Else
                        rststock!ITEM_NAME = Trim(TXTsample.Text)
                    End If
                    rststock.Update
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!ITEM_NAME = Trim(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 3
                    Dim INWARD, OUTWARD, BAL_QTY As Double
                    Dim TRXMAST As ADODB.Recordset
                    
                    Screen.MousePointer = vbHourglass
                    Set RSTITEMMAST = New ADODB.Recordset
                    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                        INWARD = 0
                        OUTWARD = 0
                        BAL_QTY = 0
'
'                        Set TRXMAST = New ADODB.Recordset
'                        TRXMAST.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE = '" & Format(Date, "yyyy/mm/dd") & "' AND TRX_TYPE = 'ST' ", db, adOpenStatic, adLockOptimistic, adCmdText
'                        Set rststock = New ADODB.Recordset
'                        If TRXMAST.RecordCount > 0 Then
'                            rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' and TRX_TYPE <> 'ST'", db, adOpenStatic, adLockReadOnly
'                        Else
'                            rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
'                        End If
'                        TRXMAST.Close
'                        Set TRXMAST = Nothing
                        
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
                        rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='MI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockReadOnly
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
                        
                        If Not (Val(TXTsample.Text) - (Val(INWARD - OUTWARD)) = 0) Then
                            Set rststock = New ADODB.Recordset
                            '.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE = '" & Format(Date, "yyyy/mm/dd") & "' AND TRX_TYPE = 'ST' ", db, adOpenStatic, adLockOptimistic, adCmdText
                            rststock.Open "SELECT * FROM RTRXFILE ", db, adOpenStatic, adLockOptimistic, adCmdText
                            'If (rststock.EOF And rststock.BOF) Then
                                rststock.AddNew
                                rststock!TRX_TYPE = "ST"
                                rststock!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                                rststock!VCH_NO = BILL_NO
                                rststock!LINE_NO = 1
                                rststock!ITEM_CODE = RSTITEMMAST!ITEM_CODE
                            'End If
                            rststock!BAL_QTY = Val(TXTsample.Text) - (Val(BAL_QTY))
                            rststock!QTY = Val(TXTsample.Text) - (Val(INWARD - OUTWARD))
                            rststock!TRX_TOTAL = 0
                            rststock!VCH_DATE = Format(DTFROM.Value, "dd/mm/yyyy")
                            rststock!ITEM_NAME = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2))
                            rststock!item_COST = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11))
                            rststock!LINE_DISC = 1
                            rststock!P_DISC = 0
                            rststock!MRP = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6))
                            rststock!PTR = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11))
                            rststock!SALES_PRICE = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
                            rststock!P_RETAIL = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
                            'rststock!P_WS = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
                            'rststock!P_VAN = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9))
                            rststock!P_CRTN = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16))
                            rststock!P_LWS = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 17))
                            rststock!CRTN_PACK = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15))
                            rststock!Category = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 18))
                            rststock!gross_amt = 0
                            rststock!COM_FLAG = "P"
                            rststock!COM_PER = 0
                            rststock!COM_AMT = 0
                            rststock!SALES_TAX = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10))
                            rststock!LOOSE_PACK = RSTITEMMAST!LOOSE_PACK
                            rststock!PACK_TYPE = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5))
                            rststock!WARRANTY = Null
                            rststock!WARRANTY_TYPE = Null
                            rststock!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
                            'rststock!VCH_DESC = "Received From " & DataList2.Text
                            rststock!REF_NO = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9))
                            'rststock!ISSUE_QTY = 0
                            rststock!CST = 0
                            rststock!DISC_FLAG = "P"
                            rststock!SCHEME = 0
                            'rststock!EXP_DATE = Null
                            rststock!FREE_QTY = 0
                            rststock!CREATE_DATE = Format(Date, "dd/mm/yyyy")
                            rststock!C_USER_ID = "SM"
                            rststock!check_flag = "V"
                            If IsDate(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) Then
                                If Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 1, 2)) <> 0 And Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 4, 5)) <= 12 And Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 1, 2)) > 0 And Val(Mid(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), 4, 5)) > 0 Then
                                    rststock!EXP_DATE = Format(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8), "dd/mm/yyyy")
                                End If
                            End If
                            'rststock!M_USER_ID = DataList2.BoundText
                            'rststock!PINV = Trim(TXTINVOICE.Text)
                            rststock.Update
                            rststock.Close
                            Set rststock = Nothing
                            
                            RSTITEMMAST!CLOSE_QTY = Val(TXTsample.Text)
                            RSTITEMMAST!RCPT_QTY = INWARD + Val(TXTsample.Text)
                            RSTITEMMAST!ISSUE_QTY = OUTWARD
                            RSTITEMMAST.Update
                        End If
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                    End If
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8) = ""
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9) = ""
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 24) = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3))
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 4) = Round(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)) / GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13), 0)
                    '=========
                    Set RSTITEMMAST = New ADODB.Recordset
                    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                        BALQTY = 0
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT SUM(BAL_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY <> 0", db, adOpenForwardOnly
                        If Not (rststock.EOF And rststock.BOF) Then
                            BALQTY = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        If Round(BALQTY, 2) = Round(RSTITEMMAST!CLOSE_QTY, 2) Then GoTo SKIP_BALCHECK
                        
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
                        
            '            Set rststock = New ADODB.Recordset
            '            rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
            '            Do Until rststock.EOF
            '                INWARD = INWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
            '                INWARD = INWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
            '    '            If IsNull(rststock!Category) Then
            '    '                MsgBox "1"
            '    '            End If
            '    '            If IsNull(RSTITEMMAST!Category) Then
            '    '                MsgBox "2"
            '    '            End If
            '                'rststock!Category = RSTITEMMAST!Category
            '                'rststock.Update
            '                rststock.MoveNext
            '            Loop
            '            rststock.Close
            '            Set rststock = Nothing
                        
                        i = i + 1
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                            OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                '            If IsNull(rststock!Category) Then
                '                MsgBox "3"
                '            End If
                '            If IsNull(RSTITEMMAST!Category) Then
                '                MsgBox "4"
                '            End If
                            'rststock!Category = RSTITEMMAST!Category
                            'rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                        
                        
                        db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY <0"
                        If Round(INWARD - OUTWARD, 2) = 0 Then
                            db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY >0"
                        End If
                        
                        
                        'If INWARD - OUTWARD <> BALQTY Then MsgBox RSTITEMMAST!ITEM_CODE
                        
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
                                    rststock!BAL_QTY = IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) - DIFFQTY
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
                                    rststock!BAL_QTY = IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) + DIFFQTY
                                    DIFFQTY = 0
                                Else
                                    If Not rststock!BAL_QTY = IIf(IsNull(rststock!QTY), 0, rststock!QTY) Then
                                        rststock!BAL_QTY = IIf(IsNull(rststock!QTY), 0, rststock!QTY)
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
                        
                        RSTITEMMAST!CLOSE_QTY = Round(INWARD - OUTWARD, 2)
                        RSTITEMMAST!RCPT_QTY = INWARD
                        RSTITEMMAST!ISSUE_QTY = OUTWARD
                        RSTITEMMAST.Update
                    End If
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
SKIP_BALCHECK:
                    '==========
                    Call Toatal_value
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    Screen.MousePointer = vbNormal
                    
                Case 7  'RT
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_RETAIL = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12)) <> 0 Then
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 20) = Format(Round((((Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) / GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) - Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12))) * 100) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12)), 2), "0.00")
                        Else
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 20) = 0
                        End If
                        
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15)) = 0 Then
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = 1
                            rststock!CRTN_PACK = 1
                        End If
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) = 0 Then
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = 1
                            rststock!LOOSE_PACK = 1
                        End If
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!P_RETAIL = Val(TXTsample.Text)
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15)) = 0 Then
                            rststock!CRTN_PACK = 1
                        End If
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) = 0 Then
                            rststock!LOOSE_PACK = 1
                        End If
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 8  'WS
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_WS = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15)) = 0 Then
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = 1
                            rststock!CRTN_PACK = 1
                        End If
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) = 0 Then
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = 1
                            rststock!LOOSE_PACK = 1
                        End If
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!P_WS = Val(TXTsample.Text)
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15)) = 0 Then
                            rststock!CRTN_PACK = 1
                        End If
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) = 0 Then
                            rststock!LOOSE_PACK = 1
                        End If
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 15  'CRTN_PACK
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!CRTN_PACK = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Val(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!CRTN_PACK = Val(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 16  'L. R. PRICE
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_CRTN = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                        If GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = 0 Then
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = 1
                            rststock!CRTN_PACK = 1
                        End If
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!P_CRTN = Val(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 17  'L. W. PRICE
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_LWS = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                        If GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = 0 Then
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = 1
                            rststock!CRTN_PACK = 1
                        End If
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!P_LWS = Val(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 9  'batch
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (rststock.EOF And rststock.BOF) Then
'                        rststock!P_VAN = Val(TXTsample.Text)
'                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
'                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15)) = 0 Then
'                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = 1
'                            rststock!CRTN_PACK = 1
'                        End If
'                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) = 0 Then
'                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = 1
'                            rststock!LOOSE_PACK = 1
'                        End If
'                        rststock.Update
'                    End If
'                    rststock.Close
'                    Set rststock = Nothing
'
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                    Do Until rststock.EOF
'                        rststock!P_WS = Val(TXTsample.Text)
'                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15)) = 0 Then
'                            rststock!CRTN_PACK = 1
'                        End If
'                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) = 0 Then
'                            rststock!LOOSE_PACK = 1
'                        End If
'                        rststock.Update
'                        rststock.MoveNext
'                    Loop
'                    rststock.Close
'                    Set rststock = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 18  'CATEGORY
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!Category = Trim(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 27  'item spec
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_SPEC = Trim(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 28  'REMARKS
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!BIN_LOCATION = Trim(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
'                Case 11  'LOC
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (rststock.EOF And rststock.BOF) Then
'                        rststock!BIN_LOCATION = Trim(TXTsample.Text)
'                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
'                        rststock.Update
'                    End If
'                    rststock.Close
'                    Set rststock = Nothing
'                    GRDSTOCK.Enabled = True
'                    TXTsample.Visible = False
'                    GRDSTOCK.SetFocus
                
                Case 11  'COST
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!item_COST = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) / 100)
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12)) <> 0 Then
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 20) = Format(Round(((Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) - Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12))) * 100) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12)), 2), "0.00")
                        Else
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 20) = 0
                        End If
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 24) = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3))
                    Call Toatal_value
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 24  'VALUE
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)) <> 0 Then
                            rststock!item_COST = Round(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), 3)
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11) = Format(Round(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), 3), "0.000")
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) / 100)
                        End If
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) <> 0 Then
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 20) = Format(Round(((Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) - Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11))) * 100) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)), 2), "0.00")
                        Else
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 20) = 0
                        End If
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 24) = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3))
                    Call Toatal_value
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 6  'MRP
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND (MRP =0 OR ISNULL(MRP)) ", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!MRP = Val(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 10  'TAX
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!SALES_TAX = Val(TXTsample.Text)
                        rststock!check_flag = "V"
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Val(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) / 100)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 20  'Profit %
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7) = Format(Round(((Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12)) * GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) * Val(TXTsample.Text) / 100) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12)) * GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)), 2), "0.000")
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_RETAIL = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 21  'Cust Disc
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!CUST_DISC = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 14  'BARCODE
                    If Trim(TXTsample.Text) = "" Then
                        TXTsample.Text = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)) & Int(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)))
                        If BARTEMPLATE = "Y" And Len(TXTsample.Text) Mod 2 <> 0 Then TXTsample.Text = TXTsample.Text & "9"
                    End If
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!BARCODE = Trim(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    db.Execute "Update RTRXFILE Set BARCODE ='" & Trim(TXTsample.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' and (isnull(BARCODE) or BARCODE ='')"
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 25  'CESS%
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!CESS_PER = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 26  'CESS RATE
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!cess_amt = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 13  'UNIT
                    If Val(TXTsample.Text) <= 0 Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!LOOSE_PACK = Val(TXTsample.Text)
                        
                        rststock!P_CRTN = Round(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) / Val(TXTsample.Text), 2)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) = Format(Round(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) / Val(TXTsample.Text), 2), "0.000")
                        
                        rststock!P_LWS = Round(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) / Val(TXTsample.Text), 2)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 17) = Format(Round(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) / Val(TXTsample.Text), 2), "0.000")

                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
eRRhAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 3, 6, 7, 8, 10, 11, 13, 15, 16, 17, 20, 21, 25, 26
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        
        Case 1, 2, 14, 9
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub TXTCODE_Change()
    On Error GoTo eRRhAND
    If Trim(tXTMEDICINE.Text) <> "" Then Call Fillgrid
    If REPFLAG = True Then
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.Text)
    'Call Fillgrid
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'tXTMEDICINE.SetFocus
            Call CmdLoad_Click
        Case vbKeyEscape
            Call CmdExit_Click
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

Private Sub TxtComper_GotFocus()
    TxtComper.SelStart = 0
    TxtComper.SelLength = Len(TxtComper.Text)
End Sub

Private Sub TxtComper_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            cmdOK_Click
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub TxtComper_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case 65, 97
            OptAmt.Value = True
            KeyAscii = 0
        Case 112, 80
            OptPercent.Value = True
            KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtComper_LostFocus()
    TxtComper.Text = Format(TxtComper.Text, "0.00")
End Sub

Private Sub OptPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            TxtComper.SetFocus
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub OptAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
             TxtComper.SetFocus
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub


Private Sub cmdOK_Click()
    Dim rststock As ADODB.Recordset
    
    If Not IsNumeric(TxtComper.Text) Then
        MsgBox " Enter proper value", vbOKOnly, "Commission !!!"
        TxtComper.SetFocus
        Exit Sub
    End If
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (rststock.EOF And rststock.BOF) Then
        If Val(TxtComper.Text) = 0 Then
            rststock!COM_FLAG = ""
            rststock!COM_PER = 0
            rststock!COM_AMT = 0
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 20) = "0.00"
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 21) = ""
        Else
            If OptAmt.Value = True Then
                rststock!COM_FLAG = "A"
                rststock!COM_PER = 0
                rststock!COM_AMT = Val(TxtComper.Text)
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 20) = Format(Val(TxtComper.Text), "0.00")
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 21) = "Rs"
            Else
                rststock!COM_FLAG = "P"
                rststock!COM_PER = Val(TxtComper.Text)
                rststock!COM_AMT = 0
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 20) = Format(Val(TxtComper.Text), "0.00")
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 21) = "%"
            End If
        End If
        rststock.Update
    End If
    rststock.Close
    Set rststock = Nothing
    GRDSTOCK.Enabled = True
    FRAME.Visible = False
    GRDSTOCK.SetFocus
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub cmdcancel_Click()
    FRAME.Visible = False
    GRDSTOCK.SetFocus
End Sub

Private Sub CmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Function Toatal_value()
    Dim Stk_Val As Double
    Dim i As Integer
    lblpvalue.Caption = ""
    lblnetvalue.Caption = ""
    For i = 1 To GRDSTOCK.rows - 1
        lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 24)), "0.00")
        lblnetvalue.Caption = Format(Val(lblnetvalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 12)) * Val(GRDSTOCK.TextMatrix(i, 3))), "0.00")
    Next i
End Function

Private Sub TXTDEALER2_Change()
    
    On Error GoTo eRRhAND
    If flagchange2.Caption <> "1" Then
        If chkcategory.Value = 1 Then
            If PHY_FLAG = True Then
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
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
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
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
eRRhAND:
    MsgBox err.Description
    
End Sub


Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.Text)
    'CHKCATEGORY2.value = 1
End Sub

Private Sub TXTDEALER2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
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
        
    TXTDEALER2.Text = DataList1.Text
    LBLDEALER2.Caption = TXTDEALER2.Text
    Call Fillgrid
    tXTMEDICINE.SetFocus
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Trim(TXTDEALER2.Text) = "" Then Exit Sub
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
    flagchange2.Caption = 1
    TXTDEALER2.Text = LBLDEALER2.Caption
    DataList1.Text = TXTDEALER2.Text
    Call DataList1_Click
    'CHKCATEGORY2.value = 1
End Sub

Private Sub DataList1_LostFocus()
     flagchange2.Caption = ""
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

