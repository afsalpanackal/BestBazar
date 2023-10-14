VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPriceAnalysis 
   BackColor       =   &H00E8DFEC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price Analysis"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   20430
   ClipControls    =   0   'False
   Icon            =   "FrmPriceAnalysis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   20430
   Begin VB.Frame Frame7 
      BackColor       =   &H00BFDFEC&
      Height          =   450
      Left            =   6480
      TabIndex        =   101
      Top             =   -75
      Width           =   5250
      Begin VB.OptionButton OptAllStock 
         BackColor       =   &H00BFDFEC&
         Caption         =   "All"
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
         Left            =   3645
         TabIndex        =   104
         Top             =   150
         Width           =   1395
      End
      Begin VB.OptionButton OptBrStock 
         BackColor       =   &H00BFDFEC&
         Caption         =   "Branch Stock"
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
         Left            =   1755
         TabIndex        =   103
         Top             =   150
         Width           =   1605
      End
      Begin VB.OptionButton OptMainStk 
         BackColor       =   &H00BFDFEC&
         Caption         =   "Main Stock"
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
         Left            =   30
         TabIndex        =   102
         Top             =   150
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Import Qty"
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
      Left            =   19290
      TabIndex        =   100
      Top             =   5535
      Width           =   1080
   End
   Begin VB.TextBox txtstkcrct 
      Alignment       =   2  'Center
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
      Left            =   18480
      MaxLength       =   2
      TabIndex        =   95
      Top             =   60
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton CmdStkCrct 
      Caption         =   "Stock Correction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10515
      TabIndex        =   94
      Top             =   2190
      Width           =   1035
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
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
      Left            =   10515
      TabIndex        =   93
      Top             =   1770
      Width           =   1020
   End
   Begin VB.TextBox TxtName 
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
      Left            =   3000
      TabIndex        =   2
      Top             =   270
      Width           =   690
   End
   Begin VB.CommandButton cmdchangeunbill 
      Caption         =   "Change all selected items with / without"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   19185
      TabIndex        =   79
      Top             =   435
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Assign Selling Rates"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   17925
      TabIndex        =   67
      Top             =   435
      Width           =   1245
   End
   Begin VB.CommandButton cmdPriceChange 
      Caption         =   "Assign Price Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   16635
      TabIndex        =   70
      Top             =   435
      Width           =   1245
   End
   Begin VB.ComboBox CmbPrChange 
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
      ItemData        =   "FrmPriceAnalysis.frx":030A
      Left            =   16620
      List            =   "FrmPriceAnalysis.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   69
      Top             =   45
      Width           =   1245
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Assign Cust Disc to all"
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
      Left            =   14370
      TabIndex        =   66
      Top             =   945
      Width           =   1050
   End
   Begin VB.TextBox txtcustdic 
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
      Left            =   15450
      TabIndex        =   65
      Top             =   930
      Width           =   1110
   End
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
      Left            =   8220
      TabIndex        =   64
      Top             =   1770
      Width           =   1110
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
      Height          =   435
      Left            =   15450
      TabIndex        =   63
      Top             =   1365
      Width           =   1125
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
      Height          =   435
      Left            =   14370
      TabIndex        =   62
      Top             =   1380
      Width           =   1050
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
      Left            =   9375
      TabIndex        =   61
      Top             =   1770
      Width           =   1125
   End
   Begin VB.Frame Frame4 
      Height          =   1425
      Left            =   8670
      TabIndex        =   55
      Top             =   300
      Width           =   3060
      Begin VB.CheckBox ChkLR 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Hide LR Price"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1740
         TabIndex        =   85
         Top             =   1155
         Width           =   1290
      End
      Begin VB.CheckBox ChkUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Hide Unit"
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
         Height          =   240
         Left            =   1740
         TabIndex        =   84
         Top             =   945
         Width           =   1290
      End
      Begin VB.CheckBox ChkNetCost 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Hide NetCost"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1740
         TabIndex        =   83
         Top             =   735
         Width           =   1290
      End
      Begin VB.CheckBox ChkCost 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Hide Cost"
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
         Left            =   1740
         TabIndex        =   82
         Top             =   525
         Width           =   1290
      End
      Begin VB.CheckBox ChkMRP 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Hide MRP"
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
         Height          =   240
         Left            =   1740
         TabIndex        =   81
         Top             =   315
         Width           =   1290
      End
      Begin VB.CheckBox ChkUOM 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Hide UOM"
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
         Height          =   240
         Left            =   1740
         TabIndex        =   80
         Top             =   105
         Width           =   1290
      End
      Begin VB.CheckBox chkdeaditems 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide Dead Stock"
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
         TabIndex        =   75
         Top             =   1155
         Value           =   1  'Checked
         Width           =   1740
      End
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
         TabIndex        =   60
         Top             =   945
         Width           =   1740
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
         TabIndex        =   59
         Top             =   735
         Width           =   1740
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
         TabIndex        =   58
         Top             =   525
         Width           =   1740
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
         TabIndex        =   57
         Top             =   315
         Width           =   1740
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
         TabIndex        =   56
         Top             =   105
         Width           =   1740
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
      Left            =   5580
      TabIndex        =   5
      Top             =   270
      Width           =   870
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
      Left            =   5115
      TabIndex        =   4
      Top             =   270
      Width           =   450
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
      Left            =   3705
      TabIndex        =   3
      Top             =   270
      Width           =   1395
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
      Left            =   15450
      TabIndex        =   48
      Top             =   1830
      Width           =   1125
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
      Left            =   14370
      TabIndex        =   47
      Top             =   1830
      Width           =   1050
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
      Left            =   14370
      TabIndex        =   45
      Top             =   495
      Width           =   1050
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
      Left            =   15450
      TabIndex        =   44
      Top             =   480
      Width           =   1110
   End
   Begin VB.Frame Frame 
      Height          =   2190
      Left            =   1920
      TabIndex        =   35
      Top             =   3900
      Visible         =   0   'False
      Width           =   3945
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   2640
         TabIndex        =   42
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   1335
         TabIndex        =   41
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
         TabIndex        =   36
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   40
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
      Height          =   240
      Left            =   13065
      TabIndex        =   33
      Top             =   15
      Width           =   1305
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
      Left            =   11760
      TabIndex        =   29
      Top             =   255
      Width           =   2565
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
      Left            =   11760
      TabIndex        =   28
      Top             =   15
      Width           =   1620
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
      Left            =   15450
      TabIndex        =   27
      Top             =   45
      Width           =   1110
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
      Left            =   14370
      TabIndex        =   26
      Top             =   45
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   19290
      Top             =   4785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   435
      Left            =   19305
      TabIndex        =   24
      Top             =   1800
      Width           =   1140
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
      Height          =   435
      Left            =   19305
      TabIndex        =   23
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      Height          =   2415
      Left            =   19305
      TabIndex        =   22
      Top             =   2250
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
   Begin VB.CommandButton CmdLoad 
      Caption         =   "&Re- Load"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6465
      TabIndex        =   6
      Top             =   1215
      Width           =   1125
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
      Left            =   1950
      TabIndex        =   1
      Top             =   270
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   900
      Left            =   6480
      TabIndex        =   9
      Top             =   285
      Width           =   2190
      Begin VB.OptionButton OptPC 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Price Changing Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   30
         TabIndex        =   43
         Top             =   630
         Width           =   2115
      End
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Display All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   30
         TabIndex        =   11
         Top             =   165
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptStock 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Stock Items Only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   30
         TabIndex        =   10
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
      Width           =   1890
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
      Height          =   390
      Left            =   7620
      TabIndex        =   7
      Top             =   1215
      Width           =   1035
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   45
      TabIndex        =   8
      Top             =   630
      Width           =   6405
      _ExtentX        =   11298
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
   Begin MSComCtl2.DTPicker DTFROM 
      Height          =   360
      Left            =   6450
      TabIndex        =   20
      Top             =   1830
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   635
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
      Format          =   112394241
      CurrentDate     =   40498
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   780
      Left            =   11760
      TabIndex        =   30
      Top             =   600
      Width           =   2565
      _ExtentX        =   4524
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
   Begin VB.Frame frmunbill 
      BackColor       =   &H00FFC0C0&
      Height          =   630
      Left            =   16620
      TabIndex        =   76
      Top             =   780
      Visible         =   0   'False
      Width           =   2565
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
         TabIndex        =   78
         Top             =   375
         Visible         =   0   'False
         Width           =   2385
      End
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
         TabIndex        =   77
         Top             =   120
         Width           =   1875
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Height          =   1305
      Left            =   16620
      TabIndex        =   71
      Top             =   1320
      Width           =   2565
      Begin VB.OptionButton OptItemcode 
         BackColor       =   &H00FFC0C0&
         Caption         =   "By Item Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1020
         TabIndex        =   109
         Top             =   1050
         Width           =   1485
      End
      Begin VB.OptionButton OptSortTax 
         BackColor       =   &H00FFC0C0&
         Caption         =   "By &Tax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   45
         TabIndex        =   108
         Top             =   1035
         Width           =   1005
      End
      Begin VB.OptionButton OptSortHSN 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by &HSN Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   45
         TabIndex        =   107
         Top             =   795
         Width           =   2340
      End
      Begin VB.OptionButton OptSortQty 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by &Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   45
         TabIndex        =   74
         Top             =   345
         Width           =   1935
      End
      Begin VB.OptionButton OptSortName 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by &Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   45
         TabIndex        =   73
         Top             =   120
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.OptionButton OptSortPrice 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by &Cost"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   45
         TabIndex        =   72
         Top             =   555
         Width           =   2340
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   6030
      Left            =   45
      TabIndex        =   12
      Top             =   2520
      Width           =   19260
      Begin VB.Frame Frmebatch 
         Caption         =   "`"
         Height          =   4620
         Left            =   4845
         TabIndex        =   86
         Top             =   810
         Visible         =   0   'False
         Width           =   9885
         Begin VB.TextBox TXTEDIT 
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
            Left            =   2010
            TabIndex        =   87
            Top             =   1350
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSMask.MaskEdBox TXTEXP 
            Height          =   360
            Left            =   0
            TabIndex        =   88
            Top             =   585
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   635
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
         Begin MSFlexGridLib.MSFlexGrid grdbatch 
            Height          =   4170
            Left            =   15
            TabIndex        =   89
            Top             =   435
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   7355
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   450
            BackColor       =   14995141
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
         Begin VB.Label LBLQTY 
            Alignment       =   1  'Right Justify
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
            Left            =   5895
            TabIndex        =   92
            Top             =   135
            Width           =   2070
         End
         Begin VB.Label LBLITEMCODE 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   105
            TabIndex        =   91
            Top             =   135
            Width           =   1725
         End
         Begin VB.Label LblItem 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   1935
            TabIndex        =   90
            Top             =   120
            Width           =   3375
         End
      End
      Begin VB.ComboBox CMBCHANGE 
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
         ItemData        =   "FrmPriceAnalysis.frx":0321
         Left            =   3770
         List            =   "FrmPriceAnalysis.frx":032B
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   3555
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSDataListLib.DataCombo Cmbcategory 
         Height          =   360
         Left            =   9645
         TabIndex        =   46
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
         TabIndex        =   17
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
         TabIndex        =   15
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
         ItemData        =   "FrmPriceAnalysis.frx":0338
         Left            =   2385
         List            =   "FrmPriceAnalysis.frx":038D
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   795
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
         Height          =   5940
         Left            =   15
         TabIndex        =   13
         Top             =   105
         Width           =   19200
         _ExtentX        =   33867
         _ExtentY        =   10478
         _Version        =   393216
         Rows            =   1
         Cols            =   31
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
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
      Begin VB.Image Image2 
         Height          =   15
         Left            =   4395
         Top             =   4380
         Width           =   15
      End
   End
   Begin VB.Label lblstock 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   45
      TabIndex        =   106
      Top             =   2265
      Width           =   6405
   End
   Begin VB.Label lblstktype 
      Height          =   135
      Left            =   5070
      TabIndex        =   105
      Top             =   2265
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Value"
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
      Index           =   7
      Left            =   11655
      TabIndex        =   99
      Top             =   2040
      Width           =   1185
   End
   Begin VB.Label lblsalevalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   12720
      TabIndex        =   98
      Top             =   2055
      Width           =   1620
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Index           =   44
      Left            =   19230
      TabIndex        =   97
      Top             =   60
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
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
      Index           =   43
      Left            =   17910
      TabIndex        =   96
      Top             =   75
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   5
      Left            =   5580
      TabIndex        =   54
      Top             =   15
      Width           =   870
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
      Left            =   5115
      TabIndex        =   53
      Top             =   15
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Barcode/Code"
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
      Left            =   3720
      TabIndex        =   52
      Top             =   15
      Width           =   1365
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
      TabIndex        =   51
      Top             =   15
      Width           =   3645
   End
   Begin VB.Label lblnetvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   12720
      TabIndex        =   50
      Top             =   1725
      Width           =   1620
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
      Left            =   11745
      TabIndex        =   49
      Top             =   1740
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
      Height          =   330
      Left            =   6465
      TabIndex        =   34
      Top             =   2190
      Width           =   4005
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   0
      TabIndex        =   32
      Top             =   810
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   0
      TabIndex        =   31
      Top             =   450
      Width           =   495
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
      Height          =   300
      Index           =   34
      Left            =   19425
      TabIndex        =   25
      Top             =   3660
      Width           =   1455
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
      Left            =   6375
      TabIndex        =   21
      Top             =   1605
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tot Value"
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
      Left            =   11775
      TabIndex        =   19
      Top             =   1425
      Width           =   1500
   End
   Begin VB.Label lblpvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   12720
      TabIndex        =   18
      Top             =   1395
      Width           =   1620
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
      TabIndex        =   16
      Top             =   660
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "FrmPriceAnalysis"
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
Dim frmloadflag As Boolean

Private Sub CHKCATEGORY_Click()
    CHKCATEGORY2.Value = 0
End Sub

Private Sub CHKCATEGORY2_Click()
    chkcategory.Value = 0
End Sub

Private Sub ChkCost_Click()
    If frmLogin.rs!Level = "0" Then
        If ChkCost.Value = 1 Then
            GRDSTOCK.TextMatrix(0, 11) = ""
            GRDSTOCK.ColWidth(11) = 0
        Else
            GRDSTOCK.TextMatrix(0, 11) = "Per Rate"
            GRDSTOCK.ColWidth(11) = 1000
        End If
    End If
End Sub

Private Sub chkhidecat_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo ERRHAND
    If chkhidecat.Value = 1 Then
        db.Execute "Update COMPINFO set hide_category = 'Y' where COMP_CODE = '001' "
        GRDSTOCK.TextMatrix(0, 18) = ""
        GRDSTOCK.ColWidth(18) = 0
    Else
        db.Execute "Update COMPINFO set hide_category = 'N' where COMP_CODE = '001' "
        GRDSTOCK.TextMatrix(0, 18) = "Category"
        GRDSTOCK.ColWidth(18) = 1000
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
        GRDSTOCK.TextMatrix(0, 19) = ""
        GRDSTOCK.ColWidth(19) = 0
    Else
        db.Execute "Update COMPINFO set hide_company = 'N' where COMP_CODE = '001' "
        GRDSTOCK.TextMatrix(0, 19) = "Company"
        GRDSTOCK.ColWidth(19) = 1000
    End If
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub ChkLR_Click()
    If ChkLR.Value = 1 Then
        GRDSTOCK.TextMatrix(0, 16) = ""
        GRDSTOCK.ColWidth(16) = 0
    Else
        GRDSTOCK.TextMatrix(0, 16) = "L.R.Price"
        GRDSTOCK.ColWidth(16) = 1000
    End If
End Sub

Private Sub chklwp_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo ERRHAND
    If chklwp.Value = 1 Then
        db.Execute "Update COMPINFO set hide_lwp = 'Y' where COMP_CODE = '001' "
        GRDSTOCK.TextMatrix(0, 17) = ""
        GRDSTOCK.ColWidth(17) = 0
    Else
        db.Execute "Update COMPINFO set hide_lwp = 'N' where COMP_CODE = '001' "
        GRDSTOCK.TextMatrix(0, 17) = "L.W.Price"
        GRDSTOCK.ColWidth(17) = 1000
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub ChkMRP_Click()
    If chkmrp.Value = 1 Then
        GRDSTOCK.TextMatrix(0, 6) = ""
        GRDSTOCK.ColWidth(6) = 0
    Else
        GRDSTOCK.TextMatrix(0, 6) = "MRP"
        GRDSTOCK.ColWidth(6) = 1000
    End If
End Sub

Private Sub ChkNetCost_Click()
    If frmLogin.rs!Level = "0" Then
        If ChkNetCost.Value = 1 Then
            GRDSTOCK.TextMatrix(0, 12) = ""
            GRDSTOCK.ColWidth(12) = 0
        Else
            GRDSTOCK.TextMatrix(0, 12) = "Net Cost"
            GRDSTOCK.ColWidth(12) = 900
        End If
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

Private Sub ChkUnit_Click()
    If ChkUnit.Value = 1 Then
        GRDSTOCK.TextMatrix(0, 13) = ""
        GRDSTOCK.ColWidth(13) = 0
    Else
        GRDSTOCK.TextMatrix(0, 13) = "Unit"
        GRDSTOCK.ColWidth(13) = 800
    End If
End Sub

Private Sub ChkUOM_Click()
    If ChkUOM.Value = 1 Then
        GRDSTOCK.TextMatrix(0, 5) = ""
        GRDSTOCK.ColWidth(5) = 0
    Else
        GRDSTOCK.TextMatrix(0, 5) = "UOM"
        GRDSTOCK.ColWidth(5) = 900
    End If
End Sub

Private Sub chkvp_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo ERRHAND
    If chkvp.Value = 1 Then
        db.Execute "Update COMPINFO set hide_van = 'Y' where COMP_CODE = '001' "
        GRDSTOCK.TextMatrix(0, 9) = ""
        GRDSTOCK.ColWidth(9) = 0
    Else
        db.Execute "Update COMPINFO set hide_van = 'N' where COMP_CODE = '001' "
        GRDSTOCK.TextMatrix(0, 9) = "VP"
        GRDSTOCK.ColWidth(9) = 1000
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
        GRDSTOCK.TextMatrix(0, 8) = ""
        GRDSTOCK.ColWidth(8) = 0
    Else
        db.Execute "Update COMPINFO set hide_ws = 'N' where COMP_CODE = '001' "
        GRDSTOCK.TextMatrix(0, 8) = "WS"
        GRDSTOCK.ColWidth(8) = 1000
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Cmbcategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 18  'CATEGORY
                    If Cmbcategory.Text = "" Then
                        MsgBox "Please select Category from the List", vbOKOnly, "EzBiz"
                        Exit Sub
                    End If
                    
                    db.Execute "Update ITEMMAST set Category = '" & Cmbcategory.Text & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    db.Execute "Update trxfile set Category = '" & Cmbcategory.Text & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    db.Execute "Update rtrxfile set Category = '" & Cmbcategory.Text & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Cmbcategory.Text
                    
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
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CMBCHANGE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 28  'pack
                    If CMBCHANGE.ListIndex = -1 Then CMBCHANGE.ListIndex = 0
                    Select Case CMBCHANGE.ListIndex
                        Case 0
                            
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT DISTINCT P_RETAIL from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_RETAIL) AND P_RETAIL <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                            If rststock.RecordCount > 1 Then
                                If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                                    rststock.Close
                                    Set rststock = Nothing
                                    CMBCHANGE.Visible = False
                                    GRDSTOCK.SetFocus
                                    Exit Sub
                                End If
                            End If
                            rststock.Close
                            Set rststock = Nothing
                            
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                            If rststock.RecordCount > 1 Then
                                If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                                    rststock.Close
                                    Set rststock = Nothing
                                    CMBCHANGE.Visible = False
                                    GRDSTOCK.SetFocus
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
                                    CMBCHANGE.Visible = False
                                    GRDSTOCK.SetFocus
                                    Exit Sub
                                End If
                            End If
                            rststock.Close
                            Set rststock = Nothing
                            
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rststock.EOF And rststock.BOF) Then
                                rststock!P_RETAIL = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
                                rststock!P_WS = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
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
                                If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) = 0 Then
                                    rststock!P_CRTN = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
                                    rststock!P_LWS = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
                                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)), "0.000")
                                Else
                                    rststock!P_CRTN = Round(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)), 2)
                                    rststock!P_LWS = Round(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)), 2)
                                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)), "0.000")
                                End If
                            
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                            
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                            Do Until rststock.EOF
                                rststock!P_RETAIL = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
                                rststock!P_WS = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
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
                            db.Execute "Update ITEMMAST set PRICE_CHANGE = 'Y' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' "
                        Case 1
                            db.Execute "Update ITEMMAST set PRICE_CHANGE = 'N' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' "
                    End Select
                    GRDSTOCK.Enabled = True
                    CMBCHANGE.Visible = False
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = CMBCHANGE.Text
                    GRDSTOCK.SetFocus
                Case 29  'UN_BILL
                    If CMBCHANGE.ListIndex = -1 Then CMBCHANGE.ListIndex = 0
                    Select Case CMBCHANGE.ListIndex
                        Case 0
                            db.Execute "Update ITEMMAST set UN_BILL = 'Y' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' "
                        Case 1
                            db.Execute "Update ITEMMAST set UN_BILL = 'N' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' "
                    End Select
                    GRDSTOCK.Enabled = True
                    CMBCHANGE.Visible = False
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = CMBCHANGE.Text
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            CMBCHANGE.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CMBMFGR_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 19  'MFGR
                    If CMBMFGR.Text = "" Then
                        MsgBox "Please select Company from the List", vbOKOnly, "EzBiz"
                        Exit Sub
                    End If
                    Set rststock = New ADODB.Recordset
                    
                    db.Execute "Update ITEMMAST set MANUFACTURER = '" & CMBMFGR.Text & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    db.Execute "Update trxfile set MFGR = '" & CMBMFGR.Text & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    db.Execute "Update rtrxfile set MFGR = '" & CMBMFGR.Text & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = CMBMFGR.Text
                    
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
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
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
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdchangeunbill_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    Dim i As Long
    On Error GoTo ERRHAND
    If CmbPrChange.ListIndex = -1 Then Exit Sub
    Select Case CmbPrChange.ListIndex
        Case 0
            If MsgBox("ARE YOU SURE YOU WANT TO CHANGE THE SELECTED ITEMS AS UN BILL", vbYesNo + vbDefaultButton2, "Price Analysis") = vbNo Then Exit Sub
        Case Else
            If MsgBox("ARE YOU SURE YOU WANT TO CHANGE THE SELECTED ITEMS AS BILLED", vbYesNo + vbDefaultButton2, "Price Analysis") = vbNo Then Exit Sub
    End Select
    For i = 1 To GRDSTOCK.rows - 1
        Select Case CmbPrChange.ListIndex
            Case 0
                db.Execute "Update ITEMMAST set UN_BILL = 'Y' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
                GRDSTOCK.TextMatrix(i, 29) = "Yes"
            Case Else
                db.Execute "Update ITEMMAST set UN_BILL = 'N' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
                GRDSTOCK.TextMatrix(i, 29) = "No"
        End Select
    Next i
    MsgBox "Successfully Applied", , "Price Analysis"
    CmbPrChange.ListIndex = -1
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdDelete_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If MDIMAIN.StatusBar.Panels(9).Text = "Y" Then Exit Sub
    Dim rststock As ADODB.Recordset
    
    If GRDSTOCK.rows <= 1 Then Exit Sub
    On Error GoTo ERRHAND
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
    rststock.Open "SELECT * from RTRXFILEVAN where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " Since Transactions is Available in Branch Sales", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFILEVAN where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " Since Transactions is Available in Branch Sales", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULASUB where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
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
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULAMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " ****", vbYesNo + vbDefaultButton2, "DELETING ITEM....") = vbNo Then
        GRDSTOCK.SetFocus
        Exit Sub
    End If
    'db.Execute ("DELETE from RTRXFILE where RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'")
    db.Execute ("DELETE from PRODLINK where PRODLINK.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'")
    db.Execute ("DELETE from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'")
    
    Dim selrow As Integer
    Dim i As Long
    selrow = GRDSTOCK.Row
    For i = selrow To GRDSTOCK.rows - 2
        GRDSTOCK.TextMatrix(selrow, 0) = i
        GRDSTOCK.TextMatrix(selrow, 1) = GRDSTOCK.TextMatrix(i + 1, 1)
        GRDSTOCK.TextMatrix(selrow, 2) = GRDSTOCK.TextMatrix(i + 1, 2)
        GRDSTOCK.TextMatrix(selrow, 3) = GRDSTOCK.TextMatrix(i + 1, 3)
        GRDSTOCK.TextMatrix(selrow, 4) = GRDSTOCK.TextMatrix(i + 1, 4)
        GRDSTOCK.TextMatrix(selrow, 5) = GRDSTOCK.TextMatrix(i + 1, 5)
        GRDSTOCK.TextMatrix(selrow, 6) = GRDSTOCK.TextMatrix(i + 1, 6)
        GRDSTOCK.TextMatrix(selrow, 7) = GRDSTOCK.TextMatrix(i + 1, 7)
        GRDSTOCK.TextMatrix(selrow, 8) = GRDSTOCK.TextMatrix(i + 1, 8)
        GRDSTOCK.TextMatrix(selrow, 9) = GRDSTOCK.TextMatrix(i + 1, 9)
        GRDSTOCK.TextMatrix(selrow, 10) = GRDSTOCK.TextMatrix(i + 1, 10)
        GRDSTOCK.TextMatrix(selrow, 11) = GRDSTOCK.TextMatrix(i + 1, 11)
        GRDSTOCK.TextMatrix(selrow, 12) = GRDSTOCK.TextMatrix(i + 1, 12)
        GRDSTOCK.TextMatrix(selrow, 13) = GRDSTOCK.TextMatrix(i + 1, 13)
        GRDSTOCK.TextMatrix(selrow, 14) = GRDSTOCK.TextMatrix(i + 1, 14)
        GRDSTOCK.TextMatrix(selrow, 15) = GRDSTOCK.TextMatrix(i + 1, 15)
        GRDSTOCK.TextMatrix(selrow, 16) = GRDSTOCK.TextMatrix(i + 1, 16)
        GRDSTOCK.TextMatrix(selrow, 17) = GRDSTOCK.TextMatrix(i + 1, 17)
        GRDSTOCK.TextMatrix(selrow, 18) = GRDSTOCK.TextMatrix(i + 1, 18)
        GRDSTOCK.TextMatrix(selrow, 19) = GRDSTOCK.TextMatrix(i + 1, 10)
        GRDSTOCK.TextMatrix(selrow, 20) = GRDSTOCK.TextMatrix(i + 1, 20)
        GRDSTOCK.TextMatrix(selrow, 21) = GRDSTOCK.TextMatrix(i + 1, 21)
        GRDSTOCK.TextMatrix(selrow, 22) = GRDSTOCK.TextMatrix(i + 1, 22)
        GRDSTOCK.TextMatrix(selrow, 23) = GRDSTOCK.TextMatrix(i + 1, 23)
        GRDSTOCK.TextMatrix(selrow, 24) = GRDSTOCK.TextMatrix(i + 1, 24)
        GRDSTOCK.TextMatrix(selrow, 25) = GRDSTOCK.TextMatrix(i + 1, 25)
        GRDSTOCK.TextMatrix(selrow, 26) = GRDSTOCK.TextMatrix(i + 1, 26)
        GRDSTOCK.TextMatrix(selrow, 27) = GRDSTOCK.TextMatrix(i + 1, 27)
        GRDSTOCK.TextMatrix(selrow, 28) = GRDSTOCK.TextMatrix(i + 1, 28)
        GRDSTOCK.TextMatrix(selrow, 29) = GRDSTOCK.TextMatrix(i + 1, 29)
        GRDSTOCK.TextMatrix(selrow, 30) = GRDSTOCK.TextMatrix(i + 1, 30)
        selrow = selrow + 1
    Next i
    GRDSTOCK.rows = GRDSTOCK.rows - 1
    GRDSTOCK.SetFocus
    Exit Sub
   
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdDisc_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    Dim i As Long
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If Trim(TXTDISC.Text) = "" Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THESE TAX", vbYesNo + vbDefaultButton2, "Assign TAX....") = vbNo Then Exit Sub
    For i = 1 To GRDSTOCK.rows - 1
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        rststock.Properties("Update Criteria").Value = adCriteriaKey
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
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdExport_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Price Analysis"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, n As Long
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    'Create an Excel instalce.
    Set oApp = CreateObject("Excel.Application")
    Set oWB = oApp.Workbooks.Add
    Set oWS = oWB.Worksheets(1)
    

    
    
'    xlRange = oWS.Range("A1", "C1")
'    xlRange.Font.Bold = True
'    xlRange.ColumnWidth = 15
'    'xlRange.Value = {"First Name", "Last Name", "Last Service"}
'    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
'    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
'
'    xlRange = oWS.Range("C1", "C999")
'    xlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'    xlRange.ColumnWidth = 12
    
    'If Sum_flag = False Then
        oWS.Range("A1", "J1").Merge
        oWS.Range("A1", "J1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "J2").Merge
        oWS.Range("A2", "J2").HorizontalAlignment = xlCenter
    'End If
    oWS.Range("A:A").ColumnWidth = 6
    oWS.Range("B:B").ColumnWidth = 10
    oWS.Range("C:C").ColumnWidth = 12
    oWS.Range("D:D").ColumnWidth = 12
    oWS.Range("E:E").ColumnWidth = 12
    oWS.Range("F:F").ColumnWidth = 12
    oWS.Range("G:G").ColumnWidth = 12
    oWS.Range("H:H").ColumnWidth = 12
    oWS.Range("I:I").ColumnWidth = 12
    oWS.Range("J:J").ColumnWidth = 12
    oWS.Range("K:K").ColumnWidth = 12
    oWS.Range("L:L").ColumnWidth = 12
    oWS.Range("M:M").ColumnWidth = 12
    oWS.Range("N:N").ColumnWidth = 12
    oWS.Range("O:O").ColumnWidth = 12
    oWS.Range("P:P").ColumnWidth = 12
    oWS.Range("Q:Q").ColumnWidth = 12
    oWS.Range("R:R").ColumnWidth = 12
    oWS.Range("S:S").ColumnWidth = 12
    oWS.Range("T:T").ColumnWidth = 12
    oWS.Range("U:U").ColumnWidth = 12
    oWS.Range("V:V").ColumnWidth = 12
    
    oWS.Range("A1").Select                      '-- particular cell selection
    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 14            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    'oApp.Columns("A:A").EntireColumn.AutoFit     '-- autofitted column

    oWS.Range("A2").Select                      '-- particular cell selection
    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 11            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True

'    Range("C2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("C:C").EntireColumn.AutoFit     '-- autofitted column
'
'
'    Range("D2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("D:D").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("E2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("E:E").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("F2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("F:F").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("G2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("G:G").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("H2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("H:H").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("I2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("I:I").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("J2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("J:J").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("K2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("K:K").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("L2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("L:L").EntireColumn.AutoFit     '-- autofitted column

'    oWB.ActiveSheet.Font.Name = "Arial"
'    oApp.ActiveSheet.Font.Name = "Arial"
'    oWB.Font.Size = "11"
'    oWB.Font.Bold = True
    oWS.Range("A" & 1).Value = MDIMAIN.StatusBar.Panels(5).Text
    oWS.Range("A" & 2).Value = "STOCK REPORT"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).Value = GRDSTOCK.TextMatrix(0, 0)
    oWS.Range("B" & 3).Value = GRDSTOCK.TextMatrix(0, 1)
    oWS.Range("C" & 3).Value = GRDSTOCK.TextMatrix(0, 2)
    oWS.Range("D" & 3).Value = GRDSTOCK.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 3).Value = GRDSTOCK.TextMatrix(0, 4)
    oWS.Range("F" & 3).Value = GRDSTOCK.TextMatrix(0, 5)
    oWS.Range("G" & 3).Value = GRDSTOCK.TextMatrix(0, 6)
    oWS.Range("H" & 3).Value = GRDSTOCK.TextMatrix(0, 11)
    oWS.Range("I" & 3).Value = GRDSTOCK.TextMatrix(0, 12)
    oWS.Range("J" & 3).Value = GRDSTOCK.TextMatrix(0, 10)
    oWS.Range("K" & 3).Value = GRDSTOCK.TextMatrix(0, 14)
    oWS.Range("L" & 3).Value = GRDSTOCK.TextMatrix(0, 7)
    oWS.Range("M" & 3).Value = GRDSTOCK.TextMatrix(0, 8)
    oWS.Range("N" & 3).Value = GRDSTOCK.TextMatrix(0, 9)
    oWS.Range("O" & 3).Value = GRDSTOCK.TextMatrix(0, 21)
    oWS.Range("P" & 3).Value = GRDSTOCK.TextMatrix(0, 22)
    oWS.Range("Q" & 3).Value = GRDSTOCK.TextMatrix(0, 29)
    oWS.Range("R" & 3).Value = GRDSTOCK.TextMatrix(0, 30)
    oWS.Range("S" & 3).Value = GRDSTOCK.TextMatrix(0, 18)
    oWS.Range("T" & 3).Value = GRDSTOCK.TextMatrix(0, 19)
    oWS.Range("U" & 3).Value = GRDSTOCK.TextMatrix(0, 25)
'    oWS.Range("V" & 3).Value = GRDSTOCK.TextMatrix(0, 21)
'    oWS.Range("W" & 3).Value = GRDSTOCK.TextMatrix(0, 22)
'    oWS.Range("X" & 3).Value = GRDSTOCK.TextMatrix(0, 23)
'    oWS.Range("Y" & 3).Value = GRDSTOCK.TextMatrix(0, 24)
'    oWS.Range("Z" & 3).Value = GRDSTOCK.TextMatrix(0, 25)
    On Error GoTo ERRHAND
    
    i = 4
    For n = 1 To GRDSTOCK.rows - 1
        oWS.Range("A" & i).Value = GRDSTOCK.TextMatrix(n, 0)
        oWS.Range("B" & i).Value = GRDSTOCK.TextMatrix(n, 1)
        oWS.Range("C" & i).Value = GRDSTOCK.TextMatrix(n, 2)
        oWS.Range("D" & i).Value = GRDSTOCK.TextMatrix(n, 3)
        oWS.Range("E" & i).Value = GRDSTOCK.TextMatrix(n, 4)
        oWS.Range("F" & i).Value = GRDSTOCK.TextMatrix(n, 5)
        oWS.Range("G" & i).Value = GRDSTOCK.TextMatrix(n, 6)
        oWS.Range("H" & i).Value = GRDSTOCK.TextMatrix(n, 11)
        oWS.Range("I" & i).Value = GRDSTOCK.TextMatrix(n, 12)
        oWS.Range("J" & i).Value = GRDSTOCK.TextMatrix(n, 10)
        oWS.Range("K" & i).Value = GRDSTOCK.TextMatrix(n, 14)
        oWS.Range("L" & i).Value = GRDSTOCK.TextMatrix(n, 7)
        oWS.Range("M" & i).Value = GRDSTOCK.TextMatrix(n, 8)
        oWS.Range("N" & i).Value = GRDSTOCK.TextMatrix(n, 9)
        oWS.Range("O" & i).Value = GRDSTOCK.TextMatrix(n, 21)
        oWS.Range("P" & i).Value = GRDSTOCK.TextMatrix(n, 22)
        oWS.Range("Q" & i).Value = GRDSTOCK.TextMatrix(n, 29)
        oWS.Range("R" & i).Value = GRDSTOCK.TextMatrix(n, 30)
        oWS.Range("S" & i).Value = GRDSTOCK.TextMatrix(n, 18)
        oWS.Range("T" & i).Value = GRDSTOCK.TextMatrix(n, 19)
        oWS.Range("U" & i).Value = GRDSTOCK.TextMatrix(n, 25)
'        oWS.Range("V" & i).Value = GRDSTOCK.TextMatrix(n, 21)
'        oWS.Range("W" & i).Value = GRDSTOCK.TextMatrix(n, 22)
'        oWS.Range("X" & i).Value = GRDSTOCK.TextMatrix(n, 23)
'        oWS.Range("Y" & i).Value = GRDSTOCK.TextMatrix(n, 24)
'        oWS.Range("Z" & i).Value = GRDSTOCK.TextMatrix(n, 25)
        On Error GoTo ERRHAND
        i = i + 1
    Next n
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.HorizontalAlignment = xlRight
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 13            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    
   
SKIP:
    oApp.Visible = True
    
'    If Sum_flag = True Then
        'oWS.Columns("C:C").Select
        oWS.Columns("C:C").NumberFormat = "0"
        oWS.Columns("A:Z").EntireColumn.AutoFit
'    End If
    
'    Set oWB = Nothing
'    oApp.Quit
'    Set oApp = Nothing
'
    
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox err.Description
End Sub

Private Sub CmdLoad_Click()
    On Error GoTo ERRHAND
    Call Fillgrid
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdnew_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    If frmunbill.Visible = True Then
        If MsgBox("items will be created as un bill items. Are you sure?", vbYesNo + vbDefaultButton2, "Price Analysis") = vbNo Then Exit Sub
    End If
    If GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 2) <> "" Then
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 0) = GRDSTOCK.rows - 1
        
        Dim TRXMAST As ADODB.Recordset
        On Error GoTo ERRHAND
        
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
    Frmebatch.Visible = False
    TXTEDIT.Visible = False
    TXTEXP.Visible = False
    TXTsample.Visible = False
    GRDSTOCK.TopRow = GRDSTOCK.rows - 1
    GRDSTOCK.Row = GRDSTOCK.rows - 1
    GRDSTOCK.Col = 2
    GRDSTOCK.SetFocus
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub cmdPriceChange_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    Dim i As Long
    On Error GoTo ERRHAND
    If CmbPrChange.ListIndex = -1 Then Exit Sub
    Select Case CmbPrChange.ListIndex
        Case 0
            If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THE SELECTED ITEMS AS PRICE CHANGING ITEMS", vbYesNo + vbDefaultButton2, "Price Analysis") = vbNo Then Exit Sub
        Case Else
            If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THE SELECTED ITEMS AS NON-PRICE CHANGING ITEMS", vbYesNo + vbDefaultButton2, "Price Analysis") = vbNo Then Exit Sub
    End Select
    For i = 1 To GRDSTOCK.rows - 1
        Select Case CmbPrChange.ListIndex
            Case 0
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT P_RETAIL from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_RETAIL) AND P_RETAIL <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("The Price will be affected to all the existing qty for Item " & GRDSTOCK.TextMatrix(i, 2) & ". Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        GoTo SKIP
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
                
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different MRPs Available for the item " & GRDSTOCK.TextMatrix(i, 2) & ". Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        GoTo SKIP
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
                
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different Batches Available for the item " & GRDSTOCK.TextMatrix(i, 2) & ". Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        GoTo SKIP
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
                
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                rststock.Properties("Update Criteria").Value = adCriteriaKey
                If Not (rststock.EOF And rststock.BOF) Then
                    rststock!P_RETAIL = Val(GRDSTOCK.TextMatrix(i, 7))
                    rststock!P_WS = Val(GRDSTOCK.TextMatrix(i, 8))
                    If Val(GRDSTOCK.TextMatrix(i, 12)) <> 0 Then
                        GRDSTOCK.TextMatrix(i, 20) = Format(Round((((Val(GRDSTOCK.TextMatrix(i, 7)) / GRDSTOCK.TextMatrix(i, 13)) - Val(GRDSTOCK.TextMatrix(i, 12))) * 100) / Val(GRDSTOCK.TextMatrix(i, 12)), 2), "0.00")
                    Else
                        GRDSTOCK.TextMatrix(i, 20) = 0
                    End If
                    
                    If Val(GRDSTOCK.TextMatrix(i, 15)) = 0 Then
                        GRDSTOCK.TextMatrix(i, 15) = 1
                        rststock!CRTN_PACK = 1
                    End If
                    If Val(GRDSTOCK.TextMatrix(i, 13)) = 0 Then
                        GRDSTOCK.TextMatrix(i, 13) = 1
                        rststock!LOOSE_PACK = 1
                    End If
                    If Val(GRDSTOCK.TextMatrix(i, 13)) = 0 Then
                        rststock!P_CRTN = Val(GRDSTOCK.TextMatrix(i, 7))
                        rststock!P_LWS = Val(GRDSTOCK.TextMatrix(i, 8))
                        GRDSTOCK.TextMatrix(i, 16) = Format(Val(GRDSTOCK.TextMatrix(i, 7)), "0.000")
                    Else
                        rststock!P_CRTN = Round(Val(GRDSTOCK.TextMatrix(i, 7)) / Val(GRDSTOCK.TextMatrix(i, 13)), 2)
                        rststock!P_LWS = Round(Val(GRDSTOCK.TextMatrix(i, 8)) / Val(GRDSTOCK.TextMatrix(i, 13)), 2)
                        GRDSTOCK.TextMatrix(i, 16) = Format(Val(GRDSTOCK.TextMatrix(i, 7)) / Val(GRDSTOCK.TextMatrix(i, 13)), "0.000")
                    End If
                
                    rststock.Update
                End If
                rststock.Close
                Set rststock = Nothing
                
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
                rststock.Properties("Update Criteria").Value = adCriteriaKey
                Do Until rststock.EOF
                    rststock!P_RETAIL = Val(GRDSTOCK.TextMatrix(i, 7))
                    rststock!P_WS = Val(GRDSTOCK.TextMatrix(i, 8))
                    rststock!P_VAN = Val(GRDSTOCK.TextMatrix(i, 9))
                    If Val(GRDSTOCK.TextMatrix(i, 15)) = 0 Then
                        rststock!CRTN_PACK = 1
                    End If
                    If Val(GRDSTOCK.TextMatrix(i, 13)) = 0 Then
                        rststock!LOOSE_PACK = 1
                    End If
                    rststock.Update
                    rststock.MoveNext
                Loop
                rststock.Close
                Set rststock = Nothing
                
                db.Execute "Update ITEMMAST set PRICE_CHANGE = 'Y' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
                GRDSTOCK.TextMatrix(i, 28) = "Yes"
            Case Else
                db.Execute "Update ITEMMAST set PRICE_CHANGE = 'N' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
                GRDSTOCK.TextMatrix(i, 28) = "No"
        End Select
SKIP:
    Next i
    MsgBox "Successfully Applied", , "Price Analysis"
    CmbPrChange.ListIndex = -1
    Exit Sub
ERRHAND:
    MsgBox err.Description
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
    
    If lblstktype.Caption = "A" Then Exit Sub
    Dim itemtable As String
    Dim trxtable As String
    Dim rtrxtable As String
    If OptBrStock.Value = True Then
        lblstock.Caption = "Branch Stock"
        lblstktype.Caption = "B"
        itemtable = "ITEMMASTVAN"
        trxtable = "TRXFILEVAN"
        rtrxtable = "RTRXFILEVAN"
    ElseIf OptAllStock.Value = True Then
        lblstock.Caption = "Main & Branch Stock"
        lblstktype.Caption = "A"
        itemtable = "ITEMMAST"
        trxtable = "TRXFILE"
        rtrxtable = "RTRXFILE"
    Else
        lblstock.Caption = "Main Stock"
        lblstktype.Caption = "M"
        itemtable = "ITEMMAST"
        trxtable = "TRXFILE"
        rtrxtable = "RTRXFILE"
    End If
    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    On Error GoTo ERRHAND
    
    db.Execute "Update RTRXFILE set BAL_QTY = 0 WHERE ISNULL(BAL_QTY) OR BAL_QTY <0 "
    For i = 1 To GRDSTOCK.rows - 1
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM " & itemtable & " WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTITEMMAST.Properties("Update Criteria").Value = adCriteriaKey
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            
            BALQTY = 0
            Set RSTBALQTY = New ADODB.Recordset
            RSTBALQTY.Open "Select SUM(BAL_QTY) FROM " & rtrxtable & " WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "'", db, adOpenForwardOnly
            If Not (RSTBALQTY.EOF And RSTBALQTY.BOF) Then
                BALQTY = IIf(IsNull(RSTBALQTY.Fields(0)), 0, RSTBALQTY.Fields(0))
            End If
            RSTBALQTY.Close
            Set RSTBALQTY = Nothing
            
            INWARD = 0
            OUTWARD = 0
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM(QTY) FROM " & rtrxtable & " where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                INWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
                
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM(FREE_QTY) FROM " & rtrxtable & " where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                INWARD = INWARD + IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
            
                    
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM " & trxtable & " WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until rststock.EOF
                OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            
            If Round(INWARD - OUTWARD, 2) = Round(BALQTY, 2) Then GoTo SKIP_BALCHECK
            
            
            db.Execute "Update RTRXFILE set BAL_QTY = QTY where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' "
            BALQTY = 0
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM " & rtrxtable & " WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
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
            rststock.Open "SELECT SUM(BAL_QTY) FROM " & rtrxtable & " where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                BALQTY = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
            
            If Round(INWARD - OUTWARD, 2) < BALQTY Then
                DIFFQTY = BALQTY - (Round(INWARD - OUTWARD, 2))
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * FROM " & rtrxtable & " where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ", db, adOpenStatic, adLockOptimistic, adCmdText
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
                rststock.Open "SELECT * FROM " & rtrxtable & " where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockOptimistic, adCmdText
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
    Next i
    Screen.MousePointer = vbNormal

    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Command1_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    Dim i As Long
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If Trim(txtHSN.Text) = "" Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THESE HSN CODES", vbYesNo + vbDefaultButton2, "Assign HSN CODES....") = vbNo Then Exit Sub
    For i = 1 To GRDSTOCK.rows - 1
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        rststock.Properties("Update Criteria").Value = adCriteriaKey
        If Not (rststock.EOF And rststock.BOF) Then
            rststock!REMARKS = Trim(txtHSN.Text)
            GRDSTOCK.TextMatrix(i, 14) = Trim(txtHSN.Text)
            rststock.Update
        End If
        rststock.Close
        Set rststock = Nothing
    Next i
    txtHSN.Text = ""
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Command2_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Price Analysis"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Import Stock Items") = vbNo Then Exit Sub
    If MsgBox("Sheet Name should be 'ITEMS' and First coloumn should be Item Code and Second coloumn should be Item name", vbYesNo, "Import Items") = vbNo Then Exit Sub
    On Error GoTo ERRHAND
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "Excel Files (*.xls*)|*.xls*"
    CommonDialog1.ShowOpen
    
    Screen.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    
    'Set wb = xlApp.Workbooks.Open("PATH TO YOUR EXCEL FILE")
    Set wb = xlApp.Workbooks.Open(CommonDialog1.FileName)
    
    Set ws = wb.Worksheets("ITEMS") 'Specify your worksheet name
    var = ws.Range("A1").Value
    
'''    db.Execute "dELETE FROM ITEMMAST"
'''    db.Execute "dELETE FROM RTRXFILE"
'''
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTITEMTRX As ADODB.Recordset
    Dim ITEMCODE As String
    Dim sl As Integer
    Dim lastno As Integer
    sl = 1
    lastno = 1
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'OP'", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        lastno = IIf(IsNull(RSTITEMMAST.Fields(0)), 1, RSTITEMMAST.Fields(0) + 1)
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    For i = 2 To 300000
        If ws.Range("A" & i).Value = "" Then Exit For
        
        'If MsgBox("Item not exists!!! Do You want to add this item?", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then Exit Sub
'        Set RSTITEMTRX = New ADODB.Recordset
'        RSTITEMTRX.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME = '" & Trim(ws.Range("B" & i).value) & "' AND ITEM_CODE = '" & ws.Range("A" & i).value & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (RSTITEMTRX.EOF And RSTITEMTRX.BOF) Then
'            MsgBox "Duplicate Name. Item " & Trim(ws.Range("B" & i).value) & " Skipped", vbOKOnly, "IMPORT ITEMS"
'            RSTITEMTRX.Close
'            Set RSTITEMTRX = Nothing
'            GoTo SKIP
'        End If
'        RSTITEMTRX.Close
'        Set RSTITEMTRX = Nothing
            
        Set RSTITEMTRX = New ADODB.Recordset
        RSTITEMTRX.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & ws.Range("A" & i).Value & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTITEMTRX.EOF And RSTITEMTRX.BOF) Then
            ITEMCODE = ""
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                If IsNull(RSTITEMMAST.Fields(0)) Then
                    ITEMCODE = 1
                Else
                    ITEMCODE = Val(RSTITEMMAST.Fields(0)) + 1
                End If
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        Else
            ITEMCODE = ws.Range("A" & i).Value
        End If
        RSTITEMTRX.Close
        Set RSTITEMTRX = Nothing
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & ITEMCODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            'RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
            RSTITEMMAST!ITEM_CODE = ITEMCODE
            RSTITEMMAST!ITEM_NAME = Trim(ws.Range("B" & i).Value)
            RSTITEMMAST!UNIT = 1
            RSTITEMMAST!DEAD_STOCK = "N"
            RSTITEMMAST!REMARKS = Trim(ws.Range("I" & i).Value)
            RSTITEMMAST!REORDER_QTY = 1
            RSTITEMMAST!PACK_TYPE = IIf(IsNull(ws.Range("D" & i).Value) Or Trim(ws.Range("D" & i).Value) = "", "Nos", Trim(ws.Range("D" & i).Value))
            RSTITEMMAST!FULL_PACK = IIf(IsNull(ws.Range("E" & i).Value) Or Trim(ws.Range("E" & i).Value) = "", "Nos", Trim(ws.Range("E" & i).Value))
            RSTITEMMAST!BIN_LOCATION = "" 'Trim(ws.Range("N" & i).value)
            RSTITEMMAST!MRP = Val(ws.Range("F" & i).Value)
            RSTITEMMAST!PTR = Val(ws.Range("G" & i).Value)
            RSTITEMMAST!CST = 0
            RSTITEMMAST!OPEN_QTY = 0
            RSTITEMMAST!OPEN_VAL = 0
            RSTITEMMAST!RCPT_QTY = Val(ws.Range("C" & i).Value)
            RSTITEMMAST!RCPT_VAL = Val(ws.Range("C" & i).Value) * Val(ws.Range("G" & i).Value)
            RSTITEMMAST!ISSUE_QTY = 0
            RSTITEMMAST!ISSUE_VAL = 0
            RSTITEMMAST!CLOSE_QTY = Val(ws.Range("C" & i).Value)
            RSTITEMMAST!CLOSE_VAL = Val(ws.Range("C" & i).Value) * Val(ws.Range("G" & i).Value)
            RSTITEMMAST!DAM_QTY = 0
            RSTITEMMAST!DAM_VAL = 0
            RSTITEMMAST!DISC = 0
            RSTITEMMAST!SALES_TAX = Val(ws.Range("H" & i).Value)
            If Val(ws.Range("R" & i).Value) <= 0 Then
                RSTITEMMAST!ITEM_COST = Val(ws.Range("G" & i).Value)
            Else
                RSTITEMMAST!ITEM_COST = Val(ws.Range("G" & i).Value) / Val(ws.Range("R" & i).Value)
            End If
            RSTITEMMAST!P_RETAIL = Val(ws.Range("J" & i).Value)
            RSTITEMMAST!P_WS = Val(ws.Range("K" & i).Value)
            RSTITEMMAST!P_VAN = Val(ws.Range("L" & i).Value)
            RSTITEMMAST!CUST_DISC = Val(ws.Range("M" & i).Value)
            RSTITEMMAST!BARCODE = ws.Range("O" & i).Value
            
            
            RSTITEMMAST!CRTN_PACK = 1
            RSTITEMMAST!P_CRTN = Val(ws.Range("S" & i).Value)
            RSTITEMMAST!LOOSE_PACK = Val(ws.Range("R" & i).Value)
            RSTITEMMAST!check_flag = "V"
            If Trim(ws.Range("N" & i).Value) = "Y" Then
                RSTITEMMAST!UN_BILL = "Y"
            Else
                RSTITEMMAST!UN_BILL = "N"
            End If
            If Trim(ws.Range("P" & i).Value) = "" Then
                RSTITEMMAST!Category = "GENERAL"
            Else
                RSTITEMMAST!Category = Trim(ws.Range("P" & i).Value)
            End If
            If Trim(ws.Range("Q" & i).Value) = "" Then
                RSTITEMMAST!MANUFACTURER = "GENERAL"
            Else
                RSTITEMMAST!MANUFACTURER = Left(Trim(ws.Range("Q" & i).Value), 25)
            End If
            RSTITEMMAST!BARCODE = Trim(ws.Range("O" & i).Value)
            RSTITEMMAST.Update
        End If
        db.CommitTrans
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        If Val(ws.Range("C" & i).Value) > 0 Then
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE ", db, adOpenStatic, adLockOptimistic, adCmdText
            rststock.AddNew
            rststock!TRX_TYPE = "OP"
            rststock!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            rststock!VCH_DATE = Format(Date, "DD/MM/YYYY")
            rststock!VCH_NO = lastno
            rststock!LINE_NO = sl
            rststock!ITEM_CODE = ITEMCODE
            rststock!BARCODE = Trim(ws.Range("O" & i).Value)
            rststock!BAL_QTY = Val(ws.Range("C" & i).Value)
            rststock!QTY = Val(ws.Range("C" & i).Value)
            rststock!TRX_TOTAL = Val(ws.Range("C" & i).Value) * Val(ws.Range("G" & i).Value)
            rststock!VCH_DATE = Format(Date, "dd/mm/yyyy")
            rststock!ITEM_NAME = Trim(ws.Range("B" & i).Value)
            If Val(ws.Range("R" & i).Value) <= 0 Then
                rststock!ITEM_COST = Val(ws.Range("G" & i).Value)
            Else
                rststock!ITEM_COST = Val(ws.Range("G" & i).Value) / Val(ws.Range("R" & i).Value)
            End If
            'rststock!ITEM_COST = Val(ws.Range("G" & i).value)
            rststock!LINE_DISC = 1
            rststock!P_DISC = 0
            rststock!MRP = Val(ws.Range("F" & i).Value)
            rststock!PTR = Val(ws.Range("G" & i).Value)
            rststock!SALES_PRICE = Val(ws.Range("J" & i).Value)
            rststock!P_RETAIL = Val(ws.Range("J" & i).Value)
            rststock!P_WS = Val(ws.Range("K" & i).Value)
            rststock!P_VAN = Val(ws.Range("L" & i).Value)
            rststock!P_CRTN = Val(ws.Range("S" & i).Value)
            rststock!P_LWS = Val(ws.Range("S" & i).Value)
            rststock!CRTN_PACK = 1
            rststock!gross_amt = 0
            rststock!COM_FLAG = "P"
            rststock!COM_PER = 0
            rststock!COM_AMT = 0
            rststock!SALES_TAX = Val(ws.Range("H" & i).Value)
            rststock!LOOSE_PACK = Val(ws.Range("R" & i).Value)
            rststock!PACK_TYPE = IIf(IsNull(ws.Range("D" & i).Value) Or Trim(ws.Range("D" & i).Value) = "", "Nos", Trim(ws.Range("D" & i).Value))
            'rststock!WARRANTY = Null
            'rststock!WARRANTY_TYPE = Null
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
            rststock!Category = Trim(ws.Range("P" & i).Value)
            rststock!MFGR = Trim(ws.Range("Q" & i).Value)
            'rststock!M_USER_ID = DataList2.BoundText
            'rststock!PINV = Trim(TXTINVOICE.Text)
            rststock.Update
            rststock.Close
            Set rststock = Nothing
            sl = sl + 1
        End If
                        
SKIP:
    Next i
    wb.Close
    
    xlApp.Quit
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbNormal
    
    Call CmdLoad_Click
    MsgBox "Success", vbOKOnly
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = 9 Then
        MsgBox "NO SUCH FILE PRESENT!!", vbOKOnly, "IMPORT ITEMS"
        wb.Close
        xlApp.Quit
        Set ws = Nothing
        Set wb = Nothing
        Set xlApp = Nothing
    ElseIf err.Number = 32755 Then
        
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub Command3_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    Dim i As Long
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If Trim(TXTDEALER2.Text) = "" Then
        MsgBox "Please enter a Category Name in the Text Box", vbOKOnly, "Price Analysis"
        TXTDEALER2.SetFocus
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THE CATEGORY TO ALL LISTED ITEMS", vbYesNo + vbDefaultButton2, "Assign CATEGORY....") = vbNo Then Exit Sub
    For i = 1 To GRDSTOCK.rows - 1
        db.Execute "Update ITEMMAST set Category = '" & Trim(TXTDEALER2.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
        db.Execute "Update trxfile set Category = '" & Trim(TXTDEALER2.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
        db.Execute "Update rtrxfile set Category = '" & Trim(TXTDEALER2.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
        GRDSTOCK.TextMatrix(i, 18) = Trim(TXTDEALER2.Text)
    Next i
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT DISTINCT CATEGORY FROM CATEGORY where CATEGORY = '" & Trim(TXTDEALER2.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (rststock.EOF And rststock.BOF) Then
        rststock.AddNew
        rststock!Category = Trim(TXTDEALER2.Text)
        rststock.Update
    End If
    rststock.Close
    Set rststock = Nothing
                    
    TXTDEALER2.Text = ""
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Command4_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    Dim i As Long
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If Trim(TXTDEALER2.Text) = "" Then
        MsgBox "Please enter a Company Name in the Text Box", vbOKOnly, "Price Analysis"
        TXTDEALER2.SetFocus
        Exit Sub
    End If
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN THE COMPANY TO ALL LISTED ITEMS", vbYesNo + vbDefaultButton2, "Assign COMPANY....") = vbNo Then Exit Sub
    For i = 1 To GRDSTOCK.rows - 1
        db.Execute "Update ITEMMAST set MANUFACTURER = '" & Trim(TXTDEALER2.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
        db.Execute "Update trxfile set MFGR = '" & Trim(TXTDEALER2.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
        db.Execute "Update rtrxfile set MFGR = '" & Trim(TXTDEALER2.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'"
        GRDSTOCK.TextMatrix(i, 19) = Trim(TXTDEALER2.Text)
    Next i
    TXTDEALER2.Text = ""
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Command5_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    Dim i As Long
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    If Trim(txtcustdic.Text) = "" Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO ASSIGN CUSTOMER DISC TO ALL", vbYesNo + vbDefaultButton2, "Assign Cust Disc....") = vbNo Then Exit Sub
    For i = 1 To GRDSTOCK.rows - 1
        Set rststock = New ADODB.Recordset
        rststock.Properties("Update Criteria").Value = adCriteriaKey
        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (rststock.EOF And rststock.BOF) Then
            rststock!CUST_DISC = Val(txtcustdic.Text)
            'rststock!P_RETAIL = rststock!MRP
            GRDSTOCK.TextMatrix(i, 21) = Val(txtcustdic.Text)
            rststock.Update
        End If
        rststock.Close
        Set rststock = Nothing
        
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' WHERE BAL_QTY >0", db, adOpenStatic, adLockOptimistic, adCmdText
'        Do Until rststock.EOF
'            rststock!CUST_DISC = Val(txtcustdic.Text)
'            'rststock!P_RETAIL = rststock!MRP
'            GRDSTOCK.TextMatrix(i, 17) = Val(txtcustdic.Text)
'            rststock.Update
'            rststock.MoveNext
'        Loop
'        rststock.Close
'        Set rststock = Nothing
        
    Next i
    txtcustdic.Text = ""
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Command6_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Price Analysis"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "ASSIGN RATES") = vbNo Then Exit Sub
    If MsgBox("Sheet Name should be 'RATES' and First coloumn should be Item Code 2nd Cloumn should be Item name", vbYesNo, "ASSIGN RATES") = vbNo Then Exit Sub
    On Error GoTo ERRHAND
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExistWAMP
    CommonDialog1.Filter = "Excel Files (*.xls*)|*.xls*"
    CommonDialog1.ShowOpen
    
    Screen.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    
    'Set wb = xlApp.Workbooks.Open("PATH TO YOUR EXCEL FILE")
    Set wb = xlApp.Workbooks.Open(CommonDialog1.FileName)
    
    Set ws = wb.Worksheets("RATES") 'Specify your worksheet name
    var = ws.Range("A1").Value
    
'    db.Execute "dELETE FROM ITEMMAST"
'    db.Execute "dELETE FROM RTRXFILE"
        
    Dim rststock As ADODB.Recordset
    For i = 2 To 30000
        If ws.Range("A" & i).Value = "" Then Exit For
        db.Execute "Update ITEMMAST set PACK_TYPE = '" & Trim(ws.Range("F" & i).Value) & "',  UQC = '" & Trim(ws.Range("F" & i).Value) & "',  FULL_PACK = '" & Trim(ws.Range("F" & i).Value) & "', MRP = " & Val(ws.Range("E" & i).Value) & ", P_WS = " & Val(ws.Range("D" & i).Value) & ", P_VAN = " & Val(ws.Range("C" & i).Value) & " where ITEM_CODE = '" & ws.Range("A" & i).Value & "' "
        db.Execute "Update RTRXFILE set PACK_TYPE = '" & Trim(ws.Range("F" & i).Value) & "', FULL_PACK = '" & Trim(ws.Range("F" & i).Value) & "', MRP = " & Val(ws.Range("E" & i).Value) & " , P_WS = " & Val(ws.Range("D" & i).Value) & ", P_VAN = " & Val(ws.Range("C" & i).Value) & " where ITEM_CODE = '" & ws.Range("A" & i).Value & "' "
        
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & ws.Range("A" & i).value & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (rststock.EOF And rststock.BOF) Then
'            rststock!ITEM_CODE = Left(rststock!ITEM_CODE, 20)
'            rststock.Update
'        End If
'        rststock.Close
'        Set rststock = Nothing
'
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & ws.Range("A" & i).value & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        Do Until rststock.EOF
'            rststock!ITEM_CODE = Left(rststock!ITEM_CODE, 20)
'            rststock.Update
'            rststock.MoveNext
'        Loop
'        rststock.Close
'        Set rststock = Nothing
                    
    Next i
    
    
    wb.Close
    
    xlApp.Quit
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbNormal
    
    Call CmdLoad_Click
    MsgBox "Success", vbOKOnly
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = 9 Then
        MsgBox "NO SUCH FILE PRESENT!!", vbOKOnly, "IMPORT ITEMS"
        wb.Close
        xlApp.Quit
        Set ws = Nothing
        Set wb = Nothing
        Set xlApp = Nothing
    ElseIf err.Number = 32755 Then
        
    Else
        MsgBox err.Description
    End If
End Sub


Private Sub Command7_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Price Analysis"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Import Stock Items") = vbNo Then Exit Sub
    If MsgBox("Sheet Name should be 'QTY' and First coloumn should be Item Code and Second coloumn should be Qty", vbYesNo, "Import Qty") = vbNo Then Exit Sub
    On Error GoTo ERRHAND
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "Excel Files (*.xls*)|*.xls*"
    CommonDialog1.ShowOpen
    
    Screen.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    
    'Set wb = xlApp.Workbooks.Open("PATH TO YOUR EXCEL FILE")
    Set wb = xlApp.Workbooks.Open(CommonDialog1.FileName)
    
    Set ws = wb.Worksheets("QTY") 'Specify your worksheet name
    var = ws.Range("A1").Value
    
    'db.Execute "dELETE FROM ITEMMAST"
    'db.Execute "dELETE FROM RTRXFILE"
    
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTITEMTRX As ADODB.Recordset
    Dim ITEMCODE As String
    Dim sl As Integer
    Dim lastno As Integer
    sl = 1
    lastno = 1
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'OP'", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        lastno = IIf(IsNull(RSTITEMMAST.Fields(0)), 1, RSTITEMMAST.Fields(0) + 1)
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    For i = 2 To 300000
        If ws.Range("A" & i).Value = "" Then Exit For
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & ws.Range("A" & i).Value & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        'db.BeginTrans
        If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            If Val(ws.Range("C" & i).Value) > 0 Then
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * FROM RTRXFILE ", db, adOpenStatic, adLockOptimistic, adCmdText
                rststock.AddNew
                rststock!TRX_TYPE = "OP"
                rststock!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                rststock!VCH_DATE = Format(Date, "DD/MM/YYYY")
                rststock!VCH_NO = lastno
                rststock!LINE_NO = sl
                rststock!ITEM_CODE = ws.Range("A" & i).Value
                
                rststock!BARCODE = IIf(IsNull(RSTITEMMAST!BARCODE), "", RSTITEMMAST!BARCODE)
                rststock!BAL_QTY = Val(ws.Range("B" & i).Value)
                rststock!QTY = Val(ws.Range("B" & i).Value)
                rststock!TRX_TOTAL = Val(ws.Range("B" & i).Value) * Val(ws.Range("D" & i).Value)
                rststock!VCH_DATE = Format(Date, "dd/mm/yyyy")
                rststock!ITEM_NAME = IIf(IsNull(RSTITEMMAST!ITEM_NAME), "", RSTITEMMAST!ITEM_NAME)
                rststock!ITEM_COST = Val(ws.Range("D" & i).Value)
                rststock!LINE_DISC = 1
                rststock!P_DISC = 0
                rststock!MRP = Val(ws.Range("C" & i).Value)
                rststock!PTR = Val(ws.Range("D" & i).Value)
                rststock!SALES_PRICE = Val(ws.Range("E" & i).Value)
                rststock!P_RETAIL = Val(ws.Range("E" & i).Value)
                rststock!P_WS = Val(ws.Range("F" & i).Value)
                rststock!P_VAN = Val(ws.Range("G" & i).Value)
                rststock!P_CRTN = Val(ws.Range("E" & i).Value)
                rststock!P_LWS = Val(ws.Range("F" & i).Value)
                rststock!CRTN_PACK = 1
                rststock!gross_amt = 0
                rststock!COM_FLAG = "P"
                rststock!COM_PER = 0
                rststock!COM_AMT = 0
                rststock!SALES_TAX = IIf(IsNull(RSTITEMMAST!SALES_TAX), 0, RSTITEMMAST!SALES_TAX)
                rststock!LOOSE_PACK = 1
                rststock!PACK_TYPE = IIf(IsNull(RSTITEMMAST!PACK_TYPE), "Nos", RSTITEMMAST!PACK_TYPE)
                'rststock!WARRANTY = Null
                'rststock!WARRANTY_TYPE = Null
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
                rststock!Category = IIf(IsNull(RSTITEMMAST!Category), "", RSTITEMMAST!Category)
                rststock!MFGR = IIf(IsNull(RSTITEMMAST!MANUFACTURER), "", RSTITEMMAST!MANUFACTURER)
                'rststock!M_USER_ID = DataList2.BoundText
                'rststock!PINV = Trim(TXTINVOICE.Text)
                rststock.Update
                rststock.Close
                Set rststock = Nothing
                sl = sl + 1
            End If
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
SKIP:
    Next i
    wb.Close
    
    xlApp.Quit
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbNormal
    
    Call CmdLoad_Click
    MsgBox "Success", vbOKOnly
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = 9 Then
        MsgBox "NO SUCH FILE PRESENT!!", vbOKOnly, "IMPORT ITEMS"
        wb.Close
        xlApp.Quit
        Set ws = Nothing
        Set wb = Nothing
        Set xlApp = Nothing
    ElseIf err.Number = 32755 Then
        
    Else
        MsgBox err.Description
    End If
End Sub

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
            Case 68
                Call CmdDelete_Click
        End Select
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ERRHAND
    
    Set CMBMFGR.DataSource = Nothing
    MFG_REC.Open "SELECT DISTINCT MANUFACTURER FROM ITEMMAST ORDER BY MANUFACTURER", db, adOpenForwardOnly
    Set CMBMFGR.RowSource = MFG_REC
    CMBMFGR.ListField = "MANUFACTURER"
    
    Set Cmbcategory.DataSource = Nothing
    CAT_REC.Open "SELECT DISTINCT CATEGORY FROM CATEGORY ORDER BY CATEGORY", db, adOpenForwardOnly
    Set Cmbcategory.RowSource = CAT_REC
    Cmbcategory.ListField = "CATEGORY"
    
    
    db.Execute "Update itemmast set category = '' where isnull(category) "
    db.Execute "Update itemmast set REMARKS = '' where isnull(REMARKS) "
    
    txtstkcrct.Text = 0
    
    frmloadflag = True
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        txtstkcrct.Text = IIf(IsNull(RSTCOMPANY!STOCK_CRCT), 0, RSTCOMPANY!STOCK_CRCT)
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
    
    If frmLogin.rs!Level <> "0" Then
        CmdExport.Visible = False
        Command2.Visible = False
        CMDPRINT.Visible = False
        Label1(4).Visible = False
        Label1(6).Visible = False
        lblpvalue.Visible = False
        lblnetvalue.Visible = False
        lblsalevalue.Visible = False
    End If
    
    frmloadflag = False
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
    If chkws.Value = 1 Then
        GRDSTOCK.TextMatrix(0, 8) = ""
        GRDSTOCK.ColWidth(8) = 0
    Else
        GRDSTOCK.TextMatrix(0, 8) = "WS"
        GRDSTOCK.ColWidth(8) = 1000
    End If
    If chkvp.Value = 1 Then
        GRDSTOCK.TextMatrix(0, 9) = ""
        GRDSTOCK.ColWidth(9) = 0
    Else
        GRDSTOCK.TextMatrix(0, 9) = "VP"
        GRDSTOCK.ColWidth(9) = 1000
    End If
    GRDSTOCK.TextMatrix(0, 10) = "Tax"
    GRDSTOCK.TextMatrix(0, 13) = "Unit"
    GRDSTOCK.TextMatrix(0, 14) = "HSN Code"
    GRDSTOCK.TextMatrix(0, 15) = "" '"L.Pack"
    GRDSTOCK.TextMatrix(0, 16) = "L.R.Price"
    If chklwp.Value = 1 Then
        GRDSTOCK.TextMatrix(0, 17) = ""
        GRDSTOCK.ColWidth(17) = 0
    Else
        GRDSTOCK.TextMatrix(0, 17) = "L.W.Price"
        GRDSTOCK.ColWidth(17) = 1000
    End If
    If chkhidecat.Value = 1 Then
        GRDSTOCK.TextMatrix(0, 18) = ""
        GRDSTOCK.ColWidth(18) = 0
    Else
        GRDSTOCK.TextMatrix(0, 18) = "Category"
        GRDSTOCK.ColWidth(18) = 1000
    End If
    If chkhidecomp.Value = 1 Then
        GRDSTOCK.TextMatrix(0, 19) = ""
        GRDSTOCK.ColWidth(19) = 0
    Else
        GRDSTOCK.TextMatrix(0, 19) = "Company"
        GRDSTOCK.ColWidth(19) = 1000
    End If
    If frmLogin.rs!Level = "0" Then
        GRDSTOCK.TextMatrix(0, 11) = "Per Rate"
        GRDSTOCK.TextMatrix(0, 12) = "Net Cost"
        GRDSTOCK.TextMatrix(0, 20) = "Profit%"
        GRDSTOCK.ColWidth(11) = 1000
        GRDSTOCK.ColWidth(12) = 900
        GRDSTOCK.ColWidth(20) = 900
    Else
        GRDSTOCK.TextMatrix(0, 11) = ""
        GRDSTOCK.TextMatrix(0, 12) = ""
        GRDSTOCK.TextMatrix(0, 20) = ""
        GRDSTOCK.ColWidth(11) = 0
        GRDSTOCK.ColWidth(12) = 0
        GRDSTOCK.ColWidth(20) = 0
    End If
    GRDSTOCK.TextMatrix(0, 21) = "Disc %"
    GRDSTOCK.TextMatrix(0, 22) = "Disc Amt"
    GRDSTOCK.TextMatrix(0, 23) = "Commi"
    GRDSTOCK.TextMatrix(0, 24) = "Type"
    GRDSTOCK.TextMatrix(0, 25) = "Value"
    GRDSTOCK.TextMatrix(0, 26) = "Cess%"
    GRDSTOCK.TextMatrix(0, 27) = "Cess Rate"
    GRDSTOCK.TextMatrix(0, 28) = "Price Change"
    GRDSTOCK.TextMatrix(0, 29) = "U_BILL"
    GRDSTOCK.TextMatrix(0, 30) = "BARCODE"
    
    GRDSTOCK.ColWidth(0) = 400
    GRDSTOCK.ColWidth(1) = 900
    GRDSTOCK.ColWidth(2) = 4300
    GRDSTOCK.ColWidth(3) = 1000
    GRDSTOCK.ColWidth(4) = 1000
    GRDSTOCK.ColWidth(5) = 900
    GRDSTOCK.ColWidth(6) = 1000
    GRDSTOCK.ColWidth(7) = 1000
    GRDSTOCK.ColWidth(10) = 800
    
    GRDSTOCK.ColWidth(13) = 800
    GRDSTOCK.ColWidth(14) = 1000
    GRDSTOCK.ColWidth(15) = 0 '1000 'pACK
    GRDSTOCK.ColWidth(16) = 1000
    
    GRDSTOCK.ColWidth(21) = 1000
    GRDSTOCK.ColWidth(22) = 1000
    GRDSTOCK.ColWidth(23) = 1000
    GRDSTOCK.ColWidth(24) = 1000
    GRDSTOCK.ColWidth(25) = 1500
    GRDSTOCK.ColWidth(26) = 900
    GRDSTOCK.ColWidth(27) = 1200
    GRDSTOCK.ColWidth(28) = 1200
    GRDSTOCK.ColWidth(29) = 0
    GRDSTOCK.ColWidth(30) = 2500
    
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
    GRDSTOCK.ColAlignment(29) = 4
    GRDSTOCK.ColAlignment(30) = 1
        
    grdbatch.TextMatrix(0, 0) = "SL"
    grdbatch.TextMatrix(0, 1) = "QTY"
    grdbatch.TextMatrix(0, 2) = "PRICE"
    grdbatch.TextMatrix(0, 3) = "MRP"
    grdbatch.TextMatrix(0, 4) = "EXPIRY"
    grdbatch.TextMatrix(0, 5) = "BATCH"
    grdbatch.TextMatrix(0, 6) = "PACK"
    grdbatch.TextMatrix(0, 7) = "BARCODE"
    grdbatch.TextMatrix(0, 8) = "TRX TYPE"
    grdbatch.TextMatrix(0, 9) = "VCH NO"
    grdbatch.TextMatrix(0, 10) = "LINE NO"
    grdbatch.TextMatrix(0, 11) = "WS"
    
    grdbatch.ColWidth(0) = 400
    grdbatch.ColWidth(1) = 1000
    grdbatch.ColWidth(2) = 900
    grdbatch.ColWidth(3) = 900
    grdbatch.ColWidth(4) = 1000
    grdbatch.ColWidth(5) = 1000
    grdbatch.ColWidth(6) = 800
    grdbatch.ColWidth(7) = 2500
    grdbatch.ColWidth(8) = 0
    grdbatch.ColWidth(9) = 0
    grdbatch.ColWidth(10) = 0
    grdbatch.ColWidth(11) = 900
    
    grdbatch.ColAlignment(0) = 1
    grdbatch.ColAlignment(1) = 4
    grdbatch.ColAlignment(2) = 4
    grdbatch.ColAlignment(3) = 4
    grdbatch.ColAlignment(4) = 4
    grdbatch.ColAlignment(5) = 1
    grdbatch.ColAlignment(6) = 4
    grdbatch.ColAlignment(7) = 1
    grdbatch.ColAlignment(11) = 4
    
    CmbPrChange.ListIndex = -1
    DTFROM.Value = Format(Date, "DD/MM/YYYY")
    DTFROM.Value = Null
    'Call Fillgrid
    'Me.Height = 8415
    'Me.Width = 6465
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ERRHAND:
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

Private Sub grdbatch_LostFocus()
    'Frmebatch.Visible = False
End Sub

Private Sub GRDSTOCK_Click()
    'Frmebatch.Visible = False
    If lblstktype.Caption <> "M" Then Exit Sub
'    Dim PHY As ADODB.Recordset
'    Frame6.Visible = False
'    Set Image1.DataSource = Nothing
'    bytData = ""
'    Set PHY = New ADODB.Recordset
'    PHY.Open "Select * FROM ITEMMAST WHERE ITEM_CODE ='" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockReadOnly
'    If Not (PHY.BOF And PHY.EOF) Then
'        On Error Resume Next
'        Set Image1.DataSource = PHY
'        If IsNull(PHY!PHOTO) Then
'            Frame6.Visible = False
'            Set Image1.DataSource = Nothing
'            bytData = ""
'        Else
'            If Err.Number = 545 Then
'                Frame6.Visible = False
'                Set Image1.DataSource = Nothing
'                bytData = ""
'            Else
'                Frame6.Visible = True
'                Set Image1.DataSource = PHY 'setting image1’s datasource
'                Image1.DataField = "PHOTO"
'                bytData = PHY!PHOTO
'            End If
'        End If
'    End If
'    PHY.Close
'    Set PHY = Nothing
    
    'Frmebatch.Visible = False
    TXTEDIT.Visible = False
    TXTEXP.Visible = False
    TXTsample.Visible = False
    CmbPack.Visible = False
    CMBCHANGE.Visible = False
    CMBMFGR.Visible = False
    FRAME.Visible = False
    Select Case GRDSTOCK.Col
        Case 1
        Case Else
            'GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            If lblstktype.Caption <> "M" Then Exit Sub
            If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
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
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) = 0 Then
                            MsgBox "Please enter the cost", vbOKOnly, "Price Analysis"
                            GRDSTOCK.SetFocus
                            Exit Sub
                        End If
                        TXTsample.Visible = True
                        TXTsample.Top = GRDSTOCK.CellTop + 100
                        TXTsample.Left = GRDSTOCK.CellLeft '+ 60
                        TXTsample.Width = GRDSTOCK.CellWidth
                        TXTsample.Height = GRDSTOCK.CellHeight
                        TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        TXTsample.SetFocus
                    Case 1, 2, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17, 20, 21, 22, 25, 26, 27
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
                    Case 28, 29
                        CMBCHANGE.Visible = True
                        CMBCHANGE.Top = GRDSTOCK.CellTop + 100
                        CMBCHANGE.Left = GRDSTOCK.CellLeft '+ 60
                        CMBCHANGE.Width = GRDSTOCK.CellWidth
                        'CmbPack.Height = GRDSTOCK.CellHeight
                        On Error Resume Next
                        CMBCHANGE.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        CMBCHANGE.SetFocus
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
                    Case 23
                        FRAME.Visible = True
                        FRAME.Top = GRDSTOCK.CellTop - 300
                        FRAME.Left = GRDSTOCK.CellLeft - 1500
                        'Frame.Width = GRDSTOCK.CellWidth - 25
                        TxtComper.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        If GRDSTOCK.TextMatrix(GRDSTOCK.Row, 24) = "Rs" Then
                            OptAmt.Value = True
                        Else
                            OptPercent.Value = True
                        End If
                        TxtComper.SetFocus
                End Select
            End If
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
        Case vbKeyEscape
                Frmebatch.Visible = False
                TXTEDIT.Visible = False
                TXTEXP.Visible = False
                TXTsample.Visible = False
                CmbPack.Visible = False
                CMBCHANGE.Visible = False
                CMBMFGR.Visible = False
                FRAME.Visible = False
    End Select
End Sub

Private Sub GRDSTOCK_RowColChange()
    lblitemname.Caption = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)
End Sub

Private Sub GRDSTOCK_Scroll()
    Frmebatch.Visible = False
    TXTEDIT.Visible = False
    TXTEXP.Visible = False
    TXTsample.Visible = False
    CmbPack.Visible = False
    CMBCHANGE.Visible = False
    CMBMFGR.Visible = False
    FRAME.Visible = False
End Sub

Private Sub Label1_DblClick(index As Integer)
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If frmunbill.Visible = False Then
        GRDSTOCK.ColWidth(29) = 1200
        frmunbill.Visible = True
        chkunbill.Value = 0
        chkonlyunbill.Value = 0
        cmdchangeunbill.Visible = True
        Label1(43).Visible = True
        Label1(44).Visible = True
        txtstkcrct.Visible = True
    Else
        frmunbill.Visible = False
        chkunbill.Value = 0
        chkonlyunbill.Value = 0
        GRDSTOCK.ColWidth(29) = 0
        cmdchangeunbill.Visible = False
        Label1(43).Visible = False
        Label1(44).Visible = False
        txtstkcrct.Visible = False
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

Private Sub TxtHSNCODE_Change()
    Call tXTMEDICINE_Change
End Sub

Private Sub TXTITEMCODE_Change()
    On Error GoTo ERRHAND
    'Call Fillgrid
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        If REPFLAG = True Then
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        Else
            RSTREP.Close
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        End If
    '===================================================================
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        If REPFLAG = True Then
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        Else
            RSTREP.Close
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        End If
    Else
    '===========================================================================
        If REPFLAG = True Then
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        Else
            RSTREP.Close
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%'  AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        End If
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
    'Call Fillgrid
    If Trim(TxtCode.Text) <> "" Or Trim(TxtName.Text) <> "" Then Call Fillgrid
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        If REPFLAG = True Then
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        Else
            RSTREP.Close
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  UN_BILL = 'Y' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        End If
    '===================================================================
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        If REPFLAG = True Then
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        Else
            RSTREP.Close
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        End If
    Else
    '===========================================================================
        If REPFLAG = True Then
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        Else
            RSTREP.Close
            If CHKCATEGORY2.Value = 0 Then
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND REMARKS Like '%" & Me.TxtHSNCODE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            End If
            REPFLAG = False
        End If
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
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.Text)
    'Call Fillgrid
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
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
    
    If OptAllStock.Value = True Then
        MsgBox "Will be completed soon", , "EzBiz"
        Exit Function
    End If
    Dim itemtable As String
    If OptBrStock.Value = True Then
        lblstock.Caption = "Branch Stock"
        lblstktype.Caption = "B"
        itemtable = "ITEMMASTVAN"
    ElseIf OptAllStock.Value = True Then
        lblstock.Caption = "Main & Branch Stock"
        lblstktype.Caption = "A"
        itemtable = "ITEMMAST"
    Else
        lblstock.Caption = "Main Stock"
        lblstktype.Caption = "M"
        itemtable = "ITEMMAST"
    End If
    
    Frmebatch.Visible = False
    Dim rststock As ADODB.Recordset
    Dim rstopstock As ADODB.Recordset
    Dim i As Long
    
    
    On Error GoTo ERRHAND
    
    i = 0
    Screen.MousePointer = vbHourglass
        
    lblpvalue.Caption = ""
    lblnetvalue.Caption = ""
    lblsalevalue.Caption = ""
    TXTTAX.Text = ""
    GRDSTOCK.rows = 1
    Dim SORT_STRING As String
    
    If OptSortName.Value = True Then
        SORT_STRING = "ITEM_NAME"
    ElseIf OptSortQty.Value = True Then
        SORT_STRING = "CLOSE_QTY"
    ElseIf OptSortPrice.Value = True Then
        SORT_STRING = "ITEM_COST"
    ElseIf OptSortHSN.Value = True Then
        SORT_STRING = "REMARKS"
    ElseIf OptSortTax.Value = True Then
        SORT_STRING = "SALES_TAX"
    ElseIf OptItemcode.Value = True Then
        SORT_STRING = "CONVERT(ITEM_CODE, SIGNED INTEGER)"
    Else
        SORT_STRING = "ITEM_NAME"
    End If
                    
    Set rststock = New ADODB.Recordset
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        If chkdeaditems.Value = 1 Then
            If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                End If
            Else
                If CHKCATEGORY2.Value = 1 Then
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
                Else
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
        
                End If
            End If
        Else
            If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                End If
            Else
                If CHKCATEGORY2.Value = 1 Then
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
                Else
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE UN_BILL = 'Y' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
        
                End If
            End If
        End If
    '==========================================================================================
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        If chkdeaditems.Value = 1 Then
            If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                End If
            Else
                If CHKCATEGORY2.Value = 1 Then
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
                Else
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
        
                End If
            End If
        Else
            If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                End If
            Else
                If CHKCATEGORY2.Value = 1 Then
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
                Else
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
        
                End If
            End If
        End If
    Else
    '=======================================================================================
        If chkdeaditems.Value = 1 Then
            If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                End If
            Else
                If CHKCATEGORY2.Value = 1 Then
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
                Else
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
        
                End If
            End If
        Else
            If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                If OptStock.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                Else
                    rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                End If
            Else
                If CHKCATEGORY2.Value = 1 Then
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y')  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
                Else
                    If OptStock.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    ElseIf OptPC.Value = True Then
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    Else
                        rststock.Open "SELECT * FROM " & itemtable & " WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY " & SORT_STRING & "", db, adOpenForwardOnly
                    End If
        
                End If
            End If
        End If
    End If
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        'GRDSTOCK.FixedCols = 3
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!CLOSE_QTY), 0, Round(rststock!CLOSE_QTY, 3))
        GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
        If Val(GRDSTOCK.TextMatrix(i, 13)) = 0 Then GRDSTOCK.TextMatrix(i, 13) = 1
        GRDSTOCK.TextMatrix(i, 4) = Round(Val(GRDSTOCK.TextMatrix(i, 3)) / Val(GRDSTOCK.TextMatrix(i, 13)), 0)
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!PACK_TYPE), "", rststock!PACK_TYPE)
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.00"))
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 2), "0.000"))
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 2), "0.000"))
        GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!P_VAN), "", Format(rststock!P_VAN, "0.00"))
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!SALES_TAX), "", Format(rststock!SALES_TAX, "0.00"))
        If Val(txtstkcrct.Text) > 0 Then
            GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!ITEM_COST) Or rststock!ITEM_COST = 0, "", Format(Round(rststock!ITEM_COST - (rststock!ITEM_COST * Val(txtstkcrct.Text) / 100), 2), "0.00"))
            GRDSTOCK.TextMatrix(i, 25) = IIf(IsNull(rststock!ITEM_COST) Or rststock!ITEM_COST = 0, "", Format(Round(Val(GRDSTOCK.TextMatrix(i, 11)) * rststock!CLOSE_QTY, 3), "0.00"))
        Else
            GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST, "0.0000"))
            GRDSTOCK.TextMatrix(i, 25) = IIf(IsNull(rststock!ITEM_COST), "", Format(Round(rststock!ITEM_COST * rststock!CLOSE_QTY, 3), "0.00"))
        End If
        GRDSTOCK.TextMatrix(i, 26) = IIf(IsNull(rststock!CESS_PER), "", Format(rststock!CESS_PER, "0.00"))
        GRDSTOCK.TextMatrix(i, 27) = IIf(IsNull(rststock!cess_amt), "", Format(rststock!cess_amt, "0.00"))
        If IsNull(rststock!ITEM_NET_COST) Or rststock!ITEM_NET_COST = 0 Or rststock!ITEM_NET_COST <= Val(GRDSTOCK.TextMatrix(i, 11)) Then
            GRDSTOCK.TextMatrix(i, 12) = Round((Val(GRDSTOCK.TextMatrix(i, 11)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 10)) / 100)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 26)) / 100) + Val(GRDSTOCK.TextMatrix(i, 27)), 4)
        Else
            GRDSTOCK.TextMatrix(i, 12) = IIf(IsNull(rststock!ITEM_NET_COST) Or rststock!ITEM_NET_COST < Val(GRDSTOCK.TextMatrix(i, 11)), Val(GRDSTOCK.TextMatrix(i, 11)), Format(rststock!ITEM_NET_COST, "0.000"))
        End If
        
        GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!REMARKS), "", rststock!REMARKS)
        GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!CRTN_PACK), "", rststock!CRTN_PACK)
        GRDSTOCK.TextMatrix(i, 16) = IIf(IsNull(rststock!P_CRTN), "", Format(Round(rststock!P_CRTN, 2), "0.000"))
        GRDSTOCK.TextMatrix(i, 17) = IIf(IsNull(rststock!P_LWS), "", Format(Round(rststock!P_LWS, 2), "0.000"))
        GRDSTOCK.TextMatrix(i, 18) = IIf(IsNull(rststock!Category), "", rststock!Category)
        'GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
        GRDSTOCK.TextMatrix(i, 19) = IIf(IsNull(rststock!MANUFACTURER), "", rststock!MANUFACTURER)
        If Val(GRDSTOCK.TextMatrix(i, 11)) <> 0 Then
            GRDSTOCK.TextMatrix(i, 20) = Round((((Val(GRDSTOCK.TextMatrix(i, 7)) / Val(GRDSTOCK.TextMatrix(i, 13))) - Val(GRDSTOCK.TextMatrix(i, 12))) * 100) / Val(GRDSTOCK.TextMatrix(i, 12)), 2)
        Else
            GRDSTOCK.TextMatrix(i, 20) = 0
        End If
        GRDSTOCK.TextMatrix(i, 21) = IIf(IsNull(rststock!CUST_DISC), "", Format(rststock!CUST_DISC, "0.00"))
        GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!DISC_AMT), "", Format(rststock!DISC_AMT, "0.00"))
        Select Case rststock!COM_FLAG
            Case "P"
                GRDSTOCK.TextMatrix(i, 23) = IIf(IsNull(rststock!COM_PER), "", Format(rststock!COM_PER, "0.00"))
                GRDSTOCK.TextMatrix(i, 24) = "%"
            Case "A"
                GRDSTOCK.TextMatrix(i, 23) = IIf(IsNull(rststock!COM_AMT), "", Format(rststock!COM_AMT, "0.00"))
                GRDSTOCK.TextMatrix(i, 24) = "Rs"
        End Select
        
        Select Case rststock!PRICE_CHANGE
            Case "Y"
                GRDSTOCK.TextMatrix(i, 28) = "Yes"
            Case Else
                GRDSTOCK.TextMatrix(i, 28) = "No"
        End Select
        
        Select Case rststock!UN_BILL
            Case "Y"
                GRDSTOCK.TextMatrix(i, 29) = "Yes"
            Case Else
                GRDSTOCK.TextMatrix(i, 29) = "No"
        End Select
        GRDSTOCK.TextMatrix(i, 30) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
        lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 25)), "0.00")
        lblnetvalue.Caption = Format(Round(Val(lblnetvalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 12)) * Val(GRDSTOCK.TextMatrix(i, 3))), 2), "0.00")
        lblsalevalue.Caption = Format(Round(Val(lblsalevalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 7)) * Val(GRDSTOCK.TextMatrix(i, 3)) / Val(GRDSTOCK.TextMatrix(i, 13))), 2), "0.00")
        
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
    'Call Toatal_value
    
    DTFROM.Value = Null
    Screen.MousePointer = vbNormal
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 1  ' Item Code
                    If Trim(TXTsample.Text) = "" Then Exit Sub
                    db.Execute "Update ITEMMAST set ITEM_CODE = '" & Trim(TXTsample.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    db.Execute "Update RTRXFILE set ITEM_CODE = '" & Trim(TXTsample.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    db.Execute "Update TRXFILE set ITEM_CODE = '" & Trim(TXTsample.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    db.Execute "Update TRXFORMULASUB set ITEM_CODE = '" & Trim(TXTsample.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    db.Execute "Update TRXFORMULASUB set FOR_NAME = '" & Trim(TXTsample.Text) & "' where FOR_NAME = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    db.Execute "Update TRXFORMULAMAST set ITEM_CODE = '" & Trim(TXTsample.Text) & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                                        
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 2  ' Item Name
                    If Trim(TXTsample.Text) = "" Then Exit Sub
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
                        rststock!ITEM_COST = 0
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
                        If frmunbill.Visible = True Then
                            rststock!UN_BILL = "Y"
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
                    'Dim RSTITEMMAST As ADODB.Recordset
                    'Dim TRXMAST As ADODB.Recordset
                    'Dim rststock As ADODB.Recordset
                    
                    Screen.MousePointer = vbHourglass
                    Set RSTITEMMAST = New ADODB.Recordset
                    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                        INWARD = 0
                        OUTWARD = 0
                        BAL_QTY = 0
                        
'                        Set TRXMAST = New ADODB.Recordset
'                        TRXMAST.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE = '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE = 'ST' ", db, adOpenStatic, adLockOptimistic, adCmdText
'                        Set rststock = New ADODB.Recordset
'                        If TRXMAST.RecordCount > 0 Then
'                            rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE <> '" & Format(DTFROM.value, "yyyy/mm/dd") & "' and TRX_TYPE <> 'ST'", db, adOpenStatic, adLockReadOnly
'                        Else
'                            rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
'                        End If
'
'                        TRXMAST.Close
'                        Set TRXMAST = Nothing
                        
                        db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY <0"
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
                        
                        If Not (Val(TXTsample.Text) - (Val(INWARD - OUTWARD)) = 0) Then
                            Dim BARCODE As String
                            BARCODE = ""
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT BARCODE FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ORDER BY TRX_YEAR DESC, VCH_NO DESC, LINE_NO DESC", db, adOpenStatic, adLockReadOnly
                            If Not (rststock.EOF And rststock.BOF) Then
                                BARCODE = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
                            End If
                            rststock.Close
                            Set rststock = Nothing
                        
                            Set rststock = New ADODB.Recordset
                            'rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE = '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE = 'ST' ", db, adOpenStatic, adLockOptimistic, adCmdText
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
                            rststock!ITEM_COST = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11))
                            rststock!LINE_DISC = 1
                            rststock!P_DISC = 0
                            rststock!BARCODE = BARCODE
                            rststock!MRP = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6))
                            rststock!PTR = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11))
                            rststock!SALES_PRICE = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
                            rststock!P_RETAIL = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7))
                            rststock!P_WS = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
                            rststock!P_VAN = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9))
                            rststock!P_CRTN = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16))
                            rststock!P_LWS = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 17))
                            rststock!CRTN_PACK = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15))
                            rststock!Category = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 18))
                            rststock!CESS_PER = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 26))
                            rststock!cess_amt = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 27))
                            rststock!gross_amt = 0
                            rststock!COM_FLAG = "P"
                            rststock!COM_PER = 0
                            rststock!COM_AMT = 0
                            rststock!SALES_TAX = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10))
                            rststock!LOOSE_PACK = RSTITEMMAST!LOOSE_PACK
                            rststock!PACK_TYPE = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5))
                            'rststock!WARRANTY = Null
                            'rststock!WARRANTY_TYPE = Null
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
                            
                            Set rststock = New ADODB.Recordset
                            rststock.Open "Select * From TRANSMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='ST' AND VCH_NO = " & BILL_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
                            If (rststock.EOF And rststock.BOF) Then
                                rststock.AddNew
                                rststock!TRX_TYPE = "ST"
                                rststock!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                                rststock!VCH_NO = BILL_NO
                                rststock!C_USER_ID = frmLogin.rs!USER_ID
                                rststock!CREATE_DATE = Format(Date, "DD/MM/YYYY")
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                            
                            db.Execute "Update ITEMMAST set CLOSE_QTY = " & Val(TXTsample.Text) & ", RCPT_QTY = " & INWARD & " + " & Val(TXTsample.Text) & ",  ISSUE_QTY = " & OUTWARD & " WHERE  ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                            'RSTITEMMAST!CLOSE_QTY = Val(TXTsample.Text)
                            'RSTITEMMAST!RCPT_QTY = INWARD + Val(TXTsample.Text)
                            'RSTITEMMAST!ISSUE_QTY = OUTWARD
                            'RSTITEMMAST.Update
                        End If
                        RSTITEMMAST.Close
                        Set RSTITEMMAST = Nothing
                    End If
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 25) = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3))
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 4) = Round(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)) / GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13), 0)
                    Call Toatal_value
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    Screen.MousePointer = vbNormal
                    
                Case 7  'RT
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT P_RETAIL from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_RETAIL) AND P_RETAIL <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
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
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) = 0 Then
                            rststock!P_CRTN = Val(TXTsample.Text)
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) = Format(Val(TXTsample.Text), "0.000")
                        Else
                            rststock!P_CRTN = Round(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)), 2)
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) = Format(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)), "0.000")
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
                    rststock.Open "SELECT DISTINCT P_WS from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_WS) AND P_WS <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
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
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) = 0 Then
                            rststock!P_LWS = Val(TXTsample.Text)
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 17) = Format(Val(TXTsample.Text), "0.000")
                        Else
                            rststock!P_LWS = Round(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)), 2)
                            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 17) = Format(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)), "0.000")
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
                    rststock.Open "SELECT DISTINCT P_CRTN from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_CRTN) AND P_CRTN <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
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
                    rststock.Open "SELECT DISTINCT P_LWS from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_LWS) AND P_LWS <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
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
                
                Case 9  'VAN
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT DISTINCT P_VAN from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_VAN) AND P_VAN <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                    If rststock.RecordCount > 1 Then
                        If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                            rststock.Close
                            Set rststock = Nothing
                            TXTsample.SetFocus
                            Exit Sub
                        End If
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
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
                        rststock!P_VAN = Val(TXTsample.Text)
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
                        rststock!P_VAN = Val(TXTsample.Text)
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
                        rststock!ITEM_COST = Val(TXTsample.Text)
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
                    
                    If MsgBox("Do you want to reset the Cost in available stocks of Opening Stock", vbYesNo + vbDefaultButton2, "Price Change") = vbYes Then
                        db.Execute "Update RTRXFILE set GROSS_AMT = (QTY * " & Val(TXTsample.Text) & "), TRX_TOTAL = (QTY * " & Val(TXTsample.Text) & ")+((QTY * " & Val(TXTsample.Text) & ") * SALES_TAX / 100), ITEM_COST = " & Val(TXTsample.Text) & ", PTR = " & Val(TXTsample.Text) & " where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND BAL_QTY >0 AND (TRX_TYPE = 'OP' OR TRX_TYPE = 'ST')"
                    End If
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 25) = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3))
                    Call Toatal_value
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 25  'VALUE
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)) <> 0 Then
                            rststock!ITEM_COST = Round(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), 3)
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
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 25) = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3))
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
                    Call Toatal_value
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
                        rststock!DISC_AMT = 0
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 22) = ""
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 22  'Cust Disc Amt
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!DISC_AMT = Val(TXTsample.Text)
                        rststock!CUST_DISC = 0
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.00")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 21) = ""
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 14  'HSN CODE
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!REMARKS = Trim(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 26  'CESS%
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
                Case 27  'CESS RATE
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
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 3, 6, 7, 8, 9, 10, 11, 13, 15, 16, 17, 20, 21, 25, 26
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        
        Case 1, 2, 14
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub TXTCODE_Change()
    Call tXTMEDICINE_Change
    Exit Sub
    On Error GoTo ERRHAND
    If Trim(tXTMEDICINE.Text) <> "" Or Trim(TxtName.Text) <> "" Then Call Fillgrid
    'Call Fillgrid
    If REPFLAG = True Then
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
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
    TxtCode.SelLength = Len(TxtCode.Text)
    'Call Fillgrid
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'TxtItemcode.SetFocus
            Call CmdLoad_Click
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

Private Sub TxtComper_GotFocus()
    TxtComper.SelStart = 0
    TxtComper.SelLength = Len(TxtComper.Text)
End Sub

Private Sub TxtComper_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
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
        Case vbKeyReturn
            TxtComper.SetFocus
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub OptAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
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
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 23) = "0.00"
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 24) = ""
        Else
            If OptAmt.Value = True Then
                rststock!COM_FLAG = "A"
                rststock!COM_PER = 0
                rststock!COM_AMT = Val(TxtComper.Text)
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 23) = Format(Val(TxtComper.Text), "0.00")
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 24) = "Rs"
            Else
                rststock!COM_FLAG = "P"
                rststock!COM_PER = Val(TxtComper.Text)
                rststock!COM_AMT = 0
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 23) = Format(Val(TxtComper.Text), "0.00")
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 24) = "%"
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
    Dim i As Long
    lblpvalue.Caption = ""
    lblnetvalue.Caption = ""
    lblsalevalue.Caption = ""
    For i = 1 To GRDSTOCK.rows - 1
        lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 25)), "0.00")
'        If (Val(GRDSTOCK.TextMatrix(i, 12)) * Val(GRDSTOCK.TextMatrix(i, 3))) <> Val(GRDSTOCK.TextMatrix(i, 25)) Then
'            MsgBox ""
'        End If
        lblnetvalue.Caption = Format(Round(Val(lblnetvalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 12)) * Val(GRDSTOCK.TextMatrix(i, 3))), 2), "0.00")
        lblsalevalue.Caption = Format(Round(Val(lblsalevalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 7)) * Val(GRDSTOCK.TextMatrix(i, 3)) / Val(GRDSTOCK.TextMatrix(i, 13))), 2), "0.00")
    Next i
End Function

Private Sub cmddelphoto_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    On Error GoTo errHandler
    CommonDialog1.FileName = ""
    Set Image1.DataSource = Nothing
    Image1.Picture = LoadPicture("")
    
    bytData = ""
    Dim RSTITEMMAST As ADODB.Recordset
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        Frame6.Visible = False
        RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Exit Sub
errHandler:
    MsgBox "Unexpected error. Err " & err & " : " & Error
End Sub

Private Sub CMDBROWSE_Click()
    If lblstktype.Caption <> "M" Then Exit Sub
    Dim bytData() As Byte
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
    
    Dim RSTITEMMAST As ADODB.Recordset
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        Frame6.Visible = True
        RSTITEMMAST.Fields("PHOTO").AppendChunk bytData
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Exit Sub
errHandler:
    Select Case err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        MsgBox "Unexpected error. Err " & err & " : " & Error
    End Select
End Sub


Private Sub TXTDEALER2_Change()
    
    On Error GoTo ERRHAND
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
ERRHAND:
    MsgBox err.Description
    
End Sub


Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.Text)
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
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("*")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_Click()
        
    TXTDEALER2.Text = DataList1.Text
    LBLDEALER2.Caption = TXTDEALER2.Text
    'Call Fillgrid
    tXTMEDICINE.SetFocus
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
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

Private Sub TxtItemcode_GotFocus()
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
    'Call Fillgrid
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Dim itemtable As String
            OptMainStk.Value = True
            lblstock.Caption = "Main Stock"
            lblstktype.Caption = "M"

    
            If Trim(TXTITEMCODE.Text) = "" Then
                CmdLoad.SetFocus
                Exit Sub
            End If
            Frmebatch.Visible = False
            Dim rststock As ADODB.Recordset
            Dim rstopstock As ADODB.Recordset
            Dim i As Long
        
            On Error GoTo ERRHAND
            
            i = 0
            Screen.MousePointer = vbHourglass
                
            lblpvalue.Caption = ""
            lblnetvalue.Caption = ""
            lblsalevalue.Caption = ""
            TXTTAX.Text = ""
            GRDSTOCK.rows = 1
            Set rststock = New ADODB.Recordset
            If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
                If chkdeaditems.Value = 1 Then
                    If OptSortQty.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                    ElseIf OptSortPrice.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    Else
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    End If
                Else
                    If OptSortQty.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                    ElseIf OptSortPrice.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    Else
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' and (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE UN_BILL = 'Y' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    End If
                End If
            '==========================================================================================
            ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
                If chkdeaditems.Value = 1 Then
                    If OptSortQty.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                    ElseIf OptSortPrice.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    Else
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    End If
                Else
                    If OptSortQty.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                    ElseIf OptSortPrice.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    Else
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    End If
                End If
            Else
            '=======================================================================================
                If chkdeaditems.Value = 1 Then
                    If OptSortQty.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                    ElseIf OptSortPrice.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    Else
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    End If
                Else
                    If OptSortQty.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                    ElseIf OptSortPrice.Value = True Then
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_COST", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_COST", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_COST", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    Else
                        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                            If OptStock.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            ElseIf OptPC.Value = True Then
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' and (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            Else
                                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                            End If
                        Else
                            If CHKCATEGORY2.Value = 1 Then
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "')  AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                            Else
                                If OptStock.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                ElseIf OptPC.Value = True Then
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND PRICE_CHANGE = 'Y' and  ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                Else
                                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND (ITEM_CODE Like '" & Me.TXTITEMCODE.Text & "%' OR BARCODE Like '" & Me.TXTITEMCODE.Text & "') AND (ISNULL(REMARKS) OR REMARKS Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                                End If
                    
                            End If
                        End If
                
                    End If
                End If
            End If
            Do Until rststock.EOF
                i = i + 1
                GRDSTOCK.rows = GRDSTOCK.rows + 1
                GRDSTOCK.FixedRows = 1
                'GRDSTOCK.FixedCols = 3
                GRDSTOCK.TextMatrix(i, 0) = i
                GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
                GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
                GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!CLOSE_QTY), 0, Round(rststock!CLOSE_QTY, 3))
                GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                If Val(GRDSTOCK.TextMatrix(i, 13)) = 0 Then GRDSTOCK.TextMatrix(i, 13) = 1
                GRDSTOCK.TextMatrix(i, 4) = Round(Val(GRDSTOCK.TextMatrix(i, 3)) / Val(GRDSTOCK.TextMatrix(i, 13)), 0)
                GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!PACK_TYPE), "", rststock!PACK_TYPE)
                GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.00"))
                GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!P_VAN), "", Format(rststock!P_VAN, "0.00"))
                GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!SALES_TAX), "", Format(rststock!SALES_TAX, "0.00"))
                'GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!item_COST), "", Format(rststock!item_COST, "0.0000"))
                If Val(txtstkcrct.Text) > 0 Then
                    GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!ITEM_COST) Or rststock!ITEM_COST = 0, "", Format(Round(rststock!ITEM_COST - (rststock!ITEM_COST * Val(txtstkcrct.Text) / 100), 2), "0.00"))
                    GRDSTOCK.TextMatrix(i, 25) = IIf(IsNull(rststock!ITEM_COST) Or rststock!ITEM_COST = 0, "", Format(Round(Val(GRDSTOCK.TextMatrix(i, 11)) * rststock!CLOSE_QTY, 3), "0.00"))
                Else
                    GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST, "0.0000"))
                    GRDSTOCK.TextMatrix(i, 25) = IIf(IsNull(rststock!ITEM_COST), "", Format(Round(rststock!ITEM_COST * rststock!CLOSE_QTY, 3), "0.00"))
                End If
                GRDSTOCK.TextMatrix(i, 25) = IIf(IsNull(rststock!ITEM_COST), "", Format(Round(rststock!ITEM_COST * rststock!CLOSE_QTY, 3), "0.00"))
                GRDSTOCK.TextMatrix(i, 26) = IIf(IsNull(rststock!CESS_PER), "", Format(rststock!CESS_PER, "0.00"))
                GRDSTOCK.TextMatrix(i, 27) = IIf(IsNull(rststock!cess_amt), "", Format(rststock!cess_amt, "0.00"))
                'GRDSTOCK.TextMatrix(i, 12) = Round((Val(GRDSTOCK.TextMatrix(i, 11)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 10)) / 100)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 26)) / 100) + Val(GRDSTOCK.TextMatrix(i, 27)), 4)
                If IsNull(rststock!ITEM_NET_COST) Or rststock!ITEM_NET_COST = 0 Or rststock!ITEM_NET_COST <= Val(GRDSTOCK.TextMatrix(i, 11)) Then
                    GRDSTOCK.TextMatrix(i, 12) = Round((Val(GRDSTOCK.TextMatrix(i, 11)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 10)) / 100)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 26)) / 100) + Val(GRDSTOCK.TextMatrix(i, 27)), 4)
                Else
                    GRDSTOCK.TextMatrix(i, 12) = IIf(IsNull(rststock!ITEM_NET_COST) Or rststock!ITEM_NET_COST < Val(GRDSTOCK.TextMatrix(i, 11)), Val(GRDSTOCK.TextMatrix(i, 11)), Format(rststock!ITEM_NET_COST, "0.000"))
                End If
                GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!REMARKS), "", rststock!REMARKS)
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
                GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!DISC_AMT), "", Format(rststock!DISC_AMT, "0.00"))
                Select Case rststock!COM_FLAG
                    Case "P"
                        GRDSTOCK.TextMatrix(i, 23) = IIf(IsNull(rststock!COM_PER), "", Format(rststock!COM_PER, "0.00"))
                        GRDSTOCK.TextMatrix(i, 24) = "%"
                    Case "A"
                        GRDSTOCK.TextMatrix(i, 23) = IIf(IsNull(rststock!COM_AMT), "", Format(rststock!COM_AMT, "0.00"))
                        GRDSTOCK.TextMatrix(i, 24) = "Rs"
                End Select
                
                Select Case rststock!PRICE_CHANGE
                    Case "Y"
                        GRDSTOCK.TextMatrix(i, 28) = "Yes"
                    Case Else
                        GRDSTOCK.TextMatrix(i, 28) = "No"
                End Select
                
                Select Case rststock!UN_BILL
                    Case "Y"
                        GRDSTOCK.TextMatrix(i, 29) = "Yes"
                    Case Else
                        GRDSTOCK.TextMatrix(i, 29) = "No"
                End Select
                GRDSTOCK.TextMatrix(i, 30) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
                lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 25)), "0.00")
                lblnetvalue.Caption = Format(Round(Val(lblnetvalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 12)) * Val(GRDSTOCK.TextMatrix(i, 3))), 2), "0.00")
                lblsalevalue.Caption = Format(Round(Val(lblsalevalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 7)) * Val(GRDSTOCK.TextMatrix(i, 3)) / Val(GRDSTOCK.TextMatrix(i, 13))), 2), "0.00")
                
                
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
            'Call Toatal_value
            
            DTFROM.Value = Null
            Screen.MousePointer = vbNormal
    
        Case vbKeyEscape
            TxtCode.SetFocus
    End Select
Exit Sub

ERRHAND:
    Screen.MousePointer = vbNormal
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

Private Sub TxtTax_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
    'Call Fillgrid
End Sub

Private Sub TxtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ERRHAND
            Dim rststock As ADODB.Recordset
            Dim i As Long
        
            On Error GoTo ERRHAND
            OptMainStk.Value = True
            lblstock.Caption = "Main Stock"
            lblstktype.Caption = "M"

            i = 0
            Screen.MousePointer = vbHourglass
            
            GRDSTOCK.rows = 1
            Set rststock = New ADODB.Recordset
            If OptStock.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND SALES_TAX = " & Val(TXTTAX.Text) & " AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY CLOSE_QTY", db, adOpenForwardOnly
            ElseIf OptPC.Value = True Then
                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND PRICE_CHANGE = 'Y' AND SALES_TAX = " & Val(TXTTAX.Text) & " AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
            Else
                rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND SALES_TAX = " & Val(TXTTAX.Text) & " AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY CLOSE_QTY", db, adOpenForwardOnly
            End If
            Do Until rststock.EOF
                i = i + 1
                GRDSTOCK.rows = GRDSTOCK.rows + 1
                GRDSTOCK.FixedRows = 1
                'GRDSTOCK.FixedCols = 3
                GRDSTOCK.TextMatrix(i, 0) = i
                GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
                GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
                GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!CLOSE_QTY), 0, Round(rststock!CLOSE_QTY, 3))
                GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                If Val(GRDSTOCK.TextMatrix(i, 13)) = 0 Then GRDSTOCK.TextMatrix(i, 13) = 1
                GRDSTOCK.TextMatrix(i, 4) = Round(Val(GRDSTOCK.TextMatrix(i, 3)) / Val(GRDSTOCK.TextMatrix(i, 13)), 0)
                GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!PACK_TYPE), "", rststock!PACK_TYPE)
                GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.00"))
                GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_RETAIL), "", Format(Round(rststock!P_RETAIL, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!P_WS), "", Format(Round(rststock!P_WS, 1), "0.000"))
                GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!P_VAN), "", Format(rststock!P_VAN, "0.00"))
                GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!SALES_TAX), "", Format(rststock!SALES_TAX, "0.00"))
                GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST, "0.0000"))
                GRDSTOCK.TextMatrix(i, 25) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST * rststock!CLOSE_QTY, "0.00"))
                GRDSTOCK.TextMatrix(i, 26) = IIf(IsNull(rststock!CESS_PER), "", Format(rststock!CESS_PER, "0.00"))
                GRDSTOCK.TextMatrix(i, 27) = IIf(IsNull(rststock!cess_amt), "", Format(rststock!cess_amt, "0.00"))
                GRDSTOCK.TextMatrix(i, 12) = Round((Val(GRDSTOCK.TextMatrix(i, 11)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 10)) / 100)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 26)) / 100) + Val(GRDSTOCK.TextMatrix(i, 27)), 4)
                GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!REMARKS), "", rststock!REMARKS)
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
                GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!DISC_AMT), "", Format(rststock!DISC_AMT, "0.00"))
                Select Case rststock!COM_FLAG
                    Case "P"
                        GRDSTOCK.TextMatrix(i, 23) = IIf(IsNull(rststock!COM_PER), "", Format(rststock!COM_PER, "0.00"))
                        GRDSTOCK.TextMatrix(i, 24) = "%"
                    Case "A"
                        GRDSTOCK.TextMatrix(i, 23) = IIf(IsNull(rststock!COM_AMT), "", Format(rststock!COM_AMT, "0.00"))
                        GRDSTOCK.TextMatrix(i, 24) = "Rs"
                End Select
                
                Select Case rststock!PRICE_CHANGE
                    Case "Y"
                        GRDSTOCK.TextMatrix(i, 28) = "Yes"
                    Case Else
                        GRDSTOCK.TextMatrix(i, 28) = "No"
                End Select
                
                Select Case rststock!UN_BILL
                    Case "Y"
                        GRDSTOCK.TextMatrix(i, 29) = "Yes"
                    Case Else
                        GRDSTOCK.TextMatrix(i, 29) = "No"
                End Select
                
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
            TXTTAX.SelStart = 0
            TXTTAX.SelLength = Len(TXTTAX.Text)
        Case vbKeyEscape
            TXTITEMCODE.SetFocus
    End Select
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description

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
    TxtHSNCODE.SelLength = Len(TxtHSNCODE.Text)
    'Call Fillgrid
End Sub

Private Sub TxtHSNCODE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmdLoad_Click
        Case vbKeyEscape
            TXTTAX.SetFocus
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

Private Sub GRDSTOCK_DblClick()
    If lblstktype.Caption <> "M" Then Exit Sub
    On Error GoTo ERRHAND
    Select Case GRDSTOCK.Col
        Case 1
            frmitemmaster.TXTITEMCODE.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)
            Call frmitemmaster.TxtItemcode_KeyDown(13, 0)
            frmitemmaster.Show
            frmitemmaster.SetFocus
            Exit Sub
        Case 2
            FrmStkmovmnt.Show
            FrmStkmovmnt.SetFocus
            FrmStkmovmnt.tXTMEDICINE.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)
            FrmStkmovmnt.LBLITEMCODE.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)
            FrmStkmovmnt.DataList2.BoundText = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)
            Call FrmStkmovmnt.DataList2_Click
            Exit Sub
        Case Else
            LBLitem.Caption = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)
            LBLITEMCODE.Caption = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)
            Call Fill_Batchdetails
            lblqty.Caption = "QTY: " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)
            Frmebatch.Visible = True
            grdbatch.SetFocus
            Exit Sub
    End Select
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub grdbatch_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If grdbatch.rows = 1 Then Exit Sub
            If Not (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then Exit Sub
            Select Case grdbatch.Col
                'Case 3 '' balQty
                
                Case 1, 2, 3, 6, 11
                    TXTEDIT.Visible = True
                    TXTEDIT.Top = grdbatch.CellTop + 450
                    TXTEDIT.Left = grdbatch.CellLeft + 20
                    TXTEDIT.Width = grdbatch.CellWidth
                    TXTEDIT.Height = grdbatch.CellHeight
                    TXTEDIT.Text = grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col)
                    TXTEDIT.SetFocus
                Case 7
                    TXTEDIT.Visible = True
                    TXTEDIT.Top = grdbatch.CellTop + 450
                    TXTEDIT.Left = grdbatch.CellLeft + 20
                    TXTEDIT.Width = grdbatch.CellWidth
                    TXTEDIT.Height = grdbatch.CellHeight
                    If Trim(grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col)) = "" Then
                        TXTEDIT.Text = Trim(LBLITEMCODE.Caption) & Val(grdbatch.TextMatrix(grdbatch.Row, 3))
                    Else
                        TXTEDIT.Text = grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col)
                    End If
                    TXTEDIT.SetFocus
                    
                Case 5
                    TXTEDIT.Visible = True
                    TXTEDIT.Top = grdbatch.CellTop + 450
                    TXTEDIT.Left = grdbatch.CellLeft + 20
                    TXTEDIT.Width = grdbatch.CellWidth
                    TXTEDIT.Height = grdbatch.CellHeight
                    TXTEDIT.Text = grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col)
                    TXTEDIT.SetFocus
                Case 4
                    TXTEXP.Visible = True
                    TXTEXP.Top = grdbatch.CellTop + 450
                    TXTEXP.Left = grdbatch.CellLeft + 20
                    TXTEXP.Width = grdbatch.CellWidth '- 25
                    TXTEXP.Text = IIf(IsDate(grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col)), Format(grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col), "MM/YY"), "  /  ")
                    TXTEXP.SetFocus
            End Select
        Case vbKeyEscape
            Frmebatch.Visible = False
            TXTEDIT.Visible = False
            TXTEXP.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub grdbatch_Scroll()
    TXTEDIT.Visible = False
    TXTEXP.Visible = False
    grdbatch.SetFocus
End Sub

Private Sub TXTEDIT_GotFocus()
    TXTEDIT.SelStart = 0
    TXTEDIT.SelLength = Len(TXTEDIT.Text)
End Sub

Private Sub TXTEDIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim M_STOCK As Double
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            Select Case grdbatch.Col
                Case 1  ' Bal QTY
                    grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col) = TXTEDIT.Text
                    db.Execute "UPDATE RTRXFILE SET BAL_QTY = " & Val(TXTEDIT.Text) & " WHERE ITEM_CODE = '" & LBLITEMCODE.Caption & "' AND VCH_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 9)) & " AND LINE_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 10)) & " AND TRX_TYPE = '" & grdbatch.TextMatrix(grdbatch.Row, 8) & "' "
                    grdbatch.Enabled = True
                    TXTEDIT.Visible = False
                    grdbatch.SetFocus
                Case 6  'PACK
                    If Val(TXTEDIT.Text) = 0 Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & LBLITEMCODE.Caption & "' AND RTRXFILE.VCH_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 10)) & " AND TRX_TYPE = '" & grdbatch.TextMatrix(grdbatch.Row, 8) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        'rststock!P_RETAIL = (rststock!P_RETAIL * rststock!LOOSE_PACK) / Val(TXTEDIT.Text)
                        rststock!ITEM_COST = (rststock!ITEM_COST * rststock!LOOSE_PACK) / Val(TXTEDIT.Text)
                        rststock!P_RETAIL = rststock!MRP
                        rststock!P_CRTN = Round(rststock!MRP / Val(TXTEDIT.Text), 3)
                        rststock!BAL_QTY = (rststock!BAL_QTY / IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, 1, rststock!LOOSE_PACK)) * Val(TXTEDIT.Text)
                        rststock!QTY = (rststock!QTY / IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, 1, rststock!LOOSE_PACK)) * Val(TXTEDIT.Text)
                        grdbatch.TextMatrix(grdbatch.Row, 1) = rststock!BAL_QTY
                        rststock!LOOSE_PACK = Val(TXTEDIT.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
            
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & LBLITEMCODE.Caption & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_COST = (rststock!ITEM_COST * IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, 1, rststock!LOOSE_PACK)) / Val(TXTEDIT.Text)
                        rststock!P_RETAIL = rststock!MRP
                        rststock!P_CRTN = Round(rststock!MRP / Val(TXTEDIT.Text), 3)
                        rststock!LOOSE_PACK = Val(TXTEDIT.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col) = TXTEDIT.Text
                    grdbatch.Enabled = True
                    TXTEDIT.Visible = False
                    grdbatch.SetFocus
                    
                Case 2  'RT
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & LBLITEMCODE.Caption & "' AND RTRXFILE.VCH_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 10)) & " AND TRX_TYPE = '" & grdbatch.TextMatrix(grdbatch.Row, 8) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_RETAIL = Val(TXTEDIT.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & LBLITEMCODE.Caption & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_RETAIL = Val(TXTEDIT.Text)
                        grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col) = Format(Val(TXTEDIT.Text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    grdbatch.Enabled = True
                    TXTEDIT.Visible = False
                    grdbatch.SetFocus
                    
                Case 11  'WS
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & LBLITEMCODE.Caption & "' AND RTRXFILE.VCH_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 10)) & " AND TRX_TYPE = '" & grdbatch.TextMatrix(grdbatch.Row, 8) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_WS = Val(TXTEDIT.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & LBLITEMCODE.Caption & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_WS = Val(TXTEDIT.Text)
                        grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col) = Format(Val(TXTEDIT.Text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    grdbatch.Enabled = True
                    TXTEDIT.Visible = False
                    grdbatch.SetFocus
                    
                Case 3  'MRP
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & LBLITEMCODE.Caption & "' AND RTRXFILE.VCH_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 10)) & " AND TRX_TYPE = '" & grdbatch.TextMatrix(grdbatch.Row, 8) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTEDIT.Text)
                        rststock!P_RETAIL = Val(TXTEDIT.Text)
                        rststock!P_CRTN = Round(Val(TXTEDIT.Text) / IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, 1, rststock!LOOSE_PACK), 3)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing

                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & LBLITEMCODE.Caption & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTEDIT.Text)
                        rststock!P_RETAIL = Val(TXTEDIT.Text)
                        rststock!P_CRTN = Round(Val(TXTEDIT.Text) / IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, 1, rststock!LOOSE_PACK), 3)
                        grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col) = Format(Val(TXTEDIT.Text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    grdbatch.Enabled = True
                    TXTEDIT.Visible = False
                    grdbatch.SetFocus
                    
                
                Case 7  'Barcode
                    If Trim(TXTEDIT.Text) = "" Then
                        TXTEDIT.Text = Trim(LBLITEMCODE.Caption) & Val(grdbatch.TextMatrix(grdbatch.Row, 5))
                        If BARTEMPLATE = "Y" And Len(TXTEDIT.Text) Mod 2 <> 0 Then TXTEDIT.Text = TXTEDIT.Text & "9"
                    End If
                    Dim rstTRXMAST As ADODB.Recordset
                    Set rstTRXMAST = New ADODB.Recordset
                    rstTRXMAST.Open "Select * From RTRXFILE WHERE BARCODE= '" & Trim(TXTEDIT.Text) & "' AND ITEM_CODE <> '" & LBLITEMCODE.Caption & "' ", db, adOpenStatic, adLockReadOnly
                    If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                        MsgBox "This BARCODE is already being assigned to another Item", vbOKOnly, "Barcode Entry"
                        TXTEDIT.SetFocus
                        rstTRXMAST.Close
                        Set rstTRXMAST = Nothing
                        Exit Sub
                    End If
                    rstTRXMAST.Close
                    Set rstTRXMAST = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & LBLITEMCODE.Caption & "' AND RTRXFILE.VCH_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 10)) & " AND TRX_TYPE = '" & grdbatch.TextMatrix(grdbatch.Row, 8) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!BARCODE = Trim(TXTEDIT.Text)
                        grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col) = Trim(TXTEDIT.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                
                    grdbatch.Enabled = True
                    TXTEDIT.Visible = False
                    grdbatch.SetFocus
                Case 5  'REF
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & LBLITEMCODE.Caption & "' AND RTRXFILE.VCH_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 10)) & " AND TRX_TYPE = '" & grdbatch.TextMatrix(grdbatch.Row, 8) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!REF_NO = Trim(TXTEDIT.Text)
                        grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col) = Trim(TXTEDIT.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                
                    grdbatch.Enabled = True
                    TXTEDIT.Visible = False
                    grdbatch.SetFocus
            End Select
        Case vbKeyEscape
            TXTEDIT.Visible = False
            grdbatch.SetFocus
    End Select
        Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTEDIT_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 1, 3, 6
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 5, 7
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
    End Select
End Sub
Private Sub TXTEXP_GotFocus()
    TXTEXP.SelStart = 0
    TXTEXP.SelLength = Len(TXTEXP.Text)
End Sub

Private Sub TXTEXP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim M_DATE As Date
    Dim D As Integer
    Dim M As Integer
    Dim Y As Integer
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(Mid(TXTEXP.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXP.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXP.Text, 4, 5)) = 0 Then Exit Sub
            
            M = Val(Mid(TXTEXP.Text, 1, 2))
            Y = Val(Right(TXTEXP.Text, 2))
            Y = 2000 + Y
            M_DATE = "01" & "/" & M & "/" & Y
            D = LastDayOfMonth(M_DATE)
            M_DATE = D & "/" & M & "/" & Y
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & LBLITEMCODE.Caption & "' AND RTRXFILE.VCH_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(grdbatch.TextMatrix(grdbatch.Row, 10)) & " AND TRX_TYPE = '" & grdbatch.TextMatrix(grdbatch.Row, 8) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (rststock.EOF And rststock.BOF) Then
                rststock!EXP_DATE = Format(M_DATE, "dd/mm/yyyy")
                'rststock!VCH_DATE = Format(M_DATE, "dd/mm/yyyy")
                rststock.Update
            End If
            rststock.Close
            Set rststock = Nothing
            
            TXTEXP.Visible = False
            grdbatch.TextMatrix(grdbatch.Row, grdbatch.Col) = M_DATE
            grdbatch.Enabled = True
            grdbatch.SetFocus
        Case vbKeyEscape
            TXTEXP.Visible = False
            grdbatch.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub


Private Function Fill_Batchdetails()
'    db.Execute "Update rtrxfile set MFGR = '' where isnull(MFGR) "
'    db.Execute "Update rtrxfile set CATEGORY = '' where isnull(CATEGORY) "
'    db.Execute "Update rtrxfile set BARCODE = '' where isnull(BARCODE) "
    db.Execute "Update rtrxfile set BAL_QTY = 0 where isnull(BAL_QTY) "
    db.Execute "Update rtrxfile set BAL_QTY = 0 where BAL_QTY < 0 "
'    db.Execute "Update rtrxfile set M_USER_ID = '' where isnull(M_USER_ID) "
    
    db.Execute "Update itemmast set ISSUE_QTY = 0 where isnull(ISSUE_QTY) "
'    db.Execute "Update itemmast set CATEGORY = '' where isnull(CATEGORY) "
'    db.Execute "Update itemmast set MANUFACTURER = '' where isnull(MANUFACTURER) "
    
    
    GoTo SKIP_BALCHECK
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Dim RSTITEMMAST As ADODB.Recordset
    Dim INWARD As Double
    Dim OUTWARD As Double
    Dim BALQTY As Double
    Dim DIFFQTY As Double
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & LBLITEMCODE.Caption & "' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        BALQTY = 0
        db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY <0"
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(BAL_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY <> 0", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            BALQTY = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
                
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
                
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='MI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockOptimistic, adCmdText
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
        
        If Round(BALQTY, 2) = 0 Then
            db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY >0"
        End If
        
        'If INWARD - OUTWARD <> BALQTY Then MsgBox RSTITEMMAST!ITEM_CODE
        
        If Round(BALQTY, 2) = Round(RSTITEMMAST!CLOSE_QTY, 2) Then GoTo SKIP_BALCHECK
        
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
    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3) = Round(INWARD - OUTWARD, 2)
SKIP_BALCHECK:
    
    
    i = 0
    
    On Error Resume Next
    grdbatch.FixedRows = 4
    grdbatch.rows = 1
    lblpvalue.Caption = ""
    'On Error GoTo Errhand
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & LBLITEMCODE.Caption & "' ORDER BY BAL_QTY DESC, VCH_DATE, TRX_TYPE, VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        i = i + 1
        grdbatch.rows = grdbatch.rows + 1
        grdbatch.FixedRows = 1
        grdbatch.TextMatrix(i, 0) = i
        grdbatch.TextMatrix(i, 1) = rststock!BAL_QTY
        grdbatch.TextMatrix(i, 2) = IIf(IsNull(rststock!P_RETAIL), "", Format(rststock!P_RETAIL, "0.000"))
        grdbatch.TextMatrix(i, 3) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.000"))
        grdbatch.TextMatrix(i, 4) = IIf(IsNull(rststock!EXP_DATE), "", rststock!EXP_DATE)
        grdbatch.TextMatrix(i, 5) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
        'grdbatch.TextMatrix(i, 4) = IIf(IsNull(rststock!CUST_DISC), "", rststock!CUST_DISC)
        If IsNull(rststock!LOOSE_PACK) Then
            grdbatch.TextMatrix(i, 6) = 1
        Else
            grdbatch.TextMatrix(i, 6) = rststock!LOOSE_PACK
        End If
        grdbatch.TextMatrix(i, 7) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
        grdbatch.TextMatrix(i, 8) = IIf(IsNull(rststock!TRX_TYPE), "", rststock!TRX_TYPE)
        grdbatch.TextMatrix(i, 9) = IIf(IsNull(rststock!VCH_NO), "", rststock!VCH_NO)
        grdbatch.TextMatrix(i, 10) = IIf(IsNull(rststock!LINE_NO), "", rststock!LINE_NO)
        grdbatch.TextMatrix(i, 11) = IIf(IsNull(rststock!P_WS), "", Format(rststock!P_WS, "0.000"))

        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    M_EDIT = False
    Screen.MousePointer = vbNormal
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Sub TxtName_Change()
    Call tXTMEDICINE_Change
    Exit Sub
    On Error GoTo ERRHAND
    If Trim(tXTMEDICINE.Text) <> "" Or Trim(TxtCode.Text) <> "" Then Call Fillgrid
    'Call Fillgrid
    If REPFLAG = True Then
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_NAME Like '%" & Me.TxtName.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
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

Private Sub TxtName_GotFocus()
    TxtName.SelStart = 0
    TxtName.SelLength = Len(TxtName.Text)
    'Call Fillgrid
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'TxtItemcode.SetFocus
            Call CmdLoad_Click
        Case vbKeyEscape
            TxtCode.SetFocus
    End Select
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub CmdPrint_Click()
    Dim i As Long
    
    On Error GoTo ERRHAND
    If lblstktype.Caption <> "M" Then Exit Sub
    ReportNameVar = Rptpath & "RPTSTOCKREP"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If chkunbill.Value = 1 Then
        If chkdeaditems.Value = 0 Then
            If Optall.Value = True Then
              Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES')"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        Else
            If Optall.Value = True Then
              Report.RecordSelectionFormula = "({ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES')"
            Else
                Report.RecordSelectionFormula = "({ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        End If
    Else
        If chkdeaditems.Value = 0 Then
            If Optall.Value = True Then
              Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES')"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        Else
            If Optall.Value = True Then
              Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES')"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        End If
    End If
    Set CRXFormulaFields = Report.FormulaFields
    
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        'If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'A R STEELS' & chr(13) & 'Alappuzha'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'Stock Report'"
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "STOCK REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub
