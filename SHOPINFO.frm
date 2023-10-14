VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FRMSHOINFO 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHOP INFORMATION"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12690
   FillColor       =   &H00C0C0FF&
   Icon            =   "SHOPINFO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   12690
   Begin VB.Frame FRAME 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9150
      Left            =   0
      TabIndex        =   21
      Top             =   -120
      Width           =   12645
      Begin VB.CheckBox ChkStpThermal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Hold on thermal Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3300
         TabIndex        =   192
         Top             =   7650
         Width           =   2370
      End
      Begin VB.TextBox TxtTrCopies 
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
         Left            =   6120
         MaxLength       =   1
         TabIndex        =   186
         Top             =   4590
         Width           =   675
      End
      Begin VB.TextBox TxtTrSuf 
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
         Left            =   4275
         MaxLength       =   15
         TabIndex        =   185
         Top             =   4590
         Width           =   1185
      End
      Begin VB.TextBox TxtTrPrefix 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   184
         Top             =   4590
         Width           =   1305
      End
      Begin VB.TextBox TxtPinCode 
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   180
         Top             =   3330
         Width           =   2265
      End
      Begin VB.CheckBox CHKITEMREPEAT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Item Repeat Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   90
         TabIndex        =   160
         Top             =   8820
         Width           =   2355
      End
      Begin VB.CheckBox ChkThermalcopies 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Thermal 2 Copies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3705
         TabIndex        =   146
         Top             =   6765
         Width           =   1875
      End
      Begin VB.CheckBox ChkDMPMini 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "DMP Mini"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6345
         TabIndex        =   145
         Top             =   5295
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C4D6E6&
         Caption         =   "Printer Selection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1845
         Left            =   8295
         TabIndex        =   133
         Top             =   135
         Width           =   4350
         Begin VB.ComboBox CmbDPrint 
            Height          =   315
            ItemData        =   "SHOPINFO.frx":030A
            Left            =   1365
            List            =   "SHOPINFO.frx":0317
            Style           =   2  'Dropdown List
            TabIndex        =   143
            Top             =   1500
            Width           =   2250
         End
         Begin VB.ComboBox Cmbbarcode 
            Height          =   315
            ItemData        =   "SHOPINFO.frx":033A
            Left            =   870
            List            =   "SHOPINFO.frx":033C
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   1170
            Width           =   3420
         End
         Begin VB.ComboBox CmbBillprinter 
            Height          =   315
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Top             =   180
            Width           =   3420
         End
         Begin VB.ComboBox Cmbthermalprinter 
            Height          =   315
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Top             =   855
            Width           =   3420
         End
         Begin VB.ComboBox CmbBillprinterA5 
            Height          =   315
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   134
            Top             =   510
            Width           =   3420
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Print"
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
            Index           =   41
            Left            =   60
            TabIndex        =   142
            Top             =   1530
            Width           =   1350
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
            Height          =   255
            Index           =   40
            Left            =   30
            TabIndex        =   141
            Top             =   1200
            Width           =   885
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bill (A4)"
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
            Index           =   37
            Left            =   30
            TabIndex        =   139
            Top             =   210
            Width           =   765
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Thermal"
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
            Index           =   38
            Left            =   30
            TabIndex        =   138
            Top             =   885
            Width           =   930
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bill (A5)"
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
            Index           =   39
            Left            =   30
            TabIndex        =   137
            Top             =   540
            Width           =   810
         End
      End
      Begin VB.Frame FRMEINVISIBLE 
         BackColor       =   &H00C0E0FF&
         Height          =   3510
         Left            =   8790
         TabIndex        =   121
         Top             =   1890
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton CmdCC 
            BackColor       =   &H00400000&
            Caption         =   "Cost Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   45
            MaskColor       =   &H80000007&
            TabIndex        =   197
            Top             =   3135
            UseMaskColor    =   -1  'True
            Width           =   825
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            Height          =   555
            Left            =   45
            TabIndex        =   193
            Top             =   1980
            Width           =   1980
            Begin VB.OptionButton OptScheme2 
               Caption         =   "Points on Total"
               Height          =   375
               Left            =   1005
               TabIndex        =   195
               Top             =   150
               Width           =   945
            End
            Begin VB.OptionButton OptScheme1 
               Caption         =   "Points on Items"
               Height          =   375
               Left            =   30
               TabIndex        =   194
               Top             =   150
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.TextBox TxtPCode 
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
            IMEMode         =   3  'DISABLE
            Left            =   945
            MaxLength       =   5
            PasswordChar    =   "#"
            TabIndex        =   190
            Top             =   3165
            Width           =   870
         End
         Begin VB.TextBox txtbillformat 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2835
            MaxLength       =   4
            TabIndex        =   177
            Top             =   2025
            Width           =   960
         End
         Begin VB.TextBox TxtCompCode 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2835
            PasswordChar    =   "*"
            TabIndex        =   172
            Top             =   2340
            Width           =   960
         End
         Begin VB.TextBox txtLabel 
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
            Height          =   285
            Left            =   930
            MaxLength       =   1
            TabIndex        =   170
            Top             =   2535
            Width           =   450
         End
         Begin VB.CheckBox ChkStkadjst 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Hide Stock Adjst"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1935
            TabIndex        =   169
            Top             =   1800
            Width           =   1725
         End
         Begin VB.CheckBox Chktemplate 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Barcode template"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   168
            Top             =   1800
            Width           =   1845
         End
         Begin VB.CheckBox ChkPercPurchase 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "% Profit on Purchase"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1935
            TabIndex        =   167
            Top             =   1590
            Width           =   1860
         End
         Begin VB.CheckBox ChkMobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Mobile No. Warning"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   166
            Top             =   1380
            Width           =   1845
         End
         Begin VB.CheckBox ChkPriceSplit 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Disable Price Splitup"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1935
            TabIndex        =   164
            Top             =   1380
            Width           =   1860
         End
         Begin VB.CheckBox ChkCatPur 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Category in Purchase"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1935
            TabIndex        =   163
            Top             =   1170
            Width           =   1860
         End
         Begin VB.CheckBox ChkRstBill 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Reset Bill Nos"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1935
            TabIndex        =   162
            Top             =   960
            Width           =   1860
         End
         Begin VB.CheckBox ChkSalary 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Salary Process"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1935
            TabIndex        =   161
            Top             =   750
            Width           =   1860
         End
         Begin VB.TextBox txtCalc 
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
            IMEMode         =   3  'DISABLE
            Left            =   945
            MaxLength       =   5
            PasswordChar    =   "*"
            TabIndex        =   158
            Top             =   2850
            Width           =   870
         End
         Begin VB.CheckBox ChkBatch 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Batchwise List"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1935
            TabIndex        =   157
            Top             =   540
            Width           =   1695
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
            Left            =   2835
            MaxLength       =   2
            TabIndex        =   154
            Top             =   2685
            Width           =   615
         End
         Begin VB.CheckBox ChkRound 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Round off"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1935
            TabIndex        =   153
            Top             =   330
            Width           =   1695
         End
         Begin VB.CheckBox ChkPrnPetty 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Do not print all in Petty"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1935
            TabIndex        =   148
            Top             =   120
            Width           =   1890
         End
         Begin VB.CommandButton CmdWallPaper 
            BackColor       =   &H00400000&
            Caption         =   "Change Login Screen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1875
            MaskColor       =   &H80000007&
            TabIndex        =   147
            Top             =   3015
            UseMaskColor    =   -1  'True
            Width           =   945
         End
         Begin VB.CheckBox ChkMRPDisc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "MRP Disc in Purchase"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   144
            Top             =   1590
            Width           =   2415
         End
         Begin VB.CheckBox chkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Export Enabled"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   80
            TabIndex        =   128
            Top             =   1170
            Width           =   1845
         End
         Begin VB.CheckBox ChkRemoveUBILL 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Remove UB Items"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   80
            TabIndex        =   127
            Top             =   960
            Width           =   2835
         End
         Begin VB.CheckBox chkub 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "UB Edit"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   80
            TabIndex        =   126
            Top             =   540
            Width           =   1590
         End
         Begin VB.CheckBox CHKDUP 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Do not save masters"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   80
            TabIndex        =   125
            Top             =   750
            Width           =   2235
         End
         Begin VB.CommandButton CmdClear 
            BackColor       =   &H00400000&
            Caption         =   "Clear All Data"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   2865
            MaskColor       =   &H80000007&
            TabIndex        =   122
            Top             =   3015
            UseMaskColor    =   -1  'True
            Width           =   930
         End
         Begin VB.CheckBox CHKPRNALL 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "PRN ALL"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   80
            TabIndex        =   123
            Top             =   330
            Width           =   2235
         End
         Begin VB.CheckBox ChKNSPT 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "NS for PT"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   80
            TabIndex        =   124
            Top             =   120
            Width           =   2235
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bill Format"
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   46
            Left            =   2025
            TabIndex        =   178
            Top             =   2055
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Labels"
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   29
            Left            =   105
            TabIndex        =   171
            Top             =   2550
            Width           =   1080
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Calc Code"
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
            Height          =   300
            Index           =   45
            Left            =   60
            TabIndex        =   159
            Top             =   2865
            Width           =   1080
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
            Left            =   3495
            TabIndex        =   156
            Top             =   2715
            Width           =   390
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Crct"
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
            Height          =   300
            Index           =   43
            Left            =   1905
            TabIndex        =   155
            Top             =   2730
            Width           =   1080
         End
      End
      Begin VB.TextBox TXTSTATENAME 
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
         Left            =   6180
         MaxLength       =   2
         TabIndex        =   131
         Top             =   2385
         Width           =   630
      End
      Begin VB.TextBox Txtstate 
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
         Left            =   5595
         MaxLength       =   2
         TabIndex        =   130
         Top             =   2385
         Width           =   570
      End
      Begin VB.CheckBox CHKAMC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "AMC reminder Enabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3300
         TabIndex        =   120
         Top             =   7425
         Width           =   2355
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00B6CEE9&
         Caption         =   "Shop Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   7035
         TabIndex        =   117
         Top             =   2715
         Visible         =   0   'False
         Width           =   1695
         Begin VB.OptionButton OptRetail 
            BackColor       =   &H00B6CEE9&
            Caption         =   "Retail Shop"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   119
            Top             =   270
            Width           =   1395
         End
         Begin VB.OptionButton OptWs 
            BackColor       =   &H00B6CEE9&
            Caption         =   "Wholesale"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   118
            Top             =   540
            Value           =   -1  'True
            Width           =   1530
         End
      End
      Begin VB.CheckBox chkspace 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line Space"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5145
         TabIndex        =   115
         Top             =   5295
         Width           =   1275
      End
      Begin VB.TextBox TxtHSNSum 
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
         Left            =   1980
         MaxLength       =   6
         TabIndex        =   106
         Top             =   7470
         Width           =   1275
      End
      Begin VB.TextBox TxtMRPDisc 
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
         Left            =   7215
         MaxLength       =   2
         TabIndex        =   101
         Top             =   8790
         Width           =   840
      End
      Begin VB.TextBox TxtVPDisc 
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
         Left            =   11250
         MaxLength       =   2
         TabIndex        =   88
         Top             =   8790
         Width           =   840
      End
      Begin VB.TextBox TxtWSDisc 
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
         Left            =   9240
         MaxLength       =   2
         TabIndex        =   86
         Top             =   8790
         Width           =   840
      End
      Begin VB.TextBox txtRTDisc 
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
         Left            =   5145
         MaxLength       =   2
         TabIndex        =   84
         Top             =   8790
         Width           =   840
      End
      Begin VB.CheckBox chkgst 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "GST No. Warning Enabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2130
         TabIndex        =   83
         Top             =   7200
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00D1E2E7&
         Caption         =   "GST Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   5685
         TabIndex        =   74
         Top             =   6795
         Width           =   1695
         Begin VB.OptionButton OptNonGST 
            BackColor       =   &H00D1E2E7&
            Caption         =   "No GST "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   113
            Top             =   795
            Width           =   1395
         End
         Begin VB.OptionButton OptCompound 
            BackColor       =   &H00D1E2E7&
            Caption         =   "Compounding"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   76
            Top             =   540
            Width           =   1530
         End
         Begin VB.OptionButton OptRegular 
            BackColor       =   &H00D1E2E7&
            Caption         =   "Regular"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   75
            Top             =   285
            Value           =   -1  'True
            Width           =   1395
         End
      End
      Begin VB.CheckBox chkoutPTY 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Print Outstanding"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7440
         TabIndex        =   82
         Top             =   5790
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.CheckBox chkoutB2B 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Print Outstanding"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6840
         TabIndex        =   81
         Top             =   3975
         Value           =   1  'Checked
         Width           =   1965
      End
      Begin VB.CheckBox chkoutB2C 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Print Outstanding"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6840
         TabIndex        =   80
         Top             =   4320
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox chkoutSR 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Print Outstanding"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6840
         TabIndex        =   79
         Top             =   3660
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox ChkDMPThermal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "DMP (Thermal)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2295
         TabIndex        =   78
         Top             =   5295
         Width           =   1545
      End
      Begin VB.CheckBox ChkThPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Thermal Preview"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1635
         TabIndex        =   77
         Top             =   6765
         Width           =   1935
      End
      Begin VB.CheckBox Chk62 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Form 6(2) Pruchase"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2490
         TabIndex        =   73
         Top             =   6990
         Width           =   2445
      End
      Begin VB.CheckBox ChkTaxWarn 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Tax Warning Enabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   90
         TabIndex        =   72
         Top             =   7215
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox ChkBarcode 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Enable Barcode Printing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   90
         TabIndex        =   71
         Top             =   6990
         Value           =   1  'Checked
         Width           =   2445
      End
      Begin VB.CheckBox StDiscount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Don't focus on Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7440
         TabIndex        =   70
         Top             =   5580
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.TextBox TxtMLNO 
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
         Left            =   4545
         MaxLength       =   40
         TabIndex        =   62
         Top             =   2700
         Width           =   2280
      End
      Begin VB.CheckBox Chkwithtax 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Price including GST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7440
         TabIndex        =   61
         Top             =   5370
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00CDD2DE&
         Height          =   735
         Left            =   90
         TabIndex        =   54
         Top             =   9975
         Width           =   3240
         Begin VB.TextBox TxtCGST 
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
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   600
            MaxLength       =   2
            TabIndex        =   56
            Top             =   225
            Width           =   735
         End
         Begin VB.TextBox TxtSGST 
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
            ForeColor       =   &H00004000&
            Height          =   300
            Left            =   2190
            MaxLength       =   2
            TabIndex        =   55
            Top             =   225
            Width           =   675
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
            ForeColor       =   &H00004000&
            Height          =   300
            Index           =   24
            Left            =   2895
            TabIndex        =   60
            Top             =   240
            Width           =   270
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
            ForeColor       =   &H00000040&
            Height          =   300
            Index           =   23
            Left            =   1365
            TabIndex        =   59
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CGST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   300
            Index           =   22
            Left            =   90
            TabIndex        =   58
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "SGST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   300
            Index           =   21
            Left            =   1680
            TabIndex        =   57
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.TextBox TXTIGST 
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
         ForeColor       =   &H00404000&
         Height          =   300
         Left            =   3885
         MaxLength       =   2
         TabIndex        =   51
         Top             =   9960
         Width           =   675
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
         Height          =   300
         Left            =   6075
         MaxLength       =   25
         TabIndex        =   49
         Top             =   2070
         Width           =   1860
      End
      Begin VB.CheckBox ChkPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Preview Bill"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   90
         TabIndex        =   48
         Top             =   6750
         Width           =   1560
      End
      Begin VB.TextBox TXT8VPRE 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   44
         Top             =   4275
         Width           =   1305
      End
      Begin VB.TextBox TXT8VSUF 
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
         Left            =   4275
         MaxLength       =   15
         TabIndex        =   43
         Top             =   4275
         Width           =   1185
      End
      Begin VB.TextBox TxtCopy8V 
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
         Left            =   6120
         MaxLength       =   1
         TabIndex        =   42
         Top             =   4275
         Width           =   675
      End
      Begin VB.TextBox TxtCopy8 
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
         Left            =   6120
         MaxLength       =   1
         TabIndex        =   41
         Top             =   3960
         Width           =   675
      End
      Begin VB.TextBox TxtCopy8B 
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
         Left            =   6120
         MaxLength       =   1
         TabIndex        =   39
         Top             =   3645
         Width           =   675
      End
      Begin VB.CheckBox ChKDMP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "DMP Printer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3855
         TabIndex        =   37
         Top             =   5295
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Terms && Conditions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   5280
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.Frame frmTerms 
         BackColor       =   &H00C0C0FF&
         Height          =   1350
         Left            =   75
         TabIndex        =   36
         Top             =   5445
         Visible         =   0   'False
         Width           =   7305
         Begin VB.TextBox Terms2 
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
            Left            =   30
            MaxLength       =   255
            TabIndex        =   16
            Top             =   420
            Width           =   7230
         End
         Begin VB.TextBox Terms1 
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
            Left            =   30
            MaxLength       =   255
            TabIndex        =   15
            Top             =   120
            Width           =   7230
         End
         Begin VB.TextBox Terms3 
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
            Left            =   30
            MaxLength       =   255
            TabIndex        =   17
            Top             =   720
            Width           =   7230
         End
         Begin VB.TextBox Terms4 
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
            Left            =   30
            MaxLength       =   255
            TabIndex        =   18
            Top             =   1020
            Width           =   7230
         End
      End
      Begin VB.TextBox TXT8SUF 
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
         Left            =   4275
         MaxLength       =   15
         TabIndex        =   12
         Top             =   3960
         Width           =   1185
      End
      Begin VB.TextBox TXT8PRE 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   11
         Top             =   3960
         Width           =   1305
      End
      Begin VB.TextBox TXT8BSUF 
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
         Left            =   4275
         MaxLength       =   15
         TabIndex        =   10
         Top             =   3645
         Width           =   1185
      End
      Begin VB.TextBox TXT8BPRE 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   9
         Top             =   3645
         Width           =   1305
      End
      Begin VB.TextBox txtcst 
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
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   6
         Top             =   2385
         Width           =   2880
      End
      Begin VB.TextBox txtkgst 
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   5
         Top             =   2070
         Width           =   3375
      End
      Begin VB.TextBox txtremarks 
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
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   8
         Top             =   3015
         Width           =   5280
      End
      Begin VB.TextBox txtdlno 
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
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   7
         Top             =   2700
         Width           =   2205
      End
      Begin VB.TextBox txtemail 
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
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   4
         Top             =   1755
         Width           =   3690
      End
      Begin VB.TextBox txtfaxno 
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
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1440
         Width           =   3690
      End
      Begin VB.TextBox txttelno 
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
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1125
         Width           =   3690
      End
      Begin VB.TextBox txtsupplier 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   0
         Top             =   165
         Width           =   6375
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
         Height          =   495
         Left            =   9900
         MaskColor       =   &H80000007&
         TabIndex        =   19
         Top             =   5430
         UseMaskColor    =   -1  'True
         Width           =   1305
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
         Height          =   495
         Left            =   11280
         MaskColor       =   &H80000007&
         TabIndex        =   20
         Top             =   5430
         UseMaskColor    =   -1  'True
         Width           =   1305
      End
      Begin VB.Frame FrmKFC 
         BackColor       =   &H00C0C0FF&
         Height          =   645
         Left            =   7410
         TabIndex        =   107
         Top             =   5910
         Width           =   4590
         Begin VB.CheckBox chkkfc 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "Enable 1% KFC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   60
            TabIndex        =   109
            Top             =   150
            Width           =   1545
         End
         Begin VB.CheckBox ChkKFCNet 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "Price Including Cess"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   60
            TabIndex        =   108
            Top             =   405
            Width           =   2355
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   300
            Left            =   1605
            TabIndex        =   110
            Top             =   135
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   112001025
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTo 
            Height          =   300
            Left            =   3225
            TabIndex        =   111
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   112001025
            CurrentDate     =   40498
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   34
            Left            =   2970
            TabIndex        =   112
            Top             =   150
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00D5DDDF&
         Height          =   1050
         Left            =   90
         TabIndex        =   64
         Top             =   7785
         Width           =   7305
         Begin VB.TextBox TxtPAN 
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
            Left            =   1365
            MaxLength       =   50
            TabIndex        =   67
            Top             =   720
            Width           =   5895
         End
         Begin VB.TextBox txtinvterms 
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
            Left            =   735
            MaxLength       =   200
            TabIndex        =   66
            Top             =   120
            Width           =   6525
         End
         Begin VB.TextBox Txtbank 
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
            Left            =   1365
            MaxLength       =   200
            TabIndex        =   65
            Top             =   420
            Width           =   5895
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Details"
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
            Index           =   28
            Left            =   90
            TabIndex        =   69
            Top             =   405
            Width           =   1290
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Terms"
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
            Index           =   27
            Left            =   90
            TabIndex        =   68
            Top             =   135
            Width           =   1290
         End
      End
      Begin VB.Frame FrmeCode 
         BackColor       =   &H00E3F4EA&
         Height          =   960
         Left            =   5280
         TabIndex        =   149
         Top             =   1110
         Visible         =   0   'False
         Width           =   3000
         Begin VB.CheckBox ChkMulti 
            Appearance      =   0  'Flat
            BackColor       =   &H00E3F4EA&
            Caption         =   "Multi"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   2280
            TabIndex        =   176
            Top             =   195
            Width           =   615
         End
         Begin VB.OptionButton OptDesk 
            BackColor       =   &H00E3F4EA&
            Caption         =   "Desk"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2130
            TabIndex        =   175
            Top             =   720
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton OptApp 
            BackColor       =   &H00E3F4EA&
            Caption         =   "App"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2130
            TabIndex        =   174
            Top             =   510
            Width           =   705
         End
         Begin VB.TextBox txtServer 
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
            Height          =   285
            Left            =   30
            MaxLength       =   30
            TabIndex        =   173
            Top             =   645
            Width           =   1875
         End
         Begin VB.CheckBox ChkCloud 
            Appearance      =   0  'Flat
            BackColor       =   &H00E3F4EA&
            Caption         =   "Enable Remote Access"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   30
            TabIndex        =   152
            Top             =   420
            Width           =   1965
         End
         Begin VB.TextBox TxtCode 
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
            Left            =   1275
            MaxLength       =   6
            TabIndex        =   150
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer ID"
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
            Index           =   42
            Left            =   45
            TabIndex        =   151
            Top             =   135
            Width           =   1245
         End
      End
      Begin VB.Frame frmsales 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sales Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2280
         Left            =   7410
         TabIndex        =   90
         Top             =   6510
         Width           =   4695
         Begin VB.CheckBox ChkZeroWarn 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Zero Stock Warning"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2325
            TabIndex        =   191
            Top             =   1710
            Value           =   1  'Checked
            Width           =   2205
         End
         Begin VB.CheckBox chkpcflag 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Create Items as Price Changing Items"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   45
            TabIndex        =   179
            Top             =   1980
            Width           =   2250
         End
         Begin VB.CheckBox ChkLimitd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Limited Features"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2325
            TabIndex        =   165
            Top             =   1920
            Width           =   1710
         End
         Begin VB.CheckBox chkminusbill 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Disable Minus Billing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   45
            TabIndex        =   132
            Top             =   1770
            Width           =   2955
         End
         Begin VB.CheckBox chkrepeat 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Search any string && repeat"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2325
            TabIndex        =   116
            Top             =   1500
            Value           =   1  'Checked
            Width           =   2205
         End
         Begin VB.CheckBox chkbilltype 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Bill Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   45
            TabIndex        =   114
            Top             =   1350
            Value           =   1  'Checked
            Width           =   2250
         End
         Begin VB.CheckBox StBarcode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Barcode"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2325
            TabIndex        =   104
            Top             =   1275
            Value           =   1  'Checked
            Width           =   2145
         End
         Begin VB.CheckBox chkmrpplus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "MRP + Tax"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   45
            TabIndex        =   103
            Top             =   1560
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkdeliver 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Deliver Items"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   2325
            TabIndex        =   100
            Top             =   1035
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkhideterms 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Terms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   2325
            TabIndex        =   99
            Top             =   825
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkdisc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Disc Amt"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   2325
            TabIndex        =   98
            Top             =   600
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkfree 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Free"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   2325
            TabIndex        =   97
            Top             =   375
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkexpiry 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Expiry"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   2325
            TabIndex        =   96
            Top             =   150
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkmrp 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide MRP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   45
            TabIndex        =   95
            Top             =   1110
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkserial 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Serial No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   45
            TabIndex        =   94
            Top             =   885
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkwrnty 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Warranty"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   45
            TabIndex        =   93
            Top             =   660
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkprnname 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Print Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   45
            TabIndex        =   92
            Top             =   435
            Value           =   1  'Checked
            Width           =   2265
         End
         Begin VB.CheckBox chkspec 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hide Specifications"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   45
            TabIndex        =   91
            Top             =   210
            Value           =   1  'Checked
            Width           =   2265
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copies"
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
         Index           =   50
         Left            =   5475
         TabIndex        =   189
         Top             =   4590
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Tr Suffix"
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
         Index           =   49
         Left            =   2880
         TabIndex        =   188
         Top             =   4590
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Tr Prefix"
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
         Index           =   48
         Left            =   120
         TabIndex        =   187
         Top             =   4590
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
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
         Index           =   47
         Left            =   135
         TabIndex        =   181
         Top             =   3330
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "State Code:"
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
         Index           =   36
         Left            =   4440
         TabIndex        =   129
         Top             =   2385
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "HSN Summary above"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   35
         Left            =   135
         TabIndex        =   105
         Top             =   7485
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MRP Disc%"
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
         Index           =   33
         Left            =   6090
         TabIndex        =   102
         Top             =   8805
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "VP Disc%"
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
         Index           =   32
         Left            =   10140
         TabIndex        =   89
         Top             =   8805
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "WS Disc%"
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
         Index           =   31
         Left            =   8115
         TabIndex        =   87
         Top             =   8805
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RT Disc%"
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
         Index           =   30
         Left            =   4155
         TabIndex        =   85
         Top             =   8805
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ML No"
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
         Index           =   26
         Left            =   3825
         TabIndex        =   63
         Top             =   2700
         Width           =   675
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
         ForeColor       =   &H00404000&
         Height          =   300
         Index           =   25
         Left            =   4695
         TabIndex        =   53
         Top             =   10005
         Width           =   270
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IGST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   300
         Index           =   20
         Left            =   3420
         TabIndex        =   52
         Top             =   9990
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle No."
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
         Index           =   19
         Left            =   4965
         TabIndex        =   50
         Top             =   2145
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GST B2C Prefix"
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
         Index           =   18
         Left            =   120
         TabIndex        =   47
         Top             =   4275
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GST B2C Suffix"
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
         Left            =   2880
         TabIndex        =   46
         Top             =   4275
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copies"
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
         Left            =   5475
         TabIndex        =   45
         Top             =   4275
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copies"
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
         Left            =   5475
         TabIndex        =   40
         Top             =   3960
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copies"
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
         Left            =   5475
         TabIndex        =   38
         Top             =   3645
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GST B2B Sufix"
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
         Left            =   2880
         TabIndex        =   35
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GST B2B Prefix"
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
         Left            =   120
         TabIndex        =   34
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serv Bill Suffix"
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
         Left            =   2880
         TabIndex        =   33
         Top             =   3645
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serv Bill Prefix"
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
         Left            =   135
         TabIndex        =   32
         Top             =   3645
         Width           =   1800
      End
      Begin MSForms.TextBox txtmessage 
         Height          =   345
         Left            =   1560
         TabIndex        =   13
         Top             =   4920
         Width           =   5820
         VariousPropertyBits=   -1400879077
         ForeColor       =   255
         MaxLength       =   100
         BorderStyle     =   1
         Size            =   "10266;609"
         BorderColor     =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Msg on Invoice"
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
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   31
         Top             =   4950
         Width           =   1620
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   135
         TabIndex        =   30
         Top             =   2385
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Website"
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
         Left            =   135
         TabIndex        =   29
         Top             =   2070
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
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
         Left            =   135
         TabIndex        =   28
         Top             =   3015
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FSSAI No."
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
         Left            =   135
         TabIndex        =   27
         Top             =   2700
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "E- mail"
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
         Left            =   135
         TabIndex        =   26
         Top             =   1755
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mob No."
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
         Left            =   135
         TabIndex        =   25
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone No."
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
         Left            =   135
         TabIndex        =   24
         Top             =   1125
         Width           =   1500
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   9
         Left            =   150
         TabIndex        =   23
         Top             =   525
         Width           =   1290
      End
      Begin MSForms.TextBox txtaddress 
         Height          =   630
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   6375
         VariousPropertyBits=   -1400879077
         ForeColor       =   255
         MaxLength       =   99
         BorderStyle     =   1
         Size            =   "11245;1111"
         BorderColor     =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SHOP NAME"
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
         Left            =   135
         TabIndex        =   22
         Top             =   195
         Width           =   1365
      End
   End
   Begin VB.Frame FrmAuth 
      BackColor       =   &H00C0E0FF&
      Height          =   645
      Left            =   15
      TabIndex        =   182
      Top             =   8925
      Visible         =   0   'False
      Width           =   12630
      Begin VB.CheckBox ChkOnlineBill 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Online Bill"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   11370
         TabIndex        =   198
         Top             =   300
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox DTEInvoice 
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
         Height          =   435
         Left            =   9855
         MaxLength       =   15
         TabIndex        =   196
         Top             =   150
         Width           =   1500
      End
      Begin VB.TextBox TxtAuthKey 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   45
         TabIndex        =   183
         Top             =   150
         Width           =   9780
      End
   End
End
Attribute VB_Name = "FRMSHOINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPFLAG As Boolean
Dim RSTREP As New ADODB.Recordset
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)


Private Sub ChKDMP_Click()
    If ChKDMP.Value = 1 Or ChkDMPThermal.Value = 1 Then
        chkspace.Visible = True
    Else
        chkspace.Visible = False
    End If
'    If ChKDMP.value = 1 Then
'       ChkDMPMini.value = 0
'        ChkDMPMini.Visible = True
'    Else
'        ChkDMPMini.value = 0
'        ChkDMPMini.Visible = False
'    End If
End Sub

Private Sub ChkDMPThermal_Click()
    If ChKDMP.Value = 1 Or ChkDMPThermal.Value = 1 Then
        chkspace.Visible = True
    Else
        chkspace.Visible = False
    End If
End Sub

Private Sub chkkfc_Click()
    If chkkfc.Value = 0 Then
        DTFROM.Visible = False
        DTTo.Visible = False
        Label1(34).Visible = False
    Else
        DTFROM.Visible = True
        DTTo.Visible = True
        Label1(34).Visible = True
    End If
End Sub

Private Sub chkTerms_Click()
    If chkTerms.Value = 0 Then
        frmTerms.Visible = False
    Else
        frmTerms.Visible = True
    End If
End Sub



Private Sub CmdCC_Click()
    FrmCC.Show
    FrmCC.SetFocus
End Sub

Private Sub CmdClear_Click()
    If (MsgBox("Are you sure to delete whole data. Cannot be recovered", vbYesNo + vbDefaultButton2, "WARNING NO.1") = vbNo) Then Exit Sub
    If (MsgBox("Are you sure to delete whole data. Cannot be recovered", vbYesNo + vbDefaultButton2, "WARNING NO.2") = vbNo) Then Exit Sub
    If (MsgBox("Are you sure to delete whole data. Cannot be recovered", vbYesNo + vbDefaultButton2, "WARNING NO.3") = vbNo) Then Exit Sub
            
    On Error Resume Next
    db.Execute "TRUNCATE `trxfile`;"
    db.Execute "TRUNCATE `trxmast`;"
    db.Execute "TRUNCATE `rtrxfile`;"
    db.Execute "TRUNCATE `transmast`;"
    db.Execute "TRUNCATE `cashatrxfile`;"
    db.Execute "TRUNCATE `crdtpymt`;"
    db.Execute "TRUNCATE `trxsub`;"
    
    db.Execute "TRUNCATE `custtrxfile`;"
    db.Execute "TRUNCATE `address_book`;"
    db.Execute "TRUNCATE `arealist`;"
    db.Execute "TRUNCATE `astmast`;"
    db.Execute "TRUNCATE `astrxfile`;"
    db.Execute "TRUNCATE `astrxmast`;"
    db.Execute "TRUNCATE `atrxfile`;"
    db.Execute "TRUNCATE `atrxsub`;"
    db.Execute "TRUNCATE `bankcode`;"
    db.Execute "TRUNCATE `bankletters`;"
    db.Execute "TRUNCATE `bank_trx`;"
    db.Execute "TRUNCATE `billdetails`;"
    db.Execute "TRUNCATE `bonusmast`;"
    db.Execute "TRUNCATE `bookfile`;"
    db.Execute "TRUNCATE `cancinv`;"
    db.Execute "TRUNCATE `chqmast`;"
    db.Execute "TRUNCATE `cont_mast`;"
    db.Execute "TRUNCATE `damaged`;"
    db.Execute "TRUNCATE `DAMAGE_MAST`;"
    db.Execute "TRUNCATE `dbtpymt`;"
    db.Execute "TRUNCATE `de_ret_details`;"
    db.Execute "TRUNCATE `docmast`;"
    db.Execute "TRUNCATE `expiry`;"
    db.Execute "TRUNCATE `explist`;"
    db.Execute "TRUNCATE `expsort`;"
    db.Execute "TRUNCATE `fqtylist`;"
    db.Execute "TRUNCATE `gift`;"
    db.Execute "TRUNCATE `moleculelink`;"
    db.Execute "TRUNCATE `molecules`;"
    db.Execute "TRUNCATE `nonrcvd`;"
    db.Execute "TRUNCATE `ordissue`;"
    db.Execute "TRUNCATE `ordsub`;"
    db.Execute "TRUNCATE `paste errors`;"
    db.Execute "TRUNCATE `pomast`;"
    db.Execute "TRUNCATE `posub`;"
    db.Execute "TRUNCATE `pricetable`;"
    db.Execute "TRUNCATE `prodlink`;"
    db.Execute "TRUNCATE `ptnmast`;"
    db.Execute "TRUNCATE `purcahsereturn`;"
    db.Execute "TRUNCATE `purch_return`;"
    db.Execute "TRUNCATE `qtnmast`;"
    db.Execute "TRUNCATE `qtnsub`;"
    db.Execute "TRUNCATE `reorder`;"
    db.Execute "TRUNCATE `replcn`;"
    db.Execute "TRUNCATE `returnmast`;"
    db.Execute "TRUNCATE `salereturn`;"
    db.Execute "TRUNCATE `salesledger`;"
    db.Execute "TRUNCATE `salesman`;"
    db.Execute "TRUNCATE `salesreg`;"
    db.Execute "TRUNCATE `schedule`;"
    db.Execute "TRUNCATE `seldist`;"
    db.Execute "TRUNCATE `service_stk`;"
    db.Execute "TRUNCATE `slip_reg`;"
    db.Execute "TRUNCATE `srtrxfile`;"
    db.Execute "TRUNCATE `stockreport`;"
    db.Execute "TRUNCATE `tempcn`;"
    db.Execute "TRUNCATE `tempstk`;"
    db.Execute "TRUNCATE `temptrxfile`;"
    db.Execute "TRUNCATE `tmporderlist`;"
    db.Execute "TRUNCATE `transsub`;"
    db.Execute "TRUNCATE `trnxrcpt`;"
    db.Execute "TRUNCATE `trxexpense`;"
    db.Execute "TRUNCATE `trxexpmast`;"
    db.Execute "TRUNCATE `trxexp_mast`;"
    db.Execute "TRUNCATE `trxfileexp`;"
    db.Execute "TRUNCATE `trxfile_exp`;"
    db.Execute "TRUNCATE `trxfile_sp`;"
    db.Execute "TRUNCATE `trxformulamast`;"
    db.Execute "TRUNCATE `trxformulasub`;"
    db.Execute "TRUNCATE `trxincmast`;"
    db.Execute "TRUNCATE `trxincome`;"
    db.Execute "TRUNCATE `trxmast_sp`;"
    db.Execute "TRUNCATE `vanstock`;"
    db.Execute "TRUNCATE `war_list`;"
    db.Execute "TRUNCATE `war_trxfile`;"
    db.Execute "TRUNCATE `war_trxns`;"
    db.Execute "delete from `compinfo` where COMP_CODE <> 1;"
    
    If (MsgBox("Are you sure to delete Item List", vbYesNo + vbDefaultButton2, "DELETE") = vbYes) Then db.Execute "TRUNCATE `itemmast`;"
    If (MsgBox("Are you sure to delete Supplier List", vbYesNo + vbDefaultButton2, "DELETE") = vbYes) Then db.Execute "delete from `actmast` where (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3);"
    If (MsgBox("Are you sure to delete Customer List", vbYesNo + vbDefaultButton2, "DELETE") = vbYes) Then db.Execute "delete from `custmast` where ACT_CODE <> '130000' and ACT_CODE <> '130001' ;"
    If (MsgBox("Are you sure to delete Category List", vbYesNo + vbDefaultButton2, "DELETE") = vbYes) Then db.Execute "TRUNCATE `category`;"
    If (MsgBox("Are you sure to delete Manufacturer List", vbYesNo + vbDefaultButton2, "DELETE") = vbYes) Then db.Execute "TRUNCATE `manufact`;"
    
    MsgBox "Success", , "EzBiz"
    
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTCOMPANY As ADODB.Recordset
    
    If txtsupplier.Text = "" Then
        MsgBox "ENTER NAME OF SHOP", vbOKOnly, "SHOP INFO...."
        txtsupplier.SetFocus
        Exit Sub
    End If
    
'    If Val(TxtCGST.Text) = 0 Then
'        MsgBox "Please enter the CGST amount", vbOKOnly, "SHOP INFO...."
'        TxtCGST.SetFocus
'        Exit Sub
'    End If
'
'    If Val(TxtSGST.Text) = 0 Then
'        MsgBox "Please enter the SGST amount", vbOKOnly, "SHOP INFO...."
'        TxtSGST.SetFocus
'        Exit Sub
'    End If
    TxtPinCode = Trim(TxtPinCode.Text)
    If Trim(Txtstate.Text) = "" Then Txtstate.Text = "32"
    If Val(Txtstate.Text) = 0 Then Txtstate.Text = "32"
    If Len(Txtstate.Text) <> 2 Then Txtstate.Text = "32"
    If Val(Txtstate.Text) = 32 Then TXTSTATENAME.Text = "KL"
    If Len(TxtPinCode.Text) <> 6 Then
        MsgBox "Please enter a valid pincode", vbOKOnly, "SHOP INFO...."
        TxtPinCode.SetFocus
        Exit Sub
    End If
    If chkkfc.Value = 1 And DTFROM.Value = DTTo.Value Then
        MsgBox "Kerala Flood Cess Start date should not be equal to End date", vbOKOnly, "SHOP INFO...."
        DTFROM.SetFocus
        Exit Sub
    End If
    
    If chkkfc.Value = 1 And DTFROM.Value > DTTo.Value Then
        MsgBox "Kerala Flood Cess Start date should not be greater than End date", vbOKOnly, "SHOP INFO...."
        DTFROM.SetFocus
        Exit Sub
    End If

'    If ChkBarcode.value = 1 And Cmbprofile.ListIndex = -1 Then
'        MsgBox "Please select a profile for barcode label printing", vbOKOnly, "SHOP INFO...."
'        Cmbprofile.Visible = True
'        Cmbprofile.SetFocus
'        Exit Sub
'    End If
    If chkTerms.Value = True And (Trim(Terms1.Text) = "" Or Trim(Terms2.Text) = "" Or Trim(Terms3.Text) = "" Or Trim(Terms4.Text) = "") Then
        MsgBox "Please enter the Terms & Conditions"
        Terms1.SetFocus
        Exit Sub
    End If
    If Trim(TxtMLNO.Text) = "*101*" Then TxtMLNO.Text = ""
    If Trim(TxtMLNO.Text) = "*111#" Then TxtMLNO.Text = ""
    If Trim(TxtMLNO.Text) = "*1707#" Then TxtMLNO.Text = ""
    On Error GoTo ERRHAND
    
    Set RSTCOMPANY = New ADODB.Recordset

    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        RSTCOMPANY.AddNew
        RSTCOMPANY!COMP_CODE = "001"
        RSTCOMPANY!FIN_YR = Year(MDIMAIN.DTFROM.Value)
    End If
    RSTCOMPANY!CUST_CODE = Trim(TxtCode.Text)
    RSTCOMPANY!SERVER_ADD = Trim(txtServer.Text)
    If OptApp.Value = True Then
        RSTCOMPANY!CLOUD_TYPE = 1
    Else
        RSTCOMPANY!CLOUD_TYPE = 0
    End If
    If ChkCloud.Value = 1 Then
        RSTCOMPANY!REMOTE_FLAG = "Y"
    Else
        RSTCOMPANY!REMOTE_FLAG = "N"
    End If
    If ChkMulti.Value = 1 Then
        RSTCOMPANY!MULTI_FLAG = "Y"
    Else
        RSTCOMPANY!MULTI_FLAG = "N"
    End If
    RSTCOMPANY!COMP_NAME = Trim(txtsupplier.Text)
    RSTCOMPANY!Address = Trim(txtaddress.Text)
    RSTCOMPANY!TEL_NO = Trim(txttelno.Text)
    RSTCOMPANY!FAX_NO = Trim(txtfaxno.Text)
    RSTCOMPANY!EMAIL_ADD = Trim(txtemail.Text)
    RSTCOMPANY!ST_ADDRESS = Trim(txtkgst.Text)
    RSTCOMPANY!CST = Trim(TxtCST.Text)
    RSTCOMPANY!SCODE = Trim(Txtstate.Text)
    RSTCOMPANY!SNAME = Trim(TXTSTATENAME.Text)
    RSTCOMPANY!DL_NO = Trim(txtdlno.Text)
    RSTCOMPANY!ML_NO = Trim(TxtMLNO.Text)
    RSTCOMPANY!HO_NAME = Trim(TXTREMARKS.Text)
    RSTCOMPANY!PINCODE = Trim(TxtPinCode.Text)
    RSTCOMPANY!auth_key = Trim(TxtAuthKey.Text)
    If Trim(TxtAuthKey.Text) <> "" Then
        RSTCOMPANY!auth_date = DTEInvoice.Text
    Else
        RSTCOMPANY!auth_date = ""
    End If
    RSTCOMPANY!INV_MSGS = Trim(txtmessage.Text)
    RSTCOMPANY!PREFIX_8B = Trim(TXT8BPRE.Text)
    RSTCOMPANY!SUFIX_8B = Trim(TXT8BSUF.Text)
    RSTCOMPANY!PREFIX_8 = Trim(TXT8PRE.Text)
    RSTCOMPANY!SUFIX_8 = Trim(TXT8SUF.Text)
    RSTCOMPANY!PREFIX_8V = Trim(TXT8VPRE.Text)
    RSTCOMPANY!SUFIX_8V = Trim(TXT8VSUF.Text)
    RSTCOMPANY!PREFIX_TR = Trim(TxtTrPrefix.Text)
    RSTCOMPANY!SUFIX_TR = Trim(TxtTrSuf.Text)
    RSTCOMPANY!VEHICLE = Trim(TxtVehicle.Text)
    RSTCOMPANY!CGST = Val(TxtCGST.Text)
    RSTCOMPANY!SGST = Val(TxtSGST.Text)
    RSTCOMPANY!IGST = Val(TXTIGST.Text)
    RSTCOMPANY!RTDISC = Val(txtRTDisc.Text)
    RSTCOMPANY!WSDISC = Val(TxtWSDisc.Text)
    RSTCOMPANY!VPDISC = Val(TxtVPDisc.Text)
    RSTCOMPANY!MRPDISC = Val(TxtMRPDisc.Text)
    RSTCOMPANY!HSN_SUM = Val(TxtHSNSum.Text)
    If OptCompound.Value = True Then
        RSTCOMPANY!GST_FLAG = "C"
    ElseIf OptNonGST.Value = True Then
        RSTCOMPANY!GST_FLAG = "N"
    Else
        RSTCOMPANY!GST_FLAG = "R"
    End If
    RSTCOMPANY!Copy_8B = Val(TxtCopy8B.Text)
    RSTCOMPANY!Copy_8 = Val(TxtCopy8.Text)
    RSTCOMPANY!Copy_8V = Val(TxtCopy8V.Text)
    RSTCOMPANY!Copy_TR = Val(TxtTrCopies.Text)
    RSTCOMPANY!D_PRINT = CmbDPrint.ListIndex
    
    If ChkThermalcopies.Value = 1 Then
        RSTCOMPANY!T2_COPIES = "Y"
    Else
        RSTCOMPANY!T2_COPIES = "N"
    End If
    If ChkStkadjst.Value = 1 Then
        RSTCOMPANY!STK_ADJ = "Y"
        MDIMAIN.MNUOPSTK.Visible = False
    Else
        RSTCOMPANY!STK_ADJ = "N"
        MDIMAIN.MNUOPSTK.Visible = True
    End If
'        If Val(TxtCopy8B.Text) <= 0 Then
'            RSTCOMPANY!Copy_8B = 1
'        Else
'            RSTCOMPANY!Copy_8B = Val(TxtCopy8B.Text)
'        End If
'        If Val(TxtCopy8.Text) <= 0 Then
'            RSTCOMPANY!Copy_8 = 1
'        Else
'            RSTCOMPANY!Copy_8 = Val(TxtCopy8.Text)
'        End If
'        If Val(TxtCopy8V.Text) <= 0 Then
'            RSTCOMPANY!Copy_8V = 1
'        Else
'            RSTCOMPANY!Copy_8V = Val(TxtCopy8V.Text)
'        End If
    If ChKNSPT.Value = 1 Then
        RSTCOMPANY!NSPT = "Y"
    Else
        RSTCOMPANY!NSPT = "N"
    End If
    If CHKPRNALL.Value = 1 Then
        RSTCOMPANY!ALL_PRN = "Y"
    Else
        RSTCOMPANY!ALL_PRN = "N"
    End If
    If chkub.Value = 1 Then
        RSTCOMPANY!UB = "Y"
    Else
        RSTCOMPANY!UB = "N"
    End If
    If ChkOnlineBill.Value = 1 Then
        RSTCOMPANY!ONLINE_BILL = "Y"
    Else
        RSTCOMPANY!ONLINE_BILL = "N"
    End If
    If ChKDMP.Value = 1 Then
        RSTCOMPANY!DMP_FLAG = "Y"
    Else
        RSTCOMPANY!DMP_FLAG = "N"
    End If
    
    If ChkDMPMini.Value = 1 Then
        RSTCOMPANY!DMP_MINI = "Y"
    Else
        RSTCOMPANY!DMP_MINI = "N"
    End If
    
    If chkspace.Value = 1 Then
        RSTCOMPANY!LINE_SPACE = "Y"
    Else
        RSTCOMPANY!LINE_SPACE = "N"
    End If
    If ChkDMPThermal.Value = 1 Then
        RSTCOMPANY!DMPTH_FLAG = "Y"
    Else
        RSTCOMPANY!DMPTH_FLAG = "N"
    End If
    If Chkwithtax.Value = 1 Then
        RSTCOMPANY!TAX_FLAG = "Y"
    Else
        RSTCOMPANY!TAX_FLAG = "N"
    End If
    If ChkKFCNet.Value = 1 Then
        RSTCOMPANY!KFCNET = "Y"
    Else
        RSTCOMPANY!KFCNET = "N"
    End If
    If ChkTaxWarn.Value = 1 Then
        RSTCOMPANY!TAXWRN_FLAG = "Y"
    Else
        RSTCOMPANY!TAXWRN_FLAG = "N"
    End If
    If chkgst.Value = 1 Then
        RSTCOMPANY!GSTWRN_FLAG = "Y"
    Else
        RSTCOMPANY!GSTWRN_FLAG = "N"
    End If
    If CHKITEMREPEAT.Value = 1 Then
        RSTCOMPANY!ITEM_WARN = "Y"
    Else
        RSTCOMPANY!ITEM_WARN = "N"
    End If
    If CHKAMC.Value = 1 Then
        RSTCOMPANY!AMC_FLAG = "Y"
    Else
        RSTCOMPANY!AMC_FLAG = "N"
    End If
    If ChkStpThermal.Value = 1 Then
        RSTCOMPANY!HOLD_THERMAL_FLAG = "Y"
    Else
        RSTCOMPANY!HOLD_THERMAL_FLAG = "N"
    End If
    If Chk62.Value = 1 Then
        RSTCOMPANY!FORM_62 = "Y"
    Else
        RSTCOMPANY!FORM_62 = "N"
    End If
    If CHKDUP.Value = 0 Then
        RSTCOMPANY!DUP_FLAG = "N"
    Else
        RSTCOMPANY!DUP_FLAG = "Y"
    End If
    If ChkRound.Value = 1 Then
        RSTCOMPANY!ROUND_FLAG = "Y"
    Else
        RSTCOMPANY!ROUND_FLAG = "N"
    End If
    If ChkBatch.Value = 1 Then
        RSTCOMPANY!BATCH_FLAG = "Y"
    Else
        RSTCOMPANY!BATCH_FLAG = "N"
    End If
    RSTCOMPANY!STOCK_CRCT = Val(txtstkcrct.Text)
    RSTCOMPANY!CLCODE = Val(txtCalc.Text)
    RSTCOMPANY!PCODE = Trim(TxtPCode.Text)
    DUPCODE = Trim(TxtPCode.Text)
    If Val(txtLabel.Text) <= 0 Then
        RSTCOMPANY!BAR_LABELS = 1
    Else
        RSTCOMPANY!BAR_LABELS = Val(txtLabel.Text)
    End If
    RSTCOMPANY!BILL_FORMAT = Trim(txtbillformat.Text)
    If ChkSalary.Value = 1 Then
        RSTCOMPANY!SAL_PROC = "Y"
    Else
        RSTCOMPANY!SAL_PROC = "N"
    End If
    If ChkCatPur.Value = 1 Then
        RSTCOMPANY!CAT_PURCHASE = "Y"
        MDIMAIN.lblcategory = "Y"
    Else
        RSTCOMPANY!CAT_PURCHASE = "N"
        MDIMAIN.lblcategory = "N"
    End If
    If ChkRstBill.Value = 1 Then
        RSTCOMPANY!RST_BILL = "Y"
        RstBill_Flag = "Y"
    Else
        RSTCOMPANY!RST_BILL = "N"
        RstBill_Flag = "N"
    End If
    If ChkPriceSplit.Value = 1 Then
        RSTCOMPANY!PRICE_SPLIT = "Y"
        MDIMAIN.lblPriceSplit.Caption = "Y"
    Else
        RSTCOMPANY!PRICE_SPLIT = "N"
        MDIMAIN.lblPriceSplit.Caption = "N"
    End If
    If ChkPercPurchase.Value = 1 Then
        RSTCOMPANY!PER_PURCHASE = "Y"
        MDIMAIN.lblPerPurchase.Caption = "Y"
    Else
        RSTCOMPANY!PER_PURCHASE = "N"
        MDIMAIN.lblPerPurchase.Caption = "N"
    End If
    If ChkPrnPetty.Value = 1 Then
        RSTCOMPANY!PRN_PETTY_FLAG = "Y"
    Else
        RSTCOMPANY!PRN_PETTY_FLAG = "N"
    End If
    If ChkPreview.Value = 1 Then
        RSTCOMPANY!PREVIEW_FLAG = "Y"
    Else
        RSTCOMPANY!PREVIEW_FLAG = "N"
    End If
    If ChkThPreview.Value = 1 Then
        RSTCOMPANY!PREVIEWTH_FLAG = "Y"
    Else
        RSTCOMPANY!PREVIEWTH_FLAG = "N"
    End If
    If StBarcode.Value = 1 Then
        RSTCOMPANY!CODE_FLAG = "Y"
    Else
        RSTCOMPANY!CODE_FLAG = "N"
    End If
    If StDiscount.Value = 1 Then
        RSTCOMPANY!DISC_FLAG = "Y"
    Else
        RSTCOMPANY!DISC_FLAG = "N"
    End If
    If ChkBarcode.Value = 1 Then
        RSTCOMPANY!BARCODE_FLAG = "Y"
        'RSTCOMPANY!barcode_profile = Cmbprofile.ListIndex
    Else
        RSTCOMPANY!BARCODE_FLAG = "N"
       ' RSTCOMPANY!barcode_profile = ""
    End If
    RSTCOMPANY!barcode_profile = 0
    If chkoutSR.Value = 1 Then
        RSTCOMPANY!OSSR_FLAG = "Y"
    Else
        RSTCOMPANY!OSSR_FLAG = "N"
    End If
    If chkoutB2C.Value = 1 Then
        RSTCOMPANY!OSB2C_FLAG = "Y"
    Else
        RSTCOMPANY!OSB2C_FLAG = "N"
    End If
    If chkoutB2B.Value = 1 Then
        RSTCOMPANY!OSB2B_FLAG = "Y"
    Else
        RSTCOMPANY!OSB2B_FLAG = "N"
    End If
    If chkoutPTY.Value = 1 Then
        RSTCOMPANY!OSPTY_FLAG = "Y"
    Else
        RSTCOMPANY!OSPTY_FLAG = "N"
    End If
    If ChkRemoveUBILL.Value = 1 Then
        RSTCOMPANY!REMOVE_UBILL = "Y"
    Else
        RSTCOMPANY!REMOVE_UBILL = "N"
    End If
    If ChkMRPDisc.Value = 1 Then
        RSTCOMPANY!MRP_DISC = "Y"
    Else
        RSTCOMPANY!MRP_DISC = "N"
    End If
    If chkexport.Value = 1 Then
        RSTCOMPANY!EXP_ENABLED = "Y"
    Else
        RSTCOMPANY!EXP_ENABLED = "N"
    End If
    If chkTerms.Value = 1 Then
        RSTCOMPANY!TERMS_FLAG = "Y"
        RSTCOMPANY!Terms1 = Trim(Terms1.Text)
        RSTCOMPANY!Terms2 = Trim(Terms2.Text)
        RSTCOMPANY!Terms3 = Trim(Terms3.Text)
        RSTCOMPANY!Terms4 = Trim(Terms4.Text)
    Else
        RSTCOMPANY!TERMS_FLAG = "N"
        RSTCOMPANY!Terms1 = ""
        RSTCOMPANY!Terms2 = ""
        RSTCOMPANY!Terms3 = ""
        RSTCOMPANY!Terms4 = ""
    End If
    RSTCOMPANY!INV_TERMS = Trim(txtinvterms.Text)
    RSTCOMPANY!bank_details = Trim(Txtbank.Text)
    RSTCOMPANY!PAN_NO = Trim(TxtPAN.Text)
    If OptRetail.Value = True Then
        RSTCOMPANY!SHOP_RT = "Y"
        MDIMAIN.LBLSHOPRT.Caption = "Y"
    Else
        RSTCOMPANY!SHOP_RT = "N"
        MDIMAIN.LBLSHOPRT.Caption = "N"
    End If
    If chkspec.Value = 1 Then
        RSTCOMPANY!hide_spec = "Y"
    Else
        RSTCOMPANY!hide_spec = "N"
    End If
    If chkprnname.Value = 1 Then
        RSTCOMPANY!hide_pr_name = "Y"
    Else
        RSTCOMPANY!hide_pr_name = "N"
    End If
    If chkwrnty.Value = 1 Then
        RSTCOMPANY!hide_wrnty = "Y"
    Else
        RSTCOMPANY!hide_wrnty = "N"
    End If
    If chkserial.Value = 1 Then
        RSTCOMPANY!hide_serial = "Y"
    Else
        RSTCOMPANY!hide_serial = "N"
    End If
    If chkmrp.Value = 1 Then
        RSTCOMPANY!hide_mrp = "Y"
    Else
        RSTCOMPANY!hide_mrp = "N"
    End If
    If chkbilltype.Value = 1 Then
        RSTCOMPANY!billtype_flag = "Y"
    Else
        RSTCOMPANY!billtype_flag = "N"
    End If
'        If chkrepeat.value = 1 Then
'            RSTCOMPANY!item_repeat = "Y"
'        Else
'            RSTCOMPANY!item_repeat = "N"
'        End If
    If chkexpiry.Value = 1 Then
        RSTCOMPANY!hide_expiry = "Y"
    Else
        RSTCOMPANY!hide_expiry = "N"
    End If
    If ChkFree.Value = 1 Then
        RSTCOMPANY!hide_free = "Y"
    Else
        RSTCOMPANY!hide_free = "N"
    End If
    If chkdisc.Value = 1 Then
        RSTCOMPANY!hide_disc = "Y"
    Else
        RSTCOMPANY!hide_disc = "N"
    End If
    If chkhideterms.Value = 1 Then
        RSTCOMPANY!hide_terms = "Y"
    Else
        RSTCOMPANY!hide_terms = "N"
    End If
    If chkdeliver.Value = 1 Then
        RSTCOMPANY!hide_deliver = "Y"
    Else
        RSTCOMPANY!hide_deliver = "N"
    End If
    If chkmrpplus.Value = 1 Then
        RSTCOMPANY!mrp_plus = "Y"
    Else
        RSTCOMPANY!mrp_plus = "N"
    End If
    If chkbilltype.Value = 1 Then
        RSTCOMPANY!billtype_flag = "Y"
    Else
        RSTCOMPANY!billtype_flag = "N"
    End If
'        If chkrepeat.value = 1 Then
'            RSTCOMPANY!item_repeat = "Y"
'        Else
'            RSTCOMPANY!item_repeat = "N"
'        End If
    
    If chkkfc.Value = 1 Then
        RSTCOMPANY!kfc_flag = "Y"
        RSTCOMPANY!KFCFROM_DATE = DTFROM.Value
        RSTCOMPANY!KFCTO_DATE = DTTo.Value
    Else
        RSTCOMPANY!kfc_flag = "N"
    End If
    If chkexport.Value = 1 Then
        MDIMAIN.lblRemoveUbill.Caption = "Y"
    Else
        MDIMAIN.lblRemoveUbill.Caption = "N"
    End If
    If chkexport.Value = 1 Then
        MDIMAIN.lblExpEnable.Caption = "Y"
    Else
        MDIMAIN.lblExpEnable.Caption = "N"
    End If
    If ChkZeroWarn.Value = 1 Then
        RSTCOMPANY!Zero_Warn = "Y"
        ZERO_WARN_FLAG = True
    Else
        RSTCOMPANY!Zero_Warn = "N"
        ZERO_WARN_FLAG = False
    End If
    If chkrepeat.Value = 1 Then
        RSTCOMPANY!item_repeat = "Y"
    Else
        RSTCOMPANY!item_repeat = "N"
    End If
    If chkpcflag.Value = 1 Then
        RSTCOMPANY!VS_FLAG = "Y"
    Else
        RSTCOMPANY!VS_FLAG = "N"
    End If
    If chkminusbill.Value = 1 Then
        RSTCOMPANY!MINUS_BILL = "N"
    Else
        RSTCOMPANY!MINUS_BILL = "Y"
    End If
    If ChkLimitd.Value = 1 Then
        RSTCOMPANY!LMT_FLAG = "Y"
    Else
        RSTCOMPANY!LMT_FLAG = "N"
    End If
    If Chktemplate.Value = 1 Then
        RSTCOMPANY!BAR_TEMPLATE = "Y"
    Else
        RSTCOMPANY!BAR_TEMPLATE = "N"
    End If
    If ChkMobile.Value = 1 Then
        RSTCOMPANY!MOB_WARN_FLAG = "Y"
    Else
        RSTCOMPANY!MOB_WARN_FLAG = "N"
    End If
    
    If ChkDMPMini.Value = 1 Then
        MDIMAIN.lbldmpmini.Caption = "Y"
    Else
        MDIMAIN.lbldmpmini.Caption = "N"
    End If
    
    If OptScheme2.Value = True Then
        scheme_option = "1"
        RSTCOMPANY!SCHEME_OPT = "1"
    Else
        scheme_option = "0"
        RSTCOMPANY!SCHEME_OPT = "0"
    End If
    
    RSTCOMPANY.Update
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    MDIMAIN.LBLRT.Caption = Val(txtRTDisc.Text)
    MDIMAIN.LBLWS.Caption = Val(TxtWSDisc.Text)
    MDIMAIN.lblvp.Caption = Val(TxtVPDisc.Text)
    MDIMAIN.LBLMRP.Caption = Val(TxtMRPDisc.Text)
    MDIMAIN.LBLHSNSUM.Caption = Val(TxtHSNSum.Text)
    
    customercode = Trim(TxtCode.Text)
    serveraddress = Trim(txtServer.Text)
    If OptApp.Value = True Then
        REMOTEAPP = 1
    Else
        REMOTEAPP = 0
    End If
    
    If ChkCloud.Value = 1 Then
        remoteflag = True
        MDIMAIN.mnusync.Visible = True
    Else
        remoteflag = False
        MDIMAIN.mnusync.Visible = False
    End If
            
    MDIMAIN.StatusBar.Panels(5).Text = Trim(txtsupplier.Text)
    If chkkfc.Value = 1 Then
        MDIMAIN.lblkfc.Caption = "Y"
        MDIMAIN.DTKFCSTART = DTFROM.Value
        MDIMAIN.DTKFCEND = DTTo.Value
    Else
        MDIMAIN.lblkfc.Caption = "N"
    End If
    If ChKDMP.Value = 1 Then
        MDIMAIN.StatusBar.Panels(8).Text = "Y"
    Else
        MDIMAIN.StatusBar.Panels(8).Text = "N"
    End If
    If chkspace.Value = 1 Then
        MDIMAIN.LBLSPACE.Caption = "Y"
    Else
        MDIMAIN.LBLSPACE.Caption = "N"
    End If
    If ChKNSPT.Value = 1 Then
        MDIMAIN.lblnostock.Caption = "Y"
    Else
        MDIMAIN.lblnostock.Caption = "N"
    End If
    If CHKPRNALL.Value = 1 Then
        MDIMAIN.lblprnall.Caption = "Y"
    Else
        MDIMAIN.lblprnall.Caption = "N"
    End If
    If ChkRemoveUBILL.Value = 1 Then
        MDIMAIN.lblRemoveUbill.Caption = "Y"
    Else
        MDIMAIN.lblRemoveUbill.Caption = "N"
    End If
    If ChkMRPDisc.Value = 1 Then
        MRPDISC_FLAG = "Y"
    Else
        MRPDISC_FLAG = "N"
    End If
    If chkexport.Value = 1 Then
        MDIMAIN.lblExpEnable.Caption = "Y"
    Else
        MDIMAIN.lblExpEnable.Caption = "N"
    End If
    If chkub.Value = 1 Then
        MDIMAIN.lblub.Caption = "Y"
    Else
        MDIMAIN.lblub.Caption = "N"
    End If
    If ChkDMPThermal.Value = 1 Then
        MDIMAIN.LBLDMPTHERMAL.Caption = "Y"
    Else
        MDIMAIN.LBLDMPTHERMAL.Caption = "N"
    End If
    If Chkwithtax.Value = 1 Then
        MDIMAIN.StatusBar.Panels(14).Text = "Y"
    Else
        MDIMAIN.StatusBar.Panels(14).Text = "N"
    End If
    If ChkKFCNet.Value = 1 Then
        MDIMAIN.LblKFCNet.Caption = "Y"
    Else
        MDIMAIN.LblKFCNet.Caption = "N"
    End If
    If ChkTaxWarn.Value = 1 Then
        MDIMAIN.LBLTAXWARN.Caption = "Y"
    Else
        MDIMAIN.LBLTAXWARN.Caption = "N"
    End If
    If chkgst.Value = 1 Then
        MDIMAIN.LBLGSTWRN.Caption = "Y"
    Else
        MDIMAIN.LBLGSTWRN.Caption = "N"
    End If
    If CHKITEMREPEAT.Value = 1 Then
        MDIMAIN.LBLITMWRN.Caption = "Y"
    Else
        MDIMAIN.LBLITMWRN.Caption = "N"
    End If
    If CHKAMC.Value = 1 Then
        MDIMAIN.LBLAMC.Caption = "Y"
    Else
        MDIMAIN.LBLAMC.Caption = "N"
    End If
    If ChkStpThermal.Value = 1 Then
        hold_thermal = True
    Else
        hold_thermal = False
    End If
    If Chk62.Value = 1 Then
        MDIMAIN.lblform62.Caption = "Y"
    Else
        MDIMAIN.lblform62.Caption = "N"
    End If
    If CHKDUP.Value = 1 Then
        MDIMAIN.StatusBar.Panels(9).Text = "Y"
    Else
        MDIMAIN.StatusBar.Panels(9).Text = "N"
    End If
    If ChkRound.Value = 1 Then
        Roundflag = True
    Else
        Roundflag = False
    End If
    If ChkBatch.Value = 1 Then
        BATCH_DISPLAY = True
    Else
        BATCH_DISPLAY = False
    End If
    If ChkPrnPetty.Value = 1 Then
        PRNPETTYFLAG = True
    Else
        PRNPETTYFLAG = False
    End If
    If chkrepeat.Value = 1 Then
        MDIMAIN.lblitemrepeat.Caption = "Y"
    Else
        MDIMAIN.lblitemrepeat.Caption = "N"
    End If
    If chkpcflag.Value = 1 Then
        PC_FLAG = "Y"
    Else
        PC_FLAG = "N"
    End If
    If chkminusbill.Value = 1 Then
        MINUS_BILL = "N"
    Else
        MINUS_BILL = "Y"
    End If
    If ChkLimitd.Value = 1 Then
        SALESLT_FLAG = "Y"
    Else
        SALESLT_FLAG = "N"
    End If
    If Chktemplate.Value = 1 Then
        BARTEMPLATE = "Y"
    Else
        BARTEMPLATE = "N"
    End If
'    If Chkbarformat.Value = 1 Then
'        BARFORMAT = "Y"
'    Else
'        BARFORMAT = "N"
'    End If
    If ChkPreview.Value = 1 Then
        MDIMAIN.StatusBar.Panels(13).Text = "Y"
    Else
        MDIMAIN.StatusBar.Panels(13).Text = "N"
    End If
    If ChkThPreview.Value = 1 Then
        MDIMAIN.LBLTHPREVIEW.Caption = "Y"
    Else
        MDIMAIN.LBLTHPREVIEW.Caption = "N"
    End If
    If StBarcode.Value = 1 Then
        MDIMAIN.StatusBar.Panels(15).Text = "Y"
    Else
        MDIMAIN.StatusBar.Panels(15).Text = "N"
    End If
    If StDiscount.Value = 1 Then
        MDIMAIN.StatusBar.Panels(16).Text = "Y"
    Else
        MDIMAIN.StatusBar.Panels(16).Text = "N"
    End If
    If ChkBarcode.Value = 1 Then
        MDIMAIN.StatusBar.Panels(6).Text = "Y"
        'MDIMAIN.barcode_profile.Caption = Cmbprofile.ListIndex
    Else
        MDIMAIN.StatusBar.Panels(6).Text = "N"
        'MDIMAIN.barcode_profile.Caption = ""
    End If
    MDIMAIN.barcode_profile.Caption = ""
    If Val(TxtCopy8B.Text) <= 0 Then
        MDIMAIN.StatusBar.Panels(10).Text = "1"
    Else
        MDIMAIN.StatusBar.Panels(10).Text = Val(TxtCopy8B.Text)
    End If
    If Val(TxtCopy8.Text) <= 0 Then
        MDIMAIN.StatusBar.Panels(11).Text = "1"
    Else
        MDIMAIN.StatusBar.Panels(11).Text = Val(TxtCopy8.Text)
    End If
    If Val(TxtCopy8V.Text) <= 0 Then
        MDIMAIN.StatusBar.Panels(12).Text = "1"
    Else
        MDIMAIN.StatusBar.Panels(12).Text = Val(TxtCopy8V.Text)
    End If
    If Val(TxtTrCopies.Text) <= 0 Then
        MDIMAIN.LBLTRCopy.Caption = "1"
    Else
        MDIMAIN.LBLTRCopy.Caption = Val(TxtTrCopies.Text)
    End If
    
    bill_for = Trim(txtbillformat)
    billprinter = CmbBillprinter.ListIndex
    billprinterA5 = CmbBillprinterA5.ListIndex
    thermalprinter = Cmbthermalprinter.ListIndex
    barcodeprinter = Cmbbarcode.ListIndex
    If CmbDPrint.ListIndex = -1 Then
        D_PRINT = 0
    Else
        D_PRINT = CmbDPrint.ListIndex
    End If
    Dim ObjFso
    Dim StrFileName
    Dim ObjFile
    If FileExists(App.Path & "\BillPrint") Then
        Kill (App.Path & "\BillPrint")
    End If
    Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
    Set ObjFile = ObjFso.CreateTextFile(App.Path & "\BillPrint")
    ObjFile.WriteLine CmbBillprinter.ListIndex
    ObjFile.WriteLine CmbBillprinterA5.ListIndex
    ObjFile.WriteLine Cmbthermalprinter.ListIndex
    ObjFile.WriteLine Cmbbarcode.ListIndex
    
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "SHOP INFORMATION"
Exit Sub
ERRHAND:
    MsgBox (err.Description)
        
End Sub

Private Sub CmdWallPaper_Click()
    fRMLOGO.Show
    fRMLOGO.SetFocus
End Sub

Private Sub Form_Load()
    Dim RSTCOMPANY As ADODB.Recordset
    DTFROM.Value = Date
    DTTo.Value = Date
    On Error GoTo ERRHAND
    
    Call fill_Printer
    
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        TxtCode.Text = IIf(IsNull(RSTCOMPANY!CUST_CODE), "", RSTCOMPANY!CUST_CODE)
        txtServer.Text = IIf(IsNull(RSTCOMPANY!SERVER_ADD), "", RSTCOMPANY!SERVER_ADD)
        If RSTCOMPANY!CLOUD_TYPE = "1" Then
            OptApp.Value = True
        Else
            OptDesk.Value = True
        End If
        txtsupplier.Text = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
        If RSTCOMPANY!REMOTE_FLAG = "Y" Then
            ChkCloud.Value = 1
        Else
            ChkCloud.Value = 0
        End If
        If RSTCOMPANY!MULTI_FLAG = "Y" Then
            ChkMulti.Value = 1
        Else
            ChkMulti.Value = 0
        End If
        txtaddress.Text = IIf(IsNull(RSTCOMPANY!Address), "", RSTCOMPANY!Address)
        txttelno.Text = IIf(IsNull(RSTCOMPANY!TEL_NO), "", RSTCOMPANY!TEL_NO)
        txtfaxno.Text = IIf(IsNull(RSTCOMPANY!FAX_NO), "", RSTCOMPANY!FAX_NO)
        txtemail.Text = IIf(IsNull(RSTCOMPANY!EMAIL_ADD), "", RSTCOMPANY!EMAIL_ADD)
        txtkgst.Text = IIf(IsNull(RSTCOMPANY!ST_ADDRESS), "", RSTCOMPANY!ST_ADDRESS)
        TxtCST.Text = IIf(IsNull(RSTCOMPANY!CST), "", RSTCOMPANY!CST)
        Txtstate.Text = IIf(IsNull(RSTCOMPANY!SCODE), "32", RSTCOMPANY!SCODE)
        TXTSTATENAME.Text = IIf(IsNull(RSTCOMPANY!SNAME), "KL", RSTCOMPANY!SNAME)
        txtdlno.Text = IIf(IsNull(RSTCOMPANY!DL_NO), "", RSTCOMPANY!DL_NO)
        TxtMLNO.Text = IIf(IsNull(RSTCOMPANY!ML_NO), "", RSTCOMPANY!ML_NO)
        TXTREMARKS.Text = IIf(IsNull(RSTCOMPANY!HO_NAME), "", RSTCOMPANY!HO_NAME)
        TxtPinCode.Text = IIf(IsNull(RSTCOMPANY!PINCODE), "", RSTCOMPANY!PINCODE)
        TxtAuthKey.Text = IIf(IsNull(RSTCOMPANY!auth_key), "", RSTCOMPANY!auth_key)
        If TxtAuthKey.Text = "" Then
            DTEInvoice.Text = ""
        Else
            DTEInvoice.Text = IIf(IsNull(RSTCOMPANY!auth_date), "", RSTCOMPANY!auth_date)
        End If
        txtmessage.Text = IIf(IsNull(RSTCOMPANY!INV_MSGS), "", RSTCOMPANY!INV_MSGS)
        TXT8BPRE.Text = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
        TXT8BSUF.Text = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        TXT8PRE.Text = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
        TXT8SUF.Text = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
        TXT8VPRE.Text = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
        TXT8VSUF.Text = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        TxtTrPrefix.Text = IIf(IsNull(RSTCOMPANY!PREFIX_TR), "", RSTCOMPANY!PREFIX_TR)
        TxtTrSuf.Text = IIf(IsNull(RSTCOMPANY!SUFIX_TR), "", RSTCOMPANY!SUFIX_TR)
        TxtVehicle.Text = IIf(IsNull(RSTCOMPANY!VEHICLE), "", RSTCOMPANY!VEHICLE)
        TxtCopy8B.Text = IIf(IsNull(RSTCOMPANY!Copy_8B), "", RSTCOMPANY!Copy_8B)
        TxtCopy8.Text = IIf(IsNull(RSTCOMPANY!Copy_8), "", RSTCOMPANY!Copy_8)
        TxtCopy8V.Text = IIf(IsNull(RSTCOMPANY!Copy_8V), "", RSTCOMPANY!Copy_8V)
        TxtTrCopies.Text = IIf(IsNull(RSTCOMPANY!Copy_TR), "", RSTCOMPANY!Copy_TR)
        TxtCGST.Text = IIf(IsNull(RSTCOMPANY!CGST), 0, RSTCOMPANY!CGST)
        TxtSGST.Text = IIf(IsNull(RSTCOMPANY!SGST), 0, RSTCOMPANY!SGST)
        TXTIGST.Text = IIf(IsNull(RSTCOMPANY!IGST), 0, RSTCOMPANY!IGST)
        txtinvterms.Text = IIf(IsNull(RSTCOMPANY!INV_TERMS), "", RSTCOMPANY!INV_TERMS)
        Txtbank.Text = IIf(IsNull(RSTCOMPANY!bank_details), "", RSTCOMPANY!bank_details)
        TxtPAN.Text = IIf(IsNull(RSTCOMPANY!PAN_NO), "", RSTCOMPANY!PAN_NO)
        txtRTDisc.Text = IIf(IsNull(RSTCOMPANY!RTDISC), "", RSTCOMPANY!RTDISC)
        TxtWSDisc.Text = IIf(IsNull(RSTCOMPANY!WSDISC), "", RSTCOMPANY!WSDISC)
        TxtVPDisc.Text = IIf(IsNull(RSTCOMPANY!VPDISC), "", RSTCOMPANY!VPDISC)
        TxtMRPDisc.Text = IIf(IsNull(RSTCOMPANY!MRPDISC), "", RSTCOMPANY!MRPDISC)
        TxtHSNSum.Text = IIf(IsNull(RSTCOMPANY!HSN_SUM), "", RSTCOMPANY!HSN_SUM)
        
        If RSTCOMPANY!T2_COPIES = "Y" Then
            ChkThermalcopies.Value = 1
        Else
            ChkThermalcopies.Value = 0
        End If
        If RSTCOMPANY!STK_ADJ = "Y" Then
            ChkStkadjst.Value = 1
        Else
            ChkStkadjst.Value = 0
        End If
        If RSTCOMPANY!GST_FLAG = "C" Then
            OptCompound.Value = True
        ElseIf RSTCOMPANY!GST_FLAG = "N" Then
            OptNonGST.Value = True
        Else
            OptRegular.Value = True
        End If
        If RSTCOMPANY!TERMS_FLAG = "Y" Then
            chkTerms.Value = 1
            frmTerms.Visible = True
            Terms1.Text = IIf(IsNull(RSTCOMPANY!Terms1), "", RSTCOMPANY!Terms1)
            Terms2.Text = IIf(IsNull(RSTCOMPANY!Terms2), "", RSTCOMPANY!Terms2)
            Terms3.Text = IIf(IsNull(RSTCOMPANY!Terms3), "", RSTCOMPANY!Terms3)
            Terms4.Text = IIf(IsNull(RSTCOMPANY!Terms4), "", RSTCOMPANY!Terms4)
        Else
            chkTerms.Value = 0
            frmTerms.Visible = False
            Terms1.Text = ""
            Terms2.Text = ""
            Terms3.Text = ""
            Terms4.Text = ""
        End If
        If RSTCOMPANY!NSPT = "Y" Then
            ChKNSPT.Value = 1
        Else
            ChKNSPT.Value = 0
        End If
        If RSTCOMPANY!ALL_PRN = "Y" Then
            CHKPRNALL.Value = 1
        Else
            CHKPRNALL.Value = 0
        End If
        If RSTCOMPANY!REMOVE_UBILL = "N" Then
            ChkRemoveUBILL.Value = 0
        Else
            ChkRemoveUBILL.Value = 1
        End If
        If RSTCOMPANY!MRP_DISC = "Y" Then
            ChkMRPDisc.Value = 1
        Else
            ChkMRPDisc.Value = 0
        End If
        If RSTCOMPANY!EXP_ENABLED = "Y" Then
            chkexport.Value = 1
        Else
            chkexport.Value = 0
        End If
        If RSTCOMPANY!UB = "Y" Then
            chkub.Value = 1
        Else
            chkub.Value = 0
        End If
        If RSTCOMPANY!ONLINE_BILL = "Y" Then
            ChkOnlineBill.Value = 1
        Else
            ChkOnlineBill.Value = 0
        End If
        If RSTCOMPANY!DMP_FLAG = "Y" Then
            ChKDMP.Value = 1
        Else
            ChKDMP.Value = 0
        End If
        If RSTCOMPANY!DMP_MINI = "Y" Then
            ChkDMPMini.Value = 1
        Else
            ChkDMPMini.Value = 0
        End If
        If RSTCOMPANY!LINE_SPACE = "Y" Then
            chkspace.Value = 1
        Else
            chkspace.Value = 0
        End If
        If RSTCOMPANY!DMPTH_FLAG = "Y" Then
            ChkDMPThermal.Value = 1
        Else
            ChkDMPThermal.Value = 0
        End If
        If RSTCOMPANY!TAX_FLAG = "Y" Then
            Chkwithtax.Value = 1
        Else
            Chkwithtax.Value = 0
        End If
        If RSTCOMPANY!KFCNET = "Y" Then
            ChkKFCNet.Value = 1
        Else
            ChkKFCNet.Value = 0
        End If
        If RSTCOMPANY!TAXWRN_FLAG = "Y" Then
            ChkTaxWarn.Value = 1
        Else
            ChkTaxWarn.Value = 0
        End If
        If RSTCOMPANY!GSTWRN_FLAG = "Y" Then
            chkgst.Value = 1
        Else
            chkgst.Value = 0
        End If
        If RSTCOMPANY!ITEM_WARN = "Y" Then
            CHKITEMREPEAT.Value = 1
        Else
            CHKITEMREPEAT.Value = 0
        End If
        If RSTCOMPANY!AMC_FLAG = "Y" Then
            CHKAMC.Value = 1
        Else
            CHKAMC.Value = 0
        End If
        If RSTCOMPANY!HOLD_THERMAL_FLAG = "Y" Then
            ChkStpThermal.Value = 1
        Else
            ChkStpThermal.Value = 0
        End If
        If RSTCOMPANY!FORM_62 = "Y" Then
            Chk62.Value = 1
        Else
            Chk62.Value = 0
        End If
        If RSTCOMPANY!DUP_FLAG = "N" Then
            CHKDUP.Value = 0
        Else
            CHKDUP.Value = 1
        End If
        If RSTCOMPANY!ROUND_FLAG = "N" Then
            ChkRound.Value = 0
        Else
            ChkRound.Value = 1
        End If
        If RSTCOMPANY!BATCH_FLAG = "Y" Then
            ChkBatch.Value = 1
        Else
            ChkBatch.Value = 0
        End If
        txtstkcrct.Text = IIf(IsNull(RSTCOMPANY!STOCK_CRCT), 0, RSTCOMPANY!STOCK_CRCT)
        txtCalc.Text = IIf(IsNull(RSTCOMPANY!CLCODE) Or RSTCOMPANY!CLCODE = 0, "", RSTCOMPANY!CLCODE)
        TxtPCode.Text = IIf(IsNull(RSTCOMPANY!PCODE) Or RSTCOMPANY!PCODE = 0, "", RSTCOMPANY!PCODE)
        txtLabel.Text = IIf(IsNull(RSTCOMPANY!BAR_LABELS) Or RSTCOMPANY!BAR_LABELS = 0, 1, RSTCOMPANY!BAR_LABELS)
        txtbillformat.Text = IIf(IsNull(RSTCOMPANY!BILL_FORMAT), "", RSTCOMPANY!BILL_FORMAT)
        If RSTCOMPANY!PRN_PETTY_FLAG = "Y" Then
            ChkPrnPetty.Value = 1
        Else
            ChkPrnPetty.Value = 0
        End If
        If RSTCOMPANY!SAL_PROC = "Y" Then
            ChkSalary.Value = 1
        Else
            ChkSalary.Value = 0
        End If
        If RSTCOMPANY!CAT_PURCHASE = "Y" Then
            ChkCatPur.Value = 1
        Else
            ChkCatPur.Value = 0
        End If
        If RSTCOMPANY!RST_BILL = "Y" Then
            ChkRstBill.Value = 1
        Else
            ChkRstBill.Value = 0
        End If
        If RSTCOMPANY!PRICE_SPLIT = "Y" Then
            ChkPriceSplit.Value = 1
        Else
            ChkPriceSplit.Value = 0
        End If
        If RSTCOMPANY!PER_PURCHASE = "Y" Then
            ChkPercPurchase.Value = 1
        Else
            ChkPercPurchase.Value = 0
        End If
        If RSTCOMPANY!PREVIEW_FLAG = "Y" Then
            ChkPreview.Value = 1
        Else
            ChkPreview.Value = 0
        End If
        If RSTCOMPANY!PREVIEWTH_FLAG = "Y" Then
            ChkThPreview.Value = 1
        Else
            ChkThPreview.Value = 0
        End If
        If RSTCOMPANY!CODE_FLAG = "Y" Then
            StBarcode.Value = 1
        Else
            StBarcode.Value = 0
        End If
        If RSTCOMPANY!DISC_FLAG = "Y" Then
            StDiscount.Value = 1
        Else
            StDiscount.Value = 0
        End If
        On Error Resume Next
        If RSTCOMPANY!BARCODE_FLAG = "Y" Then
            ChkBarcode.Value = 1
            'Cmbprofile.Visible = True
            'Cmbprofile.ListIndex = IIf(IsNull(RSTCOMPANY!barcode_profile), -1, RSTCOMPANY!barcode_profile)
        Else
            ChkBarcode.Value = 0
            'Cmbprofile.ListIndex = -1
            'Cmbprofile.Visible = False
        End If
        CmbDPrint.ListIndex = IIf(IsNull(RSTCOMPANY!D_PRINT) Or RSTCOMPANY!D_PRINT = "", 0, RSTCOMPANY!D_PRINT)
        
        On Error GoTo ERRHAND
        If RSTCOMPANY!OSSR_FLAG = "Y" Then
            chkoutSR.Value = 1
        Else
            chkoutSR.Value = 0
        End If
        If RSTCOMPANY!OSB2C_FLAG = "Y" Then
            chkoutB2C.Value = 1
        Else
            chkoutB2C.Value = 0
        End If
        If RSTCOMPANY!OSB2B_FLAG = "Y" Then
            chkoutB2B.Value = 1
        Else
            chkoutB2B.Value = 0
        End If
        If RSTCOMPANY!OSPTY_FLAG = "Y" Then
            chkoutPTY.Value = 1
        Else
            chkoutPTY.Value = 0
        End If
        
        '=
        If RSTCOMPANY!SHOP_RT = "Y" Then
            OptRetail.Value = True
        Else
            OptWs.Value = True
        End If
        If RSTCOMPANY!hide_spec = "Y" Then
            chkspec.Value = 1
        Else
            chkspec.Value = 0
        End If
        If RSTCOMPANY!hide_pr_name = "Y" Then
            chkprnname.Value = 1
        Else
            chkprnname.Value = 0
        End If
        If RSTCOMPANY!hide_wrnty = "Y" Then
            chkwrnty.Value = 1
        Else
            chkwrnty.Value = 0
        End If
        If RSTCOMPANY!hide_serial = "Y" Then
            chkserial.Value = 1
        Else
            chkserial.Value = 0
        End If
        If RSTCOMPANY!hide_mrp = "Y" Then
            chkmrp.Value = 1
        Else
            chkmrp.Value = 0
        End If
        If RSTCOMPANY!billtype_flag = "Y" Then
            chkbilltype.Value = 1
        Else
            chkbilltype.Value = 0
        End If
        If RSTCOMPANY!item_repeat = "Y" Then
            chkrepeat.Value = 1
        Else
            chkrepeat.Value = 0
        End If
        If RSTCOMPANY!Zero_Warn = "N" Then
            ChkZeroWarn.Value = 0
        Else
            ChkZeroWarn.Value = 1
        End If
        If RSTCOMPANY!VS_FLAG = "Y" Then
            chkpcflag.Value = 1
        Else
            chkpcflag.Value = 0
        End If
        If RSTCOMPANY!MINUS_BILL = "N" Then
            chkminusbill.Value = 1
        Else
            chkminusbill.Value = 0
        End If
        If RSTCOMPANY!LMT_FLAG = "Y" Then
            ChkLimitd.Value = 1
        Else
            ChkLimitd.Value = 0
        End If
        If RSTCOMPANY!BAR_TEMPLATE = "Y" Then
            Chktemplate.Value = 1
        Else
            Chktemplate.Value = 0
        End If
        If RSTCOMPANY!MOB_WARN_FLAG = "Y" Then
            ChkMobile.Value = 1
        Else
            ChkMobile.Value = 0
        End If
        If RSTCOMPANY!hide_expiry = "Y" Then
            chkexpiry.Value = 1
        Else
            chkexpiry.Value = 0
        End If
        If RSTCOMPANY!hide_free = "Y" Then
            ChkFree.Value = 1
        Else
            ChkFree.Value = 0
        End If
        If RSTCOMPANY!hide_disc = "Y" Then
            chkdisc.Value = 1
        Else
            chkdisc.Value = 0
        End If
        If RSTCOMPANY!hide_terms = "Y" Then
            chkhideterms.Value = 1
        Else
            chkhideterms.Value = 0
        End If
        If RSTCOMPANY!hide_deliver = "Y" Then
            chkdeliver.Value = 1
        Else
            chkdeliver.Value = 0
        End If
        If RSTCOMPANY!mrp_plus = "Y" Then
            chkmrpplus.Value = 1
        Else
            chkmrpplus.Value = 0
        End If
        If RSTCOMPANY!billtype_flag = "Y" Then
            chkbilltype.Value = 1
        Else
            chkbilltype.Value = 0
        End If
'        If RSTCOMPANY!item_repeat = "Y" Then
'            chkrepeat.value = 1
'        Else
'            chkrepeat.value = 0
'        End If
        If RSTCOMPANY!kfc_flag = "Y" Then
            chkkfc.Value = 1
            If IsDate(RSTCOMPANY!KFCFROM_DATE) Then
                DTFROM.Value = RSTCOMPANY!KFCFROM_DATE
            End If
            If IsDate(RSTCOMPANY!KFCTO_DATE) Then
                DTTo.Value = RSTCOMPANY!KFCTO_DATE
            End If
        Else
            chkkfc.Value = 0
        End If
        
        If ChKDMP.Value = 1 Or ChkDMPThermal.Value = 1 Then
            chkspace.Visible = True
        Else
            chkspace.Visible = False
        End If
        
        If RSTCOMPANY!SCHEME_OPT = "1" Then
            OptScheme2.Value = True
        Else
            OptScheme1.Value = True
        End If
        '=
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim ObjFso
    Dim StrFileName
    Dim ObjFile
    If FileExists(App.Path & "\BillPrint") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\BillPrint")  'Reading from the file
        On Error Resume Next
        CmbBillprinter.ListIndex = ObjFile.ReadLine
        CmbBillprinterA5.ListIndex = ObjFile.ReadLine
        Cmbthermalprinter.ListIndex = ObjFile.ReadLine
        Cmbbarcode.ListIndex = ObjFile.ReadLine
        err.Clear
        On Error GoTo ERRHAND
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    
    REPFLAG = True
    'Me.Width = 7000
    'Me.Height = 7500
    Me.Left = 500
    Me.Top = 0
   
    Exit Sub
ERRHAND:
    MsgBox (err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RSTREP.State = 1 Then RSTREP.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub OptCompound_Click()
    chkkfc.Value = 0
    ChkKFCNet.Value = 0
    FrmKFC.Visible = False
End Sub

Private Sub OptNonGST_Click()
    chkkfc.Value = 0
    ChkKFCNet.Value = 0
    FrmKFC.Visible = False
End Sub

Private Sub OptRegular_Click()
    FrmKFC.Visible = True
End Sub

Private Sub txtaddress_GotFocus()
    txtaddress.SelStart = 0
    txtaddress.SelLength = Len(txtaddress.Text)
End Sub

Private Sub txtaddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txttelno.SetFocus
    End Select
End Sub

Private Sub txtbillformat_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCGST_Change()
    TxtSGST.Text = 100 - Val(TxtCGST.Text)
End Sub

Private Sub TxtCGST_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtPincode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCompCode_Change()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If TxtCompCode.Text = "*111#" Then
        FrmeCode.Visible = True
    Else
        FrmeCode.Visible = False
    End If
    If TxtCompCode.Text = "*1707#" Then
        FrmAuth.Visible = True
    Else
        FrmAuth.Visible = False
    End If
End Sub

Private Sub TxtCopy8_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCopy8B_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCopy8V_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcst_GotFocus()
    TxtCST.SelStart = 0
    TxtCST.SelLength = Len(TxtCST.Text)
End Sub

Private Sub txtcst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtdlno.SetFocus
    End Select
End Sub

Private Sub Txtstate_GotFocus()
    If Trim(Txtstate.Text) = "" Then Txtstate.Text = "32"
    If Val(Txtstate.Text) = 0 Then Txtstate.Text = "32"
    If Len(Txtstate.Text) <> 2 Then Txtstate.Text = "32"
    If Val(Txtstate.Text) = 32 Then TXTSTATENAME.Text = "KL"
    Txtstate.SelStart = 0
    Txtstate.SelLength = Len(TxtCST.Text)
End Sub

Private Sub Txtstate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtdlno.SetFocus
    End Select
End Sub

Private Sub Txtstate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTSTATENAME_GotFocus()
    If Trim(Txtstate.Text) = "" Then Txtstate.Text = "32"
    If Val(Txtstate.Text) = 0 Then Txtstate.Text = "32"
    If Val(Txtstate.Text) = 32 Then TXTSTATENAME.Text = "KL"
    TXTSTATENAME.SelStart = 0
    TXTSTATENAME.SelLength = Len(TxtCST.Text)
End Sub

Private Sub TXTSTATENAME_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtdlno.SetFocus
    End Select
End Sub

Private Sub txtdlno_GotFocus()
    txtdlno.SelStart = 0
    txtdlno.SelLength = Len(txtdlno.Text)
End Sub

Private Sub txtdlno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtMLNO.SetFocus
    End Select
End Sub

Private Sub TxtHSNSum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtMLNO_Change()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If TxtMLNO.Text = "*101*" Then
        FRMEINVISIBLE.Visible = True
        Frame4.Visible = True
    Else
        FRMEINVISIBLE.Visible = False
        Frame4.Visible = False
    End If
End Sub

Private Sub txtmlno_GotFocus()
    TxtMLNO.SelStart = 0
    TxtMLNO.SelLength = Len(TxtMLNO.Text)
End Sub

Private Sub txtmlno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTREMARKS.SetFocus
    End Select
End Sub
Private Sub txtemail_GotFocus()
    txtemail.SelStart = 0
    txtemail.SelLength = Len(txtemail.Text)
End Sub

Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtkgst.SetFocus
    End Select
End Sub

Private Sub txtfaxno_GotFocus()
    txtfaxno.SelStart = 0
    txtfaxno.SelLength = Len(txtfaxno.Text)
End Sub

Private Sub txtfaxno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtemail.SetFocus
    End Select
End Sub

Private Sub TXTIGST_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtkgst_GotFocus()
    txtkgst.SelStart = 0
    txtkgst.SelLength = Len(txtkgst.Text)
End Sub

Private Sub txtkgst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtCST.SetFocus
    End Select
End Sub

Private Sub txtmessage_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdSave.SetFocus
    End Select
End Sub

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtPinCode.SetFocus
    End Select
End Sub

Private Sub txtRTDisc_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtSGST_Change()
    TxtCGST.Text = 100 - Val(TxtSGST.Text)
End Sub

Private Sub TxtSGST_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtstkcrct_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCalc_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtLabel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtsupplier_GotFocus()
    txtsupplier.SelStart = 0
    txtsupplier.SelLength = Len(txtsupplier.Text)
   
End Sub

Private Sub txtsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtsupplier.Text = "" Then
                MsgBox "ENTER NAME OF SHOP", vbOKOnly, "SHOP INFORMATION"
                txtsupplier.SetFocus
                Exit Sub
            End If
         txtaddress.SetFocus
    End Select
    
End Sub

Private Sub txtsupplier_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case Asc("")
'            KeyAscii = 0
'        'Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'        '    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    End Select
End Sub

Private Sub txttelno_GotFocus()
    txttelno.SelStart = 0
    txttelno.SelLength = Len(txttelno.Text)
End Sub

Private Sub txttelno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtfaxno.SetFocus
    End Select
End Sub

Private Sub txtmessage_GotFocus()
    txtmessage.SelStart = 0
    txtmessage.SelLength = Len(txtmessage.Text)
End Sub

Private Sub TXT8BPRE_GotFocus()
    TXT8BPRE.SelStart = 0
    TXT8BPRE.SelLength = Len(TXT8BPRE.Text)
End Sub

Private Sub TXT8BPRE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXT8BSUF.SetFocus
    End Select
End Sub

Private Sub TXT8BSUF_GotFocus()
    TXT8BSUF.SelStart = 0
    TXT8BSUF.SelLength = Len(TXT8BSUF.Text)
End Sub

Private Sub TXT8BSUF_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXT8PRE.SetFocus
    End Select
End Sub

Private Sub TXT8PRE_GotFocus()
    TXT8PRE.SelStart = 0
    TXT8PRE.SelLength = Len(TXT8PRE.Text)
End Sub

Private Sub TXT8PRE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXT8SUF.SetFocus
    End Select
End Sub

Private Sub TXT8SUF_GotFocus()
    TXT8SUF.SelStart = 0
    TXT8SUF.SelLength = Len(TXT8SUF.Text)
End Sub

Private Sub TXT8SUF_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtmessage.SetFocus
    End Select
End Sub

Private Sub TxtVehicle_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtVPDisc_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtMRPDISC_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtWSDisc_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Function fill_Printer()
    Dim p, PNAME
    'Dim printerfound As Boolean
    'printerfound = False
    On Error GoTo ERRHAND
    For Each p In Printers
        PNAME = p.DeviceName
'        If UCase(Right(PNAME, 16)) Like "BAR CODE PRINTER" Then
'            Set Printer = P
'            printerfound = True
'            Exit For
'        End If
        CmbBillprinter.AddItem (p.DeviceName)
        CmbBillprinterA5.AddItem (p.DeviceName)
        Cmbthermalprinter.AddItem (p.DeviceName)
        Cmbbarcode.AddItem (p.DeviceName)
    Next p
'    If printerfound = False Then
'        MsgBox ("Barcode Printer not found. Please correct the printer name")
'        Exit Function
'    End If
    
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub TxtPincode_GotFocus()
    TxtPinCode.SelStart = 0
    TxtPinCode.SelLength = Len(TxtPinCode.Text)
End Sub

Private Sub TxtPincode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXT8BPRE.SetFocus
    End Select
End Sub

Private Sub TxtTrCopies_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
