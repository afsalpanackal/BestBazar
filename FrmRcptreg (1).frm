VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMRcptReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEIPT REGISTER"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRcptreg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   18645
   Begin VB.Frame Frmereceipt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PRESS ESC TO CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   7455
      TabIndex        =   60
      Top             =   2505
      Visible         =   0   'False
      Width           =   5265
      Begin VB.TextBox TXTsample2 
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
         Left            =   1890
         TabIndex        =   62
         Top             =   1050
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid grdreceipts 
         Height          =   3825
         Left            =   45
         TabIndex        =   61
         Top             =   195
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   6747
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   400
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   120
      Picture         =   "FrmRcptreg.frx":030A
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   0
      Picture         =   "FrmRcptreg.frx":064C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   44
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSFlexGridLib.MSFlexGrid grdcount 
      Height          =   5190
      Left            =   11160
      TabIndex        =   43
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   9155
      _Version        =   393216
      Rows            =   1
      Cols            =   25
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      FillStyle       =   1
      SelectionMode   =   1
      Appearance      =   0
      GridLineWidth   =   2
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
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      Caption         =   "PRESS ESC TO CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5520
      Left            =   75
      TabIndex        =   12
      Top             =   1740
      Visible         =   0   'False
      Width           =   11085
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4140
         Left            =   45
         TabIndex        =   13
         Top             =   1335
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   7303
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   400
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Qty"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   15
         Left            =   3825
         TabIndex        =   70
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Lblqty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3705
         TabIndex        =   69
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLINVDATE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   135
         TabIndex        =   24
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV DATE"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   4
         Left            =   405
         TabIndex        =   23
         Top             =   735
         Width           =   885
      End
      Begin VB.Label LBLSUPPLIER 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1140
         TabIndex        =   22
         Top             =   315
         Width           =   4410
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer\"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   21
         Top             =   345
         Width           =   885
      End
      Begin VB.Label LBLBILLAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2535
         TabIndex        =   17
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV AMT"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   2685
         TabIndex        =   16
         Top             =   735
         Width           =   810
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV NO"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   1665
         TabIndex        =   15
         Top             =   735
         Width           =   675
      End
      Begin VB.Label LBLBILLNO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1515
         TabIndex        =   14
         Top             =   975
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFC0&
      Height          =   8925
      Left            =   -75
      TabIndex        =   0
      Top             =   -150
      Width           =   18735
      Begin VB.CommandButton CmdChqRet 
         Caption         =   "Cheque Return"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   12855
         TabIndex        =   68
         Top             =   825
         Width           =   870
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Print Clo. Amt for All Customers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   17325
         TabIndex        =   67
         Top             =   825
         Width           =   1350
      End
      Begin VB.CheckBox ChkVerify 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Verify"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   17580
         TabIndex        =   66
         Top             =   180
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ChkSmry 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Summary"
         Height          =   225
         Left            =   11580
         TabIndex        =   65
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CMDDRCRNOTE 
         Caption         =   "Make Debit / Credit "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   11790
         TabIndex        =   8
         Top             =   825
         Width           =   1050
      End
      Begin VB.CommandButton CMDITEM 
         Caption         =   "Print Item Details of Selected Bills"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   13740
         TabIndex        =   9
         Top             =   825
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Ledger for All Customers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   16050
         TabIndex        =   11
         Top             =   825
         Width           =   1260
      End
      Begin VB.CommandButton CmdMKRcpt 
         Caption         =   "Make Receipts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   10800
         TabIndex        =   7
         Top             =   825
         Width           =   960
      End
      Begin VB.CommandButton CmdPrnRcpt 
         Caption         =   "Print &Receipt /Dr /Cr Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   8775
         TabIndex        =   5
         Top             =   825
         Width           =   1100
      End
      Begin VB.CommandButton CmdPrintDet 
         Caption         =   "Detai&led Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7905
         TabIndex        =   4
         Top             =   825
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print Ledger"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9885
         TabIndex        =   6
         Top             =   825
         Width           =   900
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "TOTAL"
         ForeColor       =   &H000000FF&
         Height          =   885
         Left            =   105
         TabIndex        =   26
         Top             =   8010
         Width           =   17175
         Begin VB.CommandButton CmdRemove 
            Caption         =   "Mark as &Not Paid"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   15945
            TabIndex        =   49
            Top             =   225
            Width           =   1185
         End
         Begin VB.CommandButton CmdPay 
            Caption         =   "&Mark as Paid"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   14730
            TabIndex        =   48
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Comm Pend"
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
            Height          =   315
            Index           =   14
            Left            =   13440
            TabIndex        =   55
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label lblcommpend 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   13365
            TabIndex        =   54
            Top             =   420
            Width           =   1305
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Comm Paid"
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
            Height          =   315
            Index           =   13
            Left            =   12075
            TabIndex        =   53
            Top             =   165
            Width           =   1140
         End
         Begin VB.Label lblcommpaid 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   12000
            TabIndex        =   52
            Top             =   420
            Width           =   1320
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Comm Amt"
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
            Height          =   240
            Index           =   12
            Left            =   10680
            TabIndex        =   51
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label lblcomm 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   10590
            TabIndex        =   50
            Top             =   420
            Width           =   1350
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Selected Amount"
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
            Height          =   315
            Index           =   11
            Left            =   8415
            TabIndex        =   47
            Top             =   75
            Width           =   1815
         End
         Begin VB.Label LBLSelected 
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
            ForeColor       =   &H00000080&
            Height          =   450
            Left            =   8340
            TabIndex        =   46
            Top             =   300
            Width           =   2190
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Op. Balance"
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
            Height          =   300
            Index           =   8
            Left            =   645
            TabIndex        =   35
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label lblOPBal 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   600
            TabIndex        =   34
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLPAIDAMT 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   4575
            TabIndex        =   32
            Top             =   435
            Width           =   1875
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   3
            Left            =   4590
            TabIndex        =   31
            Top             =   150
            Width           =   1875
         End
         Begin VB.Label LBLINVAMT 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   2595
            TabIndex        =   30
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Debit Amt"
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
            Height          =   300
            Index           =   6
            Left            =   2640
            TabIndex        =   29
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label LBLBALAMT 
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   6465
            TabIndex        =   28
            Top             =   435
            Width           =   1830
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bal Amt"
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
            Height          =   315
            Index           =   7
            Left            =   6465
            TabIndex        =   27
            Top             =   150
            Width           =   1815
         End
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   15090
         TabIndex        =   10
         Top             =   825
         Width           =   915
      End
      Begin VB.CommandButton CMDDISPLAY 
         Caption         =   "&DISPLAY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6960
         TabIndex        =   3
         Top             =   825
         Width           =   945
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0FFC0&
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
         Height          =   1305
         Left            =   120
         TabIndex        =   18
         Top             =   105
         Width           =   6840
         Begin VB.TextBox TxtCode 
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   1095
            TabIndex        =   36
            Top             =   210
            Width           =   1875
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
            Height          =   360
            Left            =   3030
            TabIndex        =   1
            Top             =   210
            Width           =   3735
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   3030
            TabIndex        =   2
            Top             =   585
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
         Begin VB.PictureBox rptPRINT 
            Height          =   480
            Left            =   9990
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   33
            Top             =   -45
            Width           =   1200
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Press F7 to make Dr/Cr Notes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   38
            Top             =   960
            Width           =   3030
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Press F6 to make Receipts"
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
            Height          =   300
            Index           =   8
            Left            =   90
            TabIndex        =   37
            Top             =   690
            Width           =   2730
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "CUSTOMER"
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
            Height          =   315
            Index           =   5
            Left            =   45
            TabIndex        =   25
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   19
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   20
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   8205
         TabIndex        =   39
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   114098177
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   10035
         TabIndex        =   40
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         Format          =   114098177
         CurrentDate     =   40498
      End
      Begin VB.Frame Frame1 
         Height          =   6705
         Left            =   105
         TabIndex        =   57
         Top             =   1305
         Width           =   18600
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   390
            Left            =   13875
            TabIndex        =   64
            Top             =   1305
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   688
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
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
            Left            =   13545
            TabIndex        =   59
            Top             =   2130
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSFlexGridLib.MSFlexGrid GRDTranx 
            Height          =   6570
            Left            =   15
            TabIndex        =   58
            Top             =   120
            Width           =   18585
            _ExtentX        =   32782
            _ExtentY        =   11589
            _Version        =   393216
            Rows            =   1
            Cols            =   31
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            BackColorBkg    =   12632256
            FocusRect       =   2
            AllowUserResizing=   3
            Appearance      =   0
            GridLineWidth   =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label lbladdress 
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
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   12810
         TabIndex        =   63
         Top             =   210
         Width           =   4335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblagent 
         Caption         =   " "
         Height          =   225
         Left            =   11595
         TabIndex        =   56
         Top             =   825
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Period From"
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
         Height          =   270
         Index           =   10
         Left            =   7035
         TabIndex        =   42
         Top             =   405
         Width           =   1140
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
         ForeColor       =   &H00800080&
         Height          =   270
         Index           =   9
         Left            =   9750
         TabIndex        =   41
         Top             =   405
         Width           =   285
      End
   End
End
Attribute VB_Name = "FRMRcptReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CmdChqRet_Click()
    If DataList2.BoundText = "" Then Exit Sub
    Me.Enabled = False
    FRMCHQRET2.LBLSUPPLIER.Caption = DataList2.text
    FRMCHQRET2.lblactcode.Caption = DataList2.BoundText
    'FRMCHQRET2.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
    'FRMCHQRET2.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
    'FRMCHQRET2.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
    'FRMCHQRET2.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
    'FRMCHQRET2.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
    FRMCHQRET2.Show
    FRMCHQRET2.SetFocus
End Sub

Private Sub CmDDisplay_Click()
    Call Fillgrid
End Sub

Private Sub CMDDRCRNOTE_Click()
    Call GRDTranx_KeyDown(vbKeyF7, 0)
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CMDITEM_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    db.Execute "Update trxmast set PRINT_FLAG = 'N' WHERE ACT_CODE = '" & DataList2.BoundText & "'"
    If grdcount.rows = 0 Then
        MsgBox "Nothing selected", vbOKOnly, "EzBiz"
        Exit Sub
    End If
        
    For i = 0 To grdcount.rows - 1
        db.Execute "Update trxmast set PRINT_FLAG = 'Y' WHERE ACT_CODE = '" & DataList2.BoundText & "' and TRX_YEAR = '" & Val(grdcount.TextMatrix(i, 14)) & "' AND VCH_NO = " & Val(grdcount.TextMatrix(i, 3)) & " AND TRX_TYPE = '" & grdcount.TextMatrix(i, 8) & "'"
        
    Next i
    
    ReportNameVar = Rptpath & "RptItemdetail"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='WO')  AND {TRXMAST.ACT_CODE}='" & DataList2.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    Report.RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='SV' OR {TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='WO')  AND {TRXMAST.ACT_CODE}='" & DataList2.BoundText & "' AND {TRXMAST.PRINT_FLAG}='Y' )"
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdMKRcpt_Click()
    Call GRDTranx_KeyDown(vbKeyF6, 0)
End Sub

Private Sub CmdPay_Click()
    If MsgBox("Are you sure you want to make the selected invoices as Paid", vbYesNo + vbDefaultButton2, "Receipt Entry") = vbNo Then Exit Sub
    
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    On Error GoTo ErrHand
    For i = 0 To grdcount.rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE TRX_TYPE='DR' AND CR_NO = " & grdcount.TextMatrix(i, 7) & " AND INV_TRX_TYPE = '" & grdcount.TextMatrix(i, 8) & "' AND TRX_YEAR= '" & grdcount.TextMatrix(i, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE!PAID_FLAG = "Y"
            RSTTRXFILE.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        GRDTranx.TextMatrix(grdcount.TextMatrix(i, 20), 22) = "PAID"
        
        If GRDTranx.TextMatrix(grdcount.TextMatrix(i, 20), 22) = "PAID" Then
            GRDTranx.Row = grdcount.TextMatrix(i, 20)
            GRDTranx.Col = 22
            GRDTranx.CellForeColor = vbBlue
        ElseIf GRDTranx.TextMatrix(grdcount.TextMatrix(i, 20), 22) = "PEND" Then
            GRDTranx.Row = grdcount.TextMatrix(i, 20)
            GRDTranx.Col = 22
            GRDTranx.CellForeColor = vbRed
        End If
        
    Next i
    
    For i = 1 To GRDTranx.rows - 1
        GRDTranx.TextMatrix(i, 21) = "N"
        With GRDTranx
            If .TextMatrix(.Row, 8) = "HI" Or .TextMatrix(.Row, 8) = "GI" Or .TextMatrix(.Row, 8) = "SI" Or .TextMatrix(.Row, 8) = "RI" Or .TextMatrix(.Row, 8) = "VI" Or .TextMatrix(.Row, 8) = "WO" Then
                .Row = i: .Col = 25: .CellPictureAlignment = 4 ' Align the checkbox
                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
            End If
        End With
    Next i
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub CmdPrint_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast As ADODB.Recordset
    Dim i As Long
    
    If DataList2.BoundText = "" Then
        MsgBox "please Select Customer from the List", vbOKOnly, "Statement"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    Op_Bal = 0
    OP_Sale = 0
    OP_Rcpt = 0
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OP_Sale = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='RW' OR TRX_TYPE ='SR' OR TRX_TYPE ='DR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'RD') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        Select Case RSTTRXFILE!TRX_TYPE
            Case "DR", "RD"
                OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
            Case "DB"
                OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            Case Else
                OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
        End Select
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Op_Bal = OP_Sale - OP_Rcpt
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!OPEN_CR = Op_Bal
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim BAL_AMOUNT As Double
    Dim CR_FLAG As Boolean
    CR_FLAG = False
    BAL_AMOUNT = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='DR' OR TRX_TYPE ='RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'RD') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        'RSTTRXFILE!BAL_AMT = Op_Bal
        CR_FLAG = True
        Select Case RSTTRXFILE!TRX_TYPE
            Case "DB"
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            Case "RD"
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
            Case Else
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + (IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                'RSTTRXFILE!BAL_AMT = Op_Bal

        End Select
        RSTTRXFILE!BAL_AMT = BAL_AMOUNT
        Op_Bal = 0
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    db.Execute "DELETE FROM DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE ='AA'"
    If CR_FLAG = False Then
        Dim MAXNO As Double
        MAXNO = 1
        Set RstCustmast = New ADODB.Recordset
        RstCustmast.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'AA'", db, adOpenForwardOnly
        If Not (RstCustmast.EOF And RstCustmast.BOF) Then
            MAXNO = IIf(IsNull(RstCustmast.Fields(0)), 1, RstCustmast.Fields(0) + 1)
        End If
        RstCustmast.Close
        Set RstCustmast = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * FROM DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE ='AA'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "AA"
            RSTTRXFILE!CR_NO = MAXNO
        End If
        RSTTRXFILE!INV_TRX_TYPE = ""
'        RSTTRXFILE!RCPT_DATE = Null
'        RSTTRXFILE!RCPT_AMT = Null
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.text
        RSTTRXFILE!INV_DATE = Format(DTFROM.Value, "DD/MM/YYYY")
        RSTTRXFILE!REF_NO = ""
'        RSTTRXFILE!INV_AMT = Null
'        RSTTRXFILE!INV_NO = Null
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
'        RSTTRXFILE!C_TRX_TYPE = Null
'        'RSTTRXFILE!C_REC_NO = Null
'        RSTTRXFILE!C_INV_TRX_TYPE = Null
'        RSTTRXFILE!C_INV_TYPE = Null
'        ''RSTTRXFILE!C_INV_NO = Null
        RSTTRXFILE!BANK_FLAG = "N"
'        RSTTRXFILE!B_TRX_TYPE = Null
'        'RSTTRXFILE!B_TRX_NO = Null
'        RSTTRXFILE!B_BILL_TRX_TYPE = Null
'        RSTTRXFILE!B_TRX_YEAR = Null
'        RSTTRXFILE!BANK_CODE = Null
    
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
    
    Screen.MousePointer = vbHourglass
    Sleep (300)
    ReportNameVar = Rptpath & "RptCustStatmnt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "') and ({DBTPYMT.TRX_TYPE} ='AA' OR {DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='RW' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} = 'EP' OR {DBTPYMT.TRX_TYPE} = 'VR' OR {DBTPYMT.TRX_TYPE} = 'ER' OR {DBTPYMT.TRX_TYPE} = 'PY' OR {DBTPYMT.TRX_TYPE} = 'RD') AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    'Report.RecordSelectionFormula = "({CUSTMAST.CR_FLAG}='Y')"
    'Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='SR')) "
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Statement of ' & '" & UCase(DataList2.text) & "' & ' for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdPrintDet_Click()

    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast As ADODB.Recordset

    If DataList2.BoundText = "" Then
        MsgBox "please Select Customer from the List", vbOKOnly, "Statement"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    Op_Bal = 0
    OP_Sale = 0
    OP_Rcpt = 0
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OP_Sale = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='RW' OR TRX_TYPE ='SR' OR TRX_TYPE ='DR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'RD') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        Select Case RSTTRXFILE!TRX_TYPE
            Case "DR", "RD"
                OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
            Case "DB"
                OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            Case Else
                OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
        End Select
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Op_Bal = OP_Sale - OP_Rcpt
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!OPEN_CR = Op_Bal
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim BAL_AMOUNT As Double
    BAL_AMOUNT = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='DR' OR TRX_TYPE ='RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'RD') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        'RSTTRXFILE!BAL_AMT = Op_Bal
        Select Case RSTTRXFILE!TRX_TYPE
            Case "DB"
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            Case "RD"
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
            Case Else
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + (IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                'RSTTRXFILE!BAL_AMT = Op_Bal

        End Select
        RSTTRXFILE!BAL_AMT = BAL_AMOUNT
        Op_Bal = 0
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    On Error GoTo ErrHand
    ReportNameVar = Rptpath & "RPTCustRep"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    For i = 1 To Report.OpenSubreport("RptCustStatmnt.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            If Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name = "TRXFILE" Or Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name = "TRXMAST" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name & " WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            ElseIf Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name = "itemmast" Then
                Set oRs = db.Execute("SELECT * FROM TRXFILE INNER JOIN " & Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name & " USING(ITEM_CODE) WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name & " ")
            End If
            
    '        If Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name = "TRXFILE" Then
    '            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name & " WHERE TRXFILE.VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND TRXFILE.VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "'")
    '        ElseIf Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name = "TRXMAST" Then
    '            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name & " WHERE TRXMAST.VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND TRXMAST.VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "'")
    '        Else
    '            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).Name & " ")
    '        End If
            'Report.OpenSubreport("RPTSALESDAY.rpt").Database.SetDataSource oRs, 3, i
            Report.OpenSubreport("RPTSALESDAY.rpt").Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RPTSALESDAY.rpt").RecordSelectionFormula = "(({TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='SI' OR {TRXMAST.TRX_TYPE}='RI' OR {TRXMAST.TRX_TYPE}='WO')  AND {TRXMAST.ACT_CODE}='" & DataList2.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    
    For i = 1 To Report.OpenSubreport("RptCustStatmnt.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RptCustStatmnt.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            If Report.OpenSubreport("RptCustStatmnt.rpt").Database.Tables(i).Name = "DBTPYMT" Then
                Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RptCustStatmnt.rpt").Database.Tables(i).Name & " INNER JOIN CUSTMAST USING (ACT_CODE) WHERE DBTPYMT.ACT_CODE ='" & DataList2.BoundText & "' AND DBTPYMT.INV_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND DBTPYMT.INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'")
            Else
                Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RptCustStatmnt.rpt").Database.Tables(i).Name & " INNER JOIN DBTPYMT USING (ACT_CODE) WHERE CUSTMAST.ACT_CODE ='" & DataList2.BoundText & "'")
            End If
            Report.OpenSubreport("RptCustStatmnt.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RptCustStatmnt.rpt").RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='RW' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} = 'EP' OR {DBTPYMT.TRX_TYPE} = 'VR' OR {DBTPYMT.TRX_TYPE} = 'ER' OR {DBTPYMT.TRX_TYPE} = 'PY' OR {DBTPYMT.TRX_TYPE} = 'RD') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    
    Report.OpenSubreport("RptCustStatmnt.rpt").DiscardSavedData
    Report.OpenSubreport("RptCustStatmnt.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RptCustStatmnt.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
         If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Statement of ' & '" & UCase(DataList2.text) & "' & ' for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    
    Report.OpenSubreport("RPTSALESDAY.rpt").DiscardSavedData
    Report.OpenSubreport("RPTSALESDAY.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTSALESDAY.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "Customer Ledger"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
    
    
End Sub

Private Sub CmdPrnRcpt_Click()
    'Call Sel_RCPTS
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
    
    Screen.MousePointer = vbHourglass
    Sleep (300)
    If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Receipt" Then
        ReportNameVar = Rptpath & "RptRcpt"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "({DBTPYMT.TRX_TYPE} ='RT' AND {DBTPYMT.CR_NO} = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND {DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "')"
    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Then
        ReportNameVar = Rptpath & "RptDN"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "({DBTPYMT.TRX_TYPE} ='DB' AND {DBTPYMT.CR_NO} = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND {DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "')"
    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note" Then
        ReportNameVar = Rptpath & "RptCN"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "({DBTPYMT.TRX_TYPE} ='CB' AND {DBTPYMT.CR_NO} = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND {DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "')"
    Else
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    'Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='SR')) "
    
    
    Dim i As Integer
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
    Set CRXFormulaFields = Report.FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.text = "'" & CompName & "'"
        If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.text = "'" & CompAddress1 & "'"
        If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.text = "'" & CompAddress2 & "'"
        If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.text = "'" & CompAddress3 & "'"
        If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
        If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Statement of ' & '" & UCase(DataList2.text) & "' & ' for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Cmdremove_Click()
    If MsgBox("Are you sure you want to make the selected invoices as Not Paid", vbYesNo + vbDefaultButton2, "Receipt Entry") = vbNo Then Exit Sub
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    On Error GoTo ErrHand
    For i = 0 To grdcount.rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT WHERE TRX_TYPE='DR' AND CR_NO = " & grdcount.TextMatrix(i, 7) & " AND INV_TRX_TYPE = '" & grdcount.TextMatrix(i, 8) & "' AND TRX_YEAR= '" & grdcount.TextMatrix(i, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE!PAID_FLAG = "N"
            RSTTRXFILE.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        GRDTranx.TextMatrix(grdcount.TextMatrix(i, 20), 22) = "PEND"
        
        If GRDTranx.TextMatrix(grdcount.TextMatrix(i, 20), 22) = "PAID" Then
            GRDTranx.Row = grdcount.TextMatrix(i, 20)
            GRDTranx.Col = 22
            GRDTranx.CellForeColor = vbBlue
        ElseIf GRDTranx.TextMatrix(grdcount.TextMatrix(i, 20), 22) = "PEND" Then
            GRDTranx.Row = grdcount.TextMatrix(i, 20)
            GRDTranx.Col = 22
            GRDTranx.CellForeColor = vbRed
        End If
        
    Next i
    
    For i = 1 To GRDTranx.rows - 1
        GRDTranx.TextMatrix(i, 21) = "N"
        With GRDTranx
            If .TextMatrix(.Row, 8) = "GI" Or .TextMatrix(.Row, 8) = "HI" Or .TextMatrix(.Row, 8) = "SI" Or .TextMatrix(.Row, 8) = "RI" Or .TextMatrix(.Row, 8) = "VI" Or .TextMatrix(.Row, 8) = "WO" Then
                .Row = i: .Col = 25: .CellPictureAlignment = 4 ' Align the checkbox
                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
            End If
        End With
    Next i
    
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Command1_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast, rstCustomer As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rstCustomer.EOF
        Op_Bal = 0
        OP_Sale = 0
        OP_Rcpt = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & rstCustomer!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(INV_AMT) from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and TRX_TYPE ='DR' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and TRX_TYPE ='DB' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE ='RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'RD') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
            
        
        Op_Bal = OP_Sale - OP_Rcpt
            
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & rstCustomer!ACT_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE!OPEN_CR = Op_Bal
            RSTTRXFILE.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Dim BAL_AMOUNT As Double
        Dim CR_FLAG As Boolean
        CR_FLAG = False
        BAL_AMOUNT = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='DR' OR TRX_TYPE ='RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'RD') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTTRXFILE.EOF
            'RSTTRXFILE!BAL_AMT = Op_Bal
            CR_FLAG = True
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                Case "RD"
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + (IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                    'RSTTRXFILE!BAL_AMT = Op_Bal
    
            End Select
            RSTTRXFILE!BAL_AMT = BAL_AMOUNT
            Op_Bal = 0
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        db.Execute "DELETE FROM DBTPYMT WHERE ACT_CODE = '" & rstCustomer!ACT_CODE & "' AND TRX_TYPE ='AA'"
        If CR_FLAG = False Then
            Dim MAXNO As Double
            MAXNO = 1
            Set RstCustmast = New ADODB.Recordset
            RstCustmast.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'AA'", db, adOpenForwardOnly
            If Not (RstCustmast.EOF And RstCustmast.BOF) Then
                MAXNO = IIf(IsNull(RstCustmast.Fields(0)), 1, RstCustmast.Fields(0) + 1)
            End If
            RstCustmast.Close
            Set RstCustmast = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * FROM DBTPYMT WHERE ACT_CODE = '" & rstCustomer!ACT_CODE & "' AND TRX_TYPE ='AA'", db, adOpenStatic, adLockOptimistic, adCmdText
            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE.AddNew
                RSTTRXFILE!TRX_TYPE = "AA"
                RSTTRXFILE!CR_NO = MAXNO
            End If
            RSTTRXFILE!INV_TRX_TYPE = ""
    '        RSTTRXFILE!RCPT_DATE = Null
    '        RSTTRXFILE!RCPT_AMT = Null
            RSTTRXFILE!ACT_CODE = rstCustomer!ACT_CODE
            RSTTRXFILE!ACT_NAME = rstCustomer!ACT_NAME
            RSTTRXFILE!INV_DATE = Format(DTFROM.Value, "DD/MM/YYYY")
            RSTTRXFILE!REF_NO = ""
    '        RSTTRXFILE!INV_AMT = Null
    '        RSTTRXFILE!INV_NO = Null
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    '        RSTTRXFILE!C_TRX_TYPE = Null
    '        'RSTTRXFILE!C_REC_NO = Null
    '        RSTTRXFILE!C_INV_TRX_TYPE = Null
    '        RSTTRXFILE!C_INV_TYPE = Null
    '        ''RSTTRXFILE!C_INV_NO = Null
            RSTTRXFILE!BANK_FLAG = "N"
    '        RSTTRXFILE!B_TRX_TYPE = Null
    '        'RSTTRXFILE!B_TRX_NO = Null
    '        RSTTRXFILE!B_BILL_TRX_TYPE = Null
    '        RSTTRXFILE!B_TRX_YEAR = Null
    '        RSTTRXFILE!BANK_CODE = Null
        
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        End If
        rstCustomer.MoveNext
    Loop
    rstCustomer.Close
    Set rstCustomer = Nothing
    
    Screen.MousePointer = vbHourglass
    Sleep (300)
    If ChkSmry.Value = 1 Then
        ReportNameVar = Rptpath & "RptCustStmntSmry"
    Else
        ReportNameVar = Rptpath & "RptCustStatmnt"
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "({DBTPYMT.TRX_TYPE} ='AA' OR {DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='RW' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} = 'EP' OR {DBTPYMT.TRX_TYPE} = 'VR' OR {DBTPYMT.TRX_TYPE} = 'ER' OR {DBTPYMT.TRX_TYPE} = 'PY' OR {DBTPYMT.TRX_TYPE} = 'RD') AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    'Report.RecordSelectionFormula = "({CUSTMAST.CR_FLAG}='Y')"
    'Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & RstCustomer!ACT_CODE & "' and ({DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='SR')) "
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Statement of ' & '" & UCase(DataList2.text) & "' & ' for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Command2_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, rstCustomer As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstCustomer.EOF
        Op_Bal = 0
        OP_Sale = 0
        OP_Rcpt = 0
        
        OP_Sale = IIf(IsNull(rstCustomer!OPEN_DB), 0, rstCustomer!OPEN_DB)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(INV_AMT) from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='DR' OR TRX_TYPE = 'RD') and INV_DATE < '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and TRX_TYPE ='DB' and INV_DATE < '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE ='RW' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') and INV_DATE < '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
            
        
        Op_Bal = OP_Sale - OP_Rcpt
        
        rstCustomer!OPEN_CR = Op_Bal
        rstCustomer.Update
        rstCustomer.MoveNext
        i = i + 1
        
    Loop
    rstCustomer.Close
    Set rstCustomer = Nothing
    
    Screen.MousePointer = vbHourglass
    Sleep (300)
    ReportNameVar = Rptpath & "RptCustStmntAll"
    
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "({CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.ACT_CODE} <> '130001')"
    'Report.RecordSelectionFormula = "((Mid({ACTMAST.ACT_CODE}, 1, 3)='311')And (LENGTH({ACTMAST.ACT_CODE})>3))"
    'Report.RecordSelectionFormula = "({CUSTMAST.CR_FLAG}='Y')"
    'Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & RstCustomer!ACT_CODE & "' and ({DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='SR')) "
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Statement of ' & '" & UCase(DataList2.text) & "' & ' for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Activate()
    Call Fillgrid
End Sub

Private Sub Form_Load()
    
'    db.Execute "Update DBTPYMT set RCPT_AMT =0 where isnull(RCPT_AMT)"
'    db.Execute "Update DBTPYMT set INV_AMT =0 where isnull(INV_AMT)"
    
    GRDTranx.TextMatrix(0, 0) = "TYPE"
    GRDTranx.TextMatrix(0, 1) = "SL"
    GRDTranx.TextMatrix(0, 2) = "INV /PAID DATE"
    GRDTranx.TextMatrix(0, 3) = "INV NO"
    GRDTranx.TextMatrix(0, 4) = "INV AMT"
    GRDTranx.TextMatrix(0, 5) = "RCVD AMT"
    GRDTranx.TextMatrix(0, 6) = "REF NO"
    GRDTranx.TextMatrix(0, 7) = "" '"CR NO"
    GRDTranx.TextMatrix(0, 8) = "" '"TYPE"
    GRDTranx.TextMatrix(0, 20) = "Entry Date"
    GRDTranx.TextMatrix(0, 22) = "Status"
    GRDTranx.TextMatrix(0, 23) = "Rcvd Amt"
    GRDTranx.TextMatrix(0, 24) = "Bal Amt"
    GRDTranx.TextMatrix(0, 26) = "Bank Name"
    GRDTranx.TextMatrix(0, 27) = "Days"
    GRDTranx.TextMatrix(0, 28) = "Comm"
    
    GRDTranx.ColWidth(0) = 900
    GRDTranx.ColWidth(1) = 700
    GRDTranx.ColWidth(2) = 1500
    GRDTranx.ColWidth(3) = 1200
    GRDTranx.ColWidth(4) = 1200
    GRDTranx.ColWidth(5) = 1200
    GRDTranx.ColWidth(6) = 2400
    GRDTranx.ColWidth(7) = 0
    GRDTranx.ColWidth(8) = 0
    GRDTranx.ColWidth(9) = 0
    GRDTranx.ColWidth(10) = 0
    GRDTranx.ColWidth(11) = 0
    GRDTranx.ColWidth(12) = 0
    GRDTranx.ColWidth(13) = 0
    GRDTranx.ColWidth(14) = 0
    GRDTranx.ColWidth(15) = 0
    GRDTranx.ColWidth(16) = 0
    GRDTranx.ColWidth(17) = 0
    GRDTranx.ColWidth(18) = 0
    GRDTranx.ColWidth(19) = 0
    GRDTranx.ColWidth(20) = 1100
    GRDTranx.ColWidth(21) = 0
    GRDTranx.ColWidth(22) = 800
    GRDTranx.ColWidth(23) = 1200
    GRDTranx.ColWidth(24) = 1200
    GRDTranx.ColWidth(25) = 280
    GRDTranx.ColWidth(26) = 2000
    GRDTranx.ColWidth(27) = 800
    GRDTranx.ColWidth(28) = 1000
    GRDTranx.ColWidth(29) = 1900
    GRDTranx.ColWidth(30) = 0
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 4
    GRDTranx.ColAlignment(3) = 4
    'GRDTranx.ColAlignment(4) = 4
    'GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 1
    GRDTranx.ColAlignment(26) = 1
    GRDTranx.ColAlignment(27) = 4
    'GRDTranx.ColAlignment(28) = 1
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "Rate"
    GRDBILL.TextMatrix(0, 3) = "Disc %"
    GRDBILL.TextMatrix(0, 4) = "Tax %"
    GRDBILL.TextMatrix(0, 5) = "Qty"
    GRDBILL.TextMatrix(0, 6) = "Amount"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 5500
    GRDBILL.ColWidth(2) = 800
    GRDBILL.ColWidth(3) = 800
    GRDBILL.ColWidth(4) = 800
    GRDBILL.ColWidth(5) = 900
    GRDBILL.ColWidth(6) = 1100
    
    GRDBILL.ColAlignment(0) = 4
    GRDBILL.ColAlignment(1) = 1
    GRDBILL.ColAlignment(2) = 4
    GRDBILL.ColAlignment(3) = 4
    GRDBILL.ColAlignment(4) = 4
    GRDBILL.ColAlignment(5) = 4
    
    grdreceipts.TextMatrix(0, 0) = "SL"
    grdreceipts.TextMatrix(0, 1) = "Date"
    grdreceipts.TextMatrix(0, 2) = "Amount"
    grdreceipts.TextMatrix(0, 3) = """"
    grdreceipts.TextMatrix(0, 4) = "" '"%"
    grdreceipts.TextMatrix(0, 5) = ""
    grdreceipts.TextMatrix(0, 6) = "Remarks"
    
    grdreceipts.ColWidth(0) = 500
    grdreceipts.ColWidth(1) = 1500
    grdreceipts.ColWidth(2) = 1500
    grdreceipts.ColWidth(3) = 0
    grdreceipts.ColWidth(4) = 0
    grdreceipts.ColWidth(5) = 0
    grdreceipts.ColWidth(6) = 1800
    
    grdreceipts.ColAlignment(0) = 4
    grdreceipts.ColAlignment(1) = 1
    grdreceipts.ColAlignment(2) = 4
    grdreceipts.ColAlignment(3) = 4
    grdreceipts.ColAlignment(4) = 4
    grdreceipts.ColAlignment(5) = 4
    grdreceipts.ColAlignment(6) = 1
    
    'GRDBILL.ColAlignment(6) = 4
    
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 200
    Top = 0
    TXTDEALER.text = " "
    TXTDEALER.text = ""
    
    If Month(Date) - 2 > 4 Then
        DTFROM.Value = "01/" & Format(Month(Date) - 2, "00") & "/" & Year(Date)
    Else
        If Year(Date) > Year(MDIMAIN.DTFROM.Value) Then
            DTFROM.Value = "01/12/" & Year(MDIMAIN.DTFROM.Value)
        Else
            DTFROM.Value = "01/04/" & Year(MDIMAIN.DTFROM.Value)
        End If
    End If
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    
    'MDIMAIN.MNUPYMNT.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close

    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.MNUPYMNT.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRMEMAIN.Enabled = True
            FRMEBILL.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDBILL_LostFocus()
    If FRMEBILL.Visible = True Then
        Frmeperiod.Enabled = True
        FRMEBILL.Visible = False
        GRDTranx.SetFocus
    End If
End Sub

Private Sub GRDTranx_Click()
    FRMEBILL.Visible = False
    Frmereceipt.Visible = False
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    If GRDTranx.rows = 1 Then Exit Sub
    If GRDTranx.Col <> 25 Then Exit Sub
    With GRDTranx
        If .TextMatrix(.Row, 0) = "Sale" Then
            oldx = .Col
            oldy = .Row
            .Row = oldy: .Col = 25: .CellPictureAlignment = 4
                'If GRDTranx.Col = 0 Then
                    If GRDTranx.CellPicture = picChecked Then
                        Set GRDTranx.CellPicture = picUnchecked
                        '.Col = .Col + 2  ' I use data that is in column #1, usually an Index or ID #
                        'strTextCheck = .Text
                        ' When you de-select a CheckBox, we need to strip out the #
                        'strChecked = strChecked & strTextCheck & ","
                        ' Don't forget to strip off the trailing , before passing the string
                        'Debug.Print strChecked
                        .TextMatrix(.Row, 21) = "Y"
                        Call fillcount
                    Else
                        Set GRDTranx.CellPicture = picChecked
                        '.Col = .Col + 2
                        'strTextCheck = .Text
                        'strChecked = Replace(strChecked, strTextCheck & ",", "")
                        'Debug.Print strChecked
                        .TextMatrix(.Row, 21) = "N"
                        Call fillcount
                    End If
                'End If
            .Col = oldx
            .Row = oldy
        End If
    End With
    
    TXTEXPIRY.Visible = False
    TXTsample.Visible = False
    TXTsample2.Visible = False
End Sub

Private Sub GRDTranx_DblClick()
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    
    If GRDTranx.TextMatrix(GRDTranx.Row, 0) <> "Sale" Then Exit Sub
    Select Case Trim(GRDTranx.TextMatrix(GRDTranx.Row, 8))
        Case "HI"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDTranx.TextMatrix(GRDTranx.Row, 14)) Then Exit Sub
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                If IsFormLoaded(frmsales) <> True Then
                    frmsales.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    frmsales.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    frmsales.Show
                    frmsales.SetFocus
                    Call frmsales.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMSALES1) <> True Then
                    FRMSALES1.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMSALES1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMSALES1.Show
                    FRMSALES1.SetFocus
                    Call FRMSALES1.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMSALES2) <> True Then
                    FRMSALES2.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMSALES2.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMSALES2.Show
                    FRMSALES2.SetFocus
                    Call FRMSALES2.txtBillNo_KeyDown(13, 0)
                End If
            Else
                If SALESLT_FLAG = "Y" Then
                    If IsFormLoaded(FRMGSTRSM1) <> True Then
                        FRMGSTRSM1.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM1.Show
                        FRMGSTRSM1.SetFocus
                        Call FRMGSTRSM1.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTRSM2) <> True Then
                        FRMGSTRSM2.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM2.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM2.Show
                        FRMGSTRSM2.SetFocus
                        Call FRMGSTRSM2.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTRSM3) <> True Then
                        FRMGSTRSM3.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM3.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTRSM3.Show
                        FRMGSTRSM3.SetFocus
                        Call FRMGSTRSM3.txtBillNo_KeyDown(13, 0)
                    End If
                Else
                    If IsFormLoaded(FRMGSTR) <> True Then
                        FRMGSTR.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR.Show
                        FRMGSTR.SetFocus
                        Call FRMGSTR.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTR1) <> True Then
                        FRMGSTR1.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR1.Show
                        FRMGSTR1.SetFocus
                        Call FRMGSTR1.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTR2) <> True Then
                        FRMGSTR2.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR2.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                        FRMGSTR2.Show
                        FRMGSTR2.SetFocus
                        Call FRMGSTR2.txtBillNo_KeyDown(13, 0)
                    End If
                End If
            End If
        Case "GI"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDTranx.TextMatrix(GRDTranx.Row, 14)) Then Exit Sub
            If IsFormLoaded(FRMGST) <> True Then
                FRMGST.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                FRMGST.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                FRMGST.Show
                FRMGST.SetFocus
                Call FRMGST.txtBillNo_KeyDown(13, 0)
            ElseIf IsFormLoaded(FRMGST1) <> True Then
                FRMGST1.txtBillNo.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                FRMGST1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                FRMGST1.Show
                FRMGST1.SetFocus
                Call FRMGST1.txtBillNo_KeyDown(13, 0)
            End If
        Case "WO"
    End Select
    Exit Sub
ErrHand:
        MsgBox err.Description, , "EzBiz"
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.rows = 1 Then Exit Sub
            If GRDTranx.TextMatrix(GRDTranx.Row, 0) <> "Sale" Then Exit Sub
            Select Case GRDTranx.Col
                Case 23
                    grdreceipts.rows = 1
                    i = 0
                    
'                        Set rstTRANX2 = New ADODB.Recordset
'                        rstTRANX2.Open "select SUM(RCPT_AMOUNT) from trnxrcpt WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND INV_NO  = " & rstTRANX!INV_NO & " AND INV_TRX_TYPE = '" & rstTRANX!INV_TRX_TYPE & "' AND INV_TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                        If Not (rstTRANX2.EOF And rstTRANX2.BOF) Then
'                            rstTRANX!RCVD_AMOUNT = IIf(IsNull(rstTRANX2.Fields(0)), 0, rstTRANX2.Fields(0))
'                            rstTRANX.Update
'                            'db.Execute "Update DBTPYMT set RCVD_AMOUNT = IIf(IsNull(rstTRANX2.Fields(0)), 0, rstTRANX2.Fields(0)) where ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND TRX_TYPE = 'DR' AND INV_TRX_TYPE  = '" & rstTRANX!TRX_TYPE & "' AND INV_NO = '" & rstTRANX!VCH_NO & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "'"
'                            'lblsaleret.Caption = Format(IIf(IsNull(rstTRANX2.Fields(0)), 0, rstTRANX2.Fields(0)), "0.00")
'                        End If
'                        rstTRANX2.Close
'                        Set rstTRANX2 = Nothing
    
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "Select * From trnxrcpt WHERE TRX_TYPE = 'RT' AND ACT_CODE = '" & DataList2.BoundText & "' AND INV_TRX_YEAR = '" & Val(GRDTranx.TextMatrix(GRDTranx.Row, 14)) & "' AND INV_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 3)) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "'", db, adOpenForwardOnly
                    Do Until RSTTRXFILE.EOF
                        i = i + 1
                        grdreceipts.rows = grdreceipts.rows + 1
                        grdreceipts.FixedRows = 1
                        grdreceipts.TextMatrix(i, 0) = i
                        grdreceipts.TextMatrix(i, 1) = Format(RSTTRXFILE!RCPT_DATE, "dd/mm/yyyy")
                        grdreceipts.TextMatrix(i, 2) = IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), "0.00", Format(RSTTRXFILE!RCPT_AMOUNT, "0.00"))
                        grdreceipts.TextMatrix(i, 3) = RSTTRXFILE!RCPT_NO
                        grdreceipts.TextMatrix(i, 4) = RSTTRXFILE!TRX_TYPE
                        grdreceipts.TextMatrix(i, 5) = RSTTRXFILE!INV_NO
                        grdreceipts.TextMatrix(i, 6) = IIf(IsNull(RSTTRXFILE!REMARKS), "", RSTTRXFILE!REMARKS)
                        RSTTRXFILE.MoveNext
                    Loop
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
        
                    Frmereceipt.Visible = True
                    grdreceipts.SetFocus
                    Exit Sub
            End Select
                
            
            
            
'            If GRDTranx.Rows = 1 Then Exit Sub
'            If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Receipt" Then
'                Select Case GRDTranx.Col
'                    Case 5
'                            'If GRDTranx.Cols = 20 Then Exit Sub
'                            TXTsample.MaxLength = 7
'                            TXTsample.Visible = True
'                            TXTsample.Top = GRDTranx.CellTop + 100
'                            TXTsample.Left = GRDTranx.CellLeft '+ 50
'                            TXTsample.Width = GRDTranx.CellWidth
'                            TXTsample.Height = GRDTranx.CellHeight
'                            TXTsample.Text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
'                            TXTsample.SetFocus
'                End Select
'                Exit Sub
'            End If
            LBLSUPPLIER.Caption = " " & DataList2.text
            LBLINVDATE.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 2)
            LBLBILLNO.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 3)
            LBLBILLAMT.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'LBLPAID.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'LBLBAL.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 6)

            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & Val(GRDTranx.TextMatrix(GRDTranx.Row, 14)) & "' AND VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "'", db, adOpenForwardOnly
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.rows = GRDBILL.rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ITEM_NAME
                GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!SALES_PRICE, "0.00")
                GRDBILL.TextMatrix(i, 3) = Val(RSTTRXFILE!LINE_DISC)
                GRDBILL.TextMatrix(i, 4) = Val(RSTTRXFILE!SALES_TAX)
                GRDBILL.TextMatrix(i, 5) = RSTTRXFILE!QTY
                GRDBILL.TextMatrix(i, 6) = Format(RSTTRXFILE!TRX_TOTAL, "0.00")
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
                            
            lblqty.Caption = ""
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT sum(QTY * LOOSE_PACK) FROM TRXFILE WHERE TRX_YEAR = '" & Val(GRDTranx.TextMatrix(GRDTranx.Row, 14)) & "' AND VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "'", db, adOpenForwardOnly
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                lblqty.Caption = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        
            FRMEBILL.Visible = True
            GRDBILL.SetFocus
        Case vbKeyF6
            If DataList2.BoundText = "" Then Exit Sub
            If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then Exit Sub
            'If GRDTranx.Rows <= 1 Then Exit Sub
            Enabled = False
            FRMRECEIPT.LBLSUPPLIER.Caption = DataList2.text
            FRMRECEIPT.lblactcode.Caption = DataList2.BoundText
            'FRMRECEIPTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMRECEIPTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMRECEIPTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMRECEIPTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMRECEIPTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMRECEIPT.Show
        Case vbKeyF7
            If DataList2.BoundText = "" Then Exit Sub
            If DataList2.BoundText = "130000" Or DataList2.BoundText = "130001" Then Exit Sub
            'If GRDTranx.Rows <= 1 Then Exit Sub
            Enabled = False
            FRMDRCR.LBLSUPPLIER.Caption = DataList2.text
            FRMDRCR.lblactcode.Caption = DataList2.BoundText
            'FRMRECEIPTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMRECEIPTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMRECEIPTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMRECEIPTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMRECEIPTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMDRCR.Show
        Case vbKeyF2
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
            Select Case GRDTranx.Col
                 Case 2
                    If Not (GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Receipt" Or GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note") Then Exit Sub
                    TXTEXPIRY.Visible = True
                    TXTEXPIRY.Top = GRDTranx.CellTop + 120
                    TXTEXPIRY.Left = GRDTranx.CellLeft + 20
                    TXTEXPIRY.Width = GRDTranx.CellWidth
                    TXTEXPIRY.Height = GRDTranx.CellHeight
                    If Not (IsDate(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))) Then
                        TXTEXPIRY.text = "  /  /    "
                    Else
                        TXTEXPIRY.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                    End If
                    TXTEXPIRY.SetFocus
                Case 5
                    If Not (GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Receipt" Or GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note") Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDTranx.CellTop + 120
                    TXTsample.Left = GRDTranx.CellLeft + 20
                    TXTsample.Width = GRDTranx.CellWidth
                    TXTsample.Height = GRDTranx.CellHeight
                    TXTsample.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))
                    TXTsample.SetFocus
                Case 4
                    If Not (GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note") Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDTranx.CellTop + 120
                    TXTsample.Left = GRDTranx.CellLeft + 20
                    TXTsample.Width = GRDTranx.CellWidth
                    TXTsample.Height = GRDTranx.CellHeight
                    TXTsample.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))
                    TXTsample.SetFocus
                Case 6
                    TXTsample.Visible = True
                    TXTsample.Top = GRDTranx.CellTop + 120
                    TXTsample.Left = GRDTranx.CellLeft + 20
                    TXTsample.Width = GRDTranx.CellWidth
                    TXTsample.Height = GRDTranx.CellHeight
                    TXTsample.text = Trim(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))
                    TXTsample.SetFocus
            End Select
    End Select
End Sub

Private Sub GRDTranx_KeyPress(KeyAscii As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    On Error GoTo ErrHand
    Select Case KeyAscii
        Case vbKeyD, Asc("d")
            CMDDISPLAY.Tag = KeyAscii
        Case vbKeyE, Asc("e")
            CMDEXIT.Tag = KeyAscii
        Case vbKeyL, Asc("l")
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
            If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Receipt" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From DBTPYMT WHERE TRX_TYPE='RT' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        db.Execute "UPDATE DBTPYMT SET RCVD_AMOUNT = 0 WHERE TRX_TYPE = 'DR' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 3) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        'db.Execute "delete From TRNXRCPT WHERE TRX_TYPE = 'RT' AND CR_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 7)) & " AND RCPT_AMOUNT = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 5)) & " "
                        db.Execute "delete From TRNXRCPT WHERE TRX_TYPE = 'RT' AND CR_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 7)) & " AND CR_TRX_TYPE ='DR'"
                        
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 15) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 16) = "") Then
                            db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 16) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 19) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        End If
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 9) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 10) = "") Then
                            db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 14) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 8) = "") Then
                            Set RSTTRXFILE = New ADODB.Recordset
                            RSTTRXFILE.Open "Select * FROM TRXMAST WHERE TRX_YEAR= '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND  TRX_TYPE= '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND VCH_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 3) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                                RSTTRXFILE!RCPT_AMOUNT = 0
                                RSTTRXFILE!RCPT_REFNO = ""
                                RSTTRXFILE!BANK_FLAG = "N"
                                'RSTTRXFILE!CHQ_NO = Null
                                'RSTTRXFILE!BANK_CODE = Null
                                'RSTTRXFILE!BANK_NAME = Null
                                'RSTTRXFILE!CHQ_DATE = Null
                                RSTTRXFILE!CHQ_STATUS = "N"
                                RSTTRXFILE!POST_FLAG = "N"
                                RSTTRXFILE.Update
                            End If
                            RSTTRXFILE.Close
                            Set RSTTRXFILE = Nothing
                        End If
                        db.CommitTrans
                        Call Fillgrid
                    Else
                        GRDTranx.SetFocus
                    End If
                End If
'                If OptCr.value = True Then
'                    RSTTRXFILE!TRX_TYPE = "CR"
'                    RSTTRXFILE!BILL_TRX_TYPE = "CN"
'                Else
'                    RSTTRXFILE!TRX_TYPE = "DR"
'                    RSTTRXFILE!BILL_TRX_TYPE = "DN"
'                End If
            ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From DBTPYMT WHERE TRX_TYPE='DB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'DN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 15) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 16) = "") Then
                            db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 16) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 19) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        End If
                        'db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'DN' AND INV_TRX_TYPE = 'DN' AND TRX_TYPE = 'CR'"
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = 'CR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'CN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        db.CommitTrans
                        Call Fillgrid
                    End If
                End If
            ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From DBTPYMT WHERE TRX_TYPE='CB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'CN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 15) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 16) = "") Then
                            db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 16) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 19) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        End If
                        'db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'CN' AND INV_TRX_TYPE = 'CN' AND TRX_TYPE = 'CR'"
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = 'DR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'DN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        db.Execute "delete From TRNXRCPT WHERE TRX_TYPE = 'RT' AND CR_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 7)) & " AND CR_TRX_TYPE ='DR'"
                        db.CommitTrans
                        Call Fillgrid
                    End If
                End If
            ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Cheque Return" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From DBTPYMT WHERE TRX_TYPE='RD' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                        'db.Execute "delete From trnxrcpt WHERE TRX_TYPE='PY' AND CR_TRX_TYPE= 'CR' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        db.CommitTrans
                        Call Fillgrid
                    Else
                        GRDTranx.SetFocus
                    End If
                End If
                        
            End If
            
        Case Else
            CMDEXIT.Tag = ""
            CMDDISPLAY.Tag = ""
    End Select
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub GRDTranx_Scroll()
    
    TXTEXPIRY.Visible = False
    TXTsample.Visible = False
    TXTsample2.Visible = False
End Sub

Private Sub TXTCODE_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE (ACT_CODE <> '130000' or ACT_CODE <> '130001') And ACT_CODE Like '" & Me.TxtCode.text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE (ACT_CODE <> '130000' or ACT_CODE <> '130001') And ACT_CODE Like '" & Me.TxtCode.text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
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
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            
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

Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ACT_NAME
        End If
        Set DataList2.RowSource = ACT_REC
        DataList2.ListField = "ACT_NAME"
        DataList2.BoundColumn = "ACT_CODE"
    End If
    Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    GRDTranx.rows = 1
    TxtCode.text = DataList2.BoundText
    CmDDisplay_Click
    DataList2.SetFocus
    flagchange.Caption = ""
    lbldealer.Caption = DataList2.text
    
    On Error GoTo ErrHand
    Dim rstCustomer As ADODB.Recordset
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
    Else
        lbladdress.Caption = ""
    End If
        
    'TXTDEALER.Text = lbldealer.Caption
    'LBL.Caption = ""
    Exit Sub
ErrHand:
    
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
           
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
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Function Fillgrid()
    Dim rstTRANX As ADODB.Recordset
    Dim rstTRANX2 As ADODB.Recordset
    Dim RSTBANK As ADODB.Recordset
    Dim i As Long
    
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo ErrHand
        
    Screen.MousePointer = vbHourglass
        
    If chkverify.Value = 1 Then
        Dim rstdbt As ADODB.Recordset
        Dim rstdbt2 As ADODB.Recordset
        Set rstdbt = New ADODB.Recordset
        rstdbt.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' and TRX_TYPE = 'DR' and INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until rstdbt.EOF
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(RCPT_AMOUNT) from trnxrcpt WHERE ACT_CODE = '" & rstdbt!ACT_CODE & "' AND INV_NO  = " & rstdbt!INV_NO & " AND INV_TRX_TYPE = '" & rstdbt!INV_TRX_TYPE & "' AND INV_TRX_YEAR = '" & rstdbt!TRX_YEAR & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
                rstdbt!RCVD_AMOUNT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
                rstdbt.Update
                'db.Execute "Update DBTPYMT set RCVD_AMOUNT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0)) where ACT_CODE = '" & rstdbt!ACT_CODE & "' AND TRX_TYPE = 'DR' AND INV_TRX_TYPE  = '" & rstdbt!TRX_TYPE & "' AND INV_NO = '" & rstdbt!VCH_NO & "' AND TRX_YEAR = '" & rstdbt!TRX_YEAR & "'"
                'lblsaleret.Caption = Format(IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0)), "0.00")
            End If
            rstdbt2.Close
            Set rstdbt2 = Nothing
            rstdbt.MoveNext
        Loop
        rstdbt.Close
        Set rstdbt = Nothing
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' and TRX_TYPE = 'DR' and INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until rstTRANX.EOF
                
            Set rstTRANX2 = New ADODB.Recordset
            rstTRANX2.Open "select SUM(RCPT_AMOUNT) from trnxrcpt WHERE TRX_TYPE = 'RT' AND ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND INV_NO  = " & rstTRANX!INV_NO & " AND INV_TRX_TYPE = '" & rstTRANX!INV_TRX_TYPE & "' AND INV_TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstTRANX2.EOF And rstTRANX2.BOF) Then
                rstTRANX!RCVD_AMOUNT = IIf(IsNull(rstTRANX2.Fields(0)), 0, rstTRANX2.Fields(0))
                rstTRANX.Update
                'db.Execute "Update DBTPYMT set RCVD_AMOUNT = IIf(IsNull(rstTRANX2.Fields(0)), 0, rstTRANX2.Fields(0)) where ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND TRX_TYPE = 'DR' AND INV_TRX_TYPE  = '" & rstTRANX!TRX_TYPE & "' AND INV_NO = '" & rstTRANX!VCH_NO & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "'"
                'lblsaleret.Caption = Format(IIf(IsNull(rstTRANX2.Fields(0)), 0, rstTRANX2.Fields(0)), "0.00")
            End If
            rstTRANX2.Close
            Set rstTRANX2 = Nothing
                
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
    End If
    
    
    GRDTranx.rows = 1
    LBLINVAMT.Caption = ""
    LBLPAIDAMT.Caption = ""
    lblcomm.Caption = ""
    lblcommpaid.Caption = ""
    lblcommpend.Caption = ""
    LBLBALAMT.Caption = ""
    lblOPBal.Caption = ""
    Dim m_Rcpt_Amt As Double
    Dim OP_Sale As Double
    Dim OP_Rcpt As Double
    m_Rcpt_Amt = 0
    OP_Sale = 0
    OP_Rcpt = 0
    i = 1
          
              
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblOPBal.Caption = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
        If rstTRANX!CUST_TYPE = "D" Then
            lblagent.Caption = "Y"
        Else
            lblagent.Caption = "N"
        End If
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    
'    Set rstTRANX = New ADODB.Recordset
'    rstTRANX.Open "select SUM(INV_AMT) from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and TRX_TYPE ='DR' and INV_DATE < '" & Format(DTFROM.value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
'        OP_Sale = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
'    End If
'    rstTRANX.Close
'    Set rstTRANX = Nothing
'
'    Set rstTRANX = New ADODB.Recordset
'    rstTRANX.Open "select SUM(RCPT_AMT) from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and TRX_TYPE ='DB' and INV_DATE < '" & Format(DTFROM.value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
'        OP_Sale = OP_Sale + IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
'    End If
'    rstTRANX.Close
'    Set rstTRANX = Nothing
'
'    Set rstTRANX = New ADODB.Recordset
'    rstTRANX.Open "select SUM(INV_AMT) from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') and INV_DATE < '" & Format(DTFROM.value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
'        OP_Rcpt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
'    End If
'    rstTRANX.Close
'    Set rstTRANX = Nothing
                    
                    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='SR' OR TRX_TYPE ='RW' OR TRX_TYPE ='DR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'RD') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rstTRANX.EOF
        Select Case rstTRANX!TRX_TYPE
            Case "DR", "RD"
                OP_Sale = OP_Sale + IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
            Case "DB"
                OP_Sale = OP_Sale + IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
            Case Else
                OP_Rcpt = OP_Rcpt + IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
        End Select
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    lblOPBal.Caption = Format(Round(Val(lblOPBal.Caption) + (OP_Sale - OP_Rcpt), 2), "0.00")
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'CB' OR TRX_TYPE = 'DB' OR TRX_TYPE = 'RT' OR TRX_TYPE = 'DR' OR TRX_TYPE = 'SR' OR TRX_TYPE ='RW' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'RD') and INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, CR_NO", db, adOpenForwardOnly
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0

        GRDTranx.TextMatrix(i, 1) = i
        GRDTranx.TextMatrix(i, 2) = Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!INV_AMT, "0.00")
        GRDTranx.TextMatrix(i, 28) = Format(rstTRANX!COMM_AMT, "0.00")
        GRDTranx.TextMatrix(i, 29) = IIf(IsNull(rstTRANX!BR_ADDRESS) Or rstTRANX!BR_ADDRESS = "", "", rstTRANX!BR_ADDRESS)
        Select Case rstTRANX!check_flag
            Case "Y"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!INV_AMT, "0.00")
            Case "N"
                GRDTranx.TextMatrix(i, 5) = "" '0 '""Format(rstTRANX!INV_AMT, "0.00")
        End Select
        Select Case rstTRANX!TRX_TYPE
            Case "DR"
                GRDTranx.TextMatrix(i, 0) = "Sale"
                GRDTranx.CellForeColor = vbRed
                GRDTranx.TextMatrix(i, 23) = IIf(IsNull(rstTRANX!RCVD_AMOUNT), "", Format(rstTRANX!RCVD_AMOUNT, "0.00"))
                GRDTranx.TextMatrix(i, 24) = Format(GRDTranx.TextMatrix(i, 4) - Val(GRDTranx.TextMatrix(i, 23)), "0.00")
                GRDTranx.TextMatrix(i, 27) = DateDiff("d", rstTRANX!INV_DATE, Date)
                Select Case rstTRANX!PAID_FLAG
                    Case "Y"
                        GRDTranx.TextMatrix(i, 22) = "PAID"
                        GRDTranx.TextMatrix(i, 27) = ""
                    Case Else
                        GRDTranx.TextMatrix(i, 22) = "PEND"
                End Select
                
            Case "DB"
                GRDTranx.TextMatrix(i, 0) = "Debit Note"
                GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbRed
            Case "RD"
                GRDTranx.TextMatrix(i, 0) = "Cheque Return"
                GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!INV_AMT, "0.00")
                GRDTranx.CellForeColor = vbRed
            Case "CB"
                GRDTranx.TextMatrix(i, 0) = "Credit Note"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                lblcommpaid.Caption = Format(Val(lblcommpaid.Caption) + Val(GRDTranx.TextMatrix(i, 5)), "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "PY"
                GRDTranx.TextMatrix(i, 0) = "Purchase"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "RT"
                GRDTranx.TextMatrix(i, 0) = "Receipt"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "SR"
                GRDTranx.TextMatrix(i, 0) = "SALES RETURN"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "RW"
                GRDTranx.TextMatrix(i, 0) = "SALES RETURN(W)"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "EP", "VC", "ER"
                GRDTranx.TextMatrix(i, 0) = "EXPIRY RETURN"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbBlue
        End Select
        GRDTranx.TextMatrix(i, 30) = IIf(IsNull(rstTRANX!TRX_TYPE), "", rstTRANX!TRX_TYPE)
        GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO) & IIf(IsNull(rstTRANX!CHQ_NO) Or rstTRANX!CHQ_NO = "", "", " (" & rstTRANX!CHQ_NO & ")")
        GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!CR_NO), "", rstTRANX!CR_NO)
        GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstTRANX!INV_TRX_TYPE), "", rstTRANX!INV_TRX_TYPE)
        Select Case rstTRANX!BANK_FLAG
            Case "Y"
                GRDTranx.TextMatrix(i, 9) = IIf(IsNull(rstTRANX!B_TRX_TYPE), "", rstTRANX!B_TRX_TYPE)
                GRDTranx.TextMatrix(i, 10) = IIf(IsNull(rstTRANX!B_TRX_NO), "", rstTRANX!B_TRX_NO)
                GRDTranx.TextMatrix(i, 11) = IIf(IsNull(rstTRANX!B_BILL_TRX_TYPE), "", rstTRANX!B_BILL_TRX_TYPE)
                GRDTranx.TextMatrix(i, 12) = IIf(IsNull(rstTRANX!B_TRX_YEAR), "", rstTRANX!B_TRX_YEAR)
                GRDTranx.TextMatrix(i, 13) = IIf(IsNull(rstTRANX!BANK_CODE), "", rstTRANX!BANK_CODE)
                
                Set RSTBANK = New ADODB.Recordset
                RSTBANK.Open "select * from BANKCODE  WHERE BANK_CODE = '" & GRDTranx.TextMatrix(i, 13) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTBANK.EOF And RSTBANK.BOF) Then
                    GRDTranx.TextMatrix(i, 26) = IIf(IsNull(RSTBANK!BANK_NAME), "", RSTBANK!BANK_NAME)
                End If
                RSTBANK.Close
                Set RSTBANK = Nothing
                
                GRDTranx.TextMatrix(i, 15) = ""
                GRDTranx.TextMatrix(i, 16) = ""
                GRDTranx.TextMatrix(i, 17) = ""
                GRDTranx.TextMatrix(i, 18) = ""
                GRDTranx.TextMatrix(i, 19) = ""
            Case Else
                GRDTranx.TextMatrix(i, 9) = ""
                GRDTranx.TextMatrix(i, 10) = ""
                GRDTranx.TextMatrix(i, 11) = ""
                GRDTranx.TextMatrix(i, 12) = ""
                GRDTranx.TextMatrix(i, 13) = ""
                GRDTranx.TextMatrix(i, 15) = IIf(IsNull(rstTRANX!C_TRX_TYPE), "", rstTRANX!C_TRX_TYPE)
                GRDTranx.TextMatrix(i, 16) = IIf(IsNull(rstTRANX!C_REC_NO), "", rstTRANX!C_REC_NO)
                GRDTranx.TextMatrix(i, 17) = IIf(IsNull(rstTRANX!C_INV_TRX_TYPE), "", rstTRANX!C_INV_TRX_TYPE)
                GRDTranx.TextMatrix(i, 18) = IIf(IsNull(rstTRANX!C_INV_TYPE), "", rstTRANX!C_INV_TYPE)
                GRDTranx.TextMatrix(i, 19) = IIf(IsNull(rstTRANX!C_INV_NO), "", rstTRANX!C_INV_NO)
        End Select
        GRDTranx.TextMatrix(i, 20) = IIf(IsNull(rstTRANX!ENTRY_DATE), "", rstTRANX!ENTRY_DATE)
        GRDTranx.TextMatrix(i, 14) = IIf(IsNull(rstTRANX!TRX_YEAR), "", rstTRANX!TRX_YEAR)
        GRDTranx.TextMatrix(i, 21) = "N"
        
        With GRDTranx
            If .TextMatrix(.Row, 8) = "GI" Or .TextMatrix(.Row, 8) = "HI" Or .TextMatrix(.Row, 8) = "SI" Or .TextMatrix(.Row, 8) = "RI" Or .TextMatrix(.Row, 8) = "VI" Or .TextMatrix(.Row, 8) = "WO" Then
                .Row = i: .Col = 25: .CellPictureAlignment = 4 ' Align the checkbox
                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
            End If
        End With
        
        If GRDTranx.TextMatrix(i, 0) = "Sale" And Val(GRDTranx.TextMatrix(i, 24)) <= 0 Then
            GRDTranx.TextMatrix(i, 22) = "PAID"
            GRDTranx.TextMatrix(i, 27) = ""
        End If
        
        If GRDTranx.TextMatrix(i, 22) = "PAID" Then
            GRDTranx.Row = i
            GRDTranx.Col = 22
            GRDTranx.CellForeColor = vbBlue
        ElseIf GRDTranx.TextMatrix(i, 22) = "PEND" Then
            GRDTranx.Row = i
            GRDTranx.Col = 22
            GRDTranx.CellForeColor = vbRed
        End If
        LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + Val(GRDTranx.TextMatrix(i, 4)), "0.00")
        LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + Val(GRDTranx.TextMatrix(i, 5)), "0.00")
        lblcomm.Caption = Format(Val(lblcomm.Caption) + Val(GRDTranx.TextMatrix(i, 28)), "0.00")
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    
    
    LBLBALAMT.Caption = Format(Val(lblOPBal.Caption) + (Val(LBLINVAMT.Caption) - Val(LBLPAIDAMT.Caption)), "0.00")
    lblcommpend.Caption = Val(lblcomm.Caption) - Val(lblcommpaid.Caption)
    If lblagent.Caption = "Y" Then
        GRDTranx.ColWidth(28) = 1000
        GRDTranx.ColWidth(29) = 1900
        lblcommpend.Visible = True
        lblcomm.Visible = True
        lblcommpaid.Visible = True
        LBLTOTAL(12).Visible = True
        LBLTOTAL(13).Visible = True
        LBLTOTAL(14).Visible = True
    Else
        GRDTranx.ColWidth(28) = 0
        GRDTranx.ColWidth(29) = 2900
        lblcommpend.Visible = False
        lblcomm.Visible = False
        lblcommpaid.Visible = False
        LBLTOTAL(12).Visible = False
        LBLTOTAL(13).Visible = False
        LBLTOTAL(14).Visible = False
    End If
    flagchange.Caption = ""
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    If GRDTranx.rows > 16 Then GRDTranx.TopRow = GRDTranx.rows - 1
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Function

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, 40
            
            If DataList2.VisibleCount = 0 Then TXTDEALER.SetFocus
            DataList2.text = lbldealer.Caption
            Call DataList2_Click
            'lbladdress.Caption = ""
            DataList2.SetFocus
    End Select
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Function fillcount()
    Dim i, n As Long
    
    grdcount.rows = 0
    i = 0
    LBLSelected.Caption = ""
    On Error GoTo ErrHand
    
    For n = 1 To GRDTranx.rows - 1
        If GRDTranx.TextMatrix(n, 0) <> "Sale" Then GoTo SKIP
        If GRDTranx.TextMatrix(n, 21) = "Y" Then
            grdcount.rows = grdcount.rows + 1
            grdcount.TextMatrix(i, 0) = GRDTranx.TextMatrix(n, 0)
            grdcount.TextMatrix(i, 1) = GRDTranx.TextMatrix(n, 1)
            grdcount.TextMatrix(i, 2) = GRDTranx.TextMatrix(n, 2)
            grdcount.TextMatrix(i, 3) = GRDTranx.TextMatrix(n, 3)
            grdcount.TextMatrix(i, 4) = GRDTranx.TextMatrix(n, 4)
            grdcount.TextMatrix(i, 5) = GRDTranx.TextMatrix(n, 5)
            grdcount.TextMatrix(i, 6) = GRDTranx.TextMatrix(n, 6)
            grdcount.TextMatrix(i, 7) = GRDTranx.TextMatrix(n, 7)
            grdcount.TextMatrix(i, 8) = GRDTranx.TextMatrix(n, 8)
            grdcount.TextMatrix(i, 9) = GRDTranx.TextMatrix(n, 9)
            grdcount.TextMatrix(i, 10) = GRDTranx.TextMatrix(n, 10)
            grdcount.TextMatrix(i, 11) = GRDTranx.TextMatrix(n, 11)
            grdcount.TextMatrix(i, 12) = GRDTranx.TextMatrix(n, 12)
            grdcount.TextMatrix(i, 13) = GRDTranx.TextMatrix(n, 13)
            grdcount.TextMatrix(i, 14) = GRDTranx.TextMatrix(n, 14)
            grdcount.TextMatrix(i, 15) = GRDTranx.TextMatrix(n, 15)
            grdcount.TextMatrix(i, 16) = GRDTranx.TextMatrix(n, 16)
            grdcount.TextMatrix(i, 17) = GRDTranx.TextMatrix(n, 17)
            grdcount.TextMatrix(i, 18) = GRDTranx.TextMatrix(n, 18)
            grdcount.TextMatrix(i, 19) = GRDTranx.TextMatrix(n, 19)
            grdcount.TextMatrix(i, 20) = n
            grdcount.TextMatrix(i, 21) = GRDTranx.TextMatrix(n, 21)
            grdcount.TextMatrix(i, 22) = GRDTranx.TextMatrix(n, 22)
            grdcount.TextMatrix(i, 23) = GRDTranx.TextMatrix(n, 23)
            grdcount.TextMatrix(i, 24) = GRDTranx.TextMatrix(n, 24)
            LBLSelected.Caption = Val(LBLSelected.Caption) + Val(GRDTranx.TextMatrix(n, 24))
            
            i = i + 1
        End If
SKIP:
    Next n
    
    LBLSelected.Caption = Format(LBLSelected.Caption, "0.00")
    Exit Function
ErrHand:
    MsgBox err.Description
    
End Function

'Private Function Sel_RCPTS()
'    Dim N As Long
'
'    On Error GoTo eRRhAND
'    db.Execute "Update DBTPYMT set SEL_FLAG = 'N' WHERE ACT_CODE ='" & DataList2.BoundText & "' AND SEL_FLAG = 'Y' "
'    For N = 1 To GRDTranx.Rows - 1
'        If GRDTranx.TextMatrix(N, 0) <> "Sale" And GRDTranx.TextMatrix(N, 21) = "Y" Then
'            db.Execute "Update DBTPYMT set SEL_FLAG = 'Y' WHERE (TRX_TYPE = 'RT' OR TRX_TYPE = 'DB' OR TRX_TYPE = 'CB') AND CR_NO = " & GRDTranx.TextMatrix(N, 7) & "  AND ACT_CODE ='" & DataList2.BoundText & "' AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(N, 8) & "' "
'        End If
'    Next N
'
'    Exit Function
'eRRhAND:
'    MsgBox Err.Description
'
'End Function

Private Sub grdreceipts_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If grdreceipts.rows = 1 Then Exit Sub
            'If grdreceipts.TextMatrix(grdreceipts.Row, 0) = "Receipt" Then
                Select Case grdreceipts.Col
                    Case 2
                            'If grdreceipts.Cols = 20 Then Exit Sub
                            TXTsample2.MaxLength = 7
                            TXTsample2.Visible = True
                            TXTsample2.Top = grdreceipts.CellTop + 100
                            TXTsample2.Left = grdreceipts.CellLeft '+ 50
                            TXTsample2.Width = grdreceipts.CellWidth
                            TXTsample2.Height = grdreceipts.CellHeight
                            TXTsample2.text = grdreceipts.TextMatrix(grdreceipts.Row, grdreceipts.Col)
                            TXTsample2.SetFocus
                End Select
                Exit Sub
            'End If
        Case vbKeyEscape
            FRMEMAIN.Enabled = True
            Frmereceipt.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub grdreceipts_LostFocus()
'    If Frmereceipt.Visible = True Then
'        Frmeperiod.Enabled = True
'        Frmereceipt.Visible = False
'        GRDTranx.SetFocus
'    End If
End Sub

Private Sub TXTsample2_GotFocus()
    TXTsample2.SelStart = 0
    TXTsample2.SelLength = Len(TXTsample2.text)
End Sub

Private Sub TXTsample2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdreceipts.Col
                  Case 2  'RECIEPT AMOUNT
                    
                    If Val(TXTsample2.text) = 0 Then Exit Sub
                    '==========
                    db.BeginTrans
                    'db.Execute "UPDATE `" & dbase1 & "`.`trnxrcpt` SET `RCPT_AMOUNT` = " & Val(TXTsample2.Text) & " WHERE CONCAT( `trnxrcpt`.`RCPT_NO` ) = " & Val(grdreceipts.TextMatrix(grdreceipts.Row, 3)) & " AND `trnxrcpt`.`TRX_TYPE` = '" & Val(grdreceipts.TextMatrix(grdreceipts.Row, 4)) & "' AND CONCAT( `trnxrcpt`.`INV_NO` ) = " & Val(grdreceipts.TextMatrix(grdreceipts.Row, 5)) & ""
                    db.Execute "UPDATE TRNXRCPT SET RCPT_AMOUNT = " & Val(TXTsample2.text) & " WHERE TRX_TYPE = 'RT' AND RCPT_NO = " & Val(grdreceipts.TextMatrix(grdreceipts.Row, 3)) & " AND INV_NO = " & Val(grdreceipts.TextMatrix(grdreceipts.Row, 5)) & " "
                    db.CommitTrans
                    grdreceipts.TextMatrix(grdreceipts.Row, grdreceipts.Col) = Val(TXTsample2.text)
                    grdreceipts.Enabled = True
                    TXTsample2.Visible = False
                    'Call Fillgrid
                    
                    
                    '======
                    grdreceipts.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample2.Visible = False
            grdreceipts.SetFocus
    End Select
        Exit Sub
ErrHand:
    MsgBox err.Description
    db.RollbackTrans
End Sub

Private Sub TXTsample2_KeyPress(KeyAscii As Integer)
    Select Case grdreceipts.Col
        Case 5
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
'        Case 41
'             Select Case KeyAscii
'                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
'                    KeyAscii = 0
'                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                Case Else
'                    KeyAscii = 0
'            End Select
    End Select
End Sub

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 2
                    If Not (IsDate(TXTEXPIRY.text)) Then
                        TXTEXPIRY.SetFocus
                        Exit Sub
                    End If
                    If DateValue(TXTEXPIRY.text) > DateValue(Date) Then
                        MsgBox "Date could not be higher than Today", vbOKOnly, "Receipt Register..."
                        TXTEXPIRY.SetFocus
                        Exit Sub
                    End If
                    
                    If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Receipt" Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * From DBTPYMT WHERE TRX_TYPE='RT' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock!INV_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                            
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * From DBTPYMT WHERE TRX_TYPE = 'DR' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 3) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * From TRNXRCPT WHERE TRX_TYPE = 'RT' AND CR_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 7)) & " AND CR_TRX_TYPE ='DR'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 15) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 16) = "") Then
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * FROM CASHATRXFILE WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 16) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 19) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rststock.EOF And rststock.BOF) Then
                                rststock!VCH_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                        End If
                        
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 9) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 10) = "") Then
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rststock.EOF And rststock.BOF) Then
                                rststock!TRX_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                        End If
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM DBTPYMT WHERE TRX_TYPE='DB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'DN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 15) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 16) = "") Then
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * FROM CASHATRXFILE WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 16) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 19) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rststock.EOF And rststock.BOF) Then
                                rststock!VCH_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                        End If
                        
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = 'CR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'CN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rststock.EOF And rststock.BOF) Then
                                rststock!TRX_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                        End If
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note" Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM DBTPYMT WHERE TRX_TYPE='CB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'CN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 15) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 16) = "") Then
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * FROM CASHATRXFILE WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 16) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 19) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rststock.EOF And rststock.BOF) Then
                                rststock!VCH_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                        End If
                        
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = 'DR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'DN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rststock.EOF And rststock.BOF) Then
                                rststock!TRX_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                        End If
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                    End If
                        
                    GRDTranx.Enabled = True
                    TXTEXPIRY.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                
            End Select
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            GRDTranx.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTEXPIRY_LostFocus()
    TXTEXPIRY.Visible = False
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHand
    
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 5
                    If Val(TXTsample.text) = 0 Then Exit Sub
                    
                    If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Receipt" Then
                        db.Execute "Update DBTPYMT SET RCPT_AMT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='RT' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        db.Execute "Update DBTPYMT SET RCVD_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = 'DR' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 3) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        'db.Execute "UPDATE DBTPYMT SET RCVD_AMOUNT = 0 WHERE TRX_TYPE = 'DR' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 3) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        db.Execute "Update TRNXRCPT SET RCPT_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = 'RT' AND CR_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 7)) & " AND CR_TRX_TYPE ='DR'"
                        
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 15) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 16) = "") Then
                            db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 16) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 19) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        End If
                        
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 9) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 10) = "") Then
                            db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 14) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 8) = "") Then
                            db.Execute "Update TRXMAST SET RCPT_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_YEAR= '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "' AND  TRX_TYPE= '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND VCH_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 3) & " "
                        End If
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        Call Fillgrid
                    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note" Then
                        db.Execute "Update DBTPYMT SET RCPT_AMT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='CB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'CN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 15) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 16) = "") Then
                            db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 16) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 19) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        End If
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = 'DR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'DN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        Call Fillgrid
                    End If
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                Case 4
                    If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Then
                        db.Execute "Update DBTPYMT SET RCPT_AMT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='DB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'DN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        If Not (GRDTranx.TextMatrix(GRDTranx.Row, 15) = "" Or GRDTranx.TextMatrix(GRDTranx.Row, 16) = "") Then
                            db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 15) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 16) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 17) & "' AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 19) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                        End If
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = 'CR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'CN' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        Call Fillgrid
                    End If
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                Case 6
                    db.Execute "Update DBTPYMT SET REF_NO = '" & Trim(TXTsample.text) & "' WHERE TRX_TYPE= '" & GRDTranx.TextMatrix(GRDTranx.Row, 30) & "' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 14) & "'"
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDTranx.SetFocus
    End Select
        Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDTranx.Col
        Case 5, 4
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub


