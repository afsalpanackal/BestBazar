VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FRMDAMAGEREGBR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAMAGE GOODS REPORT (BRANCH SALE)"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMDAMAGEREGBR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   18660
   Begin VB.Frame FRMEBILL 
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H00000080&
      Height          =   4725
      Left            =   60
      TabIndex        =   8
      Top             =   1950
      Visible         =   0   'False
      Width           =   10845
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4005
         Left            =   30
         TabIndex        =   9
         Top             =   540
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   7064
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "NET AMT"
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
         Height          =   255
         Index           =   6
         Left            =   8565
         TabIndex        =   17
         Top             =   210
         Width           =   825
      End
      Begin VB.Label LBLNETAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9390
         TabIndex        =   16
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "DISC"
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
         Height          =   255
         Index           =   2
         Left            =   7320
         TabIndex        =   15
         Top             =   210
         Width           =   495
      End
      Begin VB.Label LBLDISC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7785
         TabIndex        =   14
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LBLBILLAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6150
         TabIndex        =   13
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "BILL AMT"
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
         Height          =   255
         Index           =   1
         Left            =   5190
         TabIndex        =   12
         Top             =   210
         Width           =   885
      End
      Begin VB.Label LBLTOTAL 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   3300
         TabIndex        =   11
         Top             =   210
         Width           =   780
      End
      Begin VB.Label LBLBILLNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4125
         TabIndex        =   10
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame1"
      Height          =   9885
      Left            =   -120
      TabIndex        =   0
      Top             =   -285
      Width           =   18720
      Begin VB.CommandButton cmdwoprint 
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
         Height          =   405
         Left            =   135
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   44
         Top             =   9450
         Width           =   420
      End
      Begin VB.CommandButton CMDEfile 
         Caption         =   "Print Report for e-Filing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6930
         TabIndex        =   43
         Top             =   8805
         Width           =   1530
      End
      Begin VB.CommandButton CmdReport 
         Caption         =   "Print Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5400
         TabIndex        =   42
         Top             =   8805
         Width           =   1440
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0FFFF&
         Caption         =   "DAMAGEGOODSREGISTER"
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
         Height          =   1890
         Left            =   150
         TabIndex        =   31
         Top             =   285
         Width           =   18510
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0FFFF&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   75
            TabIndex        =   34
            Top             =   420
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton OPTCUSTOMER 
            BackColor       =   &H00C0FFFF&
            Caption         =   "CUSTOMER"
            Height          =   210
            Left            =   90
            TabIndex        =   33
            Top             =   870
            Width           =   1320
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
            Left            =   1845
            TabIndex        =   32
            Top             =   825
            Width           =   3720
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1860
            TabIndex        =   35
            Top             =   330
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   123076609
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   4035
            TabIndex        =   36
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   123076609
            CurrentDate     =   40498
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1845
            TabIndex        =   37
            Top             =   1170
            Width           =   3720
            _ExtentX        =   6562
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
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
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
            Index           =   4
            Left            =   1110
            TabIndex        =   41
            Top             =   405
            Width           =   555
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
            Index           =   5
            Left            =   3585
            TabIndex        =   40
            Top             =   405
            Width           =   285
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   6465
            TabIndex        =   39
            Top             =   1965
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   8685
            TabIndex        =   38
            Top             =   1905
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdview 
         Caption         =   "&View Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8490
         TabIndex        =   28
         Top             =   8310
         Width           =   1545
      End
      Begin VB.CommandButton CMDREGISTER 
         Caption         =   "PRINT REGISTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6930
         TabIndex        =   27
         Top             =   8310
         Width           =   1515
      End
      Begin VB.CommandButton CMDPRINTREGISTER 
         Caption         =   "Print Item Wise Report"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5415
         TabIndex        =   3
         Top             =   8310
         Width           =   1425
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
         Height          =   450
         Left            =   13035
         TabIndex        =   2
         Top             =   8295
         Width           =   1335
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
         Height          =   450
         Left            =   11625
         TabIndex        =   1
         Top             =   8295
         Width           =   1380
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   5595
         Left            =   165
         TabIndex        =   7
         Top             =   2205
         Width           =   18465
         _ExtentX        =   32570
         _ExtentY        =   9869
         _Version        =   393216
         Rows            =   1
         Cols            =   16
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         Appearance      =   0
         GridLineWidth   =   2
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1350
         Left            =   120
         TabIndex        =   4
         Top             =   7830
         Width           =   4995
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   12
            Left            =   45
            TabIndex        =   30
            Top             =   945
            Width           =   1245
         End
         Begin VB.Label lblcommi 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   1395
            TabIndex        =   29
            Top             =   900
            Width           =   1320
         End
         Begin VB.Label LBLNET 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   3675
            TabIndex        =   25
            Top             =   930
            Width           =   1320
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "NET AMT"
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
            Height          =   495
            Index           =   10
            Left            =   2790
            TabIndex        =   24
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label LBLDISCOUNT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   1395
            TabIndex        =   23
            Top             =   465
            Width           =   1320
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "DISCOUNT"
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
            Index           =   9
            Left            =   45
            TabIndex        =   22
            Top             =   510
            Width           =   1155
         End
         Begin VB.Label LBLPROFIT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   3675
            TabIndex        =   21
            Top             =   495
            Width           =   1320
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "PROFIT"
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
            Index           =   8
            Left            =   2775
            TabIndex        =   20
            Top             =   510
            Width           =   810
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "COST"
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
            Index           =   7
            Left            =   2775
            TabIndex        =   19
            Top             =   60
            Width           =   660
         End
         Begin VB.Label LBLCOST 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   3675
            TabIndex        =   18
            Top             =   45
            Width           =   1320
         End
         Begin VB.Label LBLTRXTOTAL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   360
            Left            =   1400
            TabIndex        =   6
            Top             =   45
            Width           =   1320
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "BILL AMOUNT"
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
            Index           =   3
            Left            =   45
            TabIndex        =   5
            Top             =   105
            Width           =   1365
         End
      End
      Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
         Height          =   330
         Left            =   5430
         TabIndex        =   26
         Tag             =   "5"
         Top             =   7860
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   582
         Picture         =   "FRMDAMAGEREGBR.frx":030A
         ForeColor       =   0
         BarPicture      =   "FRMDAMAGEREGBR.frx":0326
         Max             =   150
         Text            =   "PLEASE WAIT..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
   End
End
Attribute VB_Name = "FRMDAMAGEREGBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim selectionformla As String

Private Sub CmDDisplay_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
    db.Execute "delete From SALESREG"
    
    LBLTRXTOTAL.Caption = "0.00"
    LBLDISCOUNT.Caption = "0.00"
    LBLNET.Caption = "0.00"
    lblcost.Caption = "0.00"
    LBLPROFIT.Caption = "0.00"
    lblCommi.Caption = "0.00"
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    n = 1
    M = 0
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    
    If OPTPERIOD.Value = True Then
        rstTRANX.Open "SELECT * From DAMGMASTVAN WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='DG' ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From DAMGMASTVAN WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='DG' ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
        
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!TRX_TYPE
        Select Case rstTRANX!TRX_TYPE
            Case "GI"
                GRDTranx.TextMatrix(M, 2) = "GI"
            Case "DG"
                GRDTranx.TextMatrix(M, 2) = "DG"
            Case "DM"
                GRDTranx.TextMatrix(M, 2) = "DM"
            Case "SI"
                GRDTranx.TextMatrix(M, 2) = "WS"
            Case "WO"
                GRDTranx.TextMatrix(M, 2) = "PT"
        End Select
        GRDTranx.TextMatrix(M, 3) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 4) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 5) = Format(Round(rstTRANX!VCH_AMOUNT, 2), "0.00")
'        If rstTRANX!SLSM_CODE = "A" Then
'
'        ElseIf rstTRANX!SLSM_CODE = "P" Then
'            GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(Round((rstTRANX!DISCOUNT * 100 / rstTRANX!VCH_AMOUNT), 2), "0.00"))
'        End If
        GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(rstTRANX!DISCOUNT, "0.00"))
        GRDTranx.TextMatrix(M, 7) = Format(Round(rstTRANX!NET_AMOUNT, 2), "0.00") 'Format(Round(Val(GRDTranx.TextMatrix(M, 5)) - Val(GRDTranx.TextMatrix(M, 6)), 2), "0.00")
        
        CMDEXIT.Tag = IIf(IsNull(rstTRANX!DISCOUNT), "0", Format(rstTRANX!DISCOUNT, "0.00"))
        'GRDTranx.TextMatrix(M, 7) = Format(Round(Val(GRDTranx.TextMatrix(M, 5)), 2), "0.00")
        If frmLogin.rs!Level <> "0" Then
            GRDTranx.TextMatrix(M, 8) = "xxx"
            GRDTranx.TextMatrix(M, 9) = "xxx"
        Else
            GRDTranx.TextMatrix(M, 8) = IIf(IsNull(rstTRANX!COMM_AMT), "0", Format(rstTRANX!COMM_AMT, "0.00"))
            GRDTranx.TextMatrix(M, 9) = IIf(IsNull(rstTRANX!PAY_AMOUNT), "0", Format(rstTRANX!PAY_AMOUNT, "0.00"))
        End If
        GRDTranx.TextMatrix(M, 10) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        GRDTranx.TextMatrix(M, 11) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS), "", ", " & rstTRANX!BILL_ADDRESS)
        GRDTranx.TextMatrix(M, 15) = IIf(IsNull(rstTRANX!CST), "", Format(rstTRANX!CST, "0.00"))
        CMDDISPLAY.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
        
        'If rstTRANX!TRX_TYPE <> "SI" Then GoTo SKIP
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select DISTINCT SALES_TAX From TRXFILEVAN WHERE TRX_YEAR ='" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            EXSALEAMT = 0
            TAXSALEAMT = 0
            TaxAmt = 0
            MRPVALUE = 0
            DISCAMT = 0
            TAXRATE = RSTTRXFILE!SALES_TAX
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILEVAN WHERE TRX_YEAR ='" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & RSTTRXFILE!SALES_TAX & "", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                If RSTTRXFILE!SALES_TAX > 0 And RSTtax!check_flag = "V" Then
                    TAXSALEAMT = TAXSALEAMT + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                    TaxAmt = TaxAmt + Round((RSTtax!PTR * RSTtax!SALES_TAX / 100) * RSTtax!QTY, 2)
                    
                Else
                    If RSTtax!SALE_1_FLAG = "1" Then
                        TaxAmt = TaxAmt + Round((RSTtax!SALES_PRICE - RSTtax!PTR) * RSTtax!QTY, 2)
                        MRPVALUE = Round(MRPVALUE + (100 * RSTtax!MRP / 105) * RSTtax!QTY, 2)
                    End If
                    EXSALEAMT = EXSALEAMT + RSTtax!TRX_TOTAL
                End If
                DISCAMT = Round(DISCAMT + IIf(IsNull(RSTtax!LINE_DISC), 0, RSTtax!TRX_TOTAL * RSTtax!LINE_DISC / 100), 2)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            RSTSALEREG.AddNew
            TAXSALEAMT = TAXSALEAMT - TaxAmt
            RSTSALEREG!VCH_NO = rstTRANX!VCH_NO 'N
            RSTSALEREG!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTSALEREG!VCH_DATE = rstTRANX!VCH_DATE
            RSTSALEREG!DISCOUNT = DISCAMT
            RSTSALEREG!VCH_AMOUNT = Val(GRDTranx.TextMatrix(M, 7))
            RSTSALEREG!PAYAMOUNT = Val(GRDTranx.TextMatrix(M, 9))
            RSTSALEREG!ACT_NAME = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            RSTSALEREG!ACT_CODE = IIf(IsNull(rstTRANX!ACT_CODE), "", rstTRANX!ACT_CODE)
            RSTSALEREG!CST_PER = IIf(IsNull(rstTRANX!CST), 0, rstTRANX!CST)
            
            Dim RSTACTCODE As ADODB.Recordset
            Set RSTACTCODE = New ADODB.Recordset
            RSTACTCODE.Open "SELECT KGST FROM CUSTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTACTCODE.EOF And RSTACTCODE.BOF) Then
                RSTSALEREG!TIN_NO = RSTACTCODE!KGST
            End If
            RSTACTCODE.Close
            Set RSTACTCODE = Nothing
            RSTSALEREG!EXMPSALES_AMT = EXSALEAMT
            RSTSALEREG!TAXSALES_AMT = TAXSALEAMT
            RSTSALEREG!TAXAMOUNT = TaxAmt
            RSTSALEREG!TAXRATE = TAXRATE
            CMDDISPLAY.Tag = Val(CMDDISPLAY.Tag) + EXSALEAMT
            FRMEMAIN.Tag = Val(FRMEMAIN.Tag) + TAXSALEAMT
            FRMEBILL.Tag = Val(FRMEBILL.Tag) + TaxAmt
            RSTSALEREG.Update
            
            RSTTRXFILE.MoveNext
        Loop
    
        GRDTranx.TextMatrix(M, 12) = Format(Val(CMDDISPLAY.Tag), "0.00")
        GRDTranx.TextMatrix(M, 13) = Format(Val(FRMEMAIN.Tag), "0.00")
        GRDTranx.TextMatrix(M, 14) = Format(Val(FRMEBILL.Tag), "0.00")
        
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption) + Val(GRDTranx.TextMatrix(M, 6)), "0.00")
        If frmLogin.rs!Level <> "0" Then
            lblCommi.Caption = "xxx"
            lblcost.Caption = "xxx"
            LBLPROFIT.Caption = "xxx"
        Else
            lblCommi.Caption = Format(Val(lblCommi.Caption) + Val(GRDTranx.TextMatrix(M, 8)), "0.00")
            lblcost.Caption = Format(Val(lblcost.Caption) + Val(GRDTranx.TextMatrix(M, 9)), "0.00")
            'LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
        End If
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        n = n + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    LBLNET.Caption = Format(Val(LBLTRXTOTAL.Caption) - Val(LBLDISCOUNT.Caption), "0.00")
    If frmLogin.rs!Level = "0" Then
        LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(lblcost.Caption) + Val(lblCommi.Caption)), "0.00")
    End If
        
    flagchange.Caption = ""
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    CMDPRINTREGISTER.Enabled = True
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub CMDDISPLAY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            DTTo.SetFocus
    End Select
End Sub

Private Sub CMDEfile_Click()
    Exit Sub
   Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
    db.Execute "delete From SALESREG"
    
    LBLTRXTOTAL.Caption = "0.00"
    LBLDISCOUNT.Caption = "0.00"
    LBLNET.Caption = "0.00"
    lblcost.Caption = "0.00"
    LBLPROFIT.Caption = "0.00"
    lblCommi.Caption = "0.00"
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    n = 1
    M = 0
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    If OPTPERIOD.Value = True Then
        rstTRANX.Open "SELECT * From DAMGMASTVAN WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='DG' ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From DAMGMASTVAN WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='DG' ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
        
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG", db, adOpenStatic, adLockOptimistic, adCmdText

    Do Until rstTRANX.EOF
        M = M + 1
        
        CMDDISPLAY.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
        
        'If rstTRANX!TRX_TYPE <> "SI" Then GoTo SKIP
        
        EXSALEAMT = 0
        TAXSALEAMT = 0
        TaxAmt = 0
        MRPVALUE = 0
        DISCAMT = 0
        'TAXRATE = RSTTRXFILE!SALES_TAX
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select * From TRXFILEVAN WHERE VCH_NO = " & rstTRANX!VCH_NO & " ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            If RSTtax!SALES_TAX > 0 And RSTtax!check_flag = "V" Then
                TAXSALEAMT = TAXSALEAMT + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                TaxAmt = TaxAmt + Round((RSTtax!PTR * RSTtax!SALES_TAX / 100) * RSTtax!QTY, 2)
                
            Else
                If RSTtax!SALE_1_FLAG = "1" Then
                    TaxAmt = TaxAmt + Round((RSTtax!SALES_PRICE - RSTtax!PTR) * RSTtax!QTY, 2)
                    MRPVALUE = Round(MRPVALUE + (100 * RSTtax!MRP / 105) * RSTtax!QTY, 2)
                End If
                EXSALEAMT = EXSALEAMT + RSTtax!TRX_TOTAL
            End If
            DISCAMT = Round(DISCAMT + IIf(IsNull(RSTtax!LINE_DISC), 0, RSTtax!TRX_TOTAL * RSTtax!LINE_DISC / 100), 2)
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
        RSTSALEREG.AddNew
        TAXSALEAMT = TAXSALEAMT - TaxAmt
        RSTSALEREG!VCH_NO = rstTRANX!VCH_NO 'N
        RSTSALEREG!TRX_TYPE = rstTRANX!TRX_TYPE
        RSTSALEREG!VCH_DATE = rstTRANX!VCH_DATE
        RSTSALEREG!DISCOUNT = DISCAMT
        RSTSALEREG!VCH_AMOUNT = rstTRANX!VCH_AMOUNT - IIf(IsNull(rstTRANX!DISCOUNT), "", rstTRANX!DISCOUNT)
        'RSTSALEREG!PAYAMOUNT = Val(GRDTranx.TextMatrix(M, 9))
        RSTSALEREG!ACT_NAME = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
        RSTSALEREG!ACT_CODE = IIf(IsNull(rstTRANX!ACT_CODE), "", rstTRANX!ACT_CODE)
        
        Dim RSTACTCODE As ADODB.Recordset
        Set RSTACTCODE = New ADODB.Recordset
        RSTACTCODE.Open "SELECT KGST FROM CUSTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTACTCODE.EOF And RSTACTCODE.BOF) Then
            RSTSALEREG!TIN_NO = RSTACTCODE!KGST
        End If
        RSTACTCODE.Close
        Set RSTACTCODE = Nothing
        RSTSALEREG!EXMPSALES_AMT = EXSALEAMT
        RSTSALEREG!TAXSALES_AMT = TAXSALEAMT
        RSTSALEREG!TAXAMOUNT = TaxAmt
        RSTSALEREG!TAXRATE = TAXRATE
        CMDDISPLAY.Tag = Val(CMDDISPLAY.Tag) + EXSALEAMT
        FRMEMAIN.Tag = Val(FRMEMAIN.Tag) + TAXSALEAMT
        FRMEBILL.Tag = Val(FRMEBILL.Tag) + TaxAmt
        RSTSALEREG.Update
            
        
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        n = n + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    LBLNET.Caption = Format(Val(LBLTRXTOTAL.Caption) - Val(LBLDISCOUNT.Caption), "0.00")
    If frmLogin.rs!Level = "0" Then
        LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(lblcost.Caption) + Val(lblCommi.Caption)), "0.00")
    End If
        
    flagchange.Caption = ""
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    CMDPRINTREGISTER.Enabled = True
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub CMDREGISTER_Click()
    
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    ReportNameVar = Rptpath & "RPTDAMREGVAN"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdReport_Click()
    
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTDAMREPVAN"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.RecordSelectionFormula = "( {TRXMAST.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub cmdview_Click()
   
    
    Dim TRXFILE As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim FROMDATE As Date
    Dim TODATE As Date
    Dim SLIPAMT As Double
    Dim M As Long
    
    db.Execute "delete From SLIP_REG"
    
    FROMDATE = DTFROM.Value 'Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = DTTo.Value 'Format(DTTO.Value, "MM,DD,YYYY")

    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From DAMGMASTVAN WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='DG' ", db, adOpenStatic, adLockReadOnly
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SLIP_REG", db, adOpenStatic, adLockOptimistic, adCmdText
    M = 0
    Do Until FROMDATE > TODATE
        SLIPAMT = 0
        M = M + 1
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * From DAMGMASTVAN WHERE VCH_DATE = '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='DG' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until RSTTRXFILE.EOF
            CMDDISPLAY.Tag = ""
            If RSTTRXFILE!SLSM_CODE = "A" Then
                CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(RSTTRXFILE!DISCOUNT, "0.00"))
            ElseIf RSTTRXFILE!SLSM_CODE = "P" Then
                CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(Round((RSTTRXFILE!DISCOUNT * RSTTRXFILE!VCH_AMOUNT) / 100, 2), "0.00"))
            End If
            cmdview.Tag = ""
            'cmdview.Tag = IIf(IsNull(RSTTRXFILE!ADD_AMOUNT), "", RSTTRXFILE!ADD_AMOUNT)
            SLIPAMT = SLIPAMT + Round(RSTTRXFILE!VCH_AMOUNT - Val(CMDDISPLAY.Tag), 0) '+ Val(cmdview.Tag))
            RSTTRXFILE.MoveNext
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        Loop
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTSALEREG.AddNew
            RSTTRXFILE.MoveLast
            RSTSALEREG!VCH_END_NO = RSTTRXFILE!VCH_NO
            RSTTRXFILE.MoveFirst
            RSTSALEREG!VCH_START_NO = RSTTRXFILE!VCH_NO
            RSTSALEREG!VCH_DATE = RSTTRXFILE!VCH_DATE
            RSTSALEREG!REC_NO = M
            RSTSALEREG!TRX_TYPE = "S"
            RSTSALEREG!VCH_AMOUNT = SLIPAMT
            RSTSALEREG.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing

        FROMDATE = DateAdd("d", FROMDATE, 1)
    Loop
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    'CHECKFLAG = 1
    
    Sleep (300)
    ReportNameVar = Rptpath & "RPTDAMREGSTRVAN"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'selectionformla = "( {TRXFILE.FREE_QTY}>0 and {TRXFILE.VCH_DATE}<=# " & TODATE & " # and {TRXFILE.VCH_DATE}>=# " & FROMDATE & " # and {TRXFILE.MFGR}='" & DataList3.BoundText & "')"
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
    Set CRXFormulaFields = Report.FormulaFields
    
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
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "SALES REGISTER"
    
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    CMDPRINTREGISTER.Enabled = True
    
    GENERATEREPORT
    'GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Sub

Private Sub cmdwoprint_Click()
    FRMBillTransfer.Show
    FRMBillTransfer.SetFocus
End Sub

Private Sub DTFROM_GotFocus()
    CMDPRINTREGISTER.Enabled = False
End Sub

Private Sub DTFROM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DTTo.SetFocus
    End Select
End Sub

Private Sub DTTO_GotFocus()
    CMDPRINTREGISTER.Enabled = False
End Sub

Private Sub DTTO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            DTFROM.SetFocus
    End Select
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CMDPRINTREGISTER_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    ReportNameVar = Rptpath & "RPTDAMREPORTVAN"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "(({TRXFILE.TRX_TYPE}='DG' OR {TRXFILE.TRX_TYPE}='WO' OR {TRXFILE.TRX_TYPE}='RI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * From TRXFILEVAN ")
        Report.Database.SetDataSource oRs, 3, 1
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * From DAMGMASTVAN ")
        Report.Database.SetDataSource oRs, 3, 2
        Set oRs = Nothing
    End If
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

Private Sub Form_Load()
    'If Month(Date) > 1 Then
        'CMBMONTH.ListIndex = Month(Date) - 2
    'Else
        'CMBMONTH.ListIndex = 11
    'End If
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "Type"
    GRDTranx.TextMatrix(0, 2) = "TYPE"
    GRDTranx.TextMatrix(0, 3) = "BILL NO"
    GRDTranx.TextMatrix(0, 4) = "BILL DATE"
    GRDTranx.TextMatrix(0, 5) = "BILL AMT"
    GRDTranx.TextMatrix(0, 6) = "DISC AMT"
    GRDTranx.TextMatrix(0, 7) = "NET AMT"
    GRDTranx.TextMatrix(0, 8) = "commission"
    GRDTranx.TextMatrix(0, 9) = "COST VALUE"
    GRDTranx.TextMatrix(0, 10) = "CUSTOMER"
    GRDTranx.TextMatrix(0, 11) = "Bill Address"
    GRDTranx.TextMatrix(0, 12) = "EX. SALES"
    GRDTranx.TextMatrix(0, 13) = "TAX SALES"
    GRDTranx.TextMatrix(0, 14) = "TAX AMT"
    GRDTranx.TextMatrix(0, 15) = "CST"
    
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 0
    GRDTranx.ColWidth(2) = 1500
    GRDTranx.ColWidth(3) = 1500
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 1500
    GRDTranx.ColWidth(6) = 1500
    GRDTranx.ColWidth(7) = 1500
    If frmLogin.rs!Level <> "0" Then
        GRDTranx.ColWidth(8) = 0
        GRDTranx.ColWidth(9) = 0
    Else
        GRDTranx.ColWidth(8) = 1200
        GRDTranx.ColWidth(9) = 1200
    End If
    GRDTranx.ColWidth(10) = 3000
    GRDTranx.ColWidth(11) = 3000
    GRDTranx.ColWidth(12) = 0
    GRDTranx.ColWidth(13) = 0
    GRDTranx.ColWidth(14) = 0
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 3
    GRDTranx.ColAlignment(4) = 3
    GRDTranx.ColAlignment(5) = 3
    GRDTranx.ColAlignment(6) = 6
    GRDTranx.ColAlignment(7) = 6
    GRDTranx.ColAlignment(8) = 6
    GRDTranx.ColAlignment(9) = 6
    GRDTranx.ColAlignment(10) = 1
    GRDTranx.ColAlignment(11) = 1
    GRDTranx.ColAlignment(12) = 6
    GRDTranx.ColAlignment(13) = 6
    GRDTranx.ColAlignment(14) = 6
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "Rate"
    GRDBILL.TextMatrix(0, 3) = "Disc %"
    GRDBILL.TextMatrix(0, 4) = "Tax %"
    GRDBILL.TextMatrix(0, 5) = "Qty"
    GRDBILL.TextMatrix(0, 6) = "Amount"
    GRDBILL.TextMatrix(0, 7) = "Batch"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 2800
    GRDBILL.ColWidth(2) = 800
    GRDBILL.ColWidth(3) = 800
    GRDBILL.ColWidth(4) = 800
    GRDBILL.ColWidth(5) = 900
    GRDBILL.ColWidth(6) = 1100
    GRDBILL.ColWidth(7) = 1100
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 3
    GRDBILL.ColAlignment(6) = 6
    GRDBILL.ColAlignment(7) = 1
        
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 11130
    'Me.Height = 10125
    Me.Left = 1500
    Me.Top = 0
    ACT_FLAG = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close

    MDIMAIN.PCTMENU.Enabled = True
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
        FRMEBILL.Visible = False
        GRDTranx.SetFocus
    End If
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.rows = 1 Then Exit Sub
            LBLBILLNO.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            LBLBILLAMT.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 5), "0.00")
            LBLDISC.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 6), "0.00")
            LBLNETAMT.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 7), "0.00")
             
            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILEVAN WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "  AND TRX_TYPE = '" & Trim(GRDTranx.TextMatrix(GRDTranx.Row, 1)) & "'", db, adOpenStatic, adLockReadOnly
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
                GRDBILL.TextMatrix(i, 7) = RSTTRXFILE!REF_NO
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing

            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    End Select
End Sub

'Private Sub TMPDELETE_Click()
'    If GRDTranx.Rows = 1 Then Exit Sub
'    If MsgBox("Are You Sure You want to Delete PRINT_BILL NO." & "*** " & GRDTranx.TextMatrix(GRDTranx.Row, 2) & " ****", vbYesNo, "DELETING BILL....") = vbNo Then Exit Sub
'    DB.Execute ("DELETE from SALESREG WHERE VCH_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 2) & " AND (TRX_TYPE='DG' OR TRX_TYPE='SI')")
'    Call fillSTOCKREG
'
'End Sub
'
'Private Function fillSTOCKREG()
'    Dim rstTRANX As ADODB.Recordset
'    Dim i As lONG
'
'    LBLTRXTOTAL.Caption = "0.00"
'    LBLDISCOUNT.Caption = "0.00"
'    LBLNET.Caption = "0.00"
'    LBLCOST.Caption = "0.00"
'    LBLPROFIT.Caption = "0.00"
'
'   On Error GoTo eRRHAND
'
'
'    Screen.MousePointer = vbHourglass
'
'    GRDTranx.Rows = 1
'    i = 0
'    GRDTranx.Visible = False
'    vbalProgressBar1.Value = 0
'    vbalProgressBar1.ShowText = True
'    vbalProgressBar1.Text = "PLEASE WAIT..."
'
'    Set rstTRANX = New ADODB.Recordset
'    rstTRANX.Open "SELECT * From SALESREG", DB, adOpenStatic,adLockReadOnly
'    Do Until rstTRANX.EOF
'        i = i + 1
'        GRDTranx.Rows = GRDTranx.Rows + 1
'        GRDTranx.FixedRows = 1
'        GRDTranx.TextMatrix(i, 0) = i
'        GRDTranx.TextMatrix(i, 2) = rstTRANX!VCH_NO
'        GRDTranx.TextMatrix(i, 3) = rstTRANX!VCH_DATE
'        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!VCH_AMOUNT, "0.00")
'        GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!DISCOUNT, "0.00")
'        GRDTranx.TextMatrix(i, 6) = Format(Val(GRDTranx.TextMatrix(i, 4)) - Val(GRDTranx.TextMatrix(i, 4)) * Val(GRDTranx.TextMatrix(i, 5)) / 100)
'        GRDTranx.TextMatrix(i, 7) = Format(rstTRANX!PAYAMOUNT, "0.00")
'        GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
'
'        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
'        LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption) + rstTRANX!DISCOUNT, "0.00")
'        LBLNET.Caption = Format(Val(LBLTRXTOTAL.Caption) - Val(LBLDISCOUNT.Caption), "0.00")
'        LBLCOST.Caption = Format(Val(LBLCOST.Caption) + rstTRANX!PAYAMOUNT, "0.00")
'        LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
'
'        vbalProgressBar1.Max = rstTRANX.RecordCount
'        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
'    Loop
'
'    rstTRANX.Close
'    Set rstTRANX = Nothing
'
'    vbalProgressBar1.ShowText = False
'    vbalProgressBar1.Value = 0
'    GRDTranx.Visible = True
'    Screen.MousePointer = vbDefault
'    Exit Function
'
'eRRHAND:
'    Screen.MousePointer = vbDefault
'    MsgBox Err.Description
'End Function

Private Sub ReportGeneratION()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    Dim rstSUBfile As ADODB.Recordset
    Dim SN As Integer
    Dim TRXTOTAL As Double
    
    SN = 0
    TRXTOTAL = 0
   ' On Error GoTo errHand
    '//NOTE : Report file name should never contain blank space.
    On Error GoTo CLOSEFILE
    Open App.Path & "\Report.PRN" For Output As #1 '//Report file Creation
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open App.Path & "\Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)

    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!Address & ", " & RSTCOMPANY!HO_NAME, 140)
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(48) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(48) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(48) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "SALES SUMMARY FOR THE PERIOD"
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "FROM " & DTFROM.Value & " TO " & DTTo.Value
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Set RSTTRXFILE = New ADODB.Recordset
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(13) & RepeatString("-", 59)
    Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SN", 3) & Space(2) & _
            AlignLeft("INV DATE", 8) & Space(10) & _
            AlignLeft("INV AMT", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Print #1, Space(12) & RepeatString("-", 59)
    SN = 0
    RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        SN = SN + 1
        Print #1, Chr(27) & Chr(71) & Space(5) & Chr(14) & Chr(15) & AlignRight(str(SN), 4) & ". " & Space(1) & _
            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
            AlignRight(Format(Round(RSTTRXFILE!VCH_AMOUNT, 0), "0.00"), 16)
        'Print #1, Chr(13)
        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Print #1,
    
    Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(13) & AlignRight("TOTAL AMOUNT", 12) & AlignRight((Format(TRXTOTAL, "####.00")), 11)
    Print #1, Space(56) & RepeatString("-", 16)
    'Print #1, Chr(27) & Chr(67) & Chr(0)
    'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** WISH YOU A SPEEDY RECOVERY ****", 40)


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
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    
    'Print #1, Chr(27) & Chr(80)
    Close #1 '//Closing the file
    'MsgBox "Report file generated at " & App.Path & "\Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Sub

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Function ReportREGISTER()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    Dim rstSUBfile As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim SN As Integer
    Dim TRXTOTAL As Double
    
    SN = 0
    TRXTOTAL = 0
    '//NOTE : Report file name should never contain blank space.
    db.Execute "delete From SALESREG2"
    
    On Error GoTo CLOSEFILE
    Open App.Path & "\Report.PRN" For Output As #1 '//Report file Creation
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open App.Path & "\Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)

    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!Address & ", " & RSTCOMPANY!HO_NAME, 140)
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(48) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(48) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(48) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "SALES REGSITER FOR THE PERIOD"
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "FROM " & DTFROM.Value & " TO " & DTTo.Value
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Set RSTTRXFILE = New ADODB.Recordset
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(13) & RepeatString("-", 59)
    Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SN", 3) & Space(2) & _
            AlignLeft("INV DATE", 8) & Space(10) & _
            AlignLeft("INV AMT", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Print #1, Space(12) & RepeatString("-", 59)
    SN = 0
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG2", db, adOpenStatic, adLockOptimistic, adCmdText
    'RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_NO", DB, adOpenStatic,adLockReadOnly
    RSTTRXFILE.Open "SELECT * From DAMGMASTVAN WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='DG' OR TRX_TYPE='RI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        SN = SN + 1
        CMDDISPLAY.Tag = ""
        If RSTTRXFILE!SLSM_CODE = "A" Then
            CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(Round((RSTTRXFILE!DISCOUNT * RSTTRXFILE!VCH_AMOUNT) / 100, 2), "0.00"))
        ElseIf RSTTRXFILE!SLSM_CODE = "P" Then
            CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(RSTTRXFILE!DISCOUNT, "0.00"))
        End If
        cmdview.Tag = ""
        cmdview.Tag = IIf(IsNull(RSTTRXFILE!ADD_AMOUNT), "", RSTTRXFILE!ADD_AMOUNT)
        'SLIPAMT = SLIPAMT + RSTTRXFILE!VCH_AMOUNT - (Val(CMDDISPLAY.Tag) + Val(cmdview.Tag))
        Print #1, Chr(27) & Chr(71) & Space(5) & Chr(14) & Chr(15) & AlignRight(str(SN), 4) & ". " & Space(1) & _
            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
            AlignRight(Format(Round(RSTTRXFILE!VCH_AMOUNT - (Val(CMDDISPLAY.Tag) + Val(cmdview.Tag)), 0), "0.00"), 16)
        'Print #1, Chr(13)
        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
        
        RSTSALEREG.AddNew
        RSTSALEREG!VCH_NO = RSTTRXFILE!VCH_NO
        RSTSALEREG!TRX_TYPE = "DG"
        RSTSALEREG!VCH_DATE = RSTTRXFILE!VCH_DATE
        RSTSALEREG!VCH_AMOUNT = RSTTRXFILE!VCH_AMOUNT
        RSTSALEREG!PAYAMOUNT = 0 ' TRXFILE!PAY_AMOUNT
        RSTSALEREG!ACT_NAME = "Sales"
        RSTSALEREG!ACT_CODE = "111001"
        RSTSALEREG!DISCOUNT = 0 'rstTRANX!DISCOUNT
        RSTSALEREG.Update
        
        RSTTRXFILE.MoveNext
    Loop
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Print #1,
    
    Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(13) & AlignRight("TOTAL AMOUNT", 12) & AlignRight((Format(TRXTOTAL, "####.00")), 11)
    Print #1, Space(56) & RepeatString("-", 16)
    'Print #1, Chr(27) & Chr(67) & Chr(0)
    'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** WISH YOU A SPEEDY RECOVERY ****", 40)


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
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    
    'Print #1, Chr(27) & Chr(80)
    Close #1 '//Closing the file
    'MsgBox "Report file generated at " & App.Path & "\Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub OPTCUSTOMER_Click()
    TXTDEALER.SetFocus
End Sub

Private Sub OPTCUSTOMER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub


Private Sub TxtAgent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_GotFocus()
    OPTCUSTOMER.Value = True
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            OPTPERIOD.SetFocus
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
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    GRDTranx.rows = 1
    LBLTRXTOTAL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            OPTPERIOD.SetFocus
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

