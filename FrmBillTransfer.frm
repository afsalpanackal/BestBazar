VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FRMBillTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES REPORT"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   11025
   Begin VB.Frame FRMEBILL 
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H00000080&
      Height          =   4725
      Left            =   75
      TabIndex        =   8
      Top             =   1350
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
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
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
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame1"
      Height          =   9885
      Left            =   -120
      TabIndex        =   0
      Top             =   -270
      Width           =   11145
      Begin VB.CommandButton CMDCONVERT2 
         Caption         =   "CONVERT"
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
         Left            =   9525
         TabIndex        =   47
         Top             =   7815
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Move Un Bill Items from B2C && B2B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   5475
         TabIndex        =   46
         Top             =   8265
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete Un Bill Items from B2C && B2B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   150
         TabIndex        =   45
         Top             =   8790
         Width           =   1695
      End
      Begin VB.CommandButton cmdmove 
         Caption         =   "MOVE"
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
         Left            =   9525
         TabIndex        =   44
         Top             =   8340
         Width           =   1335
      End
      Begin VB.CommandButton CMDCONVERT 
         Caption         =   "CONVERT"
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
         Left            =   9525
         TabIndex        =   43
         Top             =   7815
         Width           =   1335
      End
      Begin VB.PictureBox picChecked 
         Height          =   285
         Left            =   255
         Picture         =   "FrmBillTransfer.frx":0000
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   38
         Top             =   1500
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picUnchecked 
         Height          =   285
         Left            =   585
         Picture         =   "FrmBillTransfer.frx":0342
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   37
         Top             =   1530
         Visible         =   0   'False
         Width           =   285
      End
      Begin MSFlexGridLib.MSFlexGrid grdcount 
         Height          =   5145
         Left            =   13605
         TabIndex        =   36
         Top             =   1875
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   9075
         _Version        =   393216
         Rows            =   1
         Cols            =   15
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
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H0080C0FF&
         Caption         =   "SALES REGISTER"
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
         Height          =   825
         Left            =   150
         TabIndex        =   23
         Top             =   285
         Width           =   10950
         Begin VB.Frame Frame3 
            BackColor       =   &H0080C0FF&
            Height          =   645
            Left            =   5655
            TabIndex        =   40
            Top             =   150
            Width           =   3840
            Begin VB.OptionButton Optpetty 
               BackColor       =   &H0080C0FF&
               Caption         =   "Petty"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   270
               Left            =   1590
               TabIndex        =   42
               Top             =   300
               Width           =   945
            End
            Begin VB.OptionButton Optb2c 
               BackColor       =   &H0080C0FF&
               Caption         =   "B2C Sales"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   270
               Left            =   45
               TabIndex        =   41
               Top             =   270
               Value           =   -1  'True
               Width           =   1485
            End
         End
         Begin VB.CheckBox CHKSELECT 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Caption         =   "Select All"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   9525
            TabIndex        =   39
            Top             =   570
            Width           =   1305
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0FF&
            Height          =   1680
            Left            =   7155
            TabIndex        =   31
            Top             =   870
            Visible         =   0   'False
            Width           =   3615
            Begin VB.OptionButton OptRT 
               BackColor       =   &H00C0C0FF&
               Caption         =   "RT"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   1950
               TabIndex        =   35
               Top             =   300
               Width           =   1635
            End
            Begin VB.OptionButton OptWS 
               BackColor       =   &H00C0C0FF&
               Caption         =   "WS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   135
               TabIndex        =   34
               Top             =   890
               Width           =   1635
            End
            Begin VB.OptionButton Optall 
               BackColor       =   &H00C0C0FF&
               Caption         =   "All"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   1950
               TabIndex        =   33
               Top             =   890
               Value           =   -1  'True
               Width           =   1620
            End
            Begin VB.OptionButton OptVan 
               BackColor       =   &H00C0C0FF&
               Caption         =   "VS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   135
               TabIndex        =   32
               Top             =   300
               Width           =   1635
            End
         End
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0C0FF&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   75
            TabIndex        =   24
            Top             =   420
            Value           =   -1  'True
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1860
            TabIndex        =   25
            Top             =   330
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   112263169
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   4035
            TabIndex        =   26
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   112263169
            CurrentDate     =   40498
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   405
            Width           =   285
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   6465
            TabIndex        =   28
            Top             =   855
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   8685
            TabIndex        =   27
            Top             =   285
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.CommandButton TMPDELETE 
         Caption         =   "DELETE"
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
         Left            =   5475
         TabIndex        =   3
         Top             =   7815
         Width           =   1350
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
         Height          =   435
         Left            =   8175
         TabIndex        =   2
         Top             =   7815
         Width           =   1290
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
         Height          =   435
         Left            =   6870
         TabIndex        =   1
         Top             =   7815
         Width           =   1260
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6300
         Left            =   165
         TabIndex        =   7
         Top             =   1125
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   11113
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1350
         Left            =   120
         TabIndex        =   4
         Top             =   7455
         Width           =   4995
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
            Left            =   1365
            TabIndex        =   21
            Top             =   930
            Width           =   1365
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
            Left            =   480
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   510
            Width           =   1155
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
         TabIndex        =   22
         Tag             =   "5"
         Top             =   7470
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   582
         Picture         =   "FrmBillTransfer.frx":0684
         ForeColor       =   0
         BarPicture      =   "FrmBillTransfer.frx":06A0
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
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Amt"
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
         Left            =   8160
         TabIndex        =   49
         Top             =   9045
         Width           =   1395
      End
      Begin VB.Label lblselamt 
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
         Left            =   9525
         TabIndex        =   48
         Top             =   9030
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FRMBillTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CMDCONVERT_Click()
    'If OptB2C.value = False Then Exit Sub
    If GRDTranx.rows <= 1 Then Exit Sub
    If grdcount.rows = 0 Then Exit Sub
    If grdcount.TextMatrix(0, 4) = "" Then Exit Sub
    If OptPetty.Value = True Then
        If MsgBox("ARE YOU SURE YOU WANT TO CONVERT THE SELECTED BILLS TO B2C", vbYesNo + vbDefaultButton2, "CONVERT.....") = vbNo Then Exit Sub
    Else
        If MsgBox("ARE YOU SURE YOU WANT TO CONVERT THE SELECTED BILLS TO PETTY", vbYesNo + vbDefaultButton2, "CONVERT.....") = vbNo Then Exit Sub
    End If
    
    
    Dim TRXMAST, TRXMASTWO As ADODB.Recordset
    Dim n, LASTWOBILL, LASTBILL As Long
    Dim M_DATE As Date
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    
'    LASTWOBILL = 0
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST_SP WHERE TRX_TYPE = 'WO'", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        LASTWOBILL = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
    Dim RSTTRXFILE, rstBILL As ADODB.Recordset
    Dim INVDETAILS As String
    LASTWOBILL = 0
    For n = 0 To grdcount.rows - 1
        If OptPetty.Value = True Then
            Set rstBILL = New ADODB.Recordset
            rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI'", db, adOpenStatic, adLockReadOnly
            If Not (rstBILL.EOF And rstBILL.BOF) Then
                LASTWOBILL = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
            End If
            rstBILL.Close
            Set rstBILL = Nothing
        Else
            Set rstBILL = New ADODB.Recordset
            rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO'", db, adOpenStatic, adLockReadOnly
            If Not (rstBILL.EOF And rstBILL.BOF) Then
                LASTWOBILL = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
            End If
            rstBILL.Close
            Set rstBILL = Nothing
        End If
        If OptB2C.Value = True Then
            db.BeginTrans
            INVDETAILS = Trim("No." & grdcount.TextMatrix(n, 4))
            db.Execute "Update TRXMAST set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO', REF_NO='" & Left(INVDETAILS, 20) & "' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update TRXSUB set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update TRXFILE set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO', INV_DETAILS='" & INVDETAILS & "' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            
            db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'HI'"
            db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTWOBILL & ", B_TRX_TYPE = 'WO' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_TRX_TYPE = 'HI' AND B_VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'HI'"
            db.Execute "Update CASHATRXFILE set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'HI'"
            
''''            db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='DR' AND INV_NO = " & grdcount.TextMatrix(N, 4) & " AND INV_TRX_TYPE = 'HI'"
''''            db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_VCH_NO = " & grdcount.TextMatrix(N, 4) & " AND B_TRX_TYPE = 'HI' "
''''            db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & grdcount.TextMatrix(N, 4) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'HI' "
''''            db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & grdcount.TextMatrix(N, 4) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'HI'"
        
            db.CommitTrans
        Else
            db.BeginTrans
            db.Execute "Update TRXMAST set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update TRXSUB set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update TRXFILE set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            
            db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'WO'"
            db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTWOBILL & ", B_TRX_TYPE = 'HI' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_TRX_TYPE = 'WO' AND B_VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'WO'"
            db.Execute "Update CASHATRXFILE set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'WO'"
            db.CommitTrans
        End If
    Next n

    
    '
    
    LASTBILL = 0
    Set TRXMAST = New ADODB.Recordset
    If OptPetty.Value = True Then
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_DATE = '" & Format(DateDiff("d", 1, DTFROM.Value), "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
    Else
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_DATE = '" & Format(DateDiff("d", 1, DTFROM.Value), "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
    End If
    'TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_TYPE = 'WO' ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        LASTBILL = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    LASTWOBILL = 0
    Set TRXMAST = New ADODB.Recordset
    If OptPetty.Value = True Then
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' ", db, adOpenStatic, adLockReadOnly
    Else
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' ", db, adOpenStatic, adLockReadOnly
    End If
    'TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_TYPE = 'WO' ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        LASTWOBILL = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    Dim startbill As Long
    startbill = 1
    Set TRXMAST = New ADODB.Recordset
    If OptPetty.Value = True Then
        TRXMAST.Open "Select MIN(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO > " & LASTBILL & " ", db, adOpenStatic, adLockReadOnly
    Else
        TRXMAST.Open "Select MIN(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO > " & LASTBILL & " ", db, adOpenStatic, adLockReadOnly
    End If
    'TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_TYPE = 'WO' ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        startbill = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    For n = startbill To LASTWOBILL
        LASTBILL = LASTBILL + 1
        If OptB2C.Value = True Then
            db.BeginTrans
            db.Execute "Update TRXMAST set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & startbill & ""
            db.Execute "Update TRXSUB set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & startbill & ""
            db.Execute "Update TRXFILE set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & startbill & ""
            
            db.Execute "Update DBTPYMT set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'HI'"
            db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTBILL & ", B_TRX_TYPE = 'HI' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_TRX_TYPE = 'HI' AND B_VCH_NO = " & startbill & ""
            db.Execute "Update DBTPYMT set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'HI'"
            db.Execute "Update CASHATRXFILE set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'HI'"
            db.CommitTrans
        Else
            db.BeginTrans
            db.Execute "Update TRXMAST set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & startbill & ""
            db.Execute "Update TRXSUB set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & startbill & ""
            db.Execute "Update TRXFILE set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & startbill & ""
            
            db.Execute "Update DBTPYMT set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'WO'"
            db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTBILL & ", B_TRX_TYPE = 'WO' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_TRX_TYPE = 'WO' AND B_VCH_NO = " & startbill & ""
            db.Execute "Update DBTPYMT set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'WO'"
            db.Execute "Update CASHATRXFILE set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'WO'"
            db.CommitTrans
        End If
        startbill = startbill + 1
    Next n
    
    
'    N = LASTBILL
'    Set TRXMASTWO = New ADODB.Recordset
'    If Optpetty.value = True Then
'        TRXMASTWO.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='WO' AND VCH_NO > " & LASTBILL & " ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
'    Else
'        TRXMASTWO.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='HI' AND VCH_NO > " & LASTBILL & " ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
'    End If
'    Do Until TRXMASTWO.EOF
'        N = N + 1
'        Set TRXMAST = New ADODB.Recordset
'        If Optpetty.value = True Then
'            TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
'        Else
'            TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
'        End If
'        Do Until TRXMAST.EOF
'            TRXMAST!VCH_NO = N
'            TRXMAST.Update
'            TRXMAST.MoveNext
'        Loop
'        TRXMAST.Close
'        Set TRXMAST = Nothing
'
'        Set TRXMAST = New ADODB.Recordset
'        If Optpetty.value = True Then
'            TRXMAST.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
'        Else
'            TRXMAST.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
'        End If
'        Do Until TRXMAST.EOF
'            TRXMAST!VCH_NO = N
'            TRXMAST.Update
'            TRXMAST.MoveNext
'        Loop
'        TRXMAST.Close
'        Set TRXMAST = Nothing
'
'        Set TRXMAST = New ADODB.Recordset
'        If Optpetty.value = True Then
'            TRXMAST.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
'        Else
'            TRXMAST.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
'        End If
'        Do Until TRXMAST.EOF
'            TRXMAST!VCH_NO = N
'            TRXMAST.Update
'            TRXMAST.MoveNext
'        Loop
'        TRXMAST.Close
'        Set TRXMAST = Nothing
'
'        TRXMASTWO.MoveNext
'    Loop
'    TRXMASTWO.Close
'    Set TRXMASTWO = Nothing
        
    Call CmDDisplay_Click
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub CmdConvert2_Click()
    'If OptB2C.value = False Then Exit Sub
    If GRDTranx.rows <= 1 Then Exit Sub
    If grdcount.rows = 0 Then Exit Sub
    If grdcount.TextMatrix(0, 4) = "" Then Exit Sub
    If OptPetty.Value = True Then
        If MsgBox("ARE YOU SURE YOU WANT TO CONVERT THE SELECTED BILLS TO B2C", vbYesNo + vbDefaultButton2, "CONVERT.....") = vbNo Then Exit Sub
    Else
        If MsgBox("ARE YOU SURE YOU WANT TO CONVERT THE SELECTED BILLS TO PETTY", vbYesNo + vbDefaultButton2, "CONVERT.....") = vbNo Then Exit Sub
    End If
    
    
    Dim TRXMAST, TRXMASTWO As ADODB.Recordset
    Dim n, LASTWOBILL, LASTBILL As Long
    Dim M_DATE As Date
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    
'    LASTWOBILL = 0
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST_SP WHERE TRX_TYPE = 'WO'", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        LASTWOBILL = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
    Dim RSTTRXFILE, rstBILL As ADODB.Recordset
    Dim INVDETAILS As String
    LASTWOBILL = 0
    For n = 0 To grdcount.rows - 1
        If OptPetty.Value = True Then
            Set rstBILL = New ADODB.Recordset
            rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI'", db, adOpenStatic, adLockReadOnly
            If Not (rstBILL.EOF And rstBILL.BOF) Then
                LASTWOBILL = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
            End If
            rstBILL.Close
            Set rstBILL = Nothing
        Else
            Set rstBILL = New ADODB.Recordset
            rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO'", db, adOpenStatic, adLockReadOnly
            If Not (rstBILL.EOF And rstBILL.BOF) Then
                LASTWOBILL = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
            End If
            rstBILL.Close
            Set rstBILL = Nothing
        End If
        If OptB2C.Value = True Then
            db.BeginTrans
            INVDETAILS = Trim("No." & grdcount.TextMatrix(n, 4))
            db.Execute "Update TRXMAST set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO', REF_NO='" & Left(INVDETAILS, 20) & "' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update TRXSUB set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update TRXFILE set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO', INV_DETAILS='" & INVDETAILS & "' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            
            db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'HI'"
            db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTWOBILL & ", B_TRX_TYPE = 'WO' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_TRX_TYPE = 'HI' AND B_VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'HI'"
            db.Execute "Update CASHATRXFILE set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'HI'"
            
''''            db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='DR' AND INV_NO = " & grdcount.TextMatrix(N, 4) & " AND INV_TRX_TYPE = 'HI'"
''''            db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_VCH_NO = " & grdcount.TextMatrix(N, 4) & " AND B_TRX_TYPE = 'HI' "
''''            db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & grdcount.TextMatrix(N, 4) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'HI' "
''''            db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_NO = " & grdcount.TextMatrix(N, 4) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'HI'"
        
            db.CommitTrans
        Else
            db.BeginTrans
            db.Execute "Update TRXMAST set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update TRXSUB set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update TRXFILE set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            
            db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'WO'"
            db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTWOBILL & ", B_TRX_TYPE = 'HI' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_TRX_TYPE = 'WO' AND B_VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'WO'"
            db.Execute "Update CASHATRXFILE set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'WO'"
            db.CommitTrans
        End If
    Next n

    
    '
    
    LASTBILL = 0
    Set TRXMAST = New ADODB.Recordset
    If OptPetty.Value = True Then
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_DATE <= '" & Format(DateDiff("d", 1, DTFROM.Value), "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
    Else
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_DATE <= '" & Format(DateDiff("d", 1, DTFROM.Value), "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
    End If
    'TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_TYPE = 'WO' ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        LASTBILL = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    LASTWOBILL = 0
    Set TRXMAST = New ADODB.Recordset
    If OptPetty.Value = True Then
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' ", db, adOpenStatic, adLockReadOnly
    Else
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' ", db, adOpenStatic, adLockReadOnly
    End If
    'TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_TYPE = 'WO' ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        LASTWOBILL = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    Dim startbill As Long
    startbill = 1
    Set TRXMAST = New ADODB.Recordset
    If OptPetty.Value = True Then
        TRXMAST.Open "Select MIN(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_NO > " & LASTBILL & " ", db, adOpenStatic, adLockReadOnly
    Else
        TRXMAST.Open "Select MIN(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_NO > " & LASTBILL & " ", db, adOpenStatic, adLockReadOnly
    End If
    'TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_TYPE = 'WO' ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        startbill = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
'    For N = startbill To LASTWOBILL
'        LASTBILL = LASTBILL + 1
'        If OptB2C.value = True Then
'            db.BeginTrans
'            db.Execute "Update TRXMAST set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & startbill & ""
'            db.Execute "Update TRXSUB set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & startbill & ""
'            db.Execute "Update TRXFILE set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & startbill & ""
'
'            db.Execute "Update DBTPYMT set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'HI'"
'            db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTBILL & ", B_TRX_TYPE = 'HI' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_TRX_TYPE = 'HI' AND B_VCH_NO = " & startbill & ""
'            db.Execute "Update DBTPYMT set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'HI'"
'            db.Execute "Update CASHATRXFILE set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'HI' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'HI'"
'            db.CommitTrans
'        Else
'            db.BeginTrans
'            db.Execute "Update TRXMAST set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & startbill & ""
'            db.Execute "Update TRXSUB set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & startbill & ""
'            db.Execute "Update TRXFILE set VCH_NO = " & LASTBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'WO' AND VCH_NO = " & startbill & ""
'
'            db.Execute "Update DBTPYMT set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'WO'"
'            db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTBILL & ", B_TRX_TYPE = 'WO' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_TRX_TYPE = 'WO' AND B_VCH_NO = " & startbill & ""
'            db.Execute "Update DBTPYMT set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'WO'"
'            db.Execute "Update CASHATRXFILE set INV_NO = " & LASTBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & startbill & " AND INV_TRX_TYPE = 'WO'"
'            db.CommitTrans
'        End If
'        startbill = startbill + 1
'    Next N
    
    
    n = LASTBILL
    Set TRXMASTWO = New ADODB.Recordset
    If OptPetty.Value = True Then
        TRXMASTWO.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO > " & LASTBILL & " ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        TRXMASTWO.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO > " & LASTBILL & " ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
    Do Until TRXMASTWO.EOF
        n = n + 1
        Set TRXMAST = New ADODB.Recordset
        If OptPetty.Value = True Then
            TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        Else
            TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        End If
        If Not (TRXMAST.EOF And TRXMAST.BOF) Then
            db.Execute "Update DBTPYMT set INV_NO = " & n & " WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE = 'DR' AND INV_NO = " & TRXMAST!VCH_NO & " AND INV_TRX_TYPE = '" & TRXMAST!TRX_TYPE & "'"
            db.Execute "Update BANK_TRX set B_VCH_NO = " & n & " WHERE B_TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND B_TRX_TYPE = '" & TRXMAST!TRX_TYPE & "' AND B_VCH_NO = " & TRXMAST!VCH_NO & ""
            db.Execute "Update DBTPYMT set INV_NO = " & n & " WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE = 'RT' AND INV_NO = " & TRXMAST!VCH_NO & " AND INV_TRX_TYPE = '" & TRXMAST!TRX_TYPE & "'"
            db.Execute "Update CASHATRXFILE set INV_NO = " & n & " WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND INV_TYPE = 'RT' AND INV_NO = " & TRXMAST!VCH_NO & " AND INV_TRX_TYPE = '" & TRXMAST!TRX_TYPE & "'"

            TRXMAST!VCH_NO = n
            TRXMAST.Update
        End If
        TRXMAST.Close
        Set TRXMAST = Nothing

        Set TRXMAST = New ADODB.Recordset
        If OptPetty.Value = True Then
            TRXMAST.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        Else
            TRXMAST.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        End If
        Do Until TRXMAST.EOF
            TRXMAST!VCH_NO = n
            TRXMAST.Update
            TRXMAST.MoveNext
        Loop
        TRXMAST.Close
        Set TRXMAST = Nothing

        Set TRXMAST = New ADODB.Recordset
        If OptPetty.Value = True Then
            TRXMAST.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        Else
            TRXMAST.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        End If
        Do Until TRXMAST.EOF
            TRXMAST!VCH_NO = n
            TRXMAST.Update
            TRXMAST.MoveNext
        Loop
        TRXMAST.Close
        Set TRXMAST = Nothing

        TRXMASTWO.MoveNext
    Loop
    TRXMASTWO.Close
    Set TRXMASTWO = Nothing
        
    Call CmDDisplay_Click
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub CmDDisplay_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    
    LBLTRXTOTAL.Caption = "0.00"
    LBLDISCOUNT.Caption = "0.00"
    LBLNET.Caption = "0.00"
    lblselAMT.Caption = ""
    'LBLCOST.Caption = "0.00"
    'LBLPROFIT.Caption = "0.00"
    'lblcommi.Caption = "0.00"
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    n = 1
    M = 0
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    If OptPetty.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='WO' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If

    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = ""
        GRDTranx.TextMatrix(M, 1) = M
        GRDTranx.TextMatrix(M, 2) = rstTRANX!TRX_TYPE
        GRDTranx.TextMatrix(M, 3) = "Sale"
        GRDTranx.TextMatrix(M, 4) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 5) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 6) = Format(Round(rstTRANX!VCH_AMOUNT, 2), "0.00")
        GRDTranx.TextMatrix(M, 7) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(rstTRANX!DISCOUNT, "0.00"))
        GRDTranx.TextMatrix(M, 8) = Format(Round(Val(GRDTranx.TextMatrix(M, 6)) - Val(GRDTranx.TextMatrix(M, 7)), 2), "0.00")
       
        CMDDISPLAY.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
        
        GRDTranx.TextMatrix(M, 9) = "N"
        'GRDTranx.TextMatrix(i, 14) = !LINE_NO
        With GRDTranx
          .Row = M: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
          Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
        End With
            
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption) + Val(GRDTranx.TextMatrix(M, 7)), "0.00")
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        n = n + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
    
    LBLNET.Caption = Format(Val(LBLTRXTOTAL.Caption) - Val(LBLDISCOUNT.Caption), "0.00")
    
    TMPDELETE.Enabled = True
    CMDCONVERT.Enabled = True
    CMDCONVERT2.Enabled = True
    flagchange.Caption = ""
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
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

Private Sub Command1_Click()
    'If grdcount.Rows = 0 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE UNBILL ITEMS FROM B2B & B2C OF SELECTED PERIOD", vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then Exit Sub
    
    Dim TOTAL_AMT As Double
    Dim TRXMAST  As ADODB.Recordset
    Dim n, LASTWOBILL, LASTBILL As Long
    Dim M_DATE As Date
    Dim crt_flag As Boolean
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    
    Dim RSTTRXFILE, TRXFILEMAST As ADODB.Recordset
        
    Set TRXMAST = New ADODB.Recordset
    'rstTRXMAST.Open "Select * From RTRXFILE LEFT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE RTRXFILE.BARCODE= '" & Trim(TxtBarcode.Text) & "' AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    'TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND (TRX_TYPE='HI' OR TRX_TYPE='HI') AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
    'TRXMAST.Open "Select * FROM TRXMAST LEFT JOIN TRXSUB ON TRXMAST.TRX_YEAR = TRXSUB.TRX_YEAR AND TRXMAST.TRX_TYPE = TRXSUB.TRX_TYPE AND TRXMAST.VCH_NO = TRXSUB.VCH_NO WHERE TRXMAST.TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND (TRXMAST.TRX_TYPE='HI' OR TRXMAST.TRX_TYPE='HI') AND TRXMAST.VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND TRXMAST.VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
    TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND (TRX_TYPE='HI' OR TRX_TYPE='GI') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
    Do Until TRXMAST.EOF
        crt_flag = False
        Set TRXFILEMAST = New ADODB.Recordset
        TRXFILEMAST.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE='" & TRXMAST!TRX_TYPE & "' AND VCH_NO = " & TRXMAST!VCH_NO & " AND UN_BILL = 'Y' ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
        Do Until TRXFILEMAST.EOF
            
'            Set RSTTRXFILE = New ADODB.Recordset
'            RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & TRXMAST!R_TRX_TYPE & "' AND RTRXFILE.VCH_NO = " & TRXMAST!R_VCH_NO & " AND RTRXFILE.LINE_NO = " & TRXMAST!R_LINE_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
'            With RSTTRXFILE
'                If Not (.EOF And .BOF) Then
'                    If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
'                    If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
'                    !ISSUE_QTY = !ISSUE_QTY - TRXFILEMAST!QTY
'                    !BAL_QTY = !BAL_QTY + TRXFILEMAST!QTY
'                    RSTTRXFILE.Update
'                End If
'            End With
'            RSTTRXFILE.Close
'            Set RSTTRXFILE = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TRXFILEMAST!ITEM_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            With RSTTRXFILE
                If Not (.EOF And .BOF) Then
                    !ISSUE_QTY = !ISSUE_QTY - TRXFILEMAST!QTY
                    If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                    !ISSUE_VAL = !ISSUE_VAL - TRXFILEMAST!TRX_TOTAL
                    !CLOSE_QTY = !CLOSE_QTY + TRXFILEMAST!QTY
                    If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                    !CLOSE_VAL = !CLOSE_VAL + TRXFILEMAST!TRX_TOTAL
                    RSTTRXFILE.Update
                End If
            End With
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
                
            db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & TRXFILEMAST!TRX_YEAR & "' AND TRX_TYPE='" & TRXFILEMAST!TRX_TYPE & "' AND VCH_NO = " & TRXFILEMAST!VCH_NO & " AND LINE_NO = " & TRXFILEMAST!LINE_NO & " "
            db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & TRXFILEMAST!TRX_YEAR & "' AND TRX_TYPE='" & TRXFILEMAST!TRX_TYPE & "' AND VCH_NO = " & TRXFILEMAST!VCH_NO & " AND LINE_NO = " & TRXFILEMAST!LINE_NO & " AND UN_BILL = 'Y'"
            crt_flag = True
            TRXFILEMAST.MoveNext
        Loop
        TRXFILEMAST.Close
        Set TRXFILEMAST = Nothing
        
        If crt_flag = True Then
            TOTAL_AMT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT SUM(TRX_TOTAL) FROM TRXFILE WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE='" & TRXMAST!TRX_TYPE & "' AND VCH_NO = " & TRXMAST!VCH_NO & " ", db, adOpenForwardOnly
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                TOTAL_AMT = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                db.Execute "Update TRXMAST SET NET_AMOUNT = " & TOTAL_AMT & ", VCH_AMOUNT = " & TOTAL_AMT & " WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE='" & TRXMAST!TRX_TYPE & "' AND VCH_NO = " & TRXMAST!VCH_NO & " "
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        End If
        TRXMAST.MoveNext
    Loop
    TRXMAST.Close
    Set TRXMAST = Nothing
        
    Call CmDDisplay_Click
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Command2_Click()
    If grdcount.rows = 0 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO MOVE UNBILL ITEMS FROM B2B & B2C OF SELECTED PERIOD", vbYesNo + vbDefaultButton2, "MOVE UN BILL ITEMS.....") = vbNo Then Exit Sub
    
    Dim TOTAL_AMT As Double
    Dim TRXMAST  As ADODB.Recordset
    Dim n, LASTBILL As Long
    Dim M_DATE As Date
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXFILEMAST As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim LASTWOBILL As Long
    Dim CRNO2 As Double
    Dim INVDETAILS As String
    Dim i As Integer
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND (TRX_TYPE='HI' OR TRX_TYPE='GI') AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
    Do Until TRXMAST.EOF
        Set rstBILL = New ADODB.Recordset
        rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO'", db, adOpenStatic, adLockReadOnly
        If Not (rstBILL.EOF And rstBILL.BOF) Then
            LASTWOBILL = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        End If
        rstBILL.Close
        Set rstBILL = Nothing
        
        i = 1
        Set TRXFILEMAST = New ADODB.Recordset
        TRXFILEMAST.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE='" & TRXMAST!TRX_TYPE & "' AND VCH_NO = " & TRXMAST!VCH_NO & " AND UN_BILL = 'Y' ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
        Do Until TRXFILEMAST.EOF
            db.Execute "Update TRXSUB set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO', LINE_NO = " & i & ", TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' WHERE TRX_YEAR= '" & TRXFILEMAST!TRX_YEAR & "' AND TRX_TYPE = '" & TRXFILEMAST!TRX_TYPE & "' AND VCH_NO = " & TRXFILEMAST!VCH_NO & " AND LINE_NO = " & TRXFILEMAST!LINE_NO & ""
            INVDETAILS = ""
            Select Case TRXMAST!TRX_TYPE
                Case "GI"
                    INVDETAILS = Left(Trim("B2B-" & TRXMAST!VCH_NO & IIf(IsDate(TRXMAST!VCH_DATE), " DTD " & TRXMAST!VCH_DATE, "")), 50)
                Case Else
                    INVDETAILS = Left(Trim("No." & TRXMAST!VCH_NO & IIf(IsDate(TRXMAST!VCH_DATE), " DTD " & TRXMAST!VCH_DATE, "")), 50)
            End Select
            db.Execute "Update TRXFILE set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO', LINE_NO = " & i & ", TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "', INV_DETAILS='" & INVDETAILS & "' WHERE TRX_YEAR= '" & TRXFILEMAST!TRX_YEAR & "' AND TRX_TYPE = '" & TRXFILEMAST!TRX_TYPE & "' AND VCH_NO = " & TRXFILEMAST!VCH_NO & " AND LINE_NO = " & TRXFILEMAST!LINE_NO & ""
            i = i + 1
            TRXFILEMAST.MoveNext
        Loop
        TRXFILEMAST.Close
        Set TRXFILEMAST = Nothing
        
        If i > 1 Then
            TOTAL_AMT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT SUM(TRX_TOTAL) FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE= 'WO' AND VCH_NO = " & LASTWOBILL & " ", db, adOpenForwardOnly
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                TOTAL_AMT = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                'db.Execute "Update TRXMAST SET NET_AMOUNT = " & TOTAL_AMT & ", VCH_AMOUNT = " & TOTAL_AMT & " WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE='" & TRXMAST!TRX_TYPE & "' AND VCH_NO = " & TRXMAST!VCH_NO & " "
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * FROM TRXMAST WHERE VCH_NO = " & LASTWOBILL & " AND TRX_TYPE = 'WO' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE.AddNew
                RSTTRXFILE!VCH_NO = LASTWOBILL
                RSTTRXFILE!TRX_TYPE = "WO"
                RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
                RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
                RSTTRXFILE!C_TIME = Format(Time, "HH:MM:SS")
                RSTTRXFILE!C_USER_NAME = frmLogin.rs!USER_NAME
            End If
            Select Case TRXMAST!TRX_TYPE
                Case "GI"
                    RSTTRXFILE!REF_NO = Left(Trim("B2B-" & TRXMAST!VCH_NO & IIf(IsDate(TRXMAST!VCH_DATE), " DTD " & TRXMAST!VCH_DATE, "")), 20)
                Case Else
                    RSTTRXFILE!REF_NO = Left(Trim(TRXMAST!VCH_NO & IIf(IsDate(TRXMAST!VCH_DATE), " DTD " & TRXMAST!VCH_DATE, "")), 20)
            End Select
            RSTTRXFILE!TIN = TRXMAST!TIN
            RSTTRXFILE!UID_NO = TRXMAST!UID_NO
            RSTTRXFILE!CUST_IGST = TRXMAST!CUST_IGST
            RSTTRXFILE!VCH_AMOUNT = TOTAL_AMT
            RSTTRXFILE!NET_AMOUNT = TOTAL_AMT
            RSTTRXFILE!VCH_DATE = TRXMAST!VCH_DATE
            RSTTRXFILE!ACT_CODE = TRXMAST!ACT_CODE
            RSTTRXFILE!ACT_NAME = TRXMAST!ACT_NAME
            RSTTRXFILE!DISCOUNT = TRXMAST!DISCOUNT
            RSTTRXFILE!DISC_PERS = TRXMAST!DISC_PERS
            RSTTRXFILE!BILL_NAME = TRXMAST!BILL_NAME
            RSTTRXFILE!BILL_ADDRESS = TRXMAST!BILL_ADDRESS
            RSTTRXFILE!ADD_AMOUNT = TRXMAST!ADD_AMOUNT
            RSTTRXFILE!ROUNDED_OFF = TRXMAST!ROUNDED_OFF
            RSTTRXFILE!PAY_AMOUNT = TRXMAST!PAY_AMOUNT
            RSTTRXFILE!Area = TRXMAST!Area
            RSTTRXFILE!CN_REF = TRXMAST!CN_REF
            RSTTRXFILE!BILL_FLAG = TRXMAST!BILL_FLAG
            RSTTRXFILE!TERMS = TRXMAST!TERMS
            RSTTRXFILE!BR_CODE = TRXMAST!BR_CODE
            RSTTRXFILE!BR_NAME = TRXMAST!BR_NAME
            RSTTRXFILE!cr_days = TRXMAST!cr_days
            RSTTRXFILE!BILL_TYPE = TRXMAST!BILL_TYPE
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
'            CRNO2 = 1
'            Set rstMaxRec = New ADODB.Recordset
'            rstMaxRec.Open "Select MAX(CR_NO) From DBTPYMT", db, adOpenForwardOnly
'            If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
'                CRNO2 = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
'            End If
'            rstMaxRec.Close
'            Set rstMaxRec = Nothing
'
'            Set RSTTRXFILE = New ADODB.Recordset
'            RSTTRXFILE.Open "SELECT * FROM DBTPYMT WHERE INV_NO = " & LASTWOBILL & " AND INV_TRX_TYPE = 'WO' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'DR' ", db, adOpenStatic, adLockOptimistic, adCmdText
'            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'                RSTTRXFILE.AddNew
'                RSTTRXFILE!TRX_TYPE = "DR"
'                RSTTRXFILE!INV_TRX_TYPE = "WO"
'                RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'                RSTTRXFILE!CR_NO = CRNO2
'                RSTTRXFILE!INV_NO = LASTWOBILL
'            End If
'            RSTTRXFILE!ACT_CODE = TRXMAST!ACT_CODE
'            RSTTRXFILE!ACT_NAME = TRXMAST!ACT_NAME
'            RSTTRXFILE!INV_DATE = TRXMAST!VCH_DATE
'            RSTTRXFILE!INV_AMT = Val(lblnetamount.Caption)
'            RSTTRXFILE!BR_ADDRESS = TRXMAST!BILL_ADDRESS
'            RSTTRXFILE!COMM_AMT = 0
'            RSTTRXFILE!PYMT_PERIOD = TRXMAST!cr_days
'            RSTTRXFILE.Update
'            RSTTRXFILE.Close
'            Set RSTTRXFILE = Nothing
            
'            Set rstMaxRec = New ADODB.Recordset
'            rstMaxRec.Open "Select MAX(CR_NO) From DBTPYMT", db, adOpenForwardOnly
'            If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
'                CRNO2 = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
'            End If
'            rstMaxRec.Close
'            Set rstMaxRec = Nothing
'
'            Set RSTTRXFILE = New ADODB.Recordset
'            RSTTRXFILE.Open "SELECT * FROM DBTPYMT WHERE INV_NO = " & LASTWOBILL & " AND INV_TRX_TYPE = 'WO' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'RT' ", db, adOpenStatic, adLockOptimistic, adCmdText
'            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'                RSTTRXFILE.AddNew
'                RSTTRXFILE!TRX_TYPE = "RT"
'                RSTTRXFILE!INV_TRX_TYPE = "WO"
'                RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'                RSTTRXFILE!CR_NO = CRNO2
'                RSTTRXFILE!INV_NO = LASTWOBILL
'            End If
'            RSTTRXFILE!ACT_CODE = TRXMAST!ACT_CODE
'            RSTTRXFILE!ACT_NAME = TRXMAST!ACT_NAME
'            RSTTRXFILE!INV_DATE = TRXMAST!VCH_DATE
'            RSTTRXFILE!INV_AMT = Val(lblnetamount.Caption)
'            RSTTRXFILE!BR_ADDRESS = TRXMAST!BILL_ADDRESS
'            RSTTRXFILE!COMM_AMT = 0
'            RSTTRXFILE!PYMT_PERIOD = TRXMAST!cr_days
'            RSTTRXFILE.Update
'            RSTTRXFILE.Close
'            Set RSTTRXFILE = Nothing
                
            Set rstMaxRec = New ADODB.Recordset
            rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
            If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
                CRNO2 = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
            End If
            rstMaxRec.Close
            Set rstMaxRec = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * FROM CASHATRXFILE WHERE INV_NO = " & LASTWOBILL & " AND INV_TRX_TYPE = 'WO' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_TYPE = 'RT' ", db, adOpenStatic, adLockOptimistic, adCmdText
            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE.AddNew
                RSTTRXFILE!INV_TYPE = "RT"
                RSTTRXFILE!INV_TRX_TYPE = "WO"
                RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                RSTTRXFILE!REC_NO = CRNO2
                RSTTRXFILE!INV_NO = LASTWOBILL
            End If
            RSTTRXFILE!AMOUNT = TOTAL_AMT
            RSTTRXFILE!TRX_TYPE = "CR"
            RSTTRXFILE!check_flag = "S"
            RSTTRXFILE!ACT_CODE = TRXMAST!ACT_CODE
            RSTTRXFILE!ACT_NAME = TRXMAST!ACT_NAME
            RSTTRXFILE!VCH_DATE = TRXMAST!VCH_DATE
            RSTTRXFILE!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
        End If
        
'        db.BeginTrans
'        db.Execute "Update TRXMAST set VCH_NO = " & LASTWOBILL & ", TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'HI' AND VCH_NO = " & Val(txtBillNo.Text) & ""
'
'        db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'DR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TRX_TYPE = 'HI'"
'        db.Execute "Update BANK_TRX set B_VCH_NO = " & LASTWOBILL & ", B_TRX_TYPE = 'WO' WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_TRX_TYPE = 'HI' AND B_VCH_NO = " & Val(txtBillNo.Text) & ""
'        db.Execute "Update DBTPYMT set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = 'RT' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TRX_TYPE = 'HI'"
'        db.Execute "Update CASHATRXFILE set INV_NO = " & LASTWOBILL & ", INV_TRX_TYPE = 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND INV_TYPE = 'RT' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TRX_TYPE = 'HI'"
'        db.Execute "Update RTRXFILE set TRX_TYPE= 'WO' WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= 'HI' AND VCH_NO = " & Val(TxtCN.Text) & ""

'        db.CommitTrans
        
        If i > 1 Then
            TOTAL_AMT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT SUM(TRX_TOTAL) FROM TRXFILE WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE='" & TRXMAST!TRX_TYPE & "' AND VCH_NO = " & TRXMAST!VCH_NO & " ", db, adOpenForwardOnly
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                TOTAL_AMT = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
                db.Execute "Update TRXMAST SET NET_AMOUNT = " & TOTAL_AMT & ", VCH_AMOUNT = " & TOTAL_AMT & " WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE='" & TRXMAST!TRX_TYPE & "' AND VCH_NO = " & TRXMAST!VCH_NO & " "
                db.Execute "Update DBTPYMT set INV_AMT = " & TOTAL_AMT & " WHERE TRX_YEAR='" & TRXMAST!TRX_YEAR & "' AND TRX_TYPE = 'DR' AND INV_NO = " & TRXMAST!VCH_NO & " AND INV_TRX_TYPE = '" & TRXMAST!TRX_TYPE & "'"
                db.Execute "Update BANK_TRX set TRX_AMOUNT = " & TOTAL_AMT & " WHERE B_TRX_YEAR= '" & TRXMAST!TRX_YEAR & "' AND B_TRX_TYPE = '" & TRXMAST!TRX_TYPE & "' AND B_VCH_NO = " & TRXMAST!VCH_NO & ""
                db.Execute "Update CASHATRXFILE set AMOUNT = " & TOTAL_AMT & " WHERE TRX_YEAR= '" & TRXMAST!TRX_YEAR & "' AND INV_TYPE = 'RT' AND INV_NO = " & TRXMAST!VCH_NO & " AND INV_TRX_TYPE = '" & TRXMAST!TRX_TYPE & "'"
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        End If
        
        TRXMAST.MoveNext
    Loop
    TRXMAST.Close
    Set TRXMAST = Nothing
        
    Call CmDDisplay_Click
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub DTFROM_Change()
    TMPDELETE.Enabled = False
    CMDCONVERT.Enabled = False
    CMDCONVERT2.Enabled = False
End Sub

'Private Sub Command2_Click()
'    Dim RSTITEMMAST As ADODB.Recordset
'    If MsgBox("Are You Sure You want to Delete All Bills", vbYesNo, "DELETING BILL....") = vbNo Then Exit Sub
'    db.Execute "delete From TRXFILE"
'    db.Execute "delete From TRXMAST"
'
'    db.Execute "delete From RTRXFILE WHERE TRX_TYPE='OP' OR TRX_TYPE='PI' "
'    db.Execute "delete From TRANSMAST WHERE TRX_TYPE='OP' OR TRX_TYPE='OP' "
'
'    On Error GoTo eRRHAND
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM ITEMMAST Order by ITEM_CODE", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until RSTITEMMAST.EOF
'        RSTITEMMAST!OPEN_QTY = 0
'        RSTITEMMAST!OPEN_VAL = 0
'        RSTITEMMAST!RCPT_QTY = 0
'        RSTITEMMAST!RCPT_VAL = 0
'        RSTITEMMAST!ISSUE_QTY = 0
'        RSTITEMMAST!ISSUE_VAL = 0
'        RSTITEMMAST!CLOSE_QTY = 0
'        RSTITEMMAST!CLOSE_VAL = 0
'        RSTITEMMAST!DAM_QTY = 0
'        RSTITEMMAST!DAM_VAL = 0
'        RSTITEMMAST!DISC = 0
'        RSTITEMMAST!SALES_PRICE = 0
'        RSTITEMMAST!RCVD_NOS = 0
'        RSTITEMMAST!ISSUE_NOS = 0
'        RSTITEMMAST!BAL_NOS = 0
'        RSTITEMMAST.Update
'        RSTITEMMAST.MoveNext
'    Loop
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
'
'    Exit Sub
'eRRHAND:
'    MsgBox Err.Description
'End Sub

Private Sub DTFROM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DTTo.SetFocus
    End Select
End Sub

Private Sub DTTO_Change()
    TMPDELETE.Enabled = False
    CMDCONVERT.Enabled = False
    CMDCONVERT2.Enabled = False
End Sub

Private Sub DTTO_GotFocus()
    'CMDPRINTREGISTER.Enabled = False
End Sub

Private Sub DTTO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            DTFROM.SetFocus
    End Select
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'If Month(Date) > 1 Then
        'CMBMONTH.ListIndex = Month(Date) - 2
    'Else
        'CMBMONTH.ListIndex = 11
    'End If
    
    If RstBill_Flag = "Y" Then
        CMDCONVERT2.Visible = True
        CMDCONVERT.Visible = False
    Else
        CMDCONVERT.Visible = True
        CMDCONVERT2.Visible = False
    End If
    GRDTranx.TextMatrix(0, 0) = ""
    GRDTranx.TextMatrix(0, 1) = "SL"
    GRDTranx.TextMatrix(0, 2) = ""
    GRDTranx.TextMatrix(0, 3) = ""
    GRDTranx.TextMatrix(0, 4) = "BILL NO"
    GRDTranx.TextMatrix(0, 5) = "BILL DATE"
    GRDTranx.TextMatrix(0, 6) = "BILL AMT"
    GRDTranx.TextMatrix(0, 7) = "DISC AMT"
    GRDTranx.TextMatrix(0, 8) = "NET AMT"
    GRDTranx.TextMatrix(0, 9) = ""
        
    GRDTranx.ColWidth(0) = 300
    GRDTranx.ColWidth(2) = 0
    GRDTranx.ColWidth(3) = 0
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 1500
    GRDTranx.ColWidth(6) = 2200
    GRDTranx.ColWidth(7) = 1500
    GRDTranx.ColWidth(8) = 2200
    GRDTranx.ColWidth(9) = 0
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 2
    GRDTranx.ColAlignment(3) = 2
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    GRDTranx.ColAlignment(7) = 4
    GRDTranx.ColAlignment(8) = 4
    GRDTranx.ColAlignment(9) = 4
    
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "Rate"
    GRDBILL.TextMatrix(0, 3) = "Disc %"
    GRDBILL.TextMatrix(0, 4) = "Tax %"
    GRDBILL.TextMatrix(0, 5) = "Qty"
    GRDBILL.TextMatrix(0, 6) = "Amount"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 2800
    GRDBILL.ColWidth(2) = 800
    GRDBILL.ColWidth(3) = 800
    GRDBILL.ColWidth(4) = 800
    GRDBILL.ColWidth(5) = 1500
    GRDBILL.ColWidth(6) = 2000
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 3
    GRDBILL.ColAlignment(6) = 6
    
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

Private Sub GRDTranx_Click()
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    If GRDTranx.rows = 1 Then Exit Sub
    If GRDTranx.Col <> 0 Then Exit Sub
    With GRDTranx
        oldx = .Col
        oldy = .Row
        .Row = oldy: .Col = 0: .CellPictureAlignment = 4
            'If GRDTranx.Col = 0 Then
                If GRDTranx.CellPicture = picChecked Then
                    Set GRDTranx.CellPicture = picUnchecked
                    '.Col = .Col + 2  ' I use data that is in column #1, usually an Index or ID #
                    'strTextCheck = .Text
                    ' When you de-select a CheckBox, we need to strip out the #
                    'strChecked = strChecked & strTextCheck & ","
                    ' Don't forget to strip off the trailing , before passing the string
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 9) = "Y"
                    lblselAMT.Caption = Val(lblselAMT.Caption) + Val(GRDTranx.TextMatrix(.Row, 8))
                    Call fillcount
                Else
                    Set GRDTranx.CellPicture = picChecked
                    '.Col = .Col + 2
                    'strTextCheck = .Text
                    'strChecked = Replace(strChecked, strTextCheck & ",", "")
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 9) = "N"
                    lblselAMT.Caption = Val(lblselAMT.Caption) - Val(GRDTranx.TextMatrix(.Row, 8))
                    Call fillcount
                End If
            'End If
        .Col = oldx
        .Row = oldy
    End With
End Sub

Private Sub CHKSELECT_Click()
    Dim i As Long
    If GRDTranx.rows = 1 Then Exit Sub
    lblselAMT.Caption = ""
    For i = 1 To GRDTranx.rows - 1
        If CHKSELECT.Value = 1 Then
            With GRDTranx
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picUnchecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
              .TextMatrix(.Row, 9) = "Y"
            End With
            lblselAMT.Caption = Val(lblselAMT.Caption) + Val(GRDTranx.TextMatrix(i, 8))
        Else
            With GRDTranx
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
              .TextMatrix(.Row, 9) = "N"
            End With
        End If
    Next i
    Call fillcount
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.rows = 1 Then Exit Sub
            LBLBILLNO.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            LBLBILLAMT.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 6), "0.00")
            LBLDISC.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 7), "0.00")
            LBLNETAMT.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 8), "0.00")
             
            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILE WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "  AND TRX_TYPE = '" & Trim(GRDTranx.TextMatrix(GRDTranx.Row, 2)) & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.rows = GRDBILL.rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ITEM_NAME
                GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!SALES_PRICE, "0.00")
                GRDBILL.TextMatrix(i, 3) = Val(RSTTRXFILE!LINE_DISC)
                GRDBILL.TextMatrix(i, 4) = Val(RSTTRXFILE!SALES_TAX)
                'GRDBILL.TextMatrix(i, 5) = RSTTRXFILE!M_WEIGHT & "gms"
                GRDBILL.TextMatrix(i, 6) = Format(RSTTRXFILE!TRX_TOTAL, "0.00")
                'GRDBILL.TextMatrix(i, 7) = RSTTRXFILE!REF_NO
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
'    DB.Execute ("DELETE from SALESREG WHERE VCH_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 2) & " AND (TRX_TYPE='WO' OR TRX_TYPE='WO')")
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
    'MsgBox "Report file generated at " & Rptpath & "Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
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
    If OptPetty.Value = True Then
        RSTTRXFILE.Open "SELECT * From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='WO' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        RSTTRXFILE.Open "SELECT * From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
    Do Until RSTTRXFILE.EOF
        SN = SN + 1
        CMDDISPLAY.Tag = ""
        If RSTTRXFILE!SLSM_CODE = "A" Then
            CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(Round((RSTTRXFILE!DISCOUNT * RSTTRXFILE!VCH_AMOUNT) / 100, 2), "0.00"))
        ElseIf RSTTRXFILE!SLSM_CODE = "P" Then
            CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(RSTTRXFILE!DISCOUNT, "0.00"))
        End If
        CMDEXIT.Tag = ""
        CMDEXIT.Tag = IIf(IsNull(RSTTRXFILE!ADD_AMOUNT), "", RSTTRXFILE!ADD_AMOUNT)
        'SLIPAMT = SLIPAMT + RSTTRXFILE!VCH_AMOUNT - (Val(CMDDISPLAY.Tag) + Val(cmdview.Tag))
        Print #1, Chr(27) & Chr(71) & Space(5) & Chr(14) & Chr(15) & AlignRight(str(SN), 4) & ". " & Space(1) & _
            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
            AlignRight(Format(Round(RSTTRXFILE!VCH_AMOUNT - (Val(CMDDISPLAY.Tag) + Val(CMDEXIT.Tag)), 0), "0.00"), 16)
        'Print #1, Chr(13)
        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
        
        RSTSALEREG.AddNew
        RSTSALEREG!VCH_NO = RSTTRXFILE!VCH_NO
        If OptPetty.Value = True Then
            RSTSALEREG!TRX_TYPE = "WO"
        Else
            RSTSALEREG!TRX_TYPE = "HI"
        End If
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
    'MsgBox "Report file generated at " & Rptpath & "Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Function fillcount()
    Dim i, n As Long
    
    grdcount.rows = 0
    i = 0
    On Error GoTo ERRHAND
    For n = 1 To GRDTranx.rows - 1
        If GRDTranx.TextMatrix(n, 9) = "Y" Then
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
            i = i + 1
        End If
    Next n
    Exit Function
ERRHAND:
    MsgBox err.Description
    
End Function

Private Sub Optb2c_Click()
    TMPDELETE.Enabled = False
    CMDCONVERT.Enabled = False
    CMDCONVERT2.Enabled = False
End Sub

Private Sub Optpetty_Click()
    TMPDELETE.Enabled = False
    CMDCONVERT.Enabled = False
    CMDCONVERT2.Enabled = False
End Sub

Private Sub TMPDELETE_Click()
    
    If grdcount.rows = 0 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE THE SELECTED BILLS", vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then Exit Sub
    
    Dim TRXMAST, TRXMASTWO As ADODB.Recordset
    Dim n, LASTWOBILL, LASTBILL As Long
    Dim M_DATE As Date
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    LASTBILL = 0
    Set TRXMAST = New ADODB.Recordset
    If OptPetty.Value = True Then
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'WO' AND VCH_DATE = '" & Format(DateDiff("d", 1, DTFROM.Value), "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
    Else
        TRXMAST.Open "Select MAX(VCH_NO)  From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI' AND VCH_DATE = '" & Format(DateDiff("d", 1, DTFROM.Value), "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
    End If
    'TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_TYPE = 'WO' ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        LASTBILL = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
'    LASTWOBILL = 0
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXMAST_SP WHERE TRX_TYPE = 'WO'", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        LASTWOBILL = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
    Dim RSTTRXFILE, TRXFILEMAST As ADODB.Recordset
    For n = 0 To grdcount.rows - 1

'        LASTWOBILL = LASTWOBILL + 1
'
'        Set TRXMAST = New ADODB.Recordset
'        TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_TYPE='WO' AND VCH_NO = " & grdcount.TextMatrix(N, 4) & "", db, adOpenStatic, adLockReadOnly
'        Do Until TRXMAST.EOF
'            Set TRXMASTWO = New ADODB.Recordset
'            With TRXMASTWO
'                .Open "Select * FROM TRXMAST_SP", db, adOpenStatic, adLockOptimistic, adCmdText
'                .AddNew
'                !VCH_NO = LASTWOBILL
'                !TRX_TYPE = "WO"
'                !TRX_YEAR = Year(MDIMAIN.DTFROM.value)
'                !VCH_AMOUNT = TRXMAST!VCH_AMOUNT
'                !VCH_DATE = TRXMAST!VCH_DATE
'                !ACT_CODE = TRXMAST!ACT_CODE
'                !ACT_NAME = TRXMAST!ACT_NAME
'                !DISCOUNT = TRXMAST!DISCOUNT
'                !ADD_AMOUNT = 0
'                !ROUNDED_OFF = 0
'                !PAY_AMOUNT = TRXMAST!PAY_AMOUNT
'                !REF_NO = ""
'                !SLSM_CODE = TRXMAST!SLSM_CODE
'                !DISCOUNT = TRXMAST!DISCOUNT
'                !CHECK_FLAG = "I"
'                !POST_FLAG = TRXMAST!POST_FLAG
'                !CFORM_NO = TRXMAST!CFORM_NO
'                !Remarks = TRXMAST!Remarks
'                !DISC_PERS = 0
'                !AST_PERS = 0
'                !AST_AMNT = 0
'                !BANK_CHARGE = 0
'                !PHONE = "" '!PHONE
'                !CREATE_DATE = TRXMAST!CREATE_DATE
'                !MODIFY_DATE = TRXMAST!MODIFY_DATE
'                !C_USER_ID = "SM"
'                !cr_days = TRXMAST!cr_days
'                !BILL_NAME = TRXMAST!BILL_NAME
'                !BILL_ADDRESS = TRXMAST!BILL_ADDRESS
'                !AGENT_CODE = TRXMAST!AGENT_CODE
'                !AGENT_NAME = TRXMAST!AGENT_NAME
'                !COMM_AMT = TRXMAST!COMM_AMT
'                .Update
'            End With
'            TRXMASTWO.Close
'            Set TRXMASTWO = Nothing
'
'            TRXMAST.MoveNext
'        Loop
'        TRXMAST.Close
'        Set TRXMAST = Nothing
        
        Set TRXMAST = New ADODB.Recordset
        If OptPetty.Value = True Then
            TRXMAST.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & "", db, adOpenStatic, adLockReadOnly
        Else
            TRXMAST.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & "", db, adOpenStatic, adLockReadOnly
        End If
        Do Until TRXMAST.EOF
                
            Set TRXFILEMAST = New ADODB.Recordset
            TRXFILEMAST.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='" & TRXMAST!TRX_TYPE & "' AND VCH_NO = " & TRXMAST!VCH_NO & " AND LINE_NO = " & TRXMAST!LINE_NO & "", db, adOpenStatic, adLockReadOnly
            If Not (TRXFILEMAST.EOF And TRXFILEMAST.BOF) Then
                
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & TRXMAST!R_TRX_TYPE & "' AND RTRXFILE.VCH_NO = " & TRXMAST!R_VCH_NO & " AND RTRXFILE.LINE_NO = " & TRXMAST!R_LINE_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
                With RSTTRXFILE
                    If Not (.EOF And .BOF) Then
                        If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                        If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                        !ISSUE_QTY = !ISSUE_QTY - TRXFILEMAST!QTY
                        !BAL_QTY = !BAL_QTY + TRXFILEMAST!QTY
                        RSTTRXFILE.Update
                    End If
                End With
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & TRXFILEMAST!ITEM_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                With RSTTRXFILE
                    If Not (.EOF And .BOF) Then
                        !ISSUE_QTY = !ISSUE_QTY - TRXFILEMAST!QTY
                        If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                        !ISSUE_VAL = !ISSUE_VAL - TRXFILEMAST!TRX_TOTAL
                        !CLOSE_QTY = !CLOSE_QTY + TRXFILEMAST!QTY
                        If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                        !CLOSE_VAL = !CLOSE_VAL + TRXFILEMAST!TRX_TOTAL
                        RSTTRXFILE.Update
                    End If
                End With
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
'                Set RSTTRXFILE = New ADODB.Recordset
'                With RSTTRXFILE
'                    .Open "Select * FROM TRXFILE_SP", db, adOpenStatic, adLockOptimistic, adCmdText
'                    .AddNew
'                    !VCH_NO = LASTWOBILL
'                    !TRX_TYPE = "WO"
'                    !VCH_DATE = TRXFILEMAST!VCH_DATE
'                    !LINE_NO = TRXFILEMAST!LINE_NO
'                    !CATEGORY = TRXFILEMAST!CATEGORY
'                    !BARCODE = TRXFILEMAST!BARCODE
'                    !LOOSE_PACK = TRXFILEMAST!LOOSE_PACK
'                    !ACT_WEIGHT = TRXFILEMAST!ACT_WEIGHT
'                    !ITEM_CODE = TRXFILEMAST!ITEM_CODE
'                    !ITEM_NAME = TRXFILEMAST!ITEM_NAME
'                    !M_WEIGHT = TRXFILEMAST!M_WEIGHT
'                    !WEIGHT_TYPE = "gms"
'                    !M_PURITY = TRXFILEMAST!M_PURITY
'                    !M_RATE = TRXFILEMAST!M_RATE
'                    !STONE_WT = TRXFILEMAST!STONE_WT
'                    !STONE_WT_TYPE = "gms"
'                    !STONE_AMT = TRXFILEMAST!STONE_AMT
'                    !TOUCH_RATE = 0
'                    !SALES_TAX = TRXFILEMAST!SALES_TAX
'                    !LINE_DISC = TRXFILEMAST!LINE_DISC
'                    !CST = 0
'                    !VA_RATE = TRXFILEMAST!VA_RATE
'                    !VC_RATE = TRXFILEMAST!VC_RATE
'                    !OTHER_RATE = TRXFILEMAST!OTHER_RATE
'                    !TRX_TOTAL = TRXFILEMAST!TRX_TOTAL
'                    !COM_FLAG = TRXFILEMAST!COM_FLAG
'                    !COM_AMT = TRXFILEMAST!COM_AMT
'                    !ITEM_COST = TRXFILEMAST!ITEM_COST
'                    !MODIFY_DATE = Date
'                    !CREATE_DATE = TRXFILEMAST!CREATE_DATE
'                    !C_USER_ID = "SM"
'                    .Update
'                    .Close
'                End With
'                Set RSTTRXFILE = Nothing
            End If
            TRXFILEMAST.Close
            Set TRXFILEMAST = Nothing
            
'            Set RSTTRXFILE = New ADODB.Recordset
'            With RSTTRXFILE
'                .Open "Select * FROM TRXSUB_SP", db, adOpenStatic, adLockOptimistic, adCmdText
'                .AddNew
'                !VCH_NO = LASTWOBILL
'                !TRX_TYPE = "WO"
'                !LINE_NO = TRXMAST!LINE_NO
'                !R_VCH_NO = TRXMAST!R_VCH_NO
'                !R_LINE_NO = TRXMAST!R_LINE_NO
'                !R_TRX_TYPE = TRXMAST!R_TRX_TYPE
'                !QTY = 1
'                .Update
'                .Close
'            End With
'            Set RSTTRXFILE = Nothing
            
            TRXMAST.MoveNext
        Loop
        TRXMAST.Close
        Set TRXMAST = Nothing
        
        If OptPetty.Value = True Then
            db.Execute "delete FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            
            db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='DR' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'WO'"
            db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_VCH_NO = " & grdcount.TextMatrix(n, 4) & " AND B_TRX_TYPE = 'WO' "
            db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'WO' "
            db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'WO'"
        Else
            db.Execute "delete FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & grdcount.TextMatrix(n, 4) & ""
            
            db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='DR' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TRX_TYPE = 'HI'"
            db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND B_VCH_NO = " & grdcount.TextMatrix(n, 4) & " AND B_TRX_TYPE = 'HI' "
            db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND TRX_TYPE = 'RT' AND INV_TRX_TYPE = 'HI' "
            db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & grdcount.TextMatrix(n, 4) & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'HI'"
        End If
    Next n

    n = LASTBILL
    Set TRXMASTWO = New ADODB.Recordset
    If OptPetty.Value = True Then
        TRXMASTWO.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO > " & LASTBILL & " ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        TRXMASTWO.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO > " & LASTBILL & " ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
    Do Until TRXMASTWO.EOF
        n = n + 1
        Set TRXMAST = New ADODB.Recordset
        If OptPetty.Value = True Then
            TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        Else
            TRXMAST.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        End If
        Do Until TRXMAST.EOF
            TRXMAST!VCH_NO = n
            TRXMAST.Update
            TRXMAST.MoveNext
        Loop
        TRXMAST.Close
        Set TRXMAST = Nothing
        
        Set TRXMAST = New ADODB.Recordset
        If OptPetty.Value = True Then
            TRXMAST.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        Else
            TRXMAST.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        End If
        Do Until TRXMAST.EOF
            TRXMAST!VCH_NO = n
            TRXMAST.Update
            TRXMAST.MoveNext
        Loop
        TRXMAST.Close
        Set TRXMAST = Nothing
        
        Set TRXMAST = New ADODB.Recordset
        If OptPetty.Value = True Then
            TRXMAST.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='WO' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        Else
            TRXMAST.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & TRXMASTWO!VCH_NO & "", db, adOpenStatic, adLockOptimistic, adCmdText
        End If
        Do Until TRXMAST.EOF
            TRXMAST!VCH_NO = n
            TRXMAST.Update
            TRXMAST.MoveNext
        Loop
        TRXMAST.Close
        Set TRXMAST = Nothing
        
        TRXMASTWO.MoveNext
    Loop
    TRXMASTWO.Close
    Set TRXMASTWO = Nothing
        
    Call CmDDisplay_Click
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description

End Sub

